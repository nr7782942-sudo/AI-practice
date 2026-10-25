#!/usr/bin/env python3
"""Offline, standard-library-only PPTX evidence staging. Python 3.10+.

This program does not perform OCR, vision, scientific fact checking or LLM calls.
It copies originals, extracts evidence, checks explicit review records, and publishes
only a new directory. No command overwrites a pre-existing published collection.
"""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import math
import os
import posixpath
import re
import shutil
import struct
import sys
import uuid
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET

SCHEMA_VERSION = 1
MAX_ARCHIVE_BYTES = 512 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 1024 * 1024 * 1024
MAX_PART_BYTES = 128 * 1024 * 1024
MAX_PARTS = 20000
MAX_COMPRESSION_RATIO = 1000
NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
NUMERIC = re.compile(r'(?<![\d.])[+\-−]?(?:\d+(?:[.,]\d+)*|\.\d+)(?:[eE][+\-]?\d+)?(?:[%‰])?')
RESERVED = re.compile(r'^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)', re.I)


class PipelineError(Exception):
    pass


def utc_now():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


def sha_bytes(data):
    return hashlib.sha256(data).hexdigest()


def sha_file(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def png_dimensions(path):
    """Check PNG framing, CRCs and bounded image data; not visual semantics."""
    if Path(path).stat().st_size > MAX_PART_BYTES:
        raise PipelineError('PNG exceeds the 128 MiB file limit.')
    data = Path(path).read_bytes()
    if data[:8] != b'\x89PNG\r\n\x1a\n':
        raise PipelineError('Invalid PNG signature.')
    offset, width, height = 8, 0, 0
    compressed = bytearray()
    ended = False
    seen_ihdr = False
    while offset < len(data):
        if offset + 12 > len(data):
            raise PipelineError('Truncated PNG chunk.')
        length = struct.unpack('>I', data[offset:offset + 4])[0]
        kind = data[offset + 4:offset + 8]
        end = offset + 12 + length
        if end > len(data):
            raise PipelineError('Truncated PNG chunk payload.')
        payload = data[offset + 8:offset + 8 + length]
        crc = struct.unpack('>I', data[offset + 8 + length:end])[0]
        if zlib.crc32(kind + payload) & 0xffffffff != crc:
            raise PipelineError('PNG checksum mismatch.')
        if not seen_ihdr and kind != b'IHDR':
            raise PipelineError('PNG must start with IHDR.')
        if kind == b'IHDR':
            if seen_ihdr or length != 13:
                raise PipelineError('Invalid/duplicate PNG IHDR.')
            width, height, depth, color, compression, filtering, interlace = struct.unpack('>IIBBBBB', payload)
            if width <= 0 or height <= 0 or width * height > 50000000:
                raise PipelineError('Invalid/oversized PNG dimensions.')
            allowed = {0: (1, 2, 4, 8, 16), 2: (8, 16), 3: (1, 2, 4, 8), 4: (8, 16), 6: (8, 16)}
            if depth not in allowed.get(color, ()) or compression != 0 or filtering != 0 or interlace not in (0, 1):
                raise PipelineError('Unsupported PNG encoding.')
            seen_ihdr = True
        elif kind == b'IDAT':
            compressed.extend(payload)
        elif kind == b'IEND':
            if length or end != len(data):
                raise PipelineError('Invalid PNG ending/trailing data.')
            ended = True
            break
        offset = end
    if not ended or not compressed:
        raise PipelineError('PNG is missing image data or IEND.')
    bits = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[color] * depth
    passes = [(0, 0, 1, 1)] if interlace == 0 else [(0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2)]
    expected = 0
    for x, y, dx, dy in passes:
        pw, ph = max(0, (width - x + dx - 1) // dx), max(0, (height - y + dy - 1) // dy)
        if pw and ph:
            expected += ph * (1 + (pw * bits + 7) // 8)
    if expected > 256 * 1024 * 1024:
        raise PipelineError('PNG decoded data exceeds 256 MiB.')
    try:
        decoder = zlib.decompressobj()
        raw = decoder.decompress(bytes(compressed), expected + 1)
    except zlib.error as e:
        raise PipelineError('PNG compressed image data is invalid.') from e
    if len(raw) != expected or not decoder.eof or decoder.unused_data or decoder.unconsumed_tail:
        raise PipelineError('PNG decoded image length/stream is invalid.')
    return width, height


def dump_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8', newline='\n') as f:
        json.dump(value, f, ensure_ascii=False, indent=2)
        f.write('\n')


def write_new(path, text):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8', newline='\n') as f:
        f.write(text)


def safe_relative(value):
    if not isinstance(value, str) or not value or '\\' in value or '\x00' in value:
        raise PipelineError(f'Unsafe relative path: {value!r}')
    p = PurePosixPath(value)
    if p.is_absolute() or value.startswith('/') or any(c in value for c in ':<>|?*'):
        raise PipelineError(f'Unsafe relative path: {value!r}')
    if any(v in ('', '.', '..') or v.endswith((' ', '.')) or RESERVED.match(v) for v in value.split('/')):
        raise PipelineError(f'Unsafe Windows path: {value!r}')
    return value


def is_link_or_reparse(path):
    path = Path(path)
    try:
        return path.is_symlink() or bool(getattr(path.lstat(), 'st_file_attributes', 0) & 0x400)
    except FileNotFoundError:
        return False


def require_local_path(path):
    """Reject UNC, mapped network drives and Windows reparse points before resolving."""
    value = os.fspath(path)
    if value.startswith(('\\\\', '//')):
        raise PipelineError('UNC/network paths are unsupported; use an approved local directory.')
    candidate = Path(os.path.abspath(value))
    if os.name == 'nt' and candidate.anchor:
        if ctypes.windll.kernel32.GetDriveTypeW(str(candidate.anchor)) == 4:
            raise PipelineError('Mapped network drives are unsupported; use an approved local directory.')
    for part in [candidate] + list(candidate.parents):
        if is_link_or_reparse(part):
            raise PipelineError(f'Linked/reparse path is unsupported: {part}')
    return candidate


def under(root, relative):
    relative = safe_relative(relative)
    root = Path(root).resolve()
    candidate = root.joinpath(*relative.split('/'))
    if not candidate.resolve().is_relative_to(root):
        raise PipelineError(f'Path escapes root: {relative}')
    # Do not follow directory junctions/symlinks even if they happen to point inside.
    for parent in [candidate] + list(candidate.parents):
        if parent == root:
            break
        if is_link_or_reparse(parent):
            raise PipelineError(f'Linked path is not allowed: {parent}')
    return candidate


def parse_xml(data, part):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise PipelineError(f'DTD/entity declarations are not supported: {part}')
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise PipelineError(f'Malformed XML {part}: {e}') from e


def read_archive(path):
    path = Path(path)
    if path.suffix.lower() != '.pptx':
        raise PipelineError('Only ordinary .pptx files are accepted. Export .ppt/.pptm locally to .pptx first.')
    if path.stat().st_size > MAX_ARCHIVE_BYTES:
        raise PipelineError('PPTX exceeds the 512 MiB compressed archive limit.')
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        raise PipelineError('Not a readable unencrypted PPTX ZIP file.') from e
    with z:
        infos = z.infolist()
        if len(infos) > MAX_PARTS:
            raise PipelineError('PPTX has too many ZIP entries.')
        total = 0
        seen = set()
        for info in infos:
            name = info.filename.rstrip('/') if info.is_dir() else info.filename
            safe_relative(name)
            if name.casefold() in seen:
                raise PipelineError(f'Duplicate/case-colliding ZIP entry: {name}')
            seen.add(name.casefold())
            if info.flag_bits & 1:
                raise PipelineError('Encrypted ZIP entries are unsupported.')
            if (info.external_attr >> 16) & 0o170000 == 0o120000:
                raise PipelineError('ZIP symlinks are unsupported.')
            if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                raise PipelineError(f'Unsupported ZIP compression: {name}')
            total += info.file_size
            if info.file_size > MAX_PART_BYTES or total > MAX_UNCOMPRESSED_BYTES:
                raise PipelineError('Uncompressed package exceeds resource limits.')
            if info.file_size > 1024 * 1024 and info.file_size / max(info.compress_size, 1) > MAX_COMPRESSION_RATIO:
                raise PipelineError(f'Excessive ZIP compression ratio: {name}')
        parts = {}
        for info in infos:
            if not info.is_dir():
                parts[info.filename] = z.read(info)
        for name, data in parts.items():
            if name.endswith(('.xml', '.rels')):
                parse_xml(data, name)
        for required in ('[Content_Types].xml', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels'):
            if required not in parts:
                raise PipelineError(f'Missing PPTX part: {required}')
        types = parts['[Content_Types].xml'].lower()
        if b'macroenabled' in types or any('vbaproject' in n.lower() for n in parts):
            raise PipelineError('Macro-bearing packages are unsupported even when renamed .pptx.')
        return parts


def rels_part(owner):
    return posixpath.join(posixpath.dirname(owner), '_rels', posixpath.basename(owner) + '.rels')


def resolve_target(owner, target):
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise PipelineError(f'Unsupported internal relationship target: {target}')
    target = unquote(parsed.path)
    if '\\' in target or '\x00' in target:
        raise PipelineError(f'Unsafe relationship target: {target}')
    base = '' if target.startswith('/') else posixpath.dirname(owner)
    result = posixpath.normpath(posixpath.join(base, target.lstrip('/')))
    return safe_relative(result)


def relationships(parts, owner, media_map=None):
    name = rels_part(owner)
    if name not in parts:
        return []
    root = parse_xml(parts[name], name)
    result = []
    ids = set()
    for node in root:
        rid = node.get('Id')
        if not rid or rid in ids:
            raise PipelineError(f'Missing/duplicate relationship ID in {name}')
        ids.add(rid)
        external = node.get('TargetMode') == 'External'
        target = node.get('Target', '')
        entry = {'id': rid, 'type': node.get('Type', ''), 'target': target, 'external': external}
        if not external:
            part = resolve_target(owner, target)
            entry['source_part'] = part
            entry['exists'] = part in parts
            if part not in parts:
                raise PipelineError(f'Missing relationship target {part} (from {owner})')
            entry['asset_path'] = (media_map or {}).get(part, f'source/package/{part}')
        result.append(entry)
    return result


def local_name(tag):
    return tag.split('}')[-1]


def text_records(data, part, origin):
    root = parse_xml(data, part)
    output = []
    for i, paragraph in enumerate(root.iter(f"{{{NS['a']}}}p"), 1):
        chunks = []
        for node in paragraph.iter():
            tag = local_name(node.tag)
            if tag == 't':
                chunks.append(node.text or '')
            elif tag == 'br':
                chunks.append('\n')
            elif tag == 'tab':
                chunks.append('\t')
        output.append({'locator': f'{part}#p:{i:05d}', 'origin': origin, 'text': ''.join(chunks)})
    # Cached chart values, diagram strings, and equations outside DrawingML paragraphs.
    if origin == 'related_xml':
        covered = {id(x) for p in root.iter(f"{{{NS['a']}}}p") for x in p.iter()}
        for i, node in enumerate(root.iter(), 1):
            if id(node) not in covered and not list(node) and node.text is not None and node.text.strip():
                output.append({'locator': f'{part}#leaf:{i:05d}', 'origin': origin,
                               'tag': local_name(node.tag), 'text': node.text})
    return output


def shape_records(root, rels, records):
    relation_by_id = {r['id']: r for r in rels}
    paragraph_ids = {id(node): i for i, node in enumerate(root.iter(f"{{{NS['a']}}}p"), 1)}
    shapes = []
    tree = root.find('p:cSld/p:spTree', NS)
    if tree is None:
        return shapes

    def visit(parent, parent_group=None):
        for child in parent:
            kind = local_name(child.tag)
            if kind not in ('sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'contentPart'):
                continue
            nv = next(child.iter(f"{{{NS['p']}}}cNvPr"), None)
            sid = nv.get('id') if nv is not None else None
            rid_values = []
            for node in child.iter():
                for key, value in node.attrib.items():
                    if key.startswith('{' + NS['r'] + '}') and value not in rid_values:
                        rid_values.append(value)
            xfrm = next((x for x in child.iter() if local_name(x.tag) == 'xfrm'), None)
            geometry = None if xfrm is None else {
                'attributes': dict(xfrm.attrib),
                'coordinates': [{'kind': local_name(x.tag), **x.attrib} for x in xfrm],
                'unit': 'EMU', 'coordinate_frame': 'parent_group' if parent_group else 'slide',
                'group_transform_applied': False,
            }
            pnums = [paragraph_ids[id(x)] for x in child.iter(f"{{{NS['a']}}}p")]
            entry = {'shape_id': sid, 'kind': kind, 'name': nv.get('name', '') if nv is not None else '',
                     'description': nv.get('descr', '') if nv is not None else '',
                     'title': nv.get('title', '') if nv is not None else '',
                     'parent_group_id': parent_group, 'geometry': geometry,
                     'paragraph_indices': pnums,
                     'relationships': [relation_by_id.get(rid, {'id': rid, 'unresolved': True}) for rid in rid_values]}
            if kind == 'graphicFrame':
                table = child.find('.//a:tbl', NS)
                if table is not None:
                    entry['table'] = []
                    for row_no, row in enumerate(table.findall('a:tr', NS), 1):
                        for col_no, cell in enumerate(row.findall('a:tc', NS), 1):
                            entry['table'].append({'row': row_no, 'column': col_no, 'attributes': dict(cell.attrib),
                                                   'paragraph_indices': [paragraph_ids[id(x)] for x in cell.iter(f"{{{NS['a']}}}p")]})
            shapes.append(entry)
            if kind == 'grpSp':
                visit(child, sid)
    visit(tree)
    return shapes


def review_schema():
    return {
        '$schema': 'https://json-schema.org/draft/2020-12/schema', 'title': 'DRAM slide review record (offline; not fetched)',
        'type': 'object', 'required': ['schema_version', 'slides'],
        'properties': {'schema_version': {'const': 1}, 'slides': {'type': 'array', 'items': {'type': 'object',
            'required': ['slide_no', 'status', 'reviewer', 'reviewed_at', 'claims', 'numbers', 'figure_interpretations', 'uncertainties'],
            'properties': {'slide_no': {'type': 'integer', 'minimum': 1}, 'status': {'enum': ['pending', 'reviewed']},
                'reviewer': {'type': 'string'}, 'reviewed_at': {'type': 'string'}, 'claims': {'type': 'array'},
                'numbers': {'type': 'array'}, 'figure_interpretations': {'type': 'array'}, 'uncertainties': {'type': 'array', 'items': {'type': 'string'}}}}}},
        '$comment': 'The accompanying REVIEW_GUIDE.md and validator define the evidence contract. This schema is intentionally minimal; no jsonschema dependency is required.'}


REVIEW_GUIDE = '''# 검토 기록 형식 — 자동 사실 검증 아님

`review.json`의 모든 슬라이드는 처음에 `pending`입니다. 노트와 검토 기록을 사람이 또는 승인된 사내 모델이 원문과 대조한 뒤에만 `reviewed`로 바꿉니다. `reviewer`에는 실제 검토자/모델 식별자를, `reviewed_at`에는 ISO 8601 일시를 입력합니다. 도구 실행만으로 reviewed로 바꾸지 않습니다.

모든 배열은 남겨두세요. `claims`가 비어 있으면 `uncertainties`에 왜 추출 가능한 지식이 없는지/미해결인지 적으세요. 숫자가 없으면 `numbers: []`로 둡니다. 검토 완료는 불확실성 해소나 의미 정확성 보증이 아닙니다.

## claim

각 항목: `{"id":"S001-C001","text":"원문 범위를 벗어나지 않는 문장","claim_type":"text","evidence":[{"kind":"pptx_text","locator":"ppt/slides/slide7.xml#p:00001","quote":"정확한 원문 인용"}],"numbers":[]}`.

`claim_type`: `text`, `visual`, `inference`. `inference`는 text에 `[추론]`을 표시하고 `uncertainties`에 근거 한계를 기록하세요. `visual`은 실제 확인한 vision 증거가 반드시 필요합니다.

`kind`: `pptx_text`, `speaker_notes`, `vision`, `ocr`. 텍스트 locator는 해당 slide JSON의 text_records에서 그대로 복사하고, quote는 그 record.text의 정확한 부분문자열이어야 합니다. `pptx_text`에는 관련 chart/diagram XML도 포함할 수 있지만 XML cache를 현재 그래프의 렌더 모습으로 단정하지 마세요. speaker_notes는 발표자 메모에만 사용합니다.

`vision`/`ocr`의 locator는 `assets/slides/slide-0001.png` 또는 manifest에 실린 해당 슬라이드의 `assets/media/...` 파일 경로입니다. `region`으로 화면 위치/축/범례를 명시하고, OCR은 `quote`에 실제 OCR 결과를 기록하세요. OCR 결과는 그림 해석의 증거가 아닙니다. 이미지 존재만으로 vision 완료가 되지 않습니다.

## numbers

claim 안의 숫자 표기는 모두 claim.numbers에 연결하세요. 숫자 정규식은 제품명/식별자 안의 숫자까지 보수적으로 확인합니다. 예: `{"verbatim":"1.20 V","value":"1.20","unit":"V","conditions":"25 °C, 원문 조건; 없으면 미명시","source_locator":"ppt/slides/slide7.xml#p:00001","checked":true}`. verbatim은 claim.text에 존재해야 하며 숫자 토큰을 포함해야 합니다. 원문 정확 표기도 text record에서 확인할 수 있어야 합니다. 차트에서 읽은 값은 source_locator에 확인한 이미지 경로를 넣고 실제 vision figure_interpretations가 있어야 합니다. 계산/단위 변환 결과를 원문 숫자로 등록하지 마세요. 파생 계산은 Q&A 또는 별도 파생 노트에서 입력값·식·반올림을 설명하고, 이 리뷰의 원문 정량 claim에는 등록하지 않습니다. checked:true는 실제 원문/단위/조건을 대조했다는 수동 선언입니다.

## figure_interpretations

최소 하나의 검토 기록: `{"asset_path":"assets/slides/slide-0001.png","method":"vision","description":"관찰한 배치, 축, 범례와 연결 관계","uncertainties":["해상도상 판독 불가한 항목"],"assessed_by":"실제 사내 모델/사람","scope":"whole_slide"}`. `method`는 `vision`, `ocr_only`, `unavailable`만 허용합니다. 파일을 못 보면 `{"asset_path":null,"method":"unavailable","description":"렌더링/시각 확인 미수행","uncertainties":["그림 내용 미확인"],"assessed_by":"검토자","scope":"whole_slide"}`로 남기세요. `ocr_only`는 실제 OCR만 실행한 경우이며 시각 확인 성공으로 세지 않습니다. vision은 사내 Vision 지원 여부와 실제 이미지 전달을 확인한 뒤에만 기록합니다.

## 파일과 게시

원본/기계 추출 파일은 수정하지 마세요. 해석은 notes/slide-XXXX.md와 review.json에만 작성합니다. 노트에 원본 슬라이드 전체 그림을 넣을 때 실제 파일이 생긴 후 `![[<Vault 상대 staging 경로>/assets/slides/slide-XXXX.png]]`를 사용하세요. 그림 일부를 별도 추가하는 자동 자산 작업은 이 버전에 포함되지 않습니다.

validate는 해시, 파일, 인용, 링크, 필수 기록과 선언을 검사합니다. 원문 의미, 해석의 정확성, 누락, 실제로 모델이 이미지를 보았는지는 판정하지 못합니다. publish는 검토 기록 완료를 요구하고 새 디렉터리에 복사합니다. notes의 원문 근거가 부족한 자유 서술은 자동으로 사실 승인되지 않습니다. 게시 후에도 사내 전문가 표본 검토가 필요합니다.
'''


def note_template(run_rel, slide, records):
    n = slide['slide_no']
    lines = [f'---\ntype: dram-source-slide\nslide_no: {n}\nsource_part: "{slide["source_part"]}"\nreview_status: pending\n---\n',
             f'# 슬라이드 {n:04d}\n',
             f'- 원본: [[{run_rel}/source/original.pptx]]',
             f'- 원문 위치: `{slide["source_part"]}`',
             f'- 숨김 슬라이드: {slide["hidden"]}',
             f'- 기계 추출: [[{run_rel}/{slide["json_path"]}]]',
             '- 검토 상태: 미검토. 텍스트 추출·OCR·실제 Vision 해석을 구분한다.\n',
             '## 원본 슬라이드 이미지\n',
             f'렌더링 전: 이미지 없음. 실제 생성 후에만 다음 경로를 위키 임베드로 삽입: `{run_rel}/{slide["render_path"]}`\n',
             '## 원문 전사 (전체; 읽기 순서/도표 배치는 보장하지 않음)\n']
    for rec in records:
        # A fence longer than any run in source avoids source-created Markdown escapes.
        longest = max([len(x) for x in re.findall(r'`+', rec['text'])] + [2])
        fence = '`' * (longest + 1)
        lines.extend([f'### `{rec["locator"]}` — {rec["origin"]}', f'{fence}text\n{rec["text"]}\n{fence}\n'])
    lines.extend(['## 해석 및 주장\n\n미작성. 각 주장 ID와 정확한 locator/인용을 review.json에 함께 기록한다.\n',
                  '## 수치·단위·조건 대조\n\n미검증. 원문 표기, 단위, 축/범례, 조건, 비교 기준을 보존한다.\n',
                  '## 그림/표 해석\n\n미수행. Vision / OCR만 수행 / 확인 불가 중 실제 수행 경로를 명시한다.\n',
                  '## 불확실성 및 추가 확인\n\n- 원문 의미와 그림 연결 관계는 아직 검토하지 않았다.\n',
                  '## 관련 지식\n\n검증한 다른 Vault 노트가 있을 때만 위키링크를 추가한다.\n'])
    return '\n'.join(lines)


def extract(source, vault):
    source = require_local_path(source).resolve(strict=True)
    vault = require_local_path(vault).resolve()
    if not vault.is_dir():
        raise PipelineError('Vault folder must already exist. Create/select the approved local Vault first.')
    parts = read_archive(source)
    source_hash = sha_file(source)
    deck_id = source_hash[:12] + '_' + datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ') + '_' + uuid.uuid4().hex[:8]
    run_rel = f'_dram/staging/{deck_id}'
    run = under(vault, run_rel)
    run.mkdir(parents=True, exist_ok=False)
    (run / 'assets/slides').mkdir(parents=True)
    immutable = []

    def store_bytes(relative, data):
        dest = under(run, relative)
        dest.parent.mkdir(parents=True, exist_ok=True)
        with dest.open('xb') as f:
            f.write(data)
        immutable.append({'path': relative, 'sha256': sha_bytes(data), 'size': len(data)})

    # Read once again and compare after copy to detect a source modified during extraction.
    source_bytes = source.read_bytes()
    if sha_bytes(source_bytes) != source_hash:
        raise PipelineError('Source changed while reading. Retry from a stable local copy.')
    store_bytes('source/original.pptx', source_bytes)
    for part, data in parts.items():
        store_bytes(f'source/package/{part}', data)
    media_map = {}
    assets = []
    for part, data in parts.items():
        if part.startswith('ppt/media/'):
            name = re.sub(r'[^A-Za-z0-9._-]', '_', posixpath.basename(part))
            rel = f'assets/media/{hashlib.sha256(part.encode()).hexdigest()[:16]}_{name}'
            media_map[part] = rel
            store_bytes(rel, data)
            assets.append({'source_part': part, 'path': rel, 'sha256': sha_bytes(data), 'size': len(data)})
    presentation = parse_xml(parts['ppt/presentation.xml'], 'ppt/presentation.xml')
    present_rels = {x['id']: x for x in relationships(parts, 'ppt/presentation.xml', media_map)}
    ids = presentation.find('p:sldIdLst', NS)
    if ids is None or not len(ids):
        raise PipelineError('Presentation has no slides.')
    external = []
    global_flags = set()
    for part in parts:
        if part.endswith('.rels'):
            node = parse_xml(parts[part], part)
            for r in node:
                if r.get('TargetMode') == 'External':
                    external.append({'relationships_part': part, 'id': r.get('Id'), 'type': r.get('Type'), 'target': r.get('Target')})
    if external:
        global_flags.add('external_relationship')
    if any('/activeX/' in p or 'vbaProject' in p for p in parts):
        global_flags.add('macro_or_activex')
    if any('/embeddings/' in p for p in parts):
        global_flags.add('embedded_object')
    slides = []
    reviews = []
    used_parts = set()
    for ordinal, node in enumerate(ids, 1):
        rid = node.get('{' + NS['r'] + '}id')
        rel = present_rels.get(rid)
        if not rel or rel.get('external') or not rel['type'].endswith('/slide'):
            raise PipelineError(f'Invalid presentation slide relationship at ordinal {ordinal}')
        part = rel['source_part']
        if part in used_parts:
            raise PipelineError('One slide part is referenced twice in presentation order.')
        used_parts.add(part)
        root = parse_xml(parts[part], part)
        records = text_records(parts[part], part, 'slide_text')
        srels = relationships(parts, part, media_map)
        notes = []
        related = []
        queue = [r for r in srels if not r['external'] and any(k in r['type'] for k in ('notesSlide', '/chart', '/diagram', '/comments'))]
        related_seen = set()
        while queue:
            r = queue.pop(0)
            rp = r['source_part']
            if rp in related_seen:
                continue
            related_seen.add(rp)
            if '/notesSlide' in r['type']:
                notes.append(rp)
                records.extend(text_records(parts[rp], rp, 'speaker_notes'))
            elif rp.endswith('.xml'):
                related.append(rp)
                records.extend(text_records(parts[rp], rp, 'related_xml'))
            queue.extend(x for x in relationships(parts, rp, media_map) if not x['external'] and any(k in x['type'] for k in ('/chart', '/diagram')))
        flags = set()
        raw = parts[part]
        if any('/chart' in r['type'] for r in srels):
            flags.add('chart_requires_visual_review')
        if any('/diagram' in r['type'] for r in srels):
            flags.add('smartart_requires_visual_review')
        if b'oleObj' in raw:
            flags.add('ole_object')
        if b'officeDocument/2006/math' in raw or b'oMath' in raw:
            flags.add('equation_requires_visual_review')
        if b'<p:timing' in raw:
            flags.add('animation_or_build_not_captured_by_static_render')
        if any(r['external'] for r in srels):
            flags.add('external_relationship')
        flags.add('layout_master_fonts_reading_order_require_visual_review')
        base = f'slide-{ordinal:04d}'
        slide = {'slide_no': ordinal, 'presentation_slide_id': node.get('id'), 'source_part': part,
                 'hidden': root.get('show', '1') in ('0', 'false'), 'notes_parts': notes, 'related_xml_parts': related,
                 'json_path': f'slides/{base}.json', 'text_path': f'slides/{base}.txt', 'note_path': f'notes/{base}.md',
                 'render_path': f'assets/slides/{base}.png', 'raw_xml_path': f'source/package/{part}'}
        data = {**slide, 'text_records': records, 'shapes': shape_records(root, srels, records),
                'relationships': srels, 'flags': sorted(flags),
                'extraction_notes': ['No OCR or vision performed.', 'Text XML order is not guaranteed visual reading order.',
                                     'Group geometry is preserved without flattening transforms.',
                                     'Static slide rendering does not preserve animation states or interactive objects.']}
        store_bytes(slide['json_path'], (json.dumps(data, ensure_ascii=False, indent=2) + '\n').encode('utf-8'))
        full_text = '\n\n'.join(f'[{r["locator"]}] ({r["origin"]})\n{r["text"]}' for r in records)
        store_bytes(slide['text_path'], (full_text + '\n').encode('utf-8'))
        write_new(run / slide['note_path'], note_template(run_rel, slide, records))
        slides.append(slide)
        global_flags.update(flags)
        reviews.append({'slide_no': ordinal, 'status': 'pending', 'reviewer': '', 'reviewed_at': '', 'claims': [], 'numbers': [],
                        'figure_interpretations': [], 'uncertainties': ['원문 의미, 수치 및 시각 자료 미검토']})
    manifest = {'schema_version': 1, 'created_at': utc_now(), 'run_id': deck_id,
                'vault_path': str(vault), 'run_vault_relative': run_rel,
                'source': {'path': 'source/original.pptx', 'sha256': source_hash, 'original_name': source.name},
                'source_sha256': source_hash, 'slide_count': len(slides), 'slides': slides,
                'assets': assets, 'external_relationships': external, 'flags': sorted(global_flags), 'immutable_files': immutable,
                'validation_scope': 'Mechanical consistency only. No semantic, OCR, vision or scientific verification.'}
    dump_json(run / 'manifest.json', manifest)
    dump_json(run / 'review.json', {'schema_version': 1, 'slides': reviews})
    dump_json(run / 'review.schema.json', review_schema())
    write_new(run / 'REVIEW_GUIDE.md', REVIEW_GUIDE)
    report = validate(run)
    if report['errors']:
        raise PipelineError('New extraction failed consistency validation: ' + '; '.join(report['errors']))
    return {'run': str(run), 'slide_count': len(slides), 'source_sha256': source_hash,
            'status': 'staged_unreviewed', 'next': 'Review source and notes; render locally if approved; edit review.json; validate --require-reviewed.'}


def read_json(path):
    try:
        with Path(path).open(encoding='utf-8-sig') as f:
            return json.load(f)
    except (OSError, ValueError) as e:
        raise PipelineError(f'Cannot read JSON {path}: {e}') from e


def validate(run, require_reviewed=False):
    run = require_local_path(run).resolve(strict=True)
    manifest = read_json(run / 'manifest.json')
    errors, warnings = [], []
    def error(msg):
        errors.append(msg)
    def check(condition, msg):
        if not condition:
            error(msg)
        return condition
    def existing(rel):
        try:
            p = under(run, rel)
            if not p.is_file():
                raise PipelineError(f'Missing file {rel}')
            return p
        except (PipelineError, OSError) as e:
            error(str(e))
            return None
    check(manifest.get('schema_version') == 1, 'Unsupported manifest schema.')
    slide_entries = manifest.get('slides', [])
    check(isinstance(slide_entries, list) and len(slide_entries) == manifest.get('slide_count') and bool(slide_entries), 'Slide count mismatch.')
    expected = list(range(1, len(slide_entries) + 1))
    check([x.get('slide_no') for x in slide_entries] == expected, 'Slide ordinals are not complete and ordered.')
    source_info = manifest.get('source', {})
    src = existing(source_info.get('path', ''))
    source_parts = None
    if src:
        digest = sha_file(src)
        check(digest == source_info.get('sha256') == manifest.get('source_sha256'), 'Original source SHA256 mismatch.')
        try:
            source_parts = read_archive(src)
            presentation = parse_xml(source_parts['ppt/presentation.xml'], 'ppt/presentation.xml')
            rels = {x['id']: x for x in relationships(source_parts, 'ppt/presentation.xml')}
            ordered = [rels[n.get('{' + NS['r'] + '}id')]['source_part'] for n in presentation.find('p:sldIdLst', NS)]
            check(ordered == [x.get('source_part') for x in slide_entries], 'Manifest slide mapping disagrees with original PPTX order.')
        except (PipelineError, KeyError, TypeError) as e:
            error(f'Source revalidation failed: {e}')
    immutable = manifest.get('immutable_files', [])
    check(isinstance(immutable, list) and bool(immutable), 'Immutable artifact list is missing.')
    tracked = set()
    for item in immutable:
        rel = item.get('path', '')
        check(rel not in tracked, f'Duplicate immutable path: {rel}')
        tracked.add(rel)
        p = existing(rel)
        if p:
            check(p.stat().st_size == item.get('size') and sha_file(p) == item.get('sha256'), f'Extracted artifact changed: {rel}')
    if source_parts:
        for part, data in source_parts.items():
            rel = f'source/package/{part}'
            check(rel in tracked, f'Untracked source package part: {part}')
            p = existing(rel)
            if p:
                check(sha_file(p) == sha_bytes(data), f'Package part differs from source: {part}')
    assets = manifest.get('assets', [])
    asset_map = {x.get('path'): x for x in assets}
    for a in assets:
        p = existing(a.get('path', ''))
        if p:
            check(sha_file(p) == a.get('sha256'), f'Media SHA mismatch: {a.get("path")}')
            if source_parts:
                check(a.get('source_part') in source_parts and sha_bytes(source_parts[a['source_part']]) == a.get('sha256'), 'Media mapping differs from source.')
    render_map = {}
    if (run / 'render-manifest.json').exists():
        rm = read_json(run / 'render-manifest.json')
        check(rm.get('source_sha256') == source_info.get('sha256'), 'Render source SHA mismatch.')
        check(rm.get('method') in ('powerpoint_com', 'manual_powerpoint_export'), 'Unknown renderer; evidence approval requires documented renderer integration.')
        if rm.get('method') == 'manual_powerpoint_export':
            check(rm.get('mapping_confirmed') is True, 'Manual rendering requires explicit slide mapping confirmation.')
        for r in rm.get('renders', []):
            n = r.get('slide_no')
            check(n in expected and n not in render_map, f'Invalid/duplicate render ordinal: {n}')
            render_map[n] = r
            check(r.get('path') == f'assets/slides/slide-{n:04d}.png' if isinstance(n, int) else False, 'Render filename/ordinal mismatch.')
            p = existing(r.get('path', ''))
            if p:
                check(sha_file(p) == r.get('sha256'), f'Render SHA mismatch: {n}')
                try:
                    dimensions = png_dimensions(p)
                    check(dimensions == (r.get('width'), r.get('height')), f'Render dimensions mismatch: {n}')
                except PipelineError as e:
                    error(f'Invalid PNG render {n}: {e}')
            check(all(isinstance(r.get(k), (int, float)) and r[k] > 0 for k in ('width', 'height')), f'Invalid render dimensions: {n}')
        check(sorted(render_map) == expected, 'Render manifest is incomplete.')
    slides_by_no = {}
    for s in slide_entries:
        n = s.get('slide_no')
        p = existing(s.get('json_path', ''))
        existing(s.get('text_path', ''))
        existing(s.get('note_path', ''))
        if not p:
            continue
        d = read_json(p)
        slides_by_no[n] = d
        check(d.get('slide_no') == n and d.get('source_part') == s.get('source_part'), f'Slide JSON mapping mismatch: {n}')
        if source_parts and d.get('source_part') in source_parts:
            original_records = text_records(source_parts[d['source_part']], d['source_part'], 'slide_text')
            check([x for x in d.get('text_records', []) if x.get('origin') == 'slide_text'] == original_records, f'Raw slide text differs from original: {n}')
        for shape in d.get('shapes', []):
            for r in shape.get('relationships', []):
                check(not r.get('unresolved'), f'Unresolved shape relationship on slide {n}: {r.get("id")}')
        for r in d.get('relationships', []):
            if not r.get('external'):
                existing(r.get('asset_path', ''))
        render = under(run, s.get('render_path', ''))
        if render.exists() and n not in render_map:
            error(f'Unregistered render exists for slide {n}; require render-manifest.json.')
    review = read_json(run / 'review.json')
    check(review.get('schema_version') == 1, 'Unsupported review schema.')
    rslides = review.get('slides', [])
    check(isinstance(rslides, list) and sorted([x.get('slide_no', -1) for x in rslides]) == expected, 'Review slide coverage must match all slide ordinals exactly once.')
    reviewed = 0
    all_claim_ids = set()
    for rs in rslides:
        n = rs.get('slide_no')
        label = f'Slide {n}'
        d = slides_by_no.get(n, {})
        records = {x.get('locator'): x for x in d.get('text_records', [])}
        available = {r.get('asset_path') for r in d.get('relationships', []) if not r.get('external') and r.get('asset_path') in asset_map}
        if n in render_map:
            available.add(render_map[n]['path'])
        status = rs.get('status')
        check(status in ('pending', 'reviewed'), label + ': invalid status.')
        if require_reviewed:
            check(status == 'reviewed', label + ': review is still pending.')
        for key in ('claims', 'numbers', 'figure_interpretations', 'uncertainties'):
            check(isinstance(rs.get(key), list), label + f': {key} must be an array.')
        if status != 'reviewed':
            continue
        reviewed += 1
        note_entry = next((s for s in slide_entries if s.get('slide_no') == n), None)
        note_text = ''
        if note_entry and under(run, note_entry['note_path']).is_file():
            note_text = under(run, note_entry['note_path']).read_text(encoding='utf-8-sig')
        if require_reviewed:
            frontmatter = re.match(r'^---\s*\r?\n(.*?)\r?\n---(?:\r?\n|$)', note_text, re.S)
            check(frontmatter is not None and bool(re.search(r'^review_status:\s*reviewed\s*$', frontmatter.group(1), re.M)),
                  label + ': note frontmatter must explicitly say review_status: reviewed.')
        check(isinstance(rs.get('reviewer'), str) and bool(rs['reviewer'].strip()), label + ': reviewer is required.')
        try:
            datetime.fromisoformat(rs.get('reviewed_at', '').replace('Z', '+00:00'))
        except (TypeError, ValueError):
            error(label + ': reviewed_at must be ISO 8601.')
        check(all(isinstance(v, str) and v.strip() for v in rs.get('uncertainties', [])), label + ': uncertainties must contain nonempty strings.')
        figures = rs.get('figure_interpretations', [])
        check(bool(figures), label + ': explicitly record vision, ocr_only or unavailable.')
        vision_assets = set()
        ocr_assets = set()
        for f in figures:
            method = f.get('method')
            check(method in ('vision', 'ocr_only', 'unavailable'), label + ': invalid figure method.')
            check(isinstance(f.get('description'), str) and bool(f['description'].strip()), label + ': figure description is required.')
            check(isinstance(f.get('uncertainties'), list), label + ': figure uncertainties array is required.')
            check(isinstance(f.get('assessed_by'), str) and bool(f['assessed_by'].strip()), label + ': figure assessed_by is required.')
            check(f.get('scope') in ('whole_slide', 'crop', 'asset'), label + ': invalid figure scope.')
            if method in ('vision', 'ocr_only'):
                a = f.get('asset_path')
                check(a in available, label + f': figure is not mapped to this slide: {a}')
                if a in available:
                    existing(a)
                    (vision_assets if method == 'vision' else ocr_assets).add(a)
            else:
                check(f.get('asset_path') is None or f.get('asset_path') in available, label + ': unavailable figure has invalid asset path.')
                check(bool(f.get('uncertainties')), label + ': unavailable figure requires an uncertainty.')
                warnings.append(label + ': visual interpretation unavailable.')
        if not rs.get('claims'):
            check(bool(rs.get('uncertainties')), label + ': no claims; record why no knowledge was extracted.')

        def evidence(ev):
            kind, locator = ev.get('kind'), ev.get('locator')
            check(kind in ('pptx_text', 'speaker_notes', 'vision', 'ocr'), label + ': invalid evidence kind.')
            if kind in ('pptx_text', 'speaker_notes'):
                record = records.get(locator)
                check(record is not None, label + f': unknown text evidence locator: {locator}')
                if record:
                    check((record['origin'] == 'speaker_notes') == (kind == 'speaker_notes'), label + ': notes/text evidence kind mismatch.')
                    q = ev.get('quote')
                    check(isinstance(q, str) and bool(q) and q in record['text'], label + ': evidence quote must exactly match its source text record.')
            elif kind in ('vision', 'ocr'):
                check(locator in (vision_assets if kind == 'vision' else ocr_assets | vision_assets), label + ': image evidence lacks corresponding actual interpretation record.')
                check(isinstance(ev.get('region'), str) and bool(ev['region'].strip()), label + ': image evidence region is required.')
                if kind == 'ocr':
                    check(isinstance(ev.get('quote'), str) and bool(ev['quote']), label + ': OCR quote is required.')

        def number(item, claim_text=None):
            verbatim = item.get('verbatim')
            check(isinstance(verbatim, str) and bool(NUMERIC.search(verbatim)), label + ': numeric verbatim is required.')
            check(item.get('checked') is True, label + ': numeric comparison not marked checked.')
            for key in ('value', 'unit', 'conditions', 'source_locator'):
                check(isinstance(item.get(key), str) and bool(item[key].strip()), label + f': numeric {key} is required (use 미명시/해당없음 explicitly).')
            if claim_text is not None:
                check(isinstance(verbatim, str) and verbatim in claim_text, label + ': numeric verbatim is absent from claim text.')
            loc = item.get('source_locator')
            if loc in records:
                check(isinstance(verbatim, str) and verbatim in records[loc]['text'], label + ': numeric verbatim differs from source text.')
            else:
                check(loc in vision_assets, label + ': numeric source is neither exact text nor actual vision evidence.')
        for item in rs.get('numbers', []):
            number(item)
        for c in rs.get('claims', []):
            cid, txt = c.get('id'), c.get('text')
            check(isinstance(cid, str) and bool(cid) and cid not in all_claim_ids, label + ': claim id must be globally unique.')
            if isinstance(cid, str):
                all_claim_ids.add(cid)
                if require_reviewed:
                    check(cid in note_text, label + f': claim ID missing from Markdown note: {cid}')
            check(isinstance(txt, str) and bool(txt.strip()), label + ': claim text is required.')
            ct = c.get('claim_type')
            check(ct in ('text', 'visual', 'inference'), label + ': invalid claim_type.')
            evs = c.get('evidence', [])
            check(isinstance(evs, list) and bool(evs), label + ': each claim requires evidence.')
            for ev in evs:
                evidence(ev)
            if ct == 'visual':
                check(any(ev.get('kind') == 'vision' for ev in evs), label + ': visual claims require vision; OCR alone is insufficient.')
            if ct == 'inference':
                check(isinstance(txt, str) and '[추론]' in txt and bool(rs.get('uncertainties')), label + ': inference requires label and uncertainty.')
            cn = c.get('numbers', [])
            check(isinstance(cn, list), label + ': claim numbers must be an array.')
            for item in cn:
                number(item, txt)
            if isinstance(txt, str):
                for token in NUMERIC.findall(txt):
                    check(any(token in item.get('verbatim', '') for item in cn), label + f': unchecked numerical token in claim: {token}')
    # Only vault-relative links are accepted in generated source notes. External knowledge
    # links may point to existing Vault files but ambiguous basename-only links are rejected.
    vault = require_local_path(manifest.get('vault_path', '')).resolve()
    origin_rel = manifest.get('run_vault_relative', '')
    published_rel = f'DRAM/Sources/{manifest.get("run_id", "")}'
    for s in slide_entries:
        p = under(run, s['note_path'])
        if not p.is_file():
            continue
        contents = strip_code(p.read_text(encoding='utf-8-sig'))
        for target in re.findall(r'!?\[\[([^\]\n]+)\]\]', contents):
            target = target.split('|', 1)[0].split('#', 1)[0]
            if not target:
                continue
            try:
                if target.startswith(origin_rel + '/'):
                    dest = under(run, target[len(origin_rel) + 1:])
                elif target.startswith(published_rel + '/'):
                    dest = under(run, target[len(published_rel) + 1:])
                else:
                    dest = under(vault, target)
                check(dest.is_file() or (not dest.suffix and dest.with_suffix('.md').is_file()), f'Broken/ambiguous wiki link in {s["note_path"]}: {target}')
            except PipelineError as e:
                error(str(e))
    return {'ok': not errors, 'run': str(run), 'require_reviewed': require_reviewed, 'slide_count': len(slide_entries),
            'reviewed_slide_count': reviewed, 'errors': errors, 'warnings': warnings,
            'scope': 'Mechanical consistency only; review declarations are not proof of semantic accuracy or actual vision execution.'}


def publish(run):
    run = require_local_path(run).resolve(strict=True)
    result = validate(run, require_reviewed=True)
    if not result['ok']:
        raise PipelineError('Publication blocked: ' + '; '.join(result['errors']))
    m = read_json(run / 'manifest.json')
    vault = require_local_path(m['vault_path']).resolve(strict=True)
    expected_run = under(vault, m['run_vault_relative'])
    if run != expected_run:
        raise PipelineError('Publish only from the original staging path in this Vault.')
    destination_rel = f'DRAM/Sources/{safe_relative(m["run_id"])}'
    destination = under(vault, destination_rel)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise PipelineError(f'Publication already exists; nothing overwritten: {destination}')
    # Reject any user-added symlink/junction before copytree follows it.
    for path in run.rglob('*'):
        if is_link_or_reparse(path):
            raise PipelineError(f'Cannot publish a linked path: {path}')
    # Only a complete collection is moved into the searchable Sources directory.
    building = under(vault, f'_dram/publishing/{m["run_id"]}_{uuid.uuid4().hex[:12]}')
    building.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(run, building, dirs_exist_ok=False)
    for s in m['slides']:
        p = under(building, s['note_path'])
        contents = p.read_text(encoding='utf-8-sig')
        contents = contents.replace(m['run_vault_relative'] + '/', destination_rel + '/')
        # These are fresh copies owned by this publication; source staging is unchanged.
        p.write_text(contents, encoding='utf-8', newline='\n')
    index = ['---\ntype: dram-source-deck\nreview_status: records-complete\nsemantic_accuracy: not-automatically-verified\n---\n',
             f'# {m["source"]["original_name"]}\n', f'- 원본 SHA256: `{m["source"]["sha256"]}`',
             f'- 원본: [[{destination_rel}/source/original.pptx]]',
             '- 모든 슬라이드의 검토 기록이 있습니다. 의미 정확성/완전성은 자동 검증하지 않았습니다.',
             f'- 검토 기록: [[{destination_rel}/review.json]]\n', '## 슬라이드\n']
    for s in m['slides']:
        index.append(f'- [[{destination_rel}/{s["note_path"]}|슬라이드 {s["slide_no"]:04d}]]' + (' (숨김)' if s['hidden'] else ''))
    write_new(building / 'INDEX.md', '\n'.join(index) + '\n')
    dump_json(building / 'publication.json', {'published_at': utc_now(), 'source_staging': str(run),
              'vault_relative': destination_rel, 'validation': result, 'source_sha256': m['source']['sha256']})
    check_copy = validate(building, require_reviewed=True)
    if not check_copy['ok']:
        raise PipelineError('Prepared publication failed verification; preserved outside Sources: ' + str(building))
    if destination.exists():
        raise PipelineError('Publication destination appeared during copying; prepared copy retained: ' + str(building))
    # On supported Windows filesystems os.rename fails if the target exists.
    os.rename(building, destination)
    return {'published': str(destination), 'index': str(destination / 'INDEX.md'),
            'staging_preserved': str(run), 'semantic_accuracy': 'not automatically verified'}


def strip_code(text):
    """Do not treat literal source text in fenced/inline code as live wiki links."""
    lines = []
    marker = None
    count = 0
    for line in text.splitlines():
        match = re.match(r'^ {0,3}(`{3,}|~{3,})', line)
        if marker:
            if match and match.group(1)[0] == marker and len(match.group(1)) >= count:
                marker = None
            continue
        if match:
            marker, count = match.group(1)[0], len(match.group(1))
            continue
        lines.append(re.sub(r'(`+).*?\1', '', line))
    return '\n'.join(lines)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(errors='backslashreplace')
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('extract', help='Create a new immutable evidence staging run; no OCR/Vision.')
    p.add_argument('--source', required=True)
    p.add_argument('--vault', required=True)
    p = sub.add_parser('validate', help='Check hashes, mappings, citations, links and review declarations.')
    p.add_argument('--run', required=True)
    p.add_argument('--require-reviewed', action='store_true')
    p = sub.add_parser('publish', help='Require all reviews; copy to a new DRAM/Sources directory.')
    p.add_argument('--run', required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == 'extract':
            result = extract(args.source, args.vault)
        elif args.command == 'validate':
            result = validate(args.run, args.require_reviewed)
        else:
            result = publish(args.run)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get('ok', True) else 2
    except (PipelineError, OSError, KeyError, TypeError, ValueError, zipfile.BadZipFile) as e:
        print(json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
