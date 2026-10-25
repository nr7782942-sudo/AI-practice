#!/usr/bin/env python3
"""Register manually exported PNGs against explicit slide ordinals; no OCR/Vision."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import sys
import uuid

from dram_pipeline import PipelineError, dump_json, png_dimensions, read_json, sha_file, under, utc_now, validate


def local_path(value, must_exist=True):
    p = Path(os.path.abspath(value))
    if str(p).startswith(('\\\\', '//')):
        raise PipelineError('Use an approved local path, not a UNC path.')
    for item in [p, *p.parents]:
        if item.exists() or item.is_symlink():
            st = item.lstat()
            if item.is_symlink() or getattr(st, 'st_file_attributes', 0) & 0x400:
                raise PipelineError(f'Symlink/reparse path not supported: {item}')
    if os.name == 'nt':
        import ctypes
        if ctypes.windll.kernel32.GetDriveTypeW(str(p.anchor)) == 4:
            raise PipelineError('Mapped network drives are not supported.')
    return p.resolve(strict=must_exist)


def import_images(run_path, images_path, pattern, mapping_confirmed):
    if not mapping_confirmed:
        raise PipelineError('First verify image/slide mapping, including hidden slides, then use --mapping-confirmed.')
    run = local_path(run_path)
    images = local_path(images_path)
    if not run.is_dir() or not images.is_dir():
        raise PipelineError('Run and images must be existing directories.')
    if '{n' not in pattern:
        raise PipelineError('Pattern must contain {n} or {n:04d}; no lexicographic guessing.')
    report = validate(run)
    if not report['ok']:
        raise PipelineError('Run validation failed: ' + '; '.join(report['errors']))
    manifest = read_json(run / 'manifest.json')
    if run != local_path(manifest['vault_path']) / Path(manifest['run_vault_relative']):
        raise PipelineError('Import into the original staging run only.')
    count = manifest['slide_count']
    render_manifest_path = under(run, 'render-manifest.json')
    if render_manifest_path.exists():
        raise PipelineError('Render manifest already exists; no files overwritten.')
    source_hash = sha_file(under(run, manifest['source']['path']))
    if source_hash != manifest['source']['sha256']:
        raise PipelineError('Original source hash mismatch.')
    plans = []
    names = set()
    for n in range(1, count + 1):
        try:
            name = pattern.format(n=n)
        except (KeyError, ValueError, IndexError) as e:
            raise PipelineError(f'Invalid pattern: {e}') from e
        if '/' in name or '\\' in name or Path(name).suffix.lower() != '.png':
            raise PipelineError('Pattern must produce PNG basenames, not paths.')
        image_path = under(images, name)
        local_path(image_path)
        if name.casefold() in names:
            raise PipelineError('Pattern maps multiple slides to one filename.')
        names.add(name.casefold())
        dest_rel = f'assets/slides/slide-{n:04d}.png'
        if under(run, dest_rel).exists():
            raise PipelineError(f'Render already exists: {dest_rel}')
        width, height = png_dimensions(image_path)
        plans.append((image_path, {'slide_no': n, 'path': dest_rel,
                                  'sha256': sha_file(image_path), 'width': width, 'height': height,
                                  'original_export_name': name}))
    actual_names = {p.name.casefold() for p in images.iterdir() if p.is_file() and p.suffix.lower() == '.png'}
    if actual_names != names:
        raise PipelineError('PNG files do not exactly match slide count/pattern. Use a dedicated export directory.')
    temporary = run / ('manual-render-prepared-' + uuid.uuid4().hex)
    temporary.mkdir()
    for src, item in plans:
        target = temporary / Path(item['path']).name
        with src.open('rb') as inp, target.open('xb') as out:
            shutil.copyfileobj(inp, out)
        if sha_file(target) != item['sha256']:
            raise PipelineError('Export image changed during copy; prepared files retained for diagnosis.')
    # Fail closed on partial movement: no success manifest means no valid vision evidence.
    for _, item in plans:
        target = under(run, item['path'])
        target.parent.mkdir(parents=True, exist_ok=True)
        with (temporary / target.name).open('rb') as inp, target.open('xb') as out:
            shutil.copyfileobj(inp, out)
    metadata = {'schema_version': 1, 'source_sha256': source_hash,
                'method': 'manual_powerpoint_export', 'mapping_confirmed': True,
                'created_at': utc_now(), 'renders': [item for _, item in plans],
                'limitations': 'User-declared slide mapping; no automated semantic or image-content verification.'}
    dump_json(render_manifest_path, metadata)
    final_report = validate(run)
    if not final_report['ok']:
        raise PipelineError('Imported artifacts did not validate: ' + '; '.join(final_report['errors']))
    return {'ok': True, 'run': str(run), 'registered_slides': count,
            'method': metadata['method'], 'prepared_copy': str(temporary),
            'next': 'Actually open each image through approved Vision input; import alone is not image interpretation.'}


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(errors='backslashreplace')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', required=True)
    parser.add_argument('--images', required=True)
    parser.add_argument('--pattern', default='Slide{n}.PNG')
    parser.add_argument('--mapping-confirmed', action='store_true')
    args = parser.parse_args()
    try:
        result = import_images(args.run, args.images, args.pattern, args.mapping_confirmed)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (PipelineError, OSError, KeyError, TypeError, ValueError) as e:
        print(json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
