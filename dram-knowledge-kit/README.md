# 사내 DRAM → Obsidian 정밀 지식 패키지

Windows CMD에서 사용하는 회사 내부 Claude Code용입니다. 원본 PPTX와 전체 추출 기록을 보존하고, 슬라이드별 근거 노트를 만든 뒤 검토한 자료만 Q&A에 사용합니다. 모델 재학습을 하거나 스킬만으로 환각을 없애는 패키지는 아닙니다.

**회사 자료나 API 키를 이 패키지 제작자/외부 서비스에 보내지 않아도 됩니다.** 포함 스크립트는 인터넷 호출·설치·인증 변경을 하지 않습니다. Claude의 기존 회사 승인 연결, Office 보안 정책, Vault의 저장/동기화 정책은 회사 환경에서 확인해야 합니다.

## 내일 할 최소 순서

### 1. 압축 해제와 준비

ZIP을 회사가 승인한 전달 방식으로 옮겨 로컬 폴더에 **모두 압축 해제**합니다. `.claude` 폴더도 필요합니다. 예: `C:\LocalTools\dram-knowledge-kit`. ZIP 안에서 바로 실행하지 않습니다.

준비물은 이미 회사에서 사용하는 Claude Code, 승인된 Python 3.10 이상, 로컬 Obsidian Vault입니다. PPT 자동 렌더에는 Windows 데스크톱 PowerPoint가 추가로 필요합니다. Python 패키지 설치는 없습니다. PowerPoint가 없으면 승인된 환경에서 내보낸 PNG를 사용할 수 있습니다.

현재 정상 동작하는 사내 Claude 환경의 프로젝트 경로와 Vault 경로를 정합니다. 아래 `C:\DRAM-Agent`, `D:\Obsidian\DRAM`은 예시이며 본인 경로로 바꿉니다. Vault 폴더는 먼저 만들어 둡니다.

**새 프로젝트를 쓴다면 기존 회사 모델 연결이 그 프로젝트에도 적용되는지 먼저 확인하세요.** 사내 설정이 이전 프로젝트의 `.claude/settings.json`에만 있었다면 새 폴더에서는 연결이 달라질 수 있습니다. 인증 설정을 임의 복사하거나 외부 기본 모델로 로그인하지 말고, 기존 승인 프로젝트 또는 회사가 구성한 새 프로젝트를 사용하세요.

### 2. 설치 (CMD)

```bat
cd /d "C:\LocalTools\dram-knowledge-kit"
install.cmd "C:\DRAM-Agent" "D:\Obsidian\DRAM"
```

설치기는 파일 충돌을 미리 검사합니다. 기존 `CLAUDE.md` 등 대상 파일이 있으면 **기존 파일을 덮어쓰지 않고 중단**합니다. 이 경우 기존 파일을 지우지 말고 아래 ‘기존 프로젝트와 병합’을 따르세요. 전역 `.claude`, `settings.json`, `settings.local.json`, API 키, 모델 주소, `.obsidian` 설정을 수정하지 않습니다.

Python 실행기가 없다고 나오면 회사가 승인한 Python 경로로 실행합니다.

```bat
"C:\ApprovedPython\python.exe" install.py --project "C:\DRAM-Agent" --vault "D:\Obsidian\DRAM"
```

미리 볼 때는 같은 Python 명령 끝에 `--dry-run`을 붙입니다. 실행 완료 후 프로젝트의 `INSTALL-RECEIPT.json`에서 설치 목록을 확인할 수 있습니다.

### 3. Claude 시작과 스킬 확인

```bat
cd /d "C:\DRAM-Agent"
claude --version
claude
```

`/memory`에서 프로젝트 `CLAUDE.md`가 읽혔는지 확인합니다. 최신 버전은 `/skills`로 목록을 볼 수 있습니다. 없으면 `/help` 또는 `/` 메뉴에서 확인하고 재시작합니다. 아래 스킬 이름이 인식되지 않는 사내 버전은 파일을 직접 읽도록 지시할 수 있습니다.

```text
이 프로젝트의 CLAUDE.md와 .claude/skills/dram-ppt-ingest/SKILL.md를 읽고 그 절차에 따라 진행해 줘.
```

### 4. PPT 한 개, 3~5장으로 먼저 검토

Claude 입력창에 아래를 붙입니다. 실제 PPT 경로는 **회사 PC 안에서만** 넣습니다.

```text
/dram-ppt-ingest "D:\ApprovedDocs\DRAM교육.pptx"
처음에는 원문 전체를 보존·추출하고, 텍스트/표/그림이 포함된 3~5장만 정밀 검토해 줘.
실제 이미지를 Vision 모델에 전달했는지 확인하고 OCR 결과와 구분해 줘.
슬라이드별 원문, 해석, 수치·단위·조건, 그림 연결, 미확인 항목을 보여 줘.
우선 staging에서 멈추고 내가 원본과 비교할 수 있게 해 줘.
```

원본 PPT와 노트의 수치·단위·각주·그림 방향/범례를 비교합니다. 특히 그림을 실제로 열었는지 확인하세요. 괜찮으면 같은 대화에서 아래처럼 이어갑니다.

```text
파일럿 결과를 확인했어. 같은 run의 나머지 슬라이드도 같은 기준으로 검토하고,
미확인 부분은 표시한 채 전체 검증을 거쳐 Vault에 공개해 줘.
```

### 5. 근거 기반 질문

```text
/dram-knowledge-qa 방금 적재한 문서에서 refresh와 retention을 어떻게 설명하는지,
측정 조건과 예외를 포함해서 알려 줘. 주장마다 문서·슬라이드 근거를 달고,
자료에 없는 내용은 미확인으로 남겨 줘.
```

내일은 **한 파일의 파일럿을 제대로 비교하는 것**을 첫 완료 기준으로 삼으세요. 전체 공개는 모든 슬라이드 검토가 끝난 뒤에 진행합니다.

## 설치되는 파일

```text
C:\DRAM-Agent\
  CLAUDE.md
  dram-config.json                    ← 설치기가 Vault 절대경로 기록
  INSTALL-RECEIPT.json                ← 설치 파일 목록/해시
  INSTALL-JOURNAL.jsonl               ← 설치 진행·부분 실패 추적
  .claude\skills\
    dram-ppt-ingest\SKILL.md
    dram-knowledge-qa\SKILL.md
  scripts\
    dram_pipeline.py                 ← 추출·기계 검증·새 공개본 생성
    render_powerpoint.ps1             ← 선택: 로컬 PowerPoint 렌더
    import_rendered.py               ← 선택: 수동 내보낸 PNG 연결
  docs\
    WORKFLOW.md
    REVIEWING.md
    OFFICIAL-SOURCES.md
    VALIDATION.md
```

원본·추출·노트·이미지는 `dram-config.json`에 적힌 Vault에 저장됩니다. 프로젝트 폴더와 Vault는 달라도 됩니다. Vault 바깥 접근이 차단된 사내 환경은 회사 승인 권한 절차를 따르세요. 스크립트가 Claude 권한을 자동으로 넓히지 않습니다.

## 기존 프로젝트와 병합

- 가장 간단한 방식은 회사 연결이 적용되는 새 전용 프로젝트에 설치하는 것입니다.
- 기존 프로젝트에 `CLAUDE.md`가 있으면 설치기는 병합하지 않습니다. 패키지의 지침을 기존 지침과 검토·병합하고, 기존 파일은 회사 승인 위치에 백업한 뒤 직접 반영하세요. `.claude/CLAUDE.md`, 상위 폴더 지침, 동명 사용자 스킬도 충돌 여부를 확인합니다.
- 프로젝트 스킬만 따로 복사하면 `scripts`, `docs`, `dram-config.json`이 빠질 수 있습니다. 부분 설치보다 전용 프로젝트를 권장합니다.
- 설치 중 디스크/권한 오류가 나면 새 파일 일부가 남을 수 있습니다. 기존 파일을 삭제해 재시도하지 말고 설치 로그/영수증 상태를 확인해 새 대상에 설치하세요. 원본 Vault 콘텐츠는 설치 과정에서 변경하지 않습니다.
- 업데이트 때도 같은 폴더에 강제 덮어쓰지 않습니다. 새 프로젝트에 새 버전을 설치하고 검토 후 사용합니다. 프로젝트 위치가 달라지면 회사 연결 적용 범위를 다시 확인합니다.

## 중요한 범위와 한계

- 이 패키지는 `.pptx`를 처리합니다. `.ppt`, `.pptm`, 암호화/손상 문서는 회사 승인 PowerPoint에서 **새 PPTX 사본**으로 저장한 뒤 원본도 별도 보존하세요. 자동 매크로 실행/보호 해제는 하지 않습니다.
- 텍스트 추출은 문서의 의미를 완전히 읽는 작업이 아닙니다. 마스터/레이아웃, SmartArt, 차트, OLE, 수식, 애니메이션, 숨김/잘린 내용은 렌더와 원본 확인이 필요합니다. 원본 PPTX와 원본 패키지를 함께 보존하는 이유입니다.
- 포함 검증기는 경로·해시·슬라이드 매핑·리뷰 형식·일부 인용/수치 기록을 검사합니다. 누락된 모든 주장이나 잘못된 물리적 해석을 자동으로 판정하지는 못합니다.
- `vision`은 이미지를 실제로 모델이 본 기록이고, `ocr_only`는 글자 추출 기록입니다. `reviewed`는 검토 절차 수행 기록입니다. 어느 것도 무오류 인증이 아닙니다.
- 스킬은 지침입니다. 회사 모델/엔드포인트 제한과 외부 전송 차단은 회사 정책과 기술적 통제가 담당해야 합니다. 이 패키지는 그 설정을 대신하지 않습니다.

세부 처리: [WORKFLOW](docs/WORKFLOW.md) · 리뷰 작성: [REVIEWING](docs/REVIEWING.md) · 공식 경로 확인: [OFFICIAL-SOURCES](docs/OFFICIAL-SOURCES.md) · 실제 테스트 범위: [VALIDATION](docs/VALIDATION.md).
