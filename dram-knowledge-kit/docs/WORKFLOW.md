# 원본부터 근거 기반 답변까지

## 데이터 구조와 불변 원칙

추출 명령은 Vault의 `_dram/staging/<run_id>/`에 새 run을 만듭니다. 문서 이름에 특수문자가 있어도 내부 식별자는 해시와 고유값을 사용합니다. 실제 경로는 명령 출력으로 확인하세요.

```text
<run>/
  source/original.pptx                 원본 바이트 사본
  source/package/...                  원본 OOXML 파트
  manifest.json                       원본 SHA-256·순번·관계·파일 해시
  slides/slide-0001.json               원문 텍스트·도형·관계·좌표
  slides/slide-0001.txt                검색용 원문
  assets/media/...                    원본 미디어와 슬라이드 관계
  assets/slides/slide-0001.png         선택: 전체 렌더
  render-manifest.json                선택: 렌더 해시·순번·방법
  notes/slide-0001.md                  작성 중 근거 노트
  review.json                         작성 중 검토 기록
  review.schema.json / REVIEW_GUIDE.md 해당 버전의 리뷰 규칙
```

`source`, 추출 `slides`, 원본 media와 manifest는 보존층입니다. 이를 고쳐 모델의 주장을 맞추지 않습니다. 노트와 리뷰는 해석층입니다. 전체 문서 요약은 검색 안내층이며 보존층/해석층을 대신하지 않습니다.

원본 파일명을 정렬해서 슬라이드 번호를 추정하지 않습니다. PPTX `presentation.xml`의 실제 슬라이드 목록과 관계를 따라 순번을 정합니다. 숨김 슬라이드도 포함합니다. 같은 그림을 여러 슬라이드가 재사용하면 파일은 공용이어도 **각 슬라이드 관계와 객체**로 출처를 구별합니다.

## A. 추출

CMD, 프로젝트 루트 기준:

```bat
py -3 scripts\dram_pipeline.py extract --source "D:\ApprovedDocs\DRAM교육.pptx" --vault "D:\Obsidian\DRAM"
```

출력된 run 절대경로를 이후 명령에 사용합니다. PPTX 원본 전체를 보존하므로 파일 크기가 커질 수 있습니다. 원문 패키지와 media 복사본도 있어 원본보다 디스크 사용량이 큽니다.

과도하게 큰 입력은 조용히 일부만 처리하지 않고 중단합니다. 기본 한도는 PPTX 512 MiB, 압축 해제 총량 1 GiB, 단일 파트 128 MiB, ZIP 항목 20,000개입니다. 한도에 걸리면 회사 승인 도구로 원본을 유지한 별도 분할 사본을 만든 뒤 처리합니다.

체크포인트 A: 해시, 슬라이드 개수/순서, 숨김 여부, 발표자 노트, 이미지 관계, 외부 관계, 미지원 객체 플래그를 확인합니다. 추출 TXT는 도형/테이블 구조의 원문이며 완성된 화면 읽기 순서가 아닙니다. 그룹 좌표는 로컬 좌표일 수 있고 회전/크롭/변환을 반영한 화면 bbox라고 간주하지 않습니다.

## B. 렌더 또는 수동 PNG

### 승인된 데스크톱 PowerPoint 사용

```bat
powershell -NoProfile -File scripts\render_powerpoint.ps1 -RunPath "D:\Obsidian\DRAM\_dram\staging\실제_run_id"
```

이 스크립트는 저장된 원본 사본을 읽기 전용으로 열고 실제 슬라이드 순번별 PNG를 만듭니다. 실행 중인 PowerPoint가 있으면 기존 세션 보호를 위해 중단하므로 문서를 저장하고 PowerPoint를 닫은 뒤 실행합니다. 매크로/활성 콘텐츠 또는 외부 관계가 발견되면 열기 전에 중단합니다. 단순 외부 하이퍼링크만 있어도 보수적으로 차단될 수 있습니다. 보호 보기 해제, 정책 우회, 전역 보안 설정 변경은 하지 않습니다. PowerPoint가 없거나 회사가 스크립트 실행을 막으면 승인된 서명/실행 절차 또는 아래 수동 경로를 사용합니다.

PowerPoint의 추가 기능, 폰트, 기업 보안 정책에 따른 동작까지 이 스크립트가 격리하지는 못합니다. 승인된 Office 환경에서 사용하며 네트워크 격리가 필요하면 회사 정책으로 적용해야 합니다.

### 회사 승인 방식으로 수동 PNG 내보내기

1. 원본 사본을 회사 승인 PowerPoint에서 열고, 외부 콘텐츠 업데이트를 허용하지 않은 상태로 모든 슬라이드를 PNG로 내보냅니다. 숨김 슬라이드도 출력되었는지 확인합니다. 사용할 수 없는 폰트/개체는 미확인으로 기록합니다.
2. `Slide1.PNG`, `Slide2.PNG` 또는 `슬라이드1.PNG`, `슬라이드2.PNG`처럼 **실제 순번에 대응하는 파일명 규칙**을 확인합니다. 이름 정렬 순서를 근거로 삼지 않습니다.
3. 첫/마지막뿐 아니라 숨김 슬라이드 전후와 복잡한 도표도 원본과 대응 확인한 뒤 가져옵니다. 아래 `{n}`은 문자 그대로 입력합니다.

```bat
py -3 scripts\import_rendered.py --run "D:\Obsidian\DRAM\_dram\staging\실제_run_id" --images "D:\LocalExports\DRAM교육" --pattern "슬라이드{n}.PNG" --mapping-confirmed
```

파일명이 영문이면 `--pattern "Slide{n}.PNG"`를 사용합니다. 별도의 개별 PNG를 정확한 `slide-0001.png` 이름으로 준비했다면 `--pattern "slide-{n:04d}.png"`를 사용합니다. 가져오기 스크립트는 파일 수/번호/PNG 서명·크기·해시를 검사하지만 PNG 내용과 원본의 시각적 일치까지 판정하지 않습니다.

기존 PNG 또는 렌더 manifest가 있으면 덮어쓰지 않습니다. 잘못 연결했다면 새 추출 run에서 다시 작업합니다.

체크포인트 B: Claude가 실제 이미지 입력 도구로 PNG를 열었는지 확인합니다. 파일 경로만 읽기, OCR 문자열 반환, 텍스트 전용 모델 응답은 Vision 확인이 아닙니다. 불가능하면 `unavailable`로 남기고 중요한 도표의 해석은 보류합니다.

## C. 노트와 검토 기록

각 슬라이드의 전체 원문, 발표자 노트, 그림, 조건/각주를 따라 읽어 `notes/slide-NNNN.md`와 `review.json`을 작성합니다. 세부 규칙은 [REVIEWING](REVIEWING.md)을 따릅니다. 요약 때문에 상세 정보가 사라지는 것을 막기 위해 원문을 삭제하지 않고 보존층 링크를 남깁니다.

체크포인트 C: 처음 한 문서의 3~5장을 사람과 비교합니다. 오류가 발견되면 해당 관찰/수치/조건을 고친 뒤 나머지 문서에 적용합니다. 자동으로 모든 장을 `reviewed`로 바꾸지 않습니다.

## D. 기계 검사와 공개

```bat
py -3 scripts\dram_pipeline.py validate --run "실제_run_절대경로"
py -3 scripts\dram_pipeline.py validate --run "실제_run_절대경로" --require-reviewed
py -3 scripts\dram_pipeline.py publish --run "실제_run_절대경로"
```

첫 명령은 구조 점검, 두 번째는 전체 검토 기록을 요구하는 점검입니다. 마지막 명령도 공개 전에 검토 조건을 확인합니다. 성공하면 Vault의 `DRAM/Sources/<run_id>/`에 새 복사본과 인덱스가 생깁니다. staging은 남습니다. 같은 대상이 있으면 공개를 중단합니다.

체크포인트 D: 출판된 인덱스에서 슬라이드 노트, 원문 파일, 전체 이미지 링크를 열어 봅니다. 공개 후 Q&A 검색 범위는 `DRAM/Sources/`로 제한합니다. 자료가 새 버전이면 새 run으로 공개하고 문서 버전/대상 조건을 함께 비교합니다.

## Obsidian 링크 규칙

Vault 루트 기준 경로, `/` 구분자, 파일 존재 확인을 원칙으로 합니다. 생성된 실제 run_id로 바꿉니다.

```md
[[DRAM/Sources/<run_id>/notes/slide-0001|문서명 · 슬라이드 1]]
[[DRAM/Sources/<run_id>/slides/slide-0001.txt|원문 전문]]
![[DRAM/Sources/<run_id>/assets/slides/slide-0001.png]]
```

텍스트가 없는 도표도 빈 페이지로 버리지 않습니다. 이미지 접근 여부와 관찰/미확인을 남깁니다. 임의로 원시 그림을 순서대로 붙이지 않고 해당 슬라이드 관계/설명과 연결합니다. 표/수식/벡터 그림은 원시 media 파일이 없어도 전체 슬라이드 렌더로 문맥을 확인합니다.
