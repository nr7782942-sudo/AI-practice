# 공식 문서 확인과 Windows 파일 위치

확인일: **2026-10-25 (Asia/Seoul)**. 아래는 Anthropic 공식 Claude Code 문서 기준입니다. 회사 배포판의 버전과 관리 정책이 우선이며, 설치/업데이트/연결 변경을 지시하는 문서가 아닙니다.

## 지침과 스킬

- 프로젝트 지침: `<프로젝트>\CLAUDE.md` 또는 `<프로젝트>\.claude\CLAUDE.md`. 이 패키지는 루트 `CLAUDE.md`를 사용합니다. 사용자 지침 기본 위치는 `%USERPROFILE%\.claude\CLAUDE.md`입니다. [Anthropic: Memory](https://code.claude.com/docs/en/memory)
- 프로젝트 스킬: `<프로젝트>\.claude\skills\<스킬명>\SKILL.md`. 사용자 스킬: `%USERPROFILE%\.claude\skills\<스킬명>\SKILL.md`. 스킬은 관련 작업에서 선택되거나 `/스킬명`으로 호출됩니다. 이 패키지는 프로젝트 스킬만 설치합니다. [Anthropic: Skills](https://code.claude.com/docs/en/skills)
- Windows의 `~/.claude` 기본 경로는 `%USERPROFILE%\.claude`입니다. `CLAUDE_CONFIG_DIR`가 설정되어 있으면 사용자 구성 디렉터리 기준이 달라집니다. 프로젝트 경로와 혼동하지 마세요. [Anthropic: The .claude directory](https://code.claude.com/docs/en/claude-directory)

## 설정 파일

- 사용자 설정: `%USERPROFILE%\.claude\settings.json`
- 프로젝트 공유 설정: `<프로젝트>\.claude\settings.json`
- 프로젝트 개인 설정: `<프로젝트>\.claude\settings.local.json`

회사 관리 설정이 이들보다 우선할 수 있습니다. 이 패키지는 위 파일을 생성/교체하지 않고 인증 키도 다루지 않습니다. 경로 확인은 설정 적용 범위를 이해하기 위한 것입니다. [Anthropic: Settings](https://code.claude.com/docs/en/settings)

## 실행 환경

최신 공식 문서의 `/skills`, `/reload-skills`, `/memory`, `/help` 명령은 사내 구버전/제한 환경에 없을 수 있습니다. 그때는 승인된 버전을 유지하고 재시작, `/` 메뉴, 스킬 파일 직접 읽기로 확인합니다. [Anthropic: Commands](https://code.claude.com/docs/en/commands)

CMD에서 `claude`를 실행하더라도 Claude가 명령을 실행하는 내부 셸이 CMD라는 뜻은 아닙니다. 현재 공식 Windows 안내는 Git for Windows가 있으면 Bash, 없으면 PowerShell 도구를 설명합니다. README의 사용자용 명령은 CMD 기준이며 모델은 실제 내부 셸에 맞게 인수를 인용해야 합니다. [Anthropic: Setup](https://code.claude.com/docs/en/setup)

`CLAUDE.md`는 모델에 제공하는 지침이며 보안 정책 강제 수단은 아닙니다. 이 패키지의 ‘로컬만 사용’ 규칙도 회사의 승인 모델·네트워크·저장 정책 안에서 실행되어야 합니다. [Anthropic: Memory](https://code.claude.com/docs/en/memory)

## PowerPoint 구현 근거

- 읽기 전용 문서 열기와 문서창 표시 여부는 `Presentations.Open` 인자로 지정합니다. [Microsoft: Presentations.Open](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)
- 슬라이드별 `Slide.Export`로 실제 순번과 PNG 파일명을 연결합니다. [Microsoft: Slide.Export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export)
- `AutomationSecurity`의 강제 매크로 비활성화 값을 사용하며 기존 값을 복구합니다. [Microsoft: AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.automationsecurity)
- 읽기 전용/매크로 차단만으로 외부 연결 갱신이 차단되지는 않습니다. 외부 관계를 사전 탐지하면 자동 COM 열기를 중단하는 것은 이 패키지의 보수적 구현 선택입니다. [Microsoft: LinkFormat.AutoUpdate](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.linkformat.autoupdate)

공식 문서는 파일 배치와 도구 API의 근거입니다. 사내 비표준 모델 게이트웨이의 실제 이미지 지원/저장 정책은 여기서 검증하지 않았습니다.
