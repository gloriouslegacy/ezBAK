| 일자 | 에이전트·발주명 | 오류 | 원인 | 수정 내용 |
|---|---|---|---|---|
| 2026-10-27 | gate-auditor · gate_log_1027 | self_protect 오탐 후보 8건(`why: unresolved`, 정규식 인자만 걸린 grep 등) | 경로를 못 푼 인자를 보수적으로 차단 | 미정 — `unresolved` 판정을 cmd 단위로 재검토 제안(claudeOrchestration 패키지 쪽) |
| 2026-10-27 | gate-auditor · gate_log_1027 | self_protect 오탐 후보 1건(`why: fable`, `git update-index --chmod=+x -- package/*.sh`) | 경로 글롭이 든 일반 git 명령을 보호 경로로 판정 | 미정 — 확정 전 cmd 전문 확인 필요 |
| 2026-10-27 | gate-auditor · gate_log_1027 | dispatch_log 규칙 공백 WARN 122건("meta.json toolUseId 없음") | Claude Code meta.json 형식과 대장 종료 연결 불일치(추정) | 미정 — 키 이름 점검, 대체 필드 연결이나 세션당 1회 경고 제안 |
| 2026-10-27 | runner · winbuild | WSL→Windows `powershell.exe` 호출(읽기 전용 레지스트리 조회 포함)·`git archive` 가 runner·메인 모두 거부됨 | 설치기 규칙이 경로를 못 푼 `.exe` 인자를 보수적으로 차단(추정) | 미정 — 실기기 확인은 사용자 PowerShell 로 돌림, 게이트 규칙 점검은 claudeOrchestration 쪽 |
