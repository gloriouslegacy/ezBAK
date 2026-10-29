# 세션 메모 (갱신 2026-10-27 · 커밋 1dd3fee 기준 · v0.10.2 릴리스)
## 끝낸 것
- 리팩터링 T0·P1~P6, UI U0~U8·U7 main 병합(PR #45~#49). v0.10.0 릴리스(5bb3bb3, 실기기 확인 없이)
- v0.10.1(d805b00, PR #50): UAC 뒤 검은 창. app.py 시작 withdraw 가 CTk `_withdraw_called_before_window_exists` 를 켜 mainloop 가 deiconify 생략 → `tk.Tk.withdraw(self)` + tests/test_gui_window_shown.py
- v0.10.2(1dd3fee, PR #51): Backup Options 가 1000px 에서 System 잘림 → 두 줄 배치(spec docs/spec/ui-fixes.md F1) + tests/test_gui_options_fit.py, golden tree_* 4개 봇 재생성. 실기기 확인함
- 각 태그 전 게이트 시나리오 FAIL 0(최근 PASS 1772/SKIP 9). main 외 브랜치 모두 삭제(미병합 feature/ui-redesign 9b0346f 포함, 사용자 결정)
- 게이트 로그 분류는 docs/issue/agent-issue.md(오탐·공백 4줄, claudeOrchestration 쪽 조치)
## 열린 것 (사용자, Windows 실기기)
- 릴리스 exe(v0.10.2)로 `docs/nav-drafts/u8-checklist.md`: 다른 화면 잘림(Restore·Scheduled·Tools·Settings), 배율 125·150%×테마, 한·영, 모니터 이동, 6절 홈 카드. 본 것: 홈 Free Space `-`, 밝은 테마인데 진행 막대가 검게 보임
- 결함은 docs/spec/ui-fixes.md 에 F 로 적고 v0.10.x 패치
- 사용자 판단 대기: `ezbak/` 안내 부채 기한 지남(`fable debt keep`), 기계적 발주 implementer Sonnet 제안
- 미커밋: 이 메모 — 다음 `claude/…` 브랜치에서 커밋
## 함정
- conftest 는 mainloop 를 안 돌려 창 표시·배치 결함을 못 잡는다 → 자식 Python 에서 mainloop(test_gui_window_shown 방식)
- golden 재생성: 커밋 메시지 `Regen-Golden: <파일>` 는 head 커밋에만. 봇 push 는 Tests 를 다시 안 돌리지만 regen 워크플로가 전체 테스트를 돌린다
- 게이트가 WSL→Windows `powershell.exe` 와 일부 `gh … -q` jq 식을 막음 → 출력 저장 후 grep. 사용자 PC customtkinter 6.0.0(고정 5.2.2)
- 로그 기본 INFO 라 시작 단계 로그는 비는 게 정상(`EZBAK_LOG_LEVEL=DEBUG`). ezBAK.py 는 Edit 만·CRLF. CLAUDE.local.md 는 사용자가 커밋
