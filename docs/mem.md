# 세션 메모 (갱신 2026-10-29 · 커밋 7e583b9 기준 · v0.10.2 재릴리스)
## 끝낸 것
- 리팩터링 T0·P1~P6, UI U0~U8 완료 → v0.10.0~v0.10.2(UAC 뒤 검은 창 수정, Backup Options 두 줄 배치 spec ui-fixes F1)
- 2026-10-29 저장소 공개 전 민감정보 점검: 비밀·개인정보 0, 커밋 메타의 개인 Gmail 만 노출 → 저장소 삭제·재생성(단일 커밋 bc5c4a6, noreply)
- 옛 이력(PR #45~#52 포함)·릴리스 5개 백업: ~/work/ezBAK-backup/(ezBAK-old.git 미러, releases/). 옛 커밋 해시·PR 번호는 새 저장소에 없음
- pytest 9.2.0 monkeypatch undo 변경으로 CI 173 실패 → requirements-dev `pytest>=8.0,<9.2`(PR #1, 7e583b9). v0.10.2 태그 이동 후 재빌드 성공
- GitHub 이메일 비공개·노출 push 차단 켬, WSL·Windows git user.email noreply
## 열린 것
- 사용자 Windows 실기기: `docs/nav-drafts/u8-checklist.md`(다른 화면 잘림, 배율 125·150%×테마, 한·영, 모니터 이동, 홈 카드). 본 것: 홈 Free Space `-`, 밝은 테마 진행 막대 검정
- 결함은 docs/spec/ui-fixes.md 에 F 로 적고 v0.10.x 패치
- 하네스를 pytest 9.2 에 맞게(facade 대신 실제 모듈 패치, App 은 gui_harness.override) → 이후 상한 해제. 별도 implementer 발주
- 사용자 판단 대기: `ezbak/` 안내 부채 기한 지남(`fable debt keep`), 기계적 발주 implementer Sonnet 제안
- 미커밋: 이 메모 — 다음 `claude/…` 브랜치에서 커밋
## 함정
- conftest 는 mainloop 를 안 돌려 창 표시·배치 결함을 못 잡는다 → 자식 Python 에서 mainloop(test_gui_window_shown 방식)
- golden 재생성: `Regen-Golden: <파일>` 는 head 커밋에만. 게이트가 경로 변수에 `claude` 가 보이면 막음 → scratchpad 는 리터럴 경로
- 게이트가 WSL→Windows `powershell.exe`·일부 `gh … -q` jq 식을 막음 → 출력 저장 후 grep. 사용자 PC customtkinter 6.0.0(고정 5.2.2)
- 커밋·PR 에 Claude-Session 링크를 넣지 않는다(공개 저장소, 사용자 정리 대상). ezBAK.py 는 Edit 만·CRLF. CLAUDE.local.md 는 사용자가 커밋
