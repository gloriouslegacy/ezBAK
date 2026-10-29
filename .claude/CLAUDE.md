# ezBAK

Windows 사용자 데이터 백업·복원 도구. tkinter GUI 와 Headless(예약 작업) 실행을 지원한다. 1인 개발.

## 명령

- 테스트: `python -m pytest tests -q` (마커 `-m gui` 는 Tk 필요, `-m windows` 는 파일 속성·ctypes 필요)
- 문법 확인: `python -m py_compile ezBAK.py`
- 실행: `python ezBAK.py` (GUI). Headless 인자는 ADR 0003 §6
- 릴리스 빌드: `.github/workflows/release.yml` 이 기준이다. 태그 `v*` 를 push 하면 PyInstaller onefile 로 `ezBAK.exe` 를 만든다.
- 로컬 빌드: `.\scripts\build_pyinstaller_script.ps1 -ScriptName ezBAK.py`
- 환경: Windows 10/11, Python 3.11. pywin32·psutil 은 선택이다.

## 문서

- 요구사항 `docs/requirements/refactoring.md` · 구조 색인 `docs/structure.md` · 결정 `docs/adr/` · 로드맵 `docs/roadmap.md`
- 모듈별 규칙은 각 폴더의 `CLAUDE.md` 에 있다(`ezbak/<모듈>/`, `tests/`, `scripts/`).

## 규칙

- 호환은 고정이다(C3): exe 이름, CLI 인자, settings·schedule JSON 키·형식, 예약 작업 인자, 백업 폴더·로그 이름, 작업 로그 문구. 목록은 ADR 0003 §6.
- 동작 동일(Q1): 특성 테스트 golden 은 승인된 Q1 예외 커밋에서만 바꾸고, 커밋 메시지에 사유를 적는다.
- 리팩터링 커밋에 동작 변경을 섞지 않는다. 결함 수정은 별도 커밋이다.
- 줄 수 상한: 파일 500, 함수 60(gui 배치 함수 150), `ezBAK.py` 100(ADR 0003).
- `tkinter` 는 `ezbak/gui/` 에서만, `subprocess` 는 `ezbak/winsys/proc.py` 에서만 import 한다. 의존 방향은 ADR 0002 를 따르며 `tests/test_architecture.py` 가 검사한다.
- 표준 라이브러리와 같은 이름의 모듈·패키지를 만들지 않는다(logging, platform 등).
- 로그: 작업 로그는 `ezbak/core/joblog.py`(현행 문구), 진단 로그는 `logging` 과 `run_id` 를 쓴다(ADR 0003).
- 테스트에서 실제 schtasks 등록, 드라이버 설치, UAC, 네트워크 접속을 하지 않는다. 외부 명령은 proc 모킹으로 argv 만 확인한다.
- 하지 않을 것: 프레임워크 교체(예외: UI 리디자인의 CustomTkinter, ADR 0004), 백업 알고리즘 변경(압축·증분)
- UI 리디자인(U 단계, ADR 0004): 동작 golden 은 그대로 통과해야 하고, 창 구조 golden 만 U4~U6 해당 커밋에서 Q1 예외로 다시 만든다. `customtkinter` 도 `ezbak/gui/` 에서만 import 한다.
- 배치: 문서는 `docs/`, 보조 스크립트는 `scripts/`, 테스트는 `tests/` 에 둔다. 루트 코드 파일은 `ezBAK.py` 하나뿐이다.
- 작업 브랜치: 세션마다 main 에서 새 `claude/…` 브랜치를 딴다(리팩터링 P1~P6·UI U0~U8 은 모두 main 에 병합, 옛 브랜치는 2026-10-27 삭제). 발주 1건마다 커밋하고, 단계가 끝나면 PR 로 main 에 병합한다. tests.yml·regen-golden.yml 은 `claude/**` push 에서 돈다.
