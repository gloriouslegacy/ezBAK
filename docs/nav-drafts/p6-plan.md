# P6 계획서 — 로깅·빌드 (2026-10-27, 기준 main 68611dd · U8-1~4 병합 뒤)

따를 것: ADR 0003 §2(로깅)·§3(검사 d·f), ADR 0002(core → winsys), 루트·모듈 CLAUDE.md. roadmap P6 행: 진단 로그 전환(`%APPDATA%\ezBAK\logs`), 검사 d 켜기, spec·로컬 빌드를 `release.yml` 에 맞춤, `release.yml` 에서 빌드 전에 pytest, 기술문서 갱신.
U8 실기기 체크리스트(`docs/nav-drafts/u8-checklist.md`)는 P6 와 따로 간다. 발견이 나오면 그 수정은 P6 사이의 별도 커밋이다.

## 0. 지금 코드에서 확인한 사실
- 진단 로그 핸들러가 어디에도 없다(`basicConfig`·`addHandler` 0건). `_log.debug(...)`(P5-fix #16~18)·`_LOG`(engine·retention·storage.local·bridge) 줄은 지금 모두 버려진다. `core/diag.py` 없음.
- `print(` 은 `ezbak/` 에 139곳: gui 120곳(sources 35·dialogs/schedule 29·actions_schedule 28·sources_tree 15·bridge 4·shortcuts 3·settings 2·app_log 2·i18n 1·schedule_layout 1), automation/headless 18곳(C3 stdout, 그대로), 루트 ezBAK.py 2곳(관리자 재실행 실패 안내, 비frozen 콘솔만).
  - gui print 는 대부분 `DEBUG:` 이고, 아닌 것 5곳: `ERROR: …`(actions_schedule 3·dialogs/schedule 1)·`Task created successfully: …`(actions_schedule 1).
  - golden 이 stdout 을 담는 곳: schedule_cases(DEBUG 줄 75곳), tool_cases(get_system_info 경로 26곳), gui_snapshot `settings_*` 의 stdout, gui_cases 의 out(traceback 접기). smoke `test_enable_buttons` 는 `"DEBUG: Buttons enabled"` 를 stdout 에서 찾는다. `test_window_tree` 는 "DEBUG 없음"을 단언한다.
- 맨 `except:` 는 `ezbak/` 에 0곳(P5-fix #15). 검사 d 의 절반은 이미 지켜진다.
- 작업 스레드가 `progress_bar['maximum']` 을 읽는 finally 4곳(h-w): actions_backup 180·actions_restore 98·actions_tools 93·302.
- 버려지는 줄: NAS `net use` out·err(g-j, storage.nas → wrapper 가 버림), winget simple list rc≠0(g-l, 로그 없음).
- 빌드: release.yml 은 `--add-data "./icon;./icon" --name "ezBAK" --clean --version-file version_info.txt`(태그에서 생성), 로컬 ps1 은 `--add-data "./icon/ezbak_title.ico;./icon"`·`--version-file bak_version_info.txt`(저장소에 있음)·`--name`·`--clean` 없음. `.spec` 파일은 없다(ps1 이 지운다). release.yml 에 pytest 단계 없음, `actions/checkout@v3`·`setup-python@v4`.
- 정리 후보(호출처 다시 확인 필요): `winsys.paths.local_appdata()`, 모듈 함수 `log_message`, App 메서드 `_validate_time_input`·`on_exit`·`set_progress_indeterminate`·`switch_mode`(roadmap "P2 에서 찾은 것"). 로거 이름 `_LOG` 4파일(나머지 21파일은 `_log`).

## 1. 진단 로그 설계 (ADR 0003 §2 그대로)
- `ezbak/core/diag.py`: `setup(run_id=None, *, log_dir=None, level=None)` — 루트 `ezbak` 로거에 `RotatingFileHandler(<dir>\ezbak.log, maxBytes=1MB, backupCount=3, encoding='utf-8')` 하나, 형식 `%(asctime)s %(levelname)-5s %(name)s [%(run_id)s] %(message)s`, 레벨 INFO(`EZBAK_LOG_LEVEL=DEBUG` 로 상향). 두 번 불러도 핸들러는 하나. 폴더를 못 만들면 핸들러 없이 돌아간다(진단 실패가 백업을 막지 않음, 경고는 stderr 한 줄).
  - `new_run_id(clock)` = `YYYYMMDD-HHMMSS-xxxx`(4자리 16진), `bind(run_id)`·`current()` 는 `contextvars`. 필터가 레코드에 `run_id` 를 넣는다(없으면 `-`).
  - 위치 `%APPDATA%\ezBAK\logs`: `winsys.paths.appdata_dir()`(core → winsys paths 허용 간선). APPDATA 가 비면 핸들러 없음.
  - 로거는 `ezbak` 아래만(`propagate` 기본). 루트 로거·다른 라이브러리는 건드리지 않는다.
- 부르는 곳: ezBAK.py `__main__` 맨 앞 한 번(Headless·GUI 공통). 작업 시작(GUI 백업·복원·복사·도구·예약 생성·삭제, Headless 백업)에 `new_run_id` + `bind`, 작업 스레드 첫 줄에서 다시 `bind`. 시작 진단 줄 `event=job_start kind=backup job_log=<경로>`(작업 로그 본문에는 넣지 않음).
- 테스트: 실제 %APPDATA% 를 쓰지 않는다(`log_dir` 인자·isolated_env). setup 은 테스트 세션에서 부르지 않고 caplog 로 확인한다.

## 2. 발주 목록 (발주 1건 = 커밋 1개, Files 5개 이하 golden 제외, ★ = golden 재생성 Q1 예외)
| # | Goal | Files | 끝나는 기준 |
|---|---|---|---|
| P6-1 | `core.diag`(setup·new_run_id·bind·current·필터) + 단위 테스트. 호출처 없음 | ezbak/core/diag.py, ezbak/core/__init__.py, ezbak/core/CLAUDE.md, tests/test_core_diag.py | golden 무변경. 회전·형식·run_id·두 번 setup·폴더 실패 |
| P6-2 | 진입점 연결: ezBAK.py 에서 `diag.setup()`, Headless 백업 run_id·시작 줄 | ezBAK.py, ezbak/automation/headless.py, tests/test_automation_headless.py | golden 무변경(stdout 그대로). ezBAK.py ≤100 |
| P6-3 | GUI 작업 run_id: 백업·복원·복사·도구·예약 시작과 작업 스레드 bind, 시작 줄 | ezbak/gui/actions_backup.py, actions_restore.py, actions_copy.py, actions_tools.py, actions_schedule.py | golden 무변경 |
| P6-4 | g-j NAS out·err·g-l winget simple rc≠0 을 진단 로그로(암호 없음 단언 유지) | ezbak/gui/dialogs/nas.py, ezbak/gui/actions_tools.py, tests(해당 단위) | tool_net golden 무변경(caplog 암호 단언 통과) |
| P6-5a ★ | gui print → `_log.debug`(`DEBUG:`)·`_log.warning`(`ERROR:`)·`_log.info`(`Task created`): dialogs/schedule·schedule_layout·actions_schedule | 3파일 + tests/test_gui_smoke.py 필요 시 | schedule_cases·gui_dialogs 의 stdout 칸만 재생성(D1) |
| P6-5b ★ | 〃 sources·sources_tree | 2파일 | gui_dialogs·gui_cases stdout 칸만 |
| P6-5c ★ | 〃 bridge·shortcuts·settings·app_log·i18n, smoke `test_enable_buttons` 를 caplog 로 | 5파일 + test_gui_smoke.py | gui_snapshot settings stdout·tool_cases stdout 칸만 |
| P6-6 | 검사 d 켜기(D2 범위) + 자기 검사 | tests/test_architecture.py, docs/adr/0003-구현-규칙.md | 위반 0, 위반 예시 실패 단언 |
| P6-7 | h-w: finally 의 `progress_bar['maximum']` 읽기를 큐(메인 스레드)로 — 결함 커밋 | actions_backup·restore·tools, bridge | 해당 golden 4파일 큐 순서 확인(바뀌면 ★) |
| P6-8 | 로거 이름 `_LOG` → `_log`(4파일), 정리 후보 중 호출처 0 인 것 삭제 | engine·retention·storage/local·bridge + 해당 | golden 무변경 |
| P6-9 | release.yml: 빌드 전 `pip install -r requirements-dev.txt` + `python -m pytest tests -q`(검사 f), 로컬 ps1 을 release.yml 플래그에 맞춤(D3·D4) | .github/workflows/release.yml, scripts/build_pyinstaller_script.ps1, scripts/CLAUDE.md | 태그 없이 확인: workflow 문법, ps1 은 사용자 실기기 빌드 |
| P6-10 | 기술문서: structure.md·roadmap·모듈 CLAUDE.md(진단 로그 위치·EZBAK_LOG_LEVEL)·README 의 로그 안내 | 문서 | — |
- 순서: 1 → 2 → 3 → 4 → 5a~c → 6 → 7 → 8 → 9 → 10. ★ 는 같은 golden 을 고치므로 서로 병렬 금지. golden 재생성은 `Regen-Golden:` 줄(regen-golden.yml).

## 3. 결정할 항목 (사용자)
결정(2026-10-27): D1 A · D2 B · D3 A · D4 A · D5 A (모두 권장안).
- D1 gui print 의 stdout: A) 모두 진단 로그로 옮기고 golden 의 stdout 칸을 Q1 예외로 재생성(권장: `--windowed` exe 에는 콘솔이 없어 사용자가 보던 출력이 아니다. C3 목록에 GUI stdout 없음) / B) `ERROR:`·`Task created` 5곳은 print 유지, DEBUG 만 옮김.
- D2 검사 d 범위: A) ADR 그대로 `print("DEBUG`·맨 `except:` 만 / B) `ezbak/` 의 모든 `print(` 금지, 예외 `automation/headless.py`(C3 stdout)만(권장: P6-5 뒤 gui print 0 이라 비용 없음, 새 print 가 다시 생기지 않음).
- D3 로컬 빌드 버전 파일: A) 저장소의 `bak_version_info.txt` 유지(권장: 로컬 빌드는 버전이 없어도 됨) / B) release.yml 처럼 ps1 이 생성.
- D4 spec: A) spec 없이 ps1·release.yml 플래그를 같게(권장: 지금 방식, 파일 하나 덜 관리) / B) `ezBAK.spec` 을 저장소에 두고 둘 다 그걸로 빌드.
- D5 진단 로그 기본 레벨: A) ADR 그대로 INFO(권장: DEBUG 줄은 `EZBAK_LOG_LEVEL=DEBUG` 일 때만, 파일이 빨리 차지 않음) / B) DEBUG.

## 4. 진행 (2026-10-27)
- 끝: P6-1~6·8~10(해시는 roadmap "P6 발주 순서"). 계획과 다른 점: P6-2 는 ezBAK.py 대신 `headless.run`·`gui.run` 이 setup 을 부른다(루트 간선에 core 없음, ADR 0002). P6-5 는 d(winsys `system_name`·automation `create_task` 의 `debug=print` 기본값)를 더했다 — gui print 를 옮긴 뒤에도 stdout 에 DEBUG 줄이 남아서.
- P6-7(h-w): 그대로 둠(결정 C). 방법 셋의 비용:
  - A) 큐 값을 `None`(메인 스레드가 maximum 으로 채움)으로: 큐 golden(backup_cases·restore_cases·tool_cases·tool_profiles, 해당 경우만)의 마지막 `update_progress` 값이 바뀜 → Q1 예외 재생성.
  - B) maximum 을 위젯 밖 속성에 거울로 두고 작업 스레드가 그걸 읽음: 큐 값은 같지만 smoke 테스트가 위젯을 직접 바꾸는 곳(`progress_bar.config(maximum=…)`)과 거울이 갈려 테스트 순서에 따라 golden 이 흔들릴 수 있음.
  - C) 그대로 둠: tkinter 는 다른 스레드의 cget 을 Tcl 스레드로 넘겨 처리하고, mainloop 가 끝난 뒤의 RuntimeError 는 이미 `except Exception` 이 받는다. 실제 고장 보고 없음.
  - 결정(2026-10-27 사용자): C. 코드는 그대로 두고 이 문단으로 닫는다.
- 남은 확인: 로컬 PyInstaller 빌드(ps1)와 진단 로그 파일 생성은 Windows 실기기에서(U8 체크리스트와 같이).
