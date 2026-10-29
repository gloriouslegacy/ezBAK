# P5 GUI 분리 발주 계획 (초안, 2026-10-13, 기준 2e1df09, ezBAK.py 5820줄, 729 passed)
따를 것: ADR 0002·0003, gui·tests·winsys CLAUDE.md. 시작 조건: P4-fix 가 병합된 뒤. 줄 번호는 P4-fix 로 바뀌므로 발주 때 `grep -n "def \|class "` 로 다시 잡는다.
완료 기준(단계): golden 전부 그대로 + ezBAK.py 100줄 이하 + tkinter 는 `ezbak/gui/` 만 + 루트에 검사 a~c.
## 1. 목표 배치 (괄호는 예상 줄 수, 500 넘으면 같은 축 안에서 둘로)
| 파일 | 옮길 것 (ezBAK.py 현재 이름) |
|---|---|
| gui/app.py (~300) | `App(<믹스인들>, tk.Tk)`: `__init__`(88→상태·변수·시작 3개로), create_menu_bar, setup_keyboard_shortcuts, load_users, _update_header_time, set_buttons_state, reset_progress_bar, open_github_link, show_keyboard_shortcuts_help · `run(entry_script)`(App(entry_script)·mainloop) |
| gui/layout.py (~430) | create_widgets 409줄 → 같은 순서의 절 함수 4개(각 ≤150, 헤더·백업·도구·로그/진행) |
| gui/theme.py · i18n.py | Win11Theme + 테마 믹스인(_configure_progressbar_style, apply_theme_to_all_widgets, switch_theme) · Translator + 언어 믹스인(update_ui_texts, switch_language) |
| gui/strings_en.py · strings_ko.py | Translator 사전(527줄)을 언어별로 |
| gui/widgets.py · shortcuts.py | ToggleSwitch, `create_win11_button`(D5) · DialogShortcuts, KeyboardShortcuts |
| gui/settings.py | settings_file_path, load_settings, save_settings, copy_default_filters (settings.json 유일 소유) |
| gui/app_log.py | log, write_detailed_log, open_log_file, close_log_file, _cleanup_old_logs, save_log, _play_sound, _on_sound_toggle |
| gui/bridge.py (+~100) | 큐 믹스인 process_queue(13종), 복원 확인 콜백 2개(_wait_restore_folder_choice, _wait_cross_user_confirm) |
| gui/actions_backup.py · actions_copy.py | 백업 시작·run_backup·_prepare·_run_backup_engine·_finish, is_hidden·*_safe 3개, copy_file_with_progress, get_free_space_reliable · copy_files·run_copy_thread, check_space·run_space_calculation·show_space_check_result |
| gui/actions_restore.py · actions_schedule.py | start_restore_thread·run_restore(97→≤60 둘) · schedule_backup·create/delete_scheduled_task·run_create…thread·_show_task_* (`__file__` → `self.entry_script`) |
| gui/actions_tools.py (~360) | winget, 브라우저 프로필, 드라이버 백업·복원(80·99→≤60), 장치 관리자, 탐색기 |
| gui/dialogs/ | message.py(Win11Dialog), filters.py, sources.py(+sources_tree.py), schedule.py(+schedule_layout.py, _create_widgets 210→≤150 둘), nas.py(get_used_drive_letters·find_free_drive_letter·연결 243/해제 121), restore_select.py(show_folder_selection_dialog·show_cross_user_warning_dialog) |
- ezBAK.py 최종(~70): import → `_run_headless_if_requested()`(automation.cli.parse → headless.run, users_root=`winsys.paths.users_root()`, clock) → frozen 이면 `gui.run(entry_script=abspath(__file__))`(현행 try/traceback/exit 1) / 비frozen 은 is_admin 아니면 relaunch_as_admin(현행 print·exit 0·os._exit) 아니면 gui.run.
- 지울 것(호출처 0 재확인 뒤): log_message, on_exit, set_progress_indeterminate, Win11Theme.switch_mode, 루트 resource_path·format_bytes·get_system_info·core_run_backup·USERS_ROOT·_have_pywin32(테스트 별칭은 facade 로).
## 2. 결정 (사용자) — 확정 2026-10-13: D1~D6 모두 A. P4-fix 에서 넘어온 결함 커밋 2건(이동 뒤, 각 1커밋·Q1 예외): g-e 드라이버 불일치 askyesno 를 작업 스레드 밖 대화상자 요청으로(P5-14 뒤), g-p 프로필 전체 백업 원본을 선택 사용자 프로필로 + 시작 확인 대화상자(F4=C, P5-14 뒤)
- D1 App 나누기: A) 기능별 믹스인, `class App(믹스인…, tk.Tk)`(권장: 메서드 이름·self·단축키 바인딩·테스트 호출이 그대로, 믹스인을 Tk 앞에 둬 지금처럼 Tk 메서드를 가림) / B) 모듈 함수 + App 위임 메서드(이름이 두 벌) / C) 컨트롤러 객체(app.backup.start, 테스트·단축키 이름 변경).
- D2 테스트 패치 이전: A) `tests/fixtures/ezbak_facade.py` — ezbak_mod 가 facade 를 돌려주고, 읽기는 루트→ezbak.gui 공개 이름→별칭 표, `setattr` 은 루트와 로드된 `ezbak.gui.*` 중 그 이름을 가진 모듈 전부에 퍼뜨림(권장: monkeypatch 호출 20여 곳이 안 바뀌고, 테스트가 gui 내부 파일을 직접 참조하지 않음) / B) 이동마다 테스트를 새 위치로(Files 초과, 내부 파일 참조) / C) ezBAK.py 에 별칭 유지(100줄·검사 a 위반).
- D3 USERS_ROOT: A) `winsys.paths._USERS_ROOT` 를 공개 `USERS_ROOT` 로, gui·루트는 호출 시점에 읽고 facade 가 그 값을 패치(권장: ADR 0002 I/O 표가 C:\Users 를 paths 소유로 둠, 정의 1벌) / B) 루트 상수 유지 + gui.run 인자(세션 App 이 값을 들고 있어 inject 가 안 닿음).
- D4 복원 확인: A) 0.1초 속성 폴링 그대로 옮기고 ADR 0002 의 "Event 전환은 P5" 를 도입 조건(취소 요구 확정 시)으로 고침(권장: Q1, 테스트 훅 `_folder_selection_result` 유지) / B) P5 끝에 Event 전환 별도 커밋.
- D5 상한 초과 함수·중복: A) 같은 순서의 절 함수로 쪼개고, `_create_win11_button` 3벌·SelectSources 의 단축키 2메서드는 글자 그대로 같을 때만 widgets·shortcuts 로 합침(권장: 상한 지키고 동작 동일) / B) 그대로 옮기고 CLAUDE.local `[예외]`.
- D6 이동 전 GUI golden: A) P5-0b~d 전부(권장: 지금 golden 이 없는 GUI 코드 ~2500줄 — 대화상자·layout·settings·번역) / B) 스모크만 늘림(layout 절 분할을 증명 못 함).
## 3. 발주 순서 (발주 1건 = 커밋 1개, Files 에 golden 출력 제외, ezBAK.py 는 매번 순감)
| # | Goal | Files | 덮는 테스트 · 위험 → 막는 법 |
|---|---|---|---|
| P5-0a | facade(D2) + `_have_pywin32` 별칭 이전 | tests/fixtures/ezbak_facade.py, tests/conftest.py, ezBAK.py | 전 특성 테스트. nopywin32 는 새로 import 한 ezbak.* 를 pop 전에 facade 에 담음. facade 자체 테스트(퍼뜨림·없는 이름 AttributeError) |
| P5-0b | App golden: 위젯 트리(en/ko×dark/light, 클래스·text·배치 옵션·색), 메뉴, 바인딩, 번역 사전, settings.json 왕복 | tests/characterization/test_gui_snapshot.py, tests/fixtures/gui_harness.py | 코드 무변경. 자동 위젯 이름(!frame2)은 순서로만 비교 |
| P5-0c | 대화상자 golden: Win11Dialog 4종, Filter(+_add_rule·_save 결과), Sources(임시 트리·_on_ok), Schedule(on_create task_data·on_delete), NAS 2개 트리, 폴더 선택 | tests/characterization/test_gui_dialogs.py, tests/fixtures/gui_harness.py | 무변경. wait_window 는 tool_harness 방식으로 막음 |
| P5-0d | 복사·공간 계산 경우 golden + test_gui_bridge.py(GuiReporter 스로틀·최종 보고, LogStream, FolderStore, GuiJob·int_or) | tests/characterization/test_gui_cases.py, tests/test_gui_bridge.py | 무변경(P2 이월 항목) |
| P5-1 | i18n.py + strings_en/ko | gui/i18n.py, gui/strings_en.py, gui/strings_ko.py, gui/__init__.py, ezBAK.py | 0b 번역·트리. 사전 키 순서 그대로 |
| P5-2 | theme.py(Win11Theme, switch_mode 삭제)·widgets.py(ToggleSwitch) | gui/theme.py, gui/widgets.py, gui/__init__.py, ezBAK.py | 0b dark/light |
| P5-3 | dialogs/message.py, 루트 resource_path → winsys.paths | gui/dialogs/__init__.py, gui/dialogs/message.py, gui/__init__.py, ezBAK.py | 0c, 모든 특성 테스트의 mod.Win11Dialog 기록(클래스 패치라 위치 무관) |
| P5-4 | shortcuts.py | gui/shortcuts.py, gui/__init__.py, ezBAK.py | 0b·0c 바인딩 |
| P5-5 | dialogs/filters.py(+D5 widgets 합침), App.open_filter_manager 중복 항목 확인(지금 1벌)·닫기 | gui/dialogs/filters.py, gui/widgets.py, gui/__init__.py, ezBAK.py | 0c Filter |
| P5-6 | dialogs/sources.py(+sources_tree.py) | gui/dialogs/sources.py, gui/dialogs/sources_tree.py, gui/__init__.py, ezBAK.py | 0c Sources |
| P5-7 | dialogs/schedule.py(+schedule_layout.py) | gui/dialogs/schedule.py, gui/dialogs/schedule_layout.py, gui/__init__.py, ezBAK.py | 0c Schedule |
| P5-8 | app_log.py, settings.py, bridge 큐 믹스인 | gui/app_log.py, gui/settings.py, gui/bridge.py, gui/__init__.py, ezBAK.py | smoke 13종·play_sound, 0b settings, 모든 작업 로그 golden. process_queue 를 쪼개면 smoke 의 getsource 대상도 같은 커밋에서 |
| P5-9 | theme·i18n 믹스인 + layout.py(절 4개) | gui/layout.py, gui/theme.py, gui/i18n.py, gui/__init__.py, ezBAK.py | 0b 트리 4벌, smoke 위젯. 생성 순서 = pack/grid 순서 |
| P5-10 | actions_backup.py + D3(paths.USERS_ROOT) | gui/actions_backup.py, gui/__init__.py, ezBAK.py, winsys/paths.py, tests/fixtures/ezbak_facade.py | backup_golden·cases, test_hidden, cli(엔진 stub 이 별칭 `core_engine_run_backup` 을 identity 로 찾음 — 별칭 이름 유지) |
| P5-11 | actions_copy.py | gui/actions_copy.py, gui/__init__.py, ezBAK.py | 0d |
| P5-12 | actions_schedule.py, `self.entry_script` | gui/actions_schedule.py, gui/__init__.py, ezBAK.py | schedule_cases(<REPO>\ezBAK.py), argv schtasks |
| P5-13 ★ | actions_restore.py + bridge 콜백 + dialogs/restore_select.py | gui/actions_restore.py, gui/bridge.py, gui/dialogs/restore_select.py, gui/__init__.py, ezBAK.py | restore_golden·cases, smoke select_folder·cross_user. time.sleep 은 전역 패치라 그대로 |
| P5-14 ★ | actions_tools.py | gui/actions_tools.py, gui/__init__.py, ezBAK.py | tool_cases·tool_net(winget)·tool_profiles, argv |
| P5-15 ★ | dialogs/nas.py, 연결·해제 중복 블록 합침, get_used_drive_letters 의 ctypes lambda → sysinfo 기본값 | gui/dialogs/nas.py, gui/__init__.py, ezBAK.py, tests/characterization/test_tool_cases.py | tool_net NAS·암호 단언, tool_cases 드라이브 문자(패치를 `winsys.sysinfo` 쪽으로) |
| P5-16 | App → gui/app.py, `gui.run`, 진입점 ≤100, 루트 별칭 삭제, 믹스인 임시 공개 이름 정리 | gui/app.py, gui/__init__.py, ezBAK.py, tests/conftest.py, tests/fixtures/ezbak_facade.py | cli golden `App()` 종료(Tk.__init__ 막기), argv 재실행, smoke. gui_app 은 `App(entry_script=EZBAK_PATH)` |
| P5-17 | 루트 검사 a·b(간선 ezBAK.py → automation, gui, winsys), `paths.local_appdata()` 삭제 | tests/test_architecture.py, winsys/paths.py, tests/test_winsys_basic.py | 아키텍처 전체 |
- P5-17 에서 정할 것(P5-12·14 발견, targets.apps·drivers·storage.nas 도 같음): automation/CLAUDE.md 는 하위 모듈 scheduler·schedule_config 를 공개 진입점으로 적지만 test_architecture 는 하위 모듈 공개를 winsys 에만 허용(SUBMODULE_PUBLIC). 지금은 `from ezbak import automation; automation.scheduler` 로 우회. 루트 검사 b 를 켜면 `from ezbak.automation import cli/headless` 도 걸린다 — `__all__` 에 있는 하위 모듈 이름 import 를 허용할지, automation 을 SUBMODULE_PUBLIC 에 넣을지(ADR 0002 갱신) 결정
- ★ = P4-fix 가 손댈 수 있는 wrapper(복원·드라이버·winget·프로필·장치 관리자·NAS). 그래서 뒤에 둔다. P5-0b~d golden 도 P4-fix 뒤 코드로 찍는다.
- 이동 규칙: 옮긴 모듈은 ezBAK.py 의 import 별칭(`threading`, `datetime`, `ctypes`, `winsys_free_space`, `core_engine_run_backup` 등)을 이름 그대로 쓴다 — facade 퍼뜨리기와 cli 엔진 stub 이 이름·identity 로 찾는다. 믹스인은 P5-16 전까지 `ezbak.gui` 공개 이름으로 내보낸다(루트가 내부 파일을 import 하지 않게).
## 4. 이전 단계 정리 반영
- P4: `_have_pywin32` 별칭 → P5-0a(facade 가 fileattrs.HAS_PYWIN32), ctypes lambda·NAS 중복 블록 → P5-15, `paths.local_appdata()` → P5-17. P3: 복원 콜백 → P5-13(D4), `core_engine_run_backup`·`HeadlessReporter` 별칭은 유지(cli 테스트 identity·golden 클래스 이름). P1: open_filter_manager 중복 → P5-5 확인. P2: 미사용 4개 삭제는 각 이동 단계, test_gui_bridge → P5-0d.
## 5. 완료 기준 (P5 끝, runner 확인)
- `python -m pytest tests -q -W error` 전부 통과, `git diff <P4-fix 끝> -- tests/fixtures/*_golden.json` 은 P5-0 신규 파일 추가만. `fable check` 위반 0(등록 부채 제외), 파일 500·함수 60(gui 배치 150) 이하(ast 일회 확인).
- `ezBAK.py` 100줄 이하·py_compile, test_architecture 검사 a~c 가 루트에도 켜짐. 메인: gui·tests·winsys CLAUDE.md, structure.md, roadmap, ADR 0002(D4) 갱신.
