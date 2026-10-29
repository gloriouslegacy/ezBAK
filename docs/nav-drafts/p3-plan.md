# P3 복원 + automation 발주 계획 (초안, 2026-10-12, 기준 891d0ff, ezBAK.py 7172줄)
따를 것: ADR 0002·0003, core·automation·winsys·targets·tests CLAUDE.md. 완료 기준(단계): restore·argv golden 과 P3-0 경우별 golden 이 그대로 + CLI·스케줄 JSON 왕복 테스트 + 아키텍처 검사 a~c.

## 1. 현재 코드 대응 (ezBAK.py 줄)
| 영역 | 줄 · 하는 일 | 호출처 | subprocess·ctypes(→ 옮길 곳) |
|---|---|---|---|
| 복원 시작 | `start_restore_thread` 2862–2888: 사용자 확인·덮어쓰기 askyesno·askdirectory → 모듈 `threading.Thread(run_restore)` | Ctrl+R 1286, 버튼 2012 | 없음 |
| 복원 본체 | `run_restore` 2890–3170: 검증 3종(show_error) → "Selected backup folder" 큐 → 직접 선택(2922, 사용자 폴더 6종 확인)/부모 선택(2980, `_backup_` 목록 → `select_folder` 큐 + `_folder_selection_result` 속성 폴링) → `open_log_file(…_restore)` → 교차 사용자(`show_cross_user_warning` 큐 + `_cross_user_confirmed` 폴링) → 크기(`is_hidden`, 실패 시 100MB) → 여유 공간 → 복사(50개마다 큐) → 북마크 → `_cleanup_old_logs` → finally(큐 6~8개) | start_restore_thread | 없음 |
| 복원 보조 | `restore_browser_bookmarks` 3172–3196(실행 계정 %LOCALAPPDATA%), 대화상자 2411–2421·2423–2540, 큐 분기 2365–2370 | run_restore, process_queue | 없음 |
| 스케줄 JSON | `_get_schedule_config_dir` 3352, `_save_schedule_config` 3360–3405(키 12개 순서·indent 2·created_at), **`_load_schedule_config` 3407(호출처 0)** | create 3627, delete 3920 | 없음 |
| 스케줄러 | **`_build_auto_backup_command` 3425(호출처 0)**, `_create_task_xml` 3443–3584(지역 `from datetime import`, `~\AppData\Local\ezBAK\schedules`), `create_scheduled_task` 3586–3756, `_analyze_schtasks_error` 3835–3885, **`_validate_time_input` 3888(호출처 0)**, `delete_scheduled_task` 3897–3937. GUI 쪽 `schedule_backup` 3279·스레드 3758·결과 대화상자 3807–3833 은 P5 | schedule_backup 3319, 스레드 3768 | 3696 run·3742 TimeoutExpired·3906 run(shell=True) → `winsys.proc` |
| Headless | **`show_notification` 1178(호출처 0, argv golden toast_* 가 고정)**, **`HeadlessBackupRunner` 1226(0)**, `_HeadlessReporter` 6941, `core_run_backup` 6958–7007 | __main__ 7069·7110, 백업 golden | 1222 powershell(삭제 대상) |
| 진입 | `__main__` 7010: `is_admin` 7011(맨 except), **`HeadlessLoggerQueue` 7017(0)**, `_run_headless_if_requested` 7031–7126(argparse, config 경로 종료 0/1, 직접 인자 항상 0), frozen 이면 App, 아니면 비관리자 재실행 7153–7169 | 모듈 실행 | 7157 list2cmdline·7158 ShellExecuteW → `winsys.elevation` |

## 2. 목표 배치 (공개 진입점 초안, 새 간선 없음)
- `core.restore`: `validate_selection(path) -> None`(RestoreSelectionError: show_error 원문) · `locate_backup(path, user, choose_backup) -> RestorePlan|None`(None=취소; plan: kind direct/parent, backup_root, folder_name, backup_user) · `run_restore(plan, user, destination_dir, out, reporter, *, skip, free_space, confirm_cross_user) -> RestoreResult(cancelled, files, bytes)`. 문구는 restore.py 상수(포맷 1벌). 의존 core 안(fileops·filters)만.
- `targets.browser.restore_bookmarks(backup_root, out, local_appdata)` — BROWSERS 표 재사용. targets→core·winsys.paths.
- `automation.schedule_config`: `config_dir()`(APPDATA 없으면 RuntimeError 현행 문구) · `config_path(task_name)`(`[<>:"/\\|?*]`→`_`) · `save(spec, *, now) -> path` · `load(path) -> dict`(open+json.load 만) · `delete(task_name) -> path|None`. `ScheduleSpec`(쓰기용, 키 순서 고정) · `job_from_config(dict) -> BackupJob`(형 변환 없음, .get 기본값 현행).
- `automation.scheduler`: 순수 `normalize_time`·`schedule_type`·`task_arguments(entry_script, config_file)`·`run_as_user(environ)`·`build_task_xml(…, now)`·`build_create_cmd(name, xml)`·`build_delete_cmd(name)`·`explain_error(rc, stderr, out)` + `write_task_xml(name, xml) -> path` · `create_task(spec, entry_script, *, out, now) -> 성공 문구`(RuntimeError 현행 문구) · `delete_task(name)`. 의존 core(DEFAULT_FILTERS)·winsys.proc.
- `automation.headless`: `run_backup(user, dest, …, *, users_root, clock) -> (backup_path, log_path)`(core_run_backup 본문) · `run(args, *, users_root, clock) -> 종료 코드|None`(None=GUI 로) · `HeadlessReporter`. `automation.cli`: `build_parser()` · `parse(argv=None) -> args`(parse_known_args).
- `winsys.elevation`: `is_admin()` · `relaunch_as_admin(executable, argv) -> ret`. `winsys.proc` 재노출 `TimeoutExpired`·`list2cmdline`(subprocess 는 proc 만).

## 3. golden 이 붙잡는 이름과 지키는 법
- 유지할 ezBAK.py 이름: `App.start_restore_thread`·`run_restore`·`_cleanup_old_logs`·`create_scheduled_task(…)`(문구 반환)·`delete_scheduled_task`, `core_run_backup(user, dest, …)`, `__main__` 흐름(relaunch_runas 는 runpy 로 파일 전체 실행).
- monkeypatch: `mod.threading`(복원·스케줄 스레드), `mod.USERS_ROOT`·`mod.datetime`, `Win11Dialog`·`messagebox`·`filedialog`, env `APPDATA`·`LOCALAPPDATA`·`USERPROFILE`·`USERNAME`·`USERDOMAIN`, `locale.getpreferredencoding`, `ctypes.windll.shell32`(fake_shell32), `subprocess.run`(record_subprocess, proc 를 거쳐도 그대로 잡힘).
- 지키는 법: wrapper 가 `users_root=USERS_ROOT`, `now=datetime.now()`·`clock=lambda: datetime.now()` 를 호출 시점에 넘긴다. 옮긴 모듈은 import 시점에 `datetime` 을 묶지 않는다(`_create_task_xml` 지역 import 가 fixed_clock 에 잡히던 이유). `locale.getpreferredencoding()`·`ctypes.windll.shell32` 는 호출 시점에 조회. 비frozen 인자는 wrapper 가 `entry_script=os.path.abspath(__file__)`(C3).
- 보존할 이상 동작: XML `<Arguments>`·`<Author>` 미이스케이프, 재실행 `list2cmdline(sys.argv)`(덮어써지던 `params` 줄은 효과 없어 옮기지 않음), 삭제 `shell=True`·인코딩 없음, AppData 건너뜀·속성 소실, `_restore_completed` 미설정, `"\n[Old log cleanup]"`. 새로 찾은 것: 디스크 조회 실패 시 없는 `self.get_free_space`(AttributeError → "An error occurred during restore"), 우선순위 목록을 `sorted(reverse=True)` 로 다시 정렬, 복원 북마크도 실행 계정 %LOCALAPPDATA%, 스케줄 생성 로그는 열린 로그가 없으면 버려짐, `--user` 만 주면 GUI 시작, config 값 형 변환 없음·log_folder 는 `--log-folder`. 생성 경로의 `print("DEBUG…")` 는 그대로 옮기고 P6 에서 정리.

## 4. 결정 (사용자, 2026-10-12: D1~D7 모두 A)
- D1 복원 흐름 분할: A) `validate_selection`·`locate_backup`(choose_backup 인자) → wrapper 가 로그 열기 → `run_restore`(confirm_cross_user 인자)(권장: P2 D2 처럼 로그 열기는 wrapper, ADR "함수 인자" 그대로) / B) `run_restore` 하나가 `open_log` 콜백까지 받음(호출 1번, 대신 로그 열기가 core 로 들어가 D2 와 어긋남).
- D2 대기 방식: A) GUI 어댑터가 현행 큐 + 속성 폴링(0.1초) 유지, ADR 0002 의 "threading.Event" 를 P5 로 미룬다고 고침(권장: P3-0 golden 이 속성을 미리 넣어 응답) / B) 지금 Event 로(ADR 문구대로, 대신 하네스 수정·대기 중 종료 동작 변화).
- D3 CLI golden: A) runpy 로 손대지 않은 `__main__` 실행 + `ezbak.core.run_backup` 을 sys.modules 전체에서 stub 교체, 사용자 `Public`, stdout·종료 코드·stub 인자 고정. 왕복은 GUI 생성 JSON → runpy CLI(권장: 추출 뒤에도 같은 테스트) / B) 추출 뒤 단위 테스트만(Q1 증명 없음).
- D4 CLI 파서: A) argparse 그대로 옮김(권장: C3·usage 문구·종료 코드 2 동일) / B) 하위 명령·외부 라이브러리(C3 위반).
- D5 스케줄 JSON 타입: A) 쓰기는 `ScheduleSpec`, 읽기는 dict → `job_from_config` 무변환(권장: 잘못된 JSON 동작 보존) / B) dict 만, 함수 이동 / C) 양방향 dataclass 형 변환(잘못된 JSON 동작 바뀜).
- D6 `show_notification`: A) P3 에서 삭제 + argv golden `toast_plain`·`toast_button` 키 제거(권장: 호출처 0, 0003 §6) / B) P4 `winsys.shell.toast` 로 보존.
- D7 ADR 0003 §6 "필터 type 쓸 때 소문자": A) 현행(받은 dict 그대로 쓰기) 유지하고 ADR 문구 수정(권장, Q1) / B) 쓸 때 소문자화(JSON 내용 변화, Q1 예외).

## 5. 발주 순서 (발주 1건 = 커밋 1개)
| # | Goal | Files | 추가 테스트 | golden 위험 → 막는 법 | ADR 0002 |
|---|---|---|---|---|---|
| P3-0a | 손대지 않은 코드로 복원 경우별 golden | tests/characterization/test_restore_cases.py, tests/fixtures/restore_cases_golden.json, tests/fixtures/user_tree.py | 사용자 미선택·덮어쓰기 아니오·폴더 취소, .log·파일·없음, `_backup_` 인데 사용자 폴더 없음, 백업 없음, 여러 개 선택·취소, 교차 사용자 예·아니오(직접·부모), 공간 부족, disk_usage 실패, 북마크 없음, 55개 파일(50개 메시지), 성공 큐 순서 | 없음(코드 무변경) | — |
| P3-0b | 손대지 않은 코드로 CLI·스케줄 경우별 golden | tests/characterization/test_cli_headless.py, tests/fixtures/cli_golden.json, tests/characterization/test_schedule_cases.py, tests/fixtures/schedule_cases_golden.json | D3 runpy: config 성공·없음·JSON 오류·사용자 없음, 직접 인자 성공·실패·filters 오류, `--auto-backup`·`--retention abc`. 스케줄: JSON 내용·이름 정규화, 잘못된 시각·종류, rc 1+access denied, 타임아웃, 삭제 실패, GUI 생성 → CLI 왕복 | 없음 | — |
| P3-1 | `winsys.elevation` + proc 재노출, 재실행 위임 | winsys/elevation.py, winsys/proc.py, winsys/__init__.py, ezBAK.py, tests/test_winsys_elevation.py | is_admin 예외→False, ShellExecuteW 인자 | shell32 호출 시점 조회. 맨 except → except Exception(커밋 메시지) | winsys 단독 |
| P3-2 | `validate_selection`·`locate_backup` + run_restore 앞부분 위임 | core/restore.py, core/__init__.py, ezBAK.py, tests/test_core_restore.py | 검증 3종·직접/부모·정렬·취소 | "Selected backup folder" 큐가 검증 뒤·위치 앞. 속성 폴링은 wrapper(D2) | core 안 |
| P3-3 | `run_restore` 교차 확인~완료 줄 + wrapper | core/restore.py, core/__init__.py, ezBAK.py, tests/test_core_restore.py | skip·free_space·confirm 콜백, 100MB, 50개 | 진행률은 GuiReporter 재사용(마지막 값 무조건 전송)→경우 golden 큐 순서. free_space 콜백에 AttributeError 재현. finally·북마크·정리는 wrapper | core 안 |
| P3-4 | 북마크 복원 → `targets.browser.restore_bookmarks` | targets/browser.py, targets/__init__.py, ezBAK.py, tests/test_targets_browser.py | 있음·없음·복사 오류 | `winsys.paths.local_appdata()`(미설정 시 문자열 그대로)로 현행 대상 | targets→winsys.paths |
| P3-5 | `automation.schedule_config` + App 2메서드 위임, `_load_schedule_config` 삭제 | automation/__init__.py, automation/schedule_config.py, ezBAK.py, tests/test_automation_schedule_config.py | 키 순서·indent·ensure_ascii, APPDATA 없음, 왕복 | created_at 은 now 인자 | automation→core·winsys |
| P3-6 | `scheduler` 순수 함수 + `_create_task_xml`·`_analyze_schtasks_error` 위임, `_build_auto_backup_command`·`_validate_time_input` 삭제 | automation/scheduler.py, automation/__init__.py, ezBAK.py, tests/test_automation_scheduler.py | XML 3종·미이스케이프·오류 해석 표 | 템플릿 글자 그대로 상수로(함수 60줄), now 인자, 맨 위 `timedelta` 는 grep 0 이면 삭제 | automation 단독 |
| P3-7 | `create_task`·`delete_task` 를 `proc.run` 으로, App 2메서드 thin wrapper | automation/scheduler.py, ezBAK.py, tests/test_automation_scheduler.py | proc 모킹 argv·kwargs·TimeoutExpired | kwargs 그대로, encoding 호출 시점, 큐·대화상자는 wrapper, entry_script | automation→winsys.proc |
| P3-8 | `headless.run_backup`·`HeadlessReporter`, `core_run_backup` thin, `HeadlessBackupRunner` 삭제 | automation/headless.py, automation/__init__.py, ezBAK.py, tests/test_automation_headless.py | 폴더 준비 순서·사용자 없음 | users_root·clock 은 wrapper(백업 golden) | automation→core·storage·targets |
| P3-9 | `cli` + `headless.run` → 종료 코드, `__main__` 위임, `HeadlessLoggerQueue`·argparse import 삭제 | automation/cli.py, automation/headless.py, automation/__init__.py, ezBAK.py, tests/test_automation_cli.py | 인자·기본값·종료 코드 표 | stdout `[ts]`(clock)·traceback 그대로, `--user` 만이면 None | ezBAK.py→automation |
| P3-10 | (D6=A) `show_notification` 삭제 + toast_* golden 키 제거 | ezBAK.py, tests/characterization/test_commands_argv.py, tests/fixtures/argv_golden.json | — | 다른 키 그대로, 커밋 메시지에 사유 | — |
- P3 밖: GUI 대화상자·`schedule_backup`·결과 대화상자·settings.json(P5), wmic·NAS·winget·pnputil·`get_free_space` 4827(P4), DEBUG print(P6). P3-fix 후보: 복원 북마크 대상 프로필, `get_free_space`, 정렬 우선순위, XML 이스케이프.
- P3 끝: 메인이 automation·core·targets·winsys CLAUDE.md(restore·elevation·proc 재노출, shell.toast 줄), ADR 0002(D1·D2), roadmap 갱신.
