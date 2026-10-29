# P4 대상·저장 나머지 발주 계획 (초안, 2026-10-12, 기준 0500c55, ezBAK.py 6292줄)
따를 것: ADR 0002·0003, targets·storage·winsys·gui·tests CLAUDE.md. 완료 기준(단계): `argv_golden.json` 과 P4-0 경우별 golden 이 그대로 + ezBAK.py 의 `subprocess` 0줄(import 포함) + 아키텍처 검사 a~c.

## 1. 현재 코드 대응 (ezBAK.py 줄, subprocess 23줄 = 아래 ★ 합계)
| 영역 | 줄 · 하는 일 | 호출처 | 옮길 곳 · 공개 이름 | 남길 이름(wrapper) |
|---|---|---|---|---|
| 시스템 이름 | `get_system_info` 1083–1173: 레지스트리 BaseBoardProduct → ★wmic 3곳(1127·1143·1159, encoding=호출 시점 `locale.getpreferredencoding()`) → "Unknown", `clean_name`, DEBUG print | 드라이버 백업 4212·복원 4351, argv golden | `winsys.sysinfo.system_name()`(D2) | 모듈 함수 `get_system_info`(golden 이 부름) |
| 장치 관리자 | `open_device_manager` 4466–4502: startfile → ★mmc Popen 4479 → ★`start` shell 4487 → showerror | 버튼·Ctrl+M | `winsys.shell.open_device_manager() -> 성공 문구`(실패면 FileNotFoundError 그대로) | App 메서드(self.log·showerror) |
| 드라이버 | `run_driver_backup` 4206–4296: 이름·폴더·수동 로그([POLICY] 줄) → ★pnputil export(문자열, shell=True, cp949, PIPE 4245–4246) → poll/readline → communicate → STDERR. `run_driver_restore` 4345–4464: `_drivers_backup_` 검사·이름 추출·불일치 askyesno(작업 스레드) → ★pnputil add(4420–4421) | 시작 메서드 2개 | `targets.drivers`: `build_export_cmd(path)`·`build_add_cmd(folder)`·`backup_folder_name(system, now)`·`parse_backup_folder(name) -> 시스템명|None`·`restore_log_name(system, now)`·`run_pnputil(cmd, out) -> rc` | start_* 2개, run_* 2개(로그 열기·큐·대화상자·finally) |
| 앱 목록 | `run_winget_export` 3807–3944: ★--version(3813) → ★export Popen 스트림(3824) → ★basic(3851) → ★list --accept(3877, ★TimeoutExpired 3894) → ★list(3903) → RuntimeError | start 3785 | `targets.apps`: `build_*_cmd` 4개 · `export(output_path, out, *, clock) -> AppsExport(kind, path)`(전부 실패 RuntimeError 현행 문구, FileNotFoundError 는 올림) | start·run(큐 문구·"Possible solutions" 오류 문구·finally) |
| NAS | 연결 3498–3636·해제 3700–3783: 명령 조립 → ★run cp949·ignore, 예외면 ★인코딩 없이 재실행(3521·3523·3600·3602·3719·3721·3760·3762) → 결과 해석. 예외면 재입력(simpledialog) 후 한 번 더. 드라이브 문자 3334–3365(ctypes GetLogicalDrives) | 대화상자 2개 | `storage.nas`: `NasSpec`(password repr=False) · `build_connect_cmd(spec)`·`disconnect_target(text) -> (argv, label)|None` · `run_net_use(argv) -> NetUseResult(rc, out, err, message)` · `free_drive_letter(used)`. `winsys.sysinfo.used_drive_letters()` | 대화상자 2개(UI·재입력 흐름 그대로, P5 에 gui/dialogs/nas) |
| 프로필 전체 | `run_browser_profiles_backup` 4055–4185(subprocess 없음): Chrome·Edge User Data, Firefox Profiles+profiles.ini, 정션 건너뜀(pywin32/ctypes), 크기·여유 공간(`disk_usage` 실패 시 없는 `self.get_free_space`), 복사, 큐 | start 4037 | `targets.browser.backup_profiles(dest_root, out, progress, *, local, roaming, free_space, clock) -> dest_base`(공간 부족 InsufficientSpaceError 류 → wrapper 가 현행 show_error) | start·run |
| 정리 | pywin32·ctypes 블록 1015–1052(프로필 is_reparse 만 씀), `import subprocess` 11 | — | 삭제(grep 0 확인) | — |

## 2. golden 먼저 (손대지 않은 코드, 코드 무변경 커밋)
- 이미 덮임(`argv_golden.json`): 위 23줄 전부의 argv·kwargs·호출 순서·대화상자 이름. 경우: sysinfo_fallback_chain(결과 포함), driver_backup·driver_restore(일치), winget_export_ok·winget_fallbacks, nas_connect_drive_user·retry_path, nas_disconnect_unc·retry_path, devmgr 3종.
- 안 덮임(새로): 큐 메시지 순서, 작업 로그 내용(<TS>·<TMP>·<FRAMES>), 만든 파일·폴더 이름, 반환값, 그리고 경우 —
  - sysinfo: 레지스트리 적중, O.E.M. 자리표시 건너뜀, 특수문자·공백 정리, rc≠0, wmic 예외, 전부 실패 Unknown. 드라이브 문자: 비트마스크·예외 시 exists 탐색·Z→D→C→'Z'
  - drivers: 백업 rc 0·1·Popen 예외·stderr 있음, 복원 폴더 이름 아님·불일치 예·아니오·대문자 `_DRIVERS_BACKUP_`(→Unknown, 경고 없음)·rc≠0·예외(traceback 줄). devmgr 전부 실패 문구
  - winget: 폴더 취소, basic 성공, list 성공, simple 성공·실패, 전부 실패, Popen FileNotFoundError, list TimeoutExpired
  - NAS: 연결·해제 취소, rc≠0(err·out·코드 문구), 사용자 없는 UNC, 잘못된 드라이브 문자, 재입력 흐름에서 또 예외(메서드 밖으로 올라감)
  - 프로필: 원본 없음, 3종+ini, 정션 건너뜀(make_junction), 공간 부족, disk_usage 실패(AttributeError 문구), 취소
- 보안: NAS 경우는 암호 `dummy-pw` 가 caplog(전 로거 DEBUG)·작업 로그·큐에 없음을 단언한다. golden JSON 에는 argv 안에만 있다(현행 argv_golden 과 같음).

## 3. 지키는 법 (golden 이 붙잡는 것)
- 호출 시점 조회: `locale.getpreferredencoding()`, `winreg`(테스트가 OpenKey 차단), `os.startfile`, `WINDIR`, `ctypes.windll.kernel32`, `time.sleep`, 시계는 wrapper 가 `clock=lambda: datetime.now()`·`now=datetime.now()` 로 넘김(fixed_clock).
- kwargs 는 글자 그대로(timeout 300·300·180, shell=True 문자열 명령, PIPE 는 `proc.PIPE` 재노출). 로그 줄 순서("Executing command" 가 실행 앞).
- 로그 열기·[POLICY] 줄·큐·대화상자·finally·play_sound 는 wrapper(P3 D1 과 같음). 드라이버 불일치 askyesno 는 작업 스레드에서 그대로(P5 에 정리).
- `%APPDATA%` 는 `os.path.expandvars` 값(미설정이면 문자열 그대로)을 넘긴다. `winsys.paths.appdata_dir()`('' 반환)를 쓰면 상대 경로가 돼 동작이 바뀐다.
- NAS 암호: `NasSpec` repr 에서 빼고, 진단 로그에는 proc 이 프로그램 이름만 남긴다(현행). 오류 문구에 argv 를 넣지 않는다.

## 4. 결정 (사용자)
- D1 경계: A) targets·storage 는 명령 조립·실행·출력 해석·작업 로그 줄(`out`)까지, 큐·대화상자는 wrapper 가 반환값으로(권장: P3 D1 과 같고 Reporter 포트를 늘리지 않음) / B) Reporter 를 받아 큐까지(open_folder·show_error·play_sound 가 포트에 없어 포트 확장).
- D2 시스템 이름: A) `winsys.sysinfo.system_name()` 통째(권장: 예상 변경인 wmic 퇴역이 winsys 축, clean_name 은 C3 폴더 이름 고정) / B) winsys 는 원시 조회만, 우선순위·clean_name 은 targets.drivers(함수 4개+, 이름 규칙 분리).
- D3 NAS 재시도: A) "인코딩 없이 재실행"은 `run_net_use` 안, "예외 시 재입력" 흐름은 wrapper(권장: UI 흐름이라 P5 dialogs/nas 몫) / B) 재입력까지 nas 에 콜백 5개.
- D4 도달 불가 `except Exception as e` 절(3631·3778, 같은 try 의 두 번째 절): A) 옮기며 삭제, 커밋 메시지에 기록(권장: 실행될 수 없음) / B) 그대로 옮김.
- D5 프로필 원본: A) browser.py 에 `PROFILE_SOURCES` 별도 표(권장: Firefox 를 BROWSERS 에 넣으면 북마크 대상·bm 문구가 바뀜) / B) BROWSERS 에 열 추가.
- D6 프로필 진행률: A) GuiReporter 재사용(권장: P3-3 와 같고 경우 golden 큐 순서로 확인) / B) 현행 App 스로틀 콜백을 progress 로 넘김(큐 동일 보장 대신 P5 정리 몫).
- D7 루트 검사 c: A) P4 끝에 `import subprocess` 삭제 + test_architecture 검사 c 를 ezBAK.py 에도 켬(권장: 되돌아오는 subprocess 를 바로 막음, a·b 는 P5) / B) P5 에 a~c 한꺼번에.

## 5. 발주 순서 (발주 1건 = 커밋 1개, ezBAK.py 는 매번 줄어듦)
| # | Goal | Files | 추가 테스트 | 위험 → 막는 법 |
|---|---|---|---|---|
| P4-0a | 드라이버·sysinfo·devmgr 경우별 golden | tests/characterization/test_tool_cases.py, tests/fixtures/tool_cases_golden.json, tests/fixtures/tool_harness.py | §2 sysinfo·drivers·devmgr | 코드 무변경. argv 테스트의 _fill_dialog 등은 tool_harness 로 복사(argv 테스트는 안 고침) |
| P4-0b | winget·NAS 경우별 golden + 암호 단언 | tests/characterization/test_tool_cases.py, tests/fixtures/tool_cases_golden.json, tests/fixtures/tool_harness.py | §2 winget·NAS | 무변경 |
| P4-0c | 프로필 경우별 golden | 같은 3파일 | §2 프로필, pywin32 있음·없음 두 벌 | 무변경 |
| P4-1 | `proc.PIPE` 재노출, `sysinfo.system_name`·`used_drive_letters`, `get_system_info` thin | winsys/proc.py, winsys/sysinfo.py, winsys/__init__.py, ezBAK.py, tests/test_winsys_sysinfo.py | 우선순위·clean_name·인코딩 kwarg | DEBUG print 문구 그대로, locale·winreg 호출 시점 |
| P4-2 | `winsys.shell.open_device_manager` | winsys/shell.py, winsys/__init__.py, ezBAK.py, tests/test_winsys_shell.py | 3단 fallback·실패 | startfile 뒤 log, 실패 문구는 wrapper |
| P4-3 | `targets.drivers` + run_driver_* 위임 | targets/drivers.py, targets/__init__.py, ezBAK.py, tests/test_targets_drivers.py | 명령 문자열·이름·parse·스트림 | 문자열 명령·shell=True·cp949 글자 그대로, 로그 열기는 wrapper |
| P4-4 | `targets.apps.export` + run_winget_export 위임 | targets/apps.py, targets/__init__.py, ezBAK.py, tests/test_targets_apps.py | 4단 fallback·TimeoutExpired·파일 이름 | clock 읽는 시점, FileNotFoundError 는 올림 |
| P4-5 | `storage.nas` + 대화상자 2개의 실행부 위임(D3·D4) | storage/nas.py, storage/__init__.py, ezBAK.py, tests/test_storage_nas.py | argv 표·disconnect_target·재실행·암호 비노출 | 재입력 흐름의 예외가 메서드 밖으로 나가는 현행 유지 |
| P4-6 | `browser.backup_profiles` + pywin32·ctypes 블록 삭제 | targets/browser.py, targets/__init__.py, ezBAK.py, tests/test_targets_browser.py | 원본 표·정션·공간 | APPDATA expandvars, free_space 콜백으로 AttributeError 재현, browser.py 500줄 이하 |
| P4-7 | (D7=A) `import subprocess`·남은 import 삭제, 검사 c 루트 적용 | ezBAK.py, tests/test_architecture.py | 검사 c 루트 | grep 0 확인 뒤 |
- P4 끝: 메인이 targets(drivers·apps·backup_profiles 시그니처, CLAUDE.md 의 `(dest, reporter)` 초안 교체)·storage(nas)·winsys(system_name·used_drive_letters·shell·PIPE) CLAUDE.md, roadmap 갱신.

## 6. 보존할 이상 동작 · P4-fix 후보 (사용자가 정한다)
- T0-3: NAS 예외 시 같은 명령 재실행·연결 쪽 자격 증명 재요청, pnputil `cp949` 하드코딩·문자열+shell=True(경로에 `"` 면 깨짐), wmic 3곳 퇴역 중.
- 새로 찾은 것: NAS 재입력 흐름의 예외는 잡히지 않음, 드라이버 복원 이름 검사는 소문자·추출은 대소문자 구분(대문자면 Unknown·경고 없음), 불일치 askyesno 를 작업 스레드에서 호출, 드라이버 로그에 무관한 [POLICY] 줄, 백업 finally 는 reset_progress·복원은 stop_progress, 프로필 백업도 없는 `self.get_free_space`, 프로필 원본은 실행 계정, winget simple list rc≠0 은 로그 없음.
