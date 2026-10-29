# P4-fix 계획 (초안 2026-10-13 · 기준 2e1df09, ezBAK.py 5820줄) — 후보 20건, 커밋 = 발주 1건, golden 은 그 커밋에서만(Q1 예외, 사유·바뀐 키를 커밋 메시지에)
공통: 리팩터링과 섞지 않는다. ezBAK.py 는 발주당 +30줄 이하·CRLF·Edit 도구로만. golden 재생성은 해당 테스트 파일만(`set EZBAK_REGEN_GOLDEN=1 && python -m pytest <파일>`), 재생성물은 5개 셈에서 빼지만 **병렬 판단에서는 같은 golden JSON 도 겹침으로 센다**(작업 트리 공유). 새 진단은 `_log`(ADR 0003 §2), 작업 로그 문구(C3)는 늘리거나 바꾸지 않는 것이 기본. 실제 pnputil·net use·wmic 없음(record_subprocess). NAS 암호 단언 유지
C3 표시: ● C3 항목 변경 · △ C3 문구는 같고 작업 로그 줄 수·파일 이름 값만 달라짐 · – 없음

## 1. 후보 표
| id | 증상 | 위치(확인함) | 고칠 것 | C3·동작 | golden(파일: 키) | 분류 |
|---|---|---|---|---|---|---|
| g-a | 장치 관리자 오류 대화상자 `{0}` 미치환("…Manager:\n{0}:\n<e>") | ezBAK.py `App.open_device_manager` 4027 | `get('unable_to_open_devmgr').format(e)` | – (대화상자) | tool_cases: devmgr_all_fail.dialogs | 지금 |
| g-b | C~Z 다 쓰이면 빈 문자 'B' | storage/nas.py `free_drive_letter` | — 실행 중 C 가 빈 경우가 없고 B: 매핑은 동작함 | – | — | 그대로 |
| g-c | 드라이버 복원: 로그 열기 전 3줄(선택 폴더·백업/현재 시스템명) 버려짐, 불일치 '아니오'면 로그 없음 | ezBAK.py `run_driver_restore` 3947-3949·3963 → 로그 열기 3976 | 3줄을 로그 연 직후로 옮김. '아니오'는 지금처럼 로그 없음(P3 fix-e 와 같음) | △ 작업 로그 3줄 추가(문구 그대로) | tool_cases: driver_restore_mismatch_yes·rc_nonzero·popen_raises·uppercase_marker.logs | 지금(F6) |
| g-d | 대문자 `_DRIVERS_BACKUP_` → "Unknown", 불일치 경고 없음 | targets/drivers.py `parse_backup_folder` 33-41 | 표지 위치를 소문자로 찾아 원문에서 앞부분 자름 | △ 그 경우만 복원 로그 이름 `Unknown_…`→`<시스템>_…`, 경고 대화상자 생김 | tool_cases: driver_restore_uppercase_marker(전 키) | 지금 |
| g-e | 불일치 askyesno 를 작업 스레드에서 호출 | ezBAK.py `run_driver_restore` 3961 | P5 gui 의 대화상자 요청(복원 `_wait_*` 폴링과 같은 방식)으로 | – | tool_cases: driver_restore_mismatch_* (P5-fix 때) | P5 |
| g-f | 드라이버 로그 첫 줄에 무관한 `[POLICY] Hidden=… System=…` | ezBAK.py `run_driver_backup` 3827·restore 3980 | — | ● 작업 로그 줄 삭제 | — | 그대로(F6) |
| g-g | NAS 실패·해제 대화상자에 상세 문구 빠짐(en 은 "Failed to connect." 만) | ezBAK.py en 번역 사전 중복 키 287·298·299 가 223·231·232 의 `{0}` 판을 덮음 | 그 3줄 삭제(호출처는 이미 `.format`). ko 는 정상 | – (대화상자, 키 이름 그대로) | tool_net: nas_connect_rc_code·rc_err·rc_out, nas_disconnect_drive_ok·rc_code·rc_err .dialogs | 지금 |
| g-h | 재입력 흐름이 첫 대화상자의 드라이브 문자를 잊음 | ezBAK.py `open_nas_connect_dialog` except 3408-3475 | g-k 로 흐름 자체 삭제 | – | g-k 와 같음 | 지금(F1, g-k 커밋) |
| g-i | 해제가 예외면 같은 명령 4회 | ezBAK.py `open_nas_disconnect_dialog` except 3566-3596 (+run_net_use 안 1회) | g-k 로 2회 | – | g-k 와 같음 | 지금(F1, g-k 커밋) |
| g-j | NAS out·err 줄이 열린 로그가 없어 버려짐 | ezBAK.py 3383-3392 `write_detailed_log`(2036, `_log_file` 없으면 return) | 진단 로그 전환 때 `_log.info` 로(암호 없음 확인) | – | — | P6 |
| g-k | 예외 뒤 자격 증명 재요청·같은 명령 재실행, 그 안의 예외는 메서드 밖으로(Tk 콜백 traceback) (T0-3) | ezBAK.py 연결 3408-3475·해제 3566-3596 | except 는 큐 로그 + showerror 후 return. 재입력·재실행 블록 삭제(약 −85줄) | – (큐 문구 그대로, 줄만 줄어듦) | tool_net: nas_connect_retry_ok·retry_raises·nas_disconnect_retry_raises / argv: nas_connect_retry_path·nas_disconnect_retry_path | 지금(F1) |
| g-l | winget simple list rc≠0 은 로그 없음 | targets/apps.py `_try_simple` | `_log.warning` 로 rc(작업 로그 문구 안 늘림) | – | — | P6 |
| g-m | winget·프로필 로그에 백업용 Policy/Filters 머리말 | ezBAK.py `open_log_file` 2048(prefix 공용) | — | ● 작업 로그 머리말 | — | 그대로 |
| g-n | 프로필: 공간 부족에도 상태 "complete!", 빈 `BrowserProfiles_backup_<ts>` 남음 | targets/browser.py `backup_profiles` 222-232(폴더 먼저 만듦)·ezBAK.py `run_browser_profiles_backup` finally 3783 | 공간 확인을 `free_space(dest_root)` 로 폴더 만들기 전에. 실패 경로 상태는 기존 문구 "Operation finished" | – (산출물 이름 그대로, 안 만들 뿐) | tool_profiles: insufficient_space-*(tree·queue), disk_usage_fail-*(queue) | 지금(F5) |
| g-o | 프로필: 원본이 하나도 없어도 정상 완료·빈 폴더·폴더 열기 | 같은 함수 219-223 | 원본 0개(ini 포함)면 폴더 안 만들고 None, wrapper 가 show_error(새 대화상자 문구 1개)·"Operation finished" | – (작업 로그 "total bytes: 0" 줄까지 현행) | tool_profiles: no_source-*(tree·queue·returned·logs) | 지금(F5) |
| g-p | 프로필 원본이 실행 계정 %LOCALAPPDATA%·%APPDATA% | ezBAK.py 3764-3767 → browser.py `PROFILE_SOURCES` | F4=B 면 `USERS_ROOT/<user_var>/AppData/{Local,Roaming}` | – | tool_profiles 전 경우(F4=B 때) | 그대로(F4) |
| g-q | STDERR·traceback 줄 뒤 빈 줄 | drivers.py `run_pnputil` 72, ezBAK.py 4005 | — | ● 작업 로그 형식 | — | 그대로(F6) |
| g-r | pnputil `cp949` 하드코딩·문자열+shell=True(`%VAR%` 가 든 경로는 cmd 가 전개, 비한국어 로캘은 UnicodeDecodeError 로 "Driver backup error") (T0-3) | drivers.py `build_*_cmd`·`run_pnputil` 18-25·57 | `build_*_cmd` 가 `PnputilCmd(display, argv)` 반환, 실행은 argv·shell=False·encoding='oem'·errors='replace', 로그 "Executing command:" 는 display(현행 글자) | – (작업 로그 그대로. 한국어 Windows 는 oem=cp949 로 동일) | argv: driver_backup·driver_restore / tool_cases: driver_* .commands | 지금(F2) |
| g-s | wmic 3곳 퇴역(24H2 기본 제거) → 레지스트리 BaseBoardProduct 가 없으면 "Unknown" (T0-3) | winsys/sysinfo.py `_WMIC_QUERIES`·`system_name` 61-110 | 같은 BIOS 키의 SystemProductName·SystemManufacturer 를 wmic 앞에. 값 없음은 DEBUG 출력 안 함 | △ BaseBoardProduct 없는 PC 만 드라이버 폴더 이름 값이 바뀔 수 있음(같은 SMBIOS 출처) | 기존 키 무변경(확인), 새 경우는 단위 테스트로 | 지금(F3) |
| g-t | net use 예외면 인코딩 없이 한 번 더(T0-3 재실행의 안쪽) | storage/nas.py `run_net_use` 144-149 | — 예외 때만 1회, P4 D3 로 자리 정함 | – | — | 그대로 |
분류 합계: 지금 11(a·c·d·g·h·i·k·n·o·r·s) · P5 1(e) · P6 2(j·l) · 그대로 6(b·f·m·p·q·t). 표 밖 메모(P5): en 사전 중복 블록 266-303 의 나머지 키와 f-string 호출처(`task_deleted`·`failed_to_delete_task`·`filter_manager_error`, P3 발견)는 번역을 gui 로 옮길 때 같이

## 2. 결정 (사용자) — 확정 2026-10-13: F1 A · F2 A · F3 B · F4 C(→ P5, p5-plan 에 결함 커밋으로) · F5 A · F6 A
- F1 NAS 예외 흐름(g-h·i·k): A(권장) 오류 표시 후 끝, 재입력·재실행 삭제 / B 재입력 유지, 두 번째 예외만 잡음 / C P5 로 — 예외는 net.exe 를 못 띄운 경우라 자격 증명을 다시 물어도 소용없고, 지우면 P5 로 옮길 코드도 85줄 준다
- F2 pnputil(g-r): A(권장) argv·shell=False·'oem'·replace, 로그 줄 현행 / B 인코딩만 'oem'·replace / C 그대로 — 한국어 PC 출력·로그는 같고, 다른 로캘의 디코딩 예외와 cmd 전개만 사라진다
- F3 wmic(g-s): A 레지스트리 값으로 교체·wmic 삭제 / B(권장) 레지스트리 2값을 앞에, wmic 은 마지막에 남김 / C 그대로 — B 는 기존 golden·폴더 이름이 바뀌는 경우가 가장 적다
- F4 프로필 원본 사용자(g-p): A(권장) 실행 계정 유지 / B 선택한 사용자 프로필 / C B + 시작 때 확인 대화상자(P5) — 버튼이 사용자 선택과 따로 있고 user_var 기본값이 목록 첫 폴더(Default·Public 일 수 있음)라 말없이 바꾸면 엉뚱한 프로필을 백업한다
- F5 프로필 실패 결과(g-n·o): A(권장) 둘 다 / B 공간 부족만 / C 그대로 — 사용자에게 성공으로 보이는 빈 폴더가 둘 다 같은 원인이고 한 커밋이다
- F6 작업 로그 내용 3건(g-c·f·q): A(권장) g-c 만(빠진 줄 살림) / B 셋 다 / C 셋 다 그대로 — f·q 는 줄을 지우는 C3 변경인데 얻는 것이 모양뿐이다

## 3. 발주 순서 (ezBAK.py 커밋은 순차, 같은 golden JSON 도 순차. 권장안 기준)
| 차 | 발주 | Files(≤5, golden 제외) | 추가·수정 테스트 |
|---|---|---|---|
| 1 | fix-g ∥ fix-s ∥ fix-d | g: ezBAK.py, tests/characterization/test_tool_cases_net.py(docstring) · s: winsys/sysinfo.py, tests/test_winsys_sysinfo.py · d: targets/drivers.py, tests/test_targets_drivers.py | g: 없음(golden) · s: SystemProductName 적중·Manufacturer 적중·둘 다 없으면 wmic 순서·DEBUG 무출력, test_tool_cases sysinfo_* 무변경 확인 · d: parse 대문자·섞인 대소문자 param |
| 2 | fix-np(F5) ∥ fix-r(F2, fix-d 뒤) | np: targets/browser.py, ezBAK.py(+3 이하), tests/test_targets_browser.py, tests/characterization/test_tool_cases_profiles.py · r: targets/drivers.py, tests/test_targets_drivers.py, tests/characterization/test_commands_argv.py | np: 공간 부족이면 폴더 없음·원본 0개 None·free_space 인자 dest_root · r: argv·shell=False·encoding·display 로그 줄, `%` 든 경로가 argv 에 그대로 |
| 3 | fix-k(F1, fix-g·r 뒤) | ezBAK.py(−85), tests/characterization/test_tool_cases_net.py(pytest.raises 제거), tests/characterization/test_commands_argv.py | 연결·해제 예외 각 1건: 명령 2회·대화상자 showerror 1개·예외 안 올라옴·암호 비노출 |
| 4 | fix-a | ezBAK.py | 없음(golden devmgr_all_fail) |
| 5 | fix-c(F6=A, fix-d 뒤) | ezBAK.py | 없음(golden 4건 logs) |
- golden 재생성 파일: g·k → tool_net, d·r·a·c → tool_cases, r·k → argv, np → tool_profiles. ezBAK.py 증가는 np 만(+3 이하), 나머지는 0 또는 감소
- 끝나면 메인: roadmap P4-fix 행, storage CLAUDE.md(재입력 흐름 줄 삭제)·targets(PnputilCmd·parse·backup_profiles None)·winsys(system_name 순서) 갱신, roadmap "P4 뒤 정리"의 NAS 중복 블록 항목 삭제

## 4. 완료 기준
- `python -m pytest tests -q -W error` 전부 통과(729 + 새 테스트), golden 변경은 위 표의 키뿐(커밋마다 키 단위 비교)
- folder-check 위반 0(`py -3 "$HOME/.claude/fable/hooks/folder-check.py"`), ezBAK.py `subprocess` 0줄·5820줄 이하, 아키텍처 검사 a~c 통과
