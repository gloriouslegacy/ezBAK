# ezBAK 리팩터링 로드맵

- 기준 문서: `docs/requirements/refactoring.md`, `docs/adr/0001~0003`, `docs/structure.md`
- 브랜치: `refactor/package-split`(P1~P5-fix) → `claude/continue-work-vipjnn`(U0~U6) → `claude/continue-work-r3vedp`(U8·P6·U7). 모두 main 병합(#45~#48) 뒤 2026-10-27 삭제. 다음 작업은 main 에서 새 브랜치.
- 괄호 안은 이전 계획서(`ezBAKrefactoringPlan.md`, 2026-09-03, 이 문서로 흡수 후 삭제)의 Phase 번호다.

## 단계

| 단계 | 내용 | 완료 기준 | 상태 |
|---|---|---|---|
| (0) 사전 정비 | 중복 함수·고아 함수 제거, DEFAULT_FILTERS 상수화('NAME'→'name'), Headless dead code, is_hidden 버그, 미사용 import | py_compile 통과, T0-1 테스트 통과 후 커밋 | 완료(99e4d37) |
| T0 테스트 기반 | `tests/`, pytest, `requirements-dev.txt`. 필터·숨김 판정 3벌, GUI·Headless 임시 폴더 백업 golden, 명령줄 스냅샷, GUI 스모크. 동작을 바꾸지 않는 이음매(`C:\Users` 상수 등)만 추가. `test_is_hidden.py` 를 tests/ 로 옮겨 pytest 로 바꾼다 | 손대지 않은 모놀리스에서 pytest 통과 | 완료(e72adcf, 101 passed) |
| P1 winsys + core 기반 | `proc`, `fileattrs`, `paths`, `model`, `filters`, `fileops` 를 추출해 3벌 중복을 1벌로 | 특성 테스트가 golden 변경 없이 통과, 아키텍처 검사 a~c 통과 | 완료(b00ace3, 283 passed) |
| P2 엔진 통합 (1) | `engine`, `joblog`, `retention`, `storage.local`, `targets.browser`(북마크). GUI·Headless 백업은 BackupPolicy 플래그를 쓰는 thin wrapper 로 | 두 경로의 golden 이 각각 그대로 | 완료(dac806d, 402 passed) |
| P2-fix 결함·Q1 예외 | 북마크를 실행 계정에서 복사하는 문제, Headless 삭제 핸들러 인자 불일치, Headless 작업 로그 줄바꿈 누락(Q1 예외). 항목마다 별도 커밋. 호출처 없던 모듈 수준 `_handle_rmtree_error` 는 P1-4 에서 삭제함 | 커밋마다 해당 golden 갱신과 사유 | 완료(fb42ab2·382936d·3b3b52b, 405 passed) |
| P3 복원 + automation (2·5 일부) | `core.restore`, `cli`, `headless`, `schedule_config`, `scheduler` | CLI·스케줄 JSON 왕복 테스트, schtasks·XML 명령 스냅샷이 그대로 | 완료(6859545, 575 passed) |
| P3-fix 결함·Q1 예외 | 복원 후보 순서, free_space 실패 판정, 없는 get_free_space, 북마크 복원 대상, XML 이스케이프·파일 이름, 실패 시 남는 JSON·XML·복원 로그, 삭제 때 XML | 커밋마다 해당 golden 갱신과 사유 | 완료(4187014~3765d57, 594 passed) |
| P4 대상·저장 나머지 (3) | drivers, winget, 브라우저 프로필, NAS | proc 모킹 argv 스냅샷이 그대로 | 완료(27b258f, 729 passed) |
| P4-fix 결함·Q1 예외 | NAS 대화상자 상세 문구·예외 흐름, 드라이버 표지 대소문자·복원 로그 3줄, pnputil argv·oem, wmic 앞 레지스트리, 프로필 백업 빈 폴더, 장치 관리자 `{0}` | 커밋마다 해당 golden 갱신과 사유 | 완료(a2b6f53~1375246, 737 passed) |
| P5 GUI 분리 (4·5) | `gui/*`, `dialogs`, `bridge` | GUI 스모크 통과, `ezBAK.py` 100줄 이하, core 의 tkinter import 0 | 완료(951bd49, 881 passed, ezBAK.py 69줄) |
| P5-fix 결함·Q1 예외 | 단축키 바인딩(Ctrl/Alt+숫자, `<Space>`), 번역 `{0}`·en 중복 키, 복사 총량·진행·완료 상태, 스케줄·NAS·원본 선택 대화상자, 맨 except·삼킨 예외·DEBUG print 일부 | 커밋마다 해당 golden 갱신과 사유, 정리 커밋은 golden 무변경 | 완료(e784dde~fc34320, 883 passed) |
| U0~U8 UI 리디자인 | CustomTkinter 사이드바 + 페이지, 토큰·공용 위젯, 대화상자 새 스타일. 단계·결정은 `docs/nav-drafts/ui-stages.md`, 범위·golden 정책은 ADR 0004. 브랜치 `claude/continue-work-vipjnn` | 동작 golden 무변경, 창 구조 golden 만 Q1 예외 재생성, U8 실기기 체크리스트 | U0~U8 완료(main #46), U7 새 데이터 카드 완료(2026-10-27, ui-plan §10 — 해시는 그곳). 실기기 체크리스트 남음 |
| P6 로깅·빌드 (6) | 진단 로그 전환(`%APPDATA%\ezBAK\logs`), 검사 d 켜기, spec·로컬 빌드를 `release.yml` 에 맞춤, `release.yml` 에서 빌드 전에 pytest, 기술문서 갱신 | 검사 d 통과, 로컬 PyInstaller 빌드 성공, CI pytest 단계 | 완료(P6-1~10, P6-7 은 그대로 둠 결정 C. 로컬 빌드·진단 로그 파일 실기기 확인 남음. 계획 `docs/nav-drafts/p6-plan.md`) |

## T0 발주 순서

1. T0-1: tests 기반(conftest, pytest 설정, requirements-dev.txt)과 필터·숨김 특성 테스트. `ezBAK.py` 는 고치지 않는다. 통과하면 Phase 0 를 먼저 커밋하고, 그다음 T0-1 을 커밋한다.
2. T0-2: 이음매 추가와 GUI·Headless 임시 폴더 백업 golden. 완료(979d972)
3. T0-3: 외부 명령 argv 스냅샷(proc 모킹). 호출처 30곳 모두 고정. 완료(82b507b)
4. T0-4a: 한 프로세스에서 Tk 를 여러 번 만들 때 생기는 간헐 실패 제거(공용 `gui_app` 픽스처). 완료(2e4dc10). T0-4b: GUI 스모크(App 생성·위젯·큐 13종). 완료(b2173f9)
5. T0-5: `test_is_hidden.py` 이전. 완료(031c6de)
6. T0-6: 복원 golden. 복원 경로에 남은 하드코딩 `C:\Users` 2곳도 `USERS_ROOT` 로 바꾼다(T0-2 이음매). P3 의 `core.restore` 추출 전에 필요하다. 완료(e72adcf)

## P1 발주 순서

1. P1-1: winsys `proc`·`fileattrs`·`paths` 와 단위 테스트(d1260d9). ezBAK.py 의 subprocess 호출처 30곳은 그 영역을 추출하는 P3·P4 에서 `proc` 으로 옮긴다
2. P1-2: `tests/test_architecture.py` 검사 a~c, 표준 라이브러리 이름 충돌, `ezbak`/`ezBAK` import 대상(ff5628f)
3. P1-3a: `core.model`, `DEFAULT_FILTERS` 단일 정의, BackupPolicy 플래그 7개(b7dfa69)
4. P1-3b: 판정 3벌을 `core.filters` 로 통합(b394189). `App.is_hidden` 은 `replace(BackupPolicy.GUI, …)` 변형
5. P1-4: `core.fileops`(b00ace3). 작은 동작 차이 2건(Headless 크기 계산의 맨 `except:` → `except Exception`, 큐 put 실패 시 계산 중단)은 커밋 메시지에 적었다

## P2 발주 순서

상세(줄 대응 표, wrapper 가 유지할 이름, 발주별 위험)는 `docs/nav-drafts/p2-plan.md`. 결정 D1~D4 는 ADR 0002 에 기록했다(플래그 4개 + JobLogFormat, 로그·폴더 준비는 wrapper, Reporter `progress_max`, `sysinfo.free_space` 를 P2 로).

0. P2-0: 손대지 않은 코드로 경우별 golden(오래된 백업·로그 정리, 북마크 둘 다, 같은 이름 폴더, 사용자 없음, GUI 여유 공간 부족, GUI 큐 순서). 테스트만(5e40b9c)
1. P2-1: `core.joblog`(JobLog·JobLogFormat·format_bytes)와 BackupPolicy 플래그 4개, 두 머리말을 joblog 로(b7a70db)
2. P2-2: `winsys.sysinfo.free_space` 로 `get_free_space_reliable` 위임(bdecfa9)
3. P2-3: `core.ports` Protocol 과 `storage.local.LocalFolderStore`(ec36ee9)
4. P2-4: `core.retention` 으로 두 정리 이동, `rmtree(onerror=…)` → `rmtree_readonly`, 플래그 `isolate_backup_delete_errors`(1e6f398)
5. P2-5: `targets.browser.Bookmarks` 로 북마크 이동(48cafe9)
6. P2-6: `core.engine.run_backup` 과 `core_run_backup` thin wrapper(28d98f1)
7. P2-7: `App.run_backup` thin wrapper. 게이트(ezBAK.py 작업당 +30줄)로 GUI 어댑터는 `ezbak/gui/bridge.py` 에. 플래그 `isolate_retention_errors`·`return_critical_errors`(de69299)
8. P2-8: 미사용 `import stat`·`FILE_ATTRIBUTE_*` 4개·Headless 중첩 판정 함수 정리(dac806d)

## P3 발주 순서

상세(줄 대응 표, 공개 진입점 초안, wrapper 가 유지할 이름, 발주별 위험)는 `docs/nav-drafts/p3-plan.md`. 결정 D1~D7 은 모두 권장안 A(2026-10-12 사용자 결정): D1 복원 3단계 분할·D2 속성 폴링 유지(ADR 0002 기록), D3 CLI 는 손대지 않은 `__main__` 을 runpy 로 돌린 golden + GUI 생성 JSON → CLI 왕복, D4 argparse 그대로, D5 쓰기 `ScheduleSpec`·읽기 dict 무변환, D6 `show_notification` 삭제(P3-10), D7 필터 type 은 받은 값 그대로 저장(ADR 0003 §6 기록).

0. P3-0a: 손대지 않은 코드로 복원 경우별 golden 19건. 테스트만(ac93665)
1. P3-0b: 손대지 않은 코드로 CLI golden 12건(008c2c8, runpy 하네스 `run_ezbak_main`), 예약 작업 생성·삭제 21건과 GUI JSON → CLI 왕복(a8b6637). 테스트만, 459 passed
2. P3-1: `winsys.elevation`(is_admin·관리자 재실행) + proc 재노출(d091622). is_admin 의 맨 except → except Exception
3. P3-2: `core.restore.validate_selection`·`locate_backup`, 폴더 선택 폴링은 `App._wait_restore_folder_choice`(961aa0f)
4. P3-3: `core.restore.run_restore` + GUI wrapper, 교차 확인 폴링·`get_free_space` 대체 경로는 App 에(87825ae)
5. P3-4: 북마크 복원 → `targets.browser.restore_bookmarks`(158a5bc)
6. P3-5: `automation.schedule_config`, `_load_schedule_config` 삭제(147f21c)
7. P3-6: `automation.scheduler` 순수 함수, `_build_auto_backup_command`·`_validate_time_input`·`timedelta` import 삭제(b192d29)
8. P3-7: 예약 작업 생성·삭제를 `scheduler.create_task`·`delete_task` 와 `proc.run` 으로(e4eae31)
9. P3-8: `automation.headless.run_backup`, `core_run_backup` thin wrapper, `HeadlessBackupRunner` 삭제(26a1f27). 아키텍처 검사 내부 파일 판정의 대소문자 오탐 수정(49f7cc4)
10. P3-9: 파일 수 때문에 둘로 나눔 — 9a `automation.cli`(4e6eece), 9b `headless.run`·`schedule_config.job_from_config`, `__main__` 위임, `HeadlessLoggerQueue` 삭제(1cde0d8)
11. P3-10: `show_notification` 삭제 + argv golden `toast_*` 키 제거(6859545, D6 Q1 예외). 575 passed, ezBAK.py 7172 → 6292줄

P3 계획 중 새로 찾은 것(보존, 결함으로 고칠지는 사용자가 정한다): 복원 중 디스크 조회 실패 시 없는 `self.get_free_space` 호출로 "An error occurred during restore", 백업 우선순위 목록을 `sorted(reverse=True)` 로 다시 정렬, 복원 북마크도 실행 계정 %LOCALAPPDATA%, 스케줄 생성 로그는 열린 로그가 없으면 버려짐, `--user` 만 주면 GUI 시작, config 값 형 변환 없음. P3-0a(ac93665): 교차 사용자 '아니오'·공간 부족·disk_usage 실패로 중단돼도 `*_restore_<ts>.log` 가 남음, disk_usage 실패 문구는 `'_tkinter.tkapp' object has no attribute 'get_free_space'`. P3-0b: `--user` 만 준 비관리자는 `sys.exit(0)` 을 잡아 `os._exit(0)`, config 사용자 없음이면 dest 에 빈 백업 폴더만 남음, `--filters` 리스트는 오류인데 종료 코드 0, `--auto-backup` 은 무시됨, XML 파일 이름은 정규화 안 됨(`:` 이면 NTFS 대체 스트림에 씀, `<>|?"*` 이면 JSON 만 남고 실패), schtasks 실패·타임아웃 때 JSON·XML 이 남음, 삭제는 XML 을 안 지움, APPDATA 없어도 삭제 성공 보고, 삭제 대화상자 문구 `Failed to delete task::`·`deleted.: 'name'`

P3-fix(2026-10-12 사용자 결정, 계획 `docs/nav-drafts/p3-fix-plan.md`, F2~F5 모두 A): fix-a 복원 후보 대상 사용자 먼저·최신 먼저(4187014), fix-f XML 값 이스케이프·XML 파일 이름 safe_name(fd81edb, ADR 0003 §6 한 줄), fix-b sysinfo API 실패 시 드라이브 루트로(3fc8e7e, F2), fix-c 복원·프로필 백업 공간 확인을 get_free_space_reliable 로(5e4ea38), fix-g 생성 실패 시 JSON·XML 되돌림·타임아웃은 그대로(b79cfe0, F3), fix-h 삭제 성공 뒤 XML 삭제(4560128, F5), fix-d 북마크 복원을 대상 사용자 프로필로(a24e674), fix-e 복사 전 중단이면 복원 로그 삭제(3765d57, F4). 테스트 파일 닫기 정리(046ff64). 남은 것: `winsys.paths.local_appdata()` 호출처 없음(정리 후보), ezBAK.py 진단 `_log` 에 핸들러가 없어 GUI 경로 경고가 어디에도 안 남음(P6 diag). P3 뒤 정리(P5): App 의 복원 콜백 2개(`_wait_restore_folder_choice`·`_wait_cross_user_confirm`)를 gui 로, `core_engine_run_backup` 별칭·`HeadlessReporter` 별칭 정리 여부

## P4 발주 순서

상세(줄 대응 표, golden 경우, 발주별 위험)는 `docs/nav-drafts/p4-plan.md`. 결정 D1~D7 은 모두 권장안 A(2026-10-12 사용자 결정): D1 targets·storage 는 명령 조립·실행·출력 해석·작업 로그 줄까지, 큐·대화상자는 wrapper(P3 D1 과 같음), D2 `winsys.sysinfo.system_name()` 통째, D3 인코딩 없는 재실행은 `storage.nas.run_net_use` 안·자격 증명 재입력 흐름은 wrapper(P5 에 dialogs/nas), D4 도달 불가 `except` 절 2곳은 옮기며 삭제, D5 browser.py 에 `PROFILE_SOURCES` 별도 표, D6 프로필 진행률은 GuiReporter 재사용, D7 P4 끝에 ezBAK.py `import subprocess` 삭제·검사 c 루트 적용.

P3-fix 반영: 프로필 백업의 없는 `self.get_free_space` 는 fix-c(5e4ea38)에서 고쳤다. P4-0c golden 은 고친 코드로 찍는다(계획서 P4-6 의 "AttributeError 재현"은 해당 없음). 새 결함 후보: 프로필 백업 원본이 실행 계정 %LOCALAPPDATA%·%APPDATA%(북마크와 달리 대상 사용자 아님).

0. P4-0a: 손대지 않은 코드로 시스템 이름·드라이브 문자·드라이버·장치 관리자 golden 22건(9a66b35, tool_cases, tool_harness.py)
1. P4-0b: winget·NAS golden 24건 + 암호 비노출 단언(ec841ae, tool_net)
2. P4-0c: 프로필 전체 백업 golden 6×2건(9d6afa0, tool_profiles, fix-c 뒤 코드)
3. P4-1: `winsys.sysinfo.system_name`·`used_drive_letters`·`clean_name`, `proc.PIPE`(cee40fa)
4. P4-2: `winsys.shell.open_device_manager`(6f9132f)
5. P4-3: `targets.drivers`(6d47b9e)
6. P4-4: `targets.apps`(1f53c3a)
7. P4-5: `storage.nas`, 도달 불가 except 2곳 삭제(D4)(df737ef)
8. P4-6: `targets.browser.backup_profiles`·`PROFILE_SOURCES`, pywin32·ctypes 블록 삭제(21152c9)
9. P4-7: `import subprocess`·`re` 삭제, 검사 c 루트 적용(27b258f). 729 passed, ezBAK.py 6293 → 5820줄, subprocess 0

P4 에서 고정한 이상 동작(처리 결과는 아래 P4-fix 문단): 장치 관리자 오류 문구 `{0}` 미치환, 빈 드라이브 문자가 C~Z 다 쓰이면 'B', 드라이버 복원 — 로그를 열기 전 4줄이 버려짐(불일치 '아니오'면 로그 없음)·대문자 `_DRIVERS_BACKUP_` 은 Unknown 에 경고 없음·불일치 askyesno 를 작업 스레드에서·무관한 [POLICY] 줄, NAS — 대화상자에 상세 문구가 빠짐·재입력 때 드라이브 문자를 잊음·해제 재시도 4회·작업 로그 줄이 버려짐(열린 로그 없음)·재입력 흐름 예외가 메서드 밖으로, winget — simple list rc≠0 은 로그 없음·Policy/Filters 머리말, 프로필 전체 — 공간 부족에도 "complete!" 와 빈 폴더·원본 없어도 정상 완료·원본이 실행 계정 폴더, STDERR·traceback 뒤 빈 줄, (T0-3) pnputil cp949·문자열+shell=True, wmic 퇴역 중, NAS 예외 시 재실행·자격 증명 재요청.

P4 뒤 정리(P5): `_have_pywin32` 테스트 훅 별칭(conftest·특성 테스트가 `fileattrs.HAS_PYWIN32` 를 보게), `App.get_used_drive_letters` 의 ctypes lambda(P4-0a 패치를 sysinfo 쪽으로), `winsys.paths.local_appdata()` 호출처 없음.

P4-fix(2026-10-13 사용자 결정, 계획 `docs/nav-drafts/p4-fix-plan.md`, F1 A·F2 A·F3 B·F4 C·F5 A·F6 A): fix-g en 사전 NAS 중복 키 3줄 삭제(a2b6f53), fix-d 드라이버 표지 대소문자 무시(beae11b), fix-k NAS 예외는 오류 표시 후 끝·재입력·재실행 삭제, g-h·g-i 함께 해소(aca6db4, F1, ezBAK.py −87줄), fix-s 같은 BIOS 키의 SystemProductName·SystemManufacturer 를 wmic 앞에(f9a1fb9, F3), fix-r pnputil argv·shell=False·encoding 'oem'·errors 'replace', 로그 줄은 현행(c1b1256, F2), fix-np 프로필 백업이 공간 부족·원본 없음이면 폴더를 만들지 않고 "Operation finished"(2426751, F5, 새 번역 키 `no_browser_profiles`), fix-a 장치 관리자 `{0}`(d928f39), fix-c 드라이버 복원 로그 3줄을 로그 연 뒤로(1375246, F6). 그대로 둔 것: 빈 드라이브 문자 'B', [POLICY] 줄·Policy/Filters 머리말·STDERR 뒤 빈 줄(C3), net use 내부 재실행 1회. P5 로: g-e 불일치 askyesno 를 작업 스레드 밖으로, g-p 프로필 원본을 선택 사용자로 + 시작 확인 대화상자(F4=C), en 사전 중복 블록 나머지·f-string 호출처. P6 로: NAS out·err 줄(g-j)·winget simple rc≠0(g-l)을 진단 로그로. 확인 필요: 비한국어 로캘 실기기에서 pnputil 출력과 'oem' 디코딩.

P5 진행 중 고정한 이상 동작(결함 후보 → P5-fix 에서 전부 처리, 아래 P5-fix 문단): `bind('<Control-1>'…'<Control-5>')` 는 Tk 가 Ctrl+마우스 버튼 1~5 로 읽음(키는 Ctrl+6~9 만, P5-0b), 복사 `copy_tree` 에서 복사 바이트(145)가 계산 총량(132)보다 큼·pywin32 없으면 크기 계산의 `is_hidden_safe` 에 ctypes 대체가 없어 숨김 파일도 셈·복사 최종 update_progress 가 막대 기본 최대값 100(P5-0d), 원본 선택 대화상자 `<Space>` 바인딩이 잘못된 키 이름이라 Ctrl+E·Ctrl+C·F5·F1 이 안 묶임·스케줄 대화상자 `<Alt-1/2/3>` 도 마우스 버튼·번역 `failed_to_create_schedule`·`overwrite_warning` 의 `{0}` 미치환·NAS 해제 대상이 비면 "Error"/"Error"·보관 개수가 숫자가 아니면 action="create" 가 이미 설정돼 "No result data found" 분기·NAS 드라이브 상자 활성 상태가 창이 매핑될 때만 갱신·원본 선택 결과 순서가 set 순서이고 `/`·`\` 섞임(P5-0c), en 사전 중복 키 25개 중 값이 다른 6개(`error_saving_log`·`failed_to_delete_task`·`task_deleted`·`source_selection_failed`·`filter_manager_error`·`invalid_drive_letter`)는 뒤의 `{0}` 없는 값이 쓰임(P5-1, 사전은 그대로 옮김), 원본 선택 `_on_open` 의 숨김 판정 예외를 pass 로 삼킴·`_refresh_tree` 의 checked_paths 미사용(P5-6), 스케줄 on_create·run_copy_thread 의 지역 `from datetime import datetime`(mod.datetime 패치가 안 닿음)·DEBUG print(P5-7·11), open_log_file 의 맨 `except:` 2곳(P5-8b), 옮긴 코드의 `except Exception: pass` 중 위젯 꾸미기가 아닌 것(settings 저장 실패 등, 주석만 달고 그대로).

P5(2026-10-13, 계획 `docs/nav-drafts/p5-plan.md`, D1~D6 모두 A): P5-0a facade(cf0c947)·0b App golden(4b0c34c)·0c 대화상자 golden(c2621b9)·0d 복사·bridge(48d62b4), P5-1 i18n(870a6b5)·2 theme·widgets(0f65d10)·3 Win11Dialog(57f194a)·4 shortcuts(4293b87)·5 filters·버튼 합침(98e7a4e)·6 sources(04f0ee8)·7 schedule 대화상자(320881b)·8a settings(7a152bb)·8b app_log·QueueMixin(35b0197)·9 theme/언어 믹스인·layout(968f8ba)·10 actions_backup·paths.USERS_ROOT(42ac3ba)·11 actions_copy(ccaa885)·12 actions_schedule·entry_script(d0a911b)·13 복원(3877228)·14 actions_tools(4830c33)·15 NAS(8c65787), 결함 g-e 드라이버 불일치 확인을 메인 스레드로(60b791d)·g-p 프로필 원본 선택 사용자 + 확인(be0f108, F4=C), P5-16a App → gui.app·gui.run(84eba48)·16b 루트 69줄·facade 별칭(5876d0c)·17a 루트 검사 a·b, 하위 모듈은 `__all__` 에 있을 때만(30a6a2f, 사용자 결정 A, ADR 0002)·17b 우회 import 정리(951bd49). 테스트 결함 수정 f6cb460. ezBAK.py 5738 → 69줄, App = 믹스인 13개 + tk.Tk. 테스트 결함 수정: profiles_disk_usage_fail-nopywin32 가 실제 디스크를 읽던 패치(f6cb460).

P5-fix(2026-10-13, 계획 `docs/nav-drafts/p5-fix-plan.md`, D1~D7 A·D8 h-u A·h-v A·h-w B·h-x B, 위 결함 후보 전부 처리): #1 Ctrl+1~5 키 바인딩·도움말 "(Not Working)" 삭제(1f7631c), #2 복사 총량을 복사와 같은 숨김 판정·파일 원본 크기(dbeb387), #3 원본 선택 `<space>`·스케줄 `<Alt-Key-n>`(e784dde), #4 복사 finally 의 막대 채우기 삭제(bae70d8), #5 en 같은 값 중복 키 20개 삭제(874edcf), #6 값이 다른 6개는 `{0}` 값 유지·호출처 `.format`(58acf99), #7 스케줄 생성 실패 `{0}`·새 키 `dest_folder_missing`(cf03ce4), #8 보관 개수 오류 때 action 미설정(8a91b51), #9 복사·on_create DEBUG print → `_log`(d5c2649), #10 복사 성공 상태 "Copy complete!"(9c12716), #11 NAS 해제 빈 대상 문구(69045ef), #12 NAS 드라이브 상자가 매핑 체크를 따름(d17a748), #12b NAS 대화상자 테스트를 test_gui_dialogs_nas.py 로(ab14b74), #13 원본 선택 결과 트리 순서·`\`(48a076e), 정리 #14 지역 datetime import·checked_paths(1f0b71b)·#15 맨 except 5곳(ae82c2e)·#16~18 삼킨 예외 30곳을 진단 로그로(03e74ad·fc34320·05c2da8, NAS 는 error_type 만). ezBAK.py 무변경.

## P6 발주 순서

계획 `docs/nav-drafts/p6-plan.md`, 결정 D1 A·D2 B·D3 A·D4 A·D5 A(2026-10-27 사용자 결정, 모두 권장안). 브랜치 `claude/continue-work-r3vedp`(U8 병합 #46 뒤 main 에서).

1. P6-1: `core.diag`(setup·new_run_id·bind·current, `%APPDATA%\ezBAK\logs\ezbak.log` 1MB×3, INFO/EZBAK_LOG_LEVEL)(2f6555b)
2. P6-2: 진입점 `headless.run`·`gui.run` 에서 setup, Headless run_id·job_start·job_failed(64262aa). ezBAK.py 는 간선(automation·gui·winsys) 때문에 무변경
3. P6-3: GUI 작업 스레드 `begin_job`(run_id·시작 줄·작업 로그 경로) — a 백업·복원·복사·공간 확인(ffc74f2), b 도구·예약(96eb1de)
4. P6-4: g-j NAS net use 출력·g-l winget simple rc≠0 을 진단 로그로(9c0b140)
5. P6-5: print → 진단 로그(D1 A, Q1 예외 stdout 칸만) — a 예약(37ce665+golden f85bcb7), b 원본 선택(092c62a), c 나머지 gui(32bedf9+golden 8f3c6c2), d `system_name`·`create_task` 의 `debug=print` 기본값(90f964d+golden 3163a8a)
6. P6-6: 검사 d(D2 B: ezbak/ 의 print 전부·headless.py 만 예외, 맨 except)(4b29805)
7. P6-7: h-w 그대로 둠(2026-10-27 사용자 결정 C) — tkinter 가 다른 스레드의 cget 을 Tcl 스레드로 넘기고 실패는 except 가 받음. 바꾸면 큐 golden 또는 테스트 간 위젯 상태가 갈림(p6-plan §4)
8. P6-8: `_LOG` → `_log` 4파일, 호출처 없는 App 메서드 2개 삭제(0ea3ca0)
9. P6-9: release.yml 빌드 전 pytest(검사 f), 로컬 ps1 플래그를 release.yml 에 맞춤(e5b4bbd). 로컬 빌드는 Windows 실기기 확인 필요
10. P6-10: 기술문서(§1.3·2.1·8.2·8.3·10.1)·structure·roadmap·mem

## 함께 처리할 것

- 호출처가 없는 코드(`HeadlessBackupRunner`, `_build_auto_backup_command`, `_load_schedule_config`, `show_notification`)는 그 영역을 추출하는 단계에서 지운다. `compute_total_bytes_fast` 는 P1-4 에서 지웠다.
- P1 에서 찾은 것
  - `App.open_filter_manager` 가 두 번 정의돼 있다(뒤쪽이 쓰인다). P5 GUI 분리 때 앞쪽을 지운다
  - 미사용: ezBAK.py 의 `import stat`(맨 위·`core_run_backup` 안), `FILE_ATTRIBUTE_READONLY`·`DIRECTORY`·`ARCHIVE`·`DEVICE` 상수. P2 엔진 이동 때 정리한다
  - 오래된 백업 삭제의 `shutil.rmtree(onerror=…)` 는 P2 `retention` 이동 때 `fileops.rmtree_readonly` 로 바꾼다(`onerror` 는 3.12 부터 deprecated)
  - 결함 후보: GUI 크기 계산(`run_space_calculation`)은 `App.is_hidden` 변형 정책을, 실제 백업은 `BackupPolicy.GUI` 를 써서 pywin32 가 없으면 예상 크기와 복사 대상이 다르다
- T0-1 에서 고정한 판정 차이(`tests/characterization/test_filters.py`·`test_hidden.py` golden). P1 은 BackupPolicy 플래그로 그대로 보존하고, 결함으로 고칠지는 P2-fix 때 사용자가 정한다.
  1. 점 파일·`thumbs.db`·`desktop.ini`·`$recycle.bin` 을 Headless 만 이름으로 무조건 제외
  2. `is_hidden_safe` 에 AppData·Application Data·Local Settings 이름 검사가 없음
  3. `is_hidden_safe` 는 pywin32 가 없으면 속성 검사를 건너뜀(숨김·시스템·정션 포함). 나머지 두 벌은 ctypes 로 대신 검사
  4. 필터가 None 이면 GUI 는 빈 규칙, Headless 는 `DEFAULT_FILTERS`
  5. Headless `matches_filter` 는 경로가 있는지 확인하지 않음
  6. pywin32 가 있으면 ctypes 상수·`_get_file_attributes` 가 정의되지 않음(지금은 그 분기를 타지 않아 드러나지 않음)
- T0-2 백업 golden 에서 확인한 것: 3번 때문에 pywin32 가 없으면 GUI 백업이 NTUSER.DAT·숨김 폴더를 복사하고 정션을 따라 들어간다(파일 수 GUI 7·14, Headless 5). 로그 문구도 GUI 와 Headless 가 다르다. GUI 는 "Skipping hidden directory"·"Backup successfully completed" 를, Headless 는 "Skipping filtered directory"·"Backup start for …" 를 쓴다. 모두 BackupPolicy 로 보존한다.
- 필터 공통 특이 동작(고정함): `ext` 는 마지막 확장자만 비교(`*.tar.gz` 불일치), `[ab]` 미지원, 빈 include 패턴 하나면 전부 제외, `path` 는 절대경로 전체와 비교
- T0-3 에서 고정한 외부 명령의 이상한 점(`argv_golden.json`). P3·P4 에서 추출할 때는 그대로 두고, 결함으로 고칠지는 사용자가 정한다.
  - 예약 작업 XML 의 `<Arguments>`·`<Author>` 를 이스케이프하지 않아 경로에 `&`·`<` 가 있으면 XML 이 깨진다
  - 관리자 재실행에서 만든 `params` 가 쓰이지 않고 `list2cmdline(sys.argv)` 로 덮어써진다
  - NAS `net use` 가 예외로 실패하면 같은 명령을 다시 실행하고, 연결 쪽은 자격 증명을 다시 묻는다
  - shell·인코딩이 제각각이다: schtasks 삭제만 `shell=True` 에 인코딩 없음, toast 는 리스트 argv 에 `shell=True`, pnputil 은 `cp949` 하드코딩
  - `wmic` 3곳은 최신 Windows 에서 빠지는 중이다
- T0-4b·T0-6 에서 고정한 것. 결함으로 고칠지는 사용자가 정한다.
  - 복원은 숨김·시스템을 둘 다 포함으로 하지 않으면 AppData 를 건너뛴다(백업에는 들어 있음)
  - 백업할 때 숨김·시스템 속성이 사라져, 복원하면 일반 파일이 된다
  - `run_restore` 의 `_restore_completed` 를 설정하는 곳이 없어 마지막 상태가 항상 "Operation finished" 다
  - 로그 정리 단계가 `"\n[Old log cleanup]"` 을 써서 한 항목이 두 줄로 나뉜다
  - `enable_buttons`·`play_sound`·`reset_progress` 의 `DEBUG:` print(P6 검사 d 에서 정리), 메뉴바 tearoff 항목, ezBAK.py `\e` 이스케이프 경고
- P2-0 에서 고정한 것(`backup_cases_golden.json`). 결함으로 고칠지는 사용자가 정한다.
  - 같은 이름 백업 폴더 삭제: GUI 는 `onerror` 없는 `rmtree` 라 읽기 전용 파일에서 실패하고 "Failed to delete…" 를 남긴 채 그 폴더에 백업한다. Headless 는 핸들러 인자 불일치로 조용히 실패한다(P2-fix #2). 둘 다 옛 파일이 새 백업에 남는다
  - 사용자 폴더가 없으면 실패 전에 dest 에 흔적이 남는다: GUI 는 머리말만 있는 로그, Headless 는 빈 백업 폴더(로그 없음, `FileNotFoundError`)
  - GUI 여유 공간 부족이면 빈 백업 폴더가 남는다
- P2 에서 찾은 것
  - P2 전부터 안 쓰이던 것(P2-8 범위 밖): 맨 위 `timedelta` import(3472 지역 import 만 씀), 모듈 함수 `log_message`, App 메서드 `_validate_time_input`·`on_exit`·`set_progress_indeterminate`·`switch_mode`. 그 영역을 추출할 때(P5·P6) 호출처를 다시 확인하고 지운다
  - P2-fix 완료: #1 북마크 원본을 `TargetContext.profile`(대상 사용자 프로필)에서(fb42ab2), #2 Headless 같은 이름 폴더를 `rmtree_readonly` 로(382936d, GUI 쪽 onerror 없는 rmtree 는 위 P2-0 후보로 남음), #3 HEADLESS 항목은 모두 lead 없이 `end="\n"`(3b3b52b)
  - P2-fix 뒤 정리 후보(리팩터링 커밋): `bm_saved_last`·`bm_missing_last` 키와 browser.py 의 마지막 브라우저 분기(이제 일반 항목과 같음), test_targets_browser 의 GUI/Headless 꼬리 parametrize(`winsys.paths.local_appdata()` 는 P3-4 복원 북마크가 다시 쓴다). Headless `bk_error`·`bm_error` 는 golden 경우 없이 단위 테스트만 있다
  - 동작 메모(golden 밖, 커밋 메시지에 기록): GUI 숨김·시스템·필터 설정은 백업 시작 때 읽는다, `hidden_mode_var.get()` 실패는 "Backup error:" 경로, 로그 정리 예외의 Traceback 프레임이 retention.py
  - `bm_header` 문구는 C3 고정이라 브라우저를 늘릴 때 문구를 따로 정해야 한다(targets/CLAUDE.md)
  - GUI bridge(GuiReporter·GuiLogStream) 단위 테스트는 P5 에 `tests/test_gui_bridge.py` 로
- 구조 점검 위반: `ezbak/` 에 안내 파일이 없다(코드는 빈 `__init__.py` 하나). P1 끝에 부채(`--kind 안내`) 등록으로 정했다(사용자 결정, 2026-10-12, ADR 0003 §4).
- P5-fix 에서 남긴 것
  - P6 로(D8 h-w): 백업·복원·도구의 finally 가 작업 스레드에서 `progress_bar['maximum']` 을 읽는다(actions_backup·actions_restore·actions_tools 4곳, golden 4파일)
  - 그대로(D8 h-x): 메인 창 Ctrl+C 가 로그 글상자에서도 복사 기능을 띄운다(단축키 체계 변경 = UI 동작)
  - P6 검사 d: gui 의 나머지 print(actions_schedule 28 등, dialogs/schedule 의 위젯 꾸미기 print 3곳 포함), bridge.py 로거 이름 `_LOG`(다른 모듈은 `_log`)
  - 테스트: golden 재생성(`EZBAK_REGEN_GOLDEN=1`) 때만 가끔 `OSError Errno 22`(Windows 파일 잠금 추정, 다시 돌리면 통과). test_gui_dialogs·_nas 가 작은 helper 를 따로 갖고 있다(커지면 fixtures 로)
- GUI·Headless 동작을 하나로 통일할지는 리팩터링이 끝난 뒤 따로 결정한다.
- 범위 밖 정리
  - `scripts/git_push.ps1`: commit 이 실패해도 push 가 실행된다.
  - `scripts/process_monitoring.ps1`: SMTP 비밀번호를 파일에 적는 구조다. 환경 변수로 바꾼다.

## 이력

| 날짜 | 내용 |
|---|---|
| 2026-09-03 | 이전 계획서 작성, Phase 0 적용(미커밋) |
| 2026-10-12 | V5 적용: 인터뷰, 요구사항, ADR 0001~0003 확정, 모듈 안내, scripts/ 이동(3f3a762) |
| 2026-10-12 | Phase 0 커밋(99e4d37). T0 완료: 특성 테스트 101개(필터·숨김, 백업, 외부 명령 30곳, GUI 스모크, 복원). 이음매는 `USERS_ROOT` 하나 |
| 2026-10-12 | P1 완료: winsys 기반·core model/filters/fileops, 검사 a~c. 테스트 283개, golden 변경 없음, ezBAK.py 8203 → 7699줄 |
| 2026-10-12 | P2 완료: 경우별 golden, core joblog/ports/retention/engine, storage.local, targets.browser, winsys.sysinfo, gui.bridge. GUI·Headless 백업은 엔진 위 thin wrapper. 테스트 402개, golden 변경 없음, ezBAK.py 7699 → 7178줄 |
| 2026-10-27 | U0: ADR 0004(UI 리디자인 범위·CustomTkinter·golden 정책), 루트 CLAUDE.md 하지 않을 것·작업 브랜치, roadmap U 단계 |
