# P3-fix 계획 (초안 2026-10-12 · 기준 0500c55) — 결함 5건, 커밋 = 발주 1건, golden 은 그 커밋에서만(Q1 예외, 사유를 커밋 메시지에)
공통: ezBAK.py 는 발주당 +30줄 이하·CRLF 유지(core/restore.py 도 CRLF). USERS_ROOT·datetime 은 호출 시점 조회. 실제 schtasks·UAC 없음(record_subprocess·block_subprocess). golden JSON 은 EZBAK_REGEN_GOLDEN=1 재생성물이라 손으로 고치는 파일 5개 셈에서 뺀다. 새 진단은 `_log`(ADR 0003 §2)로 쓰고 작업 로그·큐 문구는 늘리지 않는다

## 결함 1 복원 북마크 → 대상 사용자 프로필 (fix-d)
- 현재: ezBAK.py:2887 이 `winsys_paths.local_appdata()`(실행 계정)를 넘김, targets/browser.py:90-113 `restore_bookmarks(backup_root, out, local_appdata)` 는 `<local_appdata>/<relative>` 로 복사
- 대상 사용자 = `self.user_var.get()`(run_restore 의 user_name). 같은 값으로 복원 폴더 `os.path.join(USERS_ROOT, user_name)` 을 이미 만든다(2877)
- 고칠 것: `restore_bookmarks(backup_root, out, profile)` 로 바꿔 대상을 `browser.source(profile)` 로(P2-fix #1 백업과 같은 함수). 호출처는 `os.path.join(USERS_ROOT, user_name)`. 쓰지 않게 되는 `winsys_paths` import 삭제. 대상 폴더가 없으면 지금처럼 브라우저별 오류 줄(만들지 않음)
- golden: restore_golden.json 4건(parent·direct × pywin32·nopywin32: tree·file_count·localappdata), restore_cases_golden.json `success`(target·localappdata). 두 테스트 준비 단계에서 브라우저 폴더를 대상 프로필 아래에도 만들고 %LOCALAPPDATA% 쪽은 남겨, 실행 계정에 쓰지 않음을 같이 고정
- 단위: test_targets_browser.py 의 restore 4건을 profile 기준으로, "실행 계정 %LOCALAPPDATA% 에 안 씀" 1건 추가
- 파일: targets/browser.py, ezBAK.py, tests/test_targets_browser.py, tests/characterization/test_restore_golden.py, test_restore_cases.py (+golden 2) · C3 없음(Chrome_/Edge_Bookmarks 이름·dest 위치 그대로)

## 결함 2 없는 `self.get_free_space` — 복원·프로필 백업 두 곳 (fix-c, 선행 fix-b 는 F2)
- 현재: ezBAK.py:2839-2844 `_restore_free_space` 가 disk_usage 실패 시 없는 메서드 호출 → "An error occurred during restore: '_tkinter.tkapp' object has no attribute 'get_free_space'". 같은 결함이 4117-4120(브라우저 프로필 전체 백업의 공간 확인)에도 있음
- 고칠 것: `free_space=self.get_free_space_reliable` 로 넘기고 `_restore_free_space` 삭제(백업 3298·3974 와 같은 경로, 내부는 winsys.sysinfo.free_space — disk_usage 를 먼저 시도하므로 성공 경로 동일). 4117-4120 의 try/except 도 `free = self.get_free_space_reliable(dest_base)` 한 줄로
- 프로필 호출처는 같은 커밋에서 고친다(P4-0c 뒤로 미루지 않음): 원인·고침이 같은 결함 1건이고 지금 이 경로의 golden 이 없다. 미루면 P4-0c 가 결함을 고정했다가 다시 Q1 예외 커밋이 필요하다. P4-0c 는 고친 코드로 찍는다(p4-plan 의 "무변경" 전제는 그대로 성립)
- 주의: sysinfo.free_space 는 GetDiskFreeSpaceExW 반환값을 안 봐서 경로가 없으면 0 을 준다(실측). fix-b 없이 고치면 disk_usage_fail 은 "Not enough free space to restore … Available: 0 B" 로 바뀐다
- golden: restore_cases_golden.json `disk_usage_fail`(queue·logs·target). test_restore_cases.py 머리 docstring 19줄 문구 갱신
- 파일: ezBAK.py, tests/characterization/test_restore_cases.py (+golden 1) · C3 없음
- fix-b(F2=A 일 때): winsys/sysinfo.py 가 반환값 0(실패)이면 드라이브 루트 단계로 넘어감. test_winsys_sysinfo.py `test_ctypes_failure_without_raise_returns_zero` 를 드라이브 루트 기대로. golden 변화 없음(fix-c 앞에 두면 disk_usage_fail golden 이 fix-c 에서 한 번만 바뀜). 파일 2개

## 결함 3 백업 후보 우선순위 (fix-a)
- 현재: core/restore.py:114-119 `order_candidates`(대상 사용자 먼저, 원래 코드 주석 "Prioritize matching backups, but show all") 결과를 `sorted(reverse=True)` 로 다시 정렬. 대화상자는 0번을 기본 선택(ezBAK.py show_folder_selection_dialog)
- 의도: 대상 사용자 백업이 먼저, 각 묶음 안은 최신 먼저(이름의 `<YYYY-MM-DD_HHMMSS>` 역순). 고칠 것: `order_candidates` 가 두 묶음을 각각 `sorted(key=str.lower, reverse=True)` 하고 119줄은 그 순서를 그대로 넘김
- golden: 없음(`multi_select` 는 tester > other 라 이미 같은 순서). 단위: test_core_restore.py `test_parent_choice_is_plain_reverse_sort` → 사용자 먼저·최신 먼저 기대로 교체, 대소문자 섞인 같은 사용자 1건
- 파일: core/restore.py, tests/test_core_restore.py · C3 없음

## 결함 4 XML 이스케이프·XML 파일 이름 (fix-f)
- 현재: scheduler.py:163·167 `str.format` 으로 `<Author>`·`<UserId>`·`<Command>`·`<Arguments>` 에 원문 삽입(`&`·`<` 가 든 경로·이름이면 XML 깨짐). 180 `f"{task_name}_task.xml"` 미정규화(`:` → 대체 스트림, `<>|?"*` → [Errno 22], JSON 만 남음)
- 고칠 것: 네 값에 `xml.sax.saxutils.escape`(요소 본문이라 `& < >` 만). 파일 이름은 JSON 과 같은 규칙: schedule_config 에 `safe_name(task_name)`(현 `_UNSAFE_NAME_CHARS` → `_`) 공개, config_path·write_task_xml 이 같이 씀. schtasks `/tn` 은 원래 이름 그대로
- golden: schedule_cases_golden.json `create_name_colon`(files·task_xml·schtasks_calls·detailed_log), `create_name_invalid_chars`(XML 쓰기 통과 → 응답 성공, return·dialogs·queue·files). 새 경우 `create_name_ampersand`(작업 이름 `A&B` → `<Arguments>` 에 `&amp;`). argv_golden.json 변화 없음
- 단위: test_automation_scheduler.py `test_build_task_xml_does_not_escape_values` → 이스케이프 기대로, write_task_xml 정규화 1건
- 파일: automation/scheduler.py, schedule_config.py, tests/test_automation_scheduler.py, tests/characterization/test_schedule_cases.py (+golden 1)
- C3: 등록되는 인자 문자열은 같다(XML 해석 뒤 원문). 바뀌는 건 지금 실패하던 이름의 XML 파일 이름뿐 → ADR 0003 §6 "파일 이름 정규화 규칙"을 XML 에도 적용한다고 한 줄 보탬. scheduler.py 머리 docstring 갱신

## 결함 5 실패 시 남는 파일 (fix-g 스케줄 · fix-e 복원 · fix-h 삭제 XML 은 F5)
- 스케줄 현재: scheduler.py:241-266 이 JSON 저장 → XML → schtasks. 266 실패·310 타임아웃 때 JSON·XML 이 남음. 같은 이름을 고치다 실패하면 이전 JSON 은 이미 덮어써져, 기존 작업이 새 설정으로 돈다
- fix-g: 저장 전 `schedule_config.read_raw(name)`(bytes|None) 로 기존 내용을 잡고, 저장 뒤 단계(XML 쓰기·schtasks 실패)에서 `schedule_config.restore_raw(name, prev)`(prev 가 있으면 되쓰기, 없으면 삭제) + 이번에 쓴 XML 삭제. 되돌리기 실패는 `_log.warning` 만, 원래 RuntimeError 를 그대로 올림. 타임아웃은 F3
- 버린 대안: schtasks 성공 뒤에 JSON 저장 — 상세 로그 순서가 바뀌고, 저장 실패 시 config 없는 작업이 등록된 채 남는다
- golden: `create_rc1_access_denied`·`create_rc2_stdout_only`(files → []), F3=B 면 `create_timeout` 도. 새 경우 `create_overwrite_rc1_keeps_previous`(기존 JSON 내용 유지). 파일: scheduler.py, schedule_config.py, test_automation_scheduler.py, test_automation_schedule_config.py, test_schedule_cases.py (+golden 1) · C3 없음
- 복원 현재: ezBAK.py:2868 에서 로그를 연 뒤 2880(교차 '아니오')·2883(공간 부족)에서 return → `*_restore_<ts>.log` 가 남고 오래된 로그 정리(성공 때만)도 안 돔
- fix-e: 복사 전 중단(result.cancelled·result.error)이면 finally 의 close 뒤 그 경로를 지움(close 가 log_file_path 를 None 으로 만드므로 열 때 경로를 지역 변수로 보관). 예외로 끝난 경우는 진단용으로 남김. 범위는 F4
- golden: restore_cases_golden.json `cross_direct_no`·`cross_parent_no`·`low_space`(logs → {}), F2=B 면 `disk_usage_fail` 도. 파일: ezBAK.py, test_restore_cases.py(docstring) (+golden 1) · C3 없음(이름 규칙 그대로, 만들지 않을 뿐)
- fix-h: scheduler.delete_task 가 schtasks 성공 뒤 `<safe_name>_task.xml` 을 지움(없으면 무시, 실패는 `_log`). ezBAK.py 수정 없음. golden `create_then_delete`(files → []). 파일: scheduler.py, test_automation_scheduler.py (+golden 1)

## 발주 순서 (같은 파일은 순차, 커밋은 끝난 순서대로 메인이 확인)
- 1차 병렬 3: fix-a(core/restore) ∥ fix-b(winsys/sysinfo) ∥ fix-f(automation)
- 2차 병렬 2: fix-c(ezBAK.py·restore golden) ∥ fix-g(automation, fix-f 뒤)
- 3차 병렬 2: fix-d(ezBAK.py·targets·restore golden, fix-c 뒤) ∥ fix-h(scheduler, fix-g 뒤)
- 4차: fix-e(ezBAK.py·restore_cases golden, fix-d 뒤). 끝나면 roadmap·모듈 CLAUDE.md(targets 복원 대상, automation·gui 현행 보존 줄, core restore.py 줄) 갱신은 메인
## 결정 (사용자)
- F2 sysinfo.free_space 반환값 검사(fix-b): A(권장) 고친다 / B 둔다 — 안 고치면 디스크 조회 실패가 "Available: 0 B" 공간 부족 오경보로 바뀔 뿐이다
- F3 schtasks 타임아웃 때: A(권장) 되돌리지 않음(등록 여부 불명, `_log` 만) / B 실패와 같이 되돌림 — 등록됐는데 JSON 을 지우면 예약 실행이 config 없음으로 실패한다
- F4 복원 로그 삭제 범위: A(권장) 교차 '아니오'·공간 부족 둘 다 / B 교차 '아니오'만 — 공간 부족 수치는 오류 대화상자에 이미 나온다
- F5 삭제 때 XML 도 지움(fix-h): A(권장) 이번에 포함 / B 제외 — 생성 실패 정리와 같은 파일을 다루고 ezBAK.py 를 건드리지 않는다
