# P2 엔진 통합 발주 계획 (초안, 2026-10-12, 기준 cc00c33)
따를 것: ADR 0002·0003, core·storage·targets·tests CLAUDE.md. 완료 기준(단계): backup·restore golden 각각 그대로 + 아키텍처 검사 a~c.

## 1. 현재 두 경로 대응 (ezBAK.py 줄)
| 단계 | GUI (App) | Headless `core_run_backup` 7259–7534 |
|---|---|---|
| 로그 열기·머리말 | `start_backup_thread` 2823 → `open_log_file` 2237(모드 a, Policy·Filters·구분선. 복원·winget·프로필도 씀) | 사용자 확인 뒤 7321 `open(w)`, "Backup start for '…' →" + 같은 머리말 |
| 폴더 이름·준비 | `run_backup` 2847: 사용자 확인(show_error) → "Backup start:" → 같은 이름 rmtree(로그 남김) → "Created new backup folder" | 7269–7290: 이름 → 같은 이름 rmtree(인자 틀린 핸들러, 조용히) → mkdir → 사용자 없으면 FileNotFoundError(로그 없음) |
| 크기·여유 공간 | `compute_total_bytes_safe` 3058(큐 상태) → `get_free_space_reliable` 3562 | `calculate_total_size`(필터 포함 skip), 실패 시 "Note: …" |
| 트리 복사 | 2897–2975 "Processing directory" 줄, 필터 제외도 "Skipping hidden", 실패 파일도 셈, 바이트=progress 합 | 7359–7428 hidden/filtered 따로, 0바이트·실패 제외, 복사 오류는 stdout print, CRITICAL 은 raise |
| 북마크 | `backup_browser_bookmarks` 3128(브라우저별 try) + 바깥 "…failed:" | 7430–7448 하나의 try, 줄바꿈 누락 3줄 |
| 오래된 백업 | `_cleanup_old_backups` 3084(rmtree onerror 3114, 폴더별 try 없음, "[ts] \n[Old…]") | 7450–7476 "\n[ts] [Old…]\n", 폴더별 try·"Could not delete" |
| 오래된 로그 | `_cleanup_old_logs` 2317(복원도 씀. 비활성 문구가 머리줄 앞, "Total files in folder", 삭제 2회차) | 7478–7532 1회, Total files 줄 없음 |
| 완료 | "Backup successfully completed." + 큐(log·update_status·open_folder) + finally(close·enable_buttons·after·sound) | `return (backup_path, log_path)`. stdout 줄은 `__main__` 7596·7637(P3) |

## 2. 기존 플래그 7개 밖의 차이 → 담을 곳(결정 D1 권장안)
- BackupPolicy 행동 플래그 4개 추가: `count_attempted_copies`(셈·바이트), `isolate_bookmark_errors`, `log_cleanup_retry_pass`, `skip_errors_include`(판정 예외 시 포함).
- JobLogFormat(문구): Policy·Filters·구분선(공용), 시작 줄 2종, Processing directory(H None), Skipping hidden/filtered, Created folder, Copying file, 복사 오류·TIMEOUT(GUI 로그 / H stdout·None), Walk·Size 오류(H None), Calculating total size, Insufficient space(GUI), Backup completed, 북마크 6줄+오류 2종, 정리 머리줄 2개·Found/Keeping/Deleting/Cleanup finished/No old/disabled, 로그 정리 13줄(Total files 는 H None), Cleanup failed, 완료 줄(H None). 항목마다 `lead`·`end`(H 줄바꿈 누락 재현). None=줄 안 씀.
- Reporter: log·status·progress + `progress_max`(D3). GUI 어댑터는 ezBAK.py 에 둔다(스로틀 `_last_progress_ts` 포함, P5 에 gui/bridge.py). Headless 어댑터: log=print(복사 오류 stdout), 나머지 무출력. show_error·open_folder·enable_buttons·play_sound·reset_progress 는 wrapper 가 결과·예외(UserNotFound·InsufficientSpace·CriticalBackupError)로 만든다.

## 3. golden 이 붙잡는 이름 (wrapper 가 유지)
- `App.start_backup_thread()`(모듈 `threading.Thread` 로 스레드 생성 — 테스트가 `mod.threading` 교체), `App.run_backup(backup_path)`, `App._cleanup_old_logs`(복원 golden 경유), `core_run_backup(user, dest, …)` → `(backup_path, log_path)`.
- monkeypatch: `mod.USERS_ROOT`, `mod.datetime`(fixed_clock), `mod.filedialog`·`Win11Dialog`·`messagebox`, env `LOCALAPPDATA`. 그래서 wrapper 는 `users_root=USERS_ROOT`, `clock=lambda: datetime.now()` 를 호출 시점에 ezBAK 전역에서 읽어 넘긴다. core 는 `datetime.now` 를 직접 부르지 않는다. 로그 정리 기준 시각은 지금처럼 `time.time()`(표시만 clock).
- 현 golden 은 기본 설정 한 경우뿐이라 정리 삭제·북마크 둘 다·오류 경로·큐 순서를 못 잡는다 → P2-0.

## 4. 발주 순서 (발주 1건 = 커밋 1개)
| # | Goal(완료 기준 1개) | Files | 추가 테스트 | golden 위험 → 막는 법 | ADR 0002 |
|---|---|---|---|---|---|
| P2-0 | 손대지 않은 ezBAK.py 로 경우별 golden 추가·통과 | tests/characterization/test_backup_cases.py, tests/fixtures/backup_cases_golden.json, tests/fixtures/user_tree.py | 두 경로×: 오래된 백업 3개(읽기 전용 포함), 오래된 .log(utime), 북마크 둘 다, 같은 이름 폴더(읽기 전용, fixed_clock 이라 이름 고정), 사용자 없음, GUI 여유 공간 부족, GUI 큐 순서(update_progress 연속 접기) | 없음(코드 무변경) | 해당 없음 |
| P2-1 | joblog·플래그 도입 후 두 머리말이 joblog 를 거쳐도 golden 그대로 | core/joblog.py, core/model.py, core/__init__.py, ezBAK.py, tests/test_core_joblog.py | JobLog 쓰기(lead·end), header_lines, format_bytes, 새 플래그 값 | GUI Policy 값은 var 문자열·실패 시 '2'/'30', 복원은 Filters 줄 없음 → 호출처가 문자열로 넘김 | core 안, winsys 미사용 |
| P2-2 | `winsys.sysinfo.free_space` 로 `get_free_space_reliable` 위임, 결과 동일 | winsys/sysinfo.py, winsys/__init__.py, ezBAK.py, tests/test_winsys_sysinfo.py | disk_usage→ctypes→드라이브 루트→999TB 순서 | 없음(GUI 3곳 같은 값) | winsys 단독 |
| P2-3 | ports Protocol + `LocalFolderStore` 단위 테스트 통과 | core/ports.py, core/__init__.py, storage/__init__.py, storage/local.py, tests/test_storage_local.py | 이름(`<user>_backup_<ts>`), 목록(소문자 접두·isdir·정렬), 삭제(rmtree_readonly), free_space | 배선 없음 | storage→core·winsys(sysinfo) |
| P2-4 | `core.retention` 으로 두 정리 이동 + `rmtree_readonly` 교체, golden(기본·경우·복원) 그대로 | core/retention.py, core/joblog.py, core/__init__.py, ezBAK.py, tests/test_core_retention.py | 두 문구셋·비활성·삭제 실패·2회차 | 머리줄 위치·Total files·폴더별 try → 포맷 None·플래그. App._handle_rmtree_error 삭제 | core→store 는 ports 로만 |
| P2-5 | `targets.browser.Bookmarks` 로 북마크 이동, golden 그대로 | targets/__init__.py(default_targets), targets/browser.py, core/joblog.py, ezBAK.py, tests/test_targets_browser.py | BROWSERS 표, 없음·있음·복사 오류 | H 줄바꿈 누락·dest 루트 위치 유지(end="") | targets→core·winsys.paths |
| P2-6 | `core.engine.run_backup` + `core_run_backup` thin wrapper, Headless golden 그대로 | core/engine.py, core/joblog.py, core/__init__.py, ezBAK.py, tests/test_core_engine.py | 플래그별 분기, CRITICAL raise, 셈 규칙 | 폴더 준비·사용자 확인·로그 열기 순서는 wrapper(D2). 내부 `import stat` 소멸 | core→winsys.fileattrs·paths 만 |
| P2-7 | `App.run_backup` thin wrapper(GUI Reporter 어댑터), GUI golden·큐 golden·스모크 그대로 | ezBAK.py, core/engine.py, core/joblog.py, tests/test_core_engine.py | GUI 플래그 경로, 큐 문자열 13종 | finally·after·translator 상태는 wrapper 에 그대로. 크기 계산 상태 메시지 순서 | ezBAK.py 는 검사 제외(P5) |
| P2-8 | 미사용 정리 후 py_compile·전체 pytest 통과 | ezBAK.py | 없음 | 맨 위 `import stat`, FILE_ATTRIBUTE_READONLY/DIRECTORY/ARCHIVE/DEVICE(1015–1018) 삭제. REPARSE_POINT·`_get_file_attributes` 는 5113–5116 이 씀 — 남김 | 해당 없음 |
- P2 끝: 메인이 core/storage/targets/winsys CLAUDE.md 진입점 표 갱신(run_backup 시그니처, sysinfo 를 P2 로), roadmap 상태.

## 5. P2-fix 위치 (P2 뒤 별도 단계, 계획에 넣지 않음)
1. 북마크 실행 계정 복사: 지금 3131–3132·7433–7434(`%LOCALAPPDATA%`) → P2 뒤 `targets/browser.py` `Bookmarks.backup` 한 곳. 복원 3465–3466 도 같은 문제(P3 범위).
2. Headless 삭제 핸들러 인자 불일치: 7278 `onerror=_handle_rmtree_error_headless`(1167, 5인자) → P2 뒤에도 `core_run_backup` wrapper 의 폴더 준비 줄.
3. Headless 줄바꿈 누락: 7444·7446·7448·7474·7476 → P2 뒤 HEADLESS JobLogFormat 의 `end=""` 5항목. 기본 golden 에선 다음 절의 앞 "\n" 이 가려 준다 — 고칠 때 앞 "\n" 도 같이 정해야 빈 줄이 안 생긴다.

## 6. 결정 (2026-10-12 사용자 결정: D1~D4 모두 권장안. ADR 0002 에 기록)
- D1 차이 표현: A) Policy 플래그 4개 추가 + JobLogFormat 은 문구·None·lead/end(권장, ADR "엔진 하나" 그대로) / B) engine 안에 GUI·Headless 절차 함수 2개가 단계 함수 공유(플래그 0, 대신 ADR 0002 문구 수정·통일 때 두 절차 병합 비용).
- D2 로그·폴더 준비 소유: A) wrapper 가 열고 준비, engine 은 열린 JobLog·경로를 받음(권장. open_log_file 은 복원 등 공용, 순서 차이·P2-fix #2 줄이 wrapper 에 남음. ADR 0002 run_backup 시그니처에 log·fmt·backup_path 추가) / B) engine 이 전부(순서 플래그 2개 + GUI 로그 열기가 스레드 안으로, 큐 순서 위험).
- D3 Reporter 에 `progress_max` 추가(권장, ADR 0002 표 한 줄) / `progress(done,total)`(큐에 update_progress 0 이 하나 늘어 큐 golden 변화).
- D4 `winsys.sysinfo.free_space` 를 P4 에서 P2 로 당김(권장, core→sysinfo 금지라 store 가 필요) / GUI wrapper 가 계산해 넘김(V2 경계 약화).
