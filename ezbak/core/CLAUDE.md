# ezbak/core/
역할: 백업·복원 절차, 필터·숨김 판정, 보관 정리, 작업 로그 문구가 바뀔 때 수정된다. GUI·Headless·스케줄이 같이 쓰는 엔진
공개 진입점(`__init__.py` 의 이름만 쓴다):
- 모델: BackupJob·BackupPolicy·BackupResult·FilterRule·FilterSet·FILTER_TYPES·DEFAULT_FILTERS
- 포트: Reporter·BackupStore·BackupTarget(runtime_checkable Protocol), TargetContext·TargetResult
- 판정: is_skipped(path, job, policy, is_dir=None, check_filters=True) -> SKIP_HIDDEN|SKIP_FILTERED|None, matches_filter(path, filters, is_dir=False)
- 파일: copy_file_with_progress, compute_total_size, copy_sources, handle_readonly_error, rmtree_readonly, COPY_TIMEOUT_SECONDS
- 작업 로그: JobLog(stream, clock), JobLogFormat(GUI·HEADLESS), LogItem, JobLine(lead·end), header_lines, format_bytes
- 보관 정리: cleanup_old_backups(store, user_name, keep, fmt, policy, out), cleanup_old_logs(folder, days, fmt, policy, out, *, clock, now_ts). out 은 JobLine 을 받는 콜백
- 엔진: run_backup(job, backup_path, log, fmt, store, targets, reporter, policy, clock, *, users_root, log_path="", now_ts=time.time) -> BackupResult, InsufficientSpaceError(required, free)
- 진단 로그(하위 모듈 diag, P6): setup(*, log_dir, level, environ) → %APPDATA%\ezBAK\logs\ezbak.log(1MB×3, `ezbak` 로거만, 한 번만, 실패 시 None)·teardown()·new_run_id(clock)·bind(run_id)·current()·level_from_env(EZBAK_LOG_LEVEL, 기본 INFO)
- 마지막 백업(하위 모듈 history, U7-2): record_backup(*, user, dest, folder, when, mode) -> bool(%APPDATA%\ezBAK\state.json, 실패는 경고만)·last_backup() -> dict|None·state_path()
- 복원: validate_selection(path)(RestoreSelectionError(.message = 현행 show_error 원문)), locate_backup(path, user, choose_backup) -> RestorePlan(kind KIND_DIRECT|KIND_PARENT, backup_root, folder_name, backup_user, source_dir)|None(취소), run_restore(plan, user, destination_dir, out, reporter, *, skip, free_space, confirm_cross_user) -> RestoreResult(cancelled, error, files, bytes, total)
의존: 들어오는 것 gui/, automation/, targets/, storage/ · 나가는 것 winsys.fileattrs, winsys.paths · 금지: tkinter, subprocess, winsys.proc, winsys 외 ezbak 모듈 import
| 파일 | 역할 |
|---|---|
| model.py | BackupJob(filters 는 JSON dict, Optional[dict]), FilterRule(type 소문자 정규화, 생성자는 잘못된 type 에 ValueError, from_dict 는 버림), FilterSet, BackupPolicy(GUI/HEADLESS: 판정 플래그 7개 + 행동 플래그 count_attempted_copies·isolate_bookmark_errors·log_cleanup_retry_pass·skip_errors_include·isolate_backup_delete_errors·isolate_retention_errors·return_critical_errors), BackupResult(critical_error 포함), DEFAULT_FILTERS(dict, C3) |
| ports.py | Reporter(log·status·progress·progress_max), BackupStore(new_backup_path·new_log_path·list_backups·path_of·delete_backup·free_space), BackupTarget(name, backup(ctx)). TargetContext(backup_path, log, dest, fmt, policy, out, profile — 모두 기본값, profile 은 대상 사용자 프로필 폴더)·TargetResult(name, ok, errors). winsys 도 import 하지 않는다 |
| filters.py | 필터 매칭·숨김/시스템/재분석 지점 판정(유일 구현). 건너뛰는 이유를 돌려줘 호출처가 로그 문구를 고른다 |
| fileops.py | 진행률 복사(30초 타임아웃, clock 주입), 전체 크기 계산, 읽기 전용 rmtree, 여러 원본 복사. 건너뛰기 판정은 skip(path, is_dir) 인자, 로그·진행률은 콜백 |
| engine.py | 백업 본체: 크기 계산 → 여유 공간(check_free_space, 크기 계산 실패 뒤 부분 합계로도) → 프로필 트리 복사 → targets(ctx.profile = users_root/job.user) → 보관 정리. CRITICAL 은 로그 후 raise, return_critical_errors 면 result.critical_error 로 반환. isolate_retention_errors 면 정리 실패·보관 값 읽기 실패를 로그로 남기고 그 정리만 건너뜀. 로그 열기·폴더 준비·사용자 확인은 진입 wrapper 몫(ADR 0002 D2). storage·targets 는 인자로만 |
| restore.py | 복원 3단계(P3 결정 D1): 선택 검증 → 백업 위치 찾기(후보는 대상 사용자 먼저, 묶음 안은 이름 역순=최신 먼저, 대소문자 무시 — fix-a) → (wrapper 가 작업 로그 열기) → run_restore(교차 확인·크기(실패 시 SIZE_FALLBACK 100MB)·공간·복사·PROGRESS_EVERY 50). 문구는 MSG_* 상수 한 벌(GUI 전용), 공간 부족은 result.error, free_space·skip 예외는 그대로 올림. 사용자 응답은 콜백(choose_backup·confirm_cross_user)으로만, 큐·폴링·time.sleep 은 wrapper(D2) |
| retention.py | 백업 개수·로그 일수 정리. 백업 목록·삭제는 BackupStore 로만, 로그 정리 기준 시각은 now_ts(현행 time.time), 표시 시각은 clock |
| joblog.py | 이미 열린 stream 에 `[ts] msg` 를 쓴다(열기 모드·인코딩은 호출처). JobLogFormat 항목: 머리말, 엔진(size_*·트리), 정리 bk_*·lg_*, 북마크 bm_*, GUI wrapper 줄, Reporter 메시지 rp_*. None 은 그 줄을 안 쓴다 |
| history.py | `state.json` 의 last_backup(시각·사용자·대상·폴더·mode). settings.json·스케줄 JSON 은 건드리지 않는다(C3). 임시 파일 + os.replace 로 쓴다. 호출처는 GUI·Headless 백업 wrapper(성공 뒤) |
| diag.py | 진단 logging 설정(ADR 0003 §2 형식), run_id(contextvars, 작업 스레드는 시작 때 다시 bind). 작업 로그(joblog)와 섞지 않는다 — run_id 는 작업 로그 본문에 넣지 않는다 |
테스트: python -m pytest tests -q -k "core or characterization"
규칙: 작업 로그 문구·순서는 ADR 0001 기준선 그대로. 시각은 clock 인자로만 얻는다. 필터 type 비교는 정규화된 소문자만
HEADLESS 항목은 모두 lead 없이 end="\n" 이다(P2-fix #3 의 Q1 예외 — test_headless_items_are_plain_lines). GUI 포맷의 lead·end 는 기준선 그대로
GUI·Headless 차이: 행동은 BackupPolicy 플래그, 문구·줄 유무·lead/end 는 JobLogFormat. 포맷 항목의 None 여부로 동작을 고르지 않는다. 통일은 별도 결정 전까지 하지 않는다
App.is_hidden 을 쓰던 곳(복원·크기 계산·복사 스레드)은 replace(BackupPolicy.GUI, skip_appdata_names=True, attr_ctypes_fallback=True) 를 쓴다. 세 번째 이름 상수는 만들지 않는다
filters 는 필터를 JSON dict 모양 그대로 받는다. FilterSet.from_dict 를 거치면 잘못된 include 규칙이 버려져 판정이 달라진다
ADR: docs/adr/0002-패키지-구조.md, docs/adr/0003-구현-규칙.md
