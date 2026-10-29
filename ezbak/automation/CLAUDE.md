# ezbak/automation/
역할: 자동화·CLI(V3)가 바뀔 때 수정된다 — CLI 인자, Headless 실행, 스케줄 config JSON, 예약 작업 등록
공개 진입점(하위 모듈 cli·headless·scheduler·schedule_config 는 `__all__` 에 있어 `from ezbak.automation import <하위 모듈>` 허용, ADR 0002):
- cli.build_parser()/parse(argv=None) -> Namespace(parse_known_args, 모르는 인자 버림)
- headless.run(args, *, users_root, clock) -> 0/1(config)·0(직접 인자)·None(그 밖, 호출처가 GUI 로), headless.run_backup(user_name, dest_dir, …, *, users_root, clock) -> (backup_path, log_path), HeadlessReporter
- scheduler.create_task(spec, entry_script, *, out, now, debug=진단 로그 DEBUG — P6-5d) -> 성공 문구/next_run(config, now)·next_scheduled(configs, now) -> (task_name, datetime)|None(U7-1)/delete_task(name)/build_create_cmd/build_delete_cmd/build_task_xml(…, *, now)/task_xml_path(name)/explain_error
- schedule_config.config_dir()/config_path(name)/safe_name(name)/save(spec, *, now)/load(path)/load_all()(U7-1)/delete(name)/read_raw(name)/restore_raw(name, prev)/job_from_config(config, *, log_folder=None) -> BackupJob, ScheduleSpec
의존: 들어오는 것 ezBAK.py, gui/ · 나가는 것 targets, storage, core, winsys · 금지: tkinter, subprocess 직접
| 파일 | 역할 |
|---|---|
| cli.py | argparse 정의 9개 인자(글자 그대로, P3 결정 D4). prog 는 기본값(호출 시점 sys.argv[0]) — usage·종료 코드 2 가 C3 |
| headless.py | run: config/직접 인자 분기, stdout `[ts]`(clock 한 번), config 오류만 traceback. run_backup: 폴더 준비 → 사용자 확인 → 로그 열기 → 엔진. 엔진은 모듈 전역 engine_run_backup(runpy 하네스 stub 대상), 리포터 클래스 이름 _HeadlessReporter 는 CLI golden 이 기록 |
| schedule_config.py | %APPDATA% 아래 스케줄 JSON 읽기·쓰기·삭제(유일 소유자). safe_name 이 JSON·작업 XML 파일 이름 규칙(ADR 0003 §6). ScheduleSpec(쓰기용, 12키 순서 고정, D5). load 는 open+json.load 만. read_raw/restore_raw 는 생성 실패 되돌리기용 바이트 그대로. job_from_config 는 .get 기본값 현행·형 변환 없음, log_folder 는 CLI `--log-folder` |
| scheduler.py | 순수: normalize_time·schedule_type·task_arguments·run_as_user·build_task_xml(템플릿 상수, 값 4개 XML 이스케이프)·build_create_cmd·build_delete_cmd·explain_error. 실행: write_task_xml(task_xml_path), create_task(검사 → config 저장 → XML → schtasks → 오류 해석, 실패 시 config·XML 되돌림), delete_task(성공 뒤 XML 삭제). schtasks 는 winsys.proc.run, kwargs 현행 그대로(삭제는 shell=True·인코딩 없음) |
테스트: python -m pytest tests -q -k automation
규칙: CLI 인자 이름·기본값·종료 코드, JSON 키, 예약 작업 인자 형식은 C3 고정(ADR 0003 §6). 모르는 인자는 무시한다
되돌리기(P3-fix): XML 쓰기·schtasks 실패면 이전 JSON 바이트를 되쓰거나(없었으면 삭제) 이번 XML 을 지운다. 타임아웃은 등록 여부를 몰라 되돌리지 않는다(결정 F3). 되돌리기·XML 삭제 실패는 _log 경고만, 원래 오류·문구 그대로
현행 보존(결함 후보, roadmap): config 값 무검증(잘못된 값도 엔진까지), 직접 인자는 오류여도 종료 0
필터 type 은 받은 값 그대로 저장한다(대소문자 정규화 없음, P3 결정 D7). 기본 필터 선택은 호출처(gui) 몫, spec.filters 는 그대로(None 이면 null)
users_root·시각(clock·now)은 인자로 받는다. ezBAK.py·gui 가 호출 시점의 winsys.paths.USERS_ROOT·datetime.now 를 넘긴다. datetime·환경 변수·locale 인코딩을 import 시점에 묶지 않는다(fixed_clock·runpy 하네스)
큐 메시지·대화상자는 호출처(gui/actions_schedule.py) 몫. 상세 로그는 out, DEBUG print 는 debug 콜백으로 순서 그대로 낸다
비frozen 등록 인자의 스크립트 경로는 entry_script(루트 ezBAK.py 가 gui.run 으로 넘기고 App.entry_script 로 전달)만 쓴다. 이 폴더의 __file__ 을 쓰지 않는다
새 CLI 인자 = cli.py + BackupJob 필드 + schedule_config 매핑(ScheduleSpec·job_from_config) 세 곳
테스트는 실제 %APPDATA%·~\AppData\Local\ezBAK\schedules 를 건드리지 않는다(삭제 테스트도 임시 폴더 환경으로)
ADR: docs/adr/0002-패키지-구조.md, docs/adr/0003-구현-규칙.md
