# ezbak/targets/
역할: 백업 대상 종류(V1)가 늘거나 한 종류의 위치·방식이 바뀔 때 수정된다(브라우저 추가, 앱 설정 등)
공개 진입점(하위 모듈 apps·drivers 는 `__all__` 에 있어 `from ezbak.targets import apps, drivers` 허용, ADR 0002):
- default_targets() — 정기 백업 작업에 붙는 대상 목록(GUI·Headless 공통, 지금은 [Bookmarks()]), Bookmarks(BackupTarget: backup(ctx) -> TargetResult), BROWSERS·Browser(name, relative, backup_name, source(profile)), restore_bookmarks(backup_root, out, profile, browsers=BROWSERS)
- 프로필 전체: backup_profiles(dest_root, out, progress, *, local, roaming, free_space, clock) -> dest_base | None(원본 0개, 폴더 안 만듦). 공간은 폴더를 만들기 전에 free_space(dest_root) 로 확인, 부족하면 InsufficientSpaceError. PROFILE_SOURCES·ProfileSource(name, root, relative, dest, .path(local, roaming))
- drivers: PnputilCmd(display, argv)·build_export_cmd(path)·build_add_cmd(folder) -> PnputilCmd·backup_folder_name(system, now)·parse_backup_folder(name) -> 시스템명|None(표지 대소문자 무시)·restore_log_name(system, now)·BACKUP_MARKER·run_pnputil(cmd, out) -> rc
- apps: build_version_cmd()·build_export_cmd(output_path, basic=False)·build_list_cmd()·build_simple_list_cmd()·export_file_name(now)·export(output_path, out, *, clock) -> AppsExport(kind, path, rc)
의존: 들어오는 것 gui/, automation/ · 나가는 것 core 공개 이름(ports·model·joblog·fileops), winsys 공개 이름(paths, proc, fileattrs) · 금지: tkinter, subprocess 직접, storage/·automation/ import
| 파일 | 역할 |
|---|---|
| __init__.py | default_targets(), 공개 이름 |
| browser.py | BROWSERS 표(Chrome, Edge), 북마크 백업(Bookmarks)·복원(restore_bookmarks), PROFILE_SOURCES 표(Chrome, Edge, Firefox+profiles.ini)와 프로필 전체 백업(정션 건너뜀은 fileattrs REPARSE_POINT) |
| drivers.py | pnputil PnputilCmd(display=로그용 현행 문자열, argv)·폴더/로그 이름 규칙·run_pnputil(Executing 줄 → proc.stream → STDERR 줄) |
| apps.py | winget 명령 4개와 fallback 사슬(--version → export → basic → list --accept → list → RuntimeError) |
테스트: python -m pytest tests -q -k targets
규칙: 명령줄은 build_*_cmd() 순수 함수가 만들고 winsys.proc 로만 실행한다. 새 브라우저 북마크는 BROWSERS 표 한 줄, 프로필 원본은 PROFILE_SOURCES 한 줄(P4 결정 D5 — 두 표를 섞지 않는다)
실행 kwargs: apps 는 현행 글자 그대로(문자열+shell=True, cp949, timeout). drivers 는 argv·shell=False·encoding 'oem'·errors 'replace'(P4-fix F2), 작업 로그 "Executing command:" 줄만 display 로 현행 글자
경계(P4 결정 D1): 이 폴더는 명령 조립·실행·출력 해석·작업 로그 줄(out 콜백)까지. 로그 열기·큐·대화상자·play_sound·finally 는 gui/actions_tools.py wrapper. 시각은 clock/now 인자로만
북마크 파일 이름(Chrome_Bookmarks·Edge_Bookmarks)과 dest 루트 위치(ctx.dest)는 C3 고정. 북마크 백업 원본·복원 대상은 대상 사용자 프로필(browser.source(profile), P2-fix #1·fix-d). 실행 계정 %LOCALAPPDATA% 로 되돌아가지 않는다
프로필 전체 백업의 원본은 GUI 에서 선택한 사용자의 <USERS_ROOT>\<user>\AppData\{Local,Roaming}(gui 가 호출 시점에 넘김), 시작 때 Tk 스레드 확인 대화상자(P4-fix F4=C, g-p)
Bookmarks: 브라우저별 try 여부는 policy.isolate_bookmark_errors 로만 고른다. 문구는 JobLogFormat bm_* 항목. bm_header·복원 시작 문구는 C3 고정이라 브라우저를 늘릴 때 문구를 따로 정한다
드라이버·앱 목록·프로필 전체 백업의 BackupTarget 적합은 정기 백업에 묶이는 요구가 확정될 때(ADR 0002)
ADR: docs/adr/0002-패키지-구조.md
