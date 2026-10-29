# ezbak/winsys/
역할: Windows I/O 단일 출구. 외부 명령 실행 방식, Windows API 호출, 선택 의존(pywin32·psutil) fallback 이 바뀔 때 수정된다
공개 진입점: proc.run(argv, …)/proc.stream(argv, …)/proc.TimeoutExpired/proc.list2cmdline/proc.PIPE(재노출), fileattrs.get_attributes(path)/HAS_PYWIN32/FILE_ATTRIBUTE_*, paths.USERS_ROOT/users_root()/appdata_dir()/resource_path(rel)/list_users(root=None), free_space(path)(sysinfo), sysinfo.system_name(*, debug=진단 로그 DEBUG — P6-5d)/used_drive_letters(get_logical_drives=None)/clean_name(name), elevation.is_admin()/relaunch_as_admin(executable, argv), shell.open_device_manager() -> 로그 문구
의존: 들어오는 것 모든 ezbak 모듈 · 나가는 것 표준 라이브러리, pywin32·psutil(선택) · 금지: 다른 ezbak 모듈 import, tkinter
| 파일 | 역할 |
|---|---|
| proc.py | subprocess 를 import 하는 유일한 파일(루트 ezBAK.py 포함, 검사 c). subprocess.run/Popen 에 그대로 전달하고 기본 인자를 끼워 넣지 않는다. 진단 로그에는 프로그램 이름만(argv 에 NAS 암호가 있다). 다른 모듈은 TimeoutExpired·list2cmdline·PIPE 도 proc 에서 가져온다 |
| fileattrs.py | GetFileAttributes 원시 조회(pywin32 → ctypes fallback), FILE_ATTRIBUTE_* 상수. HAS_PYWIN32 는 호출마다 읽는다(테스트 monkeypatch). 숨김 판정은 core.filters 몫 |
| paths.py | USERS_ROOT(C:\Users 의 유일한 정의, P5 결정 D3 — 루트·gui 는 호출 시점에 winsys_paths.USERS_ROOT 로 읽고, 테스트는 facade 별칭으로 바꿔 넣는다), %APPDATA%, _MEIPASS 기준 경로. 환경 변수는 호출 시점에 읽는다 |
| elevation.py | is_admin(IsUserAnAdmin, 예외면 False), relaunch_as_admin(ShellExecuteW runas, params=list2cmdline(argv), 반환값 그대로). 반환값 해석·실패 문구는 호출처 |
| sysinfo.py | free_space(disk_usage → GetDiskFreeSpaceExW(_api_free_space: 예외·반환값 0 이면 실패) → 드라이브 루트 → 999TB), system_name(BIOS 레지스트리 키 BaseBoardProduct → SystemProductName → SystemManufacturer(P4-fix F3, 새 값 없음은 DEBUG 없음) → wmic 3개 표(마지막 대체) → "Unknown", DEBUG 는 debug 콜백, 기본은 진단 로그), used_drive_letters(GetLogicalDrives 비트마스크, ctypes 는 호출 시점 조회, 예외면 A~Z exists 탐색) |
| shell.py | open_device_manager(os.startfile → mmc proc.stream → `start` shell, 전부 실패면 FileNotFoundError, startfile 오류는 그대로 올림). 탐색기 열기는 gui/actions_tools |
테스트: python -m pytest tests -q -k winsys
규칙: 업무 판단(필터·백업 이름 규칙)을 넣지 않는다. 예외(P4 결정 D2): 시스템 이름 정리 clean_name 은 드라이버 폴더 이름(C3)에 쓰이지만 조회와 한 덩어리라 sysinfo 에 둔다
선택 의존이 없을 때의 동작은 현행 fallback 그대로. resource_path 기준(_MEIPASS 또는 현재 작업 디렉터리)은 바꾸지 않는다
locale 인코딩·winreg·ctypes.windll(kernel32·shell32)·os.startfile·WINDIR 등 테스트가 바꿔치기하는 대상은 import 시점이 아니라 호출 시점에 조회한다(runpy 하네스·fake_shell32)
appdata_dir() 는 원값(없으면 '')만 준다. 없을 때의 처리(RuntimeError·홈 폴더·'%APPDATA%' 문자열 그대로)는 호출처마다 현행을 유지한다
free_space 는 API 가 실패(반환값 0)하면 드라이브 루트 단계로 넘어간다(fix-b). 성공한 호출의 0 은 그대로 0. core 는 sysinfo 를 직접 쓰지 않고 storage 의 BackupStore 를 거친다(ADR 0002 D4)
ADR: docs/adr/0002-패키지-구조.md
