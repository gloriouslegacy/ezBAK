# tests/
역할: 회귀 검증(Q2). 동작 기준선(ADR 0001)이나 모듈 공개 진입점이 바뀔 때 수정된다
공개 진입점: python -m pytest tests -q (CI·로컬 공통)
의존: 나가는 것 ezbak 각 모듈의 `__init__.py` 공개 이름, 특성 테스트에 한해 루트 ezBAK.py(facade 경유) · 금지: 모듈 내부 파일 직접 import(아키텍처 테스트 제외), 실제 schtasks 등록·드라이버 설치·UAC·네트워크
| 파일·폴더 | 역할 |
|---|---|
| conftest.py | ezbak_mod·ezbak_mod_nopywin32(둘 다 facade 반환, nopywin32 는 새로 import 한 ezbak.* 를 담음), isolated_env(임시 %APPDATA%·%LOCALAPPDATA%), user_root(facade 별칭으로 winsys.paths.USERS_ROOT 를 바꿈), fixed_clock, block_subprocess, record_subprocess·fake_shell32, gui_app(세션 공용 App(entry_script=EZBAK_PATH)), run_ezbak_main(runpy, fixtures/cli_harness.py). 500줄이 가까우면 fixtures/ 로 뺀다 |
| characterization/ | 분리 전 golden: 필터·숨김 표, 백업(GUI·Headless·경우별), 외부 명령 argv, 복원(golden·경우별), Headless CLI, 예약 작업, P4 도구(tool_cases·_net·_profiles), P5 GUI(test_gui_snapshot: 창 트리 en/ko×dark/light·바인딩·메뉴·테마/언어 전환·번역·settings, test_gui_dialogs: Win11Dialog·필터·원본·스케줄·폴더 선택, test_gui_dialogs_nas: NAS — 같은 gui_dialogs golden JSON, test_gui_cases: 복사·공간 계산) |
| fixtures/ | ezbak_facade.py(읽기 루트 → ezbak.gui 공개 이름 → 별칭 표 (모듈, 속성), 속성 None 은 모듈 자체. setattr 은 루트(있으면)·별칭 대상·그 이름을 가진 ezbak.gui.* 전부), user_tree.py, cli_harness.py, tool_harness.py(ask_yes_no 는 bridge 모듈 time 을 바꿔 메인 스레드 차례에 답함, check_golden 은 자기 키만 바꿔 다시 씀), gui_harness.py(트리·메뉴 순회, recording·override·modal_hook·app_unchanged, own_binding_sequences 는 CTk 루트가 자기 메서드에 묶은 <Configure>·<FocusIn> 을 뺌) + gui_ctk.py(CTk 위젯은 cget 으로 CTK_OPTIONS, 그리기 부품 _canvas 등은 기록 안 함, U2-2), golden JSON |
| test_<모듈>_<주제>.py | 모듈 단위 테스트. test_gui_bridge 는 대역 App(Tk 없음)으로 GuiReporter·LogStream·FolderStore·GuiJob·int_or, test_ezbak_facade 는 합성 모듈로 facade 자체 |
| test_gui_harness_ctk.py | gui_harness 의 CTk 대응(옵션·부품 제외·버튼 찾기·own_members). CTk 위젯은 tk 부모의 configure 를 인스턴스에 덮어쓰므로 세션 App 에 직접 붙이지 않는다 |
| test_gui_components.py | U3-3 공용 위젯(Card·StatTile·Chip·SidebarItem·RingButton): 색·반경·선택·invoke·state·원형 진행 그리기·harness 기록. tk.Frame 을 사이에 두고 만든다 |
| test_gui_ctk_root.py | U2-3: ctk.CTk 루트가 App.__init__ 에 필요한 것(withdraw·bg·menu·아이콘 실패·ttk·__class__ 교체·DPI 끔)을 지키는지. 루트는 프로세스당 하나라 fixtures/ctk_root_probe.py 를 자식 Python 으로 돌린다 |
| test_gui_window_shown.py · test_gui_options_fit.py | v0.10.x 회귀(Windows 전용, 자식 Python 에서 mainloop — 아래 gui_app 규칙의 예외): 창이 normal·mapped·viewable 인지 / 1000×700 en·ko 에서 Backup Options 의 System Exclude 가 카드 안에 드는지(spec ui-fixes F1) |
| test_gui_smoke.py | App 생성(창 표시 없음)·위젯·큐 메시지 14종 처리·작업 스레드 확인(_wait_yes_no) |
| test_architecture.py | ast 검사 a(tkinter·customtkinter)·b(의존 방향·내부 파일 import, 하위 모듈은 패키지 `__all__` 에 있을 때만 — ast 로 읽음, P5-17 결정 A)·c(subprocess), 표준 라이브러리 이름 충돌. 루트 ezBAK.py 도 a·b·c(간선 automation·gui·winsys). d(P6-6): ezbak/ 의 print(호출·값, automation/headless.py 만 예외)·맨 except: |
테스트: python -m pytest tests -q · 마커 -m gui(Tk 필요), -m windows(속성·ctypes 필요)
golden 재생성: EZBAK_REGEN_GOLDEN=1 python -m pytest <파일> -q — Q1 예외로 승인된 커밋에서만, 사유와 바뀐 키를 커밋 메시지에. 두 번 생성해 같은지, 키 단위로 비교. 번역 키를 추가하면 gui_snapshot 의 translations 도 함께. 재생성 때만 가끔 OSError Errno 22(파일 잠금 추정) — 한 번 다시 돌린다
규칙: golden 은 분리 전 코드로 만든다. pywin32 여부로 결과가 다르면 두 경우를 각각 고정. 실제 디스크·로캘·시계·화면 크기에 따라 달라지면 안 된다(disk_usage·_api_free_space·fixed_clock 고정, 픽셀 geometry·경로 이름 기록 안 함)
테스트 전용 별칭과 루트에서 뺀 이름(_have_pywin32·USERS_ROOT·tk·os·threading·DEFAULT_FILTERS·core_engine_run_backup·get_system_info 등)은 ezBAK.py 가 아니라 facade 별칭 표에 둔다. 개별 테스트는 mod.<이름> 그대로 쓴다
_api_free_space 패치는 mod.winsys_free_space.__globals__ 로(nopywin32 사본은 sys.modules 에 없다)
Tk·App 은 gui_app 으로만 얻는다(Tk 재생성은 tk.tcl 로드가 간헐 실패). App 에 기록기를 붙일 때는 MonkeyPatch 대신 gui_harness.recording·override(MonkeyPatch undo 가 App 에 메서드를 남겨 다른 테스트를 깬 적 있음). 대화상자는 modal_hook 으로 몰고 끝에 app_unchanged 로 확인
새 허용 간선·모듈을 추가하면 test_architecture.py 의 허용 간선 표를 ADR 0002 와 같이 고친다
정션은 user_tree.make_junction(mklink /J)으로. 숨김·시스템 속성은 SetFileAttributesW 로 만든다
로그 비교는 타임스탬프 <TS>, 임시 경로 <TMP>, Windows 오류 문구 <MSG>, 트레이스백 프레임 <FRAMES>. GUI 큐는 progress_throttle_secs 를 크게 두고 연속 update_progress 를 하나로 접는다
복원 사용자 응답(_folder_selection_result·_cross_user_confirmed, 0.1초 폴링)은 실행 전에 App 속성으로 넣고 끝나면 지운다. time.sleep 은 BaseException 스텁이라 답이 없으면 바로 실패한다
run_ezbak_main: 엔진 run_backup 은 stub, Tk 생성·os._exit·ShellExecuteW 는 막힘. 사용자는 Public(없으면 skip). config 인자는 --schedule-config
외부 명령 테스트는 argv·shell·인코딩만 확인하고 실제 실행은 block_subprocess 로 막는다. 실제 %APPDATA%·스케줄 폴더·C:\Users 를 건드리는 함수는 임시 환경에서만 부른다
NAS 테스트 암호는 dummy-pw 만 쓰고, 로그·큐·print·대화상자·파일에 없음을 단언한다(golden 에는 argv 안에만)
ADR: docs/adr/0003-구현-규칙.md
