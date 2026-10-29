# U1 화면 설계서 (확정 2026-10-27 — D1~D12 모두 권장안, 기준 refactor/package-split 끝 · ADR 0004)
따를 것: ADR 0002·0003·0004, `ui-stages.md`(사용자 결정 5개), gui·tests CLAUDE.md. 시작 조건: U0 커밋 뒤, 아래 D1~D12 결정 뒤.
완료 기준(U 전체): 동작 golden 그대로 + 창 구조 golden 은 해당 커밋에서만 재생성(ADR 0004 §3) + 파일 500·함수 60(배치 150) + tkinter·customtkinter 는 `ezbak/gui/` 만.

## 0. 지금 코드에서 확인한 사실 (설계의 전제)
- 창: `app.py:_init_window` 가 `geometry("1026x790")`, minsize 없음, 기본 테마 `Win11Theme(initial_mode='dark')`, 언어 'en'. **DPI 인식 코드는 어디에도 없다**(ezBAK.py·gui·winsys grep 0건, PyInstaller 는 CLI 플래그만 쓰고 .spec 파일도 없음). 지금 exe 는 DPI 비인식 → Windows 가 비트맵 확대.
- 화면: `layout.py:create_widgets` → `_build_header_section`(메뉴바·시계·옵션 줄) · `_build_backup_section`(사용자 카드·Main Operations 카드·소리 토글·버튼 5개, 지역 함수 `create_button` 반환) · `_build_tools_section`(버튼 10개 2줄) · `_build_progress_log_section`(진행률·상태·Activity Log·F1 안내). layout.py 445줄이라 사이드바를 넣을 자리가 없다.
- 카드 제목·옵션 문구("Main Operations", "Select User:", "Backups to Keep:", "Hidden:" 등)는 하드코딩 영어다. 키 `main_operations`·`select_user`·`activity_log`·`tools_utilities` 는 사전에 있지만 안 쓰인다. 메뉴의 "Help"·"Keyboard Shortcuts (F1)"·"GitHub Releases" 도 하드코딩.
- 진행·로그는 위젯이 한 벌이다: bridge `process_queue` 가 `self.progress_bar['value'/'maximum'/'mode']`·`status_label.config` 를, `AppLogMixin.log` 가 `log_text` 를 직접 고친다. 같은 위젯을 두 페이지에 둘 수 없다(→ D9).
- 대상 폴더는 상태로 없다: `start_backup_thread`·`start_restore_thread` 가 매번 `filedialog.askdirectory` 로 묻는다. 원본 크기·여유 공간은 `check_space` 흐름의 결과(`show_space_check_result`)에만 있다. → 홈의 "사용자 → 대상 폴더"·여유 공간 지표는 지금 데이터로 못 채운다(U7).
- `get_system_info()`(actions_tools → `winsys.sysinfo.system_name`)는 레지스트리 실패 시 wmic 를 proc 로 돌리고 `DEBUG:` 를 print 한다. 창 생성 중에 부르면 `test_window_tree` 의 "DEBUG 없음" 단언과 block_subprocess 에 걸린다.
- 테스트 도구: `gui_harness.own_members` 는 모듈 이름이 `tkinter` 로 시작하는 클래스만 뺀다. App 기반을 `ctk.CTk` 로 바꾸면 CTk 메서드(configure·geometry·destroy…)가 `make_host`(tk.Toplevel 숙주)와 `recording`(메서드 바꿔치기)에 섞인다. `changed_options(widget.configure())` 는 CTk 위젯에서 dict 를 못 받는다. `_texts`(test_gui_snapshot)는 `"text" in value.keys()` 로 고르므로 CTkButton 글자가 language_switch 에서 빠진다. `gh.button`·smoke `test_select_folder` 는 `tk.Button` 으로 찾는다.
- 빌드: release.yml 은 `--add-data "./icon;./icon"`, 로컬 ps1 은 `./icon/ezbak_title.ico` 만 넣고 `bak_version_info.txt` 를 쓴다(둘이 이미 다름). customtkinter 의 테마 JSON·글꼴은 `--collect-data customtkinter` 가 필요(hooks-contrib 훅이 있어도 명시 권장).

## 1. 창 구성
- 왼쪽 사이드바(폭 220) + 오른쪽 내용 영역. 내용 영역 = 상단 바(공통) + 페이지 프레임 1개 + 하단 진행 줄(공통).
- 상단 바: 페이지 제목 · `user_combo`(사용자 선택, 모든 페이지 공통 — 백업·복원·브라우저·예약 대화상자 `_parent_user_name` 이 같은 `user_var` 를 읽는다) · `header_time_label`(시계).
- 하단 진행 줄: `progress_bar`(ttk.Progressbar 유지) + `status_label`. 어느 페이지에서 단축키로 작업을 시작해도 보인다.
- 사이드바: 로고·"ezBAK" · 상태 칩("대기"/"작업 중 42%", progress_bar·status_label 을 `after` 0.5초로 읽음, 큐·bridge 무변경) · 메뉴 8개 · 하단 "F1 단축키"(지금 footer `shortcut_hint` 자리). 광고 카드·버전 표시는 넣지 않는다(코드에 버전 상수 없음, version_info 는 빌드 때만).
- 페이지 전환: 페이지 프레임은 시작 때 모두 만들고 `pack_forget`/`pack` 으로만 바꾼다(위젯 속성·바인딩·`set_buttons_state` 대상이 그대로 산다). 위젯을 지우고 다시 만들지 않는다.
- 페이지 코드: `ezbak/gui/pages/<페이지>.py` 의 모듈 함수 `build(app, parent)`(D12). App 메서드가 아니라서 `gui_harness.recording` 이 바꿔치기하지 않는다(`app.py:_add_help_menu` 와 같은 이유).

## 2. 메뉴별 화면 (U4 는 절 단위로 통째 이동, U5 에서 아래처럼 나눔)
| 메뉴 | 내용 (지금 이름) |
|---|---|
| 홈 | 상태 카드: 원형 "백업 시작"(= `start_backup_thread`, 진행 중엔 progress_bar 값을 원형으로) · 제목 "선택 사용자"(`user_var`) · 칩 "필터 n개"(`len(filters['include'])+len(filters['exclude'])`, 누르면 `open_filter_manager`) · 지표 2개: 보관 개수(`retention_count_var`)·로그 보관 일수(`log_retention_days_var`). 시스템 카드: 시스템 이름(`get_system_info`, 아래 주의)·사용자 수(`user_combo['values']`)·숨김/시스템 포함 여부. 작업 로그 카드(D9). 원본 크기·대상 여유 공간·마지막 백업·다음 예약은 U7 |
| 백업 | `backup_btn`(Backup Data) · `filters_btn` · 포함 옵션 `hidden_mode_var`·`system_mode_var` 라디오(복사·공간 확인·복원에도 적용된다는 안내) · 보관: `retention_count_var`·`log_retention_days_var` 스핀박스 · 추가 백업: `driver_backup_btn`·`browser_profiles_btn`·`winget_export_btn` |
| 복원 | `restore_btn` · `driver_restore_btn` · 안내(덮어쓰기 경고는 지금처럼 `overwrite_warning` 대화상자) |
| 예약 백업 | `schedule_btn`(→ `schedule_backup` → ScheduleBackupDialog, 만들기·삭제 둘 다 대화상자 안) + 설명 문구 |
| 도구 | `check_space_btn` · `copy_btn` · `nas_connect_btn` · `nas_disconnect_btn` · `devmgr_btn` · `file_explorer_btn` |
| 로그 | `log_text`(+스크롤바, `bold` 태그 그대로) · `save_log_btn` |
| 설정 | 테마 라이트/다크(`switch_theme`) · 언어 English/한국어(`switch_language`) · 소리(`sound_label`·`sound_toggle` → `_on_sound_toggle`) |
| 도움말 | 단축키 보기(`show_keyboard_shortcuts_help`, F1 대화상자 그대로) · GitHub Releases(`open_github_link(None)`) · 단축키 표(`keyboard_shortcuts_content` 키 그대로 표시) |
- 홈 원형 버튼은 `home_backup_btn` 새 속성 → `set_buttons_state` 의 선택 버튼 목록(try/except 블록)에 추가. `backup_btn` 은 백업 페이지에 남긴다(smoke `ACTION_BUTTONS` 15개 유지).
- 시스템 이름: 홈 첫 표시 뒤 작업 스레드에서 `get_system_info()` → App 속성에 저장 → 메인 스레드 `after` 폴링으로 표시(bridge `_poll_answer` 방식, 새 큐 task 없음 — C3). make_host 숙주(`_NoAfter`)에서는 안 돈다.

### 2.1 전체 대응표 (지금 자리 → 새 자리, 빠짐없음)
| 지금 자리 | 위젯·항목 | U4 | U5 이후 |
|---|---|---|---|
| 메뉴바 보기 > 테마 > Light/Dark | `switch_theme('light'/'dark')` | 메뉴바 유지 | 설정 (D1) |
| 메뉴바 언어 > English/한국어 | `switch_language('en'/'ko')` | 메뉴바 유지 | 설정 (D1) |
| 메뉴바 Help > Keyboard Shortcuts (F1) | `show_keyboard_shortcuts_help` | 메뉴바 유지 | 도움말 (D1) |
| 메뉴바 Help > GitHub Releases | `open_github_link(None)` | 메뉴바 유지 | 도움말 (D1) |
| 헤더 | `header_time_label`(`_update_header_time`) | 상단 바 | 상단 바 |
| 옵션 줄 | Backups to Keep 스핀(`retention_count_var`) · Log Days 스핀 · Hidden 라디오 2 · System 라디오 2 | 백업 페이지(옵션 줄 통째) | 백업 |
| 사용자 카드 | `user_combo`/`user_var`, `load_users` | 상단 바 | 상단 바 |
| Main Operations 머리 | 소리 `sound_label`·`sound_toggle` | 백업 페이지(카드 통째) | 설정 |
| Main Operations | `backup_btn`·`restore_btn`·`filters_btn`·`driver_backup_btn`·`driver_restore_btn` | 백업 페이지(카드 통째) | 백업·복원·백업·백업·복원 |
| Tools 1줄 | `browser_profiles_btn`·`check_space_btn`·`save_log_btn`·`copy_btn`·`devmgr_btn` | 도구 페이지(카드 통째) | 백업·도구·로그·도구·도구 |
| Tools 2줄 | `schedule_btn`·`winget_export_btn`·`file_explorer_btn`·`nas_connect_btn`·`nas_disconnect_btn` | 도구 페이지 | 예약·백업·도구·도구·도구 |
| 진행 카드 | `progress_bar`·`status_label` | 하단 진행 줄 | 하단 진행 줄 + 홈 원형·사이드바 칩(읽기만) |
| Activity Log 카드 | `log_text`·스크롤바 | 로그 페이지 | 로그 + 홈(D9) |
| footer | "Press F1 for keyboard shortcuts" | 사이드바 하단 | 사이드바 하단 |
- U4 의 홈·복원·예약·설정·도움말 페이지는 제목만 있는 빈 페이지다(동작은 메뉴바·단축키·다른 페이지로 다 닿는다).

## 3. 메뉴바·단축키
- 단축키는 전부 그대로(`shortcuts.py:KeyboardShortcuts.setup_shortcuts`, 루트에 bind). 페이지와 무관하게 동작한다 — 버튼 위젯이 숨겨진 페이지에 있어도 메서드를 직접 부르기 때문. `bindings` golden 의 keys 는 U 단계 내내 같아야 한다.

| 키 | 호출 | 버튼이 있는 새 페이지 |
|---|---|---|
| Ctrl+B / Ctrl+R | start_backup_thread / start_restore_thread | 홈·백업 / 복원 |
| Ctrl+D / Ctrl+Shift+D | start_driver_backup_thread / start_driver_restore_thread | 백업 / 복원 |
| Ctrl+C · Ctrl+O · Ctrl+S | copy_files · open_file_explorer · save_log | 도구 · 도구 · 로그 |
| Ctrl+M · Ctrl+K · Ctrl+F · Ctrl+T | open_device_manager · check_space · open_filter_manager · schedule_backup | 도구 · 도구 · 백업(홈 칩) · 예약 |
| Ctrl+P · Ctrl+W | start_browser_profiles_backup_thread · start_winget_export_thread | 백업 · 백업 |
| Ctrl+N · Ctrl+Shift+N | open_nas_connect_dialog · open_nas_disconnect_dialog | 도구 · 도구 |
| F1 · F5 · Esc · Ctrl+Q | show_help · refresh_user_list · cancel_operation · safe_exit | 도움말 · 상단 바 · — · — |
| Ctrl+1~9 (`<Control-Key-n>`) | select_user_by_index | 상단 바 |
- 대화상자 단축키(`DialogShortcuts`·각 대화상자 `_setup_*_shortcuts`)는 U6 에서도 그대로.
- 주의: Ctrl+C 가 루트에 묶여 있어 CTkEntry 안 복사도 `copy_files` 를 부른다(지금 tk.Entry 와 같음, Q1 유지). 페이지 이동 단축키는 새로 넣지 않는다(D11).

## 4. 라이트·다크 토큰 (theme.py 에 `TOKENS` 추가, CTk 색은 (light, dark) 쌍)
| 토큰 | 라이트 | 다크 | Win11Theme 대응 키 |
|---|---|---|---|
| bg (창·내용 바탕) | #F3F4F6 | #202020 | bg (L #f5f5f5 / D #202020) |
| sidebar | #FFFFFF | #1B1B1B | bg_secondary |
| card | #FFFFFF | #2B2B2B | bg_elevated (L #e8e8e8 / D #252525) |
| card_hover | #F5F7FA | #323232 | hover |
| border | #E5E7EB | #3D3D3D | border |
| fg | #1F2328 | #FFFFFF | fg |
| fg_muted | #6B7280 | #B3B3B3 | fg_secondary |
| accent / hover | #0078D4 / #106EBE | #0078D4 / #1A86D9 | accent (D 는 지금 #0067C0, 흰 글자 대비 위해 통일) |
| accent_soft (칩·선택 메뉴) | #E1F1FF | #003D71 | accent_light |
| success · danger · warning | #107C10 · #D13438 · #F7630C | #0F7B0F · #C42B1C · #F7630C | 같음 |
| track (진행·스위치 꺼짐) | #E5E7EB | #484F58 | ToggleSwitch color_off |
- 지금 버튼마다 다른 색(빨강 백업 #DC3545, 파랑 복원 #007BFF, 보라 #8E44AD, 청록 #00796B 등, `_build_*_section`)은 U5 에서 accent(주 동작)·중립 card(보조)·danger(연결 해제) 3종으로 줄인다.
- 간격(px): 4·8·12·16·24·32. 내용 여백 24(지금 main_frame padx 24), 카드 안 16(지금 padx 16), 카드 사이 12, 버튼 사이 8.
- 반경: 카드 12 · 버튼·입력 8 · 칩 = 높이/2 · 사이드바 항목 8 · 원형 버튼 지름 120.
- 글꼴(D2 권장 A): 영어 "Segoe UI", 한국어 "Malgun Gothic". `ctk.CTkFont` 객체를 토큰으로 두고 `switch_language` 때 family 만 바꾼다. 크기는 CTk 단위(px): caption 11 · body 13 · body_bold 13 · section 16 bold · title 20 bold · display 28 bold(원형 %). 지금 pt 크기 대응: 9pt≈12px, 10pt≈13px, 11pt≈15px, 12pt=16px. 로그는 "Consolas" 그대로(tk.Text, `bold` 태그).
- 아이콘: Segoe MDL2 Assets 글리프(Win10·11 기본 포함, 파일·Pillow 불필요, D10). 예: 홈 E80F, 백업 E898, 복원 E777, 예약 E787, 도구 E90F, 로그 E7C3, 설정 E713, 도움말 E897.
- Win11Theme.LIGHT/DARK 값은 U6 마지막 발주까지 그대로 둔다(대화상자 트리 golden 이 그 값을 담고 있어 한 번에 바꿈). 새 CTk 위젯만 TOKENS 를 쓴다.

## 5. 창 크기·DPI
- 기본 1000×700, minsize 900×600(사이드바 220 + 내용 680, 버튼 3열이 들어가는 폭). 화면 작업 영역보다 크면 높이를 줄인다(1366×768 에서 720 은 제목 표시줄·작업 표시줄과 겹침). golden 은 geometry 를 기록하지 않아 영향 없음.
- CTk 는 `CTk()` 생성 때 프로세스를 Per-Monitor DPI 인식으로 바꾸고 자기 위젯·`geometry()` 를 배율로 늘린다. 지금 앱은 DPI 비인식이라, 켜면 아직 tk 로 남은 것(대화상자 Toplevel 의 픽셀 geometry 예: 단축키 도움말 `geometry("449x748")`, ToggleSwitch 50×24, Treeview 행 높이)은 125·150% 에서 작아지고 pt 글꼴만 커져 깨진다.
- 그래서 U4~U7 은 `ctk.deactivate_automatic_dpi_awareness()` 를 App 생성 전에 불러 지금과 같은 비인식 상태를 유지하고, U8 에서 켜고 100·125·150% 를 확인한다(D8). 켤 때 tk 위젯은 `ctk.ScalingTracker` 배율을 곱하는 helper 하나로 맞춘다.
- 제목 표시줄: CTk 가 Windows 에서 appearance mode 에 따라 제목 표시줄을 어둡게/밝게 바꾼다(사용자 결정 "색만 맞춤"과 일치). 이때 CTk 가 withdraw/deiconify 를 부르는지 — `_init_window` 의 `withdraw()` → `_init_startup` 끝 `deiconify()` 순서와 smoke `winfo_viewable()==0` — U2 에서 확인.

## 6. 테스트·golden 영향
| golden | U4 | U5 | U6 | 비고 |
|---|---|---|---|---|
| gui_snapshot `tree_{en,ko}_{dark,light}` | 재생성 | 페이지마다 재생성 | — | 창 구조 |
| gui_snapshot `theme_switch` | root_bg·toggle 만 바뀔 수 있음 | 소리 토글이 CTkSwitch 면 toggle_items | — | settings·dialogs 칸은 같아야 함 |
| gui_snapshot `language_switch` | texts 같아야 함(harness 보강 뒤) | 새 위젯 texts 추가, menu 칸은 D1=B 커밋에서 null | — | 기존 속성의 글자는 그대로 |
| gui_snapshot `bindings` | 같음 | D1=B 커밋에서 menu 칸만 비움 | 같음 | keys 는 끝까지 같음 |
| gui_snapshot `translations` | 키 추가 발주에서만(끝에 덧붙임) | 〃 | 〃 | 기존 키·값·순서 그대로 |
| gui_snapshot `settings_*` | 같음 | 같음 | 같음 | settings.json C3 |
| gui_dialogs `*_tree_*`(13쌍) | 같음 | 같음 | 대화상자별 재생성 | 창 구조 |
| gui_dialogs 동작 키(filter_edit_and_save, schedule_create_*, folder_select_*, nas_*_input_checks 등) | 같음 | 같음 | 같음 | |
| gui_cases, backup·restore·argv·cli·schedule_cases·tool_* | 같음 | 같음 | 같음 | GUI 밖 |
- smoke: `test_menu_bar_has_three_cascades` 는 D1=B 커밋에서 "메뉴바 없음"으로 바꾼다. `ACTION_BUTTONS`·진행률 `cget('maximum'/'value'/'mode')` 는 ttk.Progressbar 를 유지하므로 그대로.
- 위험과 막는 법(U2-2 에서 먼저 처리, golden 무변경으로 증명):
  - `own_members` 가 customtkinter 클래스도 빼도록(모듈 이름 `customtkinter` 시작). 안 하면 make_host·recording 이 CTk 메서드를 덮는다.
  - CTk 위젯 노드: `class`·`py`·`attr`·CTk 옵션 일부(`cget` 로 text·fg_color·hover_color·text_color·corner_radius·width·height·state·command 이름·font family/size) + 배치. 안쪽 `_canvas`·`_text_label` 등 자식과 캔버스 좌표는 기록하지 않는다(배율·그리기 엔진에 따라 달라짐).
  - `_texts`·`gh.button`·smoke 버튼 찾기: tk.Button 과 CTkButton 을 같이 보는 helper 로.
  - appearance mode 는 CTk 전역이다. make_host 의 theme 인자만으로는 light 트리가 안 나온다 → 숙주 만들 때 `ctk.set_appearance_mode` 를 맞추고 끝에 되돌린다. "system" 모드는 쓰지 않는다(darkdetect 폴링 after 루프).
  - (U2-2 에서 찾음) CTk 위젯은 부모가 tk.Tk·tk.Frame 등이면 부모 인스턴스에 `config`·`configure` 를 덮어쓴다(ctk_base_class). U4 에서 App·tk 프레임 아래 CTk 위젯을 두면 `vars(app)` 가 바뀌어 `app_unchanged`·recording 이 깨질 수 있다 → U4-1 에서 처리 방법 결정(App 이 ctk.CTk 면 해당 없음, tk 프레임은 CTkFrame 으로).
  - `test_live_app_matches_fresh_build`: 테스트 중 페이지를 바꾸면 pack 상태가 달라진다 → 페이지 전환 테스트는 끝에 홈으로 되돌린다.
- CTk 위젯은 `.config()` 를 막는다(AttributeError, configure 만). 지금 `.config(` 는 app.py 21·i18n.py 17·layout.py 6·widgets.py 5·bridge.py 3·app_log.py 2곳. 위젯을 CTk 로 바꾸는 발주에서 그 위젯을 만지는 줄만 `.configure(` 로 바꾼다(동작 같음, Q1).
- 바꾸지 않는 위젯: `user_combo`(ttk.Combobox — `current()`·`['values']` 를 shortcuts·load_users 가 씀), `progress_bar`(ttk.Progressbar — bridge 가 항목 접근), `log_text`(tk.Text — CTkTextbox 는 태그 font 금지). ttk 스타일·색만 토큰에 맞춘다.

## 7. 발주 목록 (발주 1건 = 커밋 1개, Files 에 golden JSON 제외. ★ = strings_en/ko 또는 같은 golden JSON 을 고쳐 단독으로)
| # | Goal | Files | 끝나는 기준 · 위험 |
|---|---|---|---|
| U2-1 | customtkinter·darkdetect 버전 고정(D7), 빌드에 `--collect-data customtkinter` | requirements.txt, .github/workflows/release.yml, scripts/build_pyinstaller_script.ps1 | runner: 로컬 ps1 exe 가 CTk 창 1개 띄우는 임시 스크립트(저장소 밖)로 실행됨. 위험: onefile 에서 테마 JSON 누락 → 실행 시 FileNotFoundError |
| U2-2 | 테스트 도구 CTk 대응(§6 위험 목록), 코드 무변경 | tests/fixtures/gui_harness.py, tests/characterization/test_gui_snapshot.py, tests/test_gui_smoke.py, tests/test_gui_harness_ctk.py(신규: 세션 App 아래 CTkFrame·CTkButton 노드·글자·버튼 찾기) | golden 전부 그대로 + 신규 테스트 통과. 위험: CTk 위젯이 세션 루트에 after 루프를 남김 → 테스트 끝에 destroy |
| U2-3 | 확인 결과(20550d0·후속 커밋, tests/test_gui_ctk_root.py 로 남김): withdraw 유지·menu 둘 다 됨·아이콘 실패 TclError·ttk·__class__ 교체·DPI 끔(배율 1, 비인식) 예, **`configure(bg=)` 는 조용히 무시**(→ U4-1 에서 `_init_window`·`apply_theme_to_all_widgets` 를 `fg_color` 로). 처음 계획: 확인만(커밋 없음, runner 보고): CTk 루트에서 `configure(bg=)`·`config(menu=)` 허용 여부, withdraw 상태 유지, DPI 끔 동작, Tk 1회 생성 세션에서 `__class__` 바꿔치기 | — | 보고서에 각 항목 예/아니오. 아니오면 U4-1 설계 수정 |
| U3-1 | 아키텍처 검사 a 에 customtkinter 추가 | tests/test_architecture.py, docs/adr/0002-패키지-구조.md | 위반 예시(core 에서 import customtkinter) 실패 단언 |
| U3-2 | TOKENS·CTkFont 토큰·아이콘 글리프 표(§4), 아직 아무도 안 씀 | ezbak/gui/theme.py | theme.py ≤500, golden 그대로 |
| U3-3 | 공용 위젯 Card·SidebarItem·StatTile·Chip·RingButton(원형, 캔버스) | ezbak/gui/components.py(신규), tests/test_gui_harness_ctk.py | 각 위젯 생성·텍스트·state 단위 테스트. 함수 60 |
| U3-4 ★ | 새 번역 키(사이드바 8·페이지 제목·상태 칩·홈 카드 문구, 끝에 덧붙임) | ezbak/gui/strings_en.py, ezbak/gui/strings_ko.py | translations 만 재생성, 나머지 golden 그대로. 키 이름은 `nav_*`·`home_*`·`status_*` |
| U4-1 ★ | App 기반 `tk.Tk` → `ctk.CTk`, appearance mode 를 theme 에 맞춤(`apply_theme_to_all_widgets`·`load_settings` 뒤), DPI 끔(D8), 1000×700·minsize | ezbak/gui/app.py, ezbak/gui/theme.py | 동작 golden 통과, tree·theme_switch(root 색) 만 재생성. 위험: CTk 가 `bg`·`menu` 를 거부 → U2-3 결과대로 `fg_color` 로 |
| U4-2 ★ | 사이드바·상단 바·하단 진행 줄·페이지 전환(`show_page`), 절 4개를 §2.1 U4 칸대로 통째 이동, footer → 사이드바 | ezbak/gui/layout.py, ezbak/gui/nav.py(신규 NavMixin: _build_sidebar·show_page·상태 칩 after. `class LayoutMixin(NavMixin)` 로 들여 app.py 무변경), ezbak/gui/pages/__init__.py(신규), ezbak/gui/i18n.py(nav 글자 갱신), tests/test_gui_smoke.py(show_page 8개) | tree 만 재생성, bindings·language_switch texts·settings 그대로. layout.py ≤500 |
| U5-1 ★ | 홈: RingButton(`home_backup_btn`)·사용자 제목·필터 칩·지표 2·시스템 카드(스레드+after 폴링)·로그 카드(D9) | ezbak/gui/pages/home.py, ezbak/gui/app.py(set_buttons_state), ezbak/gui/nav.py | 스크린샷. tree 재생성. 위험: get_system_info DEBUG print → 숙주에선 안 부름 |
| U5-2 ★ | 백업 페이지: 옵션·보관·추가 백업, 버튼 색 3종 | ezbak/gui/pages/backup.py, ezbak/gui/layout.py, ezbak/gui/i18n.py | 〃. 옵션 trace 저장(`_build_backup_section` 의 trace_add) 그대로 |
| U5-3 ★ | 복원 페이지 | ezbak/gui/pages/restore.py, ezbak/gui/layout.py | 〃 |
| U5-4 ★ | 예약 백업 페이지 | ezbak/gui/pages/schedule.py, ezbak/gui/layout.py | 〃 |
| U5-5 ★ | 도구 페이지 | ezbak/gui/pages/tools.py, ezbak/gui/layout.py | 〃 |
| U5-6 ★ | 로그 페이지 + 홈과 log_text 공유(D9) | ezbak/gui/pages/logs.py, ezbak/gui/nav.py | 〃. 로그 golden(작업 로그·UI 로그 문구) 그대로 |
| U5-7 ★ | 설정·도움말 페이지, 메뉴바 제거(D1=B), 소리 토글 이동, `update_ui_texts` 의 `create_menu_bar` 호출 정리 | ezbak/gui/pages/settings_page.py, ezbak/gui/pages/about.py, ezbak/gui/app.py, ezbak/gui/i18n.py, tests/test_gui_smoke.py | tree·bindings(menu 칸)·language_switch(menu 칸) 재생성. 테마·언어 전환 대화상자·settings 저장은 그대로 |
| U6-1 ★ | Win11Dialog·`create_win11_button`·단축키 도움말 창을 토큰으로 | ezbak/gui/dialogs/message.py, ezbak/gui/widgets.py, ezbak/gui/shortcuts.py | win11_* 트리만 재생성 |
| U6-2 ★ | 필터 대화상자 | ezbak/gui/dialogs/filters.py | filter_tree_* 만 |
| U6-3 ★ | 원본 선택 + Treeview ttk 스타일 | ezbak/gui/dialogs/sources.py, ezbak/gui/dialogs/sources_tree.py | sources_tree_* 만. 체크 표시 문자열(`_format`) 그대로 |
| U6-4 ★ | 예약 대화상자 | ezbak/gui/dialogs/schedule.py, ezbak/gui/dialogs/schedule_layout.py | schedule_tree_* 만. task_data 키 그대로 |
| U6-5 ★ | NAS 2개·폴더 선택·교차 사용자 경고 | ezbak/gui/dialogs/nas.py, ezbak/gui/dialogs/restore_select.py | nas_*_tree·folder_select_tree 만. 암호 단언 그대로 |
| U6-6 ★ | Win11Theme.LIGHT/DARK 를 TOKENS 값으로 맞춤 | ezbak/gui/theme.py | 모든 tree 재생성(마지막 1회) |
- 순서: U2-1 → U2-2 → U2-3 → U3-1~3 → U3-4 → U4-1 → U4-2 → U5-1~7 → U6-1~6. ★ 는 gui_snapshot 또는 gui_dialogs 를 재생성하므로 서로 병렬 금지.
- 줄 수: layout.py 는 U5 에서 절을 pages/ 로 넘기며 줄어야 한다. nav.py ~200, components.py ~350 예상(넘으면 components_ring.py 로). 새 이름은 stdlib 과 겹치지 않는다(`pages`·`nav`·`components`·`logs`·`about`·`settings_page`; `logging`·`help`·`settings` 는 피함).
- 매 커밋: 메시지에 "창 구조 golden 재생성(ADR 0004): <키>" 를 적는다. gui/CLAUDE.md 표(파일 역할)는 U4-2·U5-7 뒤 메인이 갱신.

## 8. 결정할 항목 (사용자)
결정(2026-10-27): D1 B · D2 A · D3 A · D4 A · D5 A · D6 A · D7 A · D8 B · D9 B · D10 A · D11 A · D12 A (모두 권장안).
- D1 메뉴바: A) 유지 / B) U4 에는 유지, U5-7 에서 없애고 설정·도움말로(권장: Windows 기본 메뉴바는 tk.Menu `bg` 를 무시해 다크·사이드바와 안 맞고, 4개 명령 모두 새 자리가 있다. 단축키는 그대로라 기능 손실 없음) / C) 끝까지 유지하고 도움말 항목만 페이지에도.
- D2 글꼴: A) 영어 Segoe UI·한국어 맑은 고딕, 언어 전환 때 CTkFont family 교체(권장: 시스템 글꼴이라 exe 크기·등록 코드 0, 한글 메트릭 정확) / B) Pretendard 동봉(+약 3~6MB, OFL, `ctk.FontManager.load_font` — 그림과 가장 비슷) / C) 맑은 고딕 하나.
- D3 다크 색상표: A) 지금 Win11 DARK 기반 중립 회색(§4, 권장: 기존 사용자에게 익숙하고 대화상자 golden 변경이 작음) / B) 그림 라이트를 반전한 남색 계열(#0F172A 바탕·#1E293B 카드).
- D4 설정이 없을 때 기본 테마: A) 다크 유지(권장: Q1, `Win11Theme(initial_mode='dark')`·settings 없음 golden 그대로) / B) 라이트(그림과 같음, Q1 예외 커밋 필요). 'system' 값은 settings.json 에 새 값이라 쓰지 않는다(C3).
- D5 사이드바 접기: A) 없음, 폭 220 고정(권장: U4~U6 범위 최소화, minsize 900 이면 충분) / B) 접기 버튼(아이콘만 64) / C) 창 폭 < 960 에서 자동 접기.
- D6 언어·테마 전환 위치: A) 설정 페이지만(권장: 메뉴바의 자리 그대로 옮김) / B) 사이드바 하단 빠른 토글 + 설정. 둘 다 전환 뒤 "다시 시작" 안내 대화상자는 그대로(Q1, `theme_changed_msg`).
- D7 CTk 버전: A) `customtkinter==5.2.2`·`darkdetect==0.8.0` 정확히 고정(권장: C2 재현 가능한 onefile, 5.2.2 가 마지막 안정판) / B) `~=5.2`.
- D8 DPI: A) U4 부터 CTk 자동 DPI 인식 / B) U4~U7 은 끄고 U8 에서 켜고 확인(권장: 지금도 비인식이라 Q1, tk 로 남은 대화상자가 U6 전까지 깨지지 않음).
- D9 공유 위젯: progress_bar·status_label 은 하단 공통 줄(결정 불필요). log_text 를 홈과 로그 두 곳에: A) 로그 페이지에만, 홈엔 상태 줄만 / B) 한 위젯을 `pack(in_=…)`+`lift()` 로 보이는 페이지에 옮김(권장: 위젯 한 벌·bridge·app_log 무변경, harness `layout()` 은 `in` 을 기록 안 함) / C) 홈에 읽기 전용 사본 Text(`log` 이 두 곳에 써야 해 app_log 변경).
- D10 아이콘: A) Segoe MDL2 Assets 글리프(권장: 파일·Pillow·빌드 변경 없음, Win10·11 공통) / B) PNG + CTkImage(Pillow 의존 추가, assets 폴더·`--add-data`).
- D11 페이지 이동 단축키: A) 추가 안 함(권장: bindings golden keys 고정, Ctrl+1~9 는 사용자 선택이 씀) / B) Alt+1~8(U7 이후 기능 추가로).
- D12 페이지 코드 형태: A) `gui/pages/<페이지>.py` 모듈 함수 `build(app, parent)`(권장: 믹스인 더 안 늘림, recording 이 안 건드림, 페이지별 커밋이 파일 1개) / B) 페이지 믹스인 8개 / C) CTkFrame 서브클래스(App 속성 이름이 page.xxx 로 바뀌어 bridge·smoke 가 깨짐).

## 9. U8 발주 (2026-10-27, U7 은 뒤로 — 사용자 결정 "U8 먼저")
확인한 사실(customtkinter 5.2.2 소스):
- `ScalingTracker.activate_high_dpi_awareness()` 는 `CTk.__init__` 안에서 tk.Tk 를 만든 **뒤** 불린다. Tk 는 비인식 상태에서 `tk scaling`(pt→px)을 정하므로, 켜기만 하면 tk 위젯의 pt 글꼴이 125·150% 에서 100% 크기로 남는다(CTk 글꼴은 음수 px 에 배율을 곱해 따로 맞음).
- CTk 위젯·`CTk.geometry()`·`minsize()` 는 스스로 배율을 곱한다. tk.Toplevel 대화상자의 `geometry("WxH")`·`wraplength`·Canvas 크기·Treeview 행 높이는 그대로 px 다.
- 비인식(`deactivate_automatic_dpi_awareness`)이거나 Windows 가 아니면 `get_window_dpi_scaling` 은 1 이다. 100%(CI 러너 96 DPI)도 1 → golden 무변경.
- `check_dpi_scaling` 은 CTk 위젯이 든 창에 `block_update_dimensions_event` 를 부른다. 대화상자(tk.Toplevel)에는 CTk 위젯이 없어 App 만 대상이다(대화상자에 CTk 위젯을 넣으면 모니터 이동 때 AttributeError).

| # | Goal | Files | 끝나는 기준 |
|---|---|---|---|
| U8-1 | `gui/dpi.py`: `factor(widget)`(App 루트의 CTk DPI 배율, 실패·비인식 1)·`px`·`geometry`·`sync_tk_scaling`(배율이 1 이 아닐 때만 `tk scaling` = 배율×96/72). Win11Dialog 2개·폴더 선택·F1 도움말·소리 토글에 적용 | ezbak/gui/dpi.py(신규), dialogs/message.py, dialogs/restore_select.py, shortcuts.py, layout.py, tests/test_gui_dpi.py(신규) | golden 무변경(배율 1). 단위 테스트는 배율을 바꿔 px·geometry 문자열 |
| U8-2 | 나머지 tk 대화상자: 필터 3개·원본 2개 + Treeview 행 높이(배율 ≠ 1 일 때만)·예약 2개·NAS 2개(+wraplength) | dialogs/filters.py, dialogs/sources.py, dialogs/schedule.py, dialogs/nas.py | golden 무변경 |
| U8-3 | DPI 켜기(D8): `deactivate_automatic_dpi_awareness` 삭제, `super().__init__()` 바로 뒤 `sync_tk_scaling(self)` | ezbak/gui/app.py, tests/test_gui_ctk_root.py(문구) | golden 무변경(CI 100%). 실기기 125·150% 는 U8-4 체크리스트 |
| U8-4 | 실기기 체크리스트 `docs/nav-drafts/u8-checklist.md`: 배율 100·125·150%·모니터 이동, 라이트·다크, 한·영 문구 길이(긴 문구 표), exe 빌드·실행, 대화상자 13종·페이지 8개 | 문서 1개 | 사용자가 Windows 실기기에서 표를 채움 |
| U8-5 | 문서 정리(roadmap·gui CLAUDE.md·structure·mem), 체크리스트 결과 반영 → main 병합(PR, 사용자 승인) → P6 | 문서 | — |

## 10. U7 새 데이터 카드 계획 (2026-10-27, P6 병합 #47 뒤, 기준 main ebdffe6)
U7 은 기능 추가다(ADR 0004 §1). 홈에 카드 다섯 가지: 복사 속도 그래프·마지막 백업·다음 예약·시스템 이름·대상 여유 공간.

확인한 사실:
- 진행 값: 백업·복원·복사는 `update_progress_max(전체 바이트)` 뒤 `update_progress(바이트)` 다(엔진 `progress_bytes`, 복사 `total_bytes`). 상태 칩 타이머(nav, 0.5초)가 이미 `progress_bar` 값을 읽는다 → 속도는 이 타이머에서 값 차이/시간으로 구할 수 있다(bridge·큐 무변경). 퍼센트 작업(최대값 100, 도구)은 속도를 보이지 않는다.
- 마지막 백업: 상태로 남는 곳이 없다(대상 폴더는 매번 askdirectory, settings.json 에 없음). settings.json 키는 C3 고정.
- 다음 예약: 스케줄 config JSON(`%APPDATA%\ezBAK\ezBAK_Schedule\*.json`)에 schedule·time_str·monthly_day·created_at 가 있다. 첫 실행은 created_at 다음 날 time_str(build_task_xml), Weekly 는 일요일 고정, Monthly 는 monthly_day. schtasks 조회 없이 순수 함수로 계산할 수 있다(앱 밖에서 지우거나 끈 작업은 모름).
- 시스템 이름: `get_system_info` → 레지스트리, 실패 시 wmic(proc). 창 생성 중에는 부르지 않는다(§0) → 작업 스레드 + after.
- 여유 공간: `winsys.free_space(path)`(실패 시 999TB 대체) — 대상 경로가 상태로 없다(마지막 백업과 같은 문제).

| # | Goal | Files | 끝나는 기준 |
|---|---|---|---|
| U7-1 | 순수 계산: `automation.scheduler.next_run(config, now)`·`next_runs(configs, now)`(가장 이른 것, 잘못된 값은 건너뜀) | ezbak/automation/scheduler.py, tests/test_automation_scheduler.py | 단위 테스트(Daily·Weekly 일요일·Monthly 31일·created_at 당일·없음) |
| U7-2 | 마지막 백업 기록(D1): 저장소·읽기 함수 + GUI 백업 성공 시(Headless 도, D1) 쓰기 | D1 에 따라 | 기존 golden 무변경(새 파일만 생김 — 테스트는 임시 APPDATA) |
| U7-3 ★ | 새 번역 키(카드 제목·문구, 끝에 덧붙임) | strings_en.py, strings_ko.py | translations 만 재생성 |
| U7-4 | 공용 위젯 `Sparkline`(CTkCanvas, 최근 N 개 값 꺾은선, 토큰 색) | components.py, tests/test_gui_components.py | 단위 테스트 |
| U7-5 ★ | 홈 카드: 속도(Sparkline + "12.3 MB/s", 타이머 표본) · 마지막 백업(시각·사용자·폴더) · 다음 예약(작업 이름·시각) · 시스템 이름(스레드) · 여유 공간(D3, 스레드) | pages/home.py, nav.py, (필요 시 pages/home_data.py) | tree 재생성(D5), 실행 중 글자는 LIVE_TEXTS |

## 11. U7 결정할 항목
결정(2026-10-27): D1 A · D2 A · D3 A · D4 A · D5 A(허용) (모두 권장안).
진행: U7-1 da635d3 · U7-2 1e30d7c · U7-3 b0376d4+golden 372868b · U7-4 884e0b5(테스트 수정 42bc338) · U7-5 1c65e2c+golden 849651b(첫 재생성은 U7-4 테스트 실패로 멈춰 42bc338 에서 다시). 홈 데이터는 창이 보일 때만 읽는다(테스트·시작 중엔 자리표시) — 실기기 확인은 u8-checklist 에 더할 것.
- D1 마지막 백업 저장: A) 새 파일 `%APPDATA%\ezBAK\state.json`(settings.json 은 그대로, GUI·Headless 백업 성공 때 시각·사용자·대상·폴더 기록, 권장) / B) settings.json 에 새 키(C3 변경) / C) 이번 실행 동안만(메모리).
- D2 다음 예약: A) 스케줄 config JSON 으로 계산(권장: 외부 명령 없음, 빠름) / B) `schtasks /query`(정확하지만 느리고 외부 명령).
- D3 여유 공간 대상: A) 마지막 백업의 대상 드라이브, 없으면 카드에 "-"(권장) / B) 시스템 드라이브 C:.
- D4 속도 그래프: A) 최근 30초(0.5초 × 60) 꺾은선 + 현재 속도, 바이트 작업만(권장) / B) 작업 전체 구간.
- D5 golden: U7-5 커밋에서 창 구조 golden(tree)을 Q1 예외로 다시 만든다(ADR 0004 §3 은 U4~U6 만 적음 → U7 추가 허용 여부).
