"""Korean strings for Translator (moved verbatim from ezBAK.py in P5-1).

Key order is fixed (C3).
"""

KO = {
    'app_title': 'ezBAK',
    'menu_file': '파일',
    'menu_view': '보기',
    'menu_language': '언어',
    'menu_theme': '테마',
    'theme_light': '라이트 모드',
    'theme_dark': '다크 모드',
    'lang_english': 'English',
    'lang_korean': '한국어',
    'select_user': '사용자 선택:',
    'main_operations': '주요 작업',
    'tools_utilities': '도구 및 유틸리티',
    'backup_data': '데이터 백업',
    'restore_data': '데이터 복원',
    'filters': '필터',
    'backup_drivers': '드라이버 백업',
    'restore_drivers': '드라이버 복원',
    'browser': '브라우저',
    'check_space': '공간 확인',
    'save_log': '로그 저장',
    'copy_data': '데이터 복사',
    'device_mgr': '장치 관리자',
    'schedule': '예약',
    'export_apps': '앱 내보내기',
    'explorer': '탐색기',
    'connect_nas': 'NAS 연결',
    'disconnect': '연결 해제',
    'activity_log': '활동 로그',
    'sound': '소리',
    'operation_finished': '작업 완료',
    'backup_complete': '백업 완료!',
    'ready': '준비',
    'select_user_begin': '시작하려면 사용자 선택',
    'dialog_ok': '확인',
    'dialog_yes': '예',
    'dialog_no': '아니오',
    'dialog_cancel': '취소',
    'notice_title': '안내:',
    'notice_select_user': '시작하려면 사용자를 선택하세요',
    'notice_filters': '필터:제외 규칙:이름->ondrive*',
    'notice_hidden_system': '숨김 및 시스템 속성 제외됩니다',
    'notice_appdata': 'AppData 폴더는 백업되지 않습니다',
    'notice_detailed_log': '작업 중 상세 로그 파일로 저장됩니다',
    'notice_responsibility': '이 프로그램의 사용에 대한 책임은 전적으로 사용자에게 있습니다',
    # Dialog titles and messages
    'language_changed': '언어 변경',
    'language_changed_msg': '한국어로 변경합니다. 프로그램 재시작 후 적용됩니다.',
    'theme_changed': '테마 변경됨',
    'theme_changed_msg': '테마가 {0} 모드로 변경되었습니다.\n완전한 적용을 위해 애플리케이션을 다시 시작하세요.',
    'confirm_exit': '종료 확인',
    'confirm_exit_msg': 'ezBAK을 종료하시겠습니까?',
    'error': '오류',
    'warning': '경고',
    'info': '정보',
    'space_check': '공간 확인',
    'sufficient_space': '✓ 충분한 공간이 있습니다.',
    'insufficient_space': '⚠ 공간이 부족합니다.',
    'no_log_content': '저장할 로그 내용이 없습니다.',
    'error_saving_log': 'UI 로그 저장 중 오류 발생: {0}',
    'select_user_warning': '사용자를 선택하세요.',
    'overwrite_warning': '\'{0}\'의 기존 데이터를 덮어씁니다. 계속하시겠습니까?',
    'task_creation_failed': '작업 생성 실패',
    'task_deletion_failed': '작업 삭제 실패',
    'schedule_created': '예약 생성됨',
    'task_scheduler': '작업 스케줄러',
    'task_deleted': '작업 \'{0}\'이(가) 삭제되었습니다.',
    'task_name_required': '삭제하려면 작업 이름이 필요합니다.',
    'failed_to_delete_task': '작업 삭제 실패:\n{0}',
    'source_selection_failed': '소스 선택 실패: {0}',
    'filter_manager_error': '필터 관리자 오류: {0}',
    'unc_path_required': 'UNC 경로 \\\\로 시작해야 합니다 (예: \\\\Server\\share)',
    'password_required': '사용자가 제공된 경우 비밀번호가 필요합니다.',
    'invalid_drive_letter': '올바른 드라이브 문자를 입력하세요 (예: Z)',
    'connected': '연결됨',
    'open_share_now': '지금 공유를 여시겠습니까?',
    'nas_connect': 'NAS 연결',
    'failed_to_connect': '연결 실패.\n\n{0}',
    'nas_connect_error': 'NAS 연결 오류: {0}',
    'unc_path_must_start': 'UNC 경로는 \\\\로 시작해야 합니다 (예: \\\\server\\share)',
    'map_drive': '드라이브 매핑',
    'map_to_drive_letter': '이 공유를 드라이브 문자에 매핑하시겠습니까?',
    'persistence': '지속성',
    'reconnect_at_signin': '로그인 시 다시 연결하시겠습니까? (지속)',
    'nas_disconnect': 'NAS 연결 해제',
    'disconnected': '연결 해제됨: {0}',
    'failed_to_disconnect': '연결 해제 실패.\n\n{0}',
    'nas_disconnect_error': 'NAS 연결 해제 오류: {0}',
    'input_target': '대상을 입력하세요 (드라이브 문자 또는 UNC)',
    'invalid_drive_or_unc': '잘못된 드라이브 문자 또는 UNC 경로입니다.',
    'system_mismatch_warning': '시스템 불일치 경고',
    'unable_to_open_devmgr': '장치 관리자를 열 수 없습니다:\n{0}',
    'task_name_required_create': '작업 이름이 필요합니다.',
    'destination_required': '대상 폴더가 필요합니다.',
    'invalid_time_format': '잘못된 시간 형식입니다. HH:MM 형식을 사용하세요 (예: 14:30)',
    'failed_to_create_schedule': '예약 생성 실패: {0}',
    'task_name_required_delete': '삭제하려면 작업 이름이 필요합니다.',
    'failed_to_delete_schedule': '예약 삭제 실패: {0}',
    'please_enter_pattern': '패턴을 입력하세요.',
    'continue_anyway': '계속하시겠습니까?',
    'dest_folder_missing': '대상 폴더가 없습니다',

    # Dialog titles
    'select_sources': '소스 선택',
    'select_destination_folder': '대상 폴더 선택',
    'check_required_space': '필요한 공간 확인',
    'schedule_backup': '예약 백업',
    'network_share_connect': '네트워크 공유 - 연결',
    'network_share_disconnect': '네트워크 공유 - 연결 해제',

    # NAS Connect dialog labels
    'unc_path_label': 'UNC 경로 (예: \\\\NAS\\share)',
    'map_as_drive_letter': '드라이브 문자로 매핑',
    'letter': '문자:',
    'username_optional': '사용자 이름 (선택사항)',
    'password_required_if_username': '비밀번호 (사용자 이름 제공 시 필수)',
    'reconnect_at_logon_persistent': '로그인 시 다시 연결 (영구)',

    # NAS Disconnect dialog labels
    'target_drive_or_unc': '대상 (드라이브 문자 예: Z: 또는 UNC 예: \\\\Server\\Share)',

    # Filter Manager
    'filter_manager': '필터 관리자',
    'include_rules': '포함 규칙',
    'exclude_rules': '제외 규칙',
    'add': '추가',
    'remove': '제거',
    'clear_all': '모두 지우기',
    'help': '도움말',
    'save': '저장',
    'add_include_rule': '포함 규칙 추가',
    'add_exclude_rule': '제외 규칙 추가',
    'rule_type': '규칙 유형:',
    'pattern': '패턴:',
    'extension': '확장자',
    'file_folder_name': '이름',
    'full_path': '경로',
    'extension_desc': '파일 확장자 (예: tmp, log, bak)',
    'name_desc': '와일드카드를 사용한 파일/폴더 이름 (예: Thumbs.db, *cache*)',
    'path_desc': '와일드카드를 사용한 전체 경로 (예: */temp/*, *\\AppData\\*)',
    'add_rule': '규칙 추가',
    'examples_ext': '예: tmp, log, bak, cache (점 없이)',
    'examples_name': '예: Thumbs.db, *.tmp, desktop.ini, *cache*',
    'examples_path': '예: */Temp/*, *\\AppData\\Local\\*, */cache/*',
    'filter_info_include': '포함: 파일이 포함되려면 최소 하나의 규칙과 일치해야 합니다',
    'filter_info_exclude': '제외: 규칙과 일치하는 파일은 건너뜁니다 (우선순위)',

    # Filter Manager Help Dialog
    'filter_manager_help_title': '필터 매니저 도움말',
    'filter_manager_help_content': """필터 매니저 도움말

키보드 단축키:
┌─────────────────────────────────────────────────────────────┐
│ Ctrl + N         포함 규칙 추가                              │
│ Ctrl + E         제외 규칙 추가                              │
│ Ctrl + S         필터 저장                                   │
│ Ctrl + R         활성 목록 지우기                            │
│ Delete           선택한 항목 제거                            │
│ Tab              목록 간 전환                                │
│ Esc              취소                                       │
│ Enter            저장                                       │
│ F1               도움말 표시                                 │
└─────────────────────────────────────────────────────────────┘

필터 유형:
┌─────────────────────────────────────────────────────────────┐
│ 📄 확장자        파일 확장자 (예: tmp, log, bak)             │
│ 📁 이름          와일드카드를 사용한 파일/폴더 이름           │
│                  예: Thumbs.db, *.tmp, *cache*              │
│ 🗂️ 경로          와일드카드를 사용한 전체 경로               │
│                  예: */temp/*, *\\AppData\\*                │
└─────────────────────────────────────────────────────────────┘

필터 작동 방식:
• 포함 규칙: 파일이 포함되려면 최소 하나의 규칙과 일치해야 합니다
• 제외 규칙: 규칙과 일치하는 파일은 건너뜁니다
• 제외 규칙이 포함 규칙보다 우선 적용됩니다
• 패턴 매칭에 와일드카드(*)를 사용하세요
• 경로는 슬래시(/) 또는 백슬래시(\\)를 사용합니다

예제:
• 임시 파일 제외: [확장자] tmp
• 캐시 폴더 제외: [이름] *cache*
• 문서만 포함: [확장자] pdf, [확장자] docx
• 시스템 경로 제외: [경로] */Windows/System32/*""",
    'help_close_esc': '닫기 (Esc)',

    # Keyboard Shortcuts Help Dialog
    'keyboard_shortcuts_title': '도움말 - 키보드 단축키',
    'keyboard_shortcuts_content': """ezBAK 키보드 단축키

┌─────────────────────────────────────────┐
│           주요 기능                      │
├─────────────────────────────────────────┤
│ Ctrl + B        데이터 백업              │
│ Ctrl + R        데이터 복원              │
│ Ctrl + D        드라이버 백업            │
│ Ctrl + Shift + D 드라이버 복원           │
├─────────────────────────────────────────┤
│           파일 관리                      │
├─────────────────────────────────────────┤
│ Ctrl + C        데이터 복사              │
│ Ctrl + O        탐색기                   │
│ Ctrl + S        로그 저장                │
├─────────────────────────────────────────┤
│           도구 및 유틸리티               │
├─────────────────────────────────────────┤
│ Ctrl + M        장치 관리자              │
│ Ctrl + K        공간 확인                │
│ Ctrl + F        필터                     │
│ Ctrl + T        예약                     │
├─────────────────────────────────────────┤
│           브라우저 및 앱                 │
├─────────────────────────────────────────┤
│ Ctrl + P        브라우저                 │
│ Ctrl + W        앱 내보내기              │
├─────────────────────────────────────────┤
│           네트워크                       │
├─────────────────────────────────────────┤
│ Ctrl + N        NAS 연결                 │
│ Ctrl + Shift + N 연결 해제               │
├─────────────────────────────────────────┤
│           일반                           │
├─────────────────────────────────────────┤
│ F1              도움말                   │
│ F5              (작동 안 함) 사용자 목록 새로고침    │
│ Ctrl + 1~9      사용자 빠른 선택               │
│ Escape          (작동 안 함) 현재 작업 취소        │
│ Ctrl + Q        프로그램 종료            │
└─────────────────────────────────────────┘

팁 : 단축키는 대소문자를 구분하지 않습니다.""",
    'help_exit_esc': '종료 (Esc)',
    'data_restore_cancelled': '사용자에 의해 데이터 복원이 취소되었습니다.',
    'driver_restore_cancelled': '사용자에 의해 드라이버 복원이 취소되었습니다.',
    'cross_user_restore_warning': '크로스 사용자 복원 경고\n\n다음 백업을 복원하려고 합니다:\n  {0}의 백업\n대상:\n  {1} 계정\n\n폴더: {2}\n\n데이터 충돌이나 권한 문제가 발생할 수 있습니다.\n\n계속하시겠습니까?',
    'select_backup_folder': '백업 폴더 선택',
    'select_backup_folder_msg': '복원할 백업 폴더를 선택하세요:',
    'no_browser_profiles': '백업할 브라우저 프로필 데이터를 찾지 못했습니다.',
    'confirm_browser_profiles_title': '브라우저 프로필 백업',
    'confirm_browser_profiles': '사용자 \'{0}\'의 브라우저 프로필 전체를 백업할까요?\n\n원본: {1}',
    # U3-4: sidebar, status chip and home cards of the redesigned window (ADR 0004)
    'nav_home': '홈',
    'nav_backup': '백업',
    'nav_restore': '복원',
    'nav_schedule': '예약 백업',
    'nav_tools': '도구',
    'nav_logs': '로그',
    'nav_settings': '설정',
    'nav_help': '도움말',
    'status_idle': '대기',
    'status_working': '작업 중 {0}%',
    'home_start_backup': '백업 시작',
    'home_selected_user': '선택한 사용자',
    'home_filters': '필터 {0}개',
    'home_retention': '보관 개수',
    'home_log_days': '로그 보관 일수',
    'home_system_card': '시스템',
    'home_system_name': '시스템 이름',
    'home_user_count': '사용자 수',
    'home_hidden_files': '숨김 파일',
    'home_system_files': '시스템 파일',
    'home_included': '포함',
    'home_excluded': '제외',
    'home_loading': '불러오는 중...',
    # U5-2: card titles of the backup and restore pages
    'backup_main': '사용자 데이터 백업',
    'backup_options': '백업 옵션',
    'backup_extra': '추가 백업',
    'restore_main': '사용자 데이터 복원',
    # U5-3: restore page
    'restore_hint': '백업 폴더를 선택한 사용자의 프로필로 복원합니다. 기존 데이터를 덮어쓰기 전에 확인을 받고, 그다음 백업 폴더를 고릅니다.',
    'restore_drivers_hint': '드라이버 복원은 드라이버 백업 폴더의 드라이버를 설치합니다.',
    # U5-4: scheduled backup page
    'schedule_main': '예약 백업',
    'schedule_hint': '선택한 사용자를 창 없이 백업하는 Windows 예약 작업을 등록합니다. 대화상자에서 예약을 만들거나 기존 예약을 삭제합니다.',
    # U5-5: tools page card titles
    'tools_disk': '공간·복사',
    'tools_system': '시스템·파일',
    'tools_nas': '네트워크 드라이브(NAS)',
    # U5-7: help page buttons (were the fixed English Help menu entries)
    'help_shortcuts': '단축키 보기 (F1)',
    'help_github': 'GitHub 릴리스',
    # U7-3: home data cards (speed, last backup, next schedule, free space)
    'home_speed_card': '복사 속도',
    'home_speed_idle': '실행 중인 복사 없음',
    'home_last_backup_card': '마지막 백업',
    'home_last_backup_none': '아직 기록된 백업 없음',
    'home_mode_gui': '수동',
    'home_mode_headless': '예약',
    'home_next_schedule_card': '다음 예약 백업',
    'home_next_schedule_none': '예약된 백업 없음',
    'home_free_space': '여유 공간',
}
