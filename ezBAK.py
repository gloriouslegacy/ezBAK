import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import traceback
import os
import shutil
import threading
import queue
from datetime import datetime, timedelta
import argparse
import sys
import ctypes
import subprocess
import webbrowser
import fnmatch
import smtplib
import stat
import re
import time
import json

# ============================================================
# Windows 11 Style Theme System
# ============================================================

class Win11Theme:
    """Windows 11 style theme with light and dark mode support"""

    # Light Theme Colors (Windows 11 Light)
    LIGHT = {
        'bg': '#f5f5f5',              # Main background
        'bg_secondary': '#efefef',     # Secondary background
        'bg_elevated': '#e8e8e8',      # Elevated surfaces
        'fg': '#000000',              # Text color
        'fg_secondary': '#605E5C',    # Secondary text
        'accent': '#0078D4',          # Accent color (blue)
        'accent_hover': '#106EBE',    # Accent hover
        'accent_light': '#E1F1FF',    # Light accent background
        'success': '#107C10',         # Success/green
        'success_hover': '#0E6B0E',
        'danger': '#D13438',          # Danger/red
        'danger_hover': '#B32B2F',
        'warning': '#F7630C',         # Warning/orange
        'info': '#0078D4',            # Info/blue
        'border': '#E1DFDD',          # Border color
        'divider': '#EDEBE9',         # Divider
        'shadow': '#00000015',        # Shadow
        'disabled': '#A19F9D',        # Disabled state
        'hover': '#F5F5F5',           # Hover state
        'pressed': '#EBEBEB',         # Pressed state
        'text_bg': '#ffffff',         # Text area background
        'btn_bg': '#0078D4',          # Button background
        'btn_fg': '#ffffff',          # Button text
        'scrollbar_bg': '#C8C6C4',    # Scrollbar background
    }

    # Dark Theme Colors (Windows 11 Style)
    DARK = {
        'bg': '#202020',              # Main background - Windows 11 dark
        'bg_secondary': '#2b2b2b',     # Secondary background - Windows 11 elevated
        'bg_elevated': '#252525',      # Elevated surfaces
        'fg': '#ffffff',              # Text color - Pure white
        'fg_secondary': '#b3b3b3',    # Secondary text - Light gray
        'accent': '#0067C0',          # Accent color - Windows 11 blue
        'accent_hover': '#005A9E',    # Accent hover - Darker blue
        'accent_light': '#003D71',    # Dark accent background
        'success': '#0F7B0F',         # Success/green - Windows 11 green
        'success_hover': '#0D6B0D',
        'danger': '#C42B1C',          # Danger/red - Windows 11 red
        'danger_hover': '#A82315',
        'warning': '#F7630C',         # Warning/orange - Windows 11 orange
        'info': '#0078D4',            # Info/blue - Windows 11 info
        'border': '#3d3d3d',          # Border color - Subtle gray
        'divider': '#333333',         # Divider
        'shadow': '#00000040',        # Shadow
        'disabled': '#6d6d6d',        # Disabled state
        'hover': '#2d2d2d',           # Hover state
        'pressed': '#1a1a1a',         # Pressed state
        'text_bg': '#2b2b2b',         # Text area background
        'btn_bg': '#0067C0',          # Button background
        'btn_fg': '#ffffff',          # Button text
        'scrollbar_bg': '#5d5d5d',    # Scrollbar background
    }

    def __init__(self, initial_mode='dark'):
        self.current_mode = initial_mode
        self.colors = self.DARK if initial_mode == 'dark' else self.LIGHT

    def switch_mode(self):
        """Toggle between light and dark mode"""
        self.current_mode = 'light' if self.current_mode == 'dark' else 'dark'
        self.colors = self.DARK if self.current_mode == 'dark' else self.LIGHT
        return self.current_mode

    def get(self, color_name):
        """Get color value by name"""
        return self.colors.get(color_name, '#000000')

# ============================================================
# Translation System for Multi-language Support
# ============================================================

class Translator:
    """Simple translation system for multi-language support"""

    TRANSLATIONS = {
        'en': {
            'app_title': 'ezBAK',
            'menu_file': 'File',
            'menu_view': 'View',
            'menu_language': 'Language',
            'menu_theme': 'Theme',
            'theme_light': 'Light Mode',
            'theme_dark': 'Dark Mode',
            'lang_english': 'English',
            'lang_korean': '한국어',
            'select_user': 'Select User:',
            'main_operations': 'Main Operations',
            'tools_utilities': 'Tools & Utilities',
            'backup_data': 'Backup Data',
            'restore_data': 'Restore Data',
            'filters': 'Filters',
            'backup_drivers': 'Backup Drivers',
            'restore_drivers': 'Restore Drivers',
            'browser': 'Browser',
            'check_space': 'Check Space',
            'save_log': 'Save Log',
            'copy_data': 'Copy Data',
            'device_mgr': 'Device Mgr',
            'schedule': 'Schedule',
            'export_apps': 'Export Apps',
            'explorer': 'Explorer',
            'connect_nas': 'Connect NAS',
            'disconnect': 'Disconnect',
            'activity_log': 'Activity Log',
            'sound': 'Sound',
            'operation_finished': 'Operation finished',
            'backup_complete': 'Backup complete!',
            'ready': 'Ready',
            'select_user_begin': 'Select User to begin',
            'dialog_ok': 'OK',
            'dialog_yes': 'Yes',
            'dialog_no': 'No',
            'dialog_cancel': 'Cancel',
            'notice_title': 'Notice:',
            'notice_select_user': 'Select User to begin',
            'notice_filters': 'Filters:Exclude Rules:Name->ondrive*',
            'notice_hidden_system': 'Hidden and System attributes excluded',
            'notice_appdata': 'AppData folder is not backed up',
            'notice_detailed_log': 'Detailed log will be saved to a file during operations',
            'notice_responsibility': 'Use of this program is the sole responsibility of the user',

            # Dialog titles and messages

            # Dialog messages
            'language_changed': 'Language Changed',
            'language_changed_msg': 'Language changed successfully.\nRestart the application for full effect.',
            'theme_changed': 'Theme Changed',
            'theme_changed_msg': 'Theme switched to {0} mode.\nRestart the application for full effect.',

            'confirm_exit': 'Confirm Exit',
            'confirm_exit_msg': 'Are you sure you want to exit ezBAK?',
            'error': 'Error',
            'warning': 'Warning',
            'info': 'Info',
            'space_check': 'Space Check',

            'sufficient_space': '✓ Sufficient space available.',
            'insufficient_space': '⚠ Insufficient space available.',
            'no_log_content': 'No log content to save.',
            'error_saving_log': 'Error occurred while saving UI log: {0}',
            'select_user_warning': 'Please select a user.',
            'overwrite_warning': 'Existing data for \'{0}\' will be overwritten. Do you want to continue?',
            'task_creation_failed': 'Task Creation Failed',
            'task_deletion_failed': 'Task Deletion Failed',
            'schedule_created': 'Schedule Created',
            'task_scheduler': 'Task Scheduler',
            'task_deleted': 'Task \'{0}\' deleted.',
            'task_name_required': 'Task Name is required to delete.',
            'failed_to_delete_task': 'Failed to delete task:\n{0}',
            'source_selection_failed': 'Source selection failed: {0}',
            'filter_manager_error': 'Filter manager error: {0}',
            'unc_path_required': 'It should start with UNC Path \\\\ (ex: \\\\Server\\share)',
            'password_required': 'Password is required if User is provided.',
            'invalid_drive_letter': 'Please enter a valid drive letter (ex: Z)',
            'connected': 'Connected',
            'open_share_now': 'Open the share now?',
            'nas_connect': 'NAS Connect',
            'failed_to_connect': 'Failed to connect.\n\n{0}',
            'nas_connect_error': 'NAS connect error: {0}',
            'unc_path_must_start': 'UNC path must start with \\\\ (e.g., \\\\server\\share)',
            'map_drive': 'Map Drive',
            'map_to_drive_letter': 'Map this share to a drive letter?',
            'persistence': 'Persistence',
            'reconnect_at_signin': 'Reconnect at sign-in? (Persistent)',
            'nas_disconnect': 'NAS Disconnect',
            'disconnected': 'Disconnected: {0}',
            'failed_to_disconnect': 'Failed to disconnect.\n\n{0}',
            'nas_disconnect_error': 'NAS disconnect error: {0}',
            'input_target': 'Input the target (Drive Letter or UNC)',
            'invalid_drive_or_unc': 'Invalid drive letter or UNC path.',
            'system_mismatch_warning': 'System Mismatch Warning',
            'unable_to_open_devmgr': 'Unable to open Device Manager:\n{0}',
            'task_name_required_create': 'Task name is required.',
            'destination_required': 'Destination folder is required.',
            'invalid_time_format': 'Invalid time format. Use HH:MM (e.g., 14:30)',
            'failed_to_create_schedule': 'Failed to create schedule: {0}',
            'task_name_required_delete': 'Task name is required for deletion.',
            'failed_to_delete_schedule': 'Failed to delete schedule: {0}',
            'please_enter_pattern': 'Please enter a pattern.',
            'continue_anyway': 'Continue anyway?',

            # Dialog titles
            'select_sources': 'Select Sources',
            'select_destination_folder': 'Select Destination Folder',
            'check_required_space': 'Check Required Space',
            'schedule_backup': 'Schedule Backup',
            'network_share_connect': 'Network Share - Connect',
            'network_share_disconnect': 'Network Share - Disconnect',

            # NAS Connect dialog labels
            'unc_path_label': 'UNC Path (ex: \\\\NAS\\share)',
            'map_as_drive_letter': 'Map as drive letter',
            'letter': 'Letter:',
            'username_optional': 'Username (Optional)',
            'password_required_if_username': 'Password (Required if Username is provided)',
            'reconnect_at_logon_persistent': 'Reconnect at logon (Persistent)',

            # NAS Disconnect dialog labels
            'target_drive_or_unc': 'Target (Driver Letter ex: Z: or UNC ex: \\\\Server\\Share)',

            'space_sufficient': '✓ Sufficient space available.',
            'space_insufficient': '⚠ Insufficient space available.',
            'no_log_content': 'No log content to save.',
            'error_saving_log': 'Error occurred while saving UI log:',
            'please_select_user': 'Please select a user.',
            'data_will_be_overwritten': ' will be overwritten. Do you want to continue?',
            'task_creation_failed': 'Task Creation Failed',
            'task_deletion_failed': 'Task Deletion Failed',
            'schedule_created': 'Schedule Created',
            'task_name_required': 'Task Name is required to delete.',
            'failed_to_delete_task': 'Failed to delete task:',
            'task_scheduler': 'Task Scheduler',
            'task_deleted': 'deleted.',
            'source_selection_failed': 'Source selection failed:',
            'filter_manager_error': 'Filter manager error:',
            'unc_path_required': 'It should start with UNC Path \\\\ (ex: \\\\Server\\share)',
            'password_required': 'Password is required if User is provided.',
            'valid_drive_letter': 'Please enter a valid drive letter (ex: Z)',
            'connected': 'Connected',
            'open_share_now': 'Open the share now?',
            'nas_connect': 'NAS Connect',
            'failed_to_connect': 'Failed to connect.',
            'nas_connect_error': 'NAS connect error:',
            'unc_path_must_start': 'UNC path must start with \\\\ (e.g., \\\\server\\share)',
            'map_drive': 'Map Drive',
            'map_to_drive_letter': 'Map this share to a drive letter?',
            'invalid_drive_letter': 'Invalid drive letter.',
            'password': 'Password',
            'password_required_when_username': 'Password is required when specifying username.',
            'persistence': 'Persistence',
            'reconnect_at_signin': 'Reconnect at sign-in? (Persistent)',
            'nas_disconnect': 'NAS Disconnect',
            'disconnected': 'Disconnected:',
            'failed_to_disconnect': 'Failed to disconnect.',
            'nas_disconnect_error': 'NAS disconnect error:',
            'input_target': 'Input the target (Drive Letter or UNC)',
            'invalid_drive_or_unc': 'Invalid drive letter or UNC path.',
            'system_mismatch_warning': 'System Mismatch Warning',
            'unable_to_open_device_manager': 'Unable to open Device Manager:',

            # Filter Manager
            'filter_manager': 'Filter Manager',
            'include_rules': 'Include Rules',
            'exclude_rules': 'Exclude Rules',
            'add': 'Add',
            'remove': 'Remove',
            'clear_all': 'Clear All',
            'help': 'Help',
            'save': 'Save',
            'add_include_rule': 'Add Include Rule',
            'add_exclude_rule': 'Add Exclude Rule',
            'rule_type': 'Rule Type:',
            'pattern': 'Pattern:',
            'extension': 'Extension',
            'file_folder_name': 'Name',
            'full_path': 'Path',
            'extension_desc': 'File extensions (e.g., tmp, log, bak)',
            'name_desc': 'File/folder names with wildcards (e.g., Thumbs.db, *cache*)',
            'path_desc': 'Full paths with wildcards (e.g., */temp/*, *\\AppData\\*)',
            'add_rule': 'Add Rule',
            'examples_ext': 'Examples: tmp, log, bak, cache (without dots)',
            'examples_name': 'Examples: Thumbs.db, *.tmp, desktop.ini, *cache*',
            'examples_path': 'Examples: */Temp/*, *\\AppData\\Local\\*, */cache/*',
            'filter_info_include': 'Include: Files must match at least one rule to be included',
            'filter_info_exclude': 'Exclude: Files matching any rule will be skipped (takes priority)',

            # Filter Manager Help Dialog
            'filter_manager_help_title': 'Filter Manager Help',
            'filter_manager_help_content': """Filter Manager Help

KEYBOARD SHORTCUTS:
┌─────────────────────────────────────────────────────────────┐
│ Ctrl + N         Add Include Rule                           │
│ Ctrl + E         Add Exclude Rule                           │
│ Ctrl + S         Save Filters                               │
│ Ctrl + R         Clear Active List                          │
│ Delete           Remove Selected Item                       │
│ Tab              Switch Between Lists                       │
│ Esc              Cancel                                     │
│ Enter            Save                                       │
│ F1               Show This Help                             │
└─────────────────────────────────────────────────────────────┘

FILTER TYPES:
┌─────────────────────────────────────────────────────────────┐
│ 📄 Extension     File extensions (e.g., tmp, log, bak)      │
│ 📁 Name          File/folder names with wildcards           │
│                  Examples: Thumbs.db, *.tmp, *cache*        │
│ 🗂️ Path          Full paths with wildcards                  │
│                  Examples: */temp/*, *\\AppData\\*          │
└─────────────────────────────────────────────────────────────┘

HOW FILTERS WORK:
• Include Rules: Files must match at least one rule to be included
• Exclude Rules: Files matching any rule will be skipped
• Exclude rules take priority over include rules
• Use wildcards (*) for pattern matching
• Paths use forward slashes (/) or backslashes (\\)

EXAMPLES:
• Exclude temp files: [EXT] tmp
• Exclude cache folders: [NAME] *cache*
• Include only documents: [EXT] pdf, [EXT] docx
• Exclude system paths: [PATH] */Windows/System32/*""",
            'help_close_esc': 'Close (Esc)',

            # Keyboard Shortcuts Help Dialog
            'keyboard_shortcuts_title': 'Help - Keyboard Shortcuts',
            'keyboard_shortcuts_content': """ezBAK Keyboard Shortcuts

┌─────────────────────────────────────────┐
│           Main Features                 │
├─────────────────────────────────────────┤
│ Ctrl + B        Backup Data             │
│ Ctrl + R        Restore Data            │
│ Ctrl + D        Backup Drivers          │
│ Ctrl + Shift + D Restore Drivers        │
├─────────────────────────────────────────┤
│           File management               │
├─────────────────────────────────────────┤
│ Ctrl + C        Copy Data               │
│ Ctrl + O        Explorer                │
│ Ctrl + S        Save Log                │
├─────────────────────────────────────────┤
│           Tools and Utilities           │
├─────────────────────────────────────────┤
│ Ctrl + M        Device Mgr              │
│ Ctrl + K        Check Space             │
│ Ctrl + F        Filters                 │
│ Ctrl + T        Schedule                │
├─────────────────────────────────────────┤
│           Browser and App               │
├─────────────────────────────────────────┤
│ Ctrl + P        Browser                 │
│ Ctrl + W        Export Apps             │
├─────────────────────────────────────────┤
│           Network                       │
├─────────────────────────────────────────┤
│ Ctrl + N        Connect NAS             │
│ Ctrl + Shift + N Disconnect             │
├─────────────────────────────────────────┤
│           General                       │
├─────────────────────────────────────────┤
│ F1              help                    │
│ F5              (Not Working) Refresh User List       │
│ Ctrl + 1~9      (Not Working) Quick Select User │
│ Escape          (Not Working) Cancel Current Operation│
│ Ctrl + Q        Exit Program            │
└─────────────────────────────────────────┘

Tips :  Shortcuts are not case-sensitive.""",
            'help_exit_esc': 'Exit (Esc)',
        },
        'ko': {
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
│ Ctrl + 1~9      (작동 안 함) 사용자 빠른 선택      │
│ Escape          (작동 안 함) 현재 작업 취소        │
│ Ctrl + Q        프로그램 종료            │
└─────────────────────────────────────────┘

팁 : 단축키는 대소문자를 구분하지 않습니다.""",
            'help_exit_esc': '종료 (Esc)',
        }
    }

    def __init__(self, initial_lang='en'):
        self.current_lang = initial_lang

    def set_language(self, lang_code):
        """Set current language"""
        if lang_code in self.TRANSLATIONS:
            self.current_lang = lang_code
            return True
        return False

    def get(self, key, default=None):
        """Get translated text"""
        trans = self.TRANSLATIONS.get(self.current_lang, {})
        return trans.get(key, default or key)


# ============================================================
# Toggle Switch Widget
# ============================================================

class ToggleSwitch(tk.Canvas):
    """Modern toggle switch widget with Windows 11 style"""

    def __init__(self, parent, width=50, height=24, variable=None, command=None, theme=None, **kwargs):
        """
        Initialize toggle switch

        Args:
            parent: Parent widget
            width: Switch width (default 50)
            height: Switch height (default 24)
            variable: BooleanVar to bind to
            command: Callback function when toggled
            theme: Win11Theme instance for colors
        """
        super().__init__(parent, width=width, height=height,
                        bg=theme.get('bg_elevated') if theme else '#f5f5f5',
                        highlightthickness=0, **kwargs)

        self.width = width
        self.height = height
        self.variable = variable
        self.command = command
        self.theme = theme

        # Track initial state
        self.is_on = variable.get() if variable else False

        # Colors
        self.update_colors()

        # Draw switch
        self.draw_switch()

        # Bind click event
        self.bind("<Button-1>", self.toggle)

        # Bind variable trace if provided
        if self.variable:
            self.variable.trace_add('write', self._on_variable_change)

    def update_colors(self):
        """Update colors based on theme"""
        if self.theme:
            if self.theme.current_mode == 'dark':
                self.color_on = self.theme.get('accent')  # Blue when on
                self.color_off = '#484f58'  # Gray when off
                self.knob_color = '#ffffff'  # White knob
            else:
                self.color_on = self.theme.get('accent')  # Blue when on
                self.color_off = '#c8c6c4'  # Gray when off
                self.knob_color = '#ffffff'  # White knob
        else:
            self.color_on = '#0078D4'
            self.color_off = '#c8c6c4'
            self.knob_color = '#ffffff'

    def draw_switch(self):
        """Draw the toggle switch"""
        self.delete("all")

        # Calculate positions
        radius = self.height / 2
        knob_radius = radius - 3

        # Draw background track
        bg_color = self.color_on if self.is_on else self.color_off

        # Draw rounded rectangle for track
        self.create_oval(0, 0, self.height, self.height,
                        fill=bg_color, outline="")
        self.create_oval(self.width - self.height, 0,
                        self.width, self.height,
                        fill=bg_color, outline="")
        self.create_rectangle(radius, 0, self.width - radius, self.height,
                             fill=bg_color, outline="")

        # Draw knob
        if self.is_on:
            knob_x = self.width - radius
        else:
            knob_x = radius

        self.create_oval(knob_x - knob_radius, radius - knob_radius,
                        knob_x + knob_radius, radius + knob_radius,
                        fill=self.knob_color, outline="")

    def toggle(self, event=None):
        """Toggle the switch"""
        self.is_on = not self.is_on

        # Update variable if provided
        if self.variable:
            self.variable.set(self.is_on)

        # Redraw
        self.draw_switch()

        # Call command if provided
        if self.command:
            self.command()

    def _on_variable_change(self, *args):
        """Called when the bound variable changes"""
        if self.variable:
            new_state = self.variable.get()
            if new_state != self.is_on:
                self.is_on = new_state
                self.draw_switch()

    def update_theme(self, theme):
        """Update theme colors"""
        self.theme = theme
        self.update_colors()
        self.config(bg=theme.get('bg_elevated'))
        self.draw_switch()


# ============================================================
# Windows 11 Style Dialog System
# ============================================================

class Win11Dialog:
    """Windows 11 style modern dialogs to replace default messageboxes"""

    @staticmethod
    def _create_dialog_button(parent, text, command, theme, is_primary=False):
        """Create a Windows 11 style dialog button with modern design"""
        # Windows 11 style: Primary buttons have accent color, secondary have subtle background
        if is_primary:
            btn_bg = theme.get('accent')
            btn_fg = '#FFFFFF' if theme.current_mode == 'dark' else '#FFFFFF'
            hover_bg = theme.get('accent_hover')
            hover_fg = '#FFFFFF'
            pressed_bg = theme.get('accent_hover')
        else:
            btn_bg = theme.get('bg_elevated')
            btn_fg = theme.get('fg')
            hover_bg = theme.get('hover')
            hover_fg = theme.get('fg')
            pressed_bg = theme.get('pressed')

        border_color = theme.get('border')

        btn = tk.Button(parent, text=text,
                       bg=btn_bg,
                       fg=btn_fg,
                       font=("Segoe UI", 9),
                       relief="flat",
                       bd=0,
                       highlightthickness=1,
                       highlightbackground=border_color,
                       highlightcolor=border_color,
                       borderwidth=1,
                       width=10,
                       height=1,
                       command=command,
                       cursor="hand2",
                       padx=12,
                       pady=6)

        # Windows 11 style hover and press effects
        def on_enter(e):
            e.widget.config(bg=hover_bg, fg=hover_fg, relief="flat")

        def on_leave(e):
            e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

        def on_press(e):
            e.widget.config(bg=pressed_bg, relief="flat")

        def on_release(e):
            e.widget.config(bg=hover_bg, relief="flat")

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)

        return btn

    @staticmethod
    def showinfo(title, message, parent=None, theme=None, translator=None, show_icon=True):
        """Show information dialog"""
        if theme is None:
            theme = Win11Theme('dark')
        if translator is None:
            translator = Translator('en')

        dlg = tk.Toplevel(parent)
        dlg.title(title)
        # Set titlebar icon
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg=theme.get('bg_elevated'))
        dlg.resizable(False, False)
        dlg.transient(parent)
        dlg.grab_set()

        # Center the dialog
        dlg.update_idletasks()
        width = 450
        height = 350  # Increased to accommodate longer messages like schedule creation confirmation
        x = (dlg.winfo_screenwidth() // 2) - (width // 2)
        y = (dlg.winfo_screenheight() // 2) - (height // 2)
        dlg.geometry(f'{width}x{height}+{x}+{y}')

        # Message area
        msg_frame = tk.Frame(dlg, bg=theme.get('bg_elevated'))
        msg_frame.pack(fill='both', expand=True, padx=24, pady=20)

        # Icon and message in horizontal layout
        content_frame = tk.Frame(msg_frame, bg=theme.get('bg_elevated'))
        content_frame.pack(fill='both', expand=True)

        # Add info icon (optional)
        if show_icon:
            icon_label = tk.Label(content_frame, text="ℹ️",
                                 bg=theme.get('bg_elevated'),
                                 fg=theme.get('accent'),
                                 font=("Segoe UI", 24))
            icon_label.pack(side='left', padx=(0, 12))

        tk.Label(content_frame, text=message,
                bg=theme.get('bg_elevated'),
                fg=theme.get('fg'),
                font=("Segoe UI", 10),
                wraplength=360,
                justify='left').pack(side='left', fill='both', expand=True)

        # Button area
        btn_frame = tk.Frame(dlg, bg=theme.get('bg_elevated'))
        btn_frame.pack(fill='x', padx=24, pady=(0, 20))

        Win11Dialog._create_dialog_button(btn_frame, translator.get('dialog_ok'), dlg.destroy, theme, is_primary=True).pack(side='right')

        dlg.wait_window()

    @staticmethod
    def showwarning(title, message, parent=None, theme=None, translator=None):
        """Show warning dialog"""
        Win11Dialog.showinfo(title, message, parent, theme, translator)

    @staticmethod
    def showerror(title, message, parent=None, theme=None, translator=None):
        """Show error dialog"""
        Win11Dialog.showinfo(title, message, parent, theme, translator)

    @staticmethod
    def askyesno(title, message, parent=None, theme=None, translator=None):
        """Show yes/no dialog"""
        if theme is None:
            theme = Win11Theme('dark')
        if translator is None:
            translator = Translator('en')

        result = [False]

        dlg = tk.Toplevel(parent)
        dlg.title(title)
        # Set titlebar icon
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg=theme.get('bg_elevated'))
        dlg.resizable(False, False)
        dlg.transient(parent)
        dlg.grab_set()

        # Center the dialog
        dlg.update_idletasks()
        width = 400
        height = 200
        x = (dlg.winfo_screenwidth() // 2) - (width // 2)
        y = (dlg.winfo_screenheight() // 2) - (height // 2)
        dlg.geometry(f'{width}x{height}+{x}+{y}')

        # Message area
        msg_frame = tk.Frame(dlg, bg=theme.get('bg_elevated'))
        msg_frame.pack(fill='both', expand=True, padx=24, pady=20)

        tk.Label(msg_frame, text=message,
                bg=theme.get('bg_elevated'),
                fg=theme.get('fg'),
                font=("Segoe UI", 10),
                wraplength=350,
                justify='left').pack()

        # Button area
        btn_frame = tk.Frame(dlg, bg=theme.get('bg_elevated'))
        btn_frame.pack(fill='x', padx=24, pady=(0, 20))

        def on_yes():
            result[0] = True
            dlg.destroy()

        def on_no():
            result[0] = False
            dlg.destroy()

        Win11Dialog._create_dialog_button(btn_frame, translator.get('dialog_no'), on_no, theme, is_primary=False).pack(side='right', padx=(8, 0))
        Win11Dialog._create_dialog_button(btn_frame, translator.get('dialog_yes'), on_yes, theme, is_primary=True).pack(side='right')

        dlg.wait_window()
        return result[0]


        #     # always run finalization steps
        # finally:
        #     try:
        #         self.close_log_file()
        #     except Exception:
        #         pass
            
        #     try:
        #         # Set progress to 100%
        #         max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
        #         self.message_queue.put(('update_progress', max_val))
        #     except Exception:
        #         pass

        #     # Enable buttons
        #     self.message_queue.put(('enable_buttons', None))

        #     # Final status message
        #     if hasattr(self, '_backup_completed') and self._backup_completed:
        #         self.message_queue.put(('update_status', "Backup complete!"))
        #     else:
        #         self.message_queue.put(('update_status', "Operation finished"))


try:
    import win32api
    import win32con
    _have_pywin32 = True
except Exception:
    win32api = None
    win32con = None
    _have_pywin32 = False

    # Windows file attribute constants (from WinBase.h)
    FILE_ATTRIBUTE_READONLY = 0x01
    FILE_ATTRIBUTE_HIDDEN = 0x02
    FILE_ATTRIBUTE_SYSTEM = 0x04
    FILE_ATTRIBUTE_DIRECTORY = 0x10
    FILE_ATTRIBUTE_ARCHIVE = 0x20
    FILE_ATTRIBUTE_DEVICE = 0x40
    FILE_ATTRIBUTE_NORMAL = 0x80
    FILE_ATTRIBUTE_REPARSE_POINT = 0x400

    # ctypes wrapper for GetFileAttributesW
    try:
        from ctypes import wintypes

        _GetFileAttributesW = ctypes.windll.kernel32.GetFileAttributesW
        _GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
        _GetFileAttributesW.restype = wintypes.DWORD

        def _get_file_attributes(path):
            # Ensure path is a native string (LPCWSTR)
            try:
                res = _GetFileAttributesW(path)
            except Exception:
                # On some rare environments, kernel32 call may raise; convert to OSError behavior
                raise
            if res == 0xFFFFFFFF:
                # INVALID_FILE_ATTRIBUTES
                raise OSError(f"GetFileAttributesW failed for {path}")
            return int(res)
    except Exception:
        # If ctypes fallback fails for any reason, provide a dummy that raises
        def _get_file_attributes(path):
            raise OSError("GetFileAttributesW unavailable")

def format_bytes(size):
    """Converts bytes to a human-readable format (KB, MB, GB, etc.)."""
    try:
        size = float(size)
    except (ValueError, TypeError):
        return "N/A"
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024.0:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} PB"

def log_message(message, log_folder, prefix=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(log_folder, exist_ok=True)

    log_filename = os.path.join(
        log_folder,
        f"{prefix}_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.log"
    )

    for line in str(message).splitlines():
        log_line = f"[{timestamp}] {line}"
        # print(log_line)  
        try:
            with open(log_filename, "a", encoding="utf-8", errors="ignore") as f:
                f.write(log_line + "\n")
        except Exception:
            pass

def get_system_info():
    """
    Get baseboard model name or product name via WMI
    priority: Baseboard Product → Computer System Model  Computer System Manufacturer → Fallback "Unknown"
    """
    import locale

    def clean_name(name):
        """ Remove and clean up unusable characters in file names"""
        if not name or name.lower() in ('to be filled by o.e.m.', 'system product name', 'system version'):
            return None
        # Remove special characters and change spaces to underscores
        cleaned = re.sub(r'[<>:"/\\|?*]', '', name)
        cleaned = re.sub(r'\s+', '_', cleaned.strip())
        return cleaned if cleaned else None

    system_name = "Unknown"
    # Use system default encoding (e.g., cp949 on Korean Windows)
    system_encoding = locale.getpreferredencoding()

    # 1. Try to read from Windows Registry first
    try:
        import winreg
        reg_path = r"HARDWARE\DESCRIPTION\System\BIOS"
        reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
        
        try:
            baseboard_product, _ = winreg.QueryValueEx(reg_key, "BaseBoardProduct")
            winreg.CloseKey(reg_key)
            
            clean_product = clean_name(baseboard_product)
            if clean_product:
                system_name = clean_product
                print(f"DEBUG: Using registry BaseBoardProduct: {system_name}")
                return system_name
        except FileNotFoundError:
            winreg.CloseKey(reg_key)
            print(f"DEBUG: BaseBoardProduct not found in registry")
    except Exception as e:
        print(f"DEBUG: Registry query failed: {e}")
    
    try:
        # 2. Try WMI Baseboard product name
        cmd = ['wmic', 'baseboard', 'get', 'product', '/value']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding=system_encoding, errors='replace')
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if line.startswith('Product='):
                    product = line.split('=', 1)[1].strip()
                    clean_product = clean_name(product)
                    if clean_product:
                        system_name = clean_product
                        print(f"DEBUG: Using baseboard product: {system_name}")
                        return system_name
    except Exception as e:
        print(f"DEBUG: Baseboard product query failed: {e}")

    try:
        # 3. Try WMI Computer system model name
        cmd = ['wmic', 'computersystem', 'get', 'model', '/value']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding=system_encoding, errors='replace')
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if line.startswith('Model='):
                    model = line.split('=', 1)[1].strip()
                    clean_model = clean_name(model)
                    if clean_model:
                        system_name = clean_model
                        print(f"DEBUG: Using computer system model: {system_name}")
                        return system_name
    except Exception as e:
        print(f"DEBUG: Computer system model query failed: {e}")

    try:
        # 4. Try WMI Manufacturer's name (last resort)
        cmd = ['wmic', 'computersystem', 'get', 'manufacturer', '/value']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding=system_encoding, errors='replace')
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if line.startswith('Manufacturer='):
                    manufacturer = line.split('=', 1)[1].strip()
                    clean_manufacturer = clean_name(manufacturer)
                    if clean_manufacturer:
                        system_name = clean_manufacturer
                        print(f"DEBUG: Using manufacturer: {system_name}")
                        return system_name
    except Exception as e:
        print(f"DEBUG: Manufacturer query failed: {e}")

    print(f"DEBUG: Using fallback name: {system_name}")
    return system_name

def _handle_rmtree_error(func, path, exc_info):
    """
    Error handler for shutil.rmtree that handles read-only files.
    """
    if isinstance(exc_info[1], PermissionError):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception as e:
            log_message(f"ERROR: Could not delete {path} even after removing read-only flag: {e}")
    else:
        log_message(f"ERROR: Deletion failed for {path}: {exc_info[1]}")

def _handle_rmtree_error_headless(func, path, exc_info, log, timestamp):
    """
    Error handler for shutil.rmtree (for headless mode - logs to file)
    """
    
    if isinstance(exc_info[1], PermissionError):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
            log.write(f"[{timestamp}] INFO: Removed read-only and successfully deleted {path}\n")
        except Exception as e:
            log.write(f"[{timestamp}] ERROR: Could not delete {path} even after removing read-only flag: {e}\n")
    else:
        log.write(f"[{timestamp}] ERROR: Deletion failed for {path}: {exc_info[1]}\n")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def show_notification(title, message, app_id="ezBAK", button_text=None, button_argument=None):
    """Shows a Windows toast notification with an optional button."""
    try:
        # Escape single quotes for PowerShell
        safe_title = title.replace("'", "''")
        safe_message = message.replace("'", "''")

        if button_text and button_argument:
            safe_button_text = button_text.replace("'", "''")
            safe_button_argument = button_argument.replace("'", "''")
            
            script = f"""
            [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
            $template = @"
            <toast>
                <visual>
                    <binding template="ToastGeneric">
                        <text>{safe_title}</text>
                        <text>{safe_message}</text>
                    </binding>
                </visual>
                <actions>
                    <action
                        content="{safe_button_text}"
                        arguments="{safe_button_argument}"
                        activationType="protocol"/>
                </actions>
            </toast>
"@
            $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xml.LoadXml($template)
            $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('{app_id}').Show($toast)
            """
        else:
            script = f"""
            [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
            $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
            $toastXml = $template.GetXml()
            $toastXml.GetElementsByTagName('text')[0].AppendChild($template.CreateTextNode('{safe_title}')) | Out-Null
            $toastXml.GetElementsByTagName('text')[1].AppendChild($template.CreateTextNode('{safe_message}')) | Out-Null
            $toastNotification = [Windows.UI.Notifications.ToastNotification]::new($toastXml)
            [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('{app_id}').Show($toastNotification)
            """
        subprocess.run(["powershell", "-Command", script], check=False, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        print("INFO: Failed to show Windows notification.")

class HeadlessBackupRunner:
    """Performs backup operations in headless mode based on command-line arguments."""
    def __init__(self, user, dest, include_hidden, include_system, retention_count, log_retention_days):
        self.user = user
        self.dest = dest
        self.include_hidden = include_hidden
        self.include_system = include_system
        self.retention_count = retention_count
        self.log_retention_days = log_retention_days

    def run_backup(self):
        """Running headless mode backups"""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            # timestamp = datetime.now().strftime("%Y-%m-%d")
            backup_path, log_path = core_run_backup(
                self.user,
                self.dest,
                include_hidden=self.include_hidden,
                include_system=self.include_system,
                log_folder=self.dest,
                retention_count=self.retention_count,
                log_retention_days=self.log_retention_days
            )
            print(f"[{timestamp}] Backup finished → {backup_path}")
            print(f"[{timestamp}] Log saved at {log_path}")
            return True
        except Exception as e:
            print(f"[{timestamp}] ERROR during headless backup: {e}")
            return False
        finally:
            try:
                self.close_log_file()
            except Exception:
                pass
            
            try:
                # Set progress to 100%
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass
            
            # Button Activation
            self.message_queue.put(('enable_buttons', None))
            
            # final status message
            if hasattr(self, '_backup_completed') and self._backup_completed:
                self.message_queue.put(('update_status', "Backup complete!"))
            else:
                self.message_queue.put(('update_status', "Operation finished"))

class DialogShortcuts:
    """Common shortcut key functions for dialogs"""
    
    def setup_dialog_shortcuts(self):
        """Common Shortcut Key Features for Dialogs"""
        # Basic dialog shortcuts
        self.bind('<Escape>', lambda e: self.on_cancel() if hasattr(self, 'on_cancel') else self.destroy())
        self.bind('<Return>', lambda e: self.on_ok() if hasattr(self, 'on_ok') else None)
        self.bind('<Alt-F4>', lambda e: self.destroy())
        
        # Set focus (to receive keyboard events)
        self.focus_set()
        
    def add_shortcut_text(self, text, shortcut):
        """Add shortcut notation to text"""
        return f"{text} ({shortcut})"

class KeyboardShortcuts:
    """Keyboard shortcuts manager class for ezBAK"""
    
    def __init__(self, app):
        self.app = app
        self.setup_shortcuts()
        self.setup_help_dialog()
    
    def setup_shortcuts(self):
        """Register keyboard shortcuts"""
        
        # Main feature shortcuts
        self.app.bind('<Control-b>', lambda e: self.safe_execute(self.app.start_backup_thread))
        self.app.bind('<Control-r>', lambda e: self.safe_execute(self.app.start_restore_thread))
        self.app.bind('<Control-d>', lambda e: self.safe_execute(self.app.start_driver_backup_thread))
        self.app.bind('<Control-Shift-D>', lambda e: self.safe_execute(self.app.start_driver_restore_thread))
        
        # File management
        self.app.bind('<Control-c>', lambda e: self.safe_execute(self.app.copy_files))
        self.app.bind('<Control-o>', lambda e: self.safe_execute(self.app.open_file_explorer))
        self.app.bind('<Control-s>', lambda e: self.safe_execute(self.app.save_log))
        
        # Tools and utilities
        self.app.bind('<Control-m>', lambda e: self.safe_execute(self.app.open_device_manager))
        self.app.bind('<Control-k>', lambda e: self.safe_execute(self.app.check_space))
        self.app.bind('<Control-f>', lambda e: self.safe_execute(self.app.open_filter_manager))
        self.app.bind('<Control-t>', lambda e: self.safe_execute(self.app.schedule_backup))
        
        # Browser and app related
        self.app.bind('<Control-p>', lambda e: self.safe_execute(self.app.start_browser_profiles_backup_thread))
        self.app.bind('<Control-w>', lambda e: self.safe_execute(self.app.start_winget_export_thread))
        
        # NAS connection
        self.app.bind('<Control-n>', lambda e: self.safe_execute(self.app.open_nas_connect_dialog))
        self.app.bind('<Control-Shift-N>', lambda e: self.safe_execute(self.app.open_nas_disconnect_dialog))
        
        # General shortcuts
        self.app.bind('<F1>', lambda e: self.show_help())
        self.app.bind('<F5>', lambda e: self.refresh_user_list())
        self.app.bind('<Escape>', lambda e: self.cancel_operation())
        self.app.bind('<Control-q>', lambda e: self.safe_exit())
        
        # User selection shortcuts (number keys)
        for i in range(1, 10):
            self.app.bind(f'<Control-{i}>', lambda e, idx=i-1: self.select_user_by_index(idx))
        
        # Focus shortcuts
        self.app.focus_set()  
    
    def safe_execute(self, function):
        """Execute function (with error handling)"""
        try:
            if callable(function):
                function()
            else:
                self.app.message_queue.put(('log', f"Function {function} is not callable"))
        except Exception as e:
            self.app.message_queue.put(('log', f"Shortcut execution error: {e}"))
            print(f"DEBUG: Shortcut error: {e}")
    
    def show_help(self):
        """Show shortcuts help (F1)"""
        help_text = self.app.translator.get('keyboard_shortcuts_content')

        # Custom help dialog
        help_dialog = tk.Toplevel(self.app)
        help_dialog.title(self.app.translator.get('keyboard_shortcuts_title'))
        help_dialog.configure(bg=self.app.theme.get('bg'))
        help_dialog.geometry("449x748")
        help_dialog.transient(self.app)
        help_dialog.grab_set()

        try:
            help_dialog.iconbitmap(resource_path('./icon/ezbak.ico'))
        except:
            pass

        # Text area
        text_frame = tk.Frame(help_dialog, bg=self.app.theme.get('bg'))
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create the scrollbar first
        scrollbar = tk.Scrollbar(text_frame,
                                width=20,
                                bg=self.app.theme.get('scrollbar_bg'),
                                troughcolor=self.app.theme.get('bg'),
                                activebackground=self.app.theme.get('accent'),
                                relief="flat")
        scrollbar.pack(side="right", fill="y")

        text_widget = tk.Text(text_frame,
                             font=("Consolas", 10),
                             bg=self.app.theme.get('text_bg'),
                             fg=self.app.theme.get('fg'),
                             relief="flat",
                             wrap="word",
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)

        # Connect the scrollbar and the text widget
        scrollbar.configure(command=text_widget.yview)

        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        # Close button
        btn_frame = tk.Frame(help_dialog, bg=self.app.theme.get('bg'))
        btn_frame.pack(pady=10)
        close_btn = Win11Dialog._create_dialog_button(btn_frame,
                                                      self.app.translator.get('help_exit_esc'),
                                                      help_dialog.destroy,
                                                      self.app.theme,
                                                      is_primary=True)
        close_btn.pack()

        # Close with ESC
        help_dialog.bind('<Escape>', lambda e: help_dialog.destroy())
        help_dialog.focus_set()
    
    def refresh_user_list(self):
        """Refresh user list (F5)"""
        try:
            self.app.load_users()
            self.app.message_queue.put(('log', "Refreshed user list."))
        except Exception as e:
            self.app.message_queue.put(('log', f"Error refreshing user list: {e}"))

    def select_user_by_index(self, index):
        """Quick select user by number key (Ctrl + 1~9)"""
        try:
            users = self.app.user_combo['values']
            if 0 <= index < len(users):
                self.app.user_combo.current(index)
                self.app.message_queue.put(('log', f"Selected user: {users[index]}"))
            else:
                self.app.message_queue.put(('log', f"User index {index+1} is out of range."))
        except Exception as e:
            self.app.message_queue.put(('log', f"Error selecting user: {e}"))

    def cancel_operation(self):
        """Cancel current operation (ESC)"""
        try:
            # Attempt to cancel if an operation is in progress
            if hasattr(self.app, 'progress_bar') and self.app.progress_bar['mode'] == 'indeterminate':
                self.app.message_queue.put(('log', "Cancellation requested"))
                # The actual cancellation logic needs to be implemented in each task thread
        except Exception as e:
            print(f"DEBUG: Cancel operation error: {e}")
    
    def safe_exit(self):
        """Exit program  (Ctrl + Q)"""
        try:
            if Win11Dialog.askyesno(self.app.translator.get('confirm_exit'),
                                   self.app.translator.get('confirm_exit_msg'),
                                   parent=self.app, theme=self.app.theme, translator=self.app.translator):
                try:
                    self.app.save_settings()
                    self.app.close_log_file()
                except:
                    pass
                self.app.quit()
                self.app.destroy()
        except Exception as e:
            print(f"DEBUG: Safe exit error: {e}")
    
    def setup_help_dialog(self):
        """Display shortcuts hint in status bar or menu"""
        try:
            if hasattr(self.app, 'status_label'):
                original_text = self.app.status_label.cget("text")
                hint_text = f"{original_text}"
                self.app.status_label.configure(text=hint_text)
        except:
            pass

def setup_keyboard_shortcuts(self):
    self.shortcuts = KeyboardShortcuts(self)
    
    self.create_tooltip(self.backup_btn, "Backup User Data (Ctrl+B)")
    self.create_tooltip(self.restore_btn, "Restore User Data (Ctrl+R)")
    self.create_tooltip(self.driver_backup_btn, "Backup Drivers (Ctrl+D)")

def create_tooltip(self, widget, text):
    """Add tooltip to button (including shortcut info)"""
    def on_enter(event):
        try:
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.configure(bg="#2b2b2b", relief="solid", borderwidth=1)

            label = tk.Label(tooltip, text=text,
                           bg="#2b2b2b", fg="#ffffff",
                           font=("Segoe UI", 9))
            label.pack()
            
            x = event.x_root + 10
            y = event.y_root - 30
            tooltip.geometry(f"+{x}+{y}")
            
            widget.tooltip = tooltip
        except:
            pass
    
    def on_leave(event):
        try:
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                delattr(widget, 'tooltip')
        except:
            pass
    
    widget.bind('<Enter>', on_enter)
    widget.bind('<Leave>', on_leave)

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Hide window until fully initialized
        self.withdraw()

        # Initialize theme and translation systems
        self.theme = Win11Theme(initial_mode='dark')  # Start with dark mode
        self.translator = Translator(initial_lang='en')  # Start with English

        self.log_retention_days_var = tk.StringVar(value='30')
        self.title(self.translator.get('app_title'))
        try:
            # Try the title icon first (bundled with PyInstaller)
            self.iconbitmap(resource_path('./icon/ezbak_title.ico'))
        except tk.TclError:
            try:
                # Fallback to the regular icon
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except tk.TclError:
                pass
        self.geometry("1026x790")  # Slightly taller for menu bar
        self.configure(bg=self.theme.get('bg'))

        # UI elements setup
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Configure professional progress bar styles
        self._configure_progressbar_style()

        # Queue for thread-safe UI updates
        self.message_queue = queue.Queue()
        self.process_queue()

        # Options: control inclusion for Hidden and System attributes (defaults: exclude)
        self.hidden_mode_var = tk.StringVar(value='exclude')   # 'include' or 'exclude'
        self.system_mode_var = tk.StringVar(value='exclude')   # 'include' or 'exclude'
        self.retention_count_var = tk.StringVar(value='2') # Number of backups to keep
        self.log_retention_days_var = tk.StringVar(value='30') # Number of days to keep logs
        # Pattern filters (persisted): {'include': [{'type': 'ext'|'name'|'path', 'pattern': str}], 'exclude': [...]}
        
        # Sound notifications enabled
        self.sound_enabled_var = tk.BooleanVar(value=True)
        # self.filters = {'include': [], 'exclude': []}
        self.filters = {'include': [], 'exclude': [{'type': 'NAME', 'pattern': 'onedrive*'}]}
        # Load persisted settings (if any)
        try:
            self.load_settings()
        except Exception:
            pass

        # logging to file
        self.log_file_path = None
        self._log_file = None
        self.log_lock = threading.Lock()

        # progress/log throttle to keep UI responsive
        self._last_progress_ts = 0.0
        self.progress_throttle_secs = 0.12
        self._last_ui_log_ts = 0.0
        self.ui_log_throttle_secs = 0.08

        self.create_widgets()

        # Update UI texts with saved language settings
        self.update_ui_texts()

        # Initial message for the log box (UI short notice) with only the final sentence bolded
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.config(state="normal")
        # Insert standard notice text with translations
        notice_text = f"[{timestamp}] \n{self.translator.get('notice_title')}\n"
        notice_text += f"- {self.translator.get('notice_select_user')}\n"
        notice_text += f"- {self.translator.get('notice_filters')}\n"
        notice_text += f"- {self.translator.get('notice_hidden_system')}\n"
        notice_text += f"- {self.translator.get('notice_appdata')}\n"
        notice_text += f"- {self.translator.get('notice_detailed_log')}\n- "
        self.log_text.insert(tk.END, notice_text)
        # Insert the single bolded disclaimer
        self.log_text.insert(tk.END, self.translator.get('notice_responsibility') + "\n", 'bold')
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.update_idletasks()
        self.setup_keyboard_shortcuts()
        # Show window after all initialization is complete
        self.deiconify()

    def setup_keyboard_shortcuts(self):
        """Setting keyboard shortcuts"""
        self.shortcuts = KeyboardShortcuts(self)

    def create_menu_bar(self):
        """Create Windows 11 style menu bar"""
        menubar = tk.Menu(self,
                         bg=self.theme.get('bg_elevated'),
                         fg=self.theme.get('fg'),
                         activebackground=self.theme.get('accent'),
                         activeforeground='white',
                         bd=0,
                         relief='flat')

        # View Menu
        view_menu = tk.Menu(menubar, tearoff=0,
                           bg=self.theme.get('bg_elevated'),
                           fg=self.theme.get('fg'),
                           activebackground=self.theme.get('accent'),
                           activeforeground='white',
                           bd=0,
                           relief='flat')

        # Theme submenu
        theme_menu = tk.Menu(view_menu, tearoff=0,
                            bg=self.theme.get('bg_elevated'),
                            fg=self.theme.get('fg'),
                            activebackground=self.theme.get('accent'),
                            activeforeground='white')
        theme_menu.add_command(label=self.translator.get('theme_light'),
                              command=lambda: self.switch_theme('light'))
        theme_menu.add_command(label=self.translator.get('theme_dark'),
                              command=lambda: self.switch_theme('dark'))

        view_menu.add_cascade(label=self.translator.get('menu_theme'), menu=theme_menu)
        menubar.add_cascade(label=self.translator.get('menu_view'), menu=view_menu)

        # Language Menu
        lang_menu = tk.Menu(menubar, tearoff=0,
                           bg=self.theme.get('bg_elevated'),
                           fg=self.theme.get('fg'),
                           activebackground=self.theme.get('accent'),
                           activeforeground='white',
                           bd=0,
                           relief='flat')
        lang_menu.add_command(label="English", command=lambda: self.switch_language('en'))
        lang_menu.add_command(label="한국어", command=lambda: self.switch_language('ko'))

        menubar.add_cascade(label=self.translator.get('menu_language'), menu=lang_menu)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0,
                           bg=self.theme.get('bg_elevated'),
                           fg=self.theme.get('fg'),
                           activebackground=self.theme.get('accent'),
                           activeforeground='white',
                           bd=0,
                           relief='flat')
        help_menu.add_command(label="Keyboard Shortcuts (F1)", command=self.show_keyboard_shortcuts_help)
        help_menu.add_separator()
        help_menu.add_command(label="GitHub Releases", command=lambda: self.open_github_link(None))

        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def switch_theme(self, mode):
        """Switch between light and dark themes"""
        if mode != self.theme.current_mode:
            self.theme.current_mode = mode
            self.theme.colors = self.theme.DARK if mode == 'dark' else self.theme.LIGHT
            self.apply_theme_to_all_widgets()
            mode_text = self.translator.get('theme_dark') if mode == 'dark' else self.translator.get('theme_light')
            Win11Dialog.showinfo(self.translator.get('theme_changed'),
                              self.translator.get('theme_changed_msg').format(mode_text),
                              parent=self, theme=self.theme, translator=self.translator, show_icon=False)

    def switch_language(self, lang_code):
        """Switch application language"""
        if self.translator.set_language(lang_code):
            self.update_ui_texts()
            try:
                self.save_settings()  # Save the language preference
            except Exception:
                pass
            Win11Dialog.showinfo(self.translator.get('language_changed'),
                              self.translator.get('language_changed_msg'),
                              parent=self, theme=self.theme, translator=self.translator, show_icon=False)

    def _configure_progressbar_style(self):
        """Configure professional progress bar style based on current theme"""
        if self.theme.current_mode == 'dark':
            # Dark theme: vibrant accent with subtle glow effect
            self.style.configure("green.Horizontal.TProgressbar",
                               troughcolor=self.theme.get('bg_elevated'),
                               background=self.theme.get('accent'),
                               bordercolor=self.theme.get('border'),
                               lightcolor=self.theme.get('accent_hover'),
                               darkcolor=self.theme.get('accent'),
                               borderwidth=0,
                               thickness=20)
        else:
            # Light theme: clean modern look
            self.style.configure("green.Horizontal.TProgressbar",
                               troughcolor=self.theme.get('bg_elevated'),
                               background=self.theme.get('accent'),
                               bordercolor=self.theme.get('border'),
                               lightcolor=self.theme.get('accent_light'),
                               darkcolor=self.theme.get('accent'),
                               borderwidth=0,
                               thickness=20)

    def apply_theme_to_all_widgets(self):
        """Apply current theme to all widgets (requires app restart for full effect)"""
        # This is a simplified version - full theme switching would require rebuilding the UI
        self.configure(bg=self.theme.get('bg'))

        # Update progress bar style
        self._configure_progressbar_style()

        # Update sound toggle theme
        if hasattr(self, 'sound_toggle'):
            self.sound_toggle.update_theme(self.theme)

        try:
            self.save_settings()  # Save the theme preference
        except Exception:
            pass

    def update_ui_texts(self):
        """Update all UI text labels with current language"""
        try:
            # Update window title
            self.title(self.translator.get('app_title'))

            # Recreate menu bar to apply language changes
            self.create_menu_bar()

            # Check if Korean language (no emojis for Korean)
            is_korean = self.translator.current_lang == 'ko'

            # Update button texts (no emojis for all languages)
            if hasattr(self, 'backup_btn'):
                text = self.translator.get('backup_data')
                self.backup_btn.config(text=text)
            if hasattr(self, 'restore_btn'):
                text = self.translator.get('restore_data')
                self.restore_btn.config(text=text)
            if hasattr(self, 'filters_btn'):
                text = self.translator.get('filters')
                self.filters_btn.config(text=text)
            if hasattr(self, 'driver_backup_btn'):
                text = self.translator.get('backup_drivers')
                self.driver_backup_btn.config(text=text)
            if hasattr(self, 'driver_restore_btn'):
                text = self.translator.get('restore_drivers')
                self.driver_restore_btn.config(text=text)

            # Update tools buttons (no emojis for all languages)
            if hasattr(self, 'browser_profiles_btn'):
                text = self.translator.get('browser')
                self.browser_profiles_btn.config(text=text)
            if hasattr(self, 'check_space_btn'):
                text = self.translator.get('check_space')
                self.check_space_btn.config(text=text)
            if hasattr(self, 'save_log_btn'):
                text = self.translator.get('save_log')
                self.save_log_btn.config(text=text)
            if hasattr(self, 'copy_btn'):
                text = self.translator.get('copy_data')
                self.copy_btn.config(text=text)
            if hasattr(self, 'devmgr_btn'):
                text = self.translator.get('device_mgr')
                self.devmgr_btn.config(text=text)
            if hasattr(self, 'schedule_btn'):
                text = self.translator.get('schedule')
                self.schedule_btn.config(text=text)
            if hasattr(self, 'winget_export_btn'):
                text = self.translator.get('export_apps')
                self.winget_export_btn.config(text=text)
            if hasattr(self, 'file_explorer_btn'):
                text = self.translator.get('explorer')
                self.file_explorer_btn.config(text=text)
            if hasattr(self, 'nas_connect_btn'):
                text = self.translator.get('connect_nas')
                self.nas_connect_btn.config(text=text)
            if hasattr(self, 'nas_disconnect_btn'):
                text = self.translator.get('disconnect')
                self.nas_disconnect_btn.config(text=text)

            # Update status label
            if hasattr(self, 'status_label'):
                current_status = self.status_label.cget('text')
                if 'Select User' in current_status or '사용자 선택' in current_status:
                    self.status_label.config(text=self.translator.get('select_user_begin'))

            # Update sound label (no emojis for all languages)
            if hasattr(self, 'sound_label'):
                sound_text = self.translator.get('sound')
                self.sound_label.config(text=sound_text)

            # Note: Activity Log notice is only set once during initialization (lines 1183-1196)
            # and should not be updated when language changes to preserve existing log history

        except Exception as e:
            print(f"DEBUG: Error updating UI texts: {e}")

    def create_widgets(self):
        # Create menu bar first
        self.create_menu_bar()

        # Modern Windows 11 Header with theme colors - Minimal design with time only
        header_frame = tk.Frame(self, bg=self.theme.get('bg_secondary'), height=40)
        header_frame.pack(fill="x", pady=0)
        header_frame.pack_propagate(False)

        # Center the time in the header
        header_center = tk.Frame(header_frame, bg=self.theme.get('bg_secondary'))
        header_center.pack(expand=True, fill='both')

        self.header_time_label = tk.Label(header_center, text="",
                                         font=("Segoe UI", 11),
                                         bg=self.theme.get('bg_secondary'),
                                         fg=self.theme.get('fg'))
        self.header_time_label.pack(expand=True)
        self._update_header_time()

        # Main content frame with theme background
        main_frame = tk.Frame(self, bg=self.theme.get('bg'), padx=24, pady=20)
        main_frame.pack(fill="both", expand=True)

        # Options row with modern styling
        options_row = tk.Frame(main_frame, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        options_row.pack(fill="x", pady=(0, 12))

        opt_frame = tk.Frame(options_row, bg=self.theme.get('bg_elevated'))
        opt_frame.pack(side="right", padx=12, pady=8)

        modern_font = ("Segoe UI", 9)

        # Retention count (left side)
        retention_label = tk.Label(opt_frame, text="Backups to Keep:",
                                  bg=self.theme.get('bg_elevated'),
                                  fg=self.theme.get('fg'),
                                  font=("Segoe UI", 9, "bold"))
        retention_label.pack(side="left", padx=(8, 4))

        retention_spinbox = tk.Spinbox(opt_frame, from_=0, to=99,
                                      textvariable=self.retention_count_var,
                                      width=5, font=modern_font,
                                      bg=self.theme.get('bg_elevated'),
                                      fg=self.theme.get('fg'),
                                      buttonbackground=self.theme.get('accent'),
                                      relief="solid", bd=1)
        retention_spinbox.pack(side="left", padx=(0, 2))

        retention_info_label = tk.Label(opt_frame, text="(0=all)",
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg_secondary'),
                                       font=modern_font)
        retention_info_label.pack(side="left", padx=(0, 12))

        sep = ttk.Separator(opt_frame, orient='vertical')
        sep.pack(side="left", fill='y', padx=10, pady=4)

        # Log retention
        log_retention_label = tk.Label(opt_frame, text="Log Days:",
                                      bg=self.theme.get('bg_elevated'),
                                      fg=self.theme.get('fg'),
                                      font=("Segoe UI", 9, "bold"))
        log_retention_label.pack(side="left", padx=(0, 4))

        log_retention_spinbox = tk.Spinbox(opt_frame, from_=0, to=365,
                                          textvariable=self.log_retention_days_var,
                                          width=5, font=modern_font,
                                          bg=self.theme.get('bg_elevated'),
                                          fg=self.theme.get('fg'),
                                          buttonbackground=self.theme.get('accent'),
                                          relief="solid", bd=1)
        log_retention_spinbox.pack(side="left", padx=(0, 2))

        log_retention_info_label = tk.Label(opt_frame, text="(0=off)",
                                           bg=self.theme.get('bg_elevated'),
                                           fg=self.theme.get('fg_secondary'),
                                           font=modern_font)
        log_retention_info_label.pack(side="left", padx=(0, 12))

        sep2 = ttk.Separator(opt_frame, orient='vertical')
        sep2.pack(side="left", fill='y', padx=10, pady=4)

        # Hidden files option
        hidden_label = tk.Label(opt_frame, text="Hidden:",
                               bg=self.theme.get('bg_elevated'),
                               fg=self.theme.get('fg'),
                               font=("Segoe UI", 9, "bold"))
        hidden_label.pack(side="left", padx=(0, 6))

        hidden_include = tk.Radiobutton(opt_frame, text="Include",
                                       variable=self.hidden_mode_var, value='include',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        hidden_include.pack(side="left", padx=(0, 4))

        hidden_exclude = tk.Radiobutton(opt_frame, text="Exclude",
                                       variable=self.hidden_mode_var, value='exclude',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        hidden_exclude.pack(side="left", padx=(0, 12))

        sep3 = ttk.Separator(opt_frame, orient='vertical')
        sep3.pack(side="left", fill='y', padx=10, pady=4)

        # System files option
        system_label = tk.Label(opt_frame, text="System:",
                               bg=self.theme.get('bg_elevated'),
                               fg=self.theme.get('fg'),
                               font=("Segoe UI", 9, "bold"))
        system_label.pack(side="left", padx=(0, 6))

        system_include = tk.Radiobutton(opt_frame, text="Include",
                                       variable=self.system_mode_var, value='include',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        system_include.pack(side="left", padx=(0, 4))

        system_exclude = tk.Radiobutton(opt_frame, text="Exclude",
                                       variable=self.system_mode_var, value='exclude',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        system_exclude.pack(side="left", padx=(0, 8))

        # Persist settings
        try:
            self.hidden_mode_var.trace_add('write', lambda *args: self.save_settings())
            self.system_mode_var.trace_add('write', lambda *args: self.save_settings())
            self.retention_count_var.trace_add('write', lambda *args: self.save_settings())
            self.log_retention_days_var.trace_add('write', lambda *args: self.save_settings())
        except Exception:
            pass

        # User selection with modern card style
        user_card = tk.Frame(main_frame, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        user_card.pack(fill="x", pady=(0, 16))

        user_inner = tk.Frame(user_card, bg=self.theme.get('bg_elevated'))
        user_inner.pack(fill="x", padx=16, pady=12)

        tk.Label(user_inner, text="Select User:",
                font=("Segoe UI", 11, "bold"),
                background=self.theme.get('bg_elevated'),
                foreground=self.theme.get('fg')).pack(side="left", padx=(0, 12))

        self.user_var = tk.StringVar()
        self.user_combo = ttk.Combobox(user_inner, textvariable=self.user_var,
                                      state="readonly", width=35, font=("Segoe UI", 10))
        self.user_combo.pack(side="left", fill="x", expand=True)
        self.load_users()

        # Main Operations Card
        main_card = tk.Frame(main_frame, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        main_card.pack(fill="x", pady=(0, 12))

        # Card header
        main_header = tk.Frame(main_card, bg=self.theme.get('bg_elevated'))
        main_header.pack(fill="x", padx=16, pady=(12, 8))

        tk.Label(main_header, text="Main Operations",
                bg=self.theme.get('bg_elevated'),
                fg=self.theme.get('fg'),
                font=("Segoe UI", 12, "bold")).pack(side="left")

        # Sound toggle
        sound_frame = tk.Frame(main_header, bg=self.theme.get('bg_elevated'))
        sound_frame.pack(side="right")

        # Sound label
        self.sound_label = tk.Label(
            sound_frame,
            text=self.translator.get('sound'),
            bg=self.theme.get('bg_elevated'),
            fg=self.theme.get('fg'),
            font=("Segoe UI", 10)
        )
        self.sound_label.pack(side="left", padx=(0, 8))

        # Sound toggle switch
        self.sound_toggle = ToggleSwitch(
            sound_frame,
            width=50,
            height=24,
            variable=self.sound_enabled_var,
            command=self._on_sound_toggle,
            theme=self.theme
        )
        self.sound_toggle.pack(side="left")

        main_buttons_frame = tk.Frame(main_card, bg=self.theme.get('bg_elevated'))
        main_buttons_frame.pack(fill="x", padx=16, pady=(0, 12))

        for i in range(5):
            main_buttons_frame.columnconfigure(i, weight=1, uniform="button")

        btn_font = ("Segoe UI", 9)
        btn_height = 1

        # Create button helper function - Windows 11 style
        def create_button(parent, text, bg_color, hover_color, command, row, col):
            """Windows 11 style modern button with smooth transitions"""
            # Use custom colors if provided, otherwise use theme colors
            btn_bg = bg_color if bg_color else self.theme.get('bg_elevated')
            btn_fg = '#FFFFFF'  # White text for colored buttons
            btn_hover_bg = hover_color if hover_color else self.theme.get('accent')
            btn_hover_fg = '#FFFFFF'
            btn_pressed_bg = hover_color if hover_color else self.theme.get('accent_hover')
            btn_border = self.theme.get('border')

            btn = tk.Button(parent, text=text,
                          bg=btn_bg,
                          fg=btn_fg,
                          font=btn_font,
                          relief="flat",
                          bd=0,
                          highlightthickness=1,
                          borderwidth=1,
                          height=btn_height,
                          command=command,
                          cursor="hand2",
                          padx=12,
                          pady=6)
            btn.config(highlightbackground=btn_border, highlightcolor=btn_border)
            btn.grid(row=row, column=col, padx=4, pady=4, sticky="ew")

            # Windows 11 style hover, press and release effects
            def on_enter(e):
                e.widget.config(bg=btn_hover_bg, fg=btn_hover_fg, relief="flat")

            def on_leave(e):
                e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

            def on_press(e):
                e.widget.config(bg=btn_pressed_bg, relief="flat")

            def on_release(e):
                e.widget.config(bg=btn_hover_bg, relief="flat")

            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            btn.bind("<ButtonPress-1>", on_press)
            btn.bind("<ButtonRelease-1>", on_release)

            return btn

        # Main operation buttons with custom colors
        self.backup_btn = create_button(main_buttons_frame, "Backup Data",
                                        '#DC3545', '#C82333',  # Red background with white text
                                        self.start_backup_thread, 0, 0)

        self.restore_btn = create_button(main_buttons_frame, "Restore Data",
                                         '#007BFF', '#0056B3',  # Blue background with white text
                                         self.start_restore_thread, 0, 1)

        self.filters_btn = create_button(main_buttons_frame, "Filters",
                                         '#28A745', '#218838',  # Green background with white text
                                         self.open_filter_manager, 0, 2)

        self.driver_backup_btn = create_button(main_buttons_frame, "Backup Drivers",
                                              '#DC3545', '#C82333',  # Red background with white text
                                              self.start_driver_backup_thread, 0, 3)

        self.driver_restore_btn = create_button(main_buttons_frame, "Restore Drivers",
                                               '#007BFF', '#0056B3',  # Blue background with white text
                                               self.start_driver_restore_thread, 0, 4)

        # Tools & Utilities Card
        tools_card = tk.Frame(main_frame, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        tools_card.pack(fill="x", pady=(0, 12))

        tools_header = tk.Frame(tools_card, bg=self.theme.get('bg_elevated'))
        tools_header.pack(fill="x", padx=16, pady=(12, 8))

        tk.Label(tools_header, text="Tools & Utilities",
                bg=self.theme.get('bg_elevated'),
                fg=self.theme.get('fg'),
                font=("Segoe UI", 12, "bold")).pack(side="left")

        tools_buttons_frame = tk.Frame(tools_card, bg=self.theme.get('bg_elevated'))
        tools_buttons_frame.pack(fill="x", padx=16, pady=(0, 12))

        for i in range(5):
            tools_buttons_frame.columnconfigure(i, weight=1, uniform="button")

        # Tools buttons - Row 1
        self.browser_profiles_btn = create_button(tools_buttons_frame, "Browser",
                                                  self.theme.get('info'), self.theme.get('accent_hover'),
                                                  self.start_browser_profiles_backup_thread, 0, 0)

        self.check_space_btn = create_button(tools_buttons_frame, "Check Space",
                                            "#8E44AD", "#A569BD",
                                            self.check_space, 0, 1)

        self.save_log_btn = create_button(tools_buttons_frame, "Save Log",
                                          self.theme.get('success'), self.theme.get('success_hover'),
                                          self.save_log, 0, 2)

        self.copy_btn = create_button(tools_buttons_frame, "Copy Data",
                                      self.theme.get('warning'), "#FFB84D",
                                      self.copy_files, 0, 3)

        self.devmgr_btn = create_button(tools_buttons_frame, "Device Mgr",
                                       "#78909C", "#90A4AE",
                                       self.open_device_manager, 0, 4)

        # Tools buttons - Row 2
        self.schedule_btn = create_button(tools_buttons_frame, "Schedule",
                                         self.theme.get('success'), self.theme.get('success_hover'),
                                         self.schedule_backup, 1, 0)

        self.winget_export_btn = create_button(tools_buttons_frame, "Export Apps",
                                              "#00796B", "#009688",
                                              self.start_winget_export_thread, 1, 1)

        self.file_explorer_btn = create_button(tools_buttons_frame, "Explorer",
                                              "#7D98A1", "#90A4AE",
                                              self.open_file_explorer, 1, 2)

        self.nas_connect_btn = create_button(tools_buttons_frame, "Connect NAS",
                                            "#00796B", "#00897B",
                                            self.open_nas_connect_dialog, 1, 3)

        self.nas_disconnect_btn = create_button(tools_buttons_frame, "Disconnect",
                                               self.theme.get('danger'), self.theme.get('danger_hover'),
                                               self.open_nas_disconnect_dialog, 1, 4)

        # Progress section
        progress_card = tk.Frame(main_frame, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        progress_card.pack(fill="x", pady=(0, 12))

        progress_inner = tk.Frame(progress_card, bg=self.theme.get('bg_elevated'))
        progress_inner.pack(fill="x", padx=16, pady=12)

        self.progress_bar = ttk.Progressbar(progress_inner, orient="horizontal",
                                           mode="determinate",
                                           style="green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", expand=True, pady=(0, 8))

        self.status_label = tk.Label(progress_inner,
                                     text=self.translator.get('select_user_begin'),
                                     font=("Segoe UI", 10),
                                     background=self.theme.get('bg_elevated'),
                                     foreground=self.theme.get('fg_secondary'))
        self.status_label.pack()

        # Activity Log Card
        log_card = tk.Frame(main_frame, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        log_card.pack(fill="both", expand=True, pady=(0, 12))

        log_header = tk.Frame(log_card, bg=self.theme.get('bg_elevated'))
        log_header.pack(fill="x", padx=16, pady=(12, 8))

        tk.Label(log_header, text="Activity Log",
                bg=self.theme.get('bg_elevated'),
                fg=self.theme.get('fg'),
                font=("Segoe UI", 12, "bold")).pack(side="left")

        log_inner_frame = tk.Frame(log_card, bg=self.theme.get('bg_elevated'))
        log_inner_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        scrollbar = tk.Scrollbar(log_inner_frame, width=14,
                                bg=self.theme.get('bg_elevated'),
                                troughcolor=self.theme.get('bg'),
                                activebackground=self.theme.get('accent'))
        scrollbar.pack(side="right", fill="y")

        self.log_text = tk.Text(log_inner_frame, height=7,
                               bg=self.theme.get('bg_secondary'),
                               fg=self.theme.get('fg'),
                               relief="flat", bd=0,
                               font=("Consolas", 9),
                               state="disabled",
                               yscrollcommand=scrollbar.set,
                               padx=8, pady=6)
        self.log_text.tag_configure('bold', font=("Consolas", 9, "bold"))
        self.log_text.pack(fill="both", expand=True, side="left")

        scrollbar.config(command=self.log_text.yview)

        # Footer with theme colors - Minimal design
        footer_frame = tk.Frame(main_frame, bg=self.theme.get('bg'))
        footer_frame.pack(fill="x", pady=(8, 0))

        shortcut_hint = tk.Label(footer_frame, text="Press F1 for keyboard shortcuts",
                                font=("Segoe UI", 8),
                                fg=self.theme.get('fg_secondary'),
                                bg=self.theme.get('bg'))
        shortcut_hint.pack(side="right")

        # Note: update_ui_texts() is not called here to prevent duplicate Notice in Activity Log
        # The initial UI texts are already set during widget creation and in __init__
        # update_ui_texts() will be called when language is changed via switch_language()

    def _on_sound_toggle(self):
        """Callback Called When Sound Toggle"""
        try:
            is_enabled = self.sound_enabled_var.get()

            if is_enabled:
                self.message_queue.put(('log', "Sound enabled"))
                # Test sound
                self._play_sound('complete')
            else:
                self.message_queue.put(('log', "Sound disabled"))

            # Save setting
            self.save_settings()

        except Exception as e:
            print(f"DEBUG: Sound toggle error: {e}")

    def _play_sound(self, sound_type='complete'):
        """ windows default sound playback
        
        Args:
            sound_type: 'complete', 'error', 'warning' 등
        """
        try:
            if not self.sound_enabled_var.get():
                return
                
            import winsound

            # Windows system sounds by type
            sounds = {
                'complete': winsound.MB_OK,           # complete sound 
                'error': winsound.MB_ICONHAND,        # error sound
                'warning': winsound.MB_ICONEXCLAMATION,  # warning sound
                'info': winsound.MB_ICONASTERISK      # info sound
            }
            
            sound_flag = sounds.get(sound_type, winsound.MB_OK)

            # Play sound asynchronously (prevent UI blocking)
            threading.Thread(
                target=lambda: winsound.MessageBeep(sound_flag),
                daemon=True
            ).start()
            
        except Exception as e:
            print(f"DEBUG: Sound playback error: {e}")   

    def open_github_link(self, event):
        webbrowser.open_new("https://github.com/gloriouslegacy/ezBAK/releases")

    def show_keyboard_shortcuts_help(self):
        """Show keyboard shortcuts help dialog"""
        if hasattr(self, 'shortcuts') and self.shortcuts:
            self.shortcuts.show_help()

    # UI log (summarized)
    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.update_idletasks()

    # Detailed log -> file only
    def write_detailed_log(self, message):
        """Thread-safe write to detailed log file (file only)."""
        if not self._log_file:
            return
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.log_lock:
            try:
                self._log_file.write(f"[{timestamp}] {message}\n")
                self._log_file.flush()
            except Exception:
                pass

    def open_log_file(self, base_folder, prefix):
        """Creates/log file for an operation (thread-safe)."""
        try:
            ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            filename = f"{prefix}_{ts}.log"
            path = os.path.join(base_folder, filename)
            with self.log_lock:
                self._log_file = open(path, "a", encoding="utf-8", errors="ignore")
                self.log_file_path = path
                try:
                    # Get retention values
                    try:
                        backup_keep = self.retention_count_var.get()
                    except:
                        backup_keep = '2'
                    try:
                        log_keep = self.log_retention_days_var.get()
                    except:
                        log_keep = '30'

                    # Format timestamp for log
                    log_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                    # Write Policy line
                    self._log_file.write(f"[{log_timestamp}] Policy: Hidden={self.hidden_mode_var.get()}, System={self.system_mode_var.get()}, BackupKeep={backup_keep}, LogKeep={log_keep}\n")

                    # Format filters
                    filters = getattr(self, 'filters', {'include': [], 'exclude': []}) or {'include': [], 'exclude': []}

                    # Format include filters
                    include_list = filters.get('include', []) or []
                    if include_list:
                        include_str = ', '.join([f"{f.get('type', '').lower()}:{f.get('pattern', '')}" for f in include_list])
                    else:
                        include_str = 'None'

                    # Format exclude filters - add default hidden files pattern
                    exclude_list = filters.get('exclude', []) or []
                    exclude_patterns = ['name:^\\\\...*']  # Default: exclude hidden files starting with dot
                    for f in exclude_list:
                        pattern_type = f.get('type', '').lower()
                        pattern_val = f.get('pattern', '')
                        exclude_patterns.append(f"{pattern_type}:{pattern_val}")
                    exclude_str = ', '.join(exclude_patterns)

                    # Write Filters line
                    self._log_file.write(f"[{log_timestamp}] Filters: Include=[{include_str}], Exclude=[{exclude_str}]\n")

                    # Write separator
                    self._log_file.write(f"[{log_timestamp}] --------------------------------------------------\n")

                    self._log_file.flush()
                except Exception:
                    pass
            self.message_queue.put(('log', f"Detailed log: {path}"))
        except Exception as e:
            self.message_queue.put(('log', f"Unable to open log file: {e}"))
            self._log_file = None
            self.log_file_path = None

    def close_log_file(self):
        with self.log_lock:
            try:
                if self._log_file:
                    self._log_file.flush()
                    self._log_file.close()
            except Exception:
                pass
            self._log_file = None
            self.log_file_path = None
            

    def _cleanup_old_logs(self, log_folder, retention_days):
        try:
            try:
                retention_days = int(retention_days)
            except (ValueError, TypeError):
                retention_days = 30
            
            if retention_days <= 0:
                self.write_detailed_log("Log cleanup disabled (retention days <= 0)")
                return
            
            if not os.path.isdir(log_folder):
                self.write_detailed_log(f"Log folder does not exist: {log_folder}")
                return
            
            import time
            from datetime import datetime
        
            now = time.time()
            cutoff = now - (retention_days * 86400)
            cutoff_date = datetime.fromtimestamp(cutoff).strftime('%Y-%m-%d %H:%M:%S')
        
            self.write_detailed_log(f"\n[Old log cleanup]")
            self.write_detailed_log(f"Scanning '{log_folder}' for logs older than {retention_days} days...")
            self.write_detailed_log(f"Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            self.write_detailed_log(f"Cutoff time: {cutoff_date}")
            
            # Initialize variables outside try block
            deleted_count = 0
            log_files = []
        
            # Debug: Check all files in folder
            try:
                all_files = os.listdir(log_folder)
                self.write_detailed_log(f"Total files in folder: {len(all_files)}")
                
                log_files = [f for f in all_files if f.endswith('.log')]
                self.write_detailed_log(f"Log files found: {len(log_files)}")
                
                if not log_files:
                    self.write_detailed_log("No .log files found in the folder.")
                    return
                
                for filename in log_files:
                    filepath = os.path.join(log_folder, filename)
                    try:
                        if os.path.isfile(filepath):
                            file_mtime = os.path.getmtime(filepath)
                            file_date = datetime.fromtimestamp(file_mtime).strftime('%Y-%m-%d %H:%M:%S')
                            age_days = (now - file_mtime) / 86400
                        
                            self.write_detailed_log(f"  File: {filename}")
                            self.write_detailed_log(f"    Modified: {file_date}")
                            self.write_detailed_log(f"    Age: {age_days:.1f} days")
                            self.write_detailed_log(f"    Should delete: {file_mtime < cutoff}")
                        
                            if file_mtime < cutoff:
                                try:
                                    os.remove(filepath)
                                    self.write_detailed_log(f"    ✓ Deleted: {filename}")
                                    deleted_count += 1
                                except Exception as e:
                                    self.write_detailed_log(f"    ✗ Delete failed: {e}")
                    except Exception as e:
                        self.write_detailed_log(f"  Error checking {filename}: {e}")
                        
            except Exception as e:
                
                log_files = [f for f in all_files if f.endswith('.log')]
                self.write_detailed_log(f"Log files found: {len(log_files)}")
            
                for filename in log_files:
                    filepath = os.path.join(log_folder, filename)
                    try:
                        if os.path.isfile(filepath):
                            file_mtime = os.path.getmtime(filepath)
                            file_date = datetime.fromtimestamp(file_mtime).strftime('%Y-%m-%d %H:%M:%S')
                            age_days = (now - file_mtime) / 86400
                        
                            self.write_detailed_log(f"  File: {filename}")
                            self.write_detailed_log(f"    Modified: {file_date}")
                            self.write_detailed_log(f"    Age: {age_days:.1f} days")
                            self.write_detailed_log(f"    Should delete: {file_mtime < cutoff}")
                        
                            if file_mtime < cutoff:
                                try:
                                    os.remove(filepath)
                                    self.write_detailed_log(f"    ✓ Deleted: {filename}")
                                    deleted_count += 1
                                except Exception as e:
                                    self.write_detailed_log(f"    ✗ Delete failed: {e}")
                    except Exception as e:
                        self.write_detailed_log(f"  Error checking {filename}: {e}")
            
            except Exception as e:
                self.write_detailed_log(f"Error listing files: {e}")
                return
        
            # deleted_count = 0
            for filename in log_files:
                filepath = os.path.join(log_folder, filename)
                try:
                    if os.path.isfile(filepath) and os.path.getmtime(filepath) < cutoff:
                        os.remove(filepath)
                        self.write_detailed_log(f"  ✓ Deleted: {filename}")
                        deleted_count += 1
                except Exception as e:
                    self.write_detailed_log(f"  ✗ Error deleting {filename}: {e}")
        
            if deleted_count > 0:
                self.write_detailed_log(f"Log cleanup completed. Deleted {deleted_count} files.")
            else:
                self.write_detailed_log("No old log files found to delete.")
            
        except Exception as e:
            self.write_detailed_log(f"Error during log cleanup: {e}")
            import traceback
            self.write_detailed_log(f"Traceback: {traceback.format_exc()}")

    def process_queue(self):
        """Processes messages from the message queue to update the GUI."""
        try:
            while True:
                task, value = self.message_queue.get_nowait()
                if task == 'log':
                    now = time.time()
                    if now - self._last_ui_log_ts >= self.ui_log_throttle_secs:
                        self._last_ui_log_ts = now
                        self.log(value)
                elif task == 'update_progress':
                    try:
                        self.progress_bar['value'] = value
                        try:
                            maxv = float(self.progress_bar['maximum'])
                        except Exception:
                            maxv = 0.0
                        percent = 0
                        if maxv > 0:
                            try:
                                percent = int(min(100, (float(value) / maxv) * 100))
                            except Exception:
                                percent = 0
                        self.status_label.config(text=f"Progress: {percent}%") 
                    except Exception:
                        pass
                elif task == 'space_check_result':
                    self.show_space_check_result(value)
                elif task == 'update_progress_max':
                    try:
                        self.progress_bar['maximum'] = value
                        self.status_label.config(text="Progress: 0%")
                    except Exception:
                        pass
                elif task == 'update_status':
                    self.status_label.config(text=value)
                elif task == 'enable_buttons':
                    try:
                        self.set_buttons_state("normal")
                        print("DEBUG: Buttons enabled")
                    except Exception as e:
                        print(f"DEBUG: Button enable failed: {e}")
                elif task == 'show_error':
                    Win11Dialog.showerror(self.translator.get('error'), value,
                                        parent=self, theme=self.theme, translator=self.translator)
                elif task == 'open_folder':
                    try:
                        os.startfile(value)
                    except Exception:
                        pass
                elif task == 'stop_progress':
                    try:
                        self.progress_bar.stop()  # indeterminate 모드 정지
                        self.progress_bar['mode'] = "determinate"
                    except Exception:
                        pass
                elif task == 'play_sound':
                    # Sound playback in main thread to avoid issues
                    try:
                        self._play_sound(value)
                    except Exception as e:
                        print(f"DEBUG: Sound play failed: {e}")
                elif task == 'reset_progress':
                    # Reset progress bar to initial state
                    try:
                        self.reset_progress_bar()
                    except Exception as e:
                        print(f"DEBUG: Progress bar reset failed: {e}")

                self.message_queue.task_done()
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue) 


    def show_space_check_result(self, result):
        required = result['required']
        available = result['available']
        file_count = result['file_count']
        sources_count = result['sources_count']
        
        margin = available - required
        
        msg = (
            f"Complete:\n\n"
            f"Sources: {sources_count} items\n"
            f"Files: {file_count:,} items\n"
            f"Required Space: {self.format_bytes(required)}\n"
            f"Available: {self.format_bytes(available)}\n"
            f"Margin: {self.format_bytes(margin)}\n"
        )
        
        if required <= available:
            Win11Dialog.showinfo(self.translator.get('space_check'),
                               f"{self.translator.get('sufficient_space')}\n\n{msg}",
                               parent=self, theme=self.theme, translator=self.translator)
            self.message_queue.put(('log', f"Space OK. Required={self.format_bytes(required)} Available={self.format_bytes(available)}"))
            self.message_queue.put(('update_status', "Space check completed - OK"))
        else:
            Win11Dialog.showwarning(self.translator.get('space_check'),
                                  f"{self.translator.get('insufficient_space')}\n\n{msg}",
                                  parent=self, theme=self.theme, translator=self.translator)
            self.message_queue.put(('log', f"Space LOW. Required={self.format_bytes(required)} Available={self.format_bytes(available)}"))
            self.message_queue.put(('update_status', "Space check completed - Insufficient"))


    def set_buttons_state(self, state):
        """Enable or disable all action buttons."""
        self.backup_btn.config(state=state)
        self.restore_btn.config(state=state)
        self.copy_btn.config(state=state)
        self.driver_backup_btn.config(state=state)
        self.driver_restore_btn.config(state=state)
        self.devmgr_btn.config(state=state)
        self.file_explorer_btn.config(state=state)
        self.save_log_btn.config(state=state)
        # Optional buttons if present
        try:
            self.check_space_btn.config(state=state)
        except Exception:
            pass
        try:
            self.schedule_btn.config(state=state)
        except Exception:
            pass
        try:
            self.winget_export_btn.config(state=state)
        except Exception:
            pass
        try:
            self.filters_btn.config(state=state)
        except Exception:
            pass
        try:
            self.nas_connect_btn.config(state=state)
        except Exception:
            pass
        try:
            self.nas_disconnect_btn.config(state=state)
        except Exception:
            pass
        try:
            self.browser_profiles_btn.config(state=state)
        except Exception:
            pass
        try:
            self.exit_btn.config(state=state)
        except Exception:
            pass

    def reset_progress_bar(self):
        """Reset progress bar"""
        try:
            self.progress_bar.stop()  # stop indeterminate mode 
            self.progress_bar['mode'] = "determinate"
            self.progress_bar['value'] = 0
            self.progress_bar['maximum'] = 100
            self.status_label.config(text="Ready")
        except Exception:
            pass

    def set_progress_indeterminate(self):
        """Set progress bar to indeterminate mode"""
        try:
            self.progress_bar['mode'] = "indeterminate"
            self.progress_bar.start()
        except Exception:
            pass

    def save_log(self):
        """Saves the UI log content to a text file (UI only)."""
        if not self.log_text.get("1.0", tk.END).strip():
            Win11Dialog.showinfo(self.translator.get('info'),
                               self.translator.get('no_log_content'),
                               parent=self, theme=self.theme, translator=self.translator)
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            title="Save Log File"
        )

        if not file_path:
            self.message_queue.put(('log', "Log saving cancelled."))
            return

        try:
            log_content = self.log_text.get("1.0", tk.END)
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(log_content)
            self.message_queue.put(('log', f"UI log saved to {file_path}"))
        except Exception as e:
            self.message_queue.put(('log', f"Error saving UI log: {e}"))
            Win11Dialog.showerror(self.translator.get('error'),
                                self.translator.get('error_saving_log').format(e),
                                parent=self, theme=self.theme, translator=self.translator)

    def load_users(self):
        """Gets a list of Windows users and adds them to the combo box."""
        users_path = os.path.join("C:", os.sep, "Users")
        try:
            users = [d for d in os.listdir(users_path) if os.path.isdir(os.path.join(users_path, d)) and not d.startswith('.') and d.lower() != "all users"]
        except Exception:
            users = []
        self.user_combo['values'] = users
        if users:
            self.user_combo.current(0)
            # self._on_user_selected(None)

    def _update_header_time(self):
        """Update header time every second"""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.header_time_label.config(text=current_time)
            self.after(1000, self._update_header_time)
        except Exception:
            pass

    # --- New Function to check for hidden attribute ---
    def is_hidden(self, filepath):
        """
        Unified skip filter.
        Returns True if the entry should be excluded from processing based on:
        - Hidden/System attributes (per user policy)
        - Junctions (reparse points)
        - User-defined include/exclude patterns (extension/name/path)
        Note: Include patterns apply to files only to avoid pruning directories prematurely.
        """
        # Shortcut: nonexistent path -> do not skip (nothing to process yet)
        if not os.path.exists(filepath):
            return False

        # If user chose to include hidden files, never treat as hidden here
        try:
            if getattr(self, 'include_hidden_var', None) and self.include_hidden_var.get():
                return False
        except Exception:
            pass

        is_dir = False
        try:
            is_dir = os.path.isdir(filepath)
        except Exception:
            is_dir = False

        # Also treat certain folder names as hidden/excluded (common Windows hidden folders)
        try:
            base = os.path.basename(filepath).lower()
            if base in ('appdata', 'application data', 'local settings'):
                # Only include these special folders if user explicitly includes both Hidden and System
                try:
                    hid = getattr(self, 'hidden_mode_var', None) and self.hidden_mode_var.get() == 'include'
                    sysm = getattr(self, 'system_mode_var', None) and self.system_mode_var.get() == 'include'
                    if not (hid and sysm):
                        return True
                except Exception:
                    return True
        except Exception:
            pass

        # Attribute checks (hidden/system/reparse)
        try:
            if _have_pywin32 and win32api is not None and win32con is not None:
                attributes = win32api.GetFileAttributes(filepath)
                is_hidden_attr = bool(attributes & win32con.FILE_ATTRIBUTE_HIDDEN)
                is_system_attr = bool(attributes & win32con.FILE_ATTRIBUTE_SYSTEM)
                is_reparse = bool(attributes & win32con.FILE_ATTRIBUTE_REPARSE_POINT)
            else:
                try:
                    attributes = _get_file_attributes(filepath)
                except Exception:
                    attributes = FILE_ATTRIBUTE_NORMAL
                is_hidden_attr = bool(attributes & FILE_ATTRIBUTE_HIDDEN)
                is_system_attr = bool(attributes & FILE_ATTRIBUTE_SYSTEM)
                is_reparse = bool(attributes & FILE_ATTRIBUTE_REPARSE_POINT)

            if is_reparse:
                return True
            if is_hidden_attr and getattr(self, 'hidden_mode_var', None) and self.hidden_mode_var.get() == 'exclude':
                return True
            if is_system_attr and getattr(self, 'system_mode_var', None) and self.system_mode_var.get() == 'exclude':
                return True
        except Exception:
            # If attribute read fails, continue to pattern logic
            pass

        # Pattern-based filters
        try:
            filters = getattr(self, 'filters', {'include': [], 'exclude': []}) or {'include': [], 'exclude': []}
            name = os.path.basename(filepath)
            name_l = name.lower()
            ext_l = os.path.splitext(name_l)[1].lstrip('.')
            filepath_lower = filepath.lower()

            def _matches(rule):
                try:
                    rtype = (rule.get('type') or '').lower()
                    patt = (rule.get('pattern') or '').strip()
                    if not patt:
                        return False
                    patt_lower = patt.lower()
                    
                    if rtype == 'ext':
                        if is_dir:
                            return False
                        patt_clean = patt_lower.lstrip('*.')
                        return ext_l == patt_clean
                        
                    elif rtype == 'name':
                        if '*' in patt_lower or '?' in patt_lower:
                            return fnmatch.fnmatch(name_l, patt_lower)
                        else:
                            return name_l == patt_lower
                        
                    elif rtype == 'path':
                        # Simple contains check
                        if patt_lower in filepath_lower:
                            return True
                        # Normalized path matching
                        path_normalized = filepath_lower.replace('\\', '/')
                        patt_normalized = patt_lower.replace('\\', '/')
                        if patt_normalized in path_normalized:
                            return True
                        # Wildcard matching
                        if '*' in patt_normalized or '?' in patt_normalized:
                            return fnmatch.fnmatch(path_normalized, patt_normalized)
                    return False
                except Exception:
                    return False

            # Excludes take precedence
            for rule in filters.get('exclude', []) or []:
                if _matches(rule):
                    return True

            # Includes: if any include rules exist, only allow matching files; do not apply include to directories
            inc_rules = filters.get('include', []) or []
            if inc_rules and not is_dir:
                for rule in inc_rules:
                    if _matches(rule):
                        return False
                # no include matched -> exclude
                return True
        except Exception:
            pass

        # Not excluded by any rule
        return False

    # Helper: copy single file in chunks and report bytes progress
    def copy_file_with_progress(self, src, dst, progress_callback, buffer_size=64*1024):
        """Copies a file in chunks and calls progress_callback(bytes_copied) incrementally."""
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            copied = 0
            with open(src, "rb") as fsrc:
                with open(dst, "wb") as fdst:
                    while True:
                        buf = fsrc.read(buffer_size)
                        if not buf:
                            break
                        fdst.write(buf)
                        copied += len(buf)
                        progress_callback(len(buf))
            try:
                shutil.copystat(src, dst)
            except Exception:
                pass
            return copied
        except Exception as e:
            self.write_detailed_log(f"Error copying file {src} -> {dst}: {e}")
            return 0

    def compute_total_bytes_fast(self, root_path, max_files_per_update=100):
        """
        Fast size calculation - calculates while updating intermediate results to UI
        """
        total = 0
        file_count = 0
        last_update_time = time.time()
        
        try:
            for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=lambda e: None):
                # Filter hidden directories
                dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if self.is_hidden(fp):
                        continue
                        
                    try:
                        size = os.path.getsize(fp)
                        total += size
                        file_count += 1
                        
                        # Periodically update UI (improves responsiveness)
                        if file_count % max_files_per_update == 0:
                            current_time = time.time()
                            if current_time - last_update_time >= 0.1:  # Update every 0.1 seconds
                                self.message_queue.put(('update_status', f"Calculating... {file_count} files, {self.format_bytes(total)}"))
                                self.update_idletasks()  # Update UI
                                last_update_time = current_time
                                
                    except (OSError, IOError):
                        continue
                        
        except Exception as e:
            self.write_detailed_log(f"Fast compute error: {e}")
        
        return total

    # --- User Data Backup & Restore Functions ---
    def start_backup_thread(self):
        """Starts the backup process in a separate thread."""
        if not self.user_var.get():
            Win11Dialog.showwarning(self.translator.get('warning'),
                                  self.translator.get('select_user_warning'),
                                  parent=self, theme=self.theme, translator=self.translator)
            return

        backup_path = filedialog.askdirectory(title="Select Backup Folder")
        if not backup_path:
            self.message_queue.put(('log', "Backup folder selection cancelled."))
            return

        # open detailed log file for this operation
        self.open_log_file(backup_path, f"{self.user_var.get()}_backup")

        self.message_queue.put(('log', "User data backup process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['mode'] = "determinate"

        backup_thread = threading.Thread(target=self.run_backup, args=(backup_path,), daemon=True)
        backup_thread.start()

    def run_backup(self, backup_path):
        """Executes the actual backup logic with improved error handling and debugging."""
        bytes_copied = 0
        backup_success = False
        
        try:
            user_name = self.user_var.get()
            user_profile_path = os.path.join("C:", os.sep, "Users", user_name)

            if not os.path.exists(user_profile_path):
                self.message_queue.put(('show_error', f"User folder '{user_profile_path}' not found."))
                return

            self.write_detailed_log(f"Backup start: {user_profile_path} -> {backup_path}")
            self.message_queue.put(('log', f"Backing up user: {user_name}"))

            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            destination_dir = os.path.join(backup_path, f"{user_name}_backup_{timestamp}")

            if os.path.exists(destination_dir):
                try:
                    shutil.rmtree(destination_dir)
                    self.write_detailed_log(f"Deleted existing backup folder '{destination_dir}'.")
                except Exception as e:
                    self.write_detailed_log(f"Failed to delete existing backup folder: {e}")

            os.makedirs(destination_dir, exist_ok=True)
            self.write_detailed_log(f"Created new backup folder '{destination_dir}'.")

            self.message_queue.put(('update_status', "Calculating total size..."))
            total_bytes = self.compute_total_bytes_safe(user_profile_path)
            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Calculating total size... {self.format_bytes(total_bytes)}")

            # Check available space
            self.message_queue.put(('update_status', "Checking available space..."))
            free = self.get_free_space_reliable(backup_path)
            
            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space for backup. Required={total_bytes} Free={free}")
                self.message_queue.put(('show_error', f"There is not enough space in the destination.\nRequired: {self.format_bytes(total_bytes)}\nAvailable: {self.format_bytes(free)}"))
                return

            self.message_queue.put(('update_status', "Starting backup..."))
            
            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            files_processed = 0
            folders_processed = 0
            
            try:
                for dirpath, dirnames, filenames in os.walk(user_profile_path, topdown=True, onerror=None):
                    try:
                        # Log the current folder being processed
                        rel_path = os.path.relpath(dirpath, user_profile_path)
                        self.write_detailed_log(f"Processing directory: {dirpath} (relative: {rel_path})")

                        # Copy original dirnames
                        original_dirnames = dirnames.copy()
                        dirnames.clear()
                        
                        for d in original_dirnames:
                            full_dir_path = os.path.join(dirpath, d)
                            try:
                                if not self.is_hidden_safe(full_dir_path):
                                    dirnames.append(d)
                                else:
                                    self.write_detailed_log(f"Skipping hidden directory: {full_dir_path}")
                            except Exception as e:
                                self.write_detailed_log(f"Error checking directory {full_dir_path}: {e}")
                                # Include directory in the list (safe choice)
                                dirnames.append(d)

                        # Determine destination directory
                        rel_dir = os.path.relpath(dirpath, user_profile_path)
                        dest_dir = os.path.join(destination_dir, rel_dir) if rel_dir != "." else destination_dir
                        
                        try:
                            os.makedirs(dest_dir, exist_ok=True)
                            self.write_detailed_log(f"Created folder: {dest_dir}")
                            folders_processed += 1

                            # Update status every 5 folders
                            if folders_processed % 5 == 0:
                                self.message_queue.put(('update_status', f"Processing folder {folders_processed}: {os.path.basename(dirpath)}"))
                        except Exception as e:
                            self.write_detailed_log(f"ERROR: Failed to create folder {dest_dir}: {e}")
                            continue

                        # Copy files
                        for f in filenames:
                            try:
                                src_file = os.path.join(dirpath, f)
                                dest_file = os.path.join(dest_dir, f)

                                # Check if the file is hidden
                                if self.is_hidden_safe(src_file):
                                    self.write_detailed_log(f"Skipping hidden file: {src_file}")
                                    continue
                                
                                # Copy file with progress
                                self.write_detailed_log(f"Copying file: {src_file} -> {dest_file}")
                                copied = self.copy_file_with_progress_safe(src_file, dest_file, progress_cb)
                                files_processed += 1

                                # Update status every 10 files
                                if files_processed % 10 == 0:
                                    self.message_queue.put(('log', f"Copied {files_processed} files..."))
                                    self.message_queue.put(('update_status', f"Copied {files_processed} files"))
                                    
                            except Exception as e:
                                self.write_detailed_log(f"ERROR: Failed to copy file {src_file}: {e}")
                                continue
                                
                    except Exception as e:
                        self.write_detailed_log(f"ERROR: Failed to process directory {dirpath}: {e}")
                        continue
                        
            except Exception as e:
                self.write_detailed_log(f"CRITICAL ERROR in backup loop: {e}")
                self.message_queue.put(('show_error', f"A critical error occurred during backup: {e}"))
                return

            # Update final progress
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log(f"Backup completed. Files backed up: {files_processed}, Folders: {folders_processed}, Total size: {self.format_bytes(bytes_copied)}")

            # Backup browser bookmarks
            try:
                self.backup_browser_bookmarks(backup_path)
            except Exception as e:
                self.write_detailed_log(f"Browser bookmark backup failed: {e}")

            # Cleanup
            try:
                self._cleanup_old_backups(backup_path, user_name)
                try:
                    retention_days = int(self.log_retention_days_var.get())
                except (ValueError, TypeError):
                    retention_days = 30
                self._cleanup_old_logs(backup_path, retention_days)
            except Exception as e:
                self.write_detailed_log(f"Cleanup failed: {e}")
            
            self.write_detailed_log("Backup successfully completed.")
            self.message_queue.put(('log', "Backup complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Backup complete!"))
            self.message_queue.put(('open_folder', destination_dir))
            
            backup_success = True

        except Exception as e:
            self.write_detailed_log(f"Backup error: {e}")
            import traceback
            self.write_detailed_log(f"Backup traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"An error occurred during backup: {e}"))
            backup_success = False
            
        finally:   # Finally block for entire function
            # Always execute cleanup tasks
            try:
                self.close_log_file()
            except Exception:
                pass

            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass

            # Enable buttons
            self.message_queue.put(('enable_buttons', None))

            # Additional safeguard: also try direct button activation
            try:
                self.after(1000, lambda: self.set_buttons_state("normal"))
            except Exception:
                pass

            # Reset progress bar after 2 seconds
            try:
                self.after(2000, lambda: self.message_queue.put(('reset_progress', None)))
            except Exception:
                pass

            if backup_success:
                self.message_queue.put(('update_status', self.translator.get('backup_complete')))
                # Complete sound playback
                self.message_queue.put(('play_sound', 'complete'))
            else:
                self.message_queue.put(('update_status', self.translator.get('operation_finished')))
                # Error sound playback
                self.message_queue.put(('play_sound', 'error'))          
            
            
    def copy_file_with_progress_safe(self, src, dst, progress_callback, buffer_size=64*1024, timeout_seconds=30):
        """
        Safe file copy with timeout
        """
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            copied = 0
            start_time = time.time()
            
            with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
                while True:
                    # Check timeout
                    if time.time() - start_time > timeout_seconds:
                        self.write_detailed_log(f"TIMEOUT: File copy timeout for {src}")
                        raise TimeoutError(f"File copy timeout: {src}")
                    
                    buf = fsrc.read(buffer_size)
                    if not buf:
                        break
                    fdst.write(buf)
                    copied += len(buf)
                    progress_callback(len(buf))
            
            # Copy file attributes (with timeout)
            try:
                shutil.copystat(src, dst)
            except Exception:
                pass  # Ignore attribute copy failure
                
            return copied
        except Exception as e:
            self.write_detailed_log(f"ERROR: copying file {src} -> {dst}: {e}")
            return 0

    def compute_total_bytes_safe(self, root_path):
        """
        Safe total size calculation
        """
        total = 0
        file_count = 0
        
        try:
            for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=lambda e: self.write_detailed_log(f"Walk error: {e}")):
                # Safe directory filtering
                original_dirnames = dirnames.copy()
                dirnames.clear()
                
                for d in original_dirnames:
                    try:
                        if not self.is_hidden_safe(os.path.join(dirpath, d), timeout_seconds=1):
                            dirnames.append(d)
                    except Exception:
                        dirnames.append(d)  # Include on error
                
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    try:
                        if not self.is_hidden_safe(fp, timeout_seconds=1):
                            size = os.path.getsize(fp)
                            total += size
                            file_count += 1
                            
                            # Update progress every 100 files
                            if file_count % 100 == 0:
                                self.message_queue.put(('update_status', f"Calculating total size... {file_count} files, {self.format_bytes(total)}"))
                                
                    except Exception as e:
                        self.write_detailed_log(f"Size calculation error for {fp}: {e}")
                        continue
                        
        except Exception as e:
            self.write_detailed_log(f"Total bytes calculation error: {e}")
        
        return total

    def is_hidden_safe(self, filepath, timeout_seconds=2):
        """
        is_hidden function with timeout - includes filters
        """
        try:
            # Check existence
            if not os.path.exists(filepath):
                return False
            
            is_dir = os.path.isdir(filepath)
            
            # Check filters (OneDrive, etc.)
            filters = getattr(self, 'filters', {'include': [], 'exclude': []}) or {'include': [], 'exclude': []}
            name = os.path.basename(filepath)
            name_l = name.lower()
            ext_l = os.path.splitext(name_l)[1].lstrip('.')
            filepath_lower = filepath.lower()
            
            def _matches(rule):
                try:
                    rtype = (rule.get('type') or '').lower()
                    patt = (rule.get('pattern') or '').strip()
                    if not patt:
                        return False
                    patt_lower = patt.lower()
                    
                    if rtype == 'ext':
                        if is_dir:
                            return False
                        patt_clean = patt_lower.lstrip('*.')
                        return ext_l == patt_clean
                        
                    elif rtype == 'name':
                        if '*' in patt_lower or '?' in patt_lower:
                            return fnmatch.fnmatch(name_l, patt_lower)
                        else:
                            return name_l == patt_lower
                        
                    elif rtype == 'path':
                        if patt_lower in filepath_lower:
                            return True
                        path_normalized = filepath_lower.replace('\\', '/')
                        patt_normalized = patt_lower.replace('\\', '/')
                        if patt_normalized in path_normalized:
                            return True
                        if '*' in patt_normalized or '?' in patt_normalized:
                            return fnmatch.fnmatch(path_normalized, patt_normalized)
                    return False
                except Exception:
                    return False
            
            # Check exclude filters
            for rule in filters.get('exclude', []) or []:
                if _matches(rule):
                    return True
            
            # Check include filters (files only)
            inc_rules = filters.get('include', []) or []
            if inc_rules and not is_dir:
                has_match = False
                for rule in inc_rules:
                    if _matches(rule):
                        has_match = True
                        break
                if not has_match:
                    return True
            
            # Check basic hidden files
            base = os.path.basename(filepath).lower()
            if base.startswith('.') or base in ('thumbs.db', 'desktop.ini', '$recycle.bin'):
                return True
            
            # Check Windows attributes
            try:
                if _have_pywin32 and win32api and win32con:
                    attrs = win32api.GetFileAttributes(filepath)
                    is_reparse = bool(attrs & win32con.FILE_ATTRIBUTE_REPARSE_POINT)
                    if is_reparse:
                        return True
                    is_hidden_attr = bool(attrs & win32con.FILE_ATTRIBUTE_HIDDEN)
                    is_system_attr = bool(attrs & win32con.FILE_ATTRIBUTE_SYSTEM)
                    if is_hidden_attr and self.hidden_mode_var.get() == 'exclude':
                        return True
                    if is_system_attr and self.system_mode_var.get() == 'exclude':
                        return True
            except Exception:
                pass
            
            return False
            
        except Exception:
            return False

    def _cleanup_old_backups(self, backup_path, user_name):
        """Deletes old backups, keeping N most recent ones based on UI setting."""
        try:
            try:
                retention_count = int(self.retention_count_var.get())
            except (ValueError, TypeError):
                retention_count = 2  # Fallback to a safe default

            if retention_count <= 0:
                self.write_detailed_log("\n[Old backup cleanup]")
                self.write_detailed_log("Backup cleanup is disabled (retention count is 0 or less).")
                return

            self.write_detailed_log("\n[Old backup cleanup]")
            # backup_prefix = f"{user_name}_backup_"
            user_name_lower = user_name.lower()
            backup_prefix_lower = f"{user_name_lower}_backup_"
            all_backups = [d for d in os.listdir(backup_path) if os.path.isdir(os.path.join(backup_path, d)) and d.lower().startswith(backup_prefix_lower)]
            # all_backups = [d for d in os.listdir(backup_path) if os.path.isdir(os.path.join(backup_path, d)) and d.startswith(backup_prefix)]
            
            if len(all_backups) > retention_count:
                self.write_detailed_log(f"Found {len(all_backups)} backups. Keeping the {retention_count} most recent.")
                
                # Sorts chronologically based on 'YYYY-MM-DD' date format in the name
                all_backups.sort()
                
                backups_to_delete = all_backups[:-retention_count]
                for backup_dir_name in backups_to_delete:
                    dir_to_delete = os.path.join(backup_path, backup_dir_name)
                    self.write_detailed_log(f"  Deleting old backup: {dir_to_delete}")
                    shutil.rmtree(dir_to_delete, onerror=self._handle_rmtree_error)
                self.write_detailed_log("Cleanup finished.")
            else:
                self.write_detailed_log(f"Found {len(all_backups)} backups. No old backups to clean up (retention policy: keep {retention_count}).")
        except Exception as e:
            self.write_detailed_log(f"ERROR: An error occurred during backup cleanup: {e}")

    def _handle_rmtree_error(self, func, path, exc_info):
        """
        Error handler for shutil.rmtree that handles read-only files.
        If a file is read-only, it changes its permissions and retries the deletion.
        """
        if isinstance(exc_info[1], PermissionError):
            try:
                os.chmod(path, stat.S_IWRITE)
                func(path)
                self.write_detailed_log(f"INFO: Removed read-only and successfully deleted {path}")
            except Exception as e:
                self.write_detailed_log(f"ERROR: Could not delete {path} even after removing read-only flag: {e}")
        else:
            self.write_detailed_log(f"ERROR: Deletion failed for {path}: {exc_info[1]}")

    def backup_browser_bookmarks(self, backup_path):
        """Backs up Chrome and Edge bookmark files."""
        self.write_detailed_log("Backing up Chrome and Edge browser bookmarks...")
        chrome_bookmark_path = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Google', 'Chrome', 'User Data', 'Default', 'Bookmarks')
        edge_bookmark_path = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Microsoft', 'Edge', 'User Data', 'Default', 'Bookmarks')

        try:
            if os.path.exists(chrome_bookmark_path):
                shutil.copy(chrome_bookmark_path, os.path.join(backup_path, 'Chrome_Bookmarks'))
                self.write_detailed_log("- Chrome bookmarks backed up.")
            else:
                self.write_detailed_log("- Chrome bookmark file not found.")
        except Exception as e:
            self.write_detailed_log(f"- Chrome bookmark backup error: {e}")

        try:
            if os.path.exists(edge_bookmark_path):
                shutil.copy(edge_bookmark_path, os.path.join(backup_path, 'Edge_Bookmarks'))
                self.write_detailed_log("- Edge bookmarks backed up.")
            else:
                self.write_detailed_log("- Edge bookmark file not found.")
        except Exception as e:
            self.write_detailed_log(f"- Edge bookmark backup error: {e}")

    def start_restore_thread(self):
        """Starts the restore process in a separate thread."""
        if not self.user_var.get():
            Win11Dialog.showwarning(self.translator.get('warning'),
                                  self.translator.get('select_user_warning'),
                                  parent=self, theme=self.theme, translator=self.translator)
            return

        response = Win11Dialog.askyesno(self.translator.get('warning'),
                                       self.translator.get('overwrite_warning').format(self.user_var.get()),
                                       parent=self, theme=self.theme, translator=self.translator)
        if not response:
            self.message_queue.put(('log', "Restore process cancelled."))
            return

        backup_folder = filedialog.askdirectory(title="Select Backed-up Folder")
        if not backup_folder:
            self.message_queue.put(('log', "Restore folder selection cancelled."))
            return
        
        log_folder = os.path.dirname(backup_folder)

        # open log file in backup folder
        self.open_log_file(backup_folder, f"{self.user_var.get()}_restore")

        self.message_queue.put(('log', "User data restore process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['mode'] = "determinate"

        restore_thread = threading.Thread(target=self.run_restore, args=(backup_folder,), daemon=True)
        restore_thread.start()

    def run_restore(self, backup_folder):
        """Executes the actual restore logic (byte-based progress)."""
        bytes_copied = 0
        restore_success = False
        
        try:
            user_name = self.user_var.get()
            
            if backup_folder.lower().endswith(".log"):
                self.message_queue.put(('show_error',
                    f"The selected path is a log file.\n\n"
                    f"'{backup_folder}'\n\n"
                    f"You must select a backup folder"))
                self.write_detailed_log("Restore failed: .log file selected instead of backup folder.")
                return

            if os.path.isfile(backup_folder):
                self.message_queue.put(('show_error',
                    f"The selected path is a file. Please select a folder:\n{backup_folder}"))
                self.write_detailed_log("Restore failed: backup_folder is a file, expected directory.")
                return

            if os.path.isfile(backup_folder):
                self.message_queue.put(('show_error',
                    f"Selected path is a file, not a backup folder:\n{backup_folder}"))
                self.write_detailed_log("Restore failed: backup_folder is a file, expected directory.")
                return
            if not os.path.isdir(backup_folder):
                self.message_queue.put(('show_error', f"Backup folder not found:\n{backup_folder}"))
                self.write_detailed_log("Restore failed: backup folder not found.")
                return

            self.message_queue.put(('log', f"Selected backup folder: {backup_folder}"))
            self.write_detailed_log(f"Restore start for user '{user_name}' from {backup_folder}")

            # backup_folders = [d for d in os.listdir(backup_folder) if d.startswith(f"{user_name}_backup_")]
            # backup_folders = [d for d in os.listdir(backup_folder) if os.path.isdir(os.path.join(backup_folder, d)) and d.startswith(f"{user_name}_backup_")]
            user_name_lower = user_name.lower()
            backup_folders = [d for d in os.listdir(backup_folder) if os.path.isdir(os.path.join(backup_folder, d)) and d.lower().startswith(f"{user_name_lower}_backup_")]
            
            
            if not backup_folders:
                self.message_queue.put(('show_error', f"No backup data found for '{user_name}' in:\n{backup_folder}"))
                self.write_detailed_log("Restore failed: no matching backup folder.")
                return

            source_dir = os.path.join(backup_folder, sorted(backup_folders, reverse=True)[0])
            destination_dir = os.path.join("C:", os.sep, "Users", user_name)
            self.write_detailed_log(f"Restoring from {source_dir} to {destination_dir}")

             
            total_bytes = 0
            try:
                for dirpath, dirnames, filenames in os.walk(source_dir, topdown=True, onerror=lambda e: None):
                    dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                    for f in filenames:
                        fp = os.path.join(dirpath, f)
                        if not self.is_hidden(fp):
                            try:
                                total_bytes += os.path.getsize(fp)
                            except Exception:
                                pass
            except Exception as e:
                self.write_detailed_log(f"Total size calculation error: {e}")
                total_bytes = 1024 * 1024 * 100  # 100MB 기본값          
            
            
            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Total bytes to restore: {total_bytes}")

            try:
                free = shutil.disk_usage(destination_dir).free
            except Exception:
                free = self.get_free_space(destination_dir)
            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space. Required={total_bytes}, Free={free}")
                self.message_queue.put(('show_error',
                    f"Not enough free space to restore.\n"
                    f"Required: {self.format_bytes(total_bytes)}\nAvailable: {self.format_bytes(free)}"))
                return

            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            files_processed = 0
            for dirpath, dirnames, filenames in os.walk(source_dir, topdown=True, onerror=lambda e: None):
                dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                rel_dir = os.path.relpath(dirpath, source_dir)
                dest_dir = os.path.join(destination_dir, rel_dir) if rel_dir != "." else destination_dir
                os.makedirs(dest_dir, exist_ok=True)
                self.write_detailed_log(f"Created folder: {dest_dir}")

                for f in filenames:
                    src_file = os.path.join(dirpath, f)
                    if self.is_hidden(src_file):
                        self.write_detailed_log(f"Skipping hidden file: {src_file}")
                        continue
                    dest_file = os.path.join(dest_dir, f)
                    self.write_detailed_log(f"Restoring file: {src_file} -> {dest_file}")
                    self.copy_file_with_progress(src_file, dest_file, progress_cb)
                    files_processed += 1
                    if files_processed % 50 == 0:
                        self.message_queue.put(('log', f"Restored {files_processed} files..."))
                        self.write_detailed_log(f"{files_processed} files restored so far; bytes_copied={bytes_copied}")

            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log(f"Restore completed. Files restored: {files_processed}, "
                                    f"Total size: {self.format_bytes(bytes_copied)}")
            self.write_detailed_log("Restore successfully completed.")

            self.restore_browser_bookmarks(backup_folder)
            
            try:
                retention_days = int(self.log_retention_days_var.get())
            except (ValueError, TypeError):
                retention_days = 30
            self._cleanup_old_logs(backup_folder, retention_days)
            
            self.message_queue.put(('log', "Restore complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Restore complete!"))
            restore_success = True

        except Exception as e:
            self.write_detailed_log(f"Restore error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during restore: {e}"))
            restore_success = False
            
        finally:
            try:
                self.close_log_file()
            except Exception:
                pass

            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass

            # Reset progress bar
            self.message_queue.put(('reset_progress', None))

            self.message_queue.put(('enable_buttons', None))

            if restore_success:
                self.message_queue.put(('update_status', "Restore complete!"))
                self.message_queue.put(('play_sound', 'complete'))
            else:
                self.message_queue.put(('update_status', "Operation finished"))
                self.message_queue.put(('play_sound', 'error'))

            if hasattr(self, '_restore_completed') and self._restore_completed:
                self.message_queue.put(('update_status', "Restore complete!"))
            else:
                self.message_queue.put(('update_status', "Operation finished"))

    def restore_browser_bookmarks(self, backup_path):
        """Restores backed-up bookmark files to their original location."""
        self.write_detailed_log("Restoring Chrome and Edge browser bookmarks...")
        chrome_bookmark_dest = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Google', 'Chrome', 'User Data', 'Default', 'Bookmarks')
        edge_bookmark_dest = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Microsoft', 'Edge', 'User Data', 'Default', 'Bookmarks')

        try:
            chrome_bookmark_source = os.path.join(backup_path, 'Chrome_Bookmarks')
            if os.path.exists(chrome_bookmark_source):
                shutil.copy(chrome_bookmark_source, chrome_bookmark_dest)
                self.write_detailed_log("- Chrome bookmarks restored.")
            else:
                self.write_detailed_log("- Backed-up Chrome bookmark file not found.")
        except Exception as e:
            self.write_detailed_log(f"- Chrome bookmark restore error: {e}")

        try:
            edge_bookmark_source = os.path.join(backup_path, 'Edge_Bookmarks')
            if os.path.exists(edge_bookmark_source):
                shutil.copy(edge_bookmark_source, edge_bookmark_dest)
                self.write_detailed_log("- Edge bookmarks restored.")
            else:
                self.write_detailed_log("- Backed-up Edge bookmark file not found.")
        except Exception as e:
            self.write_detailed_log(f"- Edge bookmark restore error: {e}")

    def copy_files(self):
        """Prompts for source and destination using a checkbox tree, then starts the copy process."""
        try:
            print("DEBUG: copy_files called")
            self.message_queue.put(('log', "Copy process initiated."))

            # 1) Open a checkbox tree dialog to select files/folders
            try:
                print("DEBUG: Creating SelectSourcesDialog...")
                dlg = SelectSourcesDialog(self, is_hidden_fn=self.is_hidden, title=self.translator.get('select_sources'))
                print("DEBUG: SelectSourcesDialog created, waiting for window...")
                self.wait_window(dlg)
                print("DEBUG: Dialog closed")
                
                source_paths_list = dlg.selected_paths if getattr(dlg, 'selected_paths', None) else []
                print(f"DEBUG: Got {len(source_paths_list)} selected paths")
                
            except Exception as e:
                print(f"DEBUG: Source selection failed: {e}")
                import traceback
                print(f"DEBUG: Traceback: {traceback.format_exc()}")
                self.message_queue.put(('log', f"Source selection failed: {e}"))
                return

            if not source_paths_list:
                print("DEBUG: No sources selected")
                self.message_queue.put(('log', "Source selection cancelled."))
                return

            print(f"DEBUG: Selected sources: {source_paths_list}")

            # 2) Ask for destination folder
            print("DEBUG: Asking for destination folder...")
            destination_dir = filedialog.askdirectory(title=self.translator.get('select_destination_folder'))
            print(f"DEBUG: Destination selected: {destination_dir}")
            
            if not destination_dir:
                print("DEBUG: No destination selected")
                self.message_queue.put(('log', "Destination selection cancelled."))
                return

            # Log name for Copy Data: 'copy_yearmonthday_hourminutesecond'
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            # log_prefix = f"copy_data{timestamp_str}"
            log_prefix = f"copy_data"
            # open log file in destination folder
            self.open_log_file(destination_dir, log_prefix)

            self.message_queue.put(('log', "Copy process started."))
            self.set_buttons_state("disabled")
            self.progress_bar['value'] = 0
            self.progress_bar['mode'] = "determinate"
            self.message_queue.put(('update_status', "Starting copy process..."))

            # Start the copy thread
            copy_thread = threading.Thread(target=self.run_copy_thread, args=(source_paths_list, destination_dir), daemon=True)
            copy_thread.start()
            
        except Exception as e:
            print(f"DEBUG: copy_files failed: {e}")
            import traceback
            print(f"DEBUG: Traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"Copy process failed: {e}"))

    def format_bytes(self, size):
        try:
            size = float(size)
        except Exception:
            return str(size)
        for unit in ["B", "KB", "MB", "GB", "TB", "PB"]:
            if size < 1024.0 or unit == "PB":
                return f"{size:.2f} {unit}"
            size /= 1024.0

    def get_free_space_reliable(self, path):
        """
        More reliable free space calculation
        """
        try:
            # Try shutil.disk_usage first
            usage = shutil.disk_usage(path)
            return usage.free
        except Exception:
            try:
                # Calculate free space via Windows API
                import ctypes
                free_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    ctypes.c_wchar_p(path),
                    ctypes.pointer(free_bytes),
                    None,
                    None
                )
                return free_bytes.value
            except Exception:
                try:
                    # Retry with drive root
                    drive = os.path.splitdrive(path)[0]
                    if drive:
                        root_path = drive + os.sep
                        usage = shutil.disk_usage(root_path)
                        return usage.free
                except Exception:
                    pass
        
        # If all methods fail, return sufficiently large value (continue processing)
        return 999 * 1024**4  # 999 TB
    

    def schedule_backup(self):
        dlg = None
        try:
            print("DEBUG: Creating ScheduleBackupDialog...")
            dlg = ScheduleBackupDialog(self)
            print(f"DEBUG: Dialog created, action={getattr(dlg, 'action', 'NOT_SET')}")
            
            print("DEBUG: Calling wait_window...")
            self.wait_window(dlg)
            print("DEBUG: wait_window completed")

            # Check if dlg still exists and is valid
            if dlg and hasattr(dlg, 'action'):
                action = dlg.action
                print(f"DEBUG: Dialog action = {action}")
                
                if action == 'create':
                    if hasattr(dlg, 'result') and dlg.result:
                        data = dlg.result
                        print(f"DEBUG: Creating scheduled task with data: {data}")

                        # Disable buttons to prevent multiple clicks during task creation
                        self.set_buttons_state("disabled")

                        # Create and start thread for task creation (prevents UI blocking)
                        task_thread = threading.Thread(
                            target=self.run_create_scheduled_task_thread,
                            args=(data,),
                            daemon=True
                        )
                        task_thread.start()
                        print("DEBUG: Scheduled task creation thread started")
                    else:
                        print("DEBUG: No result data found")
                        
                elif action == 'delete':
                    task_name = getattr(dlg, 'task_name', None)
                    print(f"DEBUG: Deleting task: {task_name}")
                    if task_name:
                        try:
                            self.delete_scheduled_task(task_name)
                        except Exception as delete_error:
                            print(f"DEBUG: Task deletion failed: {delete_error}")
                            error_msg = f"Failed to delete scheduled task: {str(delete_error)}"
                            self.message_queue.put(('log', error_msg))
                            try:
                                Win11Dialog.showerror(self.translator.get('task_deletion_failed'), error_msg,
                                                    parent=self, theme=self.theme, translator=self.translator)
                            except Exception:
                                print(f"ERROR: {error_msg}")
                else:
                    print(f"DEBUG: Action '{action}' - no operation needed")
            else:
                print("DEBUG: Dialog has no action attribute or is None")
                
        except Exception as e:
            print(f"DEBUG: Exception in schedule_backup: {type(e).__name__}: {e}")
            import traceback
            print(f"DEBUG: Traceback: {traceback.format_exc()}")
            
            error_msg = f"Schedule dialog error: {str(e)}"
            self.message_queue.put(('log', error_msg))
            
            # Handle the messagebox call safely as wel
            try:
                Win11Dialog.showerror(self.translator.get('error'), error_msg,
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception as mb_error:
                print(f"DEBUG: messagebox.showerror failed: {mb_error}")


    ############# Schedule Backup Functions #############

    def _get_schedule_config_dir(self):
        """Get the directory path for schedule configuration files"""
        appdata = os.environ.get('APPDATA', '')
        if not appdata:
            raise RuntimeError("APPDATA environment variable not found")
        config_dir = os.path.join(appdata, 'ezBAK', 'ezBAK_Schedule')
        return config_dir

    def _save_schedule_config(self, task_name, user, dest, schedule, time_str, monthly_day=1,
                             include_hidden=False, include_system=False,
                             retention_count=2, log_retention_days=30, filters=None):
        """
        Save schedule configuration to a JSON file in %appdata%\ezBAK\ezBAK_Schedule\
        Returns the path to the saved configuration file
        """
        try:
            # Get config directory
            config_dir = self._get_schedule_config_dir()

            # Create directory if it doesn't exist
            os.makedirs(config_dir, exist_ok=True)

            # Create config file path (use task name as filename)
            # Sanitize task name for filename
            safe_task_name = re.sub(r'[<>:"/\\|?*]', '_', task_name)
            config_file = os.path.join(config_dir, f"{safe_task_name}.json")

            # Prepare configuration data
            config_data = {
                'task_name': task_name,
                'user': user,
                'dest': dest,
                'schedule': schedule,
                'time_str': time_str,
                'monthly_day': monthly_day,
                'include_hidden': include_hidden,
                'include_system': include_system,
                'retention_count': retention_count,
                'log_retention_days': log_retention_days,
                'filters': filters if filters else {'include': [], 'exclude': [{'type': 'name', 'pattern': 'onedrive*'}]},
                'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            # Save to JSON file
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)

            self.write_detailed_log(f"Schedule configuration saved to: {config_file}")
            return config_file

        except Exception as e:
            error_msg = f"Failed to save schedule configuration: {str(e)}"
            self.write_detailed_log(error_msg)
            raise RuntimeError(error_msg) from e

    def _load_schedule_config(self, config_file):
        """
        Load schedule configuration from a JSON file
        Returns a dictionary with configuration data
        """
        try:
            if not os.path.exists(config_file):
                raise FileNotFoundError(f"Configuration file not found: {config_file}")

            with open(config_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            return config_data

        except Exception as e:
            error_msg = f"Failed to load schedule configuration from {config_file}: {str(e)}"
            raise RuntimeError(error_msg) from e

    def _build_auto_backup_command(self, user, dest, include_hidden=False, include_system=False):
        # Build the command string used by Task Scheduler (/TR argument)
        try:
            if getattr(sys, 'frozen', False):
                exe = sys.executable
                cmd = f'"{exe}" --auto-backup --user "{user}" --dest "{dest}"'
            else:
                exe = sys.executable
                script = os.path.abspath(__file__)
                cmd = f'"{exe}" "{script}" --auto-backup --user "{user}" --dest "{dest}"'
            if include_hidden:
                cmd += ' --include-hidden'
            if include_system:
                cmd += ' --include-system'
            return cmd
        except Exception:
            return ''

    def _create_task_xml(self, task_name, executable, arguments, formatted_time, sc_type, run_as_user, monthly_day=1):
        """
        Create XML file for scheduled task to avoid 261-character limit of /tr parameter

        Args:
            task_name: Name of the scheduled task
            executable: Path to the executable (e.g., python.exe)
            arguments: Command-line arguments for the executable
            formatted_time: Time in HH:MM format
            sc_type: Schedule type (DAILY, WEEKLY, MONTHLY)
            run_as_user: User account to run the task
            monthly_day: Day of month for monthly schedules (default: 1)

        Returns path to the temporary XML file
        """
        import tempfile
        import xml.etree.ElementTree as ET
        from datetime import datetime, timedelta

        # Parse the time
        time_parts = formatted_time.split(':')
        hour = time_parts[0]
        minute = time_parts[1]

        # Create XML structure for Windows Task Scheduler
        # Get start boundary (tomorrow at the specified time to avoid immediate execution)
        start_date = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')
        start_boundary = f"{start_date}T{formatted_time}:00"

        # Build the XML structure
        xml_template = f'''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>{datetime.now().isoformat()}</Date>
    <Author>{run_as_user}</Author>
    <Description>ezBAK Automated Backup Task</Description>
  </RegistrationInfo>
  <Triggers>
'''

        # Add appropriate trigger based on schedule type
        if sc_type == "DAILY":
            xml_template += f'''    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
'''
        elif sc_type == "WEEKLY":
            xml_template += f'''    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Sunday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
'''
        elif sc_type == "MONTHLY":
            xml_template += f'''    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth>
          <Day>{monthly_day}</Day>
        </DaysOfMonth>
        <Months>
          <January />
          <February />
          <March />
          <April />
          <May />
          <June />
          <July />
          <August />
          <September />
          <October />
          <November />
          <December />
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
'''

        xml_template += f'''  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{run_as_user}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{executable}</Command>
      <Arguments>{arguments}</Arguments>
    </Exec>
  </Actions>
</Task>'''

        # Create temporary XML file
        try:
            # Create temp file in the schedules directory for better organization
            schedules_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'ezBAK', 'schedules')
            os.makedirs(schedules_dir, exist_ok=True)

            xml_file = os.path.join(schedules_dir, f"{task_name}_task.xml")

            # Write XML file with UTF-16 encoding (required by Windows Task Scheduler)
            with open(xml_file, 'w', encoding='utf-16') as f:
                f.write(xml_template)

            self.write_detailed_log(f"Task XML file created: {xml_file}")
            return xml_file

        except Exception as e:
            error_msg = f"Failed to create task XML file: {str(e)}"
            self.write_detailed_log(error_msg)
            raise RuntimeError(error_msg) from e

    def create_scheduled_task(self, task_name, schedule, time_str, user, dest, monthly_day=1, include_hidden=False, include_system=False, retention_count=2, log_retention_days=30, filters=None):
        """
        Create a scheduled task (set to run with administrator privileges)
        """
        import json
        print(f"DEBUG: create_scheduled_task called with task_name={task_name}, schedule={schedule}, time_str={time_str}")
        try:
            # Time format validation and normalization
            from datetime import datetime
            try:
                time_obj = datetime.strptime(time_str.strip(), "%H:%M")
                formatted_time = time_obj.strftime("%H:%M")
                print(f"DEBUG: Time format validated: {time_str} -> {formatted_time}")
                self.write_detailed_log(f"Time format validated: {time_str} -> {formatted_time}")
            except ValueError as ve:
                raise ValueError(f"Invalid time format '{time_str}'. Please use HH:MM format (e.g., 14:30)")

            # Schedule type mapping and validation
            schedule_lower = schedule.lower().strip()
            if schedule_lower == "daily":
                sc_type = "DAILY"
            elif schedule_lower == "weekly":
                sc_type = "WEEKLY"
            elif schedule_lower == "monthly":
                sc_type = "MONTHLY"
            else:
                raise ValueError(f"Unsupported schedule type: {schedule}. Use Daily, Weekly, or Monthly")

            # Check executable file path
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
                script_path = ""
            else:
                exe_path = sys.executable
                script_path = os.path.abspath(__file__)

            # Prepare filters if not provided
            if filters is None:
                filters = {'include': [], 'exclude': [{'type': 'name', 'pattern': 'onedrive*'}]}

            # Save schedule configuration to file
            config_file = self._save_schedule_config(
                task_name=task_name,
                user=user,
                dest=dest,
                schedule=schedule,
                time_str=time_str,
                monthly_day=monthly_day,
                include_hidden=include_hidden,
                include_system=include_system,
                retention_count=retention_count,
                log_retention_days=log_retention_days,
                filters=filters
            )

            # Configure command to use schedule config file
            # Separate executable and arguments for proper Windows Task Scheduler format
            executable = exe_path
            if script_path:
                arguments = f'"{script_path}" --schedule-config "{config_file}"'
            else:
                arguments = f'--schedule-config "{config_file}"'

            self.write_detailed_log(f"Task executable: {executable}")
            self.write_detailed_log(f"Task arguments: {arguments}")
            self.write_detailed_log(f"Schedule config saved to: {config_file}")
            self.write_detailed_log(f"Creating scheduled task with retention_count={retention_count}, log_retention_days={log_retention_days}")

            # Get current user info
            current_user = os.environ.get('USERNAME', 'Administrator')
            domain = os.environ.get('USERDOMAIN', os.environ.get('COMPUTERNAME', ''))

            # Determine user account format
            if domain and domain.upper() != current_user.upper():
                run_as_user = f"{domain}\\{current_user}"
            else:
                run_as_user = current_user

            # Create XML file for task (avoids 261-character limit of /tr parameter)
            xml_file = self._create_task_xml(
                task_name=task_name,
                executable=executable,
                arguments=arguments,
                formatted_time=formatted_time,
                sc_type=sc_type,
                run_as_user=run_as_user,
                monthly_day=monthly_day
            )

            # Configure schtasks command using XML file
            cmd_args = [
                "schtasks", "/create",
                "/tn", task_name,
                "/xml", xml_file,
                "/f"  # Overwrite if existing task exists
            ]

            # Log command to be executed
            self.write_detailed_log(f"schtasks command (XML-based): {' '.join(cmd_args)}")
            self.write_detailed_log(f"Task will run as: {run_as_user} with HIGHEST privileges")
            self.write_detailed_log(f"Using XML file to avoid command length limitations")
            
            # Execute command
            # On Windows, use system default encoding (e.g., cp949 on Korean Windows)
            # to avoid encoding errors with special characters
            import locale
            system_encoding = locale.getpreferredencoding()

            try:
                print(f"DEBUG: Executing schtasks command: {' '.join(cmd_args)}")
                result = subprocess.run(
                    cmd_args,
                    check=False,
                    shell=False,
                    capture_output=True,
                    text=True,
                    encoding=system_encoding,
                    errors='replace',  # Replace problematic characters instead of ignoring
                    timeout=180  # Increased from 90 to 180 seconds to prevent timeouts on busy systems
                )

                print(f"DEBUG: schtasks command completed with returncode={result.returncode}")
                # Log results
                self.write_detailed_log(f"schtasks return code: {result.returncode}")
                if result.stdout:
                    self.write_detailed_log(f"schtasks stdout: {result.stdout}")
                if result.stderr:
                    self.write_detailed_log(f"schtasks stderr: {result.stderr}")
                    
                # Check success
                if result.returncode == 0:
                    print(f"DEBUG: schtasks command succeeded (returncode=0)")
                    success_msg = (
                        f"Scheduled task '{task_name}' created successfully!\n\n"
                        f"Schedule: {schedule} at {formatted_time}\n"
                        f"User: {user}\n"
                        f"Destination: {dest}\n"
                        f"Retention: {retention_count} backups, {log_retention_days} days logs\n"
                        f"Hidden files: {'Included' if include_hidden else 'Excluded'}\n"
                        f"System files: {'Included' if include_system else 'Excluded'}\n\n"
                        f"Run Level: Administrator (Highest Privileges)\n"
                        f"Run As: {run_as_user}"
                    )

                    self.write_detailed_log(f"Task creation successful: {task_name} (with admin privileges)")
                    self.message_queue.put(('log', f"Scheduled task '{task_name}' created with administrator privileges"))

                    print(f"DEBUG: Returning success message from create_scheduled_task")
                    # Return success message instead of showing dialog here (dialog will be shown from main thread)
                    return success_msg
                    
                else:
                    # Detailed error analysis on failure
                    error_details = self._analyze_schtasks_error(result.returncode, result.stderr, result.stdout)
                    raise RuntimeError(error_details)
                    
            except subprocess.TimeoutExpired:
                error_msg = "Task creation timed out. The system may be busy."
                self.write_detailed_log(f"schtasks command timed out")
                raise RuntimeError(error_msg)
                
        except Exception as e:
            error_msg = f"Failed to create scheduled task '{task_name}': {str(e)}"
            self.write_detailed_log(f"Task creation failed: {e}")
            
            # Log detailed error information
            import traceback
            self.write_detailed_log(f"Full error traceback: {traceback.format_exc()}")
            
            print(f"ERROR: {error_msg}")
            raise RuntimeError(error_msg) from e

    def run_create_scheduled_task_thread(self, task_data):
        """
        Thread wrapper for create_scheduled_task to prevent UI blocking
        """
        try:
            print(f"DEBUG: Thread started with task_name={task_data.get('task_name')}")
            self.message_queue.put(('log', f"Creating scheduled task '{task_data.get('task_name')}'..."))

            print(f"DEBUG: Calling create_scheduled_task...")
            # Call the actual task creation method
            result = self.create_scheduled_task(
                task_name=task_data.get('task_name'),
                dest=task_data.get('dest'),
                schedule=task_data.get('schedule'),
                time_str=task_data.get('time_str'),
                monthly_day=task_data.get('monthly_day', 1),
                include_hidden=task_data.get('include_hidden', False),
                include_system=task_data.get('include_system', False),
                retention_count=int(task_data.get('retention_count', 2)),
                log_retention_days=int(task_data.get('log_retention_days', 30)),
                user=task_data.get('user'),
                filters=task_data.get('filters', {'include': [], 'exclude': [{'type': 'name', 'pattern': 'onedrive*'}]})
            )

            print(f"DEBUG: create_scheduled_task returned: {result}")
            self.message_queue.put(('log', f"Scheduled task '{task_data.get('task_name')}' created successfully"))

            # Show success dialog in main thread (result is the success message)
            if result:
                print(f"DEBUG: Scheduling success dialog in main thread")
                self.after(100, lambda msg=result: self._show_task_creation_success(msg))
            else:
                print(f"DEBUG: No result returned from create_scheduled_task")

        except Exception as task_error:
            # Handle errors in the thread
            print(f"DEBUG: Exception in thread: {task_error}")
            error_msg = f"Failed to create scheduled task: {str(task_error)}"
            self.write_detailed_log(f"Thread error: {error_msg}")
            self.message_queue.put(('log', error_msg))

            # Schedule error dialog to show in main thread
            print(f"DEBUG: Scheduling error dialog in main thread")
            self.after(100, lambda: self._show_task_creation_error(error_msg))
        finally:
            # Re-enable buttons in main thread
            print(f"DEBUG: Re-enabling buttons")
            self.after(100, lambda: self.set_buttons_state("normal"))

    def _show_task_creation_success(self, success_msg):
        """Show task creation success in main thread"""
        try:
            Win11Dialog.showinfo(
                self.translator.get('schedule_created'),
                success_msg,
                parent=self,
                theme=self.theme,
                translator=self.translator
            )
        except Exception as dlg_error:
            self.write_detailed_log(f"Error showing success dialog: {dlg_error}")
            print(f"Task created successfully: {success_msg}")

    def _show_task_creation_error(self, error_msg):
        """Show task creation error in main thread"""
        try:
            Win11Dialog.showerror(
                self.translator.get('task_creation_failed'),
                error_msg,
                parent=self,
                theme=self.theme,
                translator=self.translator
            )
        except Exception as dlg_error:
            self.write_detailed_log(f"Error showing dialog: {dlg_error}")
            print(f"ERROR: {error_msg}")

    def _analyze_schtasks_error(self, return_code, stderr, stdout):
        """Analyze schtasks error code and provide solutions (including admin privileges)"""
        error_msg = f"schtasks failed with return code {return_code}"
        
        # Common error codes and solutions
        error_solutions = {
            1: "General error. Check if task name is valid and you have admin privileges.",
            2: "Access denied. Try running as administrator.",
            267: "Invalid task name or path. Avoid special characters in task name.",
            2147942487: "Access denied. Administrator privileges required.",
            2147943711: "Object already exists. Task with this name already exists.",
            -2147024809: "Invalid parameter. Check task settings.",
            -2147216609: "Task Scheduler service is not available.",
        }
        
        # Search for specific error patterns in stderr (including admin privileges)
        if stderr:
            stderr_lower = stderr.lower()
            if "access" in stderr_lower and "denied" in stderr_lower:
                error_msg += "\n\nSolution: Run the application as Administrator"
            elif "invalid" in stderr_lower and "time" in stderr_lower:
                error_msg += "\n\nSolution: Check time format (use HH:MM, e.g., 14:30)"
            elif "task" in stderr_lower and "exist" in stderr_lower:
                error_msg += "\n\nSolution: Choose a different task name or delete existing task"
            elif "privilege" in stderr_lower or "permission" in stderr_lower:
                error_msg += "\n\nSolution: Administrator privileges required for task creation"
            elif "user" in stderr_lower and ("not found" in stderr_lower or "invalid" in stderr_lower):
                error_msg += "\n\nSolution: Check user account format (try DOMAIN\\Username or just Username)"
            elif "service" in stderr_lower and "not" in stderr_lower:
                error_msg += "\n\nSolution: Task Scheduler service may not be running. Try: net start schedule"
        
        # Add solutions for known error codes
        if return_code in error_solutions:
            error_msg += f"\n\n{error_solutions[return_code]}"
        
        # Additional guidance for admin privileges
        if return_code in [2, 2147942487] or "access" in (stderr or "").lower():
            error_msg += (
                "\n\nAdditional Info:"
                "\n- Make sure ezBAK is running as Administrator"
                "\n- The scheduled task will run with HIGHEST privilege level"
                "\n- Check Windows Task Scheduler service is running"
            )
        
        # Add stderr and stdout contents
        if stderr:
            error_msg += f"\n\nSystem Error: {stderr}"
        if stdout:
            error_msg += f"\n\nSystem Output: {stdout}"
            
        return error_msg

    # Improved time validation for ScheduleBackupDialog
    def _validate_time_input(self, time_str):
        """Validate time input (used in ScheduleBackupDialog)"""
        try:
            from datetime import datetime
            time_obj = datetime.strptime(time_str.strip(), "%H:%M")
            return True, time_obj.strftime("%H:%M")
        except ValueError:
            return False, "Invalid time format. Please use HH:MM (e.g., 14:30 for 2:30 PM)"

    def delete_scheduled_task(self, task_name):
        if not task_name:
            try:
                Win11Dialog.showwarning(self.translator.get('warning'), self.translator.get('task_name_required'),
                                      parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass
            return
        try:
            res = subprocess.run(["schtasks", "/Delete", "/TN", task_name, "/F"], shell=True, capture_output=True, text=True)
            if res.returncode != 0:
                raise RuntimeError(res.stderr or res.stdout)
        except Exception as e:
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('failed_to_delete_task')}:\n{e}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass
            self.message_queue.put(('log', f"Schedule delete failed: {e}"))
            return

        # Delete the schedule configuration file
        try:
            config_dir = self._get_schedule_config_dir()
            safe_task_name = re.sub(r'[<>:"/\\|?*]', '_', task_name)
            config_file = os.path.join(config_dir, f"{safe_task_name}.json")

            if os.path.exists(config_file):
                os.remove(config_file)
                self.write_detailed_log(f"Schedule configuration file deleted: {config_file}")
                self.message_queue.put(('log', f"Config file deleted: {config_file}"))
        except Exception as config_error:
            self.write_detailed_log(f"Failed to delete schedule config file: {config_error}")
            # Don't fail the whole operation if config file deletion fails

        try:
            Win11Dialog.showinfo(self.translator.get('task_scheduler'), f"{self.translator.get('task_deleted')}: '{task_name}'",
                               parent=self, theme=self.theme, translator=self.translator)
        except Exception:
            pass
        self.message_queue.put(('log', f"Task deleted: {task_name}"))

    def check_space(self):
        """
        Execute space check in separate thread (improves UI responsiveness)
        """
        # 1) Select sources via checkbox tree
        try:
            dlg = SelectSourcesDialog(self, is_hidden_fn=self.is_hidden, title=self.translator.get('check_required_space'))
            self.wait_window(dlg)
            sources = dlg.selected_paths if getattr(dlg, 'selected_paths', None) else []
        except Exception as e:
            Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('source_selection_failed')}: {e}",
                                parent=self, theme=self.theme, translator=self.translator)
            return
        
        if not sources:
            self.message_queue.put(('log', "Space check cancelled (no sources)."))
            return

        # 2) Select destination folder
        dest = filedialog.askdirectory(title=self.translator.get('select_destination_folder'))
        if not dest:
            self.message_queue.put(('log', "Space check cancelled (no destination)."))
            return

        # 3) Execute space calculation in separate thread
        self.message_queue.put(('log', "Calculating required space..."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Calculating space..."))

        calc_thread = threading.Thread(target=self.run_space_calculation, args=(sources, dest), daemon=True)
        calc_thread.start()
        
    def run_space_calculation(self, sources, dest):
        """Execute space calculation in separate thread"""
        try:
            total = 0
            file_count = 0
            
            self.message_queue.put(('update_status', "Calculating required space..."))
            
            for i, source_path in enumerate(sources):
                self.message_queue.put(('update_status', f"Checking {i+1}/{len(sources)}: {os.path.basename(source_path)}"))
                
                if os.path.isdir(source_path):
                    # Calculate directory size
                    dir_total = 0
                    dir_files = 0
                    
                    for dirpath, dirnames, filenames in os.walk(source_path, topdown=True, onerror=lambda e: None):
                        # Skip hidden directories
                        dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                        
                        for f in filenames:
                            fp = os.path.join(dirpath, f)
                            if self.is_hidden(fp):
                                continue
                                
                            try:
                                size = os.path.getsize(fp)
                                dir_total += size
                                dir_files += 1
                                
                                # Update UI every 100 files
                                if dir_files % 100 == 0:
                                    self.message_queue.put(('update_status', 
                                        f"Checking {os.path.basename(source_path)}: {dir_files} files, {self.format_bytes(dir_total)}"))
                            except (OSError, IOError):
                                continue
                    
                    total += dir_total
                    file_count += dir_files
                    
                elif os.path.isfile(source_path) and not self.is_hidden(source_path):
                    # Single file
                    try:
                        size = os.path.getsize(source_path)
                        total += size
                        file_count += 1
                    except (OSError, IOError):
                        continue
            
            # Calculate free space
            self.message_queue.put(('update_status', "Checking available space..."))
            free = self.get_free_space_reliable(dest)
            
            # Display result
            self.message_queue.put(('space_check_result', {
                'required': total,
                'available': free,
                'file_count': file_count,
                'sources_count': len(sources)
            }))
            
        except Exception as e:
            self.message_queue.put(('show_error', f"Space calculation error: {e}"))
        finally:
            # Fixed: progress bar
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('enable_buttons', None))

    def open_filter_manager(self):
        try:
            dlg = FilterManagerDialog(self, current_filters=self.filters)
            self.wait_window(dlg)
            if getattr(dlg, 'result', None):
                self.filters = dlg.result
                # persist immediately
                try:
                    self.save_settings()
                except Exception:
                    pass
                self.message_queue.put(('log', 'Filters updated.'))
        except Exception as e:
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('filter_manager_error')}: {e}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass

    def get_used_drive_letters(self):
        """Returns a set of used drive letters (uppercase, without colon)."""
        used = set()
        try:
            bitmask = ctypes.windll.kernel32.GetLogicalDrives()
            for i in range(26):
                if bitmask & (1 << i):
                    used.add(chr(65 + i))  # 'A'..'Z'
        except Exception:
            # Fallback probe
            for i in range(65, 91):
                letter = chr(i)
                try:
                    if os.path.exists(f"{letter}:\\"):
                        used.add(letter)
                except Exception:
                    continue
        return used

    def find_free_drive_letter(self):
        """Finds a free drive letter, preferring Z backward to D (avoid A,B, system)."""
        used = self.get_used_drive_letters()
        for i in range(90, 67, -1):  # Z..D
            letter = chr(i)
            if letter not in used:
                return letter
        # If none found, fall back to C range (unlikely)
        for i in range(67, 65, -1):
            letter = chr(i)
            if letter not in used:
                return letter
        return 'Z'

    def open_nas_connect_dialog(self):
        """Open a large dialog for NAS connect with full command examples, then run 'net use'."""
        # Build dialog UI
        dlg = tk.Toplevel(self)
        dlg.title(self.translator.get('network_share_connect'))
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg=self.theme.get('bg_elevated'))
        dlg.geometry("680x301")
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass

        wrap = 640
        help_text = (
            "Connect: net use [<drive_letter>:] \\Server\\share [/user:id pwd] /persistent:{yes|no}\n\n"
            "Disconnect: net use [<drive_letter>:|\\Server\\share] /delete /y"
        )
        tk.Label(dlg, text=help_text, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), justify='left', anchor='w', font=("Consolas", 10), wraplength=wrap).pack(fill='x', padx=12, pady=(12, 8))

        form = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        form.pack(fill='both', expand=True, padx=16)

        row = 0
        def add_label(text):
            nonlocal row
            tk.Label(form, text=text, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg')).grid(row=row, column=0, sticky='w', pady=6)
        def add_entry(var, width=42, show=None):
            nonlocal row
            e = tk.Entry(form, textvariable=var, width=width, show=show, bg=self.theme.get('bg'), fg=self.theme.get('fg'), insertbackground=self.theme.get('fg'))
            e.grid(row=row, column=1, sticky='w', pady=6)
            row += 1
            return e

        # UNC Path
        add_label(self.translator.get('unc_path_label'))
        unc_var = tk.StringVar()
        add_entry(unc_var)

        # Map drive option
        map_var = tk.BooleanVar(value=True)
        def_drive = self.find_free_drive_letter() or 'Z'
        drive_var = tk.StringVar(value=def_drive)

        map_frame = tk.Frame(form, bg=self.theme.get('bg_elevated'))
        tk.Checkbutton(map_frame, text=self.translator.get('map_as_drive_letter'), variable=map_var, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), selectcolor=self.theme.get('bg_elevated'), activebackground=self.theme.get('bg_elevated'), activeforeground=self.theme.get('fg')).pack(side='left')
        tk.Label(map_frame, text=self.translator.get('letter'), bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg')).pack(side='left', padx=(12,4))
        drive_entry = tk.Entry(map_frame, textvariable=drive_var, width=6, bg=self.theme.get('bg'), fg=self.theme.get('fg'), insertbackground=self.theme.get('fg'))
        drive_entry.pack(side='left')
        map_frame.grid(row=row, column=0, columnspan=2, sticky='w', pady=6)
        row += 1

        def toggle_drive():
            try:
                state = 'normal' if map_var.get() else 'disabled'
                drive_entry.config(state=state)
            except Exception:
                pass
        toggle_drive()
        try:
            map_frame.children
            # Bind after creation
            map_frame.bind('<Map>', lambda e: toggle_drive())
        except Exception:
            pass
        tk.Checkbutton(map_frame, variable=map_var).destroy()  # no-op cleanup (ensures lint calm)

        # Credentials
        add_label(self.translator.get('username_optional'))
        user_var = tk.StringVar()
        add_entry(user_var)

        add_label(self.translator.get('password_required_if_username'))
        pwd_var = tk.StringVar()
        add_entry(pwd_var, show='*')

        # Persistent
        persist_var = tk.BooleanVar(value=True)
        persist_chk = tk.Checkbutton(form, text=self.translator.get('reconnect_at_logon_persistent'), variable=persist_var, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), selectcolor=self.theme.get('bg_elevated'), activebackground=self.theme.get('bg_elevated'), activeforeground=self.theme.get('fg'))
        persist_chk.grid(row=row, column=0, columnspan=2, sticky='w', pady=8)
        row += 1

        # Buttons
        btns = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        btns.pack(fill='x', pady=(6,12))

        result = {'ok': False}
        def on_ok():
            unc = (unc_var.get() or '').strip()
            if not unc or not unc.startswith('\\\\'):
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('unc_path_required'), parent=dlg, theme=self.theme, translator=self.translator)
                return
            user = (user_var.get() or '').strip()
            pwd = (pwd_var.get() or '')
            if user and not pwd:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('password_required'), parent=dlg, theme=self.theme, translator=self.translator)
                return
            drv = None
            if map_var.get():
                d = (drive_var.get() or '').strip().rstrip(':').upper()
                if not (len(d) == 1 and 'A' <= d <= 'Z'):
                    Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_drive_letter'), parent=dlg, theme=self.theme, translator=self.translator)
                    return
                drv = d + ':'
            result.update(ok=True, unc=unc, drive=drv, user=user, pwd=pwd, persistent=persist_var.get())
            dlg.destroy()
        def on_cancel():
            dlg.destroy()

        # Windows 11 style buttons
        ok_btn = Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_ok'), on_ok, self.theme, is_primary=True)
        ok_btn.pack(side='right', padx=8)
        cancel_btn = Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_cancel'), on_cancel, self.theme, is_primary=False)
        cancel_btn.pack(side='right', padx=8)

        dlg.bind('<Return>', lambda e: on_ok())
        dlg.bind('<Escape>', lambda e: on_cancel())
        dlg.update_idletasks()
        try:
            # Center on parent
            px = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
            py = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
            dlg.geometry(f"+{max(px,0)}+{max(py,0)}")
        except Exception:
            pass
        self.wait_window(dlg)

        if not result.get('ok'):
            self.message_queue.put(('log', "NAS connect cancelled."))
            return

        # Execute connection
        unc = result['unc']
        drive = result['drive']
        username = result['user']
        password = result['pwd']
        persistent = result['persistent']

        try:
            cmd = ['net', 'use']
            if drive:
                cmd += [drive, unc]
            else:
                cmd += [unc]
            if username:
                cmd += [f'/user:{username}', password]
            cmd += [f"/persistent:{'yes' if persistent else 'no'}"]

            self.message_queue.put(('log', f"Connecting to {unc}..."))
            try:
                res = subprocess.run(cmd, shell=False, capture_output=True, text=True, encoding='cp949', errors='ignore')
            except Exception:
                res = subprocess.run(cmd, shell=False, capture_output=True, text=True)

            out = (res.stdout or '').strip()
            err = (res.stderr or '').strip()
            if out:
                try:
                    self.write_detailed_log(out)
                except Exception:
                    pass
            if err:
                try:
                    self.write_detailed_log(err)
                except Exception:
                    pass

            if res.returncode == 0:
                self.message_queue.put(('log', f"NAS connected: {drive or unc}"))
                try:
                    if Win11Dialog.askyesno(self.translator.get('connected'), self.translator.get('open_share_now'), parent=self, theme=self.theme, translator=self.translator):
                        os.startfile(drive if drive else unc)
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS connect failed: {msg}"))
                try:
                    Win11Dialog.showerror(self.translator.get('nas_connect'), self.translator.get('failed_to_connect').format(msg), parent=self, theme=self.theme, translator=self.translator)
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS connect error: {e}"))
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"NAS connect error: {e}", parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass
            if not unc:
                self.message_queue.put(('log', "NAS connect cancelled (no path)."))
                return
            if not unc.startswith('\\\\'):
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('unc_path_required'), parent=self, theme=self.theme, translator=self.translator)
                return

            map_drive = Win11Dialog.askyesno(self.translator.get('map_drive'), self.translator.get('map_to_drive_letter'), parent=self, theme=self.theme, translator=self.translator)
            drive = None
            if map_drive:
                default_drive = self.find_free_drive_letter() or 'Z'
                val = simpledialog.askstring("Drive Letter", f"Drive letter (A-Z). Default: {default_drive}", parent=self)
                if not val:
                    val = default_drive
                val = (val.strip().rstrip(':').upper() if isinstance(val, str) else str(val)).rstrip(':')
                if not (len(val) == 1 and 'A' <= val <= 'Z'):
                    Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_drive_letter'), parent=self, theme=self.theme, translator=self.translator)
                    return
                drive = val + ':'

            username = simpledialog.askstring("Username", "Username (leave empty to use current credentials):", parent=self)
            password = None
            if username:
                password = simpledialog.askstring("Password", "Password (required when username is set)", show='*', parent=self)
                if not password:
                    Win11Dialog.showerror(self.translator.get('error'), self.translator.get('password_required'), parent=self, theme=self.theme, translator=self.translator)
                    return

            persistent = Win11Dialog.askyesno(self.translator.get('persistence'), self.translator.get('reconnect_at_signin'), parent=self, theme=self.theme, translator=self.translator)

            # Build command
            cmd = ['net', 'use']
            if drive:
                cmd += [drive, unc]
            else:
                cmd += [unc]
            if username:
                cmd += [f'/user:{username}', password]
            cmd += [f"/persistent:{'yes' if persistent else 'no'}"]

            self.message_queue.put(('log', f"Connecting to {unc}..."))
            try:
                res = subprocess.run(cmd, shell=False, capture_output=True, text=True, encoding='cp949', errors='ignore')
            except Exception:
                res = subprocess.run(cmd, shell=False, capture_output=True, text=True)

            out = (res.stdout or '').strip()
            err = (res.stderr or '').strip()
            if out:
                try:
                    self.write_detailed_log(out)
                except Exception:
                    pass
            if err:
                try:
                    self.write_detailed_log(err)
                except Exception:
                    pass

            if res.returncode == 0:
                self.message_queue.put(('log', f"NAS connected: {drive or unc}"))
                try:
                    if Win11Dialog.askyesno(self.translator.get('connected'), self.translator.get('open_share_now'), parent=self, theme=self.theme, translator=self.translator):
                        os.startfile(drive if drive else unc)
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS connect failed: {msg}"))
                try:
                    Win11Dialog.showerror(self.translator.get('nas_connect'), self.translator.get('failed_to_connect').format(msg), parent=self, theme=self.theme, translator=self.translator)
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS connect error: {e}"))
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"NAS connect error: {e}", parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass

    def open_nas_disconnect_dialog(self):
        """Open a large dialog for NAS disconnect with full command examples, then run 'net use ... /delete'."""
        # Build dialog UI
        dlg = tk.Toplevel(self)
        dlg.title(self.translator.get('network_share_disconnect'))
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg=self.theme.get('bg_elevated'))
        dlg.geometry("680x178")
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass

        wrap = 640
        help_text = (
            "Connect: net use [<drive_letter>:] \\\\server\\share [/user:username password] /persistent:{yes|no}\n\n"
            "Disconnect: net use [drive_letter:|\\\\server\\share] /delete /y"
        )
        tk.Label(dlg, text=help_text, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), justify='left', anchor='w', font=("Consolas", 10), wraplength=wrap).pack(fill='x', padx=12, pady=(12, 8))

        form = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        form.pack(fill='both', expand=True, padx=16)

        tk.Label(form, text=self.translator.get('target_drive_or_unc'), bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg')).grid(row=0, column=0, sticky='w', pady=6)
        target_var = tk.StringVar()
        tk.Entry(form, textvariable=target_var, width=42, bg=self.theme.get('bg'), fg=self.theme.get('fg'), insertbackground=self.theme.get('fg')).grid(row=0, column=1, sticky='w', pady=6)

        btns = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        btns.pack(fill='x', pady=(6,12))

        result = {'ok': False}
        def on_ok():
            val = (target_var.get() or '').strip()
            if not val:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('error'), parent=dlg, theme=self.theme, translator=self.translator)
                return
            result.update(ok=True, target=val)
            dlg.destroy()
        def on_cancel():
            dlg.destroy()

        # Windows 11 style buttons
        ok_btn = Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_ok'), on_ok, self.theme, is_primary=True)
        ok_btn.pack(side='right', padx=8)
        cancel_btn = Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_cancel'), on_cancel, self.theme, is_primary=False)
        cancel_btn.pack(side='right', padx=8)

        dlg.bind('<Return>', lambda e: on_ok())
        dlg.bind('<Escape>', lambda e: on_cancel())
        dlg.update_idletasks()
        try:
            px = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
            py = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
            dlg.geometry(f"+{max(px,0)}+{max(py,0)}")
        except Exception:
            pass
        self.wait_window(dlg)

        if not result.get('ok'):
            self.message_queue.put(('log', "NAS disconnect cancelled."))
            return

        target = result['target']
        try:
            if target.startswith('\\\\'):
                args = ['net', 'use', target, '/delete', '/y']
                label = target
            else:
                drv = target.rstrip(':').upper() + ':'
                if not (len(drv) == 2 and 'A' <= drv[0] <= 'Z' and drv[1] == ':'):
                    Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_drive_letter'), parent=self, theme=self.theme, translator=self.translator)
                    return
                args = ['net', 'use', drv, '/delete', '/y']
                label = drv

            self.message_queue.put(('log', f"Disconnecting {label}..."))
            try:
                res = subprocess.run(args, shell=False, capture_output=True, text=True, encoding='cp949', errors='ignore')
            except Exception:
                res = subprocess.run(args, shell=False, capture_output=True, text=True)
            out = (res.stdout or '').strip()
            err = (res.stderr or '').strip()
            if res.returncode == 0:
                self.message_queue.put(('log', f"NAS disconnected: {label}"))
                try:
                    Win11Dialog.showinfo(self.translator.get('nas_disconnect'), self.translator.get('disconnected').format(label), parent=self, theme=self.theme, translator=self.translator)
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS disconnect failed: {msg}"))
                try:
                    Win11Dialog.showerror(self.translator.get('nas_disconnect'), self.translator.get('failed_to_disconnect').format(msg), parent=self, theme=self.theme, translator=self.translator)
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS disconnect error: {e}"))
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"NAS disconnect error: {e}", parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass
            if not target:
                self.message_queue.put(('log', "NAS disconnect cancelled (no target)."))
                return
            target = target.strip()
            if target.startswith('\\\\'):
                args = ['net', 'use', target, '/delete', '/y']
                label = target
            else:
                drv = target.rstrip(':').upper() + ':'
                if not (len(drv) == 2 and 'A' <= drv[0] <= 'Z' and drv[1] == ':'):
                    Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_drive_letter'), parent=self, theme=self.theme, translator=self.translator)
                    return
                args = ['net', 'use', drv, '/delete', '/y']
                label = drv

            self.message_queue.put(('log', f"Disconnecting {label}..."))
            try:
                res = subprocess.run(args, shell=False, capture_output=True, text=True, encoding='cp949', errors='ignore')
            except Exception:
                res = subprocess.run(args, shell=False, capture_output=True, text=True)
            out = (res.stdout or '').strip()
            err = (res.stderr or '').strip()
            if res.returncode == 0:
                self.message_queue.put(('log', f"NAS disconnected: {label}"))
                try:
                    Win11Dialog.showinfo(self.translator.get('nas_disconnect'), self.translator.get('disconnected').format(label), parent=self, theme=self.theme, translator=self.translator)
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS disconnect failed: {msg}"))
                try:
                    Win11Dialog.showerror(self.translator.get('nas_disconnect'), self.translator.get('failed_to_disconnect').format(msg), parent=self, theme=self.theme, translator=self.translator)
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS disconnect error: {e}"))
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"NAS disconnect error: {e}", parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass

    def start_winget_export_thread(self):
        """Starts a thread to export installed apps list using winget export."""
        folder_selected = filedialog.askdirectory(title="Select Folder to Save Installed Apps List")
        if not folder_selected:
            self.message_queue.put(('log', "Winget export cancelled (no folder)."))
            return

        filename = f"winget_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        output_path = os.path.join(folder_selected, filename)

        # open log in chosen folder
        self.open_log_file(folder_selected, "winget_export")

        self.message_queue.put(('log', "Exporting installed applications (winget)..."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Winget export in progress..."))

        t = threading.Thread(target=self.run_winget_export, args=(output_path,), daemon=True)
        t.start()
        
    def run_winget_export(self, output_path):
        """Runs winget export to create a JSON of installed apps. On failure, falls back to 'winget list' text."""
        folder = os.path.dirname(output_path)
        try:
            # Log winget version for diagnostics
            try:
                ver = subprocess.run(['winget', '--version'], shell=False, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                self.write_detailed_log(f"winget --version -> rc={ver.returncode}, out={ver.stdout.strip() or ver.stderr.strip()}")
            except Exception:
                pass

            # Modified winget export command - removed problematic option
            cmd = ['winget', 'export', '--output', output_path, '--include-versions',
                   '--accept-source-agreements', '--disable-interactivity']
            
            self.write_detailed_log(f"Executing winget export command: {' '.join(cmd)}")
            
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       universal_newlines=True, encoding='utf-8', errors='ignore')

            # Stream output to detailed log
            while True:
                line = process.stdout.readline()
                if not line:
                    if process.poll() is not None:
                        break
                    time.sleep(0.05)
                    continue
                self.write_detailed_log(line.strip())

            err = process.stderr.read()
            if err:
                self.write_detailed_log(f"winget export stderr: {err.strip()}")

            rc = process.returncode
            self.write_detailed_log(f"winget export completed with return code: {rc}")
            
            if rc != 0:
                # First alternative: try a more basic winget export
                self.write_detailed_log(f"winget export failed with code {rc}. Trying basic export command.")
                try:
                    basic_cmd = ['winget', 'export', '--output', output_path]
                    self.write_detailed_log(f"Executing basic winget export: {' '.join(basic_cmd)}")
                    
                    basic_result = subprocess.run(basic_cmd, shell=False, capture_output=True, 
                                                text=True, encoding='utf-8', errors='ignore', timeout=300)
                    
                    if basic_result.returncode == 0:
                        self.write_detailed_log("Basic winget export succeeded!")
                        self.write_detailed_log(f"Winget export completed: {output_path}")
                        self.message_queue.put(('log', f"Winget export completed: {output_path}"))
                        self.message_queue.put(('open_folder', folder))
                        return
                    else:
                        self.write_detailed_log(f"Basic export also failed with code {basic_result.returncode}")
                        if basic_result.stderr:
                            self.write_detailed_log(f"Basic export stderr: {basic_result.stderr.strip()}")
                except Exception as basic_e:
                    self.write_detailed_log(f"Basic export attempt failed: {basic_e}")

                # Second alternative: winget list fallback
                self.write_detailed_log("Attempting fallback: winget list.")
                fb_name = f"winget_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                fb_path = os.path.join(folder, fb_name)
                
                try:
                    # Safer winget list command
                    list_cmd = ['winget', 'list', '--accept-source-agreements']
                    self.write_detailed_log(f"Executing winget list: {' '.join(list_cmd)}")
                    
                    p2 = subprocess.run(list_cmd, shell=False, capture_output=True, 
                                      text=True, encoding='utf-8', errors='ignore', timeout=300)
                    
                    if p2.returncode == 0:
                        try:
                            with open(fb_path, 'w', encoding='utf-8', errors='ignore') as f:
                                f.write(p2.stdout)
                            self.write_detailed_log(f"Saved fallback list to {fb_path}")
                            self.message_queue.put(('log', f"winget export failed (code {rc}). Saved fallback (winget list) to {fb_path}"))
                            self.message_queue.put(('open_folder', folder))
                            return
                        except Exception as fe:
                            self.write_detailed_log(f"Failed writing fallback file: {fe}")
                    else:
                        self.write_detailed_log(f"winget list fallback failed with code {p2.returncode}")
                        if p2.stderr:
                            self.write_detailed_log(f"winget list stderr: {p2.stderr.strip()}")
                except subprocess.TimeoutExpired:
                    self.write_detailed_log("winget list command timed out after 300 seconds")
                except Exception as e2:
                    self.write_detailed_log(f"winget list fallback error: {e2}")
                
                # Third alternative: very basic winget list
                self.write_detailed_log("Attempting most basic winget list.")
                try:
                    simple_cmd = ['winget', 'list']
                    p3 = subprocess.run(simple_cmd, shell=False, capture_output=True, 
                                      text=True, encoding='utf-8', errors='ignore', timeout=180)
                    
                    if p3.returncode == 0:
                        simple_fb_name = f"winget_simple_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                        simple_fb_path = os.path.join(folder, simple_fb_name)
                        with open(simple_fb_path, 'w', encoding='utf-8', errors='ignore') as f:
                            f.write(p3.stdout)
                        self.write_detailed_log(f"Saved simple fallback list to {simple_fb_path}")
                        self.message_queue.put(('log', f"All winget export attempts failed. Saved basic list to {simple_fb_path}"))
                        self.message_queue.put(('open_folder', folder))
                        return
                except Exception as e3:
                    self.write_detailed_log(f"Simple winget list also failed: {e3}")
                    
                # If all attempts fail
                raise RuntimeError(f"All winget export/list attempts failed. Primary error code: {rc}")

            # If winget export succeeds
            self.write_detailed_log(f"Winget export completed successfully: {output_path}")
            self.message_queue.put(('log', f"Winget export completed: {output_path}"))
            self.message_queue.put(('open_folder', os.path.dirname(output_path)))
            
        except FileNotFoundError:
            self.write_detailed_log("winget not found on system.")
            self.message_queue.put(('show_error', "winget is not available. Install 'App Installer' from Microsoft Store."))
        except Exception as e:
            self.write_detailed_log(f"Winget export error: {e}")
            error_msg = (
                f"Winget export failed: {e}\n\n"
                f"Possible solutions:\n"
                f"1. Update 'App Installer' from Microsoft Store\n"
                f"2. Run: winget source reset --force\n"
                f"3. Check if winget export is enabled in settings\n"
                f"4. Try running as administrator"
            )
            self.message_queue.put(('show_error', error_msg))
        finally:
            self.close_log_file()
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('update_status', "Winget export complete!"))
            self.message_queue.put(('enable_buttons', None))
            
    def run_copy_thread(self, source_paths, destination_dir):
        """Performs the actual file and folder copying in a separate thread (byte progress)."""
        bytes_copied = 0
        try:
            # Generate timestamp
            from datetime import datetime
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Copy Data folder name: 'copy_data_YYYYMMDD_HHMMSS'
            copy_folder_name = f"copy_data_{timestamp_str}"
            final_destination = os.path.join(destination_dir, copy_folder_name)
            
            # Create folder (don't delete existing folder)
            os.makedirs(final_destination, exist_ok=True)
            self.write_detailed_log(f"Created copy destination folder: {final_destination}")
            
            # --- Calculate total size ---
            self.message_queue.put(('update_status', "Calculating total size..."))
            total_bytes = 0
            for path in source_paths:
                total_bytes += self.compute_total_bytes_safe(path)

            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Copy total bytes: {total_bytes}")
            self.message_queue.put(('log', f"Found {self.format_bytes(total_bytes)} to copy."))

            # --- Check destination free space ---
            self.message_queue.put(('update_status', "Checking available space..."))
            free = self.get_free_space_reliable(destination_dir)

            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space. Required={total_bytes} Free={free}")
                self.message_queue.put((
                    'show_error',
                    f"Not enough free space.\n"
                    f"Required: {self.format_bytes(total_bytes)}\n"
                    f"Available: {self.format_bytes(free)}"
                ))
                return

            # --- Progress callback ---
            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            # --- Actual copy loop ---
            files_processed = 0
            for path in source_paths:
                if os.path.isdir(path):
                    # For directories: copy while preserving original directory name
                    dir_name = os.path.basename(path)
                    dest_base = os.path.join(final_destination, dir_name)
                    
                    for dirpath, dirnames, filenames in os.walk(path, topdown=True, onerror=None):
                        dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]

                        # Create destination directory
                        rel_dir = os.path.relpath(dirpath, path)
                        if rel_dir == ".":
                            dest_dir = dest_base
                        else:
                            dest_dir = os.path.join(dest_base, rel_dir)
                        os.makedirs(dest_dir, exist_ok=True)

                        for f in filenames:
                            src_file = os.path.join(dirpath, f)
                            dest_file = os.path.join(dest_dir, f)

                            if self.is_hidden(src_file):
                                continue

                            self.write_detailed_log(f"Copying file: {src_file} -> {dest_file}")
                            self.copy_file_with_progress_safe(src_file, dest_file, progress_cb)
                            files_processed += 1

                            if files_processed % 10 == 0:
                                self.message_queue.put(('log', f"Copied {files_processed} files..."))
                                self.message_queue.put(('update_status', f"Copied {files_processed} files"))
                                
                elif os.path.isfile(path) and not self.is_hidden(path):
                    # For single file
                    dest_file = os.path.join(final_destination, os.path.basename(path))
                    self.write_detailed_log(f"Copying file: {path} -> {dest_file}")
                    self.copy_file_with_progress_safe(path, dest_file, progress_cb)
                    files_processed += 1

            # Final progress
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log(f"Copy completed: {files_processed} files, {self.format_bytes(bytes_copied)}")
            self.message_queue.put(('log', f"Copy complete ({files_processed} files)."))
            self.message_queue.put(('update_status', "Copy complete!"))
            self.message_queue.put(('open_folder', final_destination))

        except Exception as e:
            self.write_detailed_log(f"Copy error: {e}")
            import traceback
            self.write_detailed_log(f"Copy traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"An error occurred during copy: {e}"))
        finally:
            try:
                self.close_log_file()
            except Exception:
                pass
            
            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass
            
            self.message_queue.put(('enable_buttons', None))
            
            if hasattr(self, '_Copy_completed') and self._Copy_completed:
                self.message_queue.put(('update_status', "Copy complete!"))
            else:
                self.message_queue.put(('update_status', "Operation finished"))

    def start_browser_profiles_backup_thread(self):
        """Starts backup of full browser profiles (Chrome/Edge/Firefox)."""
        dest_root = filedialog.askdirectory(title="Select Folder to Save Browser Profiles")
        if not dest_root:
            self.message_queue.put(('log', "Browser profiles backup cancelled (no destination)."))
            return

        # Prepare destination and log
        self.open_log_file(dest_root, "browser_profiles")
        self.message_queue.put(('log', "backup Browser profiles : Chrome/Edge/Firefox profiles (inc. passwords, cookies, extensions)"))
        self.set_buttons_state("disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['mode'] = "determinate"
        self.message_queue.put(('update_status', "Preparing browser profiles backup..."))

        t = threading.Thread(target=self.run_browser_profiles_backup, args=(dest_root,), daemon=True)
        t.start()

    def run_browser_profiles_backup(self, dest_root):
        bytes_copied = 0
        try:
            # Resolve standard locations
            local = os.path.expandvars('%LOCALAPPDATA%')
            roaming = os.path.expandvars('%APPDATA%')

            # Sources: (abs_path, rel_dest_subdir)
            sources = []
            chrome_src = os.path.join(local, 'Google', 'Chrome', 'User Data')
            if os.path.exists(chrome_src):
                sources.append((chrome_src, os.path.join('Chrome', 'User Data')))
            edge_src = os.path.join(local, 'Microsoft', 'Edge', 'User Data')
            if os.path.exists(edge_src):
                sources.append((edge_src, os.path.join('Edge', 'User Data')))
            ff_src = os.path.join(roaming, 'Mozilla', 'Firefox', 'Profiles')
            if os.path.exists(ff_src):
                sources.append((ff_src, os.path.join('Firefox', 'Profiles')))
            ff_ini = os.path.join(roaming, 'Mozilla', 'Firefox', 'profiles.ini')

            # Destination root with timestamp
            ts = datetime.now().strftime('%Y-%m-%d_%H%M%S')
            dest_base = os.path.join(dest_root, f"BrowserProfiles_backup_{ts}")
            os.makedirs(dest_base, exist_ok=True)

            # Helpers to avoid reparse points
            def is_reparse(p):
                try:
                    if _have_pywin32 and win32api is not None and win32con is not None:
                        attrs = win32api.GetFileAttributes(p)
                        return bool(attrs & win32con.FILE_ATTRIBUTE_REPARSE_POINT)
                    else:
                        try:
                            attrs = _get_file_attributes(p)
                        except Exception:
                            return False
                        return bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)
                except Exception:
                    return False

            # Compute total bytes (no hidden/system filtering; copy everything)
            total = 0
            for src, _rel in sources:
                for dirpath, dirnames, filenames in os.walk(src, topdown=True, onerror=lambda e: None):
                    # skip reparse dirs
                    dirnames[:] = [d for d in dirnames if not is_reparse(os.path.join(dirpath, d))]
                    for f in filenames:
                        p = os.path.join(dirpath, f)
                        try:
                            if not is_reparse(p):
                                total += os.path.getsize(p)
                        except Exception:
                            pass
            if os.path.exists(ff_ini):
                try:
                    total += os.path.getsize(ff_ini)
                except Exception:
                    pass
            self.message_queue.put(('update_progress_max', total))
            self.write_detailed_log(f"Browser profiles total bytes: {total}")

            # Free space check
            try:
                free = shutil.disk_usage(dest_base).free
            except Exception:
                free = self.get_free_space(dest_base)
            if total > free:
                self.write_detailed_log(f"Insufficient free space. Required={total} Free={free}")
                self.message_queue.put(('show_error', f"Not enough free space.\nRequired: {self.format_bytes(total)}\nAvailable: {self.format_bytes(free)}"))
                return

            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            # Copy directories
            for src, rel in sources:
                dst_dir = os.path.join(dest_base, rel)
                self.write_detailed_log(f"Copying directory: {src} -> {dst_dir}")
                if os.path.exists(dst_dir):
                    try:
                        shutil.rmtree(dst_dir)
                        self.write_detailed_log(f"  - Removed existing destination {dst_dir}")
                    except Exception as e:
                        self.write_detailed_log(f"  - Failed to remove existing destination: {e}")
                os.makedirs(dst_dir, exist_ok=True)
                for dirpath, dirnames, filenames in os.walk(src, topdown=True, onerror=lambda e: None):
                    dirnames[:] = [d for d in dirnames if not is_reparse(os.path.join(dirpath, d))]
                    relp = os.path.relpath(dirpath, src)
                    target_dir = os.path.join(dst_dir, relp) if relp != '.' else dst_dir
                    os.makedirs(target_dir, exist_ok=True)
                    for f in filenames:
                        sfile = os.path.join(dirpath, f)
                        if is_reparse(sfile):
                            continue
                        dfile = os.path.join(target_dir, f)
                        self.write_detailed_log(f"Copying file: {sfile} -> {dfile}")
                        self.copy_file_with_progress(sfile, dfile, progress_cb)

            # Copy Firefox profiles.ini
            if os.path.exists(ff_ini):
                dst_ini = os.path.join(dest_base, 'Firefox', 'profiles.ini')
                os.makedirs(os.path.dirname(dst_ini), exist_ok=True)
                try:
                    self.write_detailed_log(f"Copying file: {ff_ini} -> {dst_ini}")
                    self.copy_file_with_progress(ff_ini, dst_ini, progress_cb)
                except Exception as e:
                    self.write_detailed_log(f"Failed to copy profiles.ini: {e}")

            # Finalize
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log("Browser profiles backup completed.")
            self.message_queue.put(('log', "Browser profiles backup complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Browser profiles backup complete!"))
            self.message_queue.put(('open_folder', dest_base))
        except Exception as e:
            self.write_detailed_log(f"Browser profiles backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during browser profiles backup: {e}"))
        finally:
            self.close_log_file()
            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))                
            except Exception:
                pass
            self.message_queue.put(('update_status', "Browser profiles backup complete!"))
            self.message_queue.put(('enable_buttons', None))

    def start_driver_backup_thread(self):
        """Starts the driver backup process in a separate thread."""
        folder_selected = filedialog.askdirectory(title="Select Driver Backup Folder")
        if not folder_selected:
            self.message_queue.put(('log', "Driver backup folder selection cancelled."))
            return

        # open log file in driver backup folder
        # self.open_log_file(folder_selected, "driver_backup")

        self.message_queue.put(('log', "Driver backup process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver backup in progress..."))

        backup_thread = threading.Thread(target=self.run_driver_backup, args=(folder_selected,), daemon=True)
        backup_thread.start()

    def run_driver_backup(self, folder_selected):
        """Execute driver backup with matching log filename"""
        driver_backup_success = False
        
        try:
            # Get system info and create backup folder name
            system_name = get_system_info()
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            
            # Custom backup folder name (with timestamp)
            backup_folder_name = f"{system_name}_drivers_backup_{timestamp}"
            backup_full_path = os.path.join(folder_selected, backup_folder_name)
            
            # Create backup folder
            os.makedirs(backup_full_path, exist_ok=True)
            
            # Create log file manually with matching name (bypass open_log_file timestamp)
            log_filename = f"{backup_folder_name}.log"
            # log_path = os.path.join(folder_selected, log_filename)
            log_path = os.path.join(backup_full_path, log_filename)
            
            with self.log_lock:
                self._log_file = open(log_path, "a", encoding="utf-8", errors="ignore")
                self.log_file_path = log_path
                try:
                    self._log_file.write(f"[POLICY] Hidden={self.hidden_mode_var.get()} System={self.system_mode_var.get()}\n")
                    self._log_file.flush()
                except Exception:
                    pass
            
            self.message_queue.put(('log', f"Detailed log: {log_path}"))
            
            self.write_detailed_log(f"Driver backup folder created: {backup_full_path}")
            self.message_queue.put(('log', f"Backup folder: {backup_folder_name}"))
            
            # Execute pnputil command
            cmd = f'pnputil.exe /export-driver * "{backup_full_path}"'
            self.write_detailed_log(f"Executing command: {cmd}")
            
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, 
                                     stderr=subprocess.PIPE, universal_newlines=True, encoding='cp949')

            # Real-time output logging
            while process.poll() is None:
                line = process.stdout.readline()
                if line:
                    self.write_detailed_log(line.strip())
            
            # Process remaining output
            remaining_output, error_output = process.communicate()
            if remaining_output:
                for line in remaining_output.splitlines():
                    if line.strip():
                        self.write_detailed_log(line.strip())
            
            if error_output:
                self.write_detailed_log(f"STDERR: {error_output}")
                
            if process.returncode == 0:
                self.write_detailed_log("Driver backup successfully completed.")
                self.message_queue.put(('log', f"Driver backup completed: {backup_folder_name}"))
                self.message_queue.put(('open_folder', backup_full_path))
                driver_backup_success = True
            else:
                self.write_detailed_log(f"Driver backup completed with return code: {process.returncode}")
                self.message_queue.put(('log', f"Driver backup finished (check log for details)"))
                driver_backup_success = False
                
        except Exception as e:
            self.write_detailed_log(f"Driver backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during driver backup: {e}"))
            driver_backup_success = False
            
        finally:
            # Cleanup
            try:
                self.close_log_file()
            except Exception:
                pass

            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass

            # Reset progress bar
            self.message_queue.put(('reset_progress', None))

            self.message_queue.put(('enable_buttons', None))
            self.message_queue.put(('update_status', "Driver backup complete!"))

            if driver_backup_success:
                self.message_queue.put(('play_sound', 'complete'))
            else:
                self.message_queue.put(('play_sound', 'error'))

    # ex folder-name :
    # - "ASUS_TUF_Gaming_B450M-PRO_S_drivers_2025-01-15"
    # - "OptiPlex_7090_drivers_2025-01-15"
    # - "HP_Pavilion_Desktop_drivers_2025-01-15"
    # - "Unknown_Desktop_drivers_2025-01-15"
    
    def start_driver_restore_thread(self):
        """Starts the driver restore process in a separate thread."""
        folder_selected = filedialog.askdirectory(title="Select Driver Restore Folder")
        if not folder_selected:
            self.message_queue.put(('log', "Driver restore folder selection cancelled."))
            return

        self.message_queue.put(('log', "Driver restore process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver restore in progress..."))

        restore_thread = threading.Thread(target=self.run_driver_restore, args=(folder_selected,), daemon=True)
        restore_thread.start()
        
    # Improved Features:

    #     System Model Name Matching Priority:

    #         Current System: PRIME_Z490-P

    #         Folder List: PRIME_Z490-P_drivers_backup_2025-01-15, ASUS_TUF_drivers_backup_2025-01-14

    #         → Prioritize selecting the folder that starts with PRIME_Z490-P.

    #     Warning on Mismatch:

    #         If no matching system name is found, use the latest backup but display a warning dialog.

    #         User can choose to continue or cancel.

    #     Log Recording:

    #         Clearly record the current system name and the backup folder information being restored.

    def run_driver_restore(self, folder_selected):
        """Execute driver restore with matching log filename"""
        driver_restore_success = False
        
        try:
            # Get current system info
            current_system_name = get_system_info()
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            
            # Normalize path
            folder_selected = folder_selected.rstrip(os.sep)
            
            # Check if selected folder is a driver backup folder itself
            folder_name = os.path.basename(folder_selected)
            
            # Selected folder must be a driver backup folder
            if '_drivers_backup_' not in folder_name.lower():
                error_msg = f"Please select a driver backup folder.\n\nSelected: {folder_name}\n\nFolder name should contain '_drivers_backup_'"
                self.message_queue.put(('show_error', error_msg))
                return
            
            # Use selected folder as driver backup folder
            driver_backup_folder = folder_selected
            
            # Extract system name from backup folder name
            # Example: PRIME_Z490-P_drivers_backup_2025-01-15_120000 -> PRIME_Z490-P
            backup_system_name = folder_name.split('_drivers_backup_')[0] if '_drivers_backup_' in folder_name else "Unknown"
            
            self.write_detailed_log(f"Selected driver backup folder: {driver_backup_folder}")
            self.write_detailed_log(f"Backup system name: {backup_system_name}")
            self.write_detailed_log(f"Current system name: {current_system_name}")
            
            # Check system name mismatch
            if backup_system_name.lower() != current_system_name.lower() and backup_system_name != "Unknown":
                warning_msg = (
                    f"Current System: {current_system_name}\n"
                    f"Backup From: {backup_system_name}\n\n"
                    f"The driver backup may not match your current system.\n"
                    f"Do you want to continue?"
                )
                
                # Ask user for confirmation
                if not Win11Dialog.askyesno(self.translator.get('system_mismatch_warning'), warning_msg,
                                          parent=self, theme=self.theme, translator=self.translator):
                    self.write_detailed_log("User cancelled restore due to system mismatch")
                    self.message_queue.put(('log', "Driver restore cancelled by user"))
                    return
                
                self.message_queue.put(('log', f"Warning: Restoring from different system backup"))
            else:
                self.message_queue.put(('log', f"Restoring from matching system backup"))
            
            # Create log filename using backup system name and save inside backup folder
            log_filename = f"{backup_system_name}_drivers_restore_{timestamp}.log"
            log_path = os.path.join(driver_backup_folder, log_filename)
            
            # Create log file
            with self.log_lock:
                self._log_file = open(log_path, "a", encoding="utf-8", errors="ignore")
                self.log_file_path = log_path
                try:
                    self._log_file.write(f"[POLICY] Hidden={self.hidden_mode_var.get()} System={self.system_mode_var.get()}\n")
                    self._log_file.flush()
                except Exception:
                    pass
            
            self.message_queue.put(('log', f"Detailed log: {log_path}"))
            self.message_queue.put(('log', f"Restoring from: {os.path.basename(driver_backup_folder)}"))
            
            self.write_detailed_log(f"Driver restore started from: {driver_backup_folder}")
            
            # Execute pnputil restore command (only search in the specific backup folder)
            cmd = f'pnputil.exe /add-driver "{driver_backup_folder}\\*.inf" /subdirs /install'
            self.write_detailed_log(f"Executing command: {cmd}")
            
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, 
                                     stderr=subprocess.PIPE, universal_newlines=True, encoding='cp949')

            # Real-time output logging
            while process.poll() is None:
                line = process.stdout.readline()
                if line:
                    self.write_detailed_log(line.strip())
            
            # Process remaining output
            remaining_output, error_output = process.communicate()
            if remaining_output:
                for line in remaining_output.splitlines():
                    if line.strip():
                        self.write_detailed_log(line.strip())
            
            if error_output:
                self.write_detailed_log(f"STDERR: {error_output}")

            if process.returncode == 0:
                self.write_detailed_log("Driver restore successfully completed.")
                self.message_queue.put(('log', "Driver restore completed."))
                driver_restore_success = True
            else:
                self.write_detailed_log(f"Driver restore completed with return code: {process.returncode}")
                self.message_queue.put(('log', f"Driver restore finished (check log for details)"))
                driver_restore_success = False
                
        except Exception as e:
            self.write_detailed_log(f"Driver restore error: {e}")
            import traceback
            self.write_detailed_log(f"Traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"An error occurred during driver restore: {e}"))
            driver_restore_success = False
            
        finally:
            self.close_log_file()
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('update_status', "Driver restore complete!"))
            self.message_queue.put(('enable_buttons', None))
            
            if driver_restore_success:
                self.message_queue.put(('play_sound', 'complete'))
            else:
                self.message_queue.put(('play_sound', 'error'))

    def open_device_manager(self):
        """Opens the Device Manager using multiple fallbacks for reliability."""
        try:
            # 1) try direct path in System32 with os.startfile (preferred)
            windir = os.environ.get("WINDIR", r"C:\Windows")
            devmgmt_path = os.path.join(windir, "system32", "devmgmt.msc")
            if os.path.exists(devmgmt_path):
                os.startfile(devmgmt_path)
                self.log("Opening Device Manager...")
                return

            # 2) try using mmc (Microsoft Management Console)
            try:
                subprocess.Popen(["mmc", "devmgmt.msc"], shell=False)
                self.log("Opening Device Manager via mmc...")
                return
            except Exception:
                pass

            # 3) fallback to shell 'start' (works in many contexts)
            try:
                subprocess.Popen('start devmgmt.msc', shell=True)
                self.log("Opening Device Manager (via shell start)...")
                return
            except Exception:
                pass

            # If none worked, raise
            raise FileNotFoundError("devmgmt.msc not found / cannot be launched")

        except Exception as e:
            self.log(f"Error opening Device Manager: {e}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('unable_to_open_devmgr')}:\n{e}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass

    def open_file_explorer(self):
        """Opens File Explorer to the user's home directory."""
        try:
            user_home = os.path.expanduser("~")
            os.startfile(user_home)
            self.log(f"Opening File Explorer to: {user_home}")
        except Exception as e:
            self.log(f"Error opening File Explorer: {e}")

    def on_exit(self):
        """(Deprecated) Exit disabled. Replaced by Browser Profiles backup button."""
        # Previous exit logic disabled per request:
        # try:
        #     self.save_settings()
        #     self.close_log_file()
        # except Exception:
        #     pass
        # self.quit()
        # self.destroy()
        pass

    # Settings persistence
    def settings_file_path(self):
        appdata = os.environ.get('APPDATA') or os.path.expanduser('~')
        cfg_dir = os.path.join(appdata, 'ezbak')
        os.makedirs(cfg_dir, exist_ok=True)
        return os.path.join(cfg_dir, 'settings.json')

    def load_settings(self):
        import json
        path = self.settings_file_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # legacy keys
                if 'hidden' in data:
                    self.hidden_mode_var.set(data.get('hidden', 'exclude'))
                if 'system' in data:
                    self.system_mode_var.set(data.get('system', 'exclude'))
                if 'retention_count' in data:
                    self.retention_count_var.set(str(data.get('retention_count', '2')))
                if 'log_retention_days' in data:
                    self.log_retention_days_var.set(str(data.get('log_retention_days', '30')))

                # Sound settings load
                if 'sound_enabled' in data:
                    sound_enabled = data.get('sound_enabled', True)
                    self.sound_enabled_var.set(sound_enabled)
                # new filter structure
                fobj = data.get('filters', {}) if isinstance(data, dict) else {}
                inc = fobj.get('include', []) if isinstance(fobj, dict) else []
                exc = fobj.get('exclude', []) if isinstance(fobj, dict) else []
                
                def _sanitize(lst):
                    out = []
                    for it in lst or []:
                        try:
                            t = (it.get('type') or '').strip().lower()
                            p = (it.get('pattern') or '').strip()
                            if t in ('ext','name','path') and p:
                                out.append({'type': t, 'pattern': p})
                        except Exception:
                            continue
                    return out
                self.filters = {'include': _sanitize(inc), 'exclude': _sanitize(exc)}

                # Load theme and language settings
                if 'theme' in data:
                    theme_mode = data.get('theme', 'dark')
                    self.theme.current_mode = theme_mode
                    self.theme.colors = self.theme.DARK if theme_mode == 'dark' else self.theme.LIGHT

                if 'language' in data:
                    lang_code = data.get('language', 'en')
                    self.translator.current_lang = lang_code
                    # Note: UI texts will be updated after widgets are created in create_widgets()

            except Exception as e:
                print(f"DEBUG: Settings load error: {e}")

    def save_settings(self):
        import json
        path = self.settings_file_path()
        try:
            data = {
                'hidden': self.hidden_mode_var.get(),
                'system': self.system_mode_var.get(),
                'retention_count': self.retention_count_var.get(),
                'log_retention_days': self.log_retention_days_var.get(),
                'sound_enabled': self.sound_enabled_var.get(),
                'filters': self.filters,
                'theme': self.theme.current_mode,
                'language': self.translator.current_lang,
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f)
        except Exception as e:
            print(f"DEBUG: Settings save error: {e}")

class ScheduleBackupDialog(tk.Toplevel, DialogShortcuts):
    def __init__(self, parent):
        print("DEBUG: ScheduleBackupDialog.__init__ started")

        # Important initialization order: set basic properties first
        self.action = None
        self.result = None
        self.task_name = None
        self.parent = parent
        self.theme = parent.theme  # Use parent's theme
        self.translator = parent.translator  # Use parent's translator
        # Initialize filters with default onedrive* exclude filter
        self.filters = {'include': [], 'exclude': [{'type': 'name', 'pattern': 'onedrive*'}]}

        try:
            tk.Toplevel.__init__(self, parent)  # Explicitly initialize Toplevel
            print("DEBUG: Toplevel initialized")

            self.title(parent.translator.get('schedule_backup'))
            self.configure(bg=self.theme.get('bg'))
            
            # Handle window close safely - set first
            self.protocol("WM_DELETE_WINDOW", self.on_close)
            print("DEBUG: WM_DELETE_WINDOW protocol set")
            
            # Set icon (continue even if it fails)
            try:
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except Exception as icon_error:
                print(f"DEBUG: Icon setting failed: {icon_error}")
                pass
                
            # Set window size and position
            w, h = 546, 380
            self.geometry(f"{w}x{h}")
            
            try:
                self.update_idletasks()
                parent_x = parent.winfo_rootx()
                parent_y = parent.winfo_rooty()
                x = parent_x + 50
                y = parent_y + 50
                self.geometry(f"{w}x{h}+{x}+{y}")
                print(f"DEBUG: Window positioned at {x},{y}")
            except Exception as pos_error:
                print(f"DEBUG: Window positioning failed: {pos_error}")
                pass

            try:
                self.transient(parent)
                print("DEBUG: Transient set")
                self.grab_set()
                print("DEBUG: Grab set")
            except Exception as modal_error:
                print(f"DEBUG: Modal setting failed: {modal_error}")
                pass

            # Create UI components
            print("DEBUG: Creating UI widgets...")
            self._create_widgets()
            print("DEBUG: UI widgets created successfully")
            
            # Setup shortcuts (after UI creation)
            self._setup_schedule_shortcuts()
            print("DEBUG: Shortcuts setup completed")
            
        except Exception as init_error:
            print(f"DEBUG: ScheduleBackupDialog init failed: {init_error}")
            self.action = 'error'
            try:
                self.destroy()
            except Exception:
                pass
            raise init_error

    def _create_win11_button(self, parent, text, command, is_primary=False, width=12):
        """Create a Windows 11 style button for dialogs"""
        if is_primary:
            btn_bg = self.theme.get('accent')
            btn_fg = '#FFFFFF'
            hover_bg = self.theme.get('accent_hover')
            hover_fg = '#FFFFFF'
            pressed_bg = self.theme.get('accent_hover')
        else:
            btn_bg = self.theme.get('bg_elevated')
            btn_fg = self.theme.get('fg')
            hover_bg = self.theme.get('hover')
            hover_fg = self.theme.get('fg')
            pressed_bg = self.theme.get('pressed')

        border_color = self.theme.get('border')

        btn = tk.Button(parent, text=text,
                       command=command,
                       bg=btn_bg,
                       fg=btn_fg,
                       font=("Segoe UI", 9),
                       relief="flat",
                       bd=0,
                       highlightthickness=1,
                       highlightbackground=border_color,
                       highlightcolor=border_color,
                       borderwidth=1,
                       width=width,
                       cursor="hand2",
                       padx=12,
                       pady=6)

        # Windows 11 style hover and press effects
        def on_enter(e):
            e.widget.config(bg=hover_bg, fg=hover_fg, relief="flat")

        def on_leave(e):
            e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

        def on_press(e):
            e.widget.config(bg=pressed_bg, relief="flat")

        def on_release(e):
            e.widget.config(bg=hover_bg, relief="flat")

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)

        return btn

    def _setup_schedule_shortcuts(self):
        """Setup shortcuts specific to schedule dialog"""
        # Apply basic dialog shortcuts
        self.setup_dialog_shortcuts()

        # Shortcuts specific to schedule dialog
        self.bind('<Control-c>', lambda e: self.on_create())
        self.bind('<Control-d>', lambda e: self.on_delete())
        self.bind('<Control-b>', lambda e: self._browse_destination())
        self.bind('<Control-f>', lambda e: self.open_filter_manager())
        self.bind('<F1>', lambda e: self._show_schedule_help())

        # Select schedule type with number keys
        self.bind('<Alt-1>', lambda e: self._set_schedule("Daily"))
        self.bind('<Alt-2>', lambda e: self._set_schedule("Weekly"))
        self.bind('<Alt-3>', lambda e: self._set_schedule("Monthly"))

    def _create_widgets(self):
        """Create UI widgets with Windows 11 styling"""
        try:
            # Safely get user name
            user_name = "Unknown"
            try:
                if hasattr(self.parent, 'user_var') and self.parent.user_var.get():
                    user_name = self.parent.user_var.get()
            except Exception:
                pass

            # Modern styling
            label_font = ("Segoe UI", 10, "bold")
            entry_font = ("Segoe UI", 10)

            # Task Name
            row = 0
            tk.Label(self, text="Task Name", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )
            self.task_var = tk.StringVar(value=f"ezBAK_Backup_{user_name}")
            tk.Entry(self, textvariable=self.task_var, width=40, font=entry_font,
                    bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                    relief="solid", bd=1).grid(
                row=row, column=1, columnspan=2, sticky="ew", padx=15, pady=8
            )

            # Destination Folder
            row += 1
            tk.Label(self, text="Destination Folder", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )
            self.dest_var = tk.StringVar()
            tk.Entry(self, textvariable=self.dest_var, width=30, font=entry_font,
                    bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                    relief="solid", bd=1).grid(
                row=row, column=1, sticky="ew", padx=15, pady=8
            )
            tk.Button(
                self, text=self.add_shortcut_text("Browse...", "Ctrl+B"),
                command=self._browse_destination,
                bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                font=("Segoe UI", 9), relief="flat", bd=0,
                cursor="hand2"
            ).grid(row=row, column=2, padx=15, pady=8, sticky="ew")

            # Schedule
            row += 1
            tk.Label(self, text=self.add_shortcut_text("Schedule", "Alt+1/2/3"),
                    bg=self.theme.get('bg'), fg=self.theme.get('fg'),
                    font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )

            schedule_frame = tk.Frame(self, bg=self.theme.get('bg'))
            schedule_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=15, pady=8)

            self.schedule_var = tk.StringVar(value="Daily")
            schedule_combo = ttk.Combobox(
                schedule_frame, textvariable=self.schedule_var,
                values=["Daily", "Weekly", "Monthly"], state="readonly",
                width=18, font=entry_font
            )
            schedule_combo.pack(side="left")

            # Day selector for Monthly schedule (1-30)
            tk.Label(schedule_frame, text="  Day:", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=("Segoe UI", 9)).pack(side="left", padx=(10, 5))

            self.monthly_day_var = tk.StringVar(value="1")
            self.monthly_day_combo = ttk.Combobox(
                schedule_frame, textvariable=self.monthly_day_var,
                values=[str(i) for i in range(1, 31)],
                state="readonly",
                width=4, font=entry_font
            )
            self.monthly_day_combo.pack(side="left")

            # Update visibility based on schedule type
            def on_schedule_change(*args):
                if self.schedule_var.get() == "Monthly":
                    self.monthly_day_combo.config(state="readonly")
                else:
                    self.monthly_day_combo.config(state="disabled")

            self.schedule_var.trace_add('write', on_schedule_change)
            on_schedule_change()  # Initialize visibility

            # Time
            row += 1
            tk.Label(self, text="Time", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )

            time_frame = tk.Frame(self, bg=self.theme.get('bg'))
            time_frame.grid(row=row, column=1, sticky="w", padx=15, pady=8)

            # Hour combobox (00-23)
            self.hour_var = tk.StringVar(value="02")
            hour_combo = ttk.Combobox(
                time_frame, textvariable=self.hour_var,
                values=[f"{i:02d}" for i in range(24)],
                state="readonly",
                width=4, font=entry_font
            )
            hour_combo.pack(side="left")

            tk.Label(time_frame, text=":", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).pack(side="left", padx=5)

            # Minute combobox (00-59)
            self.minute_var = tk.StringVar(value="00")
            minute_combo = ttk.Combobox(
                time_frame, textvariable=self.minute_var,
                values=[f"{i:02d}" for i in range(60)],
                state="readonly",
                width=4, font=entry_font
            )
            minute_combo.pack(side="left")

            # Attributes
            row += 1
            tk.Label(self, text="Attributes", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )

            attr_frame = tk.Frame(self, bg=self.theme.get('bg'))
            attr_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=15, pady=8)

            self.hidden_var = tk.BooleanVar(value=False)
            self.system_var = tk.BooleanVar(value=False)

            tk.Checkbutton(attr_frame, text="Include Hidden", variable=self.hidden_var,
                          bg=self.theme.get('bg'), fg=self.theme.get('fg'),
                          selectcolor=self.theme.get('bg'),
                          font=("Segoe UI", 9), relief="flat", bd=0).pack(side="left")

            tk.Checkbutton(attr_frame, text="Include System", variable=self.system_var,
                          bg=self.theme.get('bg'), fg=self.theme.get('fg'),
                          selectcolor=self.theme.get('bg'),
                          font=("Segoe UI", 9), relief="flat", bd=0).pack(side="left", padx=(15, 0))

            # Backups to Keep
            row += 1
            tk.Label(self, text="Backups to Keep (0=all):", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )
            self.retention_count_var = tk.StringVar(value="2")
            tk.Spinbox(self, from_=0, to=99, textvariable=self.retention_count_var,
                      width=10, font=entry_font,
                      bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                      buttonbackground=self.theme.get('accent'),
                      relief="solid", bd=1).grid(
                row=row, column=1, sticky="w", padx=15, pady=8
            )

            # Logs to Keep
            row += 1
            tk.Label(self, text="Logs to Keep (days, 0=off):", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )
            self.log_retention_days_var = tk.StringVar(value="30")
            tk.Spinbox(self, from_=0, to=365, textvariable=self.log_retention_days_var,
                      width=10, font=entry_font,
                      bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                      buttonbackground=self.theme.get('accent'),
                      relief="solid", bd=1).grid(
                row=row, column=1, sticky="w", padx=15, pady=8
            )

            # Filters
            row += 1
            tk.Label(self, text="Filters:", bg=self.theme.get('bg'),
                    fg=self.theme.get('fg'), font=label_font).grid(
                row=row, column=0, sticky="w", padx=15, pady=8
            )
            tk.Button(
                self, text=self.add_shortcut_text("Manage Filters...", "Ctrl+F"),
                command=self.open_filter_manager,
                bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                font=("Segoe UI", 9), relief="flat", bd=0,
                cursor="hand2"
            ).grid(row=row, column=1, sticky="w", padx=15, pady=8)

            # Button area with modern styling
            row += 1
            btn_frame = tk.Frame(self, bg=self.theme.get('bg'))
            btn_frame.grid(row=row, column=0, columnspan=3, sticky="e", padx=15, pady=20)

            # Windows 11 style buttons
            self._create_win11_button(btn_frame, self.add_shortcut_text("Create", "Ctrl+C"),
                                     self.on_create, is_primary=True).pack(side="right", padx=5)
            self._create_win11_button(btn_frame, self.add_shortcut_text("Delete", "Ctrl+D"),
                                     self.on_delete).pack(side="right", padx=5)
            self._create_win11_button(btn_frame, self.add_shortcut_text("Close", "Esc"),
                                     self.on_close).pack(side="right", padx=5)

            # Configure grid weights
            self.columnconfigure(1, weight=1)

        except Exception as widget_error:
            print(f"DEBUG: Widget creation failed: {widget_error}")
            raise widget_error

    def _set_schedule(self, schedule_type):
        """Set schedule type with shortcut key"""
        try:
            self.schedule_var.set(schedule_type)
            print(f"DEBUG: Schedule set to {schedule_type}")
        except Exception as e:
            print(f"DEBUG: Failed to set schedule: {e}")

    def _show_schedule_help(self):
        """Show schedule help with F1 key"""
        help_text = """
Schedule Backup Shortcuts:

Ctrl + C     Create Task
Ctrl + D     Delete Task
Ctrl + B     Browse Destination
Ctrl + F     Manage Filters
Esc          Close Dialog
Enter        Create Task
F1           Show This Help

Alt + 1      Set to Daily
Alt + 2      Set to Weekly
Alt + 3      Set to Monthly

Navigation:
Tab          Move to Next Field
Shift + Tab  Move to Previous Field
        """
        
        help_dlg = tk.Toplevel(self)
        help_dlg.title("Schedule Shortcuts Help")
        help_dlg.configure(bg=self.theme.get('bg'))
        help_dlg.geometry("400x350")
        help_dlg.transient(self)
        try:
            help_dlg.grab_set()
        except Exception:
            pass

        text_widget = tk.Text(help_dlg,
                             font=("Consolas", 10),
                             bg=self.theme.get('bg_secondary'),
                             fg=self.theme.get('fg'),
                             relief="flat",
                             wrap="word",
                             padx=12, pady=12)
        text_widget.pack(fill="both", expand=True, padx=15, pady=15)

        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        btn_frame = tk.Frame(help_dlg, bg=self.theme.get('bg'))
        btn_frame.pack(pady=15)
        close_btn = Win11Dialog._create_dialog_button(btn_frame,
                                                      "Close (Esc)",
                                                      help_dlg.destroy,
                                                      self.theme,
                                                      is_primary=True)
        close_btn.pack()
        
        help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
        help_dlg.focus_set()

    def _browse_destination(self):
        """Folder selection dialog"""
        try:
            folder = filedialog.askdirectory(parent=self, title="Select Backup Destination")
            if folder:
                self.dest_var.set(folder)
                print(f"DEBUG: Destination set to: {folder}")
        except Exception as browse_error:
            print(f"DEBUG: Browse failed: {browse_error}")

    def open_filter_manager(self):
        """Open filter manager dialog"""
        try:
            dlg = FilterManagerDialog(self, current_filters=self.filters)
            self.wait_window(dlg)
            if getattr(dlg, 'result', None):
                self.filters = dlg.result
                print(f"DEBUG: Filters updated: {self.filters}")
        except Exception as e:
            print(f"DEBUG: Filter manager error: {e}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('filter_manager_error')}: {e}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                pass

    def on_create(self):
        """Create button handler"""
        print("DEBUG: on_create called")
        try:
            # Validate input values
            task_name = self.task_var.get().strip()
            dest = self.dest_var.get().strip()
            
            if not task_name:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('task_name_required_create'),
                                    parent=self, theme=self.theme, translator=self.translator)
                return
                
            if not dest:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('destination_required'),
                                    parent=self, theme=self.theme, translator=self.translator)
                return
                
            if not os.path.isdir(dest):
                if not Win11Dialog.askyesno(self.translator.get('warning'),
                    f"{self.translator.get('overwrite_warning')}:\n{dest}\n\n{self.translator.get('continue_anyway')}",
                    parent=self, theme=self.theme, translator=self.translator):
                    return
            
            # Combine hour and minute into time string
            hour = self.hour_var.get()
            minute = self.minute_var.get()
            time_str = f"{hour}:{minute}"

            # Validate time format
            try:
                from datetime import datetime
                time_obj = datetime.strptime(time_str, "%H:%M")
                time_str = time_obj.strftime("%H:%M")
            except ValueError:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_time_format'),
                                    parent=self, theme=self.theme, translator=self.translator)
                return

            # Get user name
            user_name = "Unknown"
            try:
                if hasattr(self.parent, 'user_var') and self.parent.user_var.get():
                    user_name = self.parent.user_var.get()
            except Exception:
                pass

            self.action = "create"
            self.result = {
                "task_name": task_name,
                "dest": dest,
                "schedule": self.schedule_var.get(),
                "time_str": time_str,
                "monthly_day": int(self.monthly_day_var.get()),
                "include_hidden": self.hidden_var.get(),
                "include_system": self.system_var.get(),
                "retention_count": int(self.retention_count_var.get()),
                "log_retention_days": int(self.log_retention_days_var.get()),
                "user": user_name,
                "filters": self.filters
            }
            print(f"DEBUG: Result set: {self.result}")
            self._safe_destroy()
            
        except Exception as create_error:
            print(f"DEBUG: on_create failed: {create_error}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('failed_to_create_schedule')}: {create_error}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                print(f"ERROR: {create_error}")

    def on_delete(self):
        """Delete button handler"""
        print("DEBUG: on_delete called")
        try:
            task_name = self.task_var.get().strip()
            if not task_name:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('task_name_required_delete'),
                                    parent=self, theme=self.theme, translator=self.translator)
                return
                
            self.action = "delete"
            self.task_name = task_name
            print(f"DEBUG: Delete task_name set: {task_name}")
            self._safe_destroy()
        except Exception as delete_error:
            print(f"DEBUG: on_delete failed: {delete_error}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('failed_to_delete_schedule')}: {delete_error}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                print(f"ERROR: {delete_error}")

    def on_close(self):
        """Close button handler"""
        print("DEBUG: on_close called")
        self.action = "close"
        self._safe_destroy()
        
    def on_cancel(self):
        """Cancel handler for ESC key"""
        self.on_close()
        
    def on_ok(self):
        """OK handler for Enter key"""
        self.on_create()
        
    def _safe_destroy(self):
        """Safe window destruction"""
        print("DEBUG: _safe_destroy called")
        try:
            self.grab_release()
            print("DEBUG: grab_release successful")
        except Exception as grab_error:
            print(f"DEBUG: grab_release failed: {grab_error}")
            
        try:
            self.destroy()
            print("DEBUG: destroy successful")
        except Exception as destroy_error:
            print(f"DEBUG: destroy failed: {destroy_error}")

class FilterManagerDialog(tk.Toplevel, DialogShortcuts):
    def __init__(self, master, current_filters=None):
        tk.Toplevel.__init__(self, master)
        self.parent = master
        self.theme = master.theme  # Use parent's theme
        self.translator = master.translator  # Use parent's translator
        self.title(self.translator.get('filter_manager'))
        try:
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        self.configure(bg=self.theme.get('bg'))
        self.geometry("750x520")
        self.transient(master)
        try:
            self.grab_set()
        except Exception:
            pass

        self.result = None
        cf = current_filters or {'include': [], 'exclude': []}
        # deep copy simple
        self.inc = [dict(type=i.get('type'), pattern=i.get('pattern')) for i in (cf.get('include') or [])]
        self.exc = [dict(type=i.get('type'), pattern=i.get('pattern')) for i in (cf.get('exclude') or [])]

        self._create_filter_widgets()
        self._setup_filter_shortcuts()
        self._refresh()

    def _create_win11_button(self, parent, text, command, is_primary=False, width=12):
        """Create a Windows 11 style button for dialogs"""
        if is_primary:
            btn_bg = self.theme.get('accent')
            btn_fg = '#FFFFFF'
            hover_bg = self.theme.get('accent_hover')
            hover_fg = '#FFFFFF'
            pressed_bg = self.theme.get('accent_hover')
        else:
            btn_bg = self.theme.get('bg_elevated')
            btn_fg = self.theme.get('fg')
            hover_bg = self.theme.get('hover')
            hover_fg = self.theme.get('fg')
            pressed_bg = self.theme.get('pressed')

        border_color = self.theme.get('border')

        btn = tk.Button(parent, text=text,
                       command=command,
                       bg=btn_bg,
                       fg=btn_fg,
                       font=("Segoe UI", 9),
                       relief="flat",
                       bd=0,
                       highlightthickness=1,
                       highlightbackground=border_color,
                       highlightcolor=border_color,
                       borderwidth=1,
                       width=width,
                       cursor="hand2",
                       padx=12,
                       pady=6)

        # Windows 11 style hover and press effects
        def on_enter(e):
            e.widget.config(bg=hover_bg, fg=hover_fg, relief="flat")

        def on_leave(e):
            e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

        def on_press(e):
            e.widget.config(bg=pressed_bg, relief="flat")

        def on_release(e):
            e.widget.config(bg=hover_bg, relief="flat")

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)

        return btn

    def _setup_filter_shortcuts(self):
        """Setup shortcuts specific to filter dialog"""
        # Apply basic dialog shortcuts
        self.setup_dialog_shortcuts()

        # Filter-related shortcuts
        self.bind('<Control-n>', lambda e: self._add_rule('include'))  # New include rule
        self.bind('<Control-e>', lambda e: self._add_rule('exclude'))  # New exclude rule
        self.bind('<Delete>', lambda e: self._remove_selected_active())
        self.bind('<Control-s>', lambda e: self._save())
        self.bind('<Control-r>', lambda e: self._clear_active())  # Remove all from active list
        self.bind('<F1>', lambda e: self._show_filter_help())

        # Manage list focus
        self.bind('<Tab>', lambda e: self._cycle_focus())
        
    def _cycle_focus(self):
        """Move focus between lists with Tab key"""
        try:
            focused = self.focus_get()
            if focused == self.inc_list:
                self.exc_list.focus_set()
            else:
                self.inc_list.focus_set()
        except Exception:
            self.inc_list.focus_set()

    def _create_filter_widgets(self):
        """Create filter widgets (including shortcut notations)"""
        # Top description
        info_frame = tk.Frame(self, bg=self.theme.get('bg_elevated'))
        info_frame.pack(fill='x', padx=12, pady=(12,6))

        info_text = f"{self.translator.get('filter_info_include')}\n{self.translator.get('filter_info_exclude')}"
        tk.Label(info_frame, text=info_text, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg_secondary'),
                font=("Arial", 9), justify='left').pack(anchor='w')

        # Layout
        top = tk.Frame(self, bg=self.theme.get('bg_elevated'))
        top.pack(fill='both', expand=True, padx=12, pady=6)

        # Include panel (add shortcut notation)
        inc_frame = tk.LabelFrame(top, text=self.add_shortcut_text(self.translator.get('include_rules'), "Ctrl+N"),
                                 bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), font=("Arial", 10, "bold"))
        inc_frame.pack(side='left', fill='both', expand=True, padx=(0,6))

        self.inc_list = tk.Listbox(inc_frame, height=14, font=("Consolas", 9),
                                  bg=self.theme.get('bg_secondary'), fg=self.theme.get('fg'), selectbackground=self.theme.get('success'))
        self.inc_list.pack(fill='both', expand=True, padx=6, pady=6)

        btn_row_inc = tk.Frame(inc_frame, bg=self.theme.get('bg_elevated'))
        btn_row_inc.pack(fill='x', padx=6, pady=(0,6))

        self._create_win11_button(btn_row_inc, self.translator.get('add'),
                                 lambda: self._add_rule('include'), width=8).pack(side='left', padx=2)
        self._create_win11_button(btn_row_inc, self.add_shortcut_text(self.translator.get('remove'), "Del"),
                                 lambda: self._remove_selected('include'), width=12).pack(side='left', padx=2)
        self._create_win11_button(btn_row_inc, self.translator.get('clear_all'),
                                 lambda: self._clear('include'), width=10).pack(side='left', padx=2)

        # Exclude panel (add shortcut notation)
        exc_frame = tk.LabelFrame(top, text=self.add_shortcut_text(self.translator.get('exclude_rules'), "Ctrl+E"),
                                 bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), font=("Arial", 10, "bold"))
        exc_frame.pack(side='left', fill='both', expand=True, padx=(6,0))

        self.exc_list = tk.Listbox(exc_frame, height=14, font=("Consolas", 9),
                                  bg=self.theme.get('bg_secondary'), fg=self.theme.get('fg'), selectbackground=self.theme.get('danger'))
        self.exc_list.pack(fill='both', expand=True, padx=6, pady=6)

        btn_row_exc = tk.Frame(exc_frame, bg=self.theme.get('bg_elevated'))
        btn_row_exc.pack(fill='x', padx=6, pady=(0,6))

        self._create_win11_button(btn_row_exc, self.translator.get('add'),
                                 lambda: self._add_rule('exclude'), width=8).pack(side='left', padx=2)
        self._create_win11_button(btn_row_exc, self.add_shortcut_text(self.translator.get('remove'), "Del"),
                                 lambda: self._remove_selected('exclude'), width=12).pack(side='left', padx=2)
        self._create_win11_button(btn_row_exc, self.translator.get('clear_all'),
                                 lambda: self._clear('exclude'), width=10).pack(side='left', padx=2)

        # Bottom shortcut guide
        shortcut_frame = tk.Frame(self, bg=self.theme.get('bg_elevated'))
        shortcut_frame.pack(fill='x', padx=12, pady=(6,0))

        shortcut_text = "F1: Help  •  Tab: Switch Lists  •  Del: Remove Selected  •  Ctrl+R: Clear Active List"
        tk.Label(shortcut_frame, text=shortcut_text, bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg_secondary'),
                font=("Arial", 8)).pack(anchor='w')

        # Bottom actions (add shortcut notation)
        actions = tk.Frame(self, bg=self.theme.get('bg_elevated'))
        actions.pack(fill='x', padx=12, pady=12)

        self._create_win11_button(actions, self.add_shortcut_text(self.translator.get('help'), "F1"),
                                 self._show_filter_help, width=10).pack(side='left')

        self._create_win11_button(actions, self.add_shortcut_text(self.translator.get('save'), "Ctrl+S"),
                                 self._save, is_primary=True, width=12).pack(side='right')
        self._create_win11_button(actions, self.add_shortcut_text(self.translator.get('dialog_cancel'), "Esc"),
                                 self._cancel, width=12).pack(side='right', padx=(0,8))

    def _rule_to_text(self, r):
        """Convert rule to display text"""
        try:
            rule_type = r.get('type', 'unknown')
            pattern = r.get('pattern', '')
            
            # Add icon by type
            type_icons = {
                'ext': '📄',
                'name': '📁', 
                'path': '🗂️'
            }
            icon = type_icons.get(rule_type, '❓')
            
            return f"{icon} [{rule_type.upper()}] {pattern}"
        except Exception:
            return str(r)

    def _refresh(self):
        """Refresh list"""
        # Include List
        self.inc_list.delete(0, tk.END)
        for r in self.inc:
            self.inc_list.insert(tk.END, self._rule_to_text(r))
        
        # Exclude List
        self.exc_list.delete(0, tk.END)
        for r in self.exc:
            self.exc_list.insert(tk.END, self._rule_to_text(r))

    def _remove_selected(self, kind):
        """Remove selected items"""
        if kind == 'include':
            idxs = list(self.inc_list.curselection())
            for i in reversed(idxs):
                if 0 <= i < len(self.inc):
                    del self.inc[i]
        else:
            idxs = list(self.exc_list.curselection())
            for i in reversed(idxs):
                if 0 <= i < len(self.exc):
                    del self.exc[i]
        self._refresh()

    def _remove_selected_active(self):
        """Remove selected items from currently active list"""
        try:
            # Check focused widget
            focused = self.focus_get()
            if focused == self.inc_list:
                self._remove_selected('include')
            elif focused == self.exc_list:
                self._remove_selected('exclude')
        except Exception:
            pass

    def _clear(self, kind):
        """Clear all"""
        if kind == 'include':
            self.inc = []
        else:
            self.exc = []
        self._refresh()

    def _clear_active(self):
        """Clear all from currently active list"""
        try:
            focused = self.focus_get()
            if focused == self.inc_list:
                self._clear('include')
            elif focused == self.exc_list:
                self._clear('exclude')
        except Exception:
            pass

    def _add_rule(self, kind):
        """Add rule dialog"""
        dlg = tk.Toplevel(self)
        dlg.title(self.translator.get('add_include_rule' if kind == 'include' else 'add_exclude_rule'))
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg=self.theme.get('bg_elevated'))
        dlg.geometry("480x320")
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass

        # Select rule type
        tk.Label(dlg, text=self.translator.get('rule_type'), bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                font=("Arial", 10, "bold")).pack(anchor='w', padx=15, pady=(15,5))

        vtype = tk.StringVar(value='ext')
        tframe = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        tframe.pack(anchor='w', padx=15, fill='x')

        type_options = [
            ('ext', f'📄 {self.translator.get("extension")}', self.translator.get('extension_desc')),
            ('name', f'📁 {self.translator.get("file_folder_name")}', self.translator.get('name_desc')),
            ('path', f'🗂️ {self.translator.get("full_path")}', self.translator.get('path_desc'))
        ]

        for value, label, desc in type_options:
            rb_frame = tk.Frame(tframe, bg=self.theme.get('bg_elevated'))
            rb_frame.pack(anchor='w', pady=2)

            tk.Radiobutton(rb_frame, text=label, variable=vtype, value=value,
                          bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), selectcolor=self.theme.get('bg_elevated'),
                          font=("Arial", 9, "bold")).pack(side='left')
            tk.Label(rb_frame, text=f"  {desc}", bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg_secondary'),
                    font=("Arial", 8)).pack(side='left')

        # Pattern input
        tk.Label(dlg, text=self.translator.get('pattern'), bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'),
                font=("Arial", 10, "bold")).pack(anchor='w', padx=15, pady=(15,5))

        vpat = tk.StringVar()
        pat_frame = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        pat_frame.pack(fill='x', padx=15)

        ent = tk.Entry(pat_frame, textvariable=vpat, width=50, font=("Arial", 10),
                      bg=self.theme.get('bg_secondary'), fg=self.theme.get('fg'), insertbackground=self.theme.get('fg'))
        ent.pack(fill='x')

        # Example text
        example_frame = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        example_frame.pack(fill='x', padx=15, pady=(5,0))

        def update_example(*args):
            try:
                rule_type = vtype.get()
                examples = {
                    'ext': self.translator.get('examples_ext'),
                    'name': self.translator.get('examples_name'),
                    'path': self.translator.get('examples_path')
                }
                example_label.config(text=examples.get(rule_type, ''))
            except Exception:
                pass

        example_label = tk.Label(example_frame, text="", bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg_secondary'),
                               font=("Arial", 8))
        example_label.pack(anchor='w')

        vtype.trace('w', update_example)
        update_example()  # Initial display

        # Buttons
        btns = tk.Frame(dlg, bg=self.theme.get('bg_elevated'))
        btns.pack(fill='x', padx=15, pady=15)

        def _ok():
            p = vpat.get().strip()
            if not p:
                Win11Dialog.showwarning(self.translator.get('warning'), self.translator.get('please_enter_pattern'),
                                      parent=dlg, theme=self.theme, translator=self.translator)
                return

            rule_type = vtype.get().strip().lower()
            rule = {'type': rule_type, 'pattern': p}

            if kind == 'include':
                self.inc.append(rule)
            else:
                self.exc.append(rule)

            self._refresh()

            try:
                dlg.grab_release()
            except Exception:
                pass
            dlg.destroy()

        def _cancel():
            try:
                dlg.grab_release()
            except Exception:
                pass
            dlg.destroy()

        Win11Dialog._create_dialog_button(btns, self.translator.get('add_rule'),
                                          _ok, self.theme, is_primary=True).pack(side='right')
        Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_cancel'),
                                          _cancel, self.theme).pack(side='right', padx=(0,8))

        # Shortcuts setting
        dlg.bind('<Return>', lambda e: _ok())
        dlg.bind('<Escape>', lambda e: _cancel())
        
        try:
            ent.focus_set()
        except Exception:
            pass

    def _show_filter_help(self):
        """Show filter help"""
        help_text = self.translator.get('filter_manager_help_content')

        help_dlg = tk.Toplevel(self)
        help_dlg.title(self.translator.get('filter_manager_help_title'))
        help_dlg.configure(bg=self.theme.get('bg'))
        help_dlg.geometry("600x610")
        help_dlg.transient(self)
        try:
            help_dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        try:
            help_dlg.grab_set()
        except Exception:
            pass

        # Scrollable text area
        text_frame = tk.Frame(help_dlg, bg=self.theme.get('bg'))
        text_frame.pack(fill="both", expand=True, padx=15, pady=15)

        scrollbar = tk.Scrollbar(text_frame, width=20,
                               bg=self.theme.get('scrollbar_bg'),
                               troughcolor=self.theme.get('bg'),
                               activebackground=self.theme.get('accent'))
        scrollbar.pack(side="right", fill="y")

        text_widget = tk.Text(text_frame,
                             font=("Consolas", 9),
                             bg=self.theme.get('text_bg'),
                             fg=self.theme.get('fg'),
                             relief="flat",
                             wrap="word",
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)

        scrollbar.configure(command=text_widget.yview)

        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        # Close Button
        btn_frame = tk.Frame(help_dlg, bg=self.theme.get('bg'))
        btn_frame.pack(pady=(0,15))
        Win11Dialog._create_dialog_button(btn_frame, self.translator.get('help_close_esc'),
                                          help_dlg.destroy, self.theme, is_primary=True).pack()

        help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
        help_dlg.focus_set()

    def _save(self):
        """Save filters"""
        self.result = {'include': self.inc, 'exclude': self.exc}
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _cancel(self):
        """Cancel"""
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def on_cancel(self):
        """Cancel handler for ESC key"""
        self._cancel()
        
    def on_ok(self):
        """OK handler for Enter key"""
        self._save()

class SelectSourcesDialog(tk.Toplevel):
    def __init__(self, master, is_hidden_fn=None, title="Select Sources"):
        try:
            super().__init__(master)
            self.theme = master.theme  # Use parent's theme
            print(f"DEBUG: SelectSourcesDialog initialized with title: {title}")

            self.title(title)
            try:
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except Exception:
                pass
            self.configure(bg=self.theme.get('bg'))
            self.geometry("780x580")
            self.transient(master)
            try:
                self.grab_set()
            except Exception:
                pass

            self.is_hidden_fn = is_hidden_fn or (lambda p: False)
            self.selected_paths = []

            self._checked = set()       # item ids that are checked
            self._item_path = {}        # item id -> full filesystem path

            print("DEBUG: Creating widgets...")
            self._create_source_widgets()
            
            print("DEBUG: Setting up shortcuts...")
            self._setup_source_shortcuts()
            
            print("DEBUG: Adding drive roots...")
            self._add_drive_roots()
            
            print("DEBUG: SelectSourcesDialog initialization complete")
            
        except Exception as e:
            print(f"DEBUG: SelectSourcesDialog init failed: {e}")
            import traceback
            print(f"DEBUG: Traceback: {traceback.format_exc()}")
            raise

    def _setup_source_shortcuts(self):
        """Set shortcuts specific to the source selection dialog"""
        try:
            # Apply basic dialog shortcuts
            self.setup_dialog_shortcuts()
            
            # Source selection related shortcuts
            self.bind('<Control-a>', lambda e: self._select_all())
            self.bind('<Control-n>', lambda e: self._clear_all())  # None (clear all)
            self.bind('<Control-r>', lambda e: self._choose_root())  # Root folder
            self.bind('<Space>', lambda e: self._toggle_selected())
            self.bind('<Control-e>', lambda e: self._expand_selected())
            self.bind('<Control-c>', lambda e: self._collapse_selected())
            self.bind('<F5>', lambda e: self._refresh_tree())
            self.bind('<F1>', lambda e: self._show_source_help())
            
            print("DEBUG: Shortcuts setup complete")
        except Exception as e:
            print(f"DEBUG: _setup_source_shortcuts failed: {e}")

    def setup_dialog_shortcuts(self):
        """Set basic shortcuts for the dialog (implemented directly instead of via DialogShortcuts)"""
        try:
            # Basic dialog shortcuts
            self.bind('<Escape>', lambda e: self.on_cancel() if hasattr(self, 'on_cancel') else self.destroy())
            self.bind('<Return>', lambda e: self.on_ok() if hasattr(self, 'on_ok') else None)
            self.bind('<Alt-F4>', lambda e: self.destroy())
            
            # Set focus (to receive keyboard events)
            self.focus_set()
        except Exception as e:
            print(f"DEBUG: setup_dialog_shortcuts failed: {e}")
        
    def add_shortcut_text(self, text, shortcut):
        """Add shortcut notation to text"""
        return f"{text} ({shortcut})"

    def _create_win11_button(self, parent, text, command, is_primary=False, width=12):
        """Create a Windows 11 style button for dialogs"""
        if is_primary:
            btn_bg = self.theme.get('accent')
            btn_fg = '#FFFFFF'
            hover_bg = self.theme.get('accent_hover')
            hover_fg = '#FFFFFF'
            pressed_bg = self.theme.get('accent_hover')
        else:
            btn_bg = self.theme.get('bg_elevated')
            btn_fg = self.theme.get('fg')
            hover_bg = self.theme.get('hover')
            hover_fg = self.theme.get('fg')
            pressed_bg = self.theme.get('pressed')

        border_color = self.theme.get('border')

        btn = tk.Button(parent, text=text,
                       command=command,
                       bg=btn_bg,
                       fg=btn_fg,
                       font=("Segoe UI", 9),
                       relief="flat",
                       bd=0,
                       highlightthickness=1,
                       highlightbackground=border_color,
                       highlightcolor=border_color,
                       borderwidth=1,
                       width=width,
                       cursor="hand2",
                       padx=12,
                       pady=6)

        # Windows 11 style hover and press effects
        def on_enter(e):
            e.widget.config(bg=hover_bg, fg=hover_fg, relief="flat")

        def on_leave(e):
            e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

        def on_press(e):
            e.widget.config(bg=pressed_bg, relief="flat")

        def on_release(e):
            e.widget.config(bg=hover_bg, relief="flat")

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)

        return btn

    def _create_source_widgets(self):
        """Create source selection widgets"""
        try:
            print("DEBUG: Creating info frame...")
            # Top description
            info_frame = tk.Frame(self, bg=self.theme.get('bg_elevated'))
            info_frame.pack(fill="x", padx=15, pady=(15, 10))

            tk.Label(info_frame, text="Select folders and files to copy",
                     bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg'), font=("Arial", 12, "bold")).pack(anchor='w')

            tk.Label(info_frame, text="Use checkboxes or Space key to select items. Selected folders include all contents.",
                     bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg_secondary'), font=("Arial", 9)).pack(anchor='w', pady=(2,0))

            print("DEBUG: Creating top controls...")
            # Top controls
            top_controls = tk.Frame(self, bg=self.theme.get('bg_elevated'))
            top_controls.pack(fill="x", padx=15, pady=(0, 10))

            self.add_root_btn = self._create_win11_button(top_controls,
                                                         self.add_shortcut_text("Add Root Folder...", "Ctrl+R"),
                                                         self._choose_root, width=20)
            self.add_root_btn.pack(side="left", padx=2)

            self.sel_all_btn = self._create_win11_button(top_controls,
                                                         self.add_shortcut_text("Select All", "Ctrl+A"),
                                                         self._select_all, width=15)
            self.sel_all_btn.pack(side="left", padx=2)

            self.clear_all_btn = self._create_win11_button(top_controls,
                                                           self.add_shortcut_text("Clear All", "Ctrl+N"),
                                                           self._clear_all, width=15)
            self.clear_all_btn.pack(side="left", padx=2)

            print("DEBUG: Creating tree area...")
            # Tree area
            tree_frame = tk.Frame(self, bg=self.theme.get('bg_elevated'))
            tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

            self.tree = ttk.Treeview(tree_frame, columns=("fullpath",), displaycolumns=())
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self.tree.pack(side="left", fill="both", expand=True)

            self.tree.bind("<<TreeviewOpen>>", self._on_open)
            self.tree.bind("<Button-1>", self._on_click)

            print("DEBUG: Creating status frame...")
            # Bottom shortcut guide
            status_frame = tk.Frame(self, bg=self.theme.get('bg_elevated'))
            status_frame.pack(fill='x', padx=15, pady=(5,0))

            self.status_label = tk.Label(status_frame, text="Ready - Use mouse or keyboard to select items",
                                        bg=self.theme.get('bg_elevated'), fg=self.theme.get('fg_secondary'), font=("Arial", 9))
            self.status_label.pack(side='left')

            print("DEBUG: Creating action buttons...")
            # Action buttons
            action_frame = tk.Frame(self, bg=self.theme.get('bg_elevated'))
            action_frame.pack(fill="x", padx=15, pady=15)

            ok_btn = self._create_win11_button(action_frame,
                                              self.add_shortcut_text("OK", "Enter"),
                                              self._on_ok, is_primary=True, width=12)
            ok_btn.pack(side="right")

            cancel_btn = self._create_win11_button(action_frame,
                                                   self.add_shortcut_text("Cancel", "Esc"),
                                                   self._on_cancel, width=12)
            cancel_btn.pack(side="right", padx=(0, 8))
            
            print("DEBUG: Widget creation complete")
            
        except Exception as e:
            print(f"DEBUG: _create_source_widgets failed: {e}")
            import traceback
            print(f"DEBUG: Traceback: {traceback.format_exc()}")
            raise

    def _format(self, name, checked):
        """Format as checkbox"""
        checkbox = "[✓]" if checked else "[ ]"
        return f"{checkbox} {name}"

    def _toggle_selected(self):
        """Toggle currently selected item (for Space key)"""
        try:
            item = self.tree.focus()
            if item:
                self._toggle_item(item)
                self._update_status()
        except Exception as e:
            print(f"DEBUG: _toggle_selected failed: {e}")

    def _expand_selected(self):
        """Expand selected item"""
        try:
            item = self.tree.focus()
            if item:
                self.tree.item(item, open=True)
        except Exception as e:
            print(f"DEBUG: _expand_selected failed: {e}")

    def _collapse_selected(self):
        """Collapse selected item"""
        try:
            item = self.tree.focus()
            if item:
                self.tree.item(item, open=False)
        except Exception as e:
            print(f"DEBUG: _collapse_selected failed: {e}")

    def _refresh_tree(self):
        """Refresh tree (F5)"""
        try:
            print("DEBUG: Refreshing tree...")
            # Save currently checked items
            checked_paths = []
            for item in self._checked:
                path = self._item_path.get(item)
                if path:
                    checked_paths.append(path)
            
            # Initialize tree
            for item in self.tree.get_children(""):
                self.tree.delete(item)
            self._checked.clear()
            self._item_path.clear()
            
            # Re-add drives
            self._add_drive_roots()
            
            self._update_status("Tree refreshed")
            print("DEBUG: Tree refresh complete")
        except Exception as e:
            print(f"DEBUG: _refresh_tree failed: {e}")
            self._update_status(f"Refresh failed: {e}")

    def _update_status(self, message=None):
        """Update status bar"""
        try:
            if message:
                self.status_label.config(text=message)
            else:
                count = len(self._checked)
                if count == 0:
                    self.status_label.config(text="No items selected")
                elif count == 1:
                    self.status_label.config(text="1 item selected")
                else:
                    self.status_label.config(text=f"{count} items selected")
        except Exception as e:
            print(f"DEBUG: _update_status failed: {e}")

    def _show_source_help(self):
        """Show source selection help"""
        help_text = """Copy Data - Source Selection Help

KEYBOARD SHORTCUTS:
Ctrl + A         Select All Items
Ctrl + N         Clear All Selections  
Ctrl + R         Add Root Folder
Space            Toggle Selected Item
F5               Refresh Tree
F1               Show This Help
Enter            OK (Confirm Selection)
Esc              Cancel

SELECTION BEHAVIOR:
• [✓] Checked items will be included in the copy operation
• [ ] Unchecked items will be skipped
• Selecting a folder includes all its contents
• Only items without checked parents are included (avoids duplicates)

TIPS:
• Use "Add Root Folder" to browse for specific directories
• Check parent folders to include entire directory trees
• Use F5 to refresh if folder contents have changed
        """
        
        try:
            help_dlg = tk.Toplevel(self)
            help_dlg.title("Source Selection Help")
            help_dlg.configure(bg=self.theme.get('bg_elevated'))
            help_dlg.geometry("500x400")
            help_dlg.transient(self)
            help_dlg.grab_set()

            text_widget = tk.Text(help_dlg,
                                 font=("Consolas", 10),
                                 bg=self.theme.get('bg_secondary'),
                                 fg=self.theme.get('fg'),
                                 relief="flat",
                                 wrap="word")
            text_widget.pack(fill="both", expand=True, padx=15, pady=15)

            text_widget.insert("1.0", help_text)
            text_widget.configure(state="disabled")

            btn_frame = tk.Frame(help_dlg, bg=self.theme.get('bg'))
            btn_frame.pack(pady=(0,15))
            close_btn = Win11Dialog._create_dialog_button(btn_frame,
                                                          "Close (Esc)",
                                                          help_dlg.destroy,
                                                          self.theme,
                                                          is_primary=True)
            close_btn.pack()
            
            help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
            help_dlg.focus_set()
            
        except Exception as e:
            print(f"DEBUG: _show_source_help failed: {e}")

    def _add_drive_roots(self):
        """Add drive roots to tree"""
        try:
            print("DEBUG: Adding drive roots...")
            drives = []
            
            # Detect drives with simple method
            for code in range(ord('C'), ord('Z') + 1):
                drive = f"{chr(code)}:\\"
                if os.path.exists(drive):
                    drives.append(drive)
            
            # Add drives to the tree
            for drive in sorted(drives):
                try:
                    item = self.tree.insert("", "end", text=self._format(drive, False), open=False)
                    self._item_path[item] = drive
                    # dummy child for lazy load
                    self.tree.insert(item, "end", text="...")
                    print(f"DEBUG: Added drive: {drive}")
                except Exception as e:
                    print(f"DEBUG: Failed to add drive {drive}: {e}")
                    
            print(f"DEBUG: Added {len(drives)} drives")
        except Exception as e:
            print(f"DEBUG: _add_drive_roots failed: {e}")

    def _choose_root(self):
        """ Select root folder"""
        try:
            print("DEBUG: Choosing root folder...")
            path = filedialog.askdirectory(title="Select Root Folder", parent=self)
            if path:
                print(f"DEBUG: Selected path: {path}")
                item = self.tree.insert("", "end", text=self._format(path, False), open=False)
                self._item_path[item] = path
                self.tree.insert(item, "end", text="...")
                self._update_status(f"Added root: {os.path.basename(path)}")
                print("DEBUG: Root folder added successfully")
            else:
                print("DEBUG: No path selected")
        except Exception as e:
            print(f"DEBUG: _choose_root failed: {e}")
            self._update_status(f"Failed to add root: {e}")

    def _on_open(self, event):
        """Tree item open event"""
        try:
            item = self.tree.focus()
            if not item:
                return
            children = self.tree.get_children(item)
            if len(children) == 1 and self.tree.item(children[0], "text") == "...":
                # populate children lazily
                self.tree.delete(children[0])
                path = self._item_path.get(item)
                parent_is_checked = item in self._checked

                if not path:
                    return
                
                entries = []
                with os.scandir(path) as it:
                    for entry in it:
                        full = entry.path
                        try:
                            if self.is_hidden_fn(full):
                                continue
                        except Exception:
                            pass
                        is_dir = False
                        try:
                            is_dir = entry.is_dir(follow_symlinks=False)
                        except Exception:
                            is_dir = False
                        entries.append((entry.name, is_dir, full))
                
                # sort: directories first, then files
                for name, is_dir, full in sorted(entries, key=lambda t: (not t[1], t[0].lower())):
                    child = self.tree.insert(item, "end", text=self._format(name, parent_is_checked), open=False)
                    self._item_path[child] = full
                    if parent_is_checked:
                        self._checked.add(child)
                    if is_dir:
                        self.tree.insert(child, "end", text="...")
        except Exception as e:
            print(f"DEBUG: _on_open failed: {e}")

    def _on_click(self, event):
        """Tree item click event"""
        try:
            item = self.tree.identify_row(event.y)
            if not item:
                return
            try:
                elem = self.tree.identify("element", event.x, event.y)
            except Exception:
                elem = ""
            elem_str = str(elem).lower()
            # Only toggle when clicking on the text area
            if "text" not in elem_str:
                return
            self._toggle_item(item)
            self._update_status()
        except Exception as e:
            print(f"DEBUG: _on_click failed: {e}")

    def _toggle_item(self, item):
        """Toggle item check state"""
        try:
            text = self.tree.item(item, "text")
            if not text:
                return
            checked = item in self._checked
            new_checked = not checked
            
            # Extract only name part from text (remove checkbox part)
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]  # "[✓] " or "[ ] " remove
            else:
                name = text
            
            self.tree.item(item, text=self._format(name, new_checked))
            
            if new_checked:
                self._checked.add(item)
            else:
                self._checked.discard(item)
                
            # propagate to descendants
            for child in self.tree.get_children(item):
                if self.tree.item(child, "text") == "...":
                    continue
                self._apply_to_descendants(child, new_checked)
        except Exception as e:
            print(f"DEBUG: _toggle_item failed: {e}")

    def _apply_to_descendants(self, item, checked):
        """Apply check state to descendants"""
        try:
            text = self.tree.item(item, "text")
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]
            else:
                name = text
            
            self.tree.item(item, text=self._format(name, checked))
            
            if checked:
                self._checked.add(item)
            else:
                self._checked.discard(item)
                
            for child in self.tree.get_children(item):
                if self.tree.item(child, "text") == "...":
                    continue
                self._apply_to_descendants(child, checked)
        except Exception as e:
            print(f"DEBUG: _apply_to_descendants failed: {e}")

    def _select_all(self):
        """Select all items"""
        try:
            for item in self.tree.get_children(""):
                self._set_checked_recursive(item, True)
            self._update_status()
        except Exception as e:
            print(f"DEBUG: _select_all failed: {e}")

    def _clear_all(self):
        """Clear all selections"""
        try:
            for item in self.tree.get_children(""):
                self._set_checked_recursive(item, False)
            self._update_status()
        except Exception as e:
            print(f"DEBUG: _clear_all failed: {e}")

    def _set_checked_recursive(self, item, checked):
        """ Set check state recursively"""
        try:
            text = self.tree.item(item, "text")
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]
            else:
                name = text
            
            self.tree.item(item, text=self._format(name, checked))
            
            if checked:
                self._checked.add(item)
            else:
                self._checked.discard(item)
                
            for child in self.tree.get_children(item):
                if self.tree.item(child, "text") == "...":
                    continue
                self._set_checked_recursive(child, checked)
        except Exception as e:
            print(f"DEBUG: _set_checked_recursive failed: {e}")

    def _on_ok(self):
        """OK button"""
        try:
            print("DEBUG: _on_ok called")
            paths = []
            # Only include items that don't have a checked ancestor (avoid duplicates)
            for item in list(self._checked):
                p = self._item_path.get(item)
                if not p:
                    continue
                # skip if any ancestor checked
                parent = self.tree.parent(item)
                skip = False
                while parent:
                    if parent in self._checked:
                        skip = True
                        break
                    parent = self.tree.parent(parent)
                if not skip:
                    paths.append(p)
            
            self.selected_paths = paths
            print(f"DEBUG: Selected {len(paths)} paths")
            for path in paths:
                print(f"DEBUG: Selected path: {path}")
            
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
        except Exception as e:
            print(f"DEBUG: _on_ok failed: {e}")
            import traceback
            print(f"DEBUG: Traceback: {traceback.format_exc()}")

    def _on_cancel(self):
        """Cancel button"""
        try:
            print("DEBUG: _on_cancel called")
            self.selected_paths = []
            try:
                self.grab_release()
            except Exception:
                pass
            self.destroy()
        except Exception as e:
            print(f"DEBUG: _on_cancel failed: {e}")

    def on_cancel(self):
        """Cancel handler for ESC key"""
        self._on_cancel()
        
    def on_ok(self):
        """OK handler for Enter key"""
        self._on_ok()

def core_run_backup(user_name, dest_dir, include_hidden=False, include_system=False, log_folder=None, retention_count=2, log_retention_days=30, filters=None):
    """
    ommon backup logic for GUI and scheduler (detailed log + cleanup added)
    ame logic as run_backup() but executes without UI updates
    """
    import os, shutil
    from datetime import datetime
    import time
    import stat

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_name = f"{user_name}_backup_{timestamp}"
    backup_path = os.path.join(dest_dir, backup_name)
    log_timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    log_path = os.path.join(dest_dir, f"{user_name}_backup_{log_timestamp}.log")

    # Delete existing backup folder if exists
    if os.path.exists(backup_path):
        try:
            shutil.rmtree(backup_path, onerror=_handle_rmtree_error_headless)
        except Exception as e:
            pass

    os.makedirs(backup_path, exist_ok=True)

    total_bytes = 0
    files_copied = 0
    folders_processed = 0
    user_home = os.path.join("C:\\Users", user_name)

    if not os.path.exists(user_home):
        raise FileNotFoundError(f"User folder '{user_home}' not found.")

    # Check hidden/system files function (same logic as run_backup)
    def is_hidden_headless(filepath):
        """is_hidden function for headless mode"""
        if not os.path.exists(filepath):
            return False

        # Basic checks
        base = os.path.basename(filepath).lower()
        if base.startswith('.') or base in ('thumbs.db', 'desktop.ini', '$recycle.bin'):
            return True

        if base in ('appdata', 'application data', 'local settings'):
            if not (include_hidden and include_system):
                return True

        try:
            # Check Windows attributes
            if _have_pywin32 and win32api and win32con:
                attrs = win32api.GetFileAttributes(filepath)
                is_hidden_attr = bool(attrs & win32con.FILE_ATTRIBUTE_HIDDEN)
                is_system_attr = bool(attrs & win32con.FILE_ATTRIBUTE_SYSTEM)
                is_reparse = bool(attrs & win32con.FILE_ATTRIBUTE_REPARSE_POINT)
            else:
                try:
                    attrs = _get_file_attributes(filepath)
                except Exception:
                    return False
                is_hidden_attr = bool(attrs & FILE_ATTRIBUTE_HIDDEN)
                is_system_attr = bool(attrs & FILE_ATTRIBUTE_SYSTEM)
                is_reparse = bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)

            if is_reparse:
                return True
            if is_hidden_attr and not include_hidden:
                return True
            if is_system_attr and not include_system:
                return True
        except Exception:
            pass

        return False

    # Use provided filters or default filters for scheduled backup (same as GUI default)
    if filters is None:
        filters = {'include': [], 'exclude': [{'type': 'name', 'pattern': 'onedrive*'}]}

    default_filters = filters

    def matches_filter(filepath, is_dir=False):
        """Check if filepath matches any filter rules"""
        import fnmatch

        try:
            name = os.path.basename(filepath)
            name_l = name.lower()
            ext_l = os.path.splitext(name_l)[1].lstrip('.')
            filepath_lower = filepath.lower()

            def _matches(rule):
                try:
                    rtype = (rule.get('type') or '').lower()
                    patt = (rule.get('pattern') or '').strip()
                    if not patt:
                        return False
                    patt_lower = patt.lower()

                    if rtype == 'ext':
                        if is_dir:
                            return False
                        patt_clean = patt_lower.lstrip('*.')
                        return ext_l == patt_clean

                    elif rtype == 'name':
                        if '*' in patt_lower or '?' in patt_lower:
                            return fnmatch.fnmatch(name_l, patt_lower)
                        else:
                            return name_l == patt_lower

                    elif rtype == 'path':
                        # Simple contains check
                        if patt_lower in filepath_lower:
                            return True
                        # Normalized path matching
                        path_normalized = filepath_lower.replace('\\', '/')
                        patt_normalized = patt_lower.replace('\\', '/')
                        if patt_normalized in path_normalized:
                            return True
                        # Wildcard matching
                        if '*' in patt_normalized or '?' in patt_normalized:
                            return fnmatch.fnmatch(path_normalized, patt_normalized)
                    return False
                except Exception:
                    return False

            # Excludes take precedence
            for rule in default_filters.get('exclude', []) or []:
                if _matches(rule):
                    return True

            # Includes: if any include rules exist, only allow matching files
            inc_rules = default_filters.get('include', []) or []
            if inc_rules and not is_dir:
                for rule in inc_rules:
                    if _matches(rule):
                        return False
                # no include matched -> exclude
                return True
        except Exception:
            pass

        return False

    # Safe file copy function (same as run_backup)
    def copy_file_safe_headless(src, dst, timeout_seconds=30):
        """Safe file copy with timeout (for headless)"""
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            copied = 0
            start_time = time.time()
            buffer_size = 64*1024
            
            with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
                while True:
                    if time.time() - start_time > timeout_seconds:
                        raise TimeoutError(f"File copy timeout: {src}")
                    
                    buf = fsrc.read(buffer_size)
                    if not buf:
                        break
                    fdst.write(buf)
                    copied += len(buf)
            
            try:
                shutil.copystat(src, dst)
            except Exception:
                pass
                
            return copied
        except Exception as e:
            print(f"ERROR copying file {src} -> {dst}: {e}")
            return 0

    # Pre-calculate total size
    def calculate_total_size(root_path):
        """Calculate total size of files to backup"""
        total = 0
        for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=None):
            try:
                original_dirnames = dirnames.copy()
                dirnames.clear()
                for d in original_dirnames:
                    full_dir_path = os.path.join(dirpath, d)
                    try:
                        if not is_hidden_headless(full_dir_path) and not matches_filter(full_dir_path, is_dir=True):
                            dirnames.append(d)
                    except:
                        dirnames.append(d)

                for f in filenames:
                    try:
                        src_file = os.path.join(dirpath, f)
                        if not is_hidden_headless(src_file) and not matches_filter(src_file, is_dir=False):
                            total += os.path.getsize(src_file)
                    except:
                        pass
            except:
                pass
        return total

    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Backup start for '{user_name}' → {backup_path}\n")

        # Write Policy line
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Policy: Hidden={'include' if include_hidden else 'exclude'}, System={'include' if include_system else 'exclude'}, BackupKeep={retention_count}, LogKeep={log_retention_days}\n")

        # Write Filters line (format actual filters from default_filters)
        # Format include filters
        include_list = default_filters.get('include', []) or []
        if include_list:
            include_str = ', '.join([f"{f.get('type', '').lower()}:{f.get('pattern', '')}" for f in include_list])
        else:
            include_str = 'None'

        # Format exclude filters - add default hidden files pattern
        exclude_list = default_filters.get('exclude', []) or []
        exclude_patterns = ['name:^\\...*']  # Default: exclude hidden files starting with dot
        for f in exclude_list:
            pattern_type = f.get('type', '').lower()
            pattern_val = f.get('pattern', '')
            exclude_patterns.append(f"{pattern_type}:{pattern_val}")
        exclude_str = ', '.join(exclude_patterns)

        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Filters: Include=[{include_str}], Exclude=[{exclude_str}]\n")

        # Write separator
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] --------------------------------------------------\n")

        # Calculate and log total size
        try:
            precalc_total = calculate_total_size(user_home)
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Calculating total size... {format_bytes(precalc_total)}\n")
        except Exception as e:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Note: Could not pre-calculate total size: {e}\n")

        try:
            # Same logic as run_backup()
            for dirpath, dirnames, filenames in os.walk(user_home, topdown=True, onerror=None):
                try:
                    # Safe directory filtering
                    original_dirnames = dirnames.copy()
                    dirnames.clear()

                    for d in original_dirnames:
                        full_dir_path = os.path.join(dirpath, d)
                        try:
                            if is_hidden_headless(full_dir_path):
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Skipping hidden directory: {full_dir_path}\n")
                            elif matches_filter(full_dir_path, is_dir=True):
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Skipping filtered directory: {full_dir_path}\n")
                            else:
                                dirnames.append(d)
                        except Exception as e:
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error checking directory {full_dir_path}: {e}\n")
                            dirnames.append(d)  # Include directory on error
                    
                    # Create destination folder
                    rel_dir = os.path.relpath(dirpath, user_home)
                    dest_dir_path = os.path.join(backup_path, rel_dir) if rel_dir != "." else backup_path
                    
                    try:
                        os.makedirs(dest_dir_path, exist_ok=True)
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Created folder: {dest_dir_path}\n")
                        folders_processed += 1
                    except Exception as e:
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: Failed to create folder {dest_dir_path}: {e}\n")
                        continue

                    # Copy files
                    for f in filenames:
                        try:
                            src_file = os.path.join(dirpath, f)
                            dest_file = os.path.join(dest_dir_path, f)

                            # Check hidden files
                            if is_hidden_headless(src_file):
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Skipping hidden file: {src_file}\n")
                                continue

                            # Check filter rules (e.g., onedrive*)
                            if matches_filter(src_file, is_dir=False):
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Skipping filtered file: {src_file}\n")
                                continue

                            # Copy file
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Copying file: {src_file} -> {dest_file}\n")
                            copied_bytes = copy_file_safe_headless(src_file, dest_file)
                            if copied_bytes > 0:
                                files_copied += 1
                                total_bytes += copied_bytes
                                
                        except Exception as e:
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: Failed to copy file {src_file}: {e}\n")
                            continue
                            
                except Exception as e:
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: Failed to process directory {dirpath}: {e}\n")
                    continue

            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Backup completed. Files backed up: {files_copied}, Folders: {folders_processed}, Total size: {format_bytes(total_bytes)}\n")

        except Exception as e:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] CRITICAL ERROR in backup loop: {e}\n")
            raise

        # Browser bookmark backup (same as run_backup)
        try:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Backing up Chrome and Edge browser bookmarks...\n")
            chrome_bookmark_path = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Google', 'Chrome', 'User Data', 'Default', 'Bookmarks')
            edge_bookmark_path = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Microsoft', 'Edge', 'User Data', 'Default', 'Bookmarks')

            if os.path.exists(chrome_bookmark_path):
                shutil.copy(chrome_bookmark_path, os.path.join(dest_dir, 'Chrome_Bookmarks'))
                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Chrome bookmarks backed up.\n")
            else:
                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Chrome bookmark file not found.\n")

            if os.path.exists(edge_bookmark_path):
                shutil.copy(edge_bookmark_path, os.path.join(dest_dir, 'Edge_Bookmarks'))
                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Edge bookmarks backed up.")
            else:
                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] - Edge bookmark file not found.")
        except Exception as e:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Browser bookmark backup error: {e}")

        # Old backup cleanup (same as run_backup)
        try:
            log.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [Old backup cleanup]\n")
            if retention_count <= 0:
                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Backup cleanup is disabled (retention count is 0 or less).\n")
            else:
                # backup_prefix = f"{user_name}_backup_"
                # all_backups = [d for d in os.listdir(dest_dir) if os.path.isdir(os.path.join(dest_dir, d)) and d.startswith(backup_prefix)]
                backup_prefix_lower = f"{user_name.lower()}_backup_"
                all_backups = [d for d in os.listdir(dest_dir) if os.path.isdir(os.path.join(dest_dir, d)) and d.lower().startswith(backup_prefix_lower)]
                 
                if len(all_backups) > retention_count:
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Found {len(all_backups)} backups. Keeping the {retention_count} most recent.\n")
                    all_backups.sort()
                    backups_to_delete = all_backups[:-retention_count]
                    for backup_dir_name in backups_to_delete:
                        dir_to_delete = os.path.join(dest_dir, backup_dir_name)
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   Deleting old backup: {dir_to_delete}\n")
                        try:
                            shutil.rmtree(dir_to_delete, onerror=lambda func, path, exc_info: _handle_rmtree_error_headless(func, path, exc_info, log, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                        except Exception as e:
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   ERROR: Could not delete {dir_to_delete}: {e}\n")
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Cleanup finished.\n")
                else:
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Found {len(all_backups)} backups. No old backups to clean up (retention policy: keep {retention_count}).")
        except Exception as e:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: An error occurred during backup cleanup: {e}")

        # Old log cleanup (same as run_backup)
        try:
            log.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [Old log cleanup]\n")
            if log_retention_days <= 0:
                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Log cleanup disabled (retention days <= 0)\n")
            else:
                if not os.path.isdir(dest_dir):
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Log folder does not exist: {dest_dir}\n")
                else:
                    now = time.time()
                    cutoff = now - (log_retention_days * 86400)
                    cutoff_date = datetime.fromtimestamp(cutoff).strftime('%Y-%m-%d %H:%M:%S')
                    
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Scanning '{dest_dir}' for logs older than {log_retention_days} days...\n")
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Cutoff time: {cutoff_date}\n")
                    
                    deleted_count = 0
                    try:
                        all_files = os.listdir(dest_dir)
                        log_files = [f for f in all_files if f.endswith('.log')]
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Log files found: {len(log_files)}\n")
                        
                        for filename in log_files:
                            filepath = os.path.join(dest_dir, filename)
                            try:
                                if os.path.isfile(filepath):
                                    file_mtime = os.path.getmtime(filepath)
                                    file_date = datetime.fromtimestamp(file_mtime).strftime('%Y-%m-%d %H:%M:%S')
                                    age_days = (now - file_mtime) / 86400
                                    
                                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   File: {filename}\n")
                                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]     Modified: {file_date}\n")
                                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]     Age: {age_days:.1f} days\n")
                                    log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]     Should delete: {file_mtime < cutoff}\n")
                                    
                                    if file_mtime < cutoff:
                                        try:
                                            os.remove(filepath)
                                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]     ✓ Deleted: {filename}\n")
                                            deleted_count += 1
                                        except Exception as e:
                                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]     ✗ Delete failed: {e}\n")
                            except Exception as e:
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]   Error checking {filename}: {e}\n")
                        
                        if deleted_count > 0:
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Log cleanup completed. Deleted {deleted_count} files.\n")
                        else:
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] No old log files found to delete.\n")
                            
                    except Exception as e:
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error listing files: {e}\n")
        except Exception as e:
            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error during log cleanup: {e}\n")

    return backup_path, log_path


# 2. Add necessary helper functions

def format_bytes(size):
    """Converts bytes to a human-readable format (KB, MB, GB, etc.)."""
    try:
        size = float(size)
    except (ValueError, TypeError):
        return "N/A"
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024.0:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} PB"


def _handle_rmtree_error_headless(func, path, exc_info, log, timestamp):
    """
    Error handler for shutil.rmtree (for headless mode - logs to file)
    """
    if isinstance(exc_info[1], PermissionError):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
            log.write(f"[{timestamp}] INFO: Removed read-only and successfully deleted {path}\n")
        except Exception as e:
            log.write(f"[{timestamp}] ERROR: Could not delete {path} even after removing read-only flag: {e}\n")
    else:
        log.write(f"[{timestamp}] ERROR: Deletion failed for {path}: {exc_info[1]}\n")

if __name__ == "__main__":
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    class HeadlessLoggerQueue:
        """A dummy queue that prints important messages to stdout for headless logging."""
        def put(self, task_tuple):
            try:
                task, value = task_tuple
                if task in ('log', 'update_status'):
                    print(f"[INFO] {value}")
                elif task == 'show_error':
                    print(f"[ERROR] {value}")
            except Exception:
                pass # Ignore malformed messages or other UI tasks
        def get_nowait(self, *args, **kwargs): raise queue.Empty
        def task_done(self, *args, **kwargs): pass

    def _run_headless_if_requested():

        parser = argparse.ArgumentParser()
        parser.add_argument("--schedule-config", help="Path to schedule configuration file")
        parser.add_argument("--user", help="User name to back up")
        parser.add_argument("--dest", help="Destination folder for backup")
        parser.add_argument("--include-hidden", action="store_true")
        parser.add_argument("--include-system", action="store_true")
        parser.add_argument("--retention", type=int, default=2, help="Number of backups to keep")
        parser.add_argument("--log-retention", type=int, default=30, help="Number of days to keep logs")
        parser.add_argument("--log-folder", default="logs")
        parser.add_argument("--filters", type=str, help="JSON-encoded filter configuration")

        # Use parse_known_args to ignore other arguments
        args, unknown = parser.parse_known_args()

        # Check if schedule config file is provided
        if args.schedule_config:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] Scheduled backup start using config: {args.schedule_config}")

            try:
                # Load configuration from file
                import json
                if not os.path.exists(args.schedule_config):
                    print(f"[{timestamp}] ERROR: Configuration file not found: {args.schedule_config}")
                    sys.exit(1)

                with open(args.schedule_config, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                print(f"[{timestamp}] Configuration loaded successfully")
                print(f"[{timestamp}] Task: {config.get('task_name')}")
                print(f"[{timestamp}] User: {config.get('user')}")
                print(f"[{timestamp}] Destination: {config.get('dest')}")
                print(f"[{timestamp}] Retention: {config.get('retention_count')} backups, {config.get('log_retention_days')} days logs")

                # Run backup with loaded configuration
                backup_path, log_path = core_run_backup(
                    config.get('user'),
                    config.get('dest'),
                    include_hidden=config.get('include_hidden', False),
                    include_system=config.get('include_system', False),
                    log_folder=args.log_folder,
                    retention_count=config.get('retention_count', 2),
                    log_retention_days=config.get('log_retention_days', 30),
                    filters=config.get('filters')
                )
                print(f"[{timestamp}] Backup finished → {backup_path}")
                print(f"[{timestamp}] Log saved at {log_path}")

            except Exception as e:
                print(f"[{timestamp}] ERROR during scheduled backup: {e}")
                import traceback
                traceback.print_exc()
                sys.exit(1)

            sys.exit(0)

        # Legacy mode: direct command line arguments
        elif args.user and args.dest:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] Headless backup start for user '{args.user}'")

            # Modification: Display actual delivered retention values ​​in log
            print(f"[{timestamp}] Command line args: retention={args.retention}, log_retention={args.log_retention}")

            # Parse filters from JSON if provided
            import json
            filters = None
            if args.filters:
                try:
                    filters = json.loads(args.filters)
                    print(f"[{timestamp}] Filters loaded: {filters}")
                except Exception as e:
                    print(f"[{timestamp}] WARNING: Failed to parse filters JSON: {e}")
                    filters = None

            try:
                backup_path, log_path = core_run_backup(
                    args.user,
                    args.dest,
                    include_hidden=args.include_hidden,
                    include_system=args.include_system,
                    log_folder=args.log_folder,
                    retention_count=args.retention,  # Using values ​​passed from the command line
                    log_retention_days=args.log_retention,  # Using values ​​passed from the command line
                    filters=filters
                )
                print(f"[{timestamp}] Backup finished → {backup_path}")
                print(f"[{timestamp}] Log saved at {log_path}")

            except Exception as e:
                print(f"[{timestamp}] ERROR during headless backup: {e}")

            sys.exit(0)



    # Run headless mode if requested before any UI/elevation logic
    _run_headless_if_requested()

    # If running as a frozen executable (PyInstaller), the binary's manifest
    # controls elevation (e.g. --uac-admin). In that case we should not attempt
    # to relaunch via ShellExecuteW because Windows will already handle UAC.
    if getattr(sys, 'frozen', False):
        # When frozen, simply run the app. Any elevation prompt will be handled
        # by the executable's manifest if present.
        try:
            app = App()
            app.mainloop()
        except Exception as e:
            # Log exception to stderr so console builds can show it, and exit.
            import traceback, sys as _sys
            traceback.print_exc()
            try:
                _sys.exit(1)
            except Exception:
                os._exit(1)
    else:
        # Non-frozen script: try to relaunch as admin if not already running elevated.
        # But skip elevation when running in auto-backup headless mode (handled above and exited).
        if not is_admin():
            try:
                executable = sys.executable
                params = ' '.join(f'"{a}"' for a in sys.argv[1:]) if len(sys.argv) > 1 else None
                params = subprocess.list2cmdline(sys.argv)
                ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, params, None, 1)
                try:
                    if int(ret) <= 32:
                        print(f"Elevation failed or was cancelled (ShellExecuteW returned {ret})")
                except Exception:
                    pass
            except Exception as e:
                print(f"Unable to elevate privileges: {e}")
            try:
                sys.exit(0)
            except SystemExit:
                os._exit(0)
        else:
            app = App()
            app.mainloop()