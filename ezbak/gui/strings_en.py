"""English strings for Translator (moved verbatim from ezBAK.py in P5-1).

Key order is fixed (C3). Duplicate keys are kept as they were: the later
value wins. Removing them is a separate defect commit.
"""

EN = {
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
    'dest_folder_missing': 'The destination folder does not exist',

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
    'please_select_user': 'Please select a user.',
    'data_will_be_overwritten': ' will be overwritten. Do you want to continue?',
    'valid_drive_letter': 'Please enter a valid drive letter (ex: Z)',
    'password': 'Password',
    'password_required_when_username': 'Password is required when specifying username.',
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
│ Ctrl + 1~9      Quick Select User               │
│ Escape          (Not Working) Cancel Current Operation│
│ Ctrl + Q        Exit Program            │
└─────────────────────────────────────────┘

Tips :  Shortcuts are not case-sensitive.""",
    'help_exit_esc': 'Exit (Esc)',
    'data_restore_cancelled': 'Data restore cancelled by user.',
    'driver_restore_cancelled': 'Driver restore cancelled by user.',
    'cross_user_restore_warning': 'Cross-User Restore Warning\n\nYou are about to restore:\n  {0}\'s backup\nto:\n  {1}\'s account\n\nFolder: {2}\n\nThis may cause data conflicts or permission issues.\n\nDo you want to continue?',
    'select_backup_folder': 'Select Backup Folder',
    'select_backup_folder_msg': 'Select a backup folder to restore:',
    'no_browser_profiles': 'No browser profile data was found to back up.',
    'confirm_browser_profiles_title': 'Browser Profiles Backup',
    'confirm_browser_profiles': 'Back up the full browser profiles of user \'{0}\'?\n\nSource: {1}',
    # U3-4: sidebar, status chip and home cards of the redesigned window (ADR 0004)
    'nav_home': 'Home',
    'nav_backup': 'Backup',
    'nav_restore': 'Restore',
    'nav_schedule': 'Scheduled Backup',
    'nav_tools': 'Tools',
    'nav_logs': 'Logs',
    'nav_settings': 'Settings',
    'nav_help': 'Help',
    'status_idle': 'Idle',
    'status_working': 'Working {0}%',
    'home_start_backup': 'Start Backup',
    'home_selected_user': 'Selected User',
    'home_filters': '{0} filters',
    'home_retention': 'Backups to Keep',
    'home_log_days': 'Log Retention (days)',
    'home_system_card': 'System',
    'home_system_name': 'System Name',
    'home_user_count': 'Users',
    'home_hidden_files': 'Hidden Files',
    'home_system_files': 'System Files',
    'home_included': 'Included',
    'home_excluded': 'Excluded',
    'home_loading': 'Loading...',
    # U5-2: card titles of the backup and restore pages
    'backup_main': 'Back Up User Data',
    'backup_options': 'Backup Options',
    'backup_extra': 'Additional Backups',
    'restore_main': 'Restore User Data',
    # U5-3: restore page
    'restore_hint': "Restores a backup folder into the selected user's profile. You are asked before existing data is overwritten, then you choose the backup folder.",
    'restore_drivers_hint': 'Restore Drivers installs the drivers from a driver backup folder.',
    # U5-4: scheduled backup page
    'schedule_main': 'Scheduled Backup',
    'schedule_hint': 'Registers a Windows scheduled task that backs up the selected user without opening the window. The dialog creates a schedule or deletes an existing one.',
    # U5-5: tools page card titles
    'tools_disk': 'Space and Copy',
    'tools_system': 'System and Files',
    'tools_nas': 'Network Drive (NAS)',
    # U5-7: help page buttons (were the fixed English Help menu entries)
    'help_shortcuts': 'Keyboard Shortcuts (F1)',
    'help_github': 'GitHub Releases',
    # U7-3: home data cards (speed, last backup, next schedule, free space)
    'home_speed_card': 'Copy Speed',
    'home_speed_idle': 'No copy running',
    'home_last_backup_card': 'Last Backup',
    'home_last_backup_none': 'No backup recorded yet',
    'home_mode_gui': 'Manual',
    'home_mode_headless': 'Scheduled',
    'home_next_schedule_card': 'Next Scheduled Backup',
    'home_next_schedule_none': 'No scheduled backup',
    'home_free_space': 'Free Space',
}
