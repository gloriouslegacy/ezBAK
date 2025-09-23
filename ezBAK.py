import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import traceback
import os
import shutil
import threading
import queue
from datetime import datetime, timedelta
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

# ㄴ Check Space : 실행시 '계산중 등 메시지 필요'
# ㄴ Keep Log for : X
# ㄴ Schedule Backup : 로그 생성시 시분초 표시됨(기존 백업 날짜까지), NTUSER데이터 생성됨(4파일) 

# Try to import pywin32 (win32api/win32con). When running as a bundled exe,
# pywin32 might not be included by PyInstaller; provide a ctypes-based
# fallback that uses GetFileAttributesW so attribute checks still work.
# 필수/안전

#   중지/취소: 진행 중 작업을 즉시 중단k
#    일시정지/재개: 대용량 복사 시 일시 정지/재개
# 	용량/공간 체크: 소스 총 용량 계산 및 대상 여유공간 검사
# 무결성 검증: 완료 후 해시 검증(선택 항목만)으로 복사 검증
#   오래된 백업 정리: 보관 정책(예: 최근 N개 유지)으로 자동 삭제

# 편의/생산성

# 빠른 폴더 선택: 문서/바탕화면/다운로드/사진/음악/동영상 빠른 선택
# 	포함/제외 규칙: 확장자/이름/경로 패턴 기반 필터 관리
#   백업 폴더 열기: 마지막 작업의 대상 폴더 즉시 열기ㅎㅎ
#   최근 로그 열기: 상세 로그 저장 폴더 열기
#   설정 내보내기/가져오기: 숨김/시스템/규칙 등 구성 백업
# 고급/자동화

# 	스케줄 백업: 작업 스케줄러 등록/해제(UI로 간단히 주기/시간 설정)
# 압축 백업: 선택 항목을 ZIP으로 백업(옵션: 분할/암호)
# 	설치 프로그램 목록 내보내기: winget export로 현재 설치 목록 백업
# 드라이버 목록 내보내기: pnputil/driverquery 결과를 텍스트로 저장
# 	브라우저 프로필 백업: Chrome/Edge/Firefox 프로필 전체(확장/설정 포함)
# UI/유틸

# 테마/언어 전환: Light/Dark 및 언어 스위치
# 업데이트 확인: GitHub 릴리즈 페이지 열기
# 	네트워크 공유 연결: 백업 대상 NAS 경로 연결/해제 도우미
# 우선순위 권장

# 중지/취소, 일시정지/재개, 용량/공간 체크
# 빠른 폴더 선택, 포함/제외 규칙, 백업 폴더/로그 열기
# 스케줄 백업, 압축 백업, 무결성 검증
# 현 구조에 쉽게 추가 가능한 항목

# 중지/취소, 일시정지/재개: 현재 스레드/큐 구조에 플래그만 추가해 제어 가능
# 빠른 폴더 선택: 버튼에서 미리 정의된 경로를 소스 선택 다이얼로그에 추가
# 백업 폴더 열기/최근 로그 열기: os.startfile 호출만 추가
# 용량/공간 체크: 기존 compute_total_bytes와 shutil.disk_usage로 즉시 구현 가능


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
import time

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
    from datetime import datetime
    import os
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
        self.hidden_mode = 'include' if include_hidden else 'exclude'
        self.system_mode = 'include' if include_system else 'exclude'
        self.retention_count = retention_count
        self.log_retention_days = log_retention_days
        # Headless mode does not use complex filters from the UI
        self.filters = {'include': [], 'exclude': []}

    def is_hidden(self, filepath):
        if not os.path.exists(filepath):
            return False

        base = os.path.basename(filepath).lower()
        if base in ('appdata', 'application data', 'local settings'):
            if not (self.hidden_mode == 'include' and self.system_mode == 'include'):
                return True

        try:
            if _have_pywin32 and win32api:
                attrs = win32api.GetFileAttributes(filepath)
                is_hidden_attr = bool(attrs & win32con.FILE_ATTRIBUTE_HIDDEN)
                is_system_attr = bool(attrs & win32con.FILE_ATTRIBUTE_SYSTEM)
                is_reparse = bool(attrs & win32con.FILE_ATTRIBUTE_REPARSE_POINT)
            else:
                attrs = _get_file_attributes(filepath)
                is_hidden_attr = bool(attrs & FILE_ATTRIBUTE_HIDDEN)
                is_system_attr = bool(attrs & FILE_ATTRIBUTE_SYSTEM)
                is_reparse = bool(attrs & FILE_ATTRIBUTE_REPARSE_POINT)
            
            if is_reparse: return True
            if is_hidden_attr and self.hidden_mode == 'exclude': return True
            if is_system_attr and self.system_mode == 'exclude': return True
        except Exception as e:
            log_message(f"WARN: is_hidden check failed for {filepath}: {e}")
        return False

    def copy_file_with_progress(self, src, dst, progress_callback, buffer_size=64*1024):
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            copied = 0
            with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
                while True:






                    buf = fsrc.read(buffer_size)
                    if not buf: break
                    fdst.write(buf)
                    copied += len(buf)
                    progress_callback(len(buf))
            try: shutil.copystat(src, dst)
            except Exception: pass
            return copied
        except Exception as e:
            log_message(f"ERROR: copying file {src} -> {dst}: {e}")
            return 0

    def compute_total_bytes(self, root_path):
        total = 0
        for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=lambda e: None):
            dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if not self.is_hidden(fp):
                    try: total += os.path.getsize(fp)
                    except Exception: pass
        return total

    def backup_browser_bookmarks(self, backup_path):
        log_message("Backing up Chrome and Edge browser bookmarks...")
        chrome_bookmark_path = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Google', 'Chrome', 'User Data', 'Default', 'Bookmarks')
        edge_bookmark_path = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Microsoft', 'Edge', 'User Data', 'Default', 'Bookmarks')

        try:
            if os.path.exists(chrome_bookmark_path):
                shutil.copy(chrome_bookmark_path, os.path.join(backup_path, 'Chrome_Bookmarks'))
                log_message("- Chrome bookmarks backed up.")
            else:
                log_message("- Chrome bookmark file not found.")
        except Exception as e:
            log_message(f"- Chrome bookmark backup error: {e}")

        try:
            if os.path.exists(edge_bookmark_path):
                shutil.copy(edge_bookmark_path, os.path.join(backup_path, 'Edge_Bookmarks'))
                log_message("- Edge bookmarks backed up.")
            else:
                log_message("- Edge bookmark file not found.")
        except Exception as e:
            log_message(f"- Edge bookmark backup error: {e}")

    def _cleanup_old_backups(self, backup_path, user_name):
        try:
            if self.retention_count <= 0:
                log_message("\n[Old backup cleanup]")
                log_message("Backup cleanup is disabled (retention count is 0 or less).")
                return

            log_message("\n[Old backup cleanup]")
            backup_prefix = f"{user_name}_backup_"
            
            all_backups = [d for d in os.listdir(backup_path) if os.path.isdir(os.path.join(backup_path, d)) and d.startswith(backup_prefix)]
            
            if len(all_backups) > self.retention_count:
                log_message(f"Found {len(all_backups)} backups. Keeping the {self.retention_count} most recent.")
                all_backups.sort()
                backups_to_delete = all_backups[:-self.retention_count]
                for backup_dir_name in backups_to_delete:
                    dir_to_delete = os.path.join(backup_path, backup_dir_name)
                    log_message(f"  Deleting old backup: {dir_to_delete}")
                    shutil.rmtree(dir_to_delete, onerror=_handle_rmtree_error)
                log_message("Cleanup finished.")
            else:
                log_message(f"Found {len(all_backups)} backups. No old backups to clean up (retention policy: keep {self.retention_count}).")
        except Exception as e:
            log_message(f"ERROR: An error occurred during backup cleanup: {e}")






    def run_backup(self, backup_path):
        """Executes the actual backup logic with hidden file exclusion and byte-based progress."""
        bytes_copied = 0
        files_processed = 0
        try:
            user_name = self.user_var.get()
            user_profile_path = os.path.join("C:", os.sep, "Users", user_name)

            if not os.path.exists(user_profile_path):
                self.message_queue.put(('show_error', f"User folder '{user_profile_path}' not found."))
                self.write_detailed_log(f"Backup failed: user folder not found at {user_profile_path}")
                return

        # --- 무조건 시작 로그 ---
            self.write_detailed_log(f"Backup start for user '{user_name}': {user_profile_path} -> {backup_path}")
            self.message_queue.put(('log', f"Backing up user: {user_name}"))

            timestamp = datetime.now().strftime("%Y-%m-%d")
            destination_dir = os.path.join(backup_path, f"{user_name}_backup_{timestamp}")

            if os.path.exists(destination_dir):
                try:
                    shutil.rmtree(destination_dir)
                    self.write_detailed_log(f"Deleted existing backup folder '{destination_dir}'.")
                except Exception as e:
                    self.write_detailed_log(f"Failed to delete existing backup folder: {e}")

            os.makedirs(destination_dir, exist_ok=True)
            self.write_detailed_log(f"Created new backup folder '{destination_dir}'.")

        # Compute total bytes (excluding hidden)
            total_bytes = self.compute_total_bytes(user_profile_path)
            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Total bytes to copy: {total_bytes}")

        # Free space pre-check
            try:
                free = shutil.disk_usage(backup_path).free
            except Exception:
                free = self.get_free_space(backup_path)
            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space. Required={total_bytes}, Free={free}")
                self.message_queue.put(('show_error',
                    f"Not enough free space in destination.\n"
                    f"Required: {self.format_bytes(total_bytes)}\nAvailable: {self.format_bytes(free)}"))
                return

            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            for dirpath, dirnames, filenames in os.walk(user_profile_path, topdown=True,
                                                        onerror=lambda e: self.write_detailed_log(f"os.walk error: {e}")):
                dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                rel_dir = os.path.relpath(dirpath, user_profile_path)
                dest_dir = os.path.join(destination_dir, rel_dir) if rel_dir != '.' else destination_dir
                os.makedirs(dest_dir, exist_ok=True)
                self.write_detailed_log(f"Created folder: {dest_dir}")

                for f in filenames:
                    src_file = os.path.join(dirpath, f)
                    if self.is_hidden(src_file):
                        self.write_detailed_log(f"Skipping hidden file: {src_file}")
                        continue
                    dest_file = os.path.join(dest_dir, f)
                    self.write_detailed_log(f"Copying file: {src_file} -> {dest_file}")
                    self.copy_file_with_progress(src_file, dest_file, progress_cb)
                    files_processed += 1

        # --- 무조건 완료 로그 ---
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log(f"Backup completed. Files backed up: {files_processed}, "
                                    f"Total size: {self.format_bytes(bytes_copied)}")
            self.write_detailed_log("Backup successfully completed.")

            self.backup_browser_bookmarks(destination_dir)
            self._cleanup_old_backups(backup_path, user_name)
            self._cleanup_old_logs(backup_path, self.log_retention_days)

            self.message_queue.put(('log', "Backup complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Backup complete!"))

        except Exception as e:
            self.write_detailed_log(f"Backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during backup: {e}"))
        finally:
            self.close_log_file()
            try:
                self.message_queue.put(('update_progress', self.progress_bar['maximum']))
            except Exception:
                pass
            self.message_queue.put(('enable_buttons', None))





class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.log_retention_days_var = tk.StringVar(value='30')
        self.title("ezBAK")
        try:
            # Try the title icon first (bundled with PyInstaller)
            self.iconbitmap(resource_path('./icon/ezbak_title.ico'))
        except tk.TclError:
            try:
                # Fallback to the regular icon
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except tk.TclError:
                pass 
        self.geometry("1024x439")
        self.configure(bg="#2D3250")

        # UI elements setup
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Configure a new style for a green progress bar
        self.style.configure("green.Horizontal.TProgressbar", background="#4CAF50")

        # Queue for thread-safe UI updates
        self.message_queue = queue.Queue()
        self.process_queue()

        # Options: control inclusion for Hidden and System attributes (defaults: exclude)
        self.hidden_mode_var = tk.StringVar(value='exclude')   # 'include' or 'exclude'
        self.system_mode_var = tk.StringVar(value='exclude')   # 'include' or 'exclude'
        self.retention_count_var = tk.StringVar(value='2') # Number of backups to keep
        self.log_retention_days_var = tk.StringVar(value='30') # Number of days to keep logs
        # Pattern filters (persisted): {'include': [{'type': 'ext'|'name'|'path', 'pattern': str}], 'exclude': [...]} 
        self.filters = {'include': [], 'exclude': []}
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

        # Initial message for the log box (UI short notice) with only the final sentence bolded
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.config(state="normal")
        # Insert standard notice text
        self.log_text.insert(tk.END, f"[{timestamp}] \nNotice:\n- Select User\n- Hidden and System attributes excluded.\n- %AppData% folder is not backed up.\n- Detailed log will be saved to a file during operations.\n- ")
        # Insert the single bolded disclaimer
        self.log_text.insert(tk.END, "Use of this program is the sole responsibility of the user.\n", 'bold')
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.update_idletasks()

    def create_widgets(self):
        # Header
        # header_frame = tk.Frame(self, bg="#424769")
        # header_frame.pack(fill="x", pady=10)

        # title_label = tk.Label(header_frame, text="ezBAK", font=("Arial", 11, "bold"), bg="#424769", fg="white")
        #title_label.pack(side="left", padx=5)
        # title_label.pack(expand=True)

        # Main frame
        main_frame = tk.Frame(self, bg="#2D3250", padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # Small options row under the title (right-aligned) for Hidden/System
        options_row = tk.Frame(main_frame, bg="#2D3250")
        options_row.pack(fill="x", pady=(2,6))

        # right-aligned opt_frame placed in the options_row
        opt_frame = tk.Frame(options_row, bg="#2D3250")
        opt_frame.pack(side="right")

        # Build widgets so visible LTR order is: Hidden: Include Exclude   [gap]   System: Include Exclude
        # We pack side='right' in reverse to achieve the left-to-right appearance.
        radio_font = ("Arial", 10)

        # System group (right side)
        system_exclude = tk.Radiobutton(opt_frame, text="Exclude", variable=self.system_mode_var, value='exclude', bg="#2D3250", fg="white", selectcolor="#2D3250", font=radio_font, activebackground="#2D3250", activeforeground="white", bd=0)
        system_exclude.pack(side="right", padx=(2,4))
        system_include = tk.Radiobutton(opt_frame, text="Include", variable=self.system_mode_var, value='include', bg="#2D3250", fg="white", selectcolor="#2D3250", font=radio_font, activebackground="#2D3250", activeforeground="white", bd=0)
        system_include.pack(side="right", padx=(0,6))
        system_label = tk.Label(opt_frame, text="System:", bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
        system_label.pack(side="right", padx=(6,4))

        # visual separator between groups
        sep = ttk.Separator(opt_frame, orient='vertical')
        sep.pack(side="right", fill='y', padx=8, pady=2)

        # Hidden group (left side)
        hidden_exclude = tk.Radiobutton(opt_frame, text="Exclude", variable=self.hidden_mode_var, value='exclude', bg="#2D3250", fg="white", selectcolor="#2D3250", font=radio_font, activebackground="#2D3250", activeforeground="white", bd=0)
        hidden_exclude.pack(side="right", padx=(2,4))
        hidden_include = tk.Radiobutton(opt_frame, text="Include", variable=self.hidden_mode_var, value='include', bg="#2D3250", fg="white", selectcolor="#2D3250", font=radio_font, activebackground="#2D3250", activeforeground="white", bd=0)
        hidden_include.pack(side="right", padx=(0,6))
        hidden_label = tk.Label(opt_frame, text="Hidden:", bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
        hidden_label.pack(side="right", padx=(4,2))

        # visual separator between groups
        sep2 = ttk.Separator(opt_frame, orient='vertical')
        sep2.pack(side="right", fill='y', padx=8, pady=2)

        # Log retention count group
        log_retention_info_label = tk.Label(opt_frame, text="days (0=disable)", bg="#2D3250", fg="gray", font=radio_font)
        log_retention_info_label.pack(side="right", padx=(2, 4))
        log_retention_spinbox = tk.Spinbox(opt_frame, from_=0, to=365, textvariable=self.log_retention_days_var, width=4)
        log_retention_spinbox.pack(side="right", padx=(0, 2))
        log_retention_label = tk.Label(opt_frame, text="Keep Logs for:", bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
        log_retention_label.pack(side="right", padx=(6, 4))

        # visual separator
        sep3 = ttk.Separator(opt_frame, orient='vertical')
        sep3.pack(side="right", fill='y', padx=8, pady=2)

        # Retention count group (far left)
        retention_info_label = tk.Label(opt_frame, text="(0=keep all)", bg="#2D3250", fg="gray", font=radio_font)
        retention_info_label.pack(side="right", padx=(2, 4))
        retention_spinbox = tk.Spinbox(opt_frame, from_=0, to=99, textvariable=self.retention_count_var, width=4)
        retention_spinbox.pack(side="right", padx=(0, 2))
        retention_label = tk.Label(opt_frame, text="Backups to Keep:", bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
        retention_label.pack(side="right", padx=(6, 4))

        # persist when changed
        try:
            self.hidden_mode_var.trace_add('write', lambda *args: self.save_settings())
        except Exception:
            pass
        try:
            self.system_mode_var.trace_add('write', lambda *args: self.save_settings())
        except Exception:
            pass
        try:
            self.retention_count_var.trace_add('write', lambda *args: self.save_settings())
        except Exception:
            pass        
        try:
            self.log_retention_days_var.trace_add('write', lambda *args: self.save_settings())
        except Exception:
            pass

        # User selection row
        user_selection_frame = tk.Frame(main_frame, bg="#2D3250")
        user_selection_frame.pack(fill="x", pady=6)

        # persist when changed
        self.hidden_mode_var.trace_add('write', lambda *args: self.save_settings())
        self.system_mode_var.trace_add('write', lambda *args: self.save_settings())
        self.retention_count_var.trace_add('write', lambda *args: self.save_settings())
        self.log_retention_days_var.trace_add('write', lambda *args: self.save_settings())
        ttk.Label(user_selection_frame, text="Select User:", font=("Arial", 12), background="#2D3250", foreground="white").pack(side="left", padx=5)

        self.user_var = tk.StringVar()
        self.user_combo = ttk.Combobox(user_selection_frame, textvariable=self.user_var, state="readonly", width=30)
        self.user_combo.pack(side="left", padx=5, fill="x", expand=True)
        self.load_users()

    # Hidden/System radio controls removed per user request.
    # Defaults remain: hidden/system = 'exclude' (set in __init__).

        # Button Grid Frame - Using grid layout for better alignment
        button_grid_frame = tk.Frame(main_frame, bg="#2D3250")
        button_grid_frame.pack(fill="x", pady=(0, 10))

        # Configure grid columns to be uniform
        for i in range(5):
            button_grid_frame.columnconfigure(i, weight=1, uniform="button")

        # Button styling
        button_font_size = 8
        button_width = 14
        button_padx = 5
        button_pady = 5

        # Row 1: Main Data Operations
        self.backup_btn = tk.Button(button_grid_frame, text="Backup User Data", bg="#FF5733", fg="white", 
                                  font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                  relief="flat", width=button_width, command=self.start_backup_thread)
        self.backup_btn.grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        self.backup_btn.bind("<Enter>", lambda e: e.widget.config(bg="#FF7D5A"))
        self.backup_btn.bind("<Leave>", lambda e: e.widget.config(bg="#FF5733"))

        self.restore_btn = tk.Button(button_grid_frame, text="Restore User Data", bg="#336BFF", fg="white", 
                                   font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                   relief="flat", width=button_width, command=self.start_restore_thread)
        self.restore_btn.grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        self.restore_btn.bind("<Enter>", lambda e: e.widget.config(bg="#5A8DFF"))
        self.restore_btn.bind("<Leave>", lambda e: e.widget.config(bg="#336BFF"))

        self.copy_btn = tk.Button(button_grid_frame, text="Copy Data", bg="#FFA500", fg="white", 
                                font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                relief="flat", width=button_width, command=self.copy_files)
        self.copy_btn.grid(row=2, column=2, padx=2, pady=2, sticky="ew")
        self.copy_btn.bind("<Enter>", lambda e: e.widget.config(bg="#FFB733"))
        self.copy_btn.bind("<Leave>", lambda e: e.widget.config(bg="#FFA500"))

        self.driver_backup_btn = tk.Button(button_grid_frame, text="Backup Drivers", bg="#FF5733", fg="white", 
                                         font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                         relief="flat", width=button_width, command=self.start_driver_backup_thread)
        self.driver_backup_btn.grid(row=0, column=3, padx=2, pady=2, sticky="ew")
        self.driver_backup_btn.bind("<Enter>", lambda e: e.widget.config(bg="#FF7D5A"))
        self.driver_backup_btn.bind("<Leave>", lambda e: e.widget.config(bg="#FF5733"))

        self.driver_restore_btn = tk.Button(button_grid_frame, text="Restore Drivers", bg="#336BFF", fg="white", 
                                          font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                          relief="flat", width=button_width, command=self.start_driver_restore_thread)
        self.driver_restore_btn.grid(row=0, column=4, padx=2, pady=2, sticky="ew")
        self.driver_restore_btn.bind("<Enter>", lambda e: e.widget.config(bg="#5A8DFF"))
        self.driver_restore_btn.bind("<Leave>", lambda e: e.widget.config(bg="#336BFF"))

        # Row 2: Utility Operations
        self.devmgr_btn = tk.Button(button_grid_frame, text="Device Manager", bg="#9A9B0B", fg="white", 
                                  font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                  relief="flat", width=button_width, command=self.open_device_manager)
        self.devmgr_btn.grid(row=1, column=3, padx=2, pady=2, sticky="ew")
        self.devmgr_btn.bind("<Enter>", lambda e: e.widget.config(bg="#B4B50C"))
        self.devmgr_btn.bind("<Leave>", lambda e: e.widget.config(bg="#9A9B0B"))

        self.file_explorer_btn = tk.Button(button_grid_frame, text="File Explorer", bg="#7D98A1", fg="white", 
                                         font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                         relief="flat", width=button_width, command=self.open_file_explorer)
        self.file_explorer_btn.grid(row=1, column=4, padx=2, pady=2, sticky="ew")
        self.file_explorer_btn.bind("<Enter>", lambda e: e.widget.config(bg="#92B0BB"))
        self.file_explorer_btn.bind("<Leave>", lambda e: e.widget.config(bg="#7D98A1"))

        self.save_log_btn = tk.Button(button_grid_frame, text="Save Log", bg="#4CAF50", fg="white", 
                                    font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                    relief="flat", width=button_width, command=self.save_log)
        self.save_log_btn.grid(row=1, column=2, padx=2, pady=2, sticky="ew")
        self.save_log_btn.bind("<Enter>", lambda e: e.widget.config(bg="#6EC571"))
        self.save_log_btn.bind("<Leave>", lambda e: e.widget.config(bg="#4CAF50"))

        # Check Space button
        self.check_space_btn = tk.Button(button_grid_frame, text="Check Space", bg="#8E44AD", fg="white", 
                                        font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady, 
                                        relief="flat", width=button_width, command=self.check_space)
        self.check_space_btn.grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        self.check_space_btn.bind("<Enter>", lambda e: e.widget.config(bg="#A569BD"))
        self.check_space_btn.bind("<Leave>", lambda e: e.widget.config(bg="#8E44AD"))

        # Exit button removed per request; replaced with Browser Profiles backup button
        self.browser_profiles_btn = tk.Button(button_grid_frame, text="Backup Browser Profiles", bg="#1E88E5", fg="white",
                                             font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
                                             relief="flat", width=button_width, command=self.start_browser_profiles_backup_thread)
        self.browser_profiles_btn.grid(row=1, column=0, padx=2, pady=2, sticky="ew")
        self.browser_profiles_btn.bind("<Enter>", lambda e: e.widget.config(bg="#42A5F5"))
        self.browser_profiles_btn.bind("<Leave>", lambda e: e.widget.config(bg="#1E88E5"))

        # Row 3: Scheduling
        self.schedule_btn = tk.Button(button_grid_frame, text="Schedule Backup", bg="#2E7D32", fg="white",
                                      font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
                                      relief="flat", width=button_width, command=self.schedule_backup)
        self.schedule_btn.grid(row=2, column=0, padx=2, pady=2, sticky="ew")
        self.schedule_btn.bind("<Enter>", lambda e: e.widget.config(bg="#388E3C"))
        self.schedule_btn.bind("<Leave>", lambda e: e.widget.config(bg="#2E7D32"))

        # Winget export (installed apps)
        self.winget_export_btn = tk.Button(button_grid_frame, text="Export Apps", bg="#00695C", fg="white",
                                          font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
                                          relief="flat", width=button_width, command=self.start_winget_export_thread)
        self.winget_export_btn.grid(row=2, column=1, padx=2, pady=2, sticky="ew")
        self.winget_export_btn.bind("<Enter>", lambda e: e.widget.config(bg="#00897B"))
        self.winget_export_btn.bind("<Leave>", lambda e: e.widget.config(bg="#00695C"))

        # Filters button
        self.filters_btn = tk.Button(button_grid_frame, text="Filters", bg="#455A64", fg="white",
                                     font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
                                     relief="flat", width=button_width, command=self.open_filter_manager)
        self.filters_btn.grid(row=0, column=2, padx=2, pady=2, sticky="ew")
        self.filters_btn.bind("<Enter>", lambda e: e.widget.config(bg="#607D8B"))
        self.filters_btn.bind("<Leave>", lambda e: e.widget.config(bg="#455A64"))

        # NAS connect/disconnect
        self.nas_connect_btn = tk.Button(button_grid_frame, text="Connect NAS", bg="#00796B", fg="white",
                                         font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
                                         relief="flat", width=button_width, command=self.open_nas_connect_dialog)
        self.nas_connect_btn.grid(row=2, column=3, padx=2, pady=2, sticky="ew")
        self.nas_connect_btn.bind("<Enter>", lambda e: e.widget.config(bg="#009688"))
        self.nas_connect_btn.bind("<Leave>", lambda e: e.widget.config(bg="#00796B"))

        self.nas_disconnect_btn = tk.Button(button_grid_frame, text="Disconnect NAS", bg="#B71C1C", fg="white",
                                            font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
                                            relief="flat", width=button_width, command=self.open_nas_disconnect_dialog)
        self.nas_disconnect_btn.grid(row=2, column=4, padx=2, pady=2, sticky="ew")
        self.nas_disconnect_btn.bind("<Enter>", lambda e: e.widget.config(bg="#D32F2F"))
        self.nas_disconnect_btn.bind("<Leave>", lambda e: e.widget.config(bg="#B71C1C"))

        # Progress Bar and Status Label
        self.progress_frame = tk.Frame(main_frame, bg="#2D3250")
        self.progress_frame.pack(fill="x", pady=(10, 0))

        # Use the newly defined green style
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient="horizontal", mode="determinate", style="green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", expand=True)

        self.status_label = ttk.Label(self.progress_frame, text="Select User to begin", font=("Arial", 10), background="#2D3250", foreground="white")
        self.status_label.pack(pady=(5, 0))

        # Log text box (UI shows only summarized logs)
        log_frame = tk.Frame(main_frame, bg="#424769", bd=1, relief="sunken")
        # keep log box moderately small (7 lines)
        log_frame.pack(fill="x", expand=False, pady=10)

        scrollbar = tk.Scrollbar(log_frame)
        self.log_text = tk.Text(log_frame, height=7, bg="#424769", fg="white", relief="flat", bd=0, font=("Consolas", 10), state="disabled", yscrollcommand=scrollbar.set)
        # Configure tag for bold text (used for the single bolded disclaimer line)
        self.log_text.tag_configure('bold', font=("Consolas", 10, "bold"))

        scrollbar.pack(side="right", fill="y")
        # make it a moderate area (7 lines)
        self.log_text.pack(fill="x", expand=False, padx=5, pady=5)

        # Scrollbar configuration
        scrollbar.config(command=self.log_text.yview)

        # Github link label
        self.github_label = tk.Label(main_frame, text="gloriouslegacy", font=("Arial", 10, "underline"), fg="#2196F3", bg="#2D3250", cursor="hand2")
        self.github_label.pack(pady=(5, 0))
        self.github_label.bind("<Button-1>", self.open_github_link)
    

    # Hidden/System options removed from UI as requested; defaults remain 'exclude'

    def open_github_link(self, event):
        webbrowser.open_new("https://github.com/gloriouslegacy/ezBAK/releases")

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

                # After opening the new log, clean up old logs in the same folder.
                # try:
                #     retention_days = int(self.log_retention_days_var.get())
                # except (ValueError, TypeError):
                #     retention_days = 30 # Fallback
                # self._cleanup_old_logs(base_folder, retention_days)

                # write selected backup policies at top of detailed log
                try:
                    self._log_file.write(f"[POLICY] Hidden={self.hidden_mode_var.get()} System={self.system_mode_var.get()}\n")
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
            
    # def _cleanup_old_logs(log_folder, retention_days):
    #     if not os.path.isdir(log_folder) or retention_days <= 0:
    #         return
    #     now = time.time()
    #     cutoff = now - (retention_days * 86400)
    #     log_message(f"[old log cleanup in '{log_folder}']", log_folder=log_folder)
    #     log_message(f"scanning for log files older than {retention_days} days...", log_folder=log_folder)
    #     print(f"[old log cleanup in '{log_folder}']")
    #     print(f"scanning for log files older than {retention_days} days...")
    #     for filename in os.listdir(log_folder):
    #         if filename.endswith(".log"):
    #             filepath = os.path.join(log_folder, filename)
    #             try:
    #                 if os.path.isfile(filepath) and os.path.getmtime(filepath) < cutoff:
    #                     os.remove(filepath)
    #                     write_detailed_log(f"  - deleting old log: {filename}", log_folder=log_folder)
    #                     print(f"  - deleting old log: {filename}")
    #             except Exception as e:
    #                 log_message(f"  - error: could not process or delete log file {filename}: {e}", log_folder=log_folder)
    #                 print(f"  - error: could not process or delete log file {filename}: {e}")

    def process_queue(self):
        """Processes messages from the message queue to update the GUI."""
        try:
            while True:
                task, value = self.message_queue.get_nowait()
                if task == 'log':
                    # UI shows only summarized logs
                    now = time.time()
                    if now - self._last_ui_log_ts >= self.ui_log_throttle_secs:
                        self._last_ui_log_ts = now
                        self.log(value)
                elif task == 'update_progress':
                    # value is bytes copied
                    try:
                        self.progress_bar['value'] = value
                        # update percentage in status_label
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
                        # show percentage only
                        self.status_label.config(text=f"Progression rate: {percent}%")
                    except Exception:
                        pass
                elif task == 'update_progress_max':
                    try:
                        self.progress_bar['maximum'] = value
                        # reset percent display
                        self.status_label.config(text="진행률: 0%")
                    except Exception:
                        pass
                elif task == 'update_status':
                    self.status_label.config(text=value)
                elif task == 'set_progress_mode':
                    self.progress_bar['mode'] = value
                elif task == 'start_progress':
                    self.progress_bar.start()
                elif task == 'stop_progress':
                    self.progress_bar.stop()
                elif task == 'enable_buttons':
                    self.set_buttons_state("normal")
                elif task == 'show_error':
                    messagebox.showerror("Error", value)
                elif task == 'open_folder':
                    try:
                        os.startfile(value)
                    except Exception:
                        pass
                self.message_queue.task_done()
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue)

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

    def save_log(self):
        """Saves the UI log content to a text file (UI only)."""
        if not self.log_text.get("1.0", tk.END).strip():
            messagebox.showinfo("Info", "No log content to save.")
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
            messagebox.showerror("Error", f"Error occurred while saving UI log: {e}")

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
            def _norm(p):
                try:
                    return os.path.normcase(p).replace('\\', '/').lower()
                except Exception:
                    return str(p).lower()
            path_norm = _norm(filepath)

            def _matches(rule):
                try:
                    rtype = (rule.get('type') or '').lower()
                    patt = (rule.get('pattern') or '')
                    if not patt:
                        return False
                    if rtype == 'ext':
                        if is_dir:
                            return False
                        # allow patterns like "tmp" or "*.tmp"
                        patt_norm = patt.lstrip('.').lower()
                        return fnmatch.fnmatch(ext_l, patt_norm)
                    elif rtype == 'name':
                        return fnmatch.fnmatch(name_l, patt.lower())
                    elif rtype == 'path':
                        return fnmatch.fnmatch(path_norm, _norm(patt))
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

    # Helper: walk and sum bytes excluding hidden files/dirs
    def compute_total_bytes(self, root_path):
        total = 0
        for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=lambda e: None):
            # skip hidden directories
            dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if self.is_hidden(fp):
                    continue
                try:
                    total += os.path.getsize(fp)
                except Exception:
                    pass
        return total

    # --- User Data Backup & Restore Functions ---
    def start_backup_thread(self):
        """Starts the backup process in a separate thread."""
        if not self.user_var.get():
            messagebox.showwarning("Warning", "Please select a user.")
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
        """Executes the actual backup logic with hidden file exclusion and byte-based progress."""
        bytes_copied = 0
        try:
            user_name = self.user_var.get()
            user_profile_path = os.path.join("C:", os.sep, "Users", user_name)

            if not os.path.exists(user_profile_path):
                self.message_queue.put(('show_error', f"User folder '{user_profile_path}' not found."))
                return

            self.write_detailed_log(f"Backup start: {user_profile_path} -> {backup_path}")
            self.message_queue.put(('log', f"Backing up user: {user_name}"))

            timestamp = datetime.now().strftime("%Y-%m-%d")
            destination_dir = os.path.join(backup_path, f"{user_name}_backup_{timestamp}")

            if os.path.exists(destination_dir):
                try:
                    shutil.rmtree(destination_dir)
                    self.write_detailed_log(f"Deleted existing backup folder '{destination_dir}'.")
                except Exception as e:
                    self.write_detailed_log(f"Failed to delete existing backup folder: {e}")

            os.makedirs(destination_dir, exist_ok=True)
            self.write_detailed_log(f"Created new backup folder '{destination_dir}'.")

            # Compute total bytes (excluding hidden)
            total_bytes = self.compute_total_bytes(user_profile_path)
            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Total bytes to copy: {total_bytes}")
            # Free space pre-check (backup destination)
            try:
                free = shutil.disk_usage(backup_path).free
            except Exception:
                free = self.get_free_space(backup_path)
            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space for backup. Required={total_bytes} Free={free}")
                self.message_queue.put(('show_error', f"Not enough free space in destination.\nRequired: {self.format_bytes(total_bytes)}\nAvailable: {self.format_bytes(free)}"))
                return

            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                # throttle UI progress updates
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            files_processed = 0
            for dirpath, dirnames, filenames in os.walk(user_profile_path, topdown=True, onerror=lambda e: None):
                dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                rel_dir = os.path.relpath(dirpath, user_profile_path)
                dest_dir = os.path.join(destination_dir, rel_dir) if rel_dir != "." else destination_dir
                os.makedirs(dest_dir, exist_ok=True)
                self.write_detailed_log(f"Created folder: {dest_dir}")

                for f in filenames:
                    src_file = os.path.join(dirpath, f)
                    if self.is_hidden(src_file):
                        self.write_detailed_log(f"Skipping hidden file: {src_file}")
                        continue
                    dest_file = os.path.join(dest_dir, f)
                    self.write_detailed_log(f"Copying file: {src_file} -> {dest_file}")
                    copied = self.copy_file_with_progress(src_file, dest_file, progress_cb)
                    files_processed += 1
                    # occasionally show UI summary
                    if files_processed % 50 == 0:
                        self.message_queue.put(('log', f"Copied {files_processed} files..."))
                        self.write_detailed_log(f"{files_processed} files copied so far; bytes_copied={bytes_copied}")
            # ensure final progress update
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log("User folder backup completed.")
            self.backup_browser_bookmarks(backup_path)
            self.write_detailed_log("Backup successfully completed.")
            self._cleanup_old_backups(backup_path, user_name)
            self.message_queue.put(('log', "Backup complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Backup complete!"))
            self.message_queue.put(('open_folder', destination_dir))

        except Exception as e:
            self.write_detailed_log(f"Backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during backup: {e}"))
        finally:
            self.close_log_file()
            # ensure progress bar completes
            try:
                self.message_queue.put(('update_progress', self.progress_bar['maximum']))
            except Exception:
                pass
            self.message_queue.put(('enable_buttons', None))

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
            backup_prefix = f"{user_name}_backup_"
            
            all_backups = [d for d in os.listdir(backup_path) if os.path.isdir(os.path.join(backup_path, d)) and d.startswith(backup_prefix)]
            
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
            messagebox.showwarning("Warning", "Please select a user.")
            return

        response = messagebox.askyesno("Warning", f"Existing data for '{self.user_var.get()}' will be overwritten. Do you want to continue?")
        if not response:
            self.message_queue.put(('log', "Restore process cancelled."))
            return

        backup_folder = filedialog.askdirectory(title="Select Backed-up Folder")
        if not backup_folder:
            self.message_queue.put(('log', "Restore folder selection cancelled."))
            return

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
        try:
            user_name = self.user_var.get()
            
            if backup_folder.lower().endswith(".log"):
                self.message_queue.put(('show_error',
                    f"선택한 경로가 로그 파일입니다.\n\n"
                    f"'{backup_folder}'\n\n"
                    f"백업 폴더(예: test_backup_2025-09-22)를 선택해야 합니다."))
                self.write_detailed_log("Restore failed: .log file selected instead of backup folder.")
                return

            if os.path.isfile(backup_folder):
                self.message_queue.put(('show_error',
                    f"선택한 경로가 파일입니다. 폴더를 선택하세요:\n{backup_folder}"))
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

            # --- 무조건 시작 로그 ---
            self.message_queue.put(('log', f"Selected backup folder: {backup_folder}"))
            self.write_detailed_log(f"Restore start for user '{user_name}' from {backup_folder}")

            # backup_folders = [d for d in os.listdir(backup_folder) if d.startswith(f"{user_name}_backup_")]
            backup_folders = [d for d in os.listdir(backup_folder) if os.path.isdir(os.path.join(backup_folder, d)) and d.startswith(f"{user_name}_backup_")]
            
            if not backup_folders:
                self.message_queue.put(('show_error', f"No backup data found for '{user_name}' in:\n{backup_folder}"))
                self.write_detailed_log("Restore failed: no matching backup folder.")
                return

            source_dir = os.path.join(backup_folder, sorted(backup_folders, reverse=True)[0])
            destination_dir = os.path.join("C:", os.sep, "Users", user_name)
            self.write_detailed_log(f"Restoring from {source_dir} to {destination_dir}")

            total_bytes = self.compute_total_bytes(source_dir)
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

        # --- 무조건 완료 로그 ---
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log(f"Restore completed. Files restored: {files_processed}, "
                                    f"Total size: {self.format_bytes(bytes_copied)}")
            self.write_detailed_log("Restore successfully completed.")

            self.restore_browser_bookmarks(backup_folder)
            self.message_queue.put(('log', "Restore complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Restore complete!"))

        except Exception as e:
            self.write_detailed_log(f"Restore error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during restore: {e}"))
        finally:
            self.close_log_file()
            try:
                self.message_queue.put(('update_progress', self.progress_bar['maximum']))
            except Exception:
                pass
            self.message_queue.put(('enable_buttons', None))





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

    # --- Copy Files & Folders Functions (New) ---
    def copy_files(self):
        """Prompts for source and destination using a checkbox tree, then starts the copy process."""
        self.message_queue.put(('log', "Copy process initiated."))

        # 1) Open a checkbox tree dialog to select files/folders
        try:
            dlg = SelectSourcesDialog(self, is_hidden_fn=self.is_hidden)
            self.wait_window(dlg)
            source_paths_list = dlg.selected_paths if getattr(dlg, 'selected_paths', None) else []
        except Exception as e:
            self.message_queue.put(('log', f"Source selection failed: {e}"))
            return

        if not source_paths_list:
            self.message_queue.put(('log', "Source selection cancelled."))
            return

        # 2) Ask for destination folder
        destination_dir = filedialog.askdirectory(title="Select Destination Folder")
        if not destination_dir:
            self.message_queue.put(('log', "Destination selection cancelled."))
            return

        # open log file in destination folder
        self.open_log_file(destination_dir, "copy")

        self.message_queue.put(('log', "Copy process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['mode'] = "determinate"
        self.message_queue.put(('update_status', "Starting copy process..."))

        # Start the copy thread
        copy_thread = threading.Thread(target=self.run_copy_thread, args=(source_paths_list, destination_dir), daemon=True)
        copy_thread.start()

    def format_bytes(self, size):
        try:
            size = float(size)
        except Exception:
            return str(size)
        for unit in ["B", "KB", "MB", "GB", "TB", "PB"]:
            if size < 1024.0 or unit == "PB":
                return f"{size:.2f} {unit}"
            size /= 1024.0

    def get_free_space(self, path):
        try:
            return shutil.disk_usage(path).free
        except Exception:
            # Try drive root as fallback
            try:
                drive = os.path.splitdrive(path)[0] or path
                root = drive if drive.endswith(':') else (drive + ':')
                return shutil.disk_usage(root + os.sep).free
            except Exception:
                return 0

    def schedule_backup(self):
        try:
            dlg = ScheduleBackupDialog(self)
            self.wait_window(dlg)
            if getattr(dlg, 'action', None) == 'create':
                data = dlg.result or {}
                self.create_scheduled_task(
                    task_name=data.get('task_name'),
                    schedule=data.get('schedule'),
                    time_str=data.get('time_str'),
                    days=data.get('days'),
                    user=data.get('user'),
                    dest=data.get('dest'),
                    include_hidden=data.get('include_hidden', False),
                    include_system=data.get('include_system', False)
                )
            elif getattr(dlg, 'action', None) == 'delete':
                name = getattr(dlg, 'task_name', None)
                if name:
                    self.delete_scheduled_task(name)
        except Exception as e:
            try:
                messagebox.showerror("Error", f"Schedule dialog error: {e}")
            except Exception:
                pass

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

    def create_scheduled_task(self, task_name, schedule, time_str, days, user, dest, include_hidden=False, include_system=False):
        if not task_name or not user or not dest or not schedule or not time_str:
            try:
                messagebox.showwarning("Warning", "Please fill all required fields (Task Name, User, Destination, Schedule, Time).")
            except Exception:
                pass
            return
        tr_cmd = self._build_auto_backup_command(user, dest, include_hidden, include_system)
        if not tr_cmd:
            try:
                messagebox.showerror("Error", "Failed to build task command.")
            except Exception:
                pass
            return
        base = ["schtasks", "/Create", "/TN", task_name, "/TR", tr_cmd, "/ST", time_str, "/F"]
        if schedule.lower() == 'daily':
            base += ["/SC", "DAILY"]
        elif schedule.lower() == 'weekly':
            base += ["/SC", "WEEKLY"]
            if days:
                base += ["/D", ','.join(days)]
        else:
            try:
                messagebox.showerror("Error", f"Unsupported schedule: {schedule}")
            except Exception:
                pass
            return
        # Try with highest privileges, fallback without if it fails
        cmd1 = base + ["/RL", "HIGHEST"]
        try:
            res = subprocess.run(cmd1, shell=True, capture_output=True, text=True)
            if res.returncode != 0:
                # fallback without /RL HIGHEST
                res2 = subprocess.run(base, shell=True, capture_output=True, text=True)
                if res2.returncode != 0:
                    raise RuntimeError(res2.stderr or res2.stdout)
        except Exception as e:
            try:
                messagebox.showerror("Error", f"Failed to create task:\n{e}")
            except Exception:
                pass
            self.message_queue.put(('log', f"Schedule create failed: {e}"))
            return
        try:
            messagebox.showinfo("Task Scheduler", f"Task '{task_name}' created.")
        except Exception:
            pass
        self.message_queue.put(('log', f"Task created: {task_name} -> {schedule} {time_str}"))

    def delete_scheduled_task(self, task_name):
        if not task_name:
            try:
                messagebox.showwarning("Warning", "Task Name is required to delete.")
            except Exception:
                pass
            return
        try:
            res = subprocess.run(["schtasks", "/Delete", "/TN", task_name, "/F"], shell=True, capture_output=True, text=True)
            if res.returncode != 0:
                raise RuntimeError(res.stderr or res.stdout)
        except Exception as e:
            try:
                messagebox.showerror("Error", f"Failed to delete task:\n{e}")
            except Exception:
                pass
            self.message_queue.put(('log', f"Schedule delete failed: {e}"))
            return
        try:
            messagebox.showinfo("Task Scheduler", f"Task '{task_name}' deleted.")
        except Exception:
            pass
        self.message_queue.put(('log', f"Task deleted: {task_name}"))

    def check_space(self):
        # 1) Select sources via checkbox tree
        try:
            dlg = SelectSourcesDialog(self, is_hidden_fn=self.is_hidden, title="Check Required Space")
            self.wait_window(dlg)
            sources = dlg.selected_paths if getattr(dlg, 'selected_paths', None) else []
        except Exception as e:
            messagebox.showerror("Error", f"Source selection failed: {e}")
            return
        if not sources:
            self.message_queue.put(('log', "Space check cancelled (no sources)."))
            return

        # 2) Select destination folder
        dest = filedialog.askdirectory(title="Select Destination Folder to Check")
        if not dest:
            self.message_queue.put(('log', "Space check cancelled (no destination)."))
            return

        # 3) Compute total bytes (blocking in UI thread by design for simplicity)
        total = 0
        for p in sources:
            if os.path.isdir(p):
                total += self.compute_total_bytes(p)
            elif os.path.isfile(p) and not self.is_hidden(p):
                try:
                    total += os.path.getsize(p)
                except Exception:
                    pass

        free = self.get_free_space(dest)
        msg = (
            f"Required: {self.format_bytes(total)}\n"
            f"Available: {self.format_bytes(free)}\n"
            f"Margin: {self.format_bytes(free - total)}\n"
        )
        if total <= free:
            messagebox.showinfo("Space Check", "Enough space available.\n\n" + msg)
            self.message_queue.put(('log', f"Space OK. Required={self.format_bytes(total)} Free={self.format_bytes(free)}"))
        else:
            messagebox.showwarning("Space Check", "Not enough space.\n\n" + msg)
            self.message_queue.put(('log', f"Space LOW. Required={self.format_bytes(total)} Free={self.format_bytes(free)}"))

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
                messagebox.showerror('Error', f'Filter manager error: {e}')
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
        dlg.title("Network Share - Connect")
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg="#2D3250")
        dlg.geometry("680x380")
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
        tk.Label(dlg, text=help_text, bg="#2D3250", fg="white", justify='left', anchor='w', font=("Consolas", 10), wraplength=wrap).pack(fill='x', padx=12, pady=(12, 8))

        form = tk.Frame(dlg, bg="#2D3250")
        form.pack(fill='both', expand=True, padx=16)

        row = 0
        def add_label(text):
            nonlocal row
            tk.Label(form, text=text, bg="#2D3250", fg="white").grid(row=row, column=0, sticky='w', pady=6)
        def add_entry(var, width=42, show=None):
            nonlocal row
            e = tk.Entry(form, textvariable=var, width=width, show=show)
            e.grid(row=row, column=1, sticky='w', pady=6)
            row += 1
            return e

        # UNC Path
        add_label("UNC Path (ex: \\NAS\\share)")
        unc_var = tk.StringVar()
        add_entry(unc_var)

        # Map drive option
        map_var = tk.BooleanVar(value=True)
        def_drive = self.find_free_drive_letter() or 'Z'
        drive_var = tk.StringVar(value=def_drive)

        map_frame = tk.Frame(form, bg="#2D3250")
        tk.Checkbutton(map_frame, text="Map as drive letter", variable=map_var, bg="#2D3250", fg="white", selectcolor="#2D3250").pack(side='left')
        tk.Label(map_frame, text="Letter:", bg="#2D3250", fg="white").pack(side='left', padx=(12,4))
        drive_entry = tk.Entry(map_frame, textvariable=drive_var, width=6)
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
        add_label("Username (Optional)")
        user_var = tk.StringVar()
        add_entry(user_var)

        add_label("Password (Required if Username is provided)")
        pwd_var = tk.StringVar()
        add_entry(pwd_var, show='*')

        # Persistent
        persist_var = tk.BooleanVar(value=True)
        persist_chk = tk.Checkbutton(form, text="Reconnect at logon (Persistent)", variable=persist_var, bg="#2D3250", fg="white", selectcolor="#2D3250")
        persist_chk.grid(row=row, column=0, columnspan=2, sticky='w', pady=8)
        row += 1

        # Buttons
        btns = tk.Frame(dlg, bg="#2D3250")
        btns.pack(fill='x', pady=(6,12))

        result = {'ok': False}
        def on_ok():
            unc = (unc_var.get() or '').strip()
            if not unc or not unc.startswith('\\\\'):
                messagebox.showerror("Error", "It should start with UNC Path \\\\ (ex: \\\\Server\\share)", parent=dlg)
                return
            user = (user_var.get() or '').strip()
            pwd = (pwd_var.get() or '')
            if user and not pwd:
                messagebox.showerror("Error", "Password is required if User is provided.", parent=dlg)
                return
            drv = None
            if map_var.get():
                d = (drive_var.get() or '').strip().rstrip(':').upper()
                if not (len(d) == 1 and 'A' <= d <= 'Z'):
                    messagebox.showerror("Error", "Please enter a valid drive letter (ex: Z)", parent=dlg)
                    return
                drv = d + ':'
            result.update(ok=True, unc=unc, drive=drv, user=user, pwd=pwd, persistent=persist_var.get())
            dlg.destroy()
        def on_cancel():
            dlg.destroy()
        ok_btn = tk.Button(btns, text="OK", command=on_ok, bg="#4CAF50", fg="white", relief="flat", width=10)
        ok_btn.pack(side='right', padx=8)
        cancel_btn = tk.Button(btns, text="Cancel", command=on_cancel, bg="#7D98A1", fg="white", relief="flat", width=10)
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
                    if messagebox.askyesno("Connected", "Open the share now?"):
                        os.startfile(drive if drive else unc)
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS connect failed: {msg}"))
                try:
                    messagebox.showerror("NAS Connect", f"Failed to connect.\n\n{msg}")
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS connect error: {e}"))
            try:
                messagebox.showerror("Error", f"NAS connect error: {e}")
            except Exception:
                pass
            if not unc:
                self.message_queue.put(('log', "NAS connect cancelled (no path)."))
                return
            if not unc.startswith('\\\\'):
                messagebox.showerror("Error", "UNC path must start with \\\\ (e.g., \\\\server\\share)")
                return

            map_drive = messagebox.askyesno("Map Drive", "Map this share to a drive letter?")
            drive = None
            if map_drive:
                default_drive = self.find_free_drive_letter() or 'Z'
                val = simpledialog.askstring("Drive Letter", f"Drive letter (A-Z). Default: {default_drive}", parent=self)
                if not val:
                    val = default_drive
                val = (val.strip().rstrip(':').upper() if isinstance(val, str) else str(val)).rstrip(':')
                if not (len(val) == 1 and 'A' <= val <= 'Z'):
                    messagebox.showerror("Error", "Invalid drive letter.")
                    return
                drive = val + ':'

            username = simpledialog.askstring("Username", "Username (leave empty to use current credentials):", parent=self)
            password = None
            if username:
                password = simpledialog.askstring("Password", "Password (required when username is set)", show='*', parent=self)
                if not password:
                    messagebox.showerror("Error", "Password is required when specifying username.")
                    return

            persistent = messagebox.askyesno("Persistence", "Reconnect at sign-in? (Persistent)")

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
                    if messagebox.askyesno("Connected", "Open the share now?"):
                        os.startfile(drive if drive else unc)
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS connect failed: {msg}"))
                try:
                    messagebox.showerror("NAS Connect", f"Failed to connect.\n\n{msg}")
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS connect error: {e}"))
            try:
                messagebox.showerror("Error", f"NAS connect error: {e}")
            except Exception:
                pass

    def open_nas_disconnect_dialog(self):
        """Open a large dialog for NAS disconnect with full command examples, then run 'net use ... /delete'."""
        # Build dialog UI
        dlg = tk.Toplevel(self)
        dlg.title("Network Share - Disconnect")
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg="#2D3250")
        dlg.geometry("680x260")
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
        tk.Label(dlg, text=help_text, bg="#2D3250", fg="white", justify='left', anchor='w', font=("Consolas", 10), wraplength=wrap).pack(fill='x', padx=12, pady=(12, 8))

        form = tk.Frame(dlg, bg="#2D3250")
        form.pack(fill='both', expand=True, padx=16)

        tk.Label(form, text="Target (Driver Letter ex: Z: or UNC ex: \\Server\\Share)", bg="#2D3250", fg="white").grid(row=0, column=0, sticky='w', pady=6)
        target_var = tk.StringVar()
        tk.Entry(form, textvariable=target_var, width=42).grid(row=0, column=1, sticky='w', pady=6)

        btns = tk.Frame(dlg, bg="#2D3250")
        btns.pack(fill='x', pady=(6,12))

        result = {'ok': False}
        def on_ok():
            val = (target_var.get() or '').strip()
            if not val:
                messagebox.showerror("Error", "Input the target (Drive Letter or UNC)", parent=dlg)
                return
            result.update(ok=True, target=val)
            dlg.destroy()
        def on_cancel():
            dlg.destroy()
        ok_btn = tk.Button(btns, text="OK", command=on_ok, bg="#4CAF50", fg="white", relief="flat", width=10)
        ok_btn.pack(side='right', padx=8)
        cancel_btn = tk.Button(btns, text="Cancel", command=on_cancel, bg="#7D98A1", fg="white", relief="flat", width=10)
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
                    messagebox.showerror("Error", "Invalid drive letter or UNC path.")
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
                    messagebox.showinfo("NAS Disconnect", f"Disconnected: {label}")
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS disconnect failed: {msg}"))
                try:
                    messagebox.showerror("NAS Disconnect", f"Failed to disconnect.\n\n{msg}")
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS disconnect error: {e}"))
            try:
                messagebox.showerror("Error", f"NAS disconnect error: {e}")
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
                    messagebox.showerror("Error", "Invalid drive letter or UNC path.")
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
                    messagebox.showinfo("NAS Disconnect", f"Disconnected: {label}")
                except Exception:
                    pass
            else:
                msg = err or out or f"Failed with code {res.returncode}"
                self.message_queue.put(('log', f"NAS disconnect failed: {msg}"))
                try:
                    messagebox.showerror("NAS Disconnect", f"Failed to disconnect.\n\n{msg}")
                except Exception:
                    pass
        except Exception as e:
            self.message_queue.put(('log', f"NAS disconnect error: {e}"))
            try:
                messagebox.showerror("Error", f"NAS disconnect error: {e}")
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

            cmd = ['winget', 'export', '--output', output_path, '--include-versions',
                   '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity']
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
                self.write_detailed_log(err.strip())

            rc = process.returncode
            if rc != 0:
                # Attempt fallback using 'winget list'
                self.write_detailed_log(f"winget export failed with code {rc}. Attempting fallback: winget list.")
                fb_name = f"winget_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                fb_path = os.path.join(folder, fb_name)
                try:
                    p2 = subprocess.run(['winget', 'list', '--accept-source-agreements', '--disable-interactivity'],
                                        shell=False, capture_output=True, text=True, encoding='utf-8', errors='ignore')
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
                        self.write_detailed_log(f"winget list fallback failed with code {p2.returncode}. stderr={p2.stderr.strip()}")
                except Exception as e2:
                    self.write_detailed_log(f"winget list fallback error: {e2}")
                # If reached here, both export and fallback failed
                raise RuntimeError(f"winget export failed with code {rc}")

            self.write_detailed_log(f"Winget export completed: {output_path}")
            self.message_queue.put(('log', f"Winget export completed: {output_path}"))
            self.message_queue.put(('open_folder', os.path.dirname(output_path)))
        except FileNotFoundError:
            self.write_detailed_log("winget not found on system.")
            self.message_queue.put(('show_error', "winget is not available. Install 'App Installer' from Microsoft Store."))
        except Exception as e:
            self.write_detailed_log(f"Winget export error: {e}")
            # Add friendly guidance for common causes
            self.message_queue.put(('log', "Winget export failed. Possible causes: old winget version, disabled export feature (older versions), or corrupted sources. Try:\n- Update 'App Installer' from Microsoft Store\n- winget source reset --force\n- winget settings and enable export (older versions)\n- Re-run Export Apps"))
            self.message_queue.put(('show_error', f"An error occurred during winget export: {e}"))
        finally:
            self.close_log_file()
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('update_status', "Winget export complete!"))
            self.message_queue.put(('enable_buttons', None))

    def run_copy_thread(self, source_paths, destination_dir):
        """Performs the actual file and folder copying in a separate thread (byte progress)."""
        bytes_copied = 0
        try:
            # compute total bytes
            total_bytes = 0
            for path in source_paths:
                if os.path.isdir(path):
                    total_bytes += self.compute_total_bytes(path)
                elif os.path.isfile(path) and not self.is_hidden(path):
                    try:
                        total_bytes += os.path.getsize(path)
                    except Exception:
                        pass

            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Copy total bytes: {total_bytes}")
            self.message_queue.put(('log', f"Found {total_bytes} bytes to copy."))
            # Free space pre-check
            try:
                free = shutil.disk_usage(destination_dir).free
            except Exception:
                free = self.get_free_space(destination_dir)
            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space. Required={total_bytes} Free={free}")
                self.message_queue.put(('show_error', f"Not enough free space.\nRequired: {self.format_bytes(total_bytes)}\nAvailable: {self.format_bytes(free)}"))
                return

            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            items_processed = 0
            for source in source_paths:
                base = os.path.basename(source.rstrip("\\/"))
                # If base is empty (e.g., when source is a drive root like C:\), make a safe folder name
                if not base:
                    drive, _ = os.path.splitdrive(source)
                    base = (drive.rstrip(':') or 'root') + "_root"
                dest_path = os.path.join(destination_dir, base)
                if os.path.isdir(source):
                    self.write_detailed_log(f"Copying directory: {source} -> {dest_path}")
                    if os.path.exists(dest_path):
                        try:
                            shutil.rmtree(dest_path)
                            self.write_detailed_log(f"  - Removed existing destination {dest_path}")
                        except Exception as e:
                            self.write_detailed_log(f"  - Failed to remove existing destination: {e}")
                    os.makedirs(dest_path, exist_ok=True)
                    for dirpath, dirnames, filenames in os.walk(source, topdown=True, onerror=lambda e: None):
                        dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                        rel = os.path.relpath(dirpath, source)
                        target_dir = os.path.join(dest_path, rel) if rel != "." else dest_path
                        os.makedirs(target_dir, exist_ok=True)
                        for f in filenames:
                            src_file = os.path.join(dirpath, f)
                            if self.is_hidden(src_file):
                                self.write_detailed_log(f"Skipping hidden file: {src_file}")
                                continue
                            dest_file = os.path.join(target_dir, f)
                            self.write_detailed_log(f"Copying file: {src_file} -> {dest_file}")
                            self.copy_file_with_progress(src_file, dest_file, progress_cb)
                            items_processed += 1
                            if items_processed % 50 == 0:
                                self.message_queue.put(('log', f"Copied {items_processed} files..."))
                elif os.path.isfile(source):
                    if self.is_hidden(source):
                        self.write_detailed_log(f"Skipping hidden file: {source}")
                        continue
                    self.write_detailed_log(f"Copying file: {source} -> {dest_path}")
                    self.copy_file_with_progress(source, dest_path, progress_cb)
                    items_processed += 1
                    if items_processed % 50 == 0:
                        self.message_queue.put(('log', f"Copied {items_processed} files..."))

            # final update
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log("Copying process successfully completed.")
            self.message_queue.put(('log', "Copy complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Copy complete!"))
            self.message_queue.put(('open_folder', destination_dir))

        except Exception as e:
            self.write_detailed_log(f"Copy error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during copy: {e}"))
        finally:
            self.close_log_file()
            try:
                self.message_queue.put(('update_progress', self.progress_bar['maximum']))
            except Exception:
                pass
            self.message_queue.put(('enable_buttons', None))

    # --- Driver Backup & Restore Functions ---
    def start_browser_profiles_backup_thread(self):
        """Starts backup of full browser profiles (Chrome/Edge/Firefox)."""
        dest_root = filedialog.askdirectory(title="Select Folder to Save Browser Profiles")
        if not dest_root:
            self.message_queue.put(('log', "Browser profiles backup cancelled (no destination)."))
            return

        # Prepare destination and log
        self.open_log_file(dest_root, "browser_profiles")
        self.message_queue.put(('log', "브라우저 프로필 백업: Chrome/Edge/Firefox 프로필 전체(확장/설정 포함)"))
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
            ts = datetime.now().strftime('%Y-%m-%d')
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
                self.message_queue.put(('update_progress', self.progress_bar['maximum']))
            except Exception:
                pass
            self.message_queue.put(('enable_buttons', None))

    def start_driver_backup_thread(self):
        """Starts the driver backup process in a separate thread."""
        folder_selected = filedialog.askdirectory(title="Select Driver Backup Folder")
        if not folder_selected:
            self.message_queue.put(('log', "Driver backup folder selection cancelled."))
            return

        # open log file in driver backup folder
        self.open_log_file(folder_selected, "driver_backup")

        self.message_queue.put(('log', "Driver backup process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver backup in progress..."))

        backup_thread = threading.Thread(target=self.run_driver_backup, args=(folder_selected,), daemon=True)
        backup_thread.start()

    def run_driver_backup(self, folder_selected):
        """Executes the actual driver backup logic."""
        try:
            cmd = f'pnputil.exe /export-driver * "{folder_selected}"'
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True, encoding='cp949')

            while process.poll() is None:
                line = process.stdout.readline()
                if line:
                    self.write_detailed_log(line.strip())
            self.write_detailed_log("Driver backup successfully completed.")
            self.message_queue.put(('log', "Driver backup completed."))
            self.message_queue.put(('open_folder', folder_selected))
        except Exception as e:
            self.write_detailed_log(f"Driver backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during driver backup: {e}"))
        finally:
            self.close_log_file()
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('update_status', "Driver backup complete!"))
            self.message_queue.put(('enable_buttons', None))

    def start_driver_restore_thread(self):
        """Starts the driver restore process in a separate thread."""
        folder_selected = filedialog.askdirectory(title="Select Driver Restore Folder")
        if not folder_selected:
            self.message_queue.put(('log', "Driver restore folder selection cancelled."))
            return

        # open log file in driver restore folder
        self.open_log_file(folder_selected, "driver_restore")

        self.message_queue.put(('log', "Driver restore process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver restore in progress..."))

        restore_thread = threading.Thread(target=self.run_driver_restore, args=(folder_selected,), daemon=True)
        restore_thread.start()

    def run_driver_restore(self, folder_selected):
        """Executes the actual driver restore logic."""
        try:
            cmd = f'pnputil.exe /add-driver "{folder_selected}\\*.inf" /subdirs /install'
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True, encoding='cp949')

            while process.poll() is None:
                line = process.stdout.readline()
                if line:
                    self.write_detailed_log(line.strip())

            self.write_detailed_log("Driver restore successfully completed.")
            self.message_queue.put(('log', "Driver restore completed."))
        except Exception as e:
            self.write_detailed_log(f"Driver restore error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during driver restore: {e}"))
        finally:
            self.close_log_file()
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('update_status', "Driver restore complete!"))
            self.message_queue.put(('enable_buttons', None))

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
                messagebox.showerror("Error", f"Unable to open Device Manager:\n{e}")
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
            except Exception:
                pass

    def save_settings(self):
        import json
        path = self.settings_file_path()
        try:
            data = {
                'hidden': self.hidden_mode_var.get(),
                'system': self.system_mode_var.get(),
                'retention_count': self.retention_count_var.get(),
                'log_retention_days': self.log_retention_days_var.get(),
                'filters': self.filters,
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f)
        except Exception:
            pass

class ScheduleBackupDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Schedule Backup")
        try:
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        self.configure(bg="#2D3250")
        self.geometry("520x420")
        self.transient(master)
        try:
            self.grab_set()
        except Exception:
            pass

        self.action = None
        self.result = None
        users = []
        try:
            users = list(master.user_combo['values']) if getattr(master, 'user_combo', None) else []
        except Exception:
            users = []
        if not users:
            try:
                users_path = os.path.join("C:", os.sep, "Users")
                users = [d for d in os.listdir(users_path) if os.path.isdir(os.path.join(users_path, d)) and not d.startswith('.')]
            except Exception:
                users = []

        # Defaults
        default_user = master.user_var.get() if getattr(master, 'user_var', None) and master.user_var.get() else (users[0] if users else "")
        default_time = "02:00"
        default_task_name = f"ezBAK_Backup_{default_user}" if default_user else "ezBAK_Backup"

        # Form
        frm = tk.Frame(self, bg="#2D3250")
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        row = 0
        def add_label(text):
            nonlocal row
            tk.Label(frm, text=text, bg="#2D3250", fg="white", font=("Arial", 10, "bold")).grid(row=row, column=0, sticky='w', pady=4)

        def add_entry(var, width=34):
            nonlocal row
            e = tk.Entry(frm, textvariable=var, width=width)
            e.grid(row=row, column=1, sticky='w', pady=4)
            row += 1
            return e

        # Task name
        add_label("Task Name")
        self.task_name_var = tk.StringVar(value=default_task_name)
        add_entry(self.task_name_var)


        # Destination
        add_label("Destination Folder")
        self.dest_var = tk.StringVar()
        dest_entry = tk.Entry(frm, textvariable=self.dest_var, width=28)
        dest_entry.grid(row=row, column=1, sticky='w', pady=4)
        browse_btn = tk.Button(frm, text="Browse...", command=self._browse_dest, bg="#7D98A1", fg="white", relief="flat")
        browse_btn.grid(row=row, column=2, sticky='w', padx=6)
        row += 1

        # Schedule type
        add_label("Schedule")
        self.schedule_var = tk.StringVar(value='Daily')
        self.schedule_combo = ttk.Combobox(frm, textvariable=self.schedule_var, values=['Daily', 'Weekly'], state='readonly', width=12)
        self.schedule_combo.grid(row=row, column=1, sticky='w', pady=4)
        self.schedule_combo.bind('<<ComboboxSelected>>', self._toggle_weekly)
        row += 1

        # Time
        add_label("Time (HH:MM)")
        self.time_var = tk.StringVar(value=default_time)
        add_entry(self.time_var)

        # Weekly days
        self.days_vars = {k: tk.BooleanVar(value=(k == 'MON')) for k in ['MON','TUE','WED','THU','FRI','SAT','SUN']}
        self.days_frame = tk.Frame(frm, bg="#2D3250")
        tk.Label(self.days_frame, text="Days:", bg="#2D3250", fg="white").pack(side='left', padx=(0,6))
        for k in ['MON','TUE','WED','THU','FRI','SAT','SUN']:
            tk.Checkbutton(self.days_frame, text=k, variable=self.days_vars[k], bg="#2D3250", fg="white", selectcolor="#2D3250").pack(side='left')
        self.days_frame.grid(row=row, column=1, sticky='w', pady=4)
        row += 1

        # Options
        self.inc_hidden_var = tk.BooleanVar(value=False)
        self.inc_system_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frm, text="Include Hidden", variable=self.inc_hidden_var, bg="#2D3250", fg="white", selectcolor="#2D3250").grid(row=row, column=1, sticky='w', pady=4)
        row += 1
        tk.Checkbutton(frm, text="Include System", variable=self.inc_system_var, bg="#2D3250", fg="white", selectcolor="#2D3250").grid(row=row, column=1, sticky='w', pady=2)
        row += 1

        # Actions
        btns = tk.Frame(self, bg="#2D3250")
        btns.pack(fill='x', padx=12, pady=(6,12))
        tk.Button(btns, text="Create", command=self._create, bg="#336BFF", fg="white", relief="flat", width=12).pack(side='right')
        tk.Button(btns, text="Delete", command=self._delete, bg="#FF5733", fg="white", relief="flat", width=12).pack(side='right', padx=(0,8))
        tk.Button(btns, text="Close", command=self._close, bg="#7D98A1", fg="white", relief="flat", width=10).pack(side='right', padx=(0,8))

        self._toggle_weekly()

    def _browse_dest(self):
        path = filedialog.askdirectory(title="Select Destination Folder")
        if path:
            self.dest_var.set(path)

    def _toggle_weekly(self, event=None):
        if self.schedule_var.get().lower() == 'weekly':
            self.days_frame.grid()
        else:
            self.days_frame.grid_remove()

    def _create(self):
        user = self.master.user_var.get().strip()
        # user = "All Users"
        dest = self.dest_var.get().strip()
        name = self.task_name_var.get().strip()
        time_str = self.time_var.get().strip()
        sched = self.schedule_var.get().strip()
        days = [k for k,v in self.days_vars.items() if v.get()] if sched.lower() == 'weekly' else []
        if not (user and dest and name and time_str and sched):
            try:
                messagebox.showwarning("Warning", "Please complete all fields.")
            except Exception:
                pass
            return
        self.result = {
            'task_name': name,
            'schedule': sched,
            'time_str': time_str,
            'days': days,
            'user': user,
            'dest': dest,
            'include_hidden': self.inc_hidden_var.get(),
            'include_system': self.inc_system_var.get(),
        }
        self.action = 'create'
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _delete(self):
        self.task_name = self.task_name_var.get().strip()
        if not self.task_name:
            try:
                messagebox.showwarning("Warning", "Task Name is required to delete.")
            except Exception:
                pass
            return
        self.action = 'delete'
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _close(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


class FilterManagerDialog(tk.Toplevel):
    def __init__(self, master, current_filters=None):
        super().__init__(master)
        self.title("Filter Manager")
        try:
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        self.configure(bg="#2D3250")
        self.geometry("620x440")
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

        # Layout
        top = tk.Frame(self, bg="#2D3250")
        top.pack(fill='both', expand=True, padx=12, pady=12)

        # Include panel
        inc_frame = tk.LabelFrame(top, text="Include (files must match at least one)", bg="#2D3250", fg="white")
        inc_frame.pack(side='left', fill='both', expand=True, padx=(0,6))
        self.inc_list = tk.Listbox(inc_frame, height=14)
        self.inc_list.pack(fill='both', expand=True, padx=6, pady=6)
        btn_row_inc = tk.Frame(inc_frame, bg="#2D3250")
        btn_row_inc.pack(fill='x', padx=6, pady=(0,6))
        tk.Button(btn_row_inc, text="Add", width=8, command=lambda: self._add_rule('include'), bg="#336BFF", fg="white", relief="flat").pack(side='left')
        tk.Button(btn_row_inc, text="Remove", width=8, command=lambda: self._remove_selected('include'), bg="#FF5733", fg="white", relief="flat").pack(side='left', padx=(6,0))
        tk.Button(btn_row_inc, text="Clear", width=8, command=lambda: self._clear('include'), bg="#7D98A1", fg="white", relief="flat").pack(side='left', padx=(6,0))

        # Exclude panel
        exc_frame = tk.LabelFrame(top, text="Exclude (always skipped)", bg="#2D3250", fg="white")
        exc_frame.pack(side='left', fill='both', expand=True, padx=(6,0))
        self.exc_list = tk.Listbox(exc_frame, height=14)
        self.exc_list.pack(fill='both', expand=True, padx=6, pady=6)
        btn_row_exc = tk.Frame(exc_frame, bg="#2D3250")
        btn_row_exc.pack(fill='x', padx=6, pady=(0,6))
        tk.Button(btn_row_exc, text="Add", width=8, command=lambda: self._add_rule('exclude'), bg="#336BFF", fg="white", relief="flat").pack(side='left')
        tk.Button(btn_row_exc, text="Remove", width=8, command=lambda: self._remove_selected('exclude'), bg="#FF5733", fg="white", relief="flat").pack(side='left', padx=(6,0))
        tk.Button(btn_row_exc, text="Clear", width=8, command=lambda: self._clear('exclude'), bg="#7D98A1", fg="white", relief="flat").pack(side='left', padx=(6,0))

        # Bottom actions
        actions = tk.Frame(self, bg="#2D3250")
        actions.pack(fill='x', padx=12, pady=(0,12))
        tk.Button(actions, text="Save", width=10, command=self._save, bg="#2E7D32", fg="white", relief="flat").pack(side='right')
        tk.Button(actions, text="Cancel", width=10, command=self._cancel, bg="#7D98A1", fg="white", relief="flat").pack(side='right', padx=(0,8))

        self._refresh()

    def _rule_to_text(self, r):
        try:
            return f"[{r.get('type')}] {r.get('pattern')}"
        except Exception:
            return str(r)

    def _refresh(self):
        self.inc_list.delete(0, tk.END)
        for r in self.inc:
            self.inc_list.insert(tk.END, self._rule_to_text(r))
        self.exc_list.delete(0, tk.END)
        for r in self.exc:
            self.exc_list.insert(tk.END, self._rule_to_text(r))

    def _remove_selected(self, kind):
        if kind == 'include':
            idxs = list(self.inc_list.curselection())
            for i in reversed(idxs):
                del self.inc[i]
        else:
            idxs = list(self.exc_list.curselection())
            for i in reversed(idxs):
                del self.exc[i]
        self._refresh()

    def _clear(self, kind):
        if kind == 'include':
            self.inc = []
        else:
            self.exc = []
        self._refresh()

    def _add_rule(self, kind):
        dlg = tk.Toplevel(self)
        dlg.title("Add Rule")
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg="#2D3250")
        dlg.geometry("360x240")
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass

        tk.Label(dlg, text="Type:", bg="#2D3250", fg="white").pack(anchor='w', padx=10, pady=(10,2))
        vtype = tk.StringVar(value='ext')
        tframe = tk.Frame(dlg, bg="#2D3250")
        tframe.pack(anchor='w', padx=10)
        for t, txt in (('ext','Extension (e.g. tmp)'), ('name','Name (glob, e.g. Thumbs.db)'), ('path','Path (glob, e.g. */Cache/*)')):
            tk.Radiobutton(tframe, text=txt, variable=vtype, value=t, bg="#2D3250", fg="white", selectcolor="#2D3250").pack(anchor='w')

        tk.Label(dlg, text="Pattern:", bg="#2D3250", fg="white").pack(anchor='w', padx=10, pady=(6,2))
        vpat = tk.StringVar()
        ent = tk.Entry(dlg, textvariable=vpat, width=40)
        ent.pack(anchor='w', padx=10)

        btns = tk.Frame(dlg, bg="#2D3250")
        btns.pack(fill='x', padx=10, pady=10)
        def _ok():
            p = vpat.get().strip()
            if not p:
                dlg.destroy()
                return
            rule = {'type': vtype.get().strip().lower(), 'pattern': p}
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
        tk.Button(btns, text="Add", width=8, command=_ok, bg="#336BFF", fg="white", relief="flat").pack(side='right')
        tk.Button(btns, text="Cancel", width=8, command=lambda: (dlg.destroy()), bg="#7D98A1", fg="white", relief="flat").pack(side='right', padx=(0,6))

        try:
            ent.focus_set()
        except Exception:
            pass

    def _save(self):
        self.result = {'include': self.inc, 'exclude': self.exc}
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _cancel(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


class SelectSourcesDialog(tk.Toplevel):
    def __init__(self, master, is_hidden_fn=None, title="Select Sources"):
        super().__init__(master)
        self.title(title)
        try:
            # Try to reuse app icon if available
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        self.configure(bg="#2D3250")
        self.geometry("640x480")
        self.transient(master)
        try:
            self.grab_set()
        except Exception:
            pass

        self.is_hidden_fn = is_hidden_fn or (lambda p: False)
        self.selected_paths = []

        # Instruction
        tk.Label(self, text="Select folders and files to copy (checkbox tree).",
                 bg="#2D3250", fg="white", font=("Arial", 11, "bold")).pack(anchor='w', padx=10, pady=(10, 5))

        # Top controls
        top_controls = tk.Frame(self, bg="#2D3250")
        top_controls.pack(fill="x", padx=10, pady=(0, 5))

        self.add_root_btn = tk.Button(top_controls, text="Add Root Folder...", command=self._choose_root,
                                      bg="#7D98A1", fg="white", relief="flat")
        self.add_root_btn.pack(side="left")

        self.sel_all_btn = tk.Button(top_controls, text="Select All", command=self._select_all,
                                     bg="#4CAF50", fg="white", relief="flat")
        self.sel_all_btn.pack(side="left", padx=(8, 0))

        self.clear_all_btn = tk.Button(top_controls, text="Clear All", command=self._clear_all,
                                       bg="#FF5733", fg="white", relief="flat")
        self.clear_all_btn.pack(side="left", padx=(8, 0))

        # Tree area
        tree_frame = tk.Frame(self, bg="#2D3250")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(tree_frame, columns=("fullpath",), displaycolumns=())
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<<TreeviewOpen>>", self._on_open)
        self.tree.bind("<Button-1>", self._on_click)

        self._checked = set()       # item ids that are checked
        self._item_path = {}        # item id -> full filesystem path

        # Populate available drives as roots by default
        self._add_drive_roots()

        # Action buttons
        action_frame = tk.Frame(self, bg="#2D3250")
        action_frame.pack(fill="x", padx=10, pady=(5, 10))

        ok_btn = tk.Button(action_frame, text="OK", width=10, command=self._on_ok,
                           bg="#336BFF", fg="white", relief="flat")
        ok_btn.pack(side="right")
        cancel_btn = tk.Button(action_frame, text="Cancel", width=10, command=self._on_cancel,
                               bg="#FF3333", fg="white", relief="flat")
        cancel_btn.pack(side="right", padx=(0, 8))

    def _format(self, name, checked):
        return f"[{'x' if checked else ' '}] {name}"

    def _add_drive_roots(self):
        """
        Uses the 'mountvol' command to find all mounted drives with drive letters.
        This is more reliable than iterating through 'C:' to 'Z:'.
        """
        drives = []
        try:
            # For compatibility with different Windows language versions,
            # we run with a standard locale and decode as UTF-8 with error handling.
            env = os.environ.copy()
            env['LC_ALL'] = 'C'
            # Use check_output to get the command's output
            output = subprocess.check_output("mountvol", text=True, encoding='utf-8', errors='ignore', shell=True, env=env)
            
            # Regex to find lines that end with a drive letter (e.g., "    C:\")
            drive_pattern = re.compile(r"([A-Z]:\\)$")
            
            for line in output.splitlines():
                match = drive_pattern.search(line.strip())
                if match:
                    drives.append(match.group(1))
                    
        except (subprocess.CalledProcessError, FileNotFoundError):
            # Fallback to the old method if mountvol fails
            for code in range(ord('C'), ord('Z') + 1):
                drive = f"{chr(code)}:\\"
                if os.path.exists(drive):
                    drives.append(drive)
        
        # Add unique, sorted drives to the tree
        for drive in sorted(list(set(drives))):
            try:
                item = self.tree.insert("", "end", text=self._format(drive, False), open=False)
                self._item_path[item] = drive
                # dummy child for lazy load
                self.tree.insert(item, "end", text="...")
            except Exception:
                pass

    def _choose_root(self):
        path = filedialog.askdirectory(title="Select Root Folder")
        if path:
            try:
                item = self.tree.insert("", "end", text=self._format(path, False), open=False)
                self._item_path[item] = path
                self.tree.insert(item, "end", text="...")
            except Exception:
                pass

    def _on_open(self, event):
        item = self.tree.focus()
        if not item:
            return
        children = self.tree.get_children(item)
        if len(children) == 1 and self.tree.item(children[0], "text") == "...":
            # populate children lazily
            self.tree.delete(children[0])
            path = self._item_path.get(item)
            # Check if the parent item is checked to propagate the state to new children.
            parent_is_checked = self.tree.item(item, "text").startswith("[x]")

            if not path:
                return
            try:
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
                    # Inherit the checked state from the parent when creating the child item.
                    child = self.tree.insert(item, "end", text=self._format(name, parent_is_checked), open=False)
                    self._item_path[child] = full
                    # If the parent is checked, the new child must also be added to the checked set.
                    if parent_is_checked:
                        self._checked.add(child)
                    if is_dir:
                        self.tree.insert(child, "end", text="...")
            except Exception:
                pass

    def _toggle_item(self, item):
        text = self.tree.item(item, "text")
        if not text:
            return
        checked = text.startswith("[x]")
        new_checked = not checked
        name = text[4:] if len(text) >= 4 else text
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

    def _apply_to_descendants(self, item, checked):
        text = self.tree.item(item, "text")
        name = text[4:] if len(text) >= 4 else text
        self.tree.item(item, text=self._format(name, checked))
        if checked:
            self._checked.add(item)
        else:
            self._checked.discard(item)
        for child in self.tree.get_children(item):
            if self.tree.item(child, "text") == "...":
                continue
            self._apply_to_descendants(child, checked)

    def _on_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        try:
            elem = self.tree.identify("element", event.x, event.y)
        except Exception:
            elem = ""
        elem_str = str(elem).lower()
        # Only toggle when clicking on the text area; ignore indicator/expander/icon clicks
        if "text" not in elem_str:
            return
        self._toggle_item(item)

    def _select_all(self):
        for item in self.tree.get_children(""):
            self._set_checked_recursive(item, True)

    def _clear_all(self):
        for item in self.tree.get_children(""):
            self._set_checked_recursive(item, False)

    def _set_checked_recursive(self, item, checked):
        text = self.tree.item(item, "text")
        name = text[4:] if len(text) >= 4 else text
        self.tree.item(item, text=self._format(name, checked))
        if checked:
            self._checked.add(item)
        else:
            self._checked.discard(item)
        for child in self.tree.get_children(item):
            if self.tree.item(child, "text") == "...":
                continue
            self._set_checked_recursive(child, checked)

    def _on_ok(self):
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
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _on_cancel(self):
        self.selected_paths = []
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

def core_run_backup(user_name, dest_dir, include_hidden=False, include_system=False, log_folder=None):
    """
    GUI와 스케줄러 공통 백업 로직 (상세 로그 기록 추가)
    """
    import os, shutil
    from datetime import datetime

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_name = f"{user_name}_backup_{timestamp}"
    backup_path = os.path.join(dest_dir, backup_name)
    log_path = os.path.join(dest_dir, f"{user_name}_backup_{timestamp}.log")

    os.makedirs(backup_path, exist_ok=True)

    total_bytes = 0
    files_copied = 0
    user_home = os.path.join("C:\\Users", user_name)

    if not os.path.exists(user_home):
        raise FileNotFoundError(f"User folder '{user_home}' not found.")

    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"[{timestamp}] Backup start for '{user_name}' → {backup_path}\n")

        for root, dirs, files in os.walk(user_home):
            # 숨김/시스템 제외 처리
            if not include_hidden:
                dirs[:] = [d for d in dirs if not d.startswith('.')]
            if not include_system:
                dirs[:] = [d for d in dirs if d.lower() not in ("system32", "appdata")]

            rel_dir = os.path.relpath(root, user_home)
            dest_dir = os.path.join(backup_path, rel_dir) if rel_dir != '.' else backup_path
            os.makedirs(dest_dir, exist_ok=True)
            log.write(f"[{timestamp}] Created folder: {dest_dir}\n")

            for f in files:
                src = os.path.join(root, f)
                dst = os.path.join(dest_dir, f)
                try:
                    shutil.copy2(src, dst)
                    files_copied += 1
                    size = os.path.getsize(src)
                    total_bytes += size
                    log.write(f"[{timestamp}] Copied: {src} -> {dst} ({size} bytes)\n")
                except Exception as e:
                    log.write(f"[{timestamp}] [WARN] Skipped: {src} ({e})\n")

        log.write(f"[{timestamp}] Backup completed. Files={files_copied}, Size={total_bytes} bytes\n")

    # 오래된 백업/로그 정리
    if log_folder:
        try:
            _cleanup_old_backups(dest_dir, user_name)
            _cleanup_old_logs(log_folder, retention_days=7)
        except Exception as e:
            print(f"[WARN] cleanup skipped: {e}")

    return backup_path, log_path



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
        import argparse
        import sys
        from datetime import datetime

        parser = argparse.ArgumentParser()
        parser.add_argument("--user", help="User name to back up")
        parser.add_argument("--dest", help="Destination folder for backup")
        parser.add_argument("--include-hidden", action="store_true")
        parser.add_argument("--include-system", action="store_true")
        parser.add_argument("--log-folder", default="logs")
        args, _ = parser.parse_known_args()

        if args.user and args.dest:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] Headless backup start for user '{args.user}'")

            try:
                backup_path, log_path = core_run_backup(
                    args.user,
                    args.dest,
                    include_hidden=args.include_hidden,
                    include_system=args.include_system,
                    log_folder=args.log_folder
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