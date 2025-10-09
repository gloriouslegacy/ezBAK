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

# wim backup core functions
# wim once boot functions

        #     # 항상 실행되는 정리 작업
        # finally:
        #     try:
        #         self.close_log_file()
        #     except Exception:
        #         pass
            
        #     try:
        #         # 진행률을 100%로 설정
        #         max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
        #         self.message_queue.put(('update_progress', max_val))
        #     except Exception:
        #         pass
            
        #     # 버튼 활성화
        #     self.message_queue.put(('enable_buttons', None))
            
        #     # 최종 상태 메시지
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
    WMI를 통해 베이스보드 모델명 또는 제품명 가져오기
    우선순위: Baseboard Product → Computer System Model  Computer System Manufacturer → Fallback "Unknown"
    """
    def clean_name(name):
        """파일명에 사용할 수 없는 문자를 제거하고 정리"""
        if not name or name.lower() in ('to be filled by o.e.m.', 'system product name', 'system version'):
            return None
        # 특수문자 제거 및 공백을 언더스코어로 변경
        cleaned = re.sub(r'[<>:"/\\|?*]', '', name)
        cleaned = re.sub(r'\s+', '_', cleaned.strip())
        return cleaned if cleaned else None

    system_name = "Unknown"
    
    try:
        # 1. 베이스보드 제품명 
        cmd = ['wmic', 'baseboard', 'get', 'product', '/value']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
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
        # 2. 컴퓨터 시스템 모델명 
        cmd = ['wmic', 'computersystem', 'get', 'model', '/value']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
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
        # 3. 제조사명 (최후의 수단)
        cmd = ['wmic', 'computersystem', 'get', 'manufacturer', '/value']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
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
    오류 처리기 for shutil.rmtree (headless 모드용 - 로그에 기록)
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
        """헤드리스 모드 백업 실행"""
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
                # 진행률을 100%로 설정
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass
            
            # 버튼 활성화
            self.message_queue.put(('enable_buttons', None))
            
            # 최종 상태 메시지
            if hasattr(self, '_backup_completed') and self._backup_completed:
                self.message_queue.put(('update_status', "Backup complete!"))
            else:
                self.message_queue.put(('update_status', "Operation finished"))

class DialogShortcuts:
    """다이얼로그용 공통 단축키 기능"""
    
    def setup_dialog_shortcuts(self):
        """다이얼로그에 기본 단축키 설정"""
        # 기본 다이얼로그 단축키들
        self.bind('<Escape>', lambda e: self.on_cancel() if hasattr(self, 'on_cancel') else self.destroy())
        self.bind('<Return>', lambda e: self.on_ok() if hasattr(self, 'on_ok') else None)
        self.bind('<Alt-F4>', lambda e: self.destroy())
        
        # 포커스 설정 (키보드 이벤트를 받기 위해)
        self.focus_set()
        
    def add_shortcut_text(self, text, shortcut):
        """텍스트에 단축키 표기 추가"""
        return f"{text} ({shortcut})"

class KeyboardShortcuts:
    """ezBAK용 키보드 단축키 관리 클래스"""
    
    def __init__(self, app):
        self.app = app
        self.setup_shortcuts()
        self.setup_help_dialog()
    
    def setup_shortcuts(self):
        """키보드 단축키 등록"""
        
        # 메인 기능 단축키
        self.app.bind('<Control-b>', lambda e: self.safe_execute(self.app.start_backup_thread))
        self.app.bind('<Control-r>', lambda e: self.safe_execute(self.app.start_restore_thread))
        self.app.bind('<Control-d>', lambda e: self.safe_execute(self.app.start_driver_backup_thread))
        self.app.bind('<Control-Shift-D>', lambda e: self.safe_execute(self.app.start_driver_restore_thread))
        
        # 파일 관리
        self.app.bind('<Control-c>', lambda e: self.safe_execute(self.app.copy_files))
        self.app.bind('<Control-o>', lambda e: self.safe_execute(self.app.open_file_explorer))
        self.app.bind('<Control-s>', lambda e: self.safe_execute(self.app.save_log))
        
        # 도구 및 유틸리티
        self.app.bind('<Control-m>', lambda e: self.safe_execute(self.app.open_device_manager))
        self.app.bind('<Control-k>', lambda e: self.safe_execute(self.app.check_space))
        self.app.bind('<Control-f>', lambda e: self.safe_execute(self.app.open_filter_manager))
        self.app.bind('<Control-t>', lambda e: self.safe_execute(self.app.schedule_backup))
        
        # 브라우저 및 앱 관련
        self.app.bind('<Control-p>', lambda e: self.safe_execute(self.app.start_browser_profiles_backup_thread))
        self.app.bind('<Control-w>', lambda e: self.safe_execute(self.app.start_winget_export_thread))
        
        # NAS 연결
        self.app.bind('<Control-n>', lambda e: self.safe_execute(self.app.open_nas_connect_dialog))
        self.app.bind('<Control-Shift-N>', lambda e: self.safe_execute(self.app.open_nas_disconnect_dialog))
        
        # 일반 단축키
        self.app.bind('<F1>', lambda e: self.show_help())
        self.app.bind('<F5>', lambda e: self.refresh_user_list())
        self.app.bind('<Escape>', lambda e: self.cancel_operation())
        self.app.bind('<Control-q>', lambda e: self.safe_exit())
        
        # 사용자 선택 단축키 (숫자키)
        for i in range(1, 10):
            self.app.bind(f'<Control-{i}>', lambda e, idx=i-1: self.select_user_by_index(idx))
        
        # Focus 단축키
        self.app.focus_set()  # 윈도우가 키보드 이벤트를 받을 수 있도록
    
    def safe_execute(self, function):
        """함수 실행 (에러 핸들링 포함)"""
        try:
            if callable(function):
                function()
            else:
                self.app.message_queue.put(('log', f"Function {function} is not callable"))
        except Exception as e:
            self.app.message_queue.put(('log', f"Shortcut execution error: {e}"))
            print(f"DEBUG: Shortcut error: {e}")
    
    def show_help(self):
        """단축키 도움말 표시 (F1)"""
        help_text = """
ezBAK Keyboard Shortcuts

┌─────────────────────────────────────────┐
│           Main Features                 │
├─────────────────────────────────────────┤
│ Ctrl + B        Backup User Data        │
│ Ctrl + R        Restore User Data       │
│ Ctrl + D        Backup Drivers          │
│ Ctrl + Shift + D Restore Drivers        │
├─────────────────────────────────────────┤
│           File management               │
├─────────────────────────────────────────┤
│ Ctrl + C        Copy Data               │
│ Ctrl + O        File Explorer           │
│ Ctrl + S        Save Log                │
├─────────────────────────────────────────┤
│           Tools and Utilities           │
├─────────────────────────────────────────┤
│ Ctrl + M        Device Manager          │
│ Ctrl + K        Check Space             │
│ Ctrl + F        Filters                 │
│ Ctrl + T        Schedule Backup         │
├─────────────────────────────────────────┤
│           Browser and App               │
├─────────────────────────────────────────┤
│ Ctrl + P        Backup Browser Profiles │
│ Ctrl + W        Export Apps             │
├─────────────────────────────────────────┤
│           Network                       │
├─────────────────────────────────────────┤
│ Ctrl + N        Connect NAS             │
│ Ctrl + Shift + N Disconnect NAS         │
├─────────────────────────────────────────┤
│           General                       │
├─────────────────────────────────────────┤
│ F1              help                    │
│ F5              (Not Working) Refresh User List       │
│ Ctrl + 1~9      (Not Working) Quick Select User │
│ Escape          (Not Working) Cancel Current Operation│
│ Ctrl + Q        Exit Program            │
└─────────────────────────────────────────┘

Tips :  Shortcuts are not case-sensitive.
"""
        
        # 사용자 정의 도움말 다이얼로그
        help_dialog = tk.Toplevel(self.app)
        help_dialog.title("Help - Keyboard Shortcuts")
        help_dialog.configure(bg="#2D3250")
        help_dialog.geometry("449x748")
        help_dialog.transient(self.app)
        help_dialog.grab_set()
        
        try:
            help_dialog.iconbitmap(resource_path('./icon/ezbak.ico'))
        except:
            pass
        
        # 텍스트 영역
        text_frame = tk.Frame(help_dialog, bg="#2D3250")
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 스크롤바를 먼저 생성
        scrollbar = tk.Scrollbar(text_frame, 
                                width=20,  
                                bg="#424769",
                                troughcolor="#2D3250",
                                activebackground="#6EC571",
                                relief="flat")
        scrollbar.pack(side="right", fill="y")
        
        text_widget = tk.Text(text_frame, 
                             font=("Consolas", 10),  
                             bg="#424769", 
                             fg="white",
                             relief="flat",
                             wrap="word",
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)
        
        # 스크롤바와 텍스트 위젯 연결
        scrollbar.configure(command=text_widget.yview)
        
        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")
        
        # 닫기 버튼
        close_btn = tk.Button(help_dialog, 
                             text="Exit (Esc)", 
                             command=help_dialog.destroy,
                             bg="#4CAF50", 
                             fg="white", 
                             relief="flat")
        close_btn.pack(pady=10)
        
        # ESC로 닫기
        help_dialog.bind('<Escape>', lambda e: help_dialog.destroy())
        help_dialog.focus_set()
    
    def refresh_user_list(self):
        """사용자 목록 새로고침 (F5)"""
        try:
            self.app.load_users()
            self.app.message_queue.put(('log', "Refreshed user list."))
        except Exception as e:
            self.app.message_queue.put(('log', f"Error refreshing user list: {e}"))

    def select_user_by_index(self, index):
        """숫자키로 사용자 빠른 선택 (Ctrl + 1~9)"""
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
        """현재 작업 취소 (ESC)"""
        try:
            # 진행 중인 작업이 있으면 취소 시도
            if hasattr(self.app, 'progress_bar') and self.app.progress_bar['mode'] == 'indeterminate':
                self.app.message_queue.put(('log', "Cancellation requested"))
                # 실제 취소 로직은 각 작업 스레드에서 구현 필요
        except Exception as e:
            print(f"DEBUG: Cancel operation error: {e}")
    
    def safe_exit(self):
        """프로그램 종료 (Ctrl + Q)"""
        try:
            if messagebox.askyesno("Confirm Exit", "Are you sure you want to exit ezBAK?"):
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
        """단축키 도움말을 상태바나 메뉴에 표시"""
        # 상태바에 단축키 힌트 추가 (선택사항)
        try:
            if hasattr(self.app, 'status_label'):
                original_text = self.app.status_label.cget("text")
                hint_text = f"{original_text} | F1: help"
                self.app.status_label.configure(text=hint_text)
        except:
            pass

# App 클래스에 추가할 메서드들
def setup_keyboard_shortcuts(self):
    self.shortcuts = KeyboardShortcuts(self)
    
    # 툴팁 표시 (선택사항)
    self.create_tooltip(self.backup_btn, "Backup User Data (Ctrl+B)")
    self.create_tooltip(self.restore_btn, "Restore User Data (Ctrl+R)")
    self.create_tooltip(self.driver_backup_btn, "Backup Drivers (Ctrl+D)")
    # ... 다른 버튼들도 동일하게

def create_tooltip(self, widget, text):
    """버튼에 툴팁 추가 (단축키 정보 포함)"""
    def on_enter(event):
        try:
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.configure(bg="#2D3250", relief="solid", borderwidth=1)
            
            label = tk.Label(tooltip, text=text, 
                           bg="#2D3250", fg="white", 
                           font=("Arial", 9))
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
        self.geometry("1024x429")
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
        self.setup_keyboard_shortcuts()

    def setup_keyboard_shortcuts(self):
        """키보드 단축키 설정"""
        self.shortcuts = KeyboardShortcuts(self)

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
        log_retention_label = tk.Label(opt_frame, text="Logs to Keep:", bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
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
        button_width = 15
        button_padx = 5
        button_pady = 3

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

        # self.test_btn = tk.Button(button_grid_frame, text="Test Log Cleanup", bg="#8E44AD", fg="white", 
        #                                     font=("Arial", button_font_size, "bold"), padx=button_padx, pady=button_pady,
        #                                     relief="flat", width=button_width, command=self.test_log_cleanup)
        # self.test_btn.grid(row=3, column=0, padx=2, pady=2, sticky="ew")

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
            

    def _cleanup_old_logs(self, log_folder, retention_days):
        """오래된 로그 파일을 정리합니다."""
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
            
            # 변수를 try 블록 밖에서 초기화
            deleted_count = 0
            log_files = []
        
            # 디버그: 폴더의 모든 파일 확인
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
                    messagebox.showerror("Error", value)
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
                self.message_queue.task_done()
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue) 


    def show_space_check_result(self, result):
        """용량 체크 결과를 사용자에게 표시"""
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
            messagebox.showinfo("Space Check", f"✓ Sufficient space available.\n\n{msg}")
            self.message_queue.put(('log', f"Space OK. Required={self.format_bytes(required)} Available={self.format_bytes(available)}"))
            self.message_queue.put(('update_status', "Space check completed - OK"))
        else:
            messagebox.showwarning("Space Check", f"⚠ Insufficient space available.\n\n{msg}")
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
        """진행률 바 초기화"""
        try:
            self.progress_bar.stop()  # indeterminate 모드 정지
            self.progress_bar['mode'] = "determinate"
            self.progress_bar['value'] = 0
            self.progress_bar['maximum'] = 100
            self.status_label.config(text="Ready")
        except Exception:
            pass

    def set_progress_indeterminate(self):
        """진행률 바를 indeterminate 모드로 설정"""
        try:
            self.progress_bar['mode'] = "indeterminate"
            self.progress_bar.start()
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
                        # 단순 포함 검사
                        if patt_lower in filepath_lower:
                            return True
                        # 정규화된 경로 매칭
                        path_normalized = filepath_lower.replace('\\', '/')
                        patt_normalized = patt_lower.replace('\\', '/')
                        if patt_normalized in path_normalized:
                            return True
                        # 와일드카드 매칭
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
        빠른 용량 계산 - 중간 결과를 UI에 업데이트하면서 계산
        """
        total = 0
        file_count = 0
        last_update_time = time.time()
        
        try:
            for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=lambda e: None):
                # 숨김 디렉토리 필터링
                dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if self.is_hidden(fp):
                        continue
                        
                    try:
                        size = os.path.getsize(fp)
                        total += size
                        file_count += 1
                        
                        # 주기적으로 UI 업데이트 (응답성 향상)
                        if file_count % max_files_per_update == 0:
                            current_time = time.time()
                            if current_time - last_update_time >= 0.1:  # 0.1초마다 업데이트
                                self.message_queue.put(('update_status', f"계산 중... {file_count}개 파일, {self.format_bytes(total)}"))
                                self.update_idletasks()  # UI 업데이트
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
        """Executes the actual backup logic with improved error handling and debugging."""
        bytes_copied = 0
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
            self.write_detailed_log(f"Total bytes to copy: {total_bytes}")
            
            # 여유 공간 검사
            self.message_queue.put(('update_status', "Checking available space..."))
            free = self.get_free_space_reliable(backup_path)
            
            if total_bytes > free:
                self.write_detailed_log(f"Insufficient free space for backup. Required={total_bytes} Free={free}")
                self.message_queue.put(('show_error', f"목적지에 공간이 부족합니다.\n필요: {self.format_bytes(total_bytes)}\n사용가능: {self.format_bytes(free)}"))
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
                        # 현재 처리 중인 폴더 로깅
                        rel_path = os.path.relpath(dirpath, user_profile_path)
                        self.write_detailed_log(f"Processing directory: {dirpath} (relative: {rel_path})")
                        
                        # 숨김 디렉토리 필터링 - 안전한 방식
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
                                # 에러 발생 시 디렉토리 포함 (안전한 선택)
                                dirnames.append(d)
                        
                        # 목적지 폴더 생성
                        rel_dir = os.path.relpath(dirpath, user_profile_path)
                        dest_dir = os.path.join(destination_dir, rel_dir) if rel_dir != "." else destination_dir
                        
                        try:
                            os.makedirs(dest_dir, exist_ok=True)
                            self.write_detailed_log(f"Created folder: {dest_dir}")
                            folders_processed += 1
                            
                            # 진행상황 업데이트
                            if folders_processed % 5 == 0:
                                self.message_queue.put(('update_status', f"Processing folder {folders_processed}: {os.path.basename(dirpath)}"))
                        except Exception as e:
                            self.write_detailed_log(f"ERROR: Failed to create folder {dest_dir}: {e}")
                            continue

                        # 파일 복사
                        for f in filenames:
                            try:
                                src_file = os.path.join(dirpath, f)
                                dest_file = os.path.join(dest_dir, f)
                                
                                # 숨김 파일 체크 
                                if self.is_hidden_safe(src_file):
                                    self.write_detailed_log(f"Skipping hidden file: {src_file}")
                                    continue
                                
                                # 파일 복사
                                self.write_detailed_log(f"Copying file: {src_file} -> {dest_file}")
                                copied = self.copy_file_with_progress_safe(src_file, dest_file, progress_cb)
                                files_processed += 1
                                
                                # 진행상황 업데이트
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
                self.message_queue.put(('show_error', f"백업 중 심각한 오류 발생: {e}"))
                return

            # 최종 진행률 업데이트
            self.message_queue.put(('update_progress', bytes_copied))
            self.write_detailed_log(f"Backup completed. Files backed up: {files_processed}, Folders: {folders_processed}, Total size: {self.format_bytes(bytes_copied)}")
            
            # 브라우저 북마크 백업
            try:
                self.backup_browser_bookmarks(backup_path)
            except Exception as e:
                self.write_detailed_log(f"Browser bookmark backup failed: {e}")
            
            # 정리 작업
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

        except Exception as e:
            self.write_detailed_log(f"Backup error: {e}")
            import traceback
            self.write_detailed_log(f"Backup traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"An error occurred during backup: {e}"))
        finally:  # 전체 함수의 finally 블록
            # 항상 실행되는 정리 작업
            try:
                self.close_log_file()
            except Exception:
                pass
            
            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass
            
            # 버튼 활성화
            self.message_queue.put(('enable_buttons', None))
            
            # 추가 안전장치: 직접 버튼 활성화도 시도
            try:
                self.after(1000, lambda: self.set_buttons_state("normal"))
            except Exception:
                pass

            if hasattr(self, '_backup_completed') and self._backup_completed:
                self.message_queue.put(('update_status', "Backup complete!"))
            else:
                self.message_queue.put(('update_status', "Operation finished"))          
            
            
    def copy_file_with_progress_safe(self, src, dst, progress_callback, buffer_size=64*1024, timeout_seconds=30):
        """
        타임아웃이 있는 안전한 파일 복사
        """
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            copied = 0
            start_time = time.time()
            
            with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
                while True:
                    # 타임아웃 체크
                    if time.time() - start_time > timeout_seconds:
                        self.write_detailed_log(f"TIMEOUT: File copy timeout for {src}")
                        raise TimeoutError(f"File copy timeout: {src}")
                    
                    buf = fsrc.read(buffer_size)
                    if not buf:
                        break
                    fdst.write(buf)
                    copied += len(buf)
                    progress_callback(len(buf))
            
            # 파일 속성 복사 (타임아웃 적용)
            try:
                shutil.copystat(src, dst)
            except Exception:
                pass  # 속성 복사 실패는 무시
                
            return copied
        except Exception as e:
            self.write_detailed_log(f"ERROR: copying file {src} -> {dst}: {e}")
            return 0

    def compute_total_bytes_safe(self, root_path):
        """
        안전한 총 용량 계산
        """
        total = 0
        file_count = 0
        
        try:
            for dirpath, dirnames, filenames in os.walk(root_path, topdown=True, onerror=lambda e: self.write_detailed_log(f"Walk error: {e}")):
                # 안전한 디렉토리 필터링
                original_dirnames = dirnames.copy()
                dirnames.clear()
                
                for d in original_dirnames:
                    try:
                        if not self.is_hidden_safe(os.path.join(dirpath, d), timeout_seconds=1):
                            dirnames.append(d)
                    except Exception:
                        dirnames.append(d)  # 에러 시 포함
                
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    try:
                        if not self.is_hidden_safe(fp, timeout_seconds=1):
                            size = os.path.getsize(fp)
                            total += size
                            file_count += 1
                            
                            # 100개마다 진행상황 업데이트
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
        타임아웃이 있는 is_hidden 함수 - 필터 포함
        """
        try:
            # 존재 확인
            if not os.path.exists(filepath):
                return False
            
            is_dir = os.path.isdir(filepath)
            
            # 필터 체크 (OneDrive 등)
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
            
            # Exclude 필터 확인
            for rule in filters.get('exclude', []) or []:
                if _matches(rule):
                    return True
            
            # Include 필터 확인 (파일만)
            inc_rules = filters.get('include', []) or []
            if inc_rules and not is_dir:
                has_match = False
                for rule in inc_rules:
                    if _matches(rule):
                        has_match = True
                        break
                if not has_match:
                    return True
            
            # 기본 숨김 파일 체크
            base = os.path.basename(filepath).lower()
            if base.startswith('.') or base in ('thumbs.db', 'desktop.ini', '$recycle.bin'):
                return True
            
            # Windows 속성 체크
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

        except Exception as e:
            self.write_detailed_log(f"Restore error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during restore: {e}"))
        finally:
            # 항상 실행되는 정리 작업
            try:
                self.close_log_file()
            except Exception:
                pass
            
            try:
                # 진행률을 100%로 설정
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass
            
            # 버튼 활성화
            self.message_queue.put(('enable_buttons', None))
            
            # 최종 상태 메시지
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
                dlg = SelectSourcesDialog(self, is_hidden_fn=self.is_hidden)
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
            destination_dir = filedialog.askdirectory(title="Select Destination Folder")
            print(f"DEBUG: Destination selected: {destination_dir}")
            
            if not destination_dir:
                print("DEBUG: No destination selected")
                self.message_queue.put(('log', "Destination selection cancelled."))
                return

            # Copy Data용 로그명: 'copy_년월일_시분초'
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
        더 안정적인 여유 공간 계산
        """
        try:
            # 우선 shutil.disk_usage 시도
            usage = shutil.disk_usage(path)
            return usage.free
        except Exception:
            try:
                # Windows API를 통한 여유 공간 계산
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
                    # 드라이브 루트로 재시도
                    drive = os.path.splitdrive(path)[0]
                    if drive:
                        root_path = drive + os.sep
                        usage = shutil.disk_usage(root_path)
                        return usage.free
                except Exception:
                    pass
        
        # 모든 방법이 실패하면 충분히 큰 값 반환 (계속 진행)
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

            # dlg가 여전히 존재하고 유효한지 확인
            if dlg and hasattr(dlg, 'action'):
                action = dlg.action
                print(f"DEBUG: Dialog action = {action}")
                
                if action == 'create':
                    if hasattr(dlg, 'result') and dlg.result:
                        data = dlg.result
                        print(f"DEBUG: Creating scheduled task with data: {data}")
                        
                        try:
                            self.create_scheduled_task(
                                task_name=data.get('task_name'),
                                dest=data.get('dest'),
                                schedule=data.get('schedule'),
                                time_str=data.get('time_str'),
                                include_hidden=data.get('include_hidden', False),
                                include_system=data.get('include_system', False),
                                retention_count=int(data.get('retention_count', 2)),
                                log_retention_days=int(data.get('log_retention_days', 30)),
                                user=data.get('user')
                            )
                            print("DEBUG: Scheduled task created successfully")
                            # 성공 메시지는 create_scheduled_task 내부에서 처리
                            
                        except Exception as task_error:
                            print(f"DEBUG: Task creation failed: {task_error}")
                            # 예외를 다시 던지지 말고 사용자에게 표시
                            error_msg = f"Failed to create scheduled task: {str(task_error)}"
                            self.message_queue.put(('log', error_msg))
                            try:
                                messagebox.showerror("Task Creation Failed", error_msg)
                            except Exception:
                                print(f"ERROR: {error_msg}")
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
                                messagebox.showerror("Task Deletion Failed", error_msg)
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
            
            # messagebox 호출도 안전하게 처리
            try:
                messagebox.showerror("Error", error_msg)
            except Exception as mb_error:
                print(f"DEBUG: messagebox.showerror failed: {mb_error}")

            
    ############# Schedule Backup Functions #############
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

    def create_scheduled_task(self, task_name, schedule, time_str, user, dest, include_hidden=False, include_system=False, retention_count=2, log_retention_days=30):
        """
        스케줄된 작업 생성 (관리자 권한으로 실행되도록 설정)
        """
        try:
            # 시간 형식 검증 및 정규화
            from datetime import datetime
            try:
                time_obj = datetime.strptime(time_str.strip(), "%H:%M")
                formatted_time = time_obj.strftime("%H:%M")
                self.write_detailed_log(f"Time format validated: {time_str} -> {formatted_time}")
            except ValueError as ve:
                raise ValueError(f"Invalid time format '{time_str}'. Please use HH:MM format (e.g., 14:30)")

            # 스케줄 타입 매핑 및 검증
            schedule_lower = schedule.lower().strip()
            if schedule_lower == "daily":
                sc_type = "DAILY"
            elif schedule_lower == "weekly":
                sc_type = "WEEKLY"
            elif schedule_lower == "monthly":
                sc_type = "MONTHLY"
            else:
                raise ValueError(f"Unsupported schedule type: {schedule}. Use Daily, Weekly, or Monthly")

            # 실행 파일 경로 확인
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
                script_path = ""
            else:
                exe_path = sys.executable
                script_path = f'"{os.path.abspath(__file__)}"'

            # 명령어 구성 (retention 옵션 포함)
            if script_path:
                base_cmd = f'"{exe_path}" {script_path} --user "{user}" --dest "{dest}" --retention {retention_count} --log-retention {log_retention_days}'
            else:
                base_cmd = f'"{exe_path}" --user "{user}" --dest "{dest}" --retention {retention_count} --log-retention {log_retention_days}'

            if include_hidden:
                base_cmd += " --include-hidden"
            if include_system:
                base_cmd += " --include-system"
                
            self.write_detailed_log(f"Task command: {base_cmd}")
            self.write_detailed_log(f"Creating scheduled task with retention_count={retention_count}, log_retention_days={log_retention_days}")

            # 현재 사용자 정보 가져오기
            current_user = os.environ.get('USERNAME', 'Administrator')
            domain = os.environ.get('USERDOMAIN', os.environ.get('COMPUTERNAME', ''))
            
            # 사용자 계정 형식 결정
            if domain and domain.upper() != current_user.upper():
                run_as_user = f"{domain}\\{current_user}"
            else:
                run_as_user = current_user

            # schtasks 명령어 구성 (관리자 권한 포함)
            cmd_args = [
                "schtasks", "/create",
                "/tn", task_name,
                "/sc", sc_type,
                "/tr", base_cmd,
                "/st", formatted_time,
                "/rl", "HIGHEST",     # 최고 권한 레벨 (관리자 권한)
                "/ru", run_as_user,   # 현재 사용자로 실행
                "/it",                # Interactive token (로그인 시에만 실행)
                "/f"                  # 기존 작업이 있으면 덮어쓰기
            ]
            
            # 추가 옵션들
            if sc_type == "WEEKLY":
                cmd_args.extend(["/d", "SUN"])
            elif sc_type == "MONTHLY":
                cmd_args.extend(["/d", "1"])

            # 로그에 실행할 명령어 기록
            self.write_detailed_log(f"schtasks command with admin privileges: {' '.join(cmd_args)}")
            self.write_detailed_log(f"Task will run as: {run_as_user} with HIGHEST privileges")
            
            # 명령어 실행
            try:
                result = subprocess.run(
                    cmd_args, 
                    check=False,
                    shell=False, 
                    capture_output=True, 
                    text=True, 
                    encoding='utf-8',
                    errors='ignore',
                    timeout=30
                )
                
                # 결과 로깅
                self.write_detailed_log(f"schtasks return code: {result.returncode}")
                if result.stdout:
                    self.write_detailed_log(f"schtasks stdout: {result.stdout}")
                if result.stderr:
                    self.write_detailed_log(f"schtasks stderr: {result.stderr}")
                    
                # 성공 여부 확인
                if result.returncode == 0:
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
                    
                    try:
                        messagebox.showinfo("Schedule Created", success_msg)
                    except Exception as mb_error:
                        self.write_detailed_log(f"messagebox.showinfo failed: {mb_error}")
                        print(f"Task created successfully with admin privileges: {task_name}")
                        
                    return True
                    
                else:
                    # 실패한 경우 상세한 오류 분석
                    error_details = self._analyze_schtasks_error(result.returncode, result.stderr, result.stdout)
                    raise RuntimeError(error_details)
                    
            except subprocess.TimeoutExpired:
                error_msg = "Task creation timed out. The system may be busy."
                self.write_detailed_log(f"schtasks command timed out")
                raise RuntimeError(error_msg)
                
        except Exception as e:
            error_msg = f"Failed to create scheduled task '{task_name}': {str(e)}"
            self.write_detailed_log(f"Task creation failed: {e}")
            
            # 상세한 오류 정보를 로그에 기록
            import traceback
            self.write_detailed_log(f"Full error traceback: {traceback.format_exc()}")
            
            print(f"ERROR: {error_msg}")
            raise RuntimeError(error_msg) from e

    def _analyze_schtasks_error(self, return_code, stderr, stdout):
        """schtasks 오류 코드 분석 및 해결방안 제시 (관리자 권한 관련 추가)"""
        error_msg = f"schtasks failed with return code {return_code}"
        
        # 일반적인 오류 코드들과 해결방안
        error_solutions = {
            1: "General error. Check if task name is valid and you have admin privileges.",
            2: "Access denied. Try running as administrator.",
            267: "Invalid task name or path. Avoid special characters in task name.",
            2147942487: "Access denied. Administrator privileges required.",
            2147943711: "Object already exists. Task with this name already exists.",
            -2147024809: "Invalid parameter. Check task settings.",
            -2147216609: "Task Scheduler service is not available.",
        }
        
        # stderr에서 특정 오류 패턴 검색 (관리자 권한 관련 추가)
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
        
        # 알려진 오류 코드에 대한 해결책 추가
        if return_code in error_solutions:
            error_msg += f"\n\n{error_solutions[return_code]}"
        
        # 관리자 권한 관련 추가 안내
        if return_code in [2, 2147942487] or "access" in (stderr or "").lower():
            error_msg += (
                "\n\nAdditional Info:"
                "\n- Make sure ezBAK is running as Administrator"
                "\n- The scheduled task will run with HIGHEST privilege level"
                "\n- Check Windows Task Scheduler service is running"
            )
        
        # stderr와 stdout 내용 추가
        if stderr:
            error_msg += f"\n\nSystem Error: {stderr}"
        if stdout:
            error_msg += f"\n\nSystem Output: {stdout}"
            
        return error_msg

    # ScheduleBackupDialog의 시간 검증도 개선
    def _validate_time_input(self, time_str):
        """시간 입력 검증 (ScheduleBackupDialog에서 사용)"""
        try:
            from datetime import datetime
            time_obj = datetime.strptime(time_str.strip(), "%H:%M")
            return True, time_obj.strftime("%H:%M")
        except ValueError:
            return False, "Invalid time format. Please use HH:MM (e.g., 14:30 for 2:30 PM)"

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
        """
        용량 체크를 별도 스레드에서 실행 (UI 응답성 향상)
        """
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

        # 3) 별도 스레드에서 용량 계산 시작
        self.message_queue.put(('log', "Calculating required space..."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Calculating space..."))

        # 스레드에서 용량 계산 실행
        calc_thread = threading.Thread(target=self.run_space_calculation, args=(sources, dest), daemon=True)
        calc_thread.start()
        
    def run_space_calculation(self, sources, dest):
        """별도 스레드에서 용량 계산 실행"""
        try:
            total = 0
            file_count = 0
            
            self.message_queue.put(('update_status', "Calculating required space..."))
            
            for i, source_path in enumerate(sources):
                self.message_queue.put(('update_status', f"Checking {i+1}/{len(sources)}: {os.path.basename(source_path)}"))
                
                if os.path.isdir(source_path):
                    # 디렉토리 용량 계산
                    dir_total = 0
                    dir_files = 0
                    
                    for dirpath, dirnames, filenames in os.walk(source_path, topdown=True, onerror=lambda e: None):
                        # 숨김 디렉토리 스킵
                        dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]
                        
                        for f in filenames:
                            fp = os.path.join(dirpath, f)
                            if self.is_hidden(fp):
                                continue
                                
                            try:
                                size = os.path.getsize(fp)
                                dir_total += size
                                dir_files += 1
                                
                                # 100개 파일마다 UI 업데이트
                                if dir_files % 100 == 0:
                                    self.message_queue.put(('update_status', 
                                        f"Checking {os.path.basename(source_path)}: {dir_files} files, {self.format_bytes(dir_total)}"))
                            except (OSError, IOError):
                                continue
                    
                    total += dir_total
                    file_count += dir_files
                    
                elif os.path.isfile(source_path) and not self.is_hidden(source_path):
                    # 단일 파일
                    try:
                        size = os.path.getsize(source_path)
                        total += size
                        file_count += 1
                    except (OSError, IOError):
                        continue
            
            # 여유 공간 계산
            self.message_queue.put(('update_status', "Checking available space..."))
            free = self.get_free_space_reliable(dest)
            
            # 결과 표시
            self.message_queue.put(('space_check_result', {
                'required': total,
                'available': free,
                'file_count': file_count,
                'sources_count': len(sources)
            }))
            
        except Exception as e:
            self.message_queue.put(('show_error', f"Space calculation error: {e}"))
        finally:
            # 수정: progress bar 
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

            # 수정된 winget export 명령어 - 문제가 되는 옵션 제거 --accept-package-agreements
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
                # 첫 번째 대안: 더 기본적인 winget export 시도
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

                # 두 번째 대안: winget list fallback
                self.write_detailed_log("Attempting fallback: winget list.")
                fb_name = f"winget_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                fb_path = os.path.join(folder, fb_name)
                
                try:
                    # 더 안전한 winget list 명령어
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
                
                # 세 번째 대안: 매우 기본적인 winget list
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
                    
                # 모든 시도가 실패한 경우
                raise RuntimeError(f"All winget export/list attempts failed. Primary error code: {rc}")

            # winget export가 성공한 경우
            self.write_detailed_log(f"Winget export completed successfully: {output_path}")
            self.message_queue.put(('log', f"Winget export completed: {output_path}"))
            self.message_queue.put(('open_folder', os.path.dirname(output_path)))
            
        except FileNotFoundError:
            self.write_detailed_log("winget not found on system.")
            self.message_queue.put(('show_error', "winget is not available. Install 'App Installer' from Microsoft Store."))
        except Exception as e:
            self.write_detailed_log(f"Winget export error: {e}")
            # 더 구체적인 오류 메시지와 해결 방안 제시
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
            # 타임스탬프 생성
            from datetime import datetime
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Copy Data용 폴더명: 'copy_data_년월일_시분초'
            copy_folder_name = f"copy_data_{timestamp_str}"
            final_destination = os.path.join(destination_dir, copy_folder_name)
            
            # 폴더 생성 (기존 폴더 삭제하지 않음)
            os.makedirs(final_destination, exist_ok=True)
            self.write_detailed_log(f"Created copy destination folder: {final_destination}")
            
            # --- 총 용량 계산 ---
            self.message_queue.put(('update_status', "Calculating total size..."))
            total_bytes = 0
            for path in source_paths:
                total_bytes += self.compute_total_bytes_safe(path)

            self.message_queue.put(('update_progress_max', total_bytes))
            self.write_detailed_log(f"Copy total bytes: {total_bytes}")
            self.message_queue.put(('log', f"Found {self.format_bytes(total_bytes)} to copy."))

            # --- 대상 여유 공간 검사 ---
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

            # --- 진행률 콜백 ---
            def progress_cb(delta):
                nonlocal bytes_copied
                bytes_copied += delta
                now = time.time()
                if now - self._last_progress_ts >= self.progress_throttle_secs:
                    self._last_progress_ts = now
                    self.message_queue.put(('update_progress', bytes_copied))

            # --- 실제 복사 루프 ---
            files_processed = 0
            for path in source_paths:
                if os.path.isdir(path):
                    # 디렉토리의 경우: 원본 디렉토리명을 유지하면서 복사
                    dir_name = os.path.basename(path)
                    dest_base = os.path.join(final_destination, dir_name)
                    
                    for dirpath, dirnames, filenames in os.walk(path, topdown=True, onerror=None):
                        dirnames[:] = [d for d in dirnames if not self.is_hidden(os.path.join(dirpath, d))]

                        # 목적지 디렉토리 생성
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
                    # 단일 파일의 경우
                    dest_file = os.path.join(final_destination, os.path.basename(path))
                    self.write_detailed_log(f"Copying file: {path} -> {dest_file}")
                    self.copy_file_with_progress_safe(path, dest_file, progress_cb)
                    files_processed += 1

            # 최종 진행률
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
            # 항상 실행되는 정리 작업
            try:
                self.close_log_file()
            except Exception:
                pass
            
            try:
                # 진행률을 100%로 설정
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass
            
            # 버튼 활성화
            self.message_queue.put(('enable_buttons', None))
            
            # 최종 상태 메시지
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
        self.open_log_file(folder_selected, "driver_backup")

        self.message_queue.put(('log', "Driver backup process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver backup in progress..."))

        backup_thread = threading.Thread(target=self.run_driver_backup, args=(folder_selected,), daemon=True)
        backup_thread.start()

    def run_driver_backup(self, folder_selected):
        """드라이버 백업 실행 (커스텀 폴더명 적용)"""
        try:
            # 시스템 정보 가져오기
            system_name = get_system_info()
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            
            # 커스텀 백업 폴더명 생성
            backup_folder_name = f"{system_name}_drivers_{timestamp}"
            backup_full_path = os.path.join(folder_selected, backup_folder_name)
            
            # 백업 폴더 생성
            os.makedirs(backup_full_path, exist_ok=True)
            
            self.write_detailed_log(f"Driver backup folder created: {backup_full_path}")
            self.message_queue.put(('log', f"Backup folder: {backup_folder_name}"))
            
            # pnputil 명령어 실행 (백업 폴더로 직접 출력)
            cmd = f'pnputil.exe /export-driver * "{backup_full_path}"'
            self.write_detailed_log(f"Executing command: {cmd}")
            
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, 
                                     stderr=subprocess.PIPE, universal_newlines=True, encoding='cp949')

            # 실시간 출력 로깅
            while process.poll() is None:
                line = process.stdout.readline()
                if line:
                    self.write_detailed_log(line.strip())
            
            # 남은 출력 처리
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
            else:
                self.write_detailed_log(f"Driver backup completed with return code: {process.returncode}")
                self.message_queue.put(('log', f"Driver backup finished (check log for details)"))
                
        except Exception as e:
            self.write_detailed_log(f"Driver backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during driver backup: {e}"))
        finally:
            # 정리 작업
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
            self.message_queue.put(('update_status', "Driver backup complete!"))

    # 결과 폴더명 예시:
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

class ScheduleBackupDialog(tk.Toplevel, DialogShortcuts):
    def __init__(self, parent):
        print("DEBUG: ScheduleBackupDialog.__init__ started")
        
        # 초기화 순서 중요: 먼저 기본 속성들을 설정
        self.action = None
        self.result = None
        self.task_name = None
        self.parent = parent
        
        try:
            tk.Toplevel.__init__(self, parent)  # 명시적으로 Toplevel 초기화
            print("DEBUG: Toplevel initialized")
            
            self.title("Schedule Backup")
            self.configure(bg="#2D3250")
            
            # 창이 닫힐 때 안전하게 처리 - 가장 먼저 설정
            self.protocol("WM_DELETE_WINDOW", self.on_close)
            print("DEBUG: WM_DELETE_WINDOW protocol set")
            
            # 아이콘 설정 (실패해도 계속 진행)
            try:
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except Exception as icon_error:
                print(f"DEBUG: Icon setting failed: {icon_error}")
                pass
                
            # 창 크기 및 위치 설정
            w, h = 546, 290
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

            # UI 구성요소들 생성
            print("DEBUG: Creating UI widgets...")
            self._create_widgets()
            print("DEBUG: UI widgets created successfully")
            
            # 단축키 설정 (UI 생성 후)
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

    def _setup_schedule_shortcuts(self):
        """스케줄 다이얼로그 전용 단축키 설정"""
        # 기본 다이얼로그 단축키 적용
        self.setup_dialog_shortcuts()
        
        # 스케줄 다이얼로그 전용 단축키들
        self.bind('<Control-c>', lambda e: self.on_create())
        self.bind('<Control-d>', lambda e: self.on_delete())
        self.bind('<Control-b>', lambda e: self._browse_destination())
        self.bind('<F1>', lambda e: self._show_schedule_help())
        
        # 숫자키로 스케줄 타입 선택
        self.bind('<Alt-1>', lambda e: self._set_schedule("Daily"))
        self.bind('<Alt-2>', lambda e: self._set_schedule("Weekly"))
        self.bind('<Alt-3>', lambda e: self._set_schedule("Monthly"))

    def _create_widgets(self):
        """UI 위젯들을 생성 (단축키 표기 포함)"""
        try:
            # 사용자 이름 안전하게 가져오기
            user_name = "Unknown"
            try:
                if hasattr(self.parent, 'user_var') and self.parent.user_var.get():
                    user_name = self.parent.user_var.get()
            except Exception:
                pass

            # Task Name
            row = 0
            tk.Label(self, text="Task Name", bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            self.task_var = tk.StringVar(value=f"ezBAK_Backup_{user_name}")
            tk.Entry(self, textvariable=self.task_var, width=40).grid(
                row=row, column=1, columnspan=2, sticky="w", padx=10, pady=5
            )

            # Destination Folder (단축키 표기 추가)
            row += 1
            tk.Label(self, text="Destination Folder", bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            self.dest_var = tk.StringVar()
            tk.Entry(self, textvariable=self.dest_var, width=30).grid(
                row=row, column=1, sticky="w", padx=10, pady=5
            )
            tk.Button(
                self, text=self.add_shortcut_text("Browse...", "Ctrl+B"), 
                command=self._browse_destination
            ).grid(row=row, column=2, padx=5, pady=5, sticky="w")

            # Schedule (단축키 표기 추가)
            row += 1
            tk.Label(self, text=self.add_shortcut_text("Schedule", "Alt+1/2/3"), 
                    bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            self.schedule_var = tk.StringVar(value="Daily")
            schedule_combo = ttk.Combobox(
                self, textvariable=self.schedule_var,
                values=["Daily", "Weekly", "Monthly"], state="readonly", width=15
            )
            schedule_combo.grid(row=row, column=1, sticky="w", padx=10, pady=5)

            # Time
            row += 1
            tk.Label(self, text="Time (HH:MM)", bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            self.time_var = tk.StringVar(value="02:00")
            tk.Entry(self, textvariable=self.time_var, width=15).grid(
                row=row, column=1, sticky="w", padx=10, pady=5
            )

            # Attributes - 한 줄에 배치
            row += 1
            tk.Label(self, text="Attributes", bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            
            attr_frame = tk.Frame(self, bg="#2D3250")
            attr_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=10, pady=5)
            
            self.hidden_var = tk.BooleanVar(value=False)
            self.system_var = tk.BooleanVar(value=False)
            
            tk.Checkbutton(attr_frame, text="Include Hidden", variable=self.hidden_var,
                           bg="#2D3250", fg="white", selectcolor="#2D3250").pack(side="left")
            
            tk.Checkbutton(attr_frame, text="Include System", variable=self.system_var,
                           bg="#2D3250", fg="white", selectcolor="#2D3250").pack(side="left", padx=(15, 0))

            # Backups to Keep
            row += 1
            tk.Label(self, text="Backups to Keep (0=all):", bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            self.retention_count_var = tk.StringVar(value="2")
            tk.Spinbox(self, from_=0, to=99, textvariable=self.retention_count_var, width=8).grid(
                row=row, column=1, sticky="w", padx=10, pady=5
            )

            # Logs to Keep
            row += 1
            tk.Label(self, text="Logs to Keep (days, 0=disable):", bg="#2D3250", fg="white").grid(
                row=row, column=0, sticky="w", padx=10, pady=5
            )
            self.log_retention_days_var = tk.StringVar(value="30")
            tk.Spinbox(self, from_=0, to=365, textvariable=self.log_retention_days_var, width=8).grid(
                row=row, column=1, sticky="w", padx=10, pady=5
            )

            # 버튼 영역 (단축키 표기 추가)
            row += 1
            btn_frame = tk.Frame(self, bg="#2D3250")
            btn_frame.grid(row=row, column=0, columnspan=3, sticky="e", padx=10, pady=15)

            tk.Button(btn_frame, text=self.add_shortcut_text("Create", "Ctrl+C"), 
                      command=self.on_create,
                      bg="#1E88E5", fg="white", relief="flat", width=15).pack(side="right", padx=5)
            tk.Button(btn_frame, text=self.add_shortcut_text("Delete", "Ctrl+D"), 
                      command=self.on_delete,
                      bg="#FF5733", fg="white", relief="flat", width=15).pack(side="right", padx=5)
            tk.Button(btn_frame, text=self.add_shortcut_text("Close", "Esc"), 
                      command=self.on_close,
                      bg="#607D8B", fg="white", relief="flat", width=15).pack(side="right", padx=5)
                      
        except Exception as widget_error:
            print(f"DEBUG: Widget creation failed: {widget_error}")
            raise widget_error

    def _set_schedule(self, schedule_type):
        """단축키로 스케줄 타입 설정"""
        try:
            self.schedule_var.set(schedule_type)
            print(f"DEBUG: Schedule set to {schedule_type}")
        except Exception as e:
            print(f"DEBUG: Failed to set schedule: {e}")

    def _show_schedule_help(self):
        """F1키로 스케줄 도움말 표시"""
        help_text = """
Schedule Backup Shortcuts:

Ctrl + C     Create Task
Ctrl + D     Delete Task  
Ctrl + B     Browse Destination
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
        help_dlg.configure(bg="#2D3250")
        help_dlg.geometry("350x300")
        help_dlg.transient(self)
        try:
            help_dlg.grab_set()
        except Exception:
            pass
        
        text_widget = tk.Text(help_dlg, 
                             font=("Consolas", 10),  
                             bg="#424769", 
                             fg="white",
                             relief="flat",
                             wrap="word")
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        
        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")
        
        close_btn = tk.Button(help_dlg, 
                             text="Close (Esc)", 
                             command=help_dlg.destroy,
                             bg="#4CAF50", 
                             fg="white", 
                             relief="flat")
        close_btn.pack(pady=10)
        
        help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
        help_dlg.focus_set()

    # 기존 메서드들 유지
    def _browse_destination(self):
        """폴더 선택 다이얼로그"""
        try:
            folder = filedialog.askdirectory(parent=self, title="Select Backup Destination")
            if folder:
                self.dest_var.set(folder)
                print(f"DEBUG: Destination set to: {folder}")
        except Exception as browse_error:
            print(f"DEBUG: Browse failed: {browse_error}")

    def on_create(self):
        """생성 버튼 핸들러"""
        print("DEBUG: on_create called")
        try:
            # 입력값 검증
            task_name = self.task_var.get().strip()
            dest = self.dest_var.get().strip()
            
            if not task_name:
                messagebox.showerror("Error", "Task name is required.", parent=self)
                return
                
            if not dest:
                messagebox.showerror("Error", "Destination folder is required.", parent=self)
                return
                
            if not os.path.isdir(dest):
                if not messagebox.askyesno("Warning", 
                    f"Destination folder does not exist:\n{dest}\n\nContinue anyway?", 
                    parent=self):
                    return
            
            # 시간 형식 검증
            time_str = self.time_var.get().strip()
            try:
                from datetime import datetime
                time_obj = datetime.strptime(time_str, "%H:%M")
                time_str = time_obj.strftime("%H:%M")
            except ValueError:
                messagebox.showerror("Error", "Invalid time format. Use HH:MM (e.g., 14:30)", parent=self)
                return

            # 사용자 이름 가져오기
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
                "include_hidden": self.hidden_var.get(),
                "include_system": self.system_var.get(),
                "retention_count": int(self.retention_count_var.get()),
                "log_retention_days": int(self.log_retention_days_var.get()),
                "user": user_name
            }
            print(f"DEBUG: Result set: {self.result}")
            self._safe_destroy()
            
        except Exception as create_error:
            print(f"DEBUG: on_create failed: {create_error}")
            try:
                messagebox.showerror("Error", f"Failed to create schedule: {create_error}", parent=self)
            except Exception:
                print(f"ERROR: {create_error}")

    def on_delete(self):
        """삭제 버튼 핸들러"""
        print("DEBUG: on_delete called")
        try:
            task_name = self.task_var.get().strip()
            if not task_name:
                messagebox.showerror("Error", "Task name is required for deletion.", parent=self)
                return
                
            self.action = "delete"
            self.task_name = task_name
            print(f"DEBUG: Delete task_name set: {task_name}")
            self._safe_destroy()
        except Exception as delete_error:
            print(f"DEBUG: on_delete failed: {delete_error}")
            try:
                messagebox.showerror("Error", f"Failed to delete schedule: {delete_error}", parent=self)
            except Exception:
                print(f"ERROR: {delete_error}")

    def on_close(self):
        """닫기 버튼 핸들러"""
        print("DEBUG: on_close called")
        self.action = "close"
        self._safe_destroy()
        
    def on_cancel(self):
        """ESC키용 취소 핸들러"""
        self.on_close()
        
    def on_ok(self):
        """Enter키용 확인 핸들러"""
        self.on_create()
        
    def _safe_destroy(self):
        """안전한 윈도우 파괴"""
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
        self.title("Filter Manager")
        try:
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        self.configure(bg="#2D3250")
        self.geometry("720x480")
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

    def _setup_filter_shortcuts(self):
        """필터 다이얼로그 전용 단축키 설정"""
        # 기본 다이얼로그 단축키 적용
        self.setup_dialog_shortcuts()
        
        # 필터 관련 단축키들
        self.bind('<Control-n>', lambda e: self._add_rule('include'))  # New include rule
        self.bind('<Control-e>', lambda e: self._add_rule('exclude'))  # New exclude rule
        self.bind('<Delete>', lambda e: self._remove_selected_active())
        self.bind('<Control-s>', lambda e: self._save())
        self.bind('<Control-r>', lambda e: self._clear_active())  # Remove all from active list
        self.bind('<F1>', lambda e: self._show_filter_help())
        
        # 리스트 포커스 관리
        self.bind('<Tab>', lambda e: self._cycle_focus())
        
    def _cycle_focus(self):
        """Tab키로 리스트 간 포커스 이동"""
        try:
            focused = self.focus_get()
            if focused == self.inc_list:
                self.exc_list.focus_set()
            else:
                self.inc_list.focus_set()
        except Exception:
            self.inc_list.focus_set()

    def _create_filter_widgets(self):
        """필터 위젯들 생성 (단축키 표기 포함)"""
        # 상단 설명
        info_frame = tk.Frame(self, bg="#2D3250")
        info_frame.pack(fill='x', padx=12, pady=(12,6))
        
        info_text = "Include: Files must match at least one rule to be included\nExclude: Files matching any rule will be skipped (takes priority)"
        tk.Label(info_frame, text=info_text, bg="#2D3250", fg="#CCCCCC", 
                font=("Arial", 9), justify='left').pack(anchor='w')

        # Layout
        top = tk.Frame(self, bg="#2D3250")
        top.pack(fill='both', expand=True, padx=12, pady=6)

        # Include panel (단축키 표기 추가)
        inc_frame = tk.LabelFrame(top, text=self.add_shortcut_text("Include Rules", "Ctrl+N"), 
                                 bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
        inc_frame.pack(side='left', fill='both', expand=True, padx=(0,6))
        
        self.inc_list = tk.Listbox(inc_frame, height=14, font=("Consolas", 9),
                                  bg="#424769", fg="white", selectbackground="#6EC571")
        self.inc_list.pack(fill='both', expand=True, padx=6, pady=6)
        
        btn_row_inc = tk.Frame(inc_frame, bg="#2D3250")
        btn_row_inc.pack(fill='x', padx=6, pady=(0,6))
        
        tk.Button(btn_row_inc, text="Add", width=8, command=lambda: self._add_rule('include'), 
                 bg="#336BFF", fg="white", relief="flat").pack(side='left')
        tk.Button(btn_row_inc, text=self.add_shortcut_text("Remove", "Del"), width=12, 
                 command=lambda: self._remove_selected('include'), 
                 bg="#FF5733", fg="white", relief="flat").pack(side='left', padx=(6,0))
        tk.Button(btn_row_inc, text="Clear All", width=10, command=lambda: self._clear('include'), 
                 bg="#7D98A1", fg="white", relief="flat").pack(side='left', padx=(6,0))

        # Exclude panel (단축키 표기 추가)
        exc_frame = tk.LabelFrame(top, text=self.add_shortcut_text("Exclude Rules", "Ctrl+E"), 
                                 bg="#2D3250", fg="white", font=("Arial", 10, "bold"))
        exc_frame.pack(side='left', fill='both', expand=True, padx=(6,0))
        
        self.exc_list = tk.Listbox(exc_frame, height=14, font=("Consolas", 9),
                                  bg="#424769", fg="white", selectbackground="#FF7D5A")
        self.exc_list.pack(fill='both', expand=True, padx=6, pady=6)
        
        btn_row_exc = tk.Frame(exc_frame, bg="#2D3250")
        btn_row_exc.pack(fill='x', padx=6, pady=(0,6))
        
        tk.Button(btn_row_exc, text="Add", width=8, command=lambda: self._add_rule('exclude'), 
                 bg="#336BFF", fg="white", relief="flat").pack(side='left')
        tk.Button(btn_row_exc, text=self.add_shortcut_text("Remove", "Del"), width=12, 
                 command=lambda: self._remove_selected('exclude'), 
                 bg="#FF5733", fg="white", relief="flat").pack(side='left', padx=(6,0))
        tk.Button(btn_row_exc, text="Clear All", width=10, command=lambda: self._clear('exclude'), 
                 bg="#7D98A1", fg="white", relief="flat").pack(side='left', padx=(6,0))

        # 하단 단축키 안내
        shortcut_frame = tk.Frame(self, bg="#2D3250")
        shortcut_frame.pack(fill='x', padx=12, pady=(6,0))
        
        shortcut_text = "F1: Help  •  Tab: Switch Lists  •  Del: Remove Selected  •  Ctrl+R: Clear Active List"
        tk.Label(shortcut_frame, text=shortcut_text, bg="#2D3250", fg="#888888", 
                font=("Arial", 8)).pack(anchor='w')

        # Bottom actions (단축키 표기 추가)
        actions = tk.Frame(self, bg="#2D3250")
        actions.pack(fill='x', padx=12, pady=12)
        
        tk.Button(actions, text=self.add_shortcut_text("Help", "F1"), width=10, 
                 command=self._show_filter_help, bg="#8E44AD", fg="white", relief="flat").pack(side='left')
        
        tk.Button(actions, text=self.add_shortcut_text("Save", "Ctrl+S"), width=12, 
                 command=self._save, bg="#2E7D32", fg="white", relief="flat").pack(side='right')
        tk.Button(actions, text=self.add_shortcut_text("Cancel", "Esc"), width=12, 
                 command=self._cancel, bg="#7D98A1", fg="white", relief="flat").pack(side='right', padx=(0,8))

    def _rule_to_text(self, r):
        """규칙을 표시용 텍스트로 변환"""
        try:
            rule_type = r.get('type', 'unknown')
            pattern = r.get('pattern', '')
            
            # 타입별 아이콘 추가
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
        """리스트 새로고침"""
        # Include 리스트
        self.inc_list.delete(0, tk.END)
        for r in self.inc:
            self.inc_list.insert(tk.END, self._rule_to_text(r))
        
        # Exclude 리스트
        self.exc_list.delete(0, tk.END)
        for r in self.exc:
            self.exc_list.insert(tk.END, self._rule_to_text(r))

    def _remove_selected(self, kind):
        """선택된 항목 제거"""
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
        """현재 활성화된 리스트에서 선택된 항목 제거"""
        try:
            # 포커스된 위젯 확인
            focused = self.focus_get()
            if focused == self.inc_list:
                self._remove_selected('include')
            elif focused == self.exc_list:
                self._remove_selected('exclude')
        except Exception:
            pass

    def _clear(self, kind):
        """전체 지우기"""
        if kind == 'include':
            self.inc = []
        else:
            self.exc = []
        self._refresh()

    def _clear_active(self):
        """현재 활성화된 리스트 전체 지우기"""
        try:
            focused = self.focus_get()
            if focused == self.inc_list:
                self._clear('include')
            elif focused == self.exc_list:
                self._clear('exclude')
        except Exception:
            pass

    def _add_rule(self, kind):
        """규칙 추가 다이얼로그"""
        dlg = tk.Toplevel(self)
        dlg.title(f"Add {'Include' if kind == 'include' else 'Exclude'} Rule")
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass
        dlg.configure(bg="#2D3250")
        dlg.geometry("480x320")
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass

        # 규칙 타입 선택
        tk.Label(dlg, text="Rule Type:", bg="#2D3250", fg="white", 
                font=("Arial", 10, "bold")).pack(anchor='w', padx=15, pady=(15,5))
        
        vtype = tk.StringVar(value='ext')
        tframe = tk.Frame(dlg, bg="#2D3250")
        tframe.pack(anchor='w', padx=15, fill='x')
        
        type_options = [
            ('ext', '📄 Extension', 'File extensions (e.g., tmp, log, bak)'),
            ('name', '📁 Name', 'File/folder names with wildcards (e.g., Thumbs.db, *cache*)'),
            ('path', '🗂️ Path', 'Full paths with wildcards (e.g., */temp/*, *\\AppData\\*)')
        ]
        
        for value, label, desc in type_options:
            rb_frame = tk.Frame(tframe, bg="#2D3250")
            rb_frame.pack(anchor='w', pady=2)
            
            tk.Radiobutton(rb_frame, text=label, variable=vtype, value=value, 
                          bg="#2D3250", fg="white", selectcolor="#2D3250",
                          font=("Arial", 9, "bold")).pack(side='left')
            tk.Label(rb_frame, text=f"  {desc}", bg="#2D3250", fg="#CCCCCC", 
                    font=("Arial", 8)).pack(side='left')

        # 패턴 입력
        tk.Label(dlg, text="Pattern:", bg="#2D3250", fg="white", 
                font=("Arial", 10, "bold")).pack(anchor='w', padx=15, pady=(15,5))
        
        vpat = tk.StringVar()
        pat_frame = tk.Frame(dlg, bg="#2D3250")
        pat_frame.pack(fill='x', padx=15)
        
        ent = tk.Entry(pat_frame, textvariable=vpat, width=50, font=("Arial", 10))
        ent.pack(fill='x')
        
        # 예시 텍스트
        example_frame = tk.Frame(dlg, bg="#2D3250")
        example_frame.pack(fill='x', padx=15, pady=(5,0))
        
        def update_example(*args):
            try:
                rule_type = vtype.get()
                examples = {
                    'ext': 'Examples: tmp, log, bak, cache (without dots)',
                    'name': 'Examples: Thumbs.db, *.tmp, desktop.ini, *cache*',
                    'path': 'Examples: */Temp/*, *\\AppData\\Local\\*, */cache/*'
                }
                example_label.config(text=examples.get(rule_type, ''))
            except Exception:
                pass
        
        example_label = tk.Label(example_frame, text="", bg="#2D3250", fg="#888888", 
                               font=("Arial", 8))
        example_label.pack(anchor='w')
        
        vtype.trace('w', update_example)
        update_example()  # 초기 표시

        # 버튼
        btns = tk.Frame(dlg, bg="#2D3250")
        btns.pack(fill='x', padx=15, pady=15)
        
        def _ok():
            p = vpat.get().strip()
            if not p:
                messagebox.showwarning("Warning", "Please enter a pattern.", parent=dlg)
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
        
        tk.Button(btns, text="Add Rule", width=12, command=_ok, 
                 bg="#336BFF", fg="white", relief="flat").pack(side='right')
        tk.Button(btns, text="Cancel", width=10, command=_cancel, 
                 bg="#7D98A1", fg="white", relief="flat").pack(side='right', padx=(0,8))

        # 단축키 설정
        dlg.bind('<Return>', lambda e: _ok())
        dlg.bind('<Escape>', lambda e: _cancel())
        
        try:
            ent.focus_set()
        except Exception:
            pass

    def _show_filter_help(self):
        """필터 도움말 표시"""
        help_text = """
Filter Manager Help

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
• Exclude system paths: [PATH] */Windows/System32/*
        """
        
        help_dlg = tk.Toplevel(self)
        help_dlg.title("Filter Manager Help")
        help_dlg.configure(bg="#2D3250")
        help_dlg.geometry("600x610")
        help_dlg.transient(self)
        try:
            help_dlg.grab_set()
        except Exception:
            pass
        
        # 스크롤 가능한 텍스트 영역
        text_frame = tk.Frame(help_dlg, bg="#2D3250")
        text_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        scrollbar = tk.Scrollbar(text_frame, width=20, bg="#424769", 
                               troughcolor="#2D3250", activebackground="#6EC571")
        scrollbar.pack(side="right", fill="y")
        
        text_widget = tk.Text(text_frame, 
                             font=("Consolas", 9),  
                             bg="#424769", 
                             fg="white",
                             relief="flat",
                             wrap="word",
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)
        
        scrollbar.configure(command=text_widget.yview)
        
        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")
        
        # 닫기 버튼
        close_btn = tk.Button(help_dlg, 
                             text="Close (Esc)", 
                             command=help_dlg.destroy,
                             bg="#4CAF50", 
                             fg="white", 
                             relief="flat",
                             width=15)
        close_btn.pack(pady=(0,15))
        
        help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
        help_dlg.focus_set()

    def _save(self):
        """필터 저장"""
        self.result = {'include': self.inc, 'exclude': self.exc}
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _cancel(self):
        """취소"""
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def on_cancel(self):
        """ESC키용 취소 핸들러"""
        self._cancel()
        
    def on_ok(self):
        """Enter키용 확인 핸들러"""
        self._save()

class SelectSourcesDialog(tk.Toplevel):
    def __init__(self, master, is_hidden_fn=None, title="Select Sources"):
        try:
            super().__init__(master)
            print(f"DEBUG: SelectSourcesDialog initialized with title: {title}")
            
            self.title(title)
            try:
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except Exception:
                pass
            self.configure(bg="#2D3250")
            self.geometry("720x540")
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
        """소스 선택 다이얼로그 전용 단축키 설정"""
        try:
            # 기본 다이얼로그 단축키 적용
            self.setup_dialog_shortcuts()
            
            # 소스 선택 관련 단축키들
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
        """다이얼로그에 기본 단축키 설정 (DialogShortcuts 대신 직접 구현)"""
        try:
            # 기본 다이얼로그 단축키들
            self.bind('<Escape>', lambda e: self.on_cancel() if hasattr(self, 'on_cancel') else self.destroy())
            self.bind('<Return>', lambda e: self.on_ok() if hasattr(self, 'on_ok') else None)
            self.bind('<Alt-F4>', lambda e: self.destroy())
            
            # 포커스 설정 (키보드 이벤트를 받기 위해)
            self.focus_set()
        except Exception as e:
            print(f"DEBUG: setup_dialog_shortcuts failed: {e}")
        
    def add_shortcut_text(self, text, shortcut):
        """텍스트에 단축키 표기 추가"""
        return f"{text} ({shortcut})"

    def _create_source_widgets(self):
        """소스 선택 위젯들 생성"""
        try:
            print("DEBUG: Creating info frame...")
            # 상단 설명
            info_frame = tk.Frame(self, bg="#2D3250")
            info_frame.pack(fill="x", padx=15, pady=(15, 10))
            
            tk.Label(info_frame, text="Select folders and files to copy",
                     bg="#2D3250", fg="white", font=("Arial", 12, "bold")).pack(anchor='w')
            
            tk.Label(info_frame, text="Use checkboxes or Space key to select items. Selected folders include all contents.",
                     bg="#2D3250", fg="#CCCCCC", font=("Arial", 9)).pack(anchor='w', pady=(2,0))

            print("DEBUG: Creating top controls...")
            # Top controls
            top_controls = tk.Frame(self, bg="#2D3250")
            top_controls.pack(fill="x", padx=15, pady=(0, 10))

            self.add_root_btn = tk.Button(top_controls, text=self.add_shortcut_text("Add Root Folder...", "Ctrl+R"), 
                                          command=self._choose_root,
                                          bg="#7D98A1", fg="white", relief="flat", width=20)
            self.add_root_btn.pack(side="left")

            self.sel_all_btn = tk.Button(top_controls, text=self.add_shortcut_text("Select All", "Ctrl+A"), 
                                         command=self._select_all,
                                         bg="#4CAF50", fg="white", relief="flat", width=15)
            self.sel_all_btn.pack(side="left", padx=(8, 0))

            self.clear_all_btn = tk.Button(top_controls, text=self.add_shortcut_text("Clear All", "Ctrl+N"), 
                                           command=self._clear_all,
                                           bg="#FF5733", fg="white", relief="flat", width=15)
            self.clear_all_btn.pack(side="left", padx=(8, 0))

            print("DEBUG: Creating tree area...")
            # Tree area
            tree_frame = tk.Frame(self, bg="#2D3250")
            tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

            self.tree = ttk.Treeview(tree_frame, columns=("fullpath",), displaycolumns=())
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self.tree.pack(side="left", fill="both", expand=True)

            self.tree.bind("<<TreeviewOpen>>", self._on_open)
            self.tree.bind("<Button-1>", self._on_click)

            print("DEBUG: Creating status frame...")
            # 하단 상태 및 단축키 안내
            status_frame = tk.Frame(self, bg="#2D3250")
            status_frame.pack(fill='x', padx=15, pady=(5,0))
            
            self.status_label = tk.Label(status_frame, text="Ready - Use mouse or keyboard to select items", 
                                        bg="#2D3250", fg="#CCCCCC", font=("Arial", 9))
            self.status_label.pack(side='left')

            print("DEBUG: Creating action buttons...")
            # Action buttons
            action_frame = tk.Frame(self, bg="#2D3250")
            action_frame.pack(fill="x", padx=15, pady=15)

            ok_btn = tk.Button(action_frame, text=self.add_shortcut_text("OK", "Enter"), 
                              width=12, command=self._on_ok,
                              bg="#336BFF", fg="white", relief="flat")
            ok_btn.pack(side="right")
            
            cancel_btn = tk.Button(action_frame, text=self.add_shortcut_text("Cancel", "Esc"), 
                                  width=12, command=self._on_cancel,
                                  bg="#FF3333", fg="white", relief="flat")
            cancel_btn.pack(side="right", padx=(0, 8))
            
            print("DEBUG: Widget creation complete")
            
        except Exception as e:
            print(f"DEBUG: _create_source_widgets failed: {e}")
            import traceback
            print(f"DEBUG: Traceback: {traceback.format_exc()}")
            raise

    def _format(self, name, checked):
        """체크박스 형태로 포맷"""
        checkbox = "[✓]" if checked else "[ ]"
        return f"{checkbox} {name}"

    def _toggle_selected(self):
        """현재 선택된 항목 토글 (Space키용)"""
        try:
            item = self.tree.focus()
            if item:
                self._toggle_item(item)
                self._update_status()
        except Exception as e:
            print(f"DEBUG: _toggle_selected failed: {e}")

    def _expand_selected(self):
        """선택된 항목 확장"""
        try:
            item = self.tree.focus()
            if item:
                self.tree.item(item, open=True)
        except Exception as e:
            print(f"DEBUG: _expand_selected failed: {e}")

    def _collapse_selected(self):
        """선택된 항목 축소"""
        try:
            item = self.tree.focus()
            if item:
                self.tree.item(item, open=False)
        except Exception as e:
            print(f"DEBUG: _collapse_selected failed: {e}")

    def _refresh_tree(self):
        """트리 새로고침 (F5)"""
        try:
            print("DEBUG: Refreshing tree...")
            # 현재 체크된 항목들 저장
            checked_paths = []
            for item in self._checked:
                path = self._item_path.get(item)
                if path:
                    checked_paths.append(path)
            
            # 트리 초기화
            for item in self.tree.get_children(""):
                self.tree.delete(item)
            self._checked.clear()
            self._item_path.clear()
            
            # 드라이브 다시 추가
            self._add_drive_roots()
            
            self._update_status("Tree refreshed")
            print("DEBUG: Tree refresh complete")
        except Exception as e:
            print(f"DEBUG: _refresh_tree failed: {e}")
            self._update_status(f"Refresh failed: {e}")

    def _update_status(self, message=None):
        """상태바 업데이트"""
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
        """소스 선택 도움말 표시"""
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
            help_dlg.configure(bg="#2D3250")
            help_dlg.geometry("500x400")
            help_dlg.transient(self)
            help_dlg.grab_set()
            
            text_widget = tk.Text(help_dlg, 
                                 font=("Consolas", 10),  
                                 bg="#424769", 
                                 fg="white",
                                 relief="flat",
                                 wrap="word")
            text_widget.pack(fill="both", expand=True, padx=15, pady=15)
            
            text_widget.insert("1.0", help_text)
            text_widget.configure(state="disabled")
            
            close_btn = tk.Button(help_dlg, 
                                 text="Close (Esc)", 
                                 command=help_dlg.destroy,
                                 bg="#4CAF50", 
                                 fg="white", 
                                 relief="flat")
            close_btn.pack(pady=(0,15))
            
            help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
            help_dlg.focus_set()
            
        except Exception as e:
            print(f"DEBUG: _show_source_help failed: {e}")

    def _add_drive_roots(self):
        """드라이브 루트들을 트리에 추가"""
        try:
            print("DEBUG: Adding drive roots...")
            drives = []
            
            # 간단한 방법으로 드라이브 감지
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
        """루트 폴더 선택"""
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
        """트리 항목 열기 이벤트"""
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
        """트리 클릭 이벤트"""
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
        """항목 체크 상태 토글"""
        try:
            text = self.tree.item(item, "text")
            if not text:
                return
            checked = item in self._checked
            new_checked = not checked
            
            # 텍스트에서 이름 부분만 추출 (체크박스 부분 제거)
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]  # "[✓] " 또는 "[ ] " 제거
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
        """하위 항목들에 체크 상태 적용"""
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
        """모든 항목 선택"""
        try:
            for item in self.tree.get_children(""):
                self._set_checked_recursive(item, True)
            self._update_status()
        except Exception as e:
            print(f"DEBUG: _select_all failed: {e}")

    def _clear_all(self):
        """모든 선택 해제"""
        try:
            for item in self.tree.get_children(""):
                self._set_checked_recursive(item, False)
            self._update_status()
        except Exception as e:
            print(f"DEBUG: _clear_all failed: {e}")

    def _set_checked_recursive(self, item, checked):
        """재귀적으로 체크 상태 설정"""
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
        """확인 버튼"""
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
        """취소 버튼"""
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
        """ESC키용 취소 핸들러"""
        self._on_cancel()
        
    def on_ok(self):
        """Enter키용 확인 핸들러"""
        self._on_ok()

def core_run_backup(user_name, dest_dir, include_hidden=False, include_system=False, log_folder=None, retention_count=2, log_retention_days=30):
    """
    GUI와 스케줄러 공통 백업 로직 (상세 로그 기록 + cleanup 추가)
    run_backup()과 동일한 로직이지만 UI 업데이트 없이 실행
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

    # 기존 백업 폴더가 있으면 삭제
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

    # 숨김/시스템 파일 체크 함수 (run_backup과 동일한 로직)
    def is_hidden_headless(filepath):
        """헤드리스 모드용 is_hidden 함수"""
        if not os.path.exists(filepath):
            return False
        
        # 기본 체크들
        base = os.path.basename(filepath).lower()
        if base.startswith('.') or base in ('thumbs.db', 'desktop.ini', '$recycle.bin'):
            return True
            
        if base in ('appdata', 'application data', 'local settings'):
            if not (include_hidden and include_system):
                return True

        try:
            # Windows 속성 체크
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

    # 안전한 파일 복사 함수 (run_backup과 동일)
    def copy_file_safe_headless(src, dst, timeout_seconds=30):
        """타임아웃이 있는 안전한 파일 복사 (헤드리스용)"""
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

    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Backup start for '{user_name}' → {backup_path}\n")
        log.write(f"[POLICY] Hidden={'include' if include_hidden else 'exclude'} System={'include' if include_system else 'exclude'}\n")

        try:
            # run_backup()과 동일한 백업 로직
            for dirpath, dirnames, filenames in os.walk(user_home, topdown=True, onerror=None):
                try:
                    # 안전한 디렉토리 필터링
                    original_dirnames = dirnames.copy()
                    dirnames.clear()
                    
                    for d in original_dirnames:
                        full_dir_path = os.path.join(dirpath, d)
                        try:
                            if not is_hidden_headless(full_dir_path):
                                dirnames.append(d)
                            else:
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Skipping hidden directory: {full_dir_path}\n")
                        except Exception as e:
                            log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error checking directory {full_dir_path}: {e}\n")
                            dirnames.append(d)  # 에러 시 디렉토리 포함
                    
                    # 목적지 폴더 생성
                    rel_dir = os.path.relpath(dirpath, user_home)
                    dest_dir_path = os.path.join(backup_path, rel_dir) if rel_dir != "." else backup_path
                    
                    try:
                        os.makedirs(dest_dir_path, exist_ok=True)
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Created folder: {dest_dir_path}\n")
                        folders_processed += 1
                    except Exception as e:
                        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ERROR: Failed to create folder {dest_dir_path}: {e}\n")
                        continue

                    # 파일 복사
                    for f in filenames:
                        try:
                            src_file = os.path.join(dirpath, f)
                            dest_file = os.path.join(dest_dir_path, f)
                            
                            # 숨김 파일 체크 
                            if is_hidden_headless(src_file):
                                log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Skipping hidden file: {src_file}\n")
                                continue
                            
                            # 파일 복사
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

        # 브라우저 북마크 백업 (run_backup과 동일)
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

        # 오래된 백업 정리 (run_backup과 동일)
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

        # 오래된 로그 정리 (run_backup과 동일)
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


# 2. 필요한 헬퍼 함수들 추가

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
    에러 처리기 for shutil.rmtree (headless 모드용 - 로그에 기록)
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


# 3. 필요한 import들이 core_run_backup 함수 내에서 사용할 수 있도록 전역에서 정의되어 있는지 확인
# 현재 코드에서 이미 정의되어 있음: _have_pywin32, win32api, win32con, _get_file_attributes 등

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
        parser.add_argument("--user", help="User name to back up")
        parser.add_argument("--dest", help="Destination folder for backup")
        parser.add_argument("--include-hidden", action="store_true")
        parser.add_argument("--include-system", action="store_true")
        parser.add_argument("--retention", type=int, default=2, help="Number of backups to keep")
        parser.add_argument("--log-retention", type=int, default=30, help="Number of days to keep logs")
        parser.add_argument("--log-folder", default="logs")
        
        # 다른 인수들을 무시하도록 parse_known_args 사용
        args, unknown = parser.parse_known_args()

        if args.user and args.dest:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] Headless backup start for user '{args.user}'")
            
            # 수정: 실제 전달된 retention 값들을 로그에 표시
            print(f"[{timestamp}] Command line args: retention={args.retention}, log_retention={args.log_retention}")
            
            try:
                backup_path, log_path = core_run_backup(
                    args.user,
                    args.dest,
                    include_hidden=args.include_hidden,
                    include_system=args.include_system,
                    log_folder=args.log_folder,
                    retention_count=args.retention,  # 명령행에서 전달된 값 사용
                    log_retention_days=args.log_retention  # 명령행에서 전달된 값 사용
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