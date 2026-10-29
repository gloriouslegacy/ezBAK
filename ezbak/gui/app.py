"""App window of ezBAK (moved from ezBAK.py in P5-16) and run(entry_script).

App(<mixins>, ctk.CTk) keeps the state, Tk variables, keyboard shortcut setup (no menu bar since U5-7)
and user list; the mixins (P5-2~15) bring the rest. Module-level names (tk, ttk, os,
threading, queue, datetime, webbrowser, Win11Theme, Translator, Win11Dialog,
FilterManagerDialog, KeyboardShortcuts, copy_default_filters, resource_path,
winsys_paths) match ezBAK.py so facade patches reach this module too.
App.__init__ runs the former statements in the same order through three helpers
(_init_window, _init_state, _init_startup; P5 decision D5).
`entry_script` is the script path the scheduled task runs (ezBAK.py passes
os.path.abspath(__file__), as before).
The base class is ctk.CTk (U4-1, ADR 0004): CTk makes the process DPI aware when the
root is created (ui-plan D8, on from U8-3; until then it was switched off), App then
syncs Tk's point scale and the tk parts take their pixel sizes through gui.dpi; the
root colour is set through fg_color (CTk ignores configure(bg=)).
"""
import logging
import os
import queue
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import ttk

import customtkinter as ctk

from ezbak.core import diag
from ezbak.winsys import paths as winsys_paths  # USERS_ROOT read at call time (P5 D3)
from ezbak.winsys.paths import resource_path

from .actions_backup import BackupActionsMixin
from .actions_copy import CopyActionsMixin
from .actions_restore import RestoreActionsMixin
from .actions_schedule import ScheduleActionsMixin
from .actions_tools import ToolsActionsMixin
from .app_log import AppLogMixin
from . import dpi
from .bridge import QueueMixin
from .dialogs import FilterManagerDialog, Win11Dialog
from .dialogs.nas import NasDialogsMixin
from .dialogs.restore_select import RestoreSelectMixin
from .i18n import LanguageMixin, Translator
from .layout import LayoutMixin
from .settings import SettingsMixin, copy_default_filters
from .shortcuts import KeyboardShortcuts
from .theme import ThemeMixin, Win11Theme

_log = logging.getLogger(__name__)

class App(SettingsMixin, AppLogMixin, QueueMixin, ThemeMixin, LanguageMixin, LayoutMixin, BackupActionsMixin, CopyActionsMixin, ScheduleActionsMixin, RestoreActionsMixin, RestoreSelectMixin, ToolsActionsMixin, NasDialogsMixin, ctk.CTk):
    def __init__(self, entry_script):
        super().__init__()  # CTk makes the process DPI aware (D8, on from U8-3)
        dpi.sync_tk_scaling(self)  # Tk chose its pt scale before that (ui-plan §9)
        self.entry_script = entry_script  # script path for scheduled tasks (gui.run)

        # Hide window until fully initialized. Tk's own withdraw: CTk's override would mark
        # "withdrawn before the window exists" and its Windows mainloop then never shows it.
        tk.Tk.withdraw(self)
        self._init_window()
        self._init_state()
        self._init_startup()

    def _init_window(self):
        """Theme, translator, title, icon, geometry and ttk style (part 1 of __init__)."""
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
                pass  # no icon file: the window keeps the default Tk icon (as before)
        self.geometry("1000x700")  # ui-plan §5 (U4-1)
        self.minsize(900, 600)
        self.configure(fg_color=self.theme.get('bg'))  # CTk root: fg_color, not bg

        # UI elements setup
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Configure professional progress bar styles
        self._configure_progressbar_style()

    def _init_state(self):
        """Queue, Tk variables, filters, saved settings, log and throttle state (part 2)."""
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
        # Initialize filters with default exclude patterns
        self.filters = copy_default_filters()
        # Load persisted settings (if any)
        try:
            self.load_settings()
        except Exception:
            # unreadable settings: start with the defaults above (as before)
            _log.debug("event=settings_load_failed", exc_info=True)
        self._sync_ctk_appearance()  # CTk widgets follow the loaded theme

        # logging to file
        self.log_file_path = None
        self._log_file = None
        self.log_lock = threading.Lock()

        # progress/log throttle to keep UI responsive
        self._last_progress_ts = 0.0
        self.progress_throttle_secs = 0.12
        self._last_ui_log_ts = 0.0
        self.ui_log_throttle_secs = 0.08

    def _init_startup(self):
        """Widgets, texts, the startup notice, shortcuts and showing the window (part 3)."""
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

    def open_github_link(self, event):
        webbrowser.open_new("https://github.com/gloriouslegacy/ezBAK/releases")

    def show_keyboard_shortcuts_help(self):
        """Show keyboard shortcuts help dialog"""
        if hasattr(self, 'shortcuts') and self.shortcuts:
            self.shortcuts.show_help()

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
            pass  # optional button missing: skip it (as before)
        try:
            self.schedule_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)
        try:
            self.winget_export_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)
        try:
            self.filters_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)
        try:
            self.nas_connect_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)
        try:
            self.nas_disconnect_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)
        try:
            self.browser_profiles_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)
        try:
            self.home_backup_btn.configure(state=state)  # CTk widget: configure, not config
        except Exception:
            pass  # home page not built (snapshot host): skip it
        try:
            self.exit_btn.config(state=state)
        except Exception:
            pass  # optional button missing: skip it (as before)

    def reset_progress_bar(self):
        """Reset progress bar"""
        try:
            self.progress_bar.stop()  # stop indeterminate mode
            self.progress_bar['mode'] = "determinate"
            self.progress_bar['value'] = 0
            self.progress_bar['maximum'] = 100
            self.status_label.config(text="Ready")
        except Exception:
            pass  # widget decoration only: a failed reset leaves the bar as is (as before)

    def load_users(self):
        """Gets a list of Windows users and adds them to the combo box."""
        users_path = winsys_paths.USERS_ROOT
        try:
            users = [d for d in os.listdir(users_path) if os.path.isdir(os.path.join(users_path, d)) and not d.startswith('.') and d.lower() != "all users"]
        except Exception:
            users = []
        self.user_combo['values'] = users
        if users:
            self.user_combo.current(0)

    def _update_header_time(self):
        """Update header time every second"""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.header_time_label.config(text=current_time)
            self.after(1000, self._update_header_time)
        except Exception:
            pass  # window closing: the clock just stops (as before)

    def format_bytes(self, size):
        try:
            size = float(size)
        except Exception:
            return str(size)
        for unit in ["B", "KB", "MB", "GB", "TB", "PB"]:
            if size < 1024.0 or unit == "PB":
                return f"{size:.2f} {unit}"
            size /= 1024.0

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
                    # not saved: the new filters apply to this session only (as before)
                    _log.debug("event=settings_save_failed source=filters", exc_info=True)
                self.message_queue.put(('log', 'Filters updated.'))
        except Exception as e:
            try:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('filter_manager_error').format(e),
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                # the error dialog itself failed: nothing more to show (as before)
                _log.debug("event=filter_error_dialog_failed", exc_info=True)


def run(entry_script):
    """Create the App window (entry_script: path the scheduled task runs) and run mainloop."""
    diag.setup()  # diagnostic log (ADR 0003 §2); each job binds its own run_id
    app = App(entry_script=entry_script)
    app.mainloop()
