"""Schedule backup dialog (moved from ezBAK.py in P5-7).

Tests patch `Win11Dialog` and `filedialog.askdirectory`; the calls below go through
the module-level names so patched objects are used here too. The widget layout lives
in schedule_layout.py. The App-side actions (App.schedule_backup and the task
create/delete calls) stay in ezBAK.py until P5-12. copy_default_filters is the
single helper from ..settings, bound here under the same name (P5-8a).
Colours and fonts come from the design tokens for the theme's mode (U6-4, ADR 0004);
widgets, texts, keys and the task data are as before.
"""
import logging
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

from ezbak.winsys.paths import resource_path

from .. import dpi
from ..settings import copy_default_filters
from ..shortcuts import DialogShortcuts
from ..theme import token_colors
from ..widgets import create_win11_button
from .filters import FilterManagerDialog
from .message import Win11Dialog
from .schedule_layout import ScheduleLayoutMixin

_log = logging.getLogger(__name__)


class ScheduleBackupDialog(ScheduleLayoutMixin, tk.Toplevel, DialogShortcuts):
    def __init__(self, parent):
        _log.debug("ScheduleBackupDialog.__init__ started")

        # Important initialization order: set basic properties first
        self.action = None
        self.result = None
        self.task_name = None
        self.parent = parent
        self.theme = parent.theme  # Use parent's theme
        self._colors = token_colors(self.theme)  # design tokens for the theme's mode (U6-4)
        self.translator = parent.translator  # Use parent's translator
        # Initialize filters with default exclude patterns
        self.filters = copy_default_filters()

        try:
            tk.Toplevel.__init__(self, parent)  # Explicitly initialize Toplevel
            _log.debug("Toplevel initialized")

            self._setup_schedule_window(parent)

            # Create UI components
            _log.debug("Creating UI widgets...")
            self._create_widgets()
            _log.debug("UI widgets created successfully")

            # Setup shortcuts (after UI creation)
            self._setup_schedule_shortcuts()
            _log.debug("Shortcuts setup completed")

        except Exception as init_error:
            _log.debug(f"ScheduleBackupDialog init failed: {init_error}")
            self.action = 'error'
            try:
                self.destroy()
            except Exception:
                # the init error below is what matters; a failed destroy adds nothing
                _log.debug("event=schedule_init_destroy failed", exc_info=True)
            raise init_error

    def _setup_schedule_window(self, parent):
        """Title, close protocol, icon, size/position and modality (part of __init__)."""
        self.title(parent.translator.get('schedule_backup'))
        self.configure(bg=self._colors['bg'])

        # Handle window close safely - set first
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        _log.debug("WM_DELETE_WINDOW protocol set")

        # Set icon (continue even if it fails)
        try:
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception as icon_error:
            _log.debug(f"Icon setting failed: {icon_error}")
            pass

        # Set window size and position
        w, h = 546, 380
        self.geometry(dpi.geometry(self, w, h))  # px at 100 % (U8)

        try:
            self.update_idletasks()
            parent_x = parent.winfo_rootx()
            parent_y = parent.winfo_rooty()
            x = parent_x + 50
            y = parent_y + 50
            self.geometry(dpi.geometry(self, w, h, x, y))
            _log.debug(f"Window positioned at {x},{y}")
        except Exception as pos_error:
            _log.debug(f"Window positioning failed: {pos_error}")
            pass

        try:
            self.transient(parent)
            _log.debug("Transient set")
            self.grab_set()
            _log.debug("Grab set")
        except Exception as modal_error:
            _log.debug(f"Modal setting failed: {modal_error}")
            pass

    def _create_win11_button(self, parent, text, command, is_primary=False, width=12):
        """Create a Windows 11 style button for dialogs"""
        return create_win11_button(self.theme, parent, text, command, is_primary, width)

    def _parent_user_name(self):
        """Selected user of the parent App, "Unknown" when unset or unreadable."""
        user_name = "Unknown"
        try:
            if hasattr(self.parent, 'user_var') and self.parent.user_var.get():
                user_name = self.parent.user_var.get()
        except Exception:
            # a missing or broken user_var falls back to "Unknown"
            _log.debug("event=schedule_parent_user_read failed", exc_info=True)
        return user_name

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
        self.bind('<Alt-Key-1>', lambda e: self._set_schedule("Daily"))
        self.bind('<Alt-Key-2>', lambda e: self._set_schedule("Weekly"))
        self.bind('<Alt-Key-3>', lambda e: self._set_schedule("Monthly"))

    def _set_schedule(self, schedule_type):
        """Set schedule type with shortcut key"""
        try:
            self.schedule_var.set(schedule_type)
            _log.debug(f"Schedule set to {schedule_type}")
        except Exception as e:
            _log.debug(f"Failed to set schedule: {e}")

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
        help_dlg.configure(bg=self._colors['bg'])
        help_dlg.geometry(dpi.geometry(help_dlg, 400, 350))  # px at 100 % (U8)
        help_dlg.transient(self)
        try:
            help_dlg.grab_set()
        except Exception:
            pass  # grab can fail while another window holds it; the help still shows modeless

        text_widget = tk.Text(help_dlg,
                             font=("Consolas", 10),
                             bg=self._colors['card'],
                             fg=self._colors['fg'],
                             relief="flat",
                             wrap="word",
                             padx=12, pady=12)
        text_widget.pack(fill="both", expand=True, padx=15, pady=15)

        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        btn_frame = tk.Frame(help_dlg, bg=self._colors['bg'])
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
                _log.debug(f"Destination set to: {folder}")
        except Exception as browse_error:
            _log.debug(f"Browse failed: {browse_error}")

    def open_filter_manager(self):
        """Open filter manager dialog"""
        try:
            dlg = FilterManagerDialog(self, current_filters=self.filters)
            self.wait_window(dlg)
            if getattr(dlg, 'result', None):
                self.filters = dlg.result
                _log.debug(f"Filters updated: {self.filters}")
        except Exception as e:
            _log.debug(f"Filter manager error: {e}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('filter_manager_error').format(e),
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                # the error was already printed above; a failed error box adds nothing
                _log.debug("event=filter_manager_error_dialog failed", exc_info=True)

    def _schedule_time_str(self):
        """Hour and minute as "HH:MM", or None after showing the invalid-time error."""
        # Combine hour and minute into time string
        hour = self.hour_var.get()
        minute = self.minute_var.get()
        time_str = f"{hour}:{minute}"

        # Validate time format
        try:
            time_obj = datetime.strptime(time_str, "%H:%M")
            time_str = time_obj.strftime("%H:%M")
        except ValueError:
            Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_time_format'),
                                parent=self, theme=self.theme, translator=self.translator)
            return None
        return time_str

    def on_create(self):
        """Create button handler"""
        _log.debug("event=schedule_create_start")
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
                    f"{self.translator.get('dest_folder_missing')}:\n{dest}\n\n{self.translator.get('continue_anyway')}",
                    parent=self, theme=self.theme, translator=self.translator):
                    return

            time_str = self._schedule_time_str()
            if time_str is None:
                return

            # Get user name
            user_name = self._parent_user_name()

            # Build the result first so a bad number leaves action/result unset (h-i)
            result = {
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
            self.action = "create"
            self.result = result
            _log.debug("event=schedule_create_result result=%s", self.result)
            self._safe_destroy()

        except Exception as create_error:
            _log.debug("event=schedule_create_failed error=%s", create_error, exc_info=True)
            try:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('failed_to_create_schedule').format(create_error),
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                _log.error("event=schedule_create_error_box_failed error=%s", create_error, exc_info=True)

    def on_delete(self):
        """Delete button handler"""
        _log.debug("on_delete called")
        try:
            task_name = self.task_var.get().strip()
            if not task_name:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('task_name_required_delete'),
                                    parent=self, theme=self.theme, translator=self.translator)
                return

            self.action = "delete"
            self.task_name = task_name
            _log.debug(f"Delete task_name set: {task_name}")
            self._safe_destroy()
        except Exception as delete_error:
            _log.debug(f"on_delete failed: {delete_error}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), f"{self.translator.get('failed_to_delete_schedule')}: {delete_error}",
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                _log.warning(f"{delete_error}")

    def on_close(self):
        """Close button handler"""
        _log.debug("on_close called")
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
        _log.debug("_safe_destroy called")
        try:
            self.grab_release()
            _log.debug("grab_release successful")
        except Exception as grab_error:
            _log.debug(f"grab_release failed: {grab_error}")

        try:
            self.destroy()
            _log.debug("destroy successful")
        except Exception as destroy_error:
            _log.debug(f"destroy failed: {destroy_error}")
