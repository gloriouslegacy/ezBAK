"""User data restore actions of App (moved from ezBAK.py in P5-13).

RestoreActionsMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before. Module-level names (os, threading, filedialog, Win11Dialog,
GuiReporter, KIND_DIRECT, RestoreSelectionError, core_locate_backup, core_run_restore,
core_validate_restore_selection, targets_restore_bookmarks, winsys_paths) match
ezBAK.py so facade patches reach this module too. run_restore is split into two
module functions that run the same statements in the same order (P5 decision D5);
module functions, not App methods, so gui_harness.recording never stubs them. The
locals shared by the try body and the finally block live in a SimpleNamespace.
"""
import logging
import os
import threading
from tkinter import filedialog
from types import SimpleNamespace

from ezbak.core import (KIND_DIRECT, RestoreSelectionError, locate_backup as core_locate_backup,
                        run_restore as core_run_restore,
                        validate_selection as core_validate_restore_selection)
from ezbak.targets import restore_bookmarks as targets_restore_bookmarks
from ezbak.winsys import paths as winsys_paths  # USERS_ROOT read at call time (P5 D3)

from .bridge import GuiReporter, begin_job
from .dialogs import Win11Dialog

_log = logging.getLogger(__name__)


def _restore_steps(self, backup_folder, state):
    # Module function, not an App method: the try body of App.run_restore (P5 D5).
    # `state` carries restore_success / log_opened / log_path / discard_log.
    user_name = self.user_var.get()

    try:
        core_validate_restore_selection(backup_folder)
        self.message_queue.put(('log', f"Selected backup folder: {backup_folder}"))
        plan = core_locate_backup(backup_folder, user_name, self._wait_restore_folder_choice)
    except RestoreSelectionError as e:
        self.message_queue.put(('show_error', e.message))
        return
    if plan is None:
        self.message_queue.put(('log', self.translator.get('data_restore_cancelled')))
        return

    folder_selected = plan.folder_name
    backup_folder = plan.backup_root
    # Open log file in the folder holding the backup, named after the backup folder
    self.open_log_file(backup_folder, f"{folder_selected}_restore")
    state.log_opened = True
    state.log_path = getattr(self, 'log_file_path', None)
    _log.info("event=job_log kind=restore path=%s", state.log_path)
    if plan.kind == KIND_DIRECT:
        self.write_detailed_log(f"Direct backup folder selection detected: {folder_selected}")
        self.write_detailed_log(f"Parent folder: {backup_folder}")
    else:
        self.write_detailed_log(f"User selected backup: {folder_selected}")

    result = core_run_restore(
        plan, user_name, os.path.join(winsys_paths.USERS_ROOT, user_name), self.write_detailed_log,
        GuiReporter(self), skip=lambda path, is_dir=None: self.is_hidden(path),
        free_space=self.get_free_space_reliable, confirm_cross_user=self._wait_cross_user_confirm)
    state.discard_log = bool(result.cancelled or result.error)
    if result.cancelled:
        self.message_queue.put(('log', self.translator.get('data_restore_cancelled')))
        return
    if result.error:
        self.message_queue.put(('show_error', result.error))
        return

    targets_restore_bookmarks(backup_folder, self.write_detailed_log, os.path.join(winsys_paths.USERS_ROOT, user_name))

    try:
        retention_days = int(self.log_retention_days_var.get())
    except (ValueError, TypeError):
        retention_days = 30
    self._cleanup_old_logs(backup_folder, retention_days)

    self.message_queue.put(('log', "Restore complete (detailed log saved)."))
    self.message_queue.put(('update_status', "Restore complete!"))
    state.restore_success = True


def _finish_restore(self, state):
    # Module function, not an App method: the finally block of App.run_restore (P5 D5).
    if state.log_opened:
        try:
            self.close_log_file()
        except Exception:
            # closing the job log is best effort (as before); the restore result stands
            _log.debug("event=job_log_close failed", exc_info=True)
    if state.discard_log and state.log_path:
        try:
            os.remove(state.log_path)
        except OSError as e:
            _log.warning("restore_log_discard failed path=%s error=%s", state.log_path, type(e).__name__)

    try:
        max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
        self.message_queue.put(('update_progress', max_val))
    except Exception:
        pass  # final progress fill is cosmetic (as before)

    # Reset progress bar
    self.message_queue.put(('reset_progress', None))

    self.message_queue.put(('enable_buttons', None))

    if state.restore_success:
        self.message_queue.put(('update_status', "Restore complete!"))
        self.message_queue.put(('play_sound', 'complete'))
    else:
        self.message_queue.put(('update_status', "Operation finished"))
        self.message_queue.put(('play_sound', 'error'))

    if hasattr(self, '_restore_completed') and self._restore_completed:
        self.message_queue.put(('update_status', "Restore complete!"))
    else:
        self.message_queue.put(('update_status', "Operation finished"))


class RestoreActionsMixin:
    """start_restore_thread and run_restore of App."""

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

        self.message_queue.put(('log', "User data restore process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['mode'] = "determinate"

        restore_thread = threading.Thread(target=self.run_restore, args=(backup_folder,), daemon=True)
        restore_thread.start()

    def run_restore(self, backup_folder):
        """Executes the actual restore logic (byte-based progress)."""
        begin_job("restore")
        state = SimpleNamespace(
            restore_success=False,
            log_opened=False,  # Track if log file is opened
            log_path=None,  # close_log_file clears log_file_path, so keep it here
            discard_log=False,  # aborted before copying: the job log is removed
        )

        try:
            _restore_steps(self, backup_folder, state)

        except Exception as e:
            self.write_detailed_log(f"Restore error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during restore: {e}"))
            state.restore_success = False

        finally:
            _finish_restore(self, state)
