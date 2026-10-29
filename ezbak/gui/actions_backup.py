"""User data backup actions of App (moved from ezBAK.py in P5-10).

BackupActionsMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before. Module-level names (os, shutil, threading, traceback, time,
datetime, core_engine_run_backup, core_is_skipped, core_copy_file, winsys_free_space,
...) match ezBAK.py so facade patches reach this module too, and the cli harness
finds the engine alias `core_engine_run_backup` by identity. C:\\Users is read from
ezbak.winsys.paths.USERS_ROOT at call time (P5 decision D3).
"""
import dataclasses
import logging
import os
import shutil
import threading
import time
import traceback
from datetime import datetime
from tkinter import filedialog
from types import SimpleNamespace

from ezbak.core import BackupJob, BackupPolicy, InsufficientSpaceError, JobLog, JobLogFormat
from ezbak.core import history as core_history
from ezbak.core import run_backup as core_engine_run_backup
from ezbak.core import compute_total_size, copy_file_with_progress as core_copy_file
from ezbak.core import is_skipped as core_is_skipped
from ezbak.targets import default_targets
from ezbak.winsys import free_space as winsys_free_space
from ezbak.winsys import paths as winsys_paths

from .bridge import GuiFolderStore, GuiJob, GuiLogStream, GuiReporter, begin_job
from .dialogs import Win11Dialog

_log = logging.getLogger(__name__)


class BackupActionsMixin:
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
        # Judgment lives in ezbak.core.filters (policy: GUI + AppData names + ctypes fallback).
        try:
            include_hidden = self.hidden_mode_var.get() == 'include'
            include_system = self.system_mode_var.get() == 'include'
        except Exception:
            include_hidden = include_system = False
        flags = SimpleNamespace(include_hidden=include_hidden, include_system=include_system,
                                filters=getattr(self, 'filters', None))
        policy = dataclasses.replace(BackupPolicy.GUI, skip_appdata_names=True,
                                     attr_ctypes_fallback=True)
        return core_is_skipped(filepath, flags, policy) is not None

    # Helper: copy single file in chunks and report bytes progress
    def copy_file_with_progress(self, src, dst, progress_callback, buffer_size=64*1024):
        """Copies a file in chunks and calls progress_callback(bytes_copied) incrementally."""
        return core_copy_file(src, dst, progress_callback, buffer_size=buffer_size,
                              on_error=lambda e: self.write_detailed_log(f"Error copying file {src} -> {dst}: {e}"))

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
        """GUI backup entry (ADR 0002 D2): checks the user and prepares the backup folder,
        then runs ezbak.core.run_backup with the GUI policy. The job log is the one
        start_backup_thread opened (open_log_file); it is closed in _finish_backup."""
        begin_job("backup", job_log=getattr(self, 'log_file_path', None))
        fmt, put = JobLogFormat.GUI, self.message_queue.put
        backup_success = False
        try:
            user_name = self.user_var.get()
            user_profile_path = os.path.join(winsys_paths.USERS_ROOT, user_name)

            if not os.path.exists(user_profile_path):
                put(('show_error', f"User folder '{user_profile_path}' not found."))
                return

            self.write_detailed_log(fmt.start.line(
                source=user_profile_path, backup_path=backup_path).message)
            put(('log', fmt.rp_user.line(user=user_name).message))
            destination_dir = self._prepare_backup_folder(backup_path, user_name)

            if self._run_backup_engine(backup_path, user_name, destination_dir) is None:
                return  # not enough space or CRITICAL error, already reported

            put(('log', fmt.rp_done_log.line().message))
            put(('update_status', fmt.rp_done_status.line().message))
            put(('open_folder', destination_dir))
            backup_success = True
            core_history.record_backup(user=user_name, dest=backup_path, folder=destination_dir,
                                       when=datetime.now(), mode='gui')  # home card (U7-2)

        except Exception as e:
            # Entry boundary: any other error ends the backup with an error dialog.
            self.write_detailed_log(fmt.backup_error.line(error=e).message)
            self.write_detailed_log(fmt.backup_traceback.line(tb=traceback.format_exc()).message)
            put(('show_error', fmt.rp_error.line(error=e).message))
            backup_success = False

        finally:   # Finally block for entire function
            self._finish_backup(backup_success)

    def _prepare_backup_folder(self, backup_path, user_name):
        """New `<user>_backup_<ts>` folder under backup_path; a same-name folder is removed first."""
        fmt = JobLogFormat.GUI
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        destination_dir = os.path.join(backup_path, f"{user_name}_backup_{timestamp}")

        if os.path.exists(destination_dir):
            try:
                shutil.rmtree(destination_dir)
                self.write_detailed_log(fmt.existing_deleted.line(path=destination_dir).message)
            except Exception as e:
                # Work-unit boundary: the backup goes on into the existing folder (as before).
                self.write_detailed_log(fmt.existing_delete_failed.line(error=e).message)

        os.makedirs(destination_dir, exist_ok=True)
        self.write_detailed_log(fmt.folder_created_new.line(path=destination_dir).message)
        return destination_dir

    def _run_backup_engine(self, backup_path, user_name, destination_dir):
        """ezbak.core.run_backup (GUI policy) into destination_dir. Returns the result, or
        None after show_error (not enough space, CRITICAL error of the copy)."""
        fmt = JobLogFormat.GUI
        clock = lambda: datetime.now()  # module datetime, read at call time (tests patch it)
        # Retention values are read by GuiJob when the cleanup step asks (old timing).
        job = GuiJob(self, BackupJob(user=user_name, dest=backup_path,
                                     include_hidden=self.hidden_mode_var.get() != 'exclude',
                                     include_system=self.system_mode_var.get() != 'exclude',
                                     filters=getattr(self, 'filters', None)))
        try:
            result = core_engine_run_backup(
                job, destination_dir, JobLog(GuiLogStream(self), clock=clock), fmt,
                GuiFolderStore(self, backup_path, clock), default_targets(), GuiReporter(self),
                BackupPolicy.GUI, clock, users_root=winsys_paths.USERS_ROOT,
                log_path=getattr(self, 'log_file_path', None) or "", now_ts=lambda: time.time())
        except InsufficientSpaceError as e:
            self.message_queue.put(('show_error', fmt.rp_space_short.line(
                required=self.format_bytes(e.required), free=self.format_bytes(e.free)).message))
            return None
        if result.critical_error is not None:
            self.message_queue.put(('show_error', fmt.rp_critical.line(
                error=result.critical_error).message))
            return None
        return result

    def _finish_backup(self, backup_success):
        """End of every GUI backup (old run_backup finally): close the log, fill the bar,
        enable buttons, schedule the reset, final status and sound."""
        try:
            self.close_log_file()
        except Exception:
            # closing the job log is best effort; the backup result stands
            _log.debug("event=job_log_close failed", exc_info=True)

        try:
            max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
            self.message_queue.put(('update_progress', max_val))
        except Exception:
            pass  # progress bar decoration only

        # Enable buttons
        self.message_queue.put(('enable_buttons', None))

        # Additional safeguard: also try direct button activation
        try:
            self.after(1000, lambda: self.set_buttons_state("normal"))
        except Exception:
            pass  # widget already gone; enable_buttons above is the main path

        # Reset progress bar after 2 seconds
        try:
            self.after(2000, lambda: self.message_queue.put(('reset_progress', None)))
        except Exception:
            pass  # widget already gone; the reset is decoration only

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
        return core_copy_file(
            src, dst, progress_callback, buffer_size=buffer_size, timeout_seconds=timeout_seconds,
            on_timeout=lambda s: self.write_detailed_log(f"TIMEOUT: File copy timeout for {s}"),
            on_error=lambda e: self.write_detailed_log(f"ERROR: copying file {src} -> {dst}: {e}"))

    def compute_total_bytes_safe(self, root_path):
        """
        Safe total size calculation
        """
        total, _count = compute_total_size(
            root_path, lambda p, _is_dir: self.is_hidden(p),  # same check as copy (fix h-m, D4)
            on_progress=lambda n, t: self.message_queue.put(
                ('update_status', f"Calculating total size... {n} files, {self.format_bytes(t)}")),
            on_walk_error=lambda e: self.write_detailed_log(f"Walk error: {e}"),
            on_file_error=lambda fp, e: self.write_detailed_log(f"Size calculation error for {fp}: {e}"),
            on_error=lambda e: self.write_detailed_log(f"Total bytes calculation error: {e}"))
        return total

    def is_hidden_safe(self, filepath, timeout_seconds=2):
        """
        is_hidden function with timeout - includes filters
        """
        # Judgment lives in ezbak.core.filters (policy: GUI). timeout_seconds is unused, as before.
        try:
            flags = SimpleNamespace(include_hidden=self.hidden_mode_var.get() != 'exclude',
                                    include_system=self.system_mode_var.get() != 'exclude',
                                    filters=getattr(self, 'filters', None))
            return core_is_skipped(filepath, flags, BackupPolicy.GUI) is not None
        except Exception:
            return False

    def get_free_space_reliable(self, path):
        """
        More reliable free space calculation (delegates to ezbak.winsys.sysinfo.free_space)
        """
        return winsys_free_space(path)
