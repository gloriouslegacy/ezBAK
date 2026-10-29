"""GUI side of the core ports for App.run_backup (P2-7), App's queue (P5-8b) and the
restore question callbacks (P5-13).

The work thread never touches widgets: everything goes into App.message_queue
with the task strings of ADR 0003 section 6 (C3). No tkinter import here.
QueueMixin (process_queue, moved from ezBAK.py) drains that queue on the Tk thread;
module-level names (os, queue, time) match ezBAK.py so facade patches reach it.
"""
from __future__ import annotations

import datetime as _dt
import logging
import os
import queue
import time
from typing import Any, Callable

from ezbak.core import diag
from ezbak.storage import LocalFolderStore

from .dialogs import Win11Dialog

_log = logging.getLogger(__name__)


def begin_job(kind: str, **fields: Any) -> str:
    """First call of a GUI work thread (P6-3, ADR 0003 §2): bind a new run_id in this
    thread (threads do not inherit it) and write `event=job_start kind=<kind> k=v ...`
    to the diagnostic log. Returns the run_id. Nothing goes to the job log or the queue."""
    run_id = diag.bind(diag.new_run_id(_dt.datetime.now))
    _log.info("event=job_start kind=%s%s", kind, "".join(f" {k}={v}" for k, v in fields.items()))
    return run_id


def int_or(value: Any, default: int) -> int:
    """int(value), or `default` when it is not a number (GUI retention vars)."""
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


class GuiJob:
    """Engine job for App.run_backup: fields of a BackupJob, but the retention values
    are read from the App vars when the cleanup step asks, as the old cleanup did.
    A non-number falls back to 2 / 30; any other read error propagates to the engine
    (no count: backup cleanup skipped; no days: "Cleanup failed")."""

    def __init__(self, app: Any, job: Any) -> None:
        self._app = app
        self._job = job

    def __getattr__(self, name: str) -> Any:
        return getattr(self._job, name)

    @property
    def retention_count(self) -> int:
        return int_or(self._app.retention_count_var.get(), 2)

    @property
    def log_retention_days(self) -> int:
        return int_or(self._app.log_retention_days_var.get(), 30)


class GuiLogStream:
    """Text stream for core.JobLog over App's detailed log file: same lock and flush per
    line as App.write_detailed_log, nothing written when no log file is open."""

    def __init__(self, app: Any) -> None:
        self._app = app

    def write(self, text: str) -> None:
        app = self._app
        with app.log_lock:
            if not app._log_file:
                return
            try:
                app._log_file.write(text)
                app._log_file.flush()
            except Exception as e:
                # I/O boundary, as write_detailed_log: a failed log write never stops the backup.
                _log.debug("event=job_log_write_failed error=%s", type(e).__name__)

    def flush(self) -> None:
        """Every write is flushed already."""


class GuiReporter:
    """core.Reporter over App.message_queue ('log', 'update_status',
    'update_progress_max', 'update_progress').

    progress is throttled with App._last_progress_ts / progress_throttle_secs, as
    the old copy callback was. Copies only raise the reported total, so a value
    that does not grow is the engine's final report after the walk; it is always
    sent and does not move the throttle clock (old run_backup did the same).
    """

    def __init__(self, app: Any, now: Callable[[], float] = time.time) -> None:
        self._app = app
        self._now = now
        self._last_value = 0

    def log(self, message: str) -> None:
        self._app.message_queue.put(('log', message))

    def status(self, message: str) -> None:
        self._app.message_queue.put(('update_status', message))

    def progress_max(self, value: int) -> None:
        self._app.message_queue.put(('update_progress_max', value))

    def progress(self, value: int) -> None:
        app = self._app
        final = value <= self._last_value
        self._last_value = value
        if not final:
            now = self._now()
            if now - app._last_progress_ts < app.progress_throttle_secs:
                return
            app._last_progress_ts = now
        app.message_queue.put(('update_progress', value))


class GuiFolderStore(LocalFolderStore):
    """LocalFolderStore whose free space still comes from App.get_free_space_reliable."""

    def __init__(self, app: Any, dest: str, clock: Callable[[], Any]) -> None:
        super().__init__(dest, clock=clock)
        self._app = app

    def free_space(self) -> int:
        return self._app.get_free_space_reliable(self.dest)


class QueueMixin:
    """App.process_queue: the 13 message kinds of ADR 0003 section 6, same order as before,
    plus 'ask_yes_no' (g-e).

    Also the two restore callbacks (P5-13) and _wait_yes_no (g-e): queue the question,
    then poll the App attribute every 0.1 s (P5 decision D4; Event only once cancelling is required)."""

    def process_queue(self):
        """Processes messages from the message queue to update the GUI."""
        try:
            while True:
                task, value = self.message_queue.get_nowait()
                self._handle_queue_message(task, value)
                self.message_queue.task_done()
        except queue.Empty:
            pass  # queue drained; the next poll is scheduled below
        finally:
            self.after(100, self.process_queue)

    def _handle_queue_message(self, task, value):
        """One message of process_queue (split out for the function size limit, D5)."""
        if task == 'log':
            now = time.time()
            if now - self._last_ui_log_ts >= self.ui_log_throttle_secs:
                self._last_ui_log_ts = now
                self.log(value)
        elif task == 'update_progress':
            self._queue_update_progress(value)
        elif task == 'space_check_result':
            self.show_space_check_result(value)
        elif task == 'update_progress_max':
            try:
                self.progress_bar['maximum'] = value
                self.status_label.config(text="Progress: 0%")
            except Exception:
                pass  # progress widget update is cosmetic (as before)
        elif task == 'update_status':
            self.status_label.config(text=value)
        elif task == 'enable_buttons':
            try:
                self.set_buttons_state("normal")
                _log.debug("Buttons enabled")
            except Exception as e:
                _log.debug(f"Button enable failed: {e}")
        elif task == 'show_error':
            Win11Dialog.showerror(self.translator.get('error'), value,
                                parent=self, theme=self.theme, translator=self.translator)
        elif task == 'open_folder':
            try:
                os.startfile(value)
            except Exception:
                # opening the result folder is a convenience (as before)
                _log.debug("event=open_folder_failed", exc_info=True)
        elif task == 'stop_progress':
            try:
                self.progress_bar.stop()  # indeterminate 모드 정지
                self.progress_bar['mode'] = "determinate"
            except Exception:
                pass  # progress widget update is cosmetic (as before)
        else:
            self._handle_queue_dialog_message(task, value)

    def _handle_queue_dialog_message(self, task, value):
        """Rest of _handle_queue_message: sound, reset, the two restore dialogs, yes/no."""
        if task == 'play_sound':
            # Sound playback in main thread to avoid issues
            try:
                self._play_sound(value)
            except Exception as e:
                _log.debug(f"Sound play failed: {e}")
        elif task == 'reset_progress':
            # Reset progress bar to initial state
            try:
                self.reset_progress_bar()
            except Exception as e:
                _log.debug(f"Progress bar reset failed: {e}")
        elif task == 'show_cross_user_warning':
            # Handle cross-user restore warning
            backup_user, target_user, folder_name = value
            self.show_cross_user_warning_dialog(backup_user, target_user, folder_name)
        elif task == 'select_folder':
            # Handle backup folder selection
            backup_folders = value
            self.show_folder_selection_dialog(backup_folders)
        elif task == 'ask_yes_no':
            # Yes/no question from a worker thread (driver restore mismatch, g-e)
            self._show_yes_no_dialog(*value)

    def _queue_update_progress(self, value):
        """'update_progress': bar value and "Progress: N%" status."""
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
            pass  # progress widget update is cosmetic (as before)

    def _wait_restore_folder_choice(self, folders):
        """choose_backup callback for core.locate_backup: queue the dialog, poll the answer."""
        self.message_queue.put(('select_folder', folders))
        return self._poll_answer('_folder_selection_result')

    def _wait_cross_user_confirm(self, backup_user, user_name, folder_selected):
        """confirm_cross_user callback for core.run_restore: queue the warning, poll the answer."""
        self.message_queue.put(('show_cross_user_warning', (backup_user, user_name, folder_selected)))
        return self._poll_answer('_cross_user_confirmed')

    def _wait_yes_no(self, title, message):
        """Worker-side Win11Dialog.askyesno (g-e): the main thread shows it ('ask_yes_no')."""
        self.message_queue.put(('ask_yes_no', (title, message)))
        return self._poll_answer('_yes_no_answer')

    def _show_yes_no_dialog(self, title, message):
        """'ask_yes_no' on the Tk thread: the dialog, answer stored for _wait_yes_no."""
        self._yes_no_answer = Win11Dialog.askyesno(title, message, parent=self,
                                                   theme=self.theme, translator=self.translator)

    def _poll_answer(self, attr):
        """Wait for the Tk thread to set App.<attr> (every 0.1 s, D4); take and clear it."""
        while True:
            if hasattr(self, attr):
                answer = getattr(self, attr)
                delattr(self, attr)
                return answer
            time.sleep(0.1)
