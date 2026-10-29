"""UI log, job log file and sound of App (moved from ezBAK.py in P5-8b).

AppLogMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before (log_text, log_lock, _log_file, log_file_path, message_queue,
sound_enabled_var, retention vars). Module-level names (datetime, os, threading,
time, tk, filedialog) match ezBAK.py so facade patches reach this module too.
Job log file names, header lines and cleanup wording are C3 and unchanged.
"""
import logging
import os
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

from ezbak.core import BackupPolicy, JobLog, JobLogFormat, cleanup_old_logs, header_lines

from .dialogs import Win11Dialog

_log = logging.getLogger(__name__)


class AppLogMixin:
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
            _log.debug(f"Sound toggle error: {e}")

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
            _log.debug(f"Sound playback error: {e}")

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
                # a failed log write must not stop the job (as before)
                _log.debug("event=job_log_write failed", exc_info=True)

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
                    except Exception:
                        backup_keep = '2'
                    try:
                        log_keep = self.log_retention_days_var.get()
                    except Exception:
                        log_keep = '30'

                    # Only write filters for backup operations (not for restore)
                    filters = None
                    if '_restore' not in prefix.lower():
                        filters = getattr(self, 'filters', {'include': [], 'exclude': []}) or {'include': [], 'exclude': []}

                    # Header: Policy, Filters, separator with one timestamp (ezbak.core.joblog)
                    job_log = JobLog(self._log_file, clock=lambda: datetime.now())
                    job_log.write_lines(header_lines(
                        JobLogFormat.GUI,
                        hidden=self.hidden_mode_var.get(), system=self.system_mode_var.get(),
                        backup_keep=backup_keep, log_keep=log_keep, filters=filters),
                        single_timestamp=True)
                    job_log.flush()
                except Exception:
                    # header is best effort; the log file stays open (as before)
                    _log.debug("event=job_log_header failed", exc_info=True)
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
                # closing is best effort; the handle is dropped below (as before)
                _log.debug("event=job_log_close failed", exc_info=True)
            self._log_file = None
            self.log_file_path = None

    def _cleanup_old_logs(self, log_folder, retention_days):
        """Old .log cleanup (also used by restore); body in ezbak.core.retention."""
        try:
            retention_days = int(retention_days)
        except (ValueError, TypeError):
            retention_days = 30
        # "Current time" read datetime.datetime through a local import (fixed_clock patches it).
        from datetime import datetime as _datetime
        cleanup_old_logs(log_folder, retention_days, JobLogFormat.GUI, BackupPolicy.GUI,
                         lambda line: self.write_detailed_log(line.message),
                         clock=lambda: _datetime.now(), now_ts=lambda: time.time())

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
