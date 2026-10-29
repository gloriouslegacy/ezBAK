"""Copy Data and Check Space actions of App (moved from ezBAK.py in P5-11).

CopyActionsMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before. Module-level names (os, threading, time, datetime, filedialog,
SelectSourcesDialog, ...) match ezBAK.py so facade patches reach this module too;
free space comes through self.get_free_space_reliable. Long methods are split into helpers that run the same statements in
the same order (P5 decision D5).
"""
import logging
import os
import threading
import time
from datetime import datetime
from tkinter import filedialog

from ezbak.core import compute_total_size, copy_sources

from .bridge import begin_job
from .dialogs import SelectSourcesDialog, Win11Dialog

_log = logging.getLogger(__name__)


def _source_bytes(app, path):
    """Bytes copy will count for one source: a file by getsize (hidden ones skipped, like
    run_space_calculation), a folder by compute_total_bytes_safe (fix h-l)."""
    if os.path.isfile(path):
        if app.is_hidden(path):
            return 0
        try:
            return os.path.getsize(path)
        except OSError as e:
            app.write_detailed_log(f"Size calculation error for {path}: {e}")
            return 0
    if os.path.isdir(path):
        return app.compute_total_bytes_safe(path)
    return 0


class CopyActionsMixin:
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

    def _copy_select_sources(self):
        """Step 1 of copy_files: the checkbox tree dialog. None when it failed."""
        try:
            _log.debug("event=copy_source_dialog_create")
            dlg = SelectSourcesDialog(self, is_hidden_fn=self.is_hidden, title=self.translator.get('select_sources'))
            _log.debug("event=copy_source_dialog_wait")
            self.wait_window(dlg)
            _log.debug("event=copy_source_dialog_closed")

            source_paths_list = dlg.selected_paths if getattr(dlg, 'selected_paths', None) else []
            _log.debug("event=copy_sources_selected count=%d", len(source_paths_list))

        except Exception as e:
            _log.debug("event=copy_source_selection_failed error=%s", e, exc_info=True)
            self.message_queue.put(('log', f"Source selection failed: {e}"))
            return None
        return source_paths_list

    def copy_files(self):
        """Prompts for source and destination using a checkbox tree, then starts the copy process."""
        try:
            _log.debug("event=copy_files_start")
            self.message_queue.put(('log', "Copy process initiated."))

            # 1) Open a checkbox tree dialog to select files/folders
            source_paths_list = self._copy_select_sources()
            if source_paths_list is None:
                return

            if not source_paths_list:
                _log.debug("event=copy_no_sources")
                self.message_queue.put(('log', "Source selection cancelled."))
                return

            _log.debug("event=copy_sources sources=%s", source_paths_list)

            # 2) Ask for destination folder
            _log.debug("event=copy_ask_destination")
            destination_dir = filedialog.askdirectory(title=self.translator.get('select_destination_folder'))
            _log.debug("event=copy_destination dest=%s", destination_dir)

            if not destination_dir:
                _log.debug("event=copy_no_destination")
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
            _log.debug("event=copy_files_failed error=%s", e, exc_info=True)
            self.message_queue.put(('show_error', f"Copy process failed: {e}"))

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
            Win11Dialog.showerror(self.translator.get('error'), self.translator.get('source_selection_failed').format(e),
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
        begin_job("space_check")
        try:
            total = 0
            file_count = 0

            self.message_queue.put(('update_status', "Calculating required space..."))

            for i, source_path in enumerate(sources):
                self.message_queue.put(('update_status', f"Checking {i+1}/{len(sources)}: {os.path.basename(source_path)}"))

                if os.path.isdir(source_path):
                    # Calculate directory size (App.is_hidden policy; check errors propagate)
                    dir_total, dir_files = compute_total_size(
                        source_path, lambda p, _is_dir: self.is_hidden(p), strict=True,
                        on_progress=lambda n, t, sp=source_path: self.message_queue.put(
                            ('update_status', f"Checking {os.path.basename(sp)}: {n} files, {self.format_bytes(t)}")))

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

    def _copy_space_ok(self, source_paths, destination_dir):
        """Total size and free space step of run_copy_thread. False when space is short."""
        # --- Calculate total size ---
        self.message_queue.put(('update_status', "Calculating total size..."))
        total_bytes = 0
        for path in source_paths:
            total_bytes += _source_bytes(self, path)

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
            return False
        return True

    def _copy_to_new_folder(self, source_paths, destination_dir):
        """Body of run_copy_thread's try block. True when the copy finished (fix h-u)."""
        bytes_copied = 0
        # Generate timestamp
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Copy Data folder name: 'copy_data_YYYYMMDD_HHMMSS'
        copy_folder_name = f"copy_data_{timestamp_str}"
        final_destination = os.path.join(destination_dir, copy_folder_name)

        # Create folder (don't delete existing folder)
        os.makedirs(final_destination, exist_ok=True)
        self.write_detailed_log(f"Created copy destination folder: {final_destination}")

        if not self._copy_space_ok(source_paths, destination_dir):
            return False

        # --- Progress callback ---
        def progress_cb(delta):
            nonlocal bytes_copied
            bytes_copied += delta
            now = time.time()
            if now - self._last_progress_ts >= self.progress_throttle_secs:
                self._last_progress_ts = now
                self.message_queue.put(('update_progress', bytes_copied))

        # --- Actual copy loop ---
        def on_count(n):
            self.message_queue.put(('log', f"Copied {n} files..."))
            self.message_queue.put(('update_status', f"Copied {n} files"))

        files_processed = copy_sources(
            source_paths, final_destination, lambda p, _is_dir: self.is_hidden(p),
            copy_one=lambda s, d: self.copy_file_with_progress_safe(s, d, progress_cb),
            on_copy=lambda s, d: self.write_detailed_log(f"Copying file: {s} -> {d}"),
            on_count=on_count)

        # Final progress
        self.message_queue.put(('update_progress', bytes_copied))
        self.write_detailed_log(f"Copy completed: {files_processed} files, {self.format_bytes(bytes_copied)}")
        self.message_queue.put(('log', f"Copy complete ({files_processed} files)."))
        self.message_queue.put(('update_status', "Copy complete!"))
        self.message_queue.put(('open_folder', final_destination))
        return True

    def run_copy_thread(self, source_paths, destination_dir):
        """Performs the actual file and folder copying in a separate thread (byte progress)."""
        begin_job("copy")
        completed = False
        try:
            completed = self._copy_to_new_folder(source_paths, destination_dir)

        except Exception as e:
            self.write_detailed_log(f"Copy error: {e}")
            import traceback
            self.write_detailed_log(f"Copy traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"An error occurred during copy: {e}"))
        finally:
            try:
                self.close_log_file()
            except Exception:
                # closing the job log is best effort; the queue updates below must run
                _log.debug("event=job_log_close failed", exc_info=True)

            self.message_queue.put(('enable_buttons', None))

            self.message_queue.put(('update_status', "Copy complete!" if completed else "Operation finished"))
