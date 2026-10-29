"""Tool actions of App (moved from ezBAK.py in P5-14).

ToolsActionsMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before. Module-level names (os, threading, datetime, filedialog,
Win11Dialog, GuiReporter, InsufficientSpaceError, targets_apps, targets_drivers,
targets_backup_profiles, winsys_paths, winsys_shell, winsys_sysinfo) match ezBAK.py so facade
patches reach this module too. Command lines, folder names and job log lines are
built in ezbak.targets; these wrappers own the log file, queue, dialogs, sounds and
finally blocks (P4 decision D1). run_driver_backup / run_driver_restore are split
into module functions that run the same statements in the same order (P5 decision
D5); module functions, not App methods, so gui_harness.recording never stubs them.
"""
import logging
import os
import threading
from datetime import datetime
from tkinter import filedialog

from ezbak.core import InsufficientSpaceError
from ezbak.targets import apps as targets_apps
from ezbak.targets import backup_profiles as targets_backup_profiles
from ezbak.targets import drivers as targets_drivers
from ezbak.winsys import paths as winsys_paths  # USERS_ROOT read at call time (P5 D3)
from ezbak.winsys import shell as winsys_shell
from ezbak.winsys import sysinfo as winsys_sysinfo

from .bridge import GuiReporter, begin_job
from .dialogs import Win11Dialog

_log = logging.getLogger(__name__)


def get_system_info():
    """System name for driver folders (delegates to ezbak.winsys.sysinfo.system_name)."""
    return winsys_sysinfo.system_name()


def _open_driver_log(self, log_path):
    # Driver logs are opened by hand with a matching name (bypass open_log_file timestamp)
    with self.log_lock:
        self._log_file = open(log_path, "a", encoding="utf-8", errors="ignore")
        self.log_file_path = log_path
        _log.info("event=job_log kind=drivers path=%s", log_path)
        try:
            self._log_file.write(f"[POLICY] Hidden={self.hidden_mode_var.get()} System={self.system_mode_var.get()}\n")
            self._log_file.flush()
        except Exception:
            # the policy header line is best effort (as before); the log stays open
            _log.debug("event=job_log_policy_header failed", exc_info=True)


def _driver_backup_steps(self, folder_selected):
    # Module function: the try body of App.run_driver_backup (P5 D5). Returns success.
    # Get system info and create backup folder name
    system_name = get_system_info()
    backup_folder_name = targets_drivers.backup_folder_name(system_name, datetime.now())
    backup_full_path = os.path.join(folder_selected, backup_folder_name)

    # Create backup folder
    os.makedirs(backup_full_path, exist_ok=True)

    # Create log file manually with matching name inside the backup folder
    log_filename = f"{backup_folder_name}.log"
    log_path = os.path.join(backup_full_path, log_filename)
    _open_driver_log(self, log_path)

    self.message_queue.put(('log', f"Detailed log: {log_path}"))

    self.write_detailed_log(f"Driver backup folder created: {backup_full_path}")
    self.message_queue.put(('log', f"Backup folder: {backup_folder_name}"))

    # Execute pnputil command (ezbak.targets.drivers)
    returncode = targets_drivers.run_pnputil(
        targets_drivers.build_export_cmd(backup_full_path), self.write_detailed_log)
    if returncode == 0:
        self.write_detailed_log("Driver backup successfully completed.")
        self.message_queue.put(('log', f"Driver backup completed: {backup_folder_name}"))
        self.message_queue.put(('open_folder', backup_full_path))
        return True
    self.write_detailed_log(f"Driver backup completed with return code: {returncode}")
    self.message_queue.put(('log', f"Driver backup finished (check log for details)"))
    return False


def _driver_backup_cleanup(self, driver_backup_success):
    # Module function: the finally block of App.run_driver_backup (P5 D5).
    try:
        self.close_log_file()
    except Exception:
        # closing the job log is best effort (as before); the backup result stands
        _log.debug("event=job_log_close failed", exc_info=True)

    try:
        max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
        self.message_queue.put(('update_progress', max_val))
    except Exception:
        pass  # final progress fill is cosmetic (as before)

    # Reset progress bar
    self.message_queue.put(('reset_progress', None))

    self.message_queue.put(('enable_buttons', None))
    self.message_queue.put(('update_status', "Driver backup complete!"))

    if driver_backup_success:
        self.message_queue.put(('play_sound', 'complete'))
    else:
        self.message_queue.put(('play_sound', 'error'))


def _confirm_system_match(self, current_system_name, backup_system_name):
    # Module function: the system name check of App.run_driver_restore (P5 D5).
    # Returns False when the user cancels. The question is shown on the Tk thread
    # through the queue (g-e); this worker only waits for the answer.
    if backup_system_name.lower() != current_system_name.lower() and backup_system_name != "Unknown":
        warning_msg = (
            f"Current System: {current_system_name}\n"
            f"Backup From: {backup_system_name}\n\n"
            f"The driver backup may not match your current system.\n"
            f"Do you want to continue?"
        )

        # Ask user for confirmation
        if not self._wait_yes_no(self.translator.get('system_mismatch_warning'), warning_msg):
            self.write_detailed_log("User cancelled restore due to system mismatch")
            self.message_queue.put(('log', self.translator.get('driver_restore_cancelled')))
            return False

        self.message_queue.put(('log', f"Warning: Restoring from different system backup"))
    else:
        self.message_queue.put(('log', f"Restoring from matching system backup"))
    return True


def _driver_restore_steps(self, folder_selected):
    # Module function: the try body of App.run_driver_restore (P5 D5). Returns success.
    # Get current system info
    current_system_name = get_system_info()
    now = datetime.now()

    # Normalize path
    folder_selected = folder_selected.rstrip(os.sep)

    # Check if selected folder is a driver backup folder itself
    folder_name = os.path.basename(folder_selected)

    # Selected folder must be a driver backup folder; the name before the
    # marker is the backup system name (ezbak.targets.drivers)
    backup_system_name = targets_drivers.parse_backup_folder(folder_name)
    if backup_system_name is None:
        error_msg = f"Please select a driver backup folder.\n\nSelected: {folder_name}\n\nFolder name should contain '_drivers_backup_'"
        self.message_queue.put(('show_error', error_msg))
        return False

    # Use selected folder as driver backup folder
    driver_backup_folder = folder_selected

    # Check system name mismatch
    if not _confirm_system_match(self, current_system_name, backup_system_name):
        return False

    # Create log filename using backup system name and save inside backup folder
    log_filename = targets_drivers.restore_log_name(backup_system_name, now)
    log_path = os.path.join(driver_backup_folder, log_filename)
    _open_driver_log(self, log_path)

    # Written only now: before this point no log file is open (g-c)
    self.write_detailed_log(f"Selected driver backup folder: {driver_backup_folder}")
    self.write_detailed_log(f"Backup system name: {backup_system_name}")
    self.write_detailed_log(f"Current system name: {current_system_name}")

    self.message_queue.put(('log', f"Detailed log: {log_path}"))
    self.message_queue.put(('log', f"Restoring from: {os.path.basename(driver_backup_folder)}"))

    self.write_detailed_log(f"Driver restore started from: {driver_backup_folder}")

    # Execute pnputil restore command (only search in the specific backup folder)
    returncode = targets_drivers.run_pnputil(
        targets_drivers.build_add_cmd(driver_backup_folder), self.write_detailed_log)
    if returncode == 0:
        self.write_detailed_log("Driver restore successfully completed.")
        self.message_queue.put(('log', "Driver restore completed."))
        return True
    self.write_detailed_log(f"Driver restore completed with return code: {returncode}")
    self.message_queue.put(('log', f"Driver restore finished (check log for details)"))
    return False


class ToolsActionsMixin:
    """Tool buttons of App: winget export, browser profiles, drivers, Device Manager, Explorer."""

    def start_winget_export_thread(self):
        """Starts a thread to export installed apps list using winget export."""
        folder_selected = filedialog.askdirectory(title="Select Folder to Save Installed Apps List")
        if not folder_selected:
            self.message_queue.put(('log', "Winget export cancelled (no folder)."))
            return

        filename = targets_apps.export_file_name(datetime.now())
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
        begin_job("winget_export", output=output_path)
        try:
            result = targets_apps.export(output_path, self.write_detailed_log, clock=lambda: datetime.now())
            if result.kind == "list":
                msg = f"winget export failed (code {result.rc}). Saved fallback (winget list) to {result.path}"
            elif result.kind == "simple":
                msg = f"All winget export attempts failed. Saved basic list to {result.path}"
            else:
                msg = f"Winget export completed: {output_path}"
            self.message_queue.put(('log', msg))
            self.message_queue.put(('open_folder', os.path.dirname(output_path)))
        except FileNotFoundError:
            self.write_detailed_log("winget not found on system.")
            self.message_queue.put(('show_error', "winget is not available. Install 'App Installer' from Microsoft Store."))
        except Exception as e:
            self.write_detailed_log(f"Winget export error: {e}")
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

    def start_browser_profiles_backup_thread(self):
        """Starts backup of full browser profiles (Chrome/Edge/Firefox) of the selected user."""
        # Same user check as the regular backup (start_backup_thread)
        if not self.user_var.get():
            Win11Dialog.showwarning(self.translator.get('warning'),
                                  self.translator.get('select_user_warning'),
                                  parent=self, theme=self.theme, translator=self.translator)
            return

        dest_root = filedialog.askdirectory(title="Select Folder to Save Browser Profiles")
        if not dest_root:
            self.message_queue.put(('log', "Browser profiles backup cancelled (no destination)."))
            return

        # Confirm on the Tk thread whose profile is read (g-p); 'No' leaves nothing behind
        user_name = self.user_var.get()
        profile_path = os.path.join(winsys_paths.USERS_ROOT, user_name)
        if not Win11Dialog.askyesno(self.translator.get('confirm_browser_profiles_title'),
                                    self.translator.get('confirm_browser_profiles').format(user_name, profile_path),
                                    parent=self, theme=self.theme, translator=self.translator):
            return

        # Prepare destination and log
        self.open_log_file(dest_root, "browser_profiles")
        self.message_queue.put(('log', "backup Browser profiles : Chrome/Edge/Firefox profiles (inc. passwords, cookies, extensions)"))
        self.set_buttons_state("disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['mode'] = "determinate"
        self.message_queue.put(('update_status', "Preparing browser profiles backup..."))

        t = threading.Thread(target=self.run_browser_profiles_backup, args=(dest_root, user_name), daemon=True)
        t.start()

    def run_browser_profiles_backup(self, dest_root, user_name):
        begin_job("browser_profiles")
        completed = False
        try:
            # Sources are the selected user's AppData folders, USERS_ROOT read at call time (g-p)
            app_data = os.path.join(winsys_paths.USERS_ROOT, user_name, "AppData")
            dest_base = targets_backup_profiles(
                dest_root, self.write_detailed_log, GuiReporter(self),
                local=os.path.join(app_data, "Local"), roaming=os.path.join(app_data, "Roaming"),
                free_space=self.get_free_space_reliable, clock=lambda: datetime.now())
            if dest_base is None:  # no source at all: no folder was made (g-o)
                self.message_queue.put(('show_error', self.translator.get('no_browser_profiles')))
                return
            completed = True
            self.message_queue.put(('log', "Browser profiles backup complete (detailed log saved)."))
            self.message_queue.put(('update_status', "Browser profiles backup complete!"))
            self.message_queue.put(('open_folder', dest_base))
        except InsufficientSpaceError as e:
            self.message_queue.put(('show_error', f"Not enough free space.\nRequired: {self.format_bytes(e.required)}\nAvailable: {self.format_bytes(e.free)}"))
        except Exception as e:
            self.write_detailed_log(f"Browser profiles backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during browser profiles backup: {e}"))
        finally:
            self.close_log_file()
            try:
                max_val = self.progress_bar['maximum'] if self.progress_bar['maximum'] > 0 else 1
                self.message_queue.put(('update_progress', max_val))
            except Exception:
                pass  # final progress fill is cosmetic (as before)
            self.message_queue.put(('update_status', "Browser profiles backup complete!" if completed else "Operation finished"))
            self.message_queue.put(('enable_buttons', None))

    def start_driver_backup_thread(self):
        """Starts the driver backup process in a separate thread."""
        folder_selected = filedialog.askdirectory(title="Select Driver Backup Folder")
        if not folder_selected:
            self.message_queue.put(('log', "Driver backup folder selection cancelled."))
            return

        self.message_queue.put(('log', "Driver backup process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver backup in progress..."))

        backup_thread = threading.Thread(target=self.run_driver_backup, args=(folder_selected,), daemon=True)
        backup_thread.start()

    def run_driver_backup(self, folder_selected):
        """Execute driver backup with matching log filename"""
        # Folder name examples: "OptiPlex_7090_drivers_2025-01-15", "Unknown_Desktop_drivers_2025-01-15"
        begin_job("driver_backup")
        driver_backup_success = False
        try:
            driver_backup_success = _driver_backup_steps(self, folder_selected)
        except Exception as e:
            self.write_detailed_log(f"Driver backup error: {e}")
            self.message_queue.put(('show_error', f"An error occurred during driver backup: {e}"))
            driver_backup_success = False
        finally:
            _driver_backup_cleanup(self, driver_backup_success)

    def start_driver_restore_thread(self):
        """Starts the driver restore process in a separate thread."""
        folder_selected = filedialog.askdirectory(title="Select Driver Restore Folder")
        if not folder_selected:
            self.message_queue.put(('log', "Driver restore folder selection cancelled."))
            return

        self.message_queue.put(('log', "Driver restore process started."))
        self.set_buttons_state("disabled")
        self.progress_bar['mode'] = "indeterminate"
        self.progress_bar.start()
        self.message_queue.put(('update_status', "Driver restore in progress..."))

        restore_thread = threading.Thread(target=self.run_driver_restore, args=(folder_selected,), daemon=True)
        restore_thread.start()

    def run_driver_restore(self, folder_selected):
        """Execute driver restore with matching log filename"""
        # The selected folder must be a driver backup folder; a backup from another
        # system (name before the marker) asks to continue, "Unknown" never asks.
        begin_job("driver_restore")
        driver_restore_success = False
        try:
            driver_restore_success = _driver_restore_steps(self, folder_selected)
        except Exception as e:
            self.write_detailed_log(f"Driver restore error: {e}")
            import traceback
            self.write_detailed_log(f"Traceback: {traceback.format_exc()}")
            self.message_queue.put(('show_error', f"An error occurred during driver restore: {e}"))
            driver_restore_success = False
        finally:
            self.close_log_file()
            self.message_queue.put(('stop_progress', None))
            self.message_queue.put(('update_status', "Driver restore complete!"))
            self.message_queue.put(('enable_buttons', None))

            if driver_restore_success:
                self.message_queue.put(('play_sound', 'complete'))
            else:
                self.message_queue.put(('play_sound', 'error'))

    def open_device_manager(self):
        """Opens the Device Manager (fallbacks in ezbak.winsys.shell.open_device_manager)."""
        try:
            self.log(winsys_shell.open_device_manager())
        except Exception as e:
            self.log(f"Error opening Device Manager: {e}")
            try:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('unable_to_open_devmgr').format(e),
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                # the error dialog is best effort (as before); the error is already logged
                _log.debug("event=devmgr_error_dialog failed", exc_info=True)

    def open_file_explorer(self):
        """Opens File Explorer to the user's home directory."""
        try:
            user_home = os.path.expanduser("~")
            os.startfile(user_home)
            self.log(f"Opening File Explorer to: {user_home}")
        except Exception as e:
            self.log(f"Error opening File Explorer: {e}")
