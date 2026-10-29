"""Schedule Backup actions of App (moved from ezBAK.py in P5-12).

ScheduleActionsMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before. The script path written into the scheduled task comes from
self.entry_script (set by App to the former os.path.abspath(__file__) of ezBAK.py);
frozen-exe handling stays in ezbak.automation.scheduler. Module-level names
(threading, datetime, ScheduleSpec, automation_scheduler, automation_schedule_config,
ScheduleBackupDialog, Win11Dialog, ...) match ezBAK.py so facade patches reach this
module too. schedule_backup is split into a helper that runs the same statements in
the same order (P5 decision D5).
"""
import logging
import threading
from datetime import datetime

from ezbak.automation import ScheduleSpec
from ezbak.automation import schedule_config as automation_schedule_config
from ezbak.automation import scheduler as automation_scheduler

from .bridge import begin_job
from .dialogs import ScheduleBackupDialog, Win11Dialog
from .settings import copy_default_filters

_log = logging.getLogger(__name__)


def _handle_schedule_dialog_action(self, dlg):
    # Module function, not an App method: runs after the dialog closes (split out of
    # App.schedule_backup, P5 D5); `self` is the App, so its calls stay App calls.
    # Check if dlg still exists and is valid
    if dlg and hasattr(dlg, 'action'):
        action = dlg.action
        _log.debug(f"Dialog action = {action}")

        if action == 'create':
            if hasattr(dlg, 'result') and dlg.result:
                data = dlg.result
                _log.debug(f"Creating scheduled task with data: {data}")

                # Disable buttons to prevent multiple clicks during task creation
                self.set_buttons_state("disabled")

                # Create and start thread for task creation (prevents UI blocking)
                task_thread = threading.Thread(
                    target=self.run_create_scheduled_task_thread,
                    args=(data,),
                    daemon=True
                )
                task_thread.start()
                _log.debug("Scheduled task creation thread started")
            else:
                _log.debug("No result data found")

        elif action == 'delete':
            task_name = getattr(dlg, 'task_name', None)
            _log.debug(f"Deleting task: {task_name}")
            if task_name:
                try:
                    self.delete_scheduled_task(task_name)
                except Exception as delete_error:
                    _log.debug(f"Task deletion failed: {delete_error}")
                    error_msg = f"Failed to delete scheduled task: {str(delete_error)}"
                    self.message_queue.put(('log', error_msg))
                    try:
                        Win11Dialog.showerror(self.translator.get('task_deletion_failed'), error_msg,
                                            parent=self, theme=self.theme, translator=self.translator)
                    except Exception:
                        _log.warning(f"{error_msg}")
        else:
            _log.debug(f"Action '{action}' - no operation needed")
    else:
        _log.debug("Dialog has no action attribute or is None")


class ScheduleActionsMixin:
    def schedule_backup(self):
        dlg = None
        try:
            _log.debug("Creating ScheduleBackupDialog...")
            dlg = ScheduleBackupDialog(self)
            _log.debug(f"Dialog created, action={getattr(dlg, 'action', 'NOT_SET')}")

            _log.debug("Calling wait_window...")
            self.wait_window(dlg)
            _log.debug("wait_window completed")

            _handle_schedule_dialog_action(self, dlg)

        except Exception as e:
            _log.debug(f"Exception in schedule_backup: {type(e).__name__}: {e}")
            import traceback
            _log.debug(f"Traceback: {traceback.format_exc()}")

            error_msg = f"Schedule dialog error: {str(e)}"
            self.message_queue.put(('log', error_msg))

            # Handle the messagebox call safely as wel
            try:
                Win11Dialog.showerror(self.translator.get('error'), error_msg,
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception as mb_error:
                _log.debug(f"messagebox.showerror failed: {mb_error}")

    ############# Schedule Backup Functions #############

    def _get_schedule_config_dir(self):
        """Get the directory path for schedule configuration files"""
        return automation_schedule_config.config_dir()

    def create_scheduled_task(self, task_name, schedule, time_str, user, dest, monthly_day=1, include_hidden=False, include_system=False, retention_count=2, log_retention_days=30, filters=None):
        """
        Create a scheduled task (set to run with administrator privileges)
        """
        _log.debug(f"create_scheduled_task called with task_name={task_name}, schedule={schedule}, time_str={time_str}")
        try:
            spec = ScheduleSpec(
                task_name=task_name, user=user, dest=dest, schedule=schedule, time_str=time_str,
                monthly_day=monthly_day, include_hidden=include_hidden, include_system=include_system,
                retention_count=retention_count, log_retention_days=log_retention_days,
                filters=filters if filters else copy_default_filters())
            success_msg = automation_scheduler.create_task(
                spec, self.entry_script, out=self.write_detailed_log, now=datetime.now())
            self.message_queue.put(('log', f"Scheduled task '{task_name}' created with administrator privileges"))
            # Return success message instead of showing dialog here (dialog will be shown from main thread)
            return success_msg
        except Exception as e:
            error_msg = f"Failed to create scheduled task '{task_name}': {str(e)}"
            self.write_detailed_log(f"Task creation failed: {e}")

            # Log detailed error information
            import traceback
            self.write_detailed_log(f"Full error traceback: {traceback.format_exc()}")

            _log.warning(f"{error_msg}")
            raise RuntimeError(error_msg) from e

    def run_create_scheduled_task_thread(self, task_data):
        """
        Thread wrapper for create_scheduled_task to prevent UI blocking
        """
        begin_job("schedule_create")
        try:
            _log.debug(f"Thread started with task_name={task_data.get('task_name')}")
            self.message_queue.put(('log', f"Creating scheduled task '{task_data.get('task_name')}'..."))

            _log.debug(f"Calling create_scheduled_task...")
            # Call the actual task creation method
            result = self.create_scheduled_task(
                task_name=task_data.get('task_name'),
                dest=task_data.get('dest'),
                schedule=task_data.get('schedule'),
                time_str=task_data.get('time_str'),
                monthly_day=task_data.get('monthly_day', 1),
                include_hidden=task_data.get('include_hidden', False),
                include_system=task_data.get('include_system', False),
                retention_count=int(task_data.get('retention_count', 2)),
                log_retention_days=int(task_data.get('log_retention_days', 30)),
                user=task_data.get('user'),
                filters=task_data.get('filters', copy_default_filters())
            )

            _log.debug(f"create_scheduled_task returned: {result}")
            self.message_queue.put(('log', f"Scheduled task '{task_data.get('task_name')}' created successfully"))

            # Show success dialog in main thread (result is the success message)
            if result:
                _log.debug(f"Scheduling success dialog in main thread")
                self.after(100, lambda msg=result: self._show_task_creation_success(msg))
            else:
                _log.debug(f"No result returned from create_scheduled_task")

        except Exception as task_error:
            # Handle errors in the thread
            _log.debug(f"Exception in thread: {task_error}")
            error_msg = f"Failed to create scheduled task: {str(task_error)}"
            self.write_detailed_log(f"Thread error: {error_msg}")
            self.message_queue.put(('log', error_msg))

            # Schedule error dialog to show in main thread
            _log.debug(f"Scheduling error dialog in main thread")
            self.after(100, lambda: self._show_task_creation_error(error_msg))
        finally:
            # Re-enable buttons in main thread
            _log.debug(f"Re-enabling buttons")
            self.after(100, lambda: self.set_buttons_state("normal"))

    def _show_task_creation_success(self, success_msg):
        """Show task creation success in main thread"""
        try:
            Win11Dialog.showinfo(
                self.translator.get('schedule_created'),
                success_msg,
                parent=self,
                theme=self.theme,
                translator=self.translator
            )
        except Exception as dlg_error:
            self.write_detailed_log(f"Error showing success dialog: {dlg_error}")
            _log.info(f"Task created successfully: {success_msg}")

    def _show_task_creation_error(self, error_msg):
        """Show task creation error in main thread"""
        try:
            Win11Dialog.showerror(
                self.translator.get('task_creation_failed'),
                error_msg,
                parent=self,
                theme=self.theme,
                translator=self.translator
            )
        except Exception as dlg_error:
            self.write_detailed_log(f"Error showing dialog: {dlg_error}")
            _log.warning(f"{error_msg}")

    def delete_scheduled_task(self, task_name):
        if not task_name:
            try:
                Win11Dialog.showwarning(self.translator.get('warning'), self.translator.get('task_name_required'),
                                      parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                # dialog display failure only; the missing name is simply ignored
                _log.debug("event=task_name_warning_dialog failed", exc_info=True)
            return
        try:
            automation_scheduler.delete_task(task_name)
        except Exception as e:
            try:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('failed_to_delete_task').format(e),
                                    parent=self, theme=self.theme, translator=self.translator)
            except Exception:
                # dialog display failure only; the failure is still logged below
                _log.debug("event=task_delete_error_dialog failed", exc_info=True)
            self.message_queue.put(('log', f"Schedule delete failed: {e}"))
            return

        # Delete the schedule configuration file
        try:
            config_file = automation_schedule_config.delete(task_name)
            if config_file:
                self.write_detailed_log(f"Schedule configuration file deleted: {config_file}")
                self.message_queue.put(('log', f"Config file deleted: {config_file}"))
        except Exception as config_error:
            self.write_detailed_log(f"Failed to delete schedule config file: {config_error}")
            # Don't fail the whole operation if config file deletion fails

        try:
            Win11Dialog.showinfo(self.translator.get('task_scheduler'), self.translator.get('task_deleted').format(task_name),
                               parent=self, theme=self.theme, translator=self.translator)
        except Exception:
            # dialog display failure only; the deletion is still logged below
            _log.debug("event=task_deleted_dialog failed", exc_info=True)
        self.message_queue.put(('log', f"Task deleted: {task_name}"))
