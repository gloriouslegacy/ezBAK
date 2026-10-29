"""ezBAK entry point: Headless run (ezbak.automation) or the GUI (ezbak.gui.run)."""
import os
from datetime import datetime
import sys

from ezbak import gui
from ezbak.winsys import elevation as winsys_elevation
from ezbak.winsys import paths as winsys_paths  # USERS_ROOT read at call time (P5 D3)
from ezbak.automation import headless as automation_headless
from ezbak.automation import cli as automation_cli


def core_run_backup(user_name, dest_dir, include_hidden=False, include_system=False, log_folder=None, retention_count=2, log_retention_days=30, filters=None):
    """Headless backup entry (ezbak.automation.headless.run_backup); winsys.paths.USERS_ROOT
    and the clock are read at call time so tests can patch them."""
    return automation_headless.run_backup(
        user_name, dest_dir, include_hidden, include_system, log_folder, retention_count,
        log_retention_days, filters, users_root=winsys_paths.USERS_ROOT, clock=lambda: datetime.now())


if __name__ == "__main__":
    is_admin = winsys_elevation.is_admin  # ezbak.winsys.elevation

    def _run_headless_if_requested():
        # Headless run (ezbak.automation.headless.run); None means: continue to the GUI
        code = automation_headless.run(automation_cli.parse(), users_root=winsys_paths.USERS_ROOT,
                                       clock=lambda: datetime.now())
        if code is not None:
            sys.exit(code)

    # Run headless mode if requested before any UI/elevation logic
    _run_headless_if_requested()

    # If running as a frozen executable (PyInstaller), the binary's manifest
    # controls elevation (e.g. --uac-admin). In that case we should not attempt
    # to relaunch via ShellExecuteW because Windows will already handle UAC.
    if getattr(sys, 'frozen', False):
        # When frozen, simply run the app. Any elevation prompt will be handled
        # by the executable's manifest if present.
        try:
            gui.run(os.path.abspath(__file__))  # App(entry_script=...) + mainloop
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
                ret = winsys_elevation.relaunch_as_admin(executable, sys.argv)
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
            gui.run(os.path.abspath(__file__))
