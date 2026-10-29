"""Administrator check and UAC relaunch (IsUserAnAdmin, ShellExecuteW "runas").

ctypes.windll.shell32 is looked up on every call, not at import, so tests can
replace it. Command-line quoting comes from proc.list2cmdline (subprocess stays in proc).
"""
import ctypes
import logging

from .proc import list2cmdline

log = logging.getLogger(__name__)


def is_admin():
    """True when the process runs elevated; False if the check itself fails."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        log.debug("elevation.is_admin check failed", exc_info=True)
        return False


def relaunch_as_admin(executable, argv):
    """ShellExecuteW(None, "runas", executable, list2cmdline(argv), None, 1); returns its value.

    The caller interprets the result (<= 32 means failure) and handles exceptions.
    """
    params = list2cmdline(argv)
    log.info("elevation.relaunch_as_admin program=%s", executable)
    return ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, params, None, 1)
