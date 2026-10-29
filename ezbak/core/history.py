"""Last successful backup, kept for the home page (U7-2, decision D1 A).

%APPDATA%\\ezBAK\\state.json holds {"last_backup": {time, user, dest, folder, mode}}.
It is a new file next to the others; settings.json and the schedule configs are not
touched (C3). Writing never stops a backup: a failure is a warning in the diagnostic
log. The time comes from the caller (`when`), like every clock in core.
"""
import json
import logging
import os

from ezbak.winsys import paths

_log = logging.getLogger(__name__)

STATE_FILE = "state.json"
TIME_FORMAT = "%Y-%m-%d %H:%M:%S"
KEY = "last_backup"


def state_path():
    """%APPDATA%\\ezBAK\\state.json, or '' when APPDATA is not set."""
    appdata = paths.appdata_dir()
    return os.path.join(appdata, "ezBAK", STATE_FILE) if appdata else ""


def _read(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def record_backup(*, user, dest, folder, when, mode):
    """Store this backup as the last one; True when written. `mode` is 'gui' or 'headless'."""
    path = state_path()
    if not path:
        return False
    data = _read(path)
    data[KEY] = {"time": when.strftime(TIME_FORMAT), "user": user, "dest": dest,
                 "folder": folder, "mode": mode}
    tmp = path + ".tmp"
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except OSError as e:
        _log.warning("event=last_backup_write failed error=%s: %s", type(e).__name__, e)
        return False
    return True


def last_backup():
    """The stored last backup dict (time, user, dest, folder, mode), or None."""
    path = state_path()
    entry = _read(path).get(KEY) if path else None
    if not isinstance(entry, dict) or not isinstance(entry.get("time"), str):
        return None
    return entry
