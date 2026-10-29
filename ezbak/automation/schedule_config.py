"""Schedule config JSON in %APPDATA%\\ezBAK\\ezBAK_Schedule\\ (sole owner).

File name, key order, indent and encoding are fixed by C3 (ADR 0003 §6).
Values are written as given (filter types keep their case, D7).
Time and environment are read at call time: the caller passes `now`.
"""
import json
import os
import re
from dataclasses import dataclass

from ezbak.core import BackupJob
from ezbak.winsys import paths

_UNSAFE_NAME_CHARS = r'[<>:"/\\|?*]'


@dataclass
class ScheduleSpec:
    """Values written to a schedule config; field order is the JSON key order."""
    task_name: str
    user: str
    dest: str
    schedule: str
    time_str: str
    monthly_day: int = 1
    include_hidden: bool = False
    include_system: bool = False
    retention_count: int = 2
    log_retention_days: int = 30
    filters: list = None

    def to_dict(self, created_at):
        """JSON object with the 12 keys in fixed order (created_at last)."""
        return {
            'task_name': self.task_name,
            'user': self.user,
            'dest': self.dest,
            'schedule': self.schedule,
            'time_str': self.time_str,
            'monthly_day': self.monthly_day,
            'include_hidden': self.include_hidden,
            'include_system': self.include_system,
            'retention_count': self.retention_count,
            'log_retention_days': self.log_retention_days,
            'filters': self.filters,
            'created_at': created_at,
        }


def config_dir():
    """%APPDATA%\\ezBAK\\ezBAK_Schedule; RuntimeError when APPDATA is unset or empty."""
    appdata = paths.appdata_dir()
    if not appdata:
        raise RuntimeError("APPDATA environment variable not found")
    return os.path.join(appdata, 'ezBAK', 'ezBAK_Schedule')


def safe_name(task_name):
    """Task name for file names: characters invalid in file names (<>:"/\\|?*) become '_'."""
    return re.sub(_UNSAFE_NAME_CHARS, '_', task_name)


def config_path(task_name):
    """Config file path for a task; the file name is safe_name(task_name)."""
    return os.path.join(config_dir(), f"{safe_name(task_name)}.json")


def save(spec, *, now):
    """Write `spec` as JSON (created_at from `now`), creating the folder; return the path."""
    os.makedirs(config_dir(), exist_ok=True)
    path = config_path(spec.task_name)
    data = spec.to_dict(now.strftime("%Y-%m-%d %H:%M:%S"))
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    return path


def load(path):
    """Read a config file as a dict, without conversion or defaults."""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_all():
    """Every readable config dict in config_dir(), in file name order (U7: next run).

    Unreadable or non-JSON files are skipped; no APPDATA or no folder gives [].
    """
    try:
        folder = config_dir()
        names = sorted(n for n in os.listdir(folder) if n.lower().endswith('.json'))
    except (RuntimeError, OSError):
        return []
    configs = []
    for name in names:
        try:
            configs.append(load(os.path.join(folder, name)))
        except (OSError, ValueError):
            continue  # a broken file only hides that schedule from the home card
    return configs


def delete(task_name):
    """Remove the task's config file; return its path, or None when there was none."""
    path = config_path(task_name)
    if not os.path.exists(path):
        return None
    os.remove(path)
    return path


def read_raw(task_name):
    """The task's config file bytes as they are now, or None when there is no file."""
    path = config_path(task_name)
    if not os.path.exists(path):
        return None
    with open(path, 'rb') as f:
        return f.read()


def restore_raw(task_name, prev):
    """Put the config back as `read_raw` saw it: rewrite `prev`, or remove the file when None."""
    path = config_path(task_name)
    if prev is None:
        if os.path.exists(path):
            os.remove(path)
        return
    with open(path, 'wb') as f:
        f.write(prev)


def job_from_config(config, *, log_folder=None):
    """BackupJob from a loaded config dict, values as stored (no type conversion, D5).

    Missing keys take the historical defaults; log_folder comes from the command line
    (--log-folder), not from the config. A non-dict config raises AttributeError.
    """
    return BackupJob(
        user=config.get('user'),
        dest=config.get('dest'),
        include_hidden=config.get('include_hidden', False),
        include_system=config.get('include_system', False),
        retention_count=config.get('retention_count', 2),
        log_retention_days=config.get('log_retention_days', 30),
        filters=config.get('filters'),
        log_folder=log_folder,
    )
