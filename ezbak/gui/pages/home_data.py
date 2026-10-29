"""Data behind the home cards of U7 (ui-plan §10, decisions D1~D4).

HomeData holds what the cards show; it is created once by home.build (no App attribute
is added later). tick(app) runs on the status chip timer (nav, every 0.5 s):
- copy speed: the change of the progress bar value per second while a job reports
  bytes (maximum > 100), pushed into the Sparkline (D4 A); nothing else is read for it;
- every SLOW_EVERY ticks a daemon thread reads the last backup (core.history, D1 A),
  the next scheduled run from the schedule configs (automation, D2 A), the free space of
  the last backup's drive (D3 A) and, once, the system name (get_system_info, may run
  wmic). The thread only sets HomeData fields; home.refresh shows them on the Tk thread.
Nothing runs while the window is not viewable (it is withdrawn during start-up and in
the tests), so a fresh window always shows the placeholders.
"""
import logging
import os
import threading
import time
from datetime import datetime

from ezbak.automation import schedule_config as automation_schedule_config
from ezbak.automation import scheduler as automation_scheduler
from ezbak.core import history as core_history
from ezbak.winsys import free_space as winsys_free_space
from ezbak.winsys import sysinfo as winsys_sysinfo

from ..actions_tools import get_system_info

_log = logging.getLogger(__name__)

SLOW_EVERY = 20  # ticks: 10 s at the 0.5 s status timer
BYTES_JOB_MIN_MAXIMUM = 100  # a maximum of 100 is a percentage job: no speed


class HomeData:
    """Values of the home data cards (None: not known yet / nothing to show)."""

    def __init__(self):
        self.system_name = None
        self.last_backup = None    # core.history.last_backup() dict
        self.next_run = None       # (task_name, datetime)
        self.free_space = None     # (drive or folder, free bytes)
        self.speed = None          # bytes per second while a job runs
        self.ticks = 0
        self.loading = False
        self._sample = None        # (monotonic time, progress value)


def tick(app, now=time.monotonic):
    """One status timer step for the home data (Tk thread)."""
    data = app.home_data
    if not app.winfo_viewable():
        return
    _sample_speed(app, data, now())
    if data.ticks % SLOW_EVERY == 0 and not data.loading:
        data.loading = True
        threading.Thread(target=_load_slow, args=(data,), daemon=True).start()
    data.ticks += 1


def _sample_speed(app, data, now):
    value = float(app.progress_bar['value'])
    maximum = float(app.progress_bar['maximum'])
    if not (maximum > BYTES_JOB_MIN_MAXIMUM and 0 < value < maximum):
        data.speed, data._sample = None, None
        return
    previous = data._sample
    data._sample = (now, value)
    if previous is None or value < previous[1]:  # a new job: start a new line
        app.home_speed_chart.clear()
        return
    elapsed = now - previous[0]
    if elapsed > 0:
        data.speed = (value - previous[1]) / elapsed
        app.home_speed_chart.push(data.speed)


def _load_slow(data):
    """Work thread: file reads and the free space query; sets HomeData fields only."""
    try:
        data.last_backup = core_history.last_backup()
        data.next_run = automation_scheduler.next_scheduled(
            automation_schedule_config.load_all(), datetime.now())
        data.free_space = _free_space(data.last_backup)
        if data.system_name is None:
            data.system_name = get_system_info()
    except Exception:
        # the cards keep their last values; the next slow tick tries again
        _log.warning("event=home_data_load failed", exc_info=True)
    finally:
        data.loading = False


def _free_space(last_backup):
    """(drive, free bytes) of the last backup's destination, None without one (D3 A)."""
    dest = (last_backup or {}).get('dest')
    if not isinstance(dest, str) or not dest:
        return None
    free = winsys_free_space(dest)
    if free >= winsys_sysinfo.FALLBACK_FREE_BYTES:  # every method failed: unknown
        return None
    return os.path.splitdrive(dest)[0] or dest, free
