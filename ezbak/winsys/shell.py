"""Shell launchers: Device Manager (P4) and File Explorer (P5).

open_device_manager is moved from ezBAK.App.open_device_manager without behaviour
change (P4-2). The caller writes the job log line and shows the error dialog.
"""
import logging
import os

from . import proc

log = logging.getLogger(__name__)

MSG_STARTFILE = "Opening Device Manager..."
MSG_MMC = "Opening Device Manager via mmc..."
MSG_SHELL_START = "Opening Device Manager (via shell start)..."


def open_device_manager():
    """Open Device Manager; return the log line of the method that worked.

    1. os.startfile(%WINDIR%\\system32\\devmgmt.msc) when that file exists
       (an exception here propagates; steps 2-3 are not tried)
    2. mmc devmgmt.msc (Popen, shell=False)
    3. 'start devmgmt.msc' (Popen, shell=True)
    Raises FileNotFoundError when steps 2 and 3 both fail.
    """
    windir = os.environ.get("WINDIR", r"C:\Windows")
    devmgmt_path = os.path.join(windir, "system32", "devmgmt.msc")
    if os.path.exists(devmgmt_path):
        os.startfile(devmgmt_path)
        log.info("devmgr.open method=startfile")
        return MSG_STARTFILE

    try:
        proc.stream(["mmc", "devmgmt.msc"], shell=False)
        log.info("devmgr.open method=mmc")
        return MSG_MMC
    except Exception as exc:  # recovery: try the next launcher
        log.warning("devmgr.open method=mmc failed error=%s", type(exc).__name__)

    try:
        proc.stream('start devmgmt.msc', shell=True)
        log.info("devmgr.open method=shell_start")
        return MSG_SHELL_START
    except Exception as exc:  # recovery: fall through to the final error
        log.warning("devmgr.open method=shell_start failed error=%s", type(exc).__name__)

    raise FileNotFoundError("devmgmt.msc not found / cannot be launched")
