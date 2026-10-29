"""NAS connect and disconnect with 'net use' (P4-5).

Command building, the run (cp949, retried once without encoding when it raises)
and the result text. The dialogs, the credential re-prompt flow, the GUI queue
and the detailed log stay with the caller (decisions D1=A, D3=A).

Security: the password goes into argv only. It is kept out of NasSpec's repr,
NetUseResult, exception text and diagnostic logs (proc logs the program name).
"""
import logging
from dataclasses import dataclass, field

from ezbak.winsys import proc

_log = logging.getLogger(__name__)


@dataclass(frozen=True)
class NasSpec:
    """One connect request. drive is 'Z:' style or None (no drive mapping)."""
    unc: str
    drive: str = None
    user: str = ""
    password: str = field(default=None, repr=False)
    persistent: bool = True


@dataclass(frozen=True)
class NetUseResult:
    """rc, stripped stdout/stderr, and message: the failure text ('' when rc == 0)."""
    rc: int
    out: str
    err: str
    message: str


def build_connect_cmd(spec):
    """net use [drive] unc [/user:<user> <password>] /persistent:yes|no."""
    cmd = ['net', 'use']
    if spec.drive:
        cmd += [spec.drive, spec.unc]
    else:
        cmd += [spec.unc]
    if spec.user:
        cmd += [f'/user:{spec.user}', spec.password]
    cmd += [f"/persistent:{'yes' if spec.persistent else 'no'}"]
    return cmd


def disconnect_target(text):
    """(argv, label) for 'net use <target> /delete /y', or None for a bad drive letter.

    A UNC path (starting with two backslashes) is used as given; anything else is
    read as a drive letter with an optional colon.
    """
    if text.startswith('\\\\'):
        return ['net', 'use', text, '/delete', '/y'], text
    drv = text.rstrip(':').upper() + ':'
    if not (len(drv) == 2 and 'A' <= drv[0] <= 'Z' and drv[1] == ':'):
        return None
    return ['net', 'use', drv, '/delete', '/y'], drv


def run_net_use(argv):
    """Run argv (cp949, errors ignored); when that raises, run it once more without encoding.

    An exception from the second run propagates to the caller.
    """
    try:
        res = proc.run(argv, shell=False, capture_output=True, text=True, encoding='cp949', errors='ignore')
    except Exception as exc:
        # Recovery: same command with the default encoding (current behaviour).
        _log.warning("net use retry without encoding error_type=%s", type(exc).__name__)
        res = proc.run(argv, shell=False, capture_output=True, text=True)
    out = (res.stdout or '').strip()
    err = (res.stderr or '').strip()
    rc = res.returncode
    message = "" if rc == 0 else (err or out or f"Failed with code {rc}")
    _log.debug("net use done rc=%s", rc)
    return NetUseResult(rc, out, err, message)


def free_drive_letter(used):
    """First letter not in used: Z down to D, then C, then B; 'Z' when all are used."""
    for i in range(90, 67, -1):  # Z..D
        letter = chr(i)
        if letter not in used:
            return letter
    for i in range(67, 65, -1):  # C, B (current behaviour)
        letter = chr(i)
        if letter not in used:
            return letter
    return 'Z'
