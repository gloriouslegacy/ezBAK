"""Driver backup and restore with pnputil (P4-3).

Command strings, folder and log names, and the pnputil run with its output
handling. Opening the log file, the GUI queue, dialogs and sounds stay with the
caller (decision D1=A). Names and the logged command text are C3: kept
character for character. The command runs as an argv list with shell=False and
the OEM code page (g-r, F2=A), so cmd never expands %VAR% inside paths and
undecodable output is replaced instead of raising.
"""
import logging
from typing import NamedTuple

from ezbak.winsys import proc

_log = logging.getLogger(__name__)

BACKUP_MARKER = "_drivers_backup_"
_STAMP = "%Y-%m-%d_%H%M%S"


class PnputilCmd(NamedTuple):
    """A pnputil command: display is the logged text (C3), argv is what runs."""
    display: str
    argv: list


def build_export_cmd(path):
    """pnputil command exporting every third-party driver into path."""
    return PnputilCmd(f'pnputil.exe /export-driver * "{path}"',
                      ['pnputil.exe', '/export-driver', '*', f'{path}'])


def build_add_cmd(folder):
    """pnputil command installing every .inf found under folder (subdirs included)."""
    return PnputilCmd(f'pnputil.exe /add-driver "{folder}\\*.inf" /subdirs /install',
                      ['pnputil.exe', '/add-driver', f'{folder}\\*.inf', '/subdirs', '/install'])


def backup_folder_name(system, now):
    """Backup folder name: <system>_drivers_backup_<YYYY-mm-dd_HHMMSS>."""
    return f"{system}{BACKUP_MARKER}{now.strftime(_STAMP)}"


def parse_backup_folder(name):
    """System name in a backup folder name, or None when it is not a backup folder.

    The marker is found ignoring case (first occurrence) and the system name is the
    part of the original name before it, case unchanged.
    """
    index = name.lower().find(BACKUP_MARKER)
    if index < 0:
        return None
    return name[:index]


def restore_log_name(system, now):
    """Restore log file name: <system>_drivers_restore_<YYYY-mm-dd_HHMMSS>.log."""
    return f"{system}_drivers_restore_{now.strftime(_STAMP)}.log"


def run_pnputil(cmd, out):
    """Run cmd (a PnputilCmd), send each output line to out(line), return the return code.

    out() first gets "Executing command: <cmd.display>", then stdout lines while the
    process runs, then lines left after communicate(), then "STDERR: ..." when
    stderr is not empty. Exceptions from starting the process propagate.
    """
    out(f"Executing command: {cmd.display}")
    process = proc.stream(cmd.argv, shell=False, stdout=proc.PIPE, stderr=proc.PIPE,
                          universal_newlines=True, encoding='oem', errors='replace')

    while process.poll() is None:
        line = process.stdout.readline()
        if line:
            out(line.strip())

    remaining_output, error_output = process.communicate()
    if remaining_output:
        for line in remaining_output.splitlines():
            if line.strip():
                out(line.strip())

    if error_output:
        out(f"STDERR: {error_output}")

    _log.debug("drivers.pnputil rc=%s", process.returncode)
    return process.returncode
