"""Unit tests for ezbak.targets.drivers (P4-3). pnputil never runs: proc.stream is faked."""
from datetime import datetime

import pytest

from ezbak.targets import drivers
from ezbak.winsys import proc

NOW = datetime(2025, 1, 15, 12, 0, 5)
CMD = drivers.PnputilCmd("shown", ["run"])


class FakeProcess:
    """Popen stand-in: stream lines while 'running', then communicate() leftovers."""

    def __init__(self, streamed, remaining="", stderr="", returncode=0):
        self._streamed = list(streamed)
        self._remaining = remaining
        self._stderr = stderr
        self.returncode = None
        self._final_rc = returncode
        self.stdout = self

    def poll(self):
        return None if self._streamed else self._final_rc

    def readline(self):
        return self._streamed.pop(0) if self._streamed else ""

    def communicate(self):
        self.returncode = self._final_rc
        return self._remaining, self._stderr


@pytest.fixture
def fake_stream(monkeypatch):
    calls = []

    def install(process=None, error=None):
        def stream(argv, **kwargs):
            calls.append((argv, kwargs))
            if error is not None:
                raise error
            return process
        monkeypatch.setattr(proc, "stream", stream)
        return calls
    return install


def test_build_export_cmd():
    cmd = drivers.build_export_cmd(r"D:\bk\PC_drivers_backup_x")
    assert cmd.display == 'pnputil.exe /export-driver * "D:\\bk\\PC_drivers_backup_x"'
    assert cmd.argv == ['pnputil.exe', '/export-driver', '*', 'D:\\bk\\PC_drivers_backup_x']


def test_build_add_cmd():
    cmd = drivers.build_add_cmd(r"D:\bk\PC_drivers_backup_x")
    assert cmd.display == \
        'pnputil.exe /add-driver "D:\\bk\\PC_drivers_backup_x\\*.inf" /subdirs /install'
    assert cmd.argv == ['pnputil.exe', '/add-driver', 'D:\\bk\\PC_drivers_backup_x\\*.inf',
                        '/subdirs', '/install']


def test_build_cmds_keep_percent_literal_in_argv():
    """g-r: no shell, so %VAR% in a path reaches pnputil unexpanded."""
    path = r"D:\bk\%USERNAME%_x"
    assert drivers.build_export_cmd(path).argv[-1] == path
    assert drivers.build_add_cmd(path).argv[2] == path + "\\*.inf"


def test_backup_folder_name():
    assert drivers.backup_folder_name("PRIME_Z490-P", NOW) == \
        "PRIME_Z490-P_drivers_backup_2025-01-15_120005"


def test_restore_log_name():
    assert drivers.restore_log_name("PRIME_Z490-P", NOW) == \
        "PRIME_Z490-P_drivers_restore_2025-01-15_120005.log"


@pytest.mark.parametrize("name, expected", [
    ("PRIME_Z490-P_drivers_backup_2025-01-15_120000", "PRIME_Z490-P"),
    ("A_drivers_backup_B_drivers_backup_C", "A"),
    ("PC_DRIVERS_BACKUP_2025-01-15_120000", "PC"),
    ("Prime_Z490_Drivers_Backup_2025", "Prime_Z490"),
    ("pc_drivers_backup_2025", "pc"),
    ("A_DRIVERS_BACKUP_B_drivers_backup_C", "A"),
    ("_drivers_backup_2025", ""),
    ("PC_drivers_restore_2025", None),
    ("random folder", None),
])
def test_parse_backup_folder(name, expected):
    assert drivers.parse_backup_folder(name) == expected


def test_run_pnputil_streams_output_and_returns_rc(fake_stream):
    calls = fake_stream(FakeProcess(["line 1\r\n", "", "  line 2  \n"],
                                    remaining="rest a\n\n  \nrest b\n", returncode=0))
    out = []
    cmd = drivers.build_export_cmd(r"D:\bk\%TEMP%")
    rc = drivers.run_pnputil(cmd, out.append)

    assert rc == 0
    assert out == ['Executing command: pnputil.exe /export-driver * "D:\\bk\\%TEMP%"',
                   "line 1", "line 2", "rest a", "rest b"]
    assert calls == [(['pnputil.exe', '/export-driver', '*', 'D:\\bk\\%TEMP%'],
                      {"shell": False, "stdout": proc.PIPE, "stderr": proc.PIPE,
                       "universal_newlines": True, "encoding": "oem", "errors": "replace"})]


def test_run_pnputil_stderr_and_nonzero_rc(fake_stream):
    fake_stream(FakeProcess([], stderr="bad thing\n", returncode=1))
    out = []
    assert drivers.run_pnputil(CMD, out.append) == 1
    assert out == ["Executing command: shown", "STDERR: bad thing\n"]


def test_run_pnputil_start_error_propagates_after_command_line(fake_stream):
    fake_stream(error=OSError("no pnputil"))
    out = []
    with pytest.raises(OSError, match="no pnputil"):
        drivers.run_pnputil(CMD, out.append)
    assert out == ["Executing command: shown"]
