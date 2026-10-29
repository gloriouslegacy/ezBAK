"""Unit tests for ezbak.winsys.elevation and the proc re-exports (no real UAC / ShellExecuteW)."""
import ctypes
import subprocess

import pytest

from ezbak.winsys import elevation, proc


class _FakeShell32:
    def __init__(self, admin=1, ret=42, raise_admin=None):
        self.admin = admin
        self.ret = ret
        self.raise_admin = raise_admin
        self.calls = []

    def IsUserAnAdmin(self):
        if self.raise_admin is not None:
            raise self.raise_admin
        return self.admin

    def ShellExecuteW(self, *args):
        self.calls.append(args)
        return self.ret


@pytest.fixture
def shell32(monkeypatch):
    def install(**kwargs):
        fake = _FakeShell32(**kwargs)
        # Looked up at call time, so patching after import must take effect.
        monkeypatch.setattr(ctypes.windll, "shell32", fake, raising=False)
        return fake
    return install


@pytest.mark.parametrize("value, expected", [(1, True), (0, False), (True, True)])
def test_is_admin_returns_check_result(shell32, value, expected):
    shell32(admin=value)
    assert elevation.is_admin() is expected


@pytest.mark.parametrize("exc", [OSError("no shell32"), AttributeError("IsUserAnAdmin")])
def test_is_admin_false_when_check_raises(shell32, exc):
    shell32(raise_admin=exc)
    assert elevation.is_admin() is False


def test_relaunch_as_admin_passes_runas_arguments(shell32):
    fake = shell32(ret=33)
    argv = ["C:\\app dir\\ezBAK.py", "--user", "Public", 'a"b']
    ret = elevation.relaunch_as_admin("C:\\Python\\python.exe", argv)
    assert ret == 33
    assert fake.calls == [(None, "runas", "C:\\Python\\python.exe",
                           subprocess.list2cmdline(argv), None, 1)]


def test_relaunch_as_admin_returns_failure_value_unchanged(shell32):
    shell32(ret=5)
    assert elevation.relaunch_as_admin("x.exe", ["x.py"]) == 5


def test_relaunch_as_admin_propagates_exceptions(monkeypatch):
    class _Broken:
        def ShellExecuteW(self, *args):
            raise OSError("boom")
    monkeypatch.setattr(ctypes.windll, "shell32", _Broken(), raising=False)
    with pytest.raises(OSError):
        elevation.relaunch_as_admin("x.exe", ["x.py"])


def test_proc_reexports_subprocess_names():
    assert proc.TimeoutExpired is subprocess.TimeoutExpired
    assert proc.list2cmdline is subprocess.list2cmdline
