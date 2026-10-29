"""Unit tests for ezbak.winsys.shell.open_device_manager: 3-step fallback and failure.

os.startfile and proc.stream are replaced; Device Manager is never opened.
"""
import os

import pytest

from ezbak.winsys import proc, shell


class _Calls(list):
    """Recorded launcher calls plus the set of launchers that should fail."""

    def __init__(self):
        super().__init__()
        self.fail = set()
        self.msc = None


@pytest.fixture
def calls(monkeypatch, tmp_path):
    """Fake WINDIR, startfile and proc.stream. Set calls.fail to make a launcher raise."""
    record = _Calls()
    monkeypatch.setenv("WINDIR", str(tmp_path))
    (tmp_path / "system32").mkdir()

    def startfile(path):
        record.append(("startfile", path))
        if "startfile" in record.fail:
            raise OSError("startfile failed")

    def stream(argv, **kwargs):
        record.append(("stream", argv, kwargs))
        key = argv[0] if isinstance(argv, list) else "start"
        if key in record.fail:
            raise OSError(f"{key} failed")
        return object()

    monkeypatch.setattr(os, "startfile", startfile, raising=False)
    monkeypatch.setattr(proc, "stream", stream)
    record.msc = tmp_path / "system32" / "devmgmt.msc"
    return record


def test_startfile_when_msc_exists(calls):
    calls.msc.write_text("", encoding="utf-8")
    assert shell.open_device_manager() == "Opening Device Manager..."
    assert calls == [("startfile", str(calls.msc))]


def test_startfile_error_propagates_without_fallback(calls):
    calls.msc.write_text("", encoding="utf-8")
    calls.fail.add("startfile")
    with pytest.raises(OSError, match="startfile failed"):
        shell.open_device_manager()
    assert calls == [("startfile", str(calls.msc))]


def test_mmc_when_msc_missing(calls):
    assert shell.open_device_manager() == "Opening Device Manager via mmc..."
    assert calls == [("stream", ["mmc", "devmgmt.msc"], {"shell": False})]


def test_shell_start_when_mmc_fails(calls):
    calls.fail.add("mmc")
    assert shell.open_device_manager() == "Opening Device Manager (via shell start)..."
    assert calls == [
        ("stream", ["mmc", "devmgmt.msc"], {"shell": False}),
        ("stream", "start devmgmt.msc", {"shell": True}),
    ]


def test_all_fail_raises_file_not_found(calls):
    calls.fail.update({"mmc", "start"})
    with pytest.raises(FileNotFoundError) as info:
        shell.open_device_manager()
    assert str(info.value) == "devmgmt.msc not found / cannot be launched"
    assert len(calls) == 2


def test_windir_default(monkeypatch):
    monkeypatch.delenv("WINDIR", raising=False)
    seen = []
    monkeypatch.setattr(os.path, "exists", lambda p: seen.append(p) or False)
    monkeypatch.setattr(proc, "stream", lambda argv, **kw: object())
    assert shell.open_device_manager() == "Opening Device Manager via mmc..."
    assert seen == [os.path.join(r"C:\Windows", "system32", "devmgmt.msc")]
