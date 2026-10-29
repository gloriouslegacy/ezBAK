"""Unit tests for ezbak.winsys.sysinfo.free_space: fallback order and final value."""
import ctypes
import logging
import os
import shutil
from types import SimpleNamespace

import pytest

from ezbak.winsys import free_space, sysinfo

TB999 = 999 * 1024**4


class _FakeKernel32:
    def __init__(self, value=None, exc=None, ret=1):
        self.value = value
        self.exc = exc
        self.ret = ret
        self.calls = []

    def GetDiskFreeSpaceExW(self, path_p, free_p, total_p, total_free_p):
        self.calls.append(path_p.value)
        if self.exc is not None:
            raise self.exc
        if self.value is not None:
            free_p.contents.value = self.value
        return self.ret


def _disk_usage(table):
    """table: path -> free bytes, or an exception to raise. Records calls."""
    calls = []

    def fake(path):
        calls.append(path)
        result = table.get(path, OSError("no such path"))
        if isinstance(result, BaseException):
            raise result
        return SimpleNamespace(total=0, used=0, free=result)

    fake.calls = calls
    return fake


@pytest.fixture
def patch_io(monkeypatch):
    def apply(usage_table, kernel32):
        fake = _disk_usage(usage_table)
        monkeypatch.setattr(sysinfo.shutil, "disk_usage", fake)
        monkeypatch.setattr(ctypes, "windll", SimpleNamespace(kernel32=kernel32), raising=False)
        return fake
    return apply


PATH = "X:" + os.sep + "backup" + os.sep + "dest"
ROOT = "X:" + os.sep


def test_public_name_is_sysinfo_free_space():
    assert free_space is sysinfo.free_space
    assert sysinfo.FALLBACK_FREE_BYTES == TB999


def test_disk_usage_first(patch_io):
    k32 = _FakeKernel32(value=7)
    usage = patch_io({PATH: 123}, k32)
    assert free_space(PATH) == 123
    assert usage.calls == [PATH]
    assert k32.calls == []


def test_ctypes_when_disk_usage_fails(patch_io):
    k32 = _FakeKernel32(value=456)
    usage = patch_io({ROOT: 999}, k32)
    assert free_space(PATH) == 456
    assert usage.calls == [PATH]
    assert k32.calls == [PATH]


def test_ctypes_zero_return_goes_to_drive_root(patch_io):
    # GetDiskFreeSpaceExW returning 0 (failure, no exception) is not reported as 0 bytes.
    k32 = _FakeKernel32(value=None, ret=0)
    usage = patch_io({ROOT: 999}, k32)
    assert free_space(PATH) == 999
    assert k32.calls == [PATH]
    assert usage.calls == [PATH, ROOT]


def test_drive_root_when_ctypes_raises(patch_io):
    k32 = _FakeKernel32(exc=OSError("boom"))
    usage = patch_io({ROOT: 789}, k32)
    assert free_space(PATH) == 789
    assert usage.calls == [PATH, ROOT]


def test_no_windll_goes_to_drive_root(monkeypatch):
    fake = _disk_usage({ROOT: 55})
    monkeypatch.setattr(sysinfo.shutil, "disk_usage", fake)
    monkeypatch.delattr(ctypes, "windll", raising=False)
    assert free_space(PATH) == 55


def test_fallback_when_root_also_fails(patch_io):
    k32 = _FakeKernel32(exc=OSError("boom"))
    usage = patch_io({}, k32)
    assert free_space(PATH) == TB999
    assert usage.calls == [PATH, ROOT]


def test_fallback_when_no_drive(patch_io):
    k32 = _FakeKernel32(exc=OSError("boom"))
    rel = os.path.join("relative", "dir")
    usage = patch_io({}, k32)
    assert free_space(rel) == TB999
    assert usage.calls == [rel]


def test_real_disk_usage_on_tmp(tmp_path):
    assert free_space(str(tmp_path)) == shutil.disk_usage(str(tmp_path)).free


def test_ezbak_wrapper_delegates(ezbak_mod, monkeypatch):
    seen = []
    monkeypatch.setattr(ezbak_mod, "winsys_free_space", lambda p: seen.append(p) or 42)
    assert ezbak_mod.App.get_free_space_reliable(None, "D:\\x") == 42
    assert seen == ["D:\\x"]


# ---------------------------------------------------------------- system_name (P4-1)

import locale  # noqa: E402
import winreg  # noqa: E402

from ezbak.winsys import proc  # noqa: E402

WMIC_KWARGS = {"capture_output": True, "text": True, "encoding": "ENC-X", "errors": "replace"}


@pytest.fixture
def wmic(monkeypatch):
    """Fake proc.run: answers keyed by the wmic class/property word; records (argv, kwargs)."""
    answers = {}
    calls = []

    def run(argv, **kwargs):
        calls.append((list(argv), kwargs))
        answer = answers.get(argv[3], SimpleNamespace(returncode=0, stdout=""))
        if isinstance(answer, BaseException):
            raise answer
        return answer

    monkeypatch.setattr(proc, "run", run)
    monkeypatch.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "ENC-X")
    return SimpleNamespace(answers=answers, calls=calls)


def _registry(monkeypatch, value=None, open_exc=None):
    closed = []

    def open_key(*args):
        if open_exc is not None:
            raise open_exc
        return "KEY"

    def query(key, name):
        if value is None:
            raise FileNotFoundError(name)
        return value, winreg.REG_SZ

    monkeypatch.setattr(winreg, "OpenKey", open_key)
    monkeypatch.setattr(winreg, "QueryValueEx", query)
    monkeypatch.setattr(winreg, "CloseKey", closed.append)
    return closed


def _ok(text, rc=0):
    return SimpleNamespace(returncode=rc, stdout=text)


def test_registry_wins_and_skips_wmic(monkeypatch, wmic):
    closed = _registry(monkeypatch, value="PRIME Z490-P")
    out = []
    assert sysinfo.system_name(debug=out.append) == "PRIME_Z490-P"
    assert out == ["DEBUG: Using registry BaseBoardProduct: PRIME_Z490-P"]
    assert closed == ["KEY"] and wmic.calls == []


def test_priority_and_kwargs_through_proc(monkeypatch, wmic):
    _registry(monkeypatch, value=None)
    wmic.answers["product"] = _ok("Product=To Be Filled By O.E.M.\r\n")
    wmic.answers["model"] = _ok("Model=IGNORED\r\n", rc=1)
    wmic.answers["manufacturer"] = _ok("\r\nManufacturer=ACME Corp\r\n")
    out = []
    assert sysinfo.system_name(debug=out.append) == "ACME_Corp"
    assert [c[0] for c in wmic.calls] == [
        ["wmic", "baseboard", "get", "product", "/value"],
        ["wmic", "computersystem", "get", "model", "/value"],
        ["wmic", "computersystem", "get", "manufacturer", "/value"]]
    assert all(kwargs == WMIC_KWARGS for _, kwargs in wmic.calls)
    assert out == ["DEBUG: BaseBoardProduct not found in registry", "DEBUG: Using manufacturer: ACME_Corp"]


def test_encoding_read_at_call_time(monkeypatch, wmic):
    _registry(monkeypatch, open_exc=OSError("blocked"))
    monkeypatch.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "cp949")
    wmic.answers["product"] = _ok("Product=B450M\r\n")
    assert sysinfo.system_name(debug=lambda m: None) == "B450M"
    assert wmic.calls[0][1]["encoding"] == "cp949"


def test_all_fail_unknown_and_debug_order(monkeypatch, wmic):
    _registry(monkeypatch, open_exc=OSError("blocked"))
    wmic.answers["product"] = OSError("wmic missing")
    wmic.answers["model"] = _ok("Model=   \r\n")
    wmic.answers["manufacturer"] = _ok("", rc=1)
    out = []
    assert sysinfo.system_name(debug=out.append) == "Unknown"
    assert out == ["DEBUG: Registry query failed: blocked",
                   "DEBUG: Baseboard product query failed: wmic missing",
                   "DEBUG: Using fallback name: Unknown"]


def test_debug_defaults_to_the_diagnostic_log(monkeypatch, wmic, capsys, caplog):
    caplog.set_level(logging.DEBUG, logger="ezbak.winsys.sysinfo")
    _registry(monkeypatch, open_exc=OSError("blocked"))
    assert sysinfo.system_name() == "Unknown"
    assert caplog.messages[-1] == "Using fallback name: Unknown"  # P6-5d, was a print
    assert capsys.readouterr().out == ""


def _bios_values(monkeypatch, values):
    """Fake BIOS key with only `values`; records QueryValueEx names and CloseKey calls."""
    seen = []
    monkeypatch.setattr(winreg, "OpenKey", lambda *args: "KEY")
    monkeypatch.setattr(winreg, "CloseKey", lambda key: seen.append("close"))

    def query(key, name):
        seen.append(name)
        if name not in values:
            raise FileNotFoundError(name)
        return values[name], winreg.REG_SZ

    monkeypatch.setattr(winreg, "QueryValueEx", query)
    return seen


def test_registry_system_product_name_hit(monkeypatch, wmic):
    seen = _bios_values(monkeypatch, {"SystemProductName": "OptiPlex 7090", "SystemManufacturer": "Dell"})
    out = []
    assert sysinfo.system_name(debug=out.append) == "OptiPlex_7090"
    assert seen == ["BaseBoardProduct", "SystemProductName", "close"] and wmic.calls == []
    assert out == ["DEBUG: BaseBoardProduct not found in registry",
                   "DEBUG: Using registry SystemProductName: OptiPlex_7090"]


def test_registry_manufacturer_hit_when_product_missing(monkeypatch, wmic):
    seen = _bios_values(monkeypatch, {"BaseBoardProduct": "To be filled by O.E.M.",
                                      "SystemManufacturer": "Dell Inc."})
    out = []
    assert sysinfo.system_name(debug=out.append) == "Dell_Inc."
    assert seen == ["BaseBoardProduct", "SystemProductName", "SystemManufacturer", "close"]
    assert wmic.calls == []
    assert out == ["DEBUG: Using registry SystemManufacturer: Dell_Inc."]


def test_registry_new_values_missing_keeps_wmic_order(monkeypatch, wmic):
    seen = _bios_values(monkeypatch, {"SystemProductName": "System Product Name"})
    wmic.answers["manufacturer"] = _ok("Manufacturer=ACME\r\n")
    out = []
    assert sysinfo.system_name(debug=out.append) == "ACME"
    assert seen == ["BaseBoardProduct", "SystemProductName", "SystemManufacturer", "close"]
    assert [c[0][3] for c in wmic.calls] == ["product", "model", "manufacturer"]
    assert out == ["DEBUG: BaseBoardProduct not found in registry", "DEBUG: Using manufacturer: ACME"]


@pytest.mark.parametrize("raw, expected", [
    (None, None), ("", None), ("To be filled by O.E.M.", None), ("System Product Name", None),
    ("system version", None), ("   ", None), ('<>:"/\\|?*', None),
    ('  My<Board>:"X"/Y\\Z|?*  Pro \t Max  ', "MyBoardXYZ_Pro_Max"), ("Dell Inc.", "Dell_Inc."),
])
def test_clean_name(raw, expected):
    assert sysinfo.clean_name(raw) == expected


def test_ezbak_get_system_info_delegates(ezbak_mod, monkeypatch):
    monkeypatch.setattr(ezbak_mod.winsys_sysinfo, "system_name", lambda: "BOARD")
    assert ezbak_mod.get_system_info() == "BOARD"


# ---------------------------------------------------------------- used_drive_letters (P4-1)

def test_used_drive_letters_bitmask(monkeypatch):
    bits = (1 << 2) | (1 << 3) | (1 << 25)  # C, D, Z
    monkeypatch.setattr(ctypes, "windll", SimpleNamespace(
        kernel32=SimpleNamespace(GetLogicalDrives=lambda: bits)), raising=False)
    monkeypatch.setattr(os.path, "exists", lambda p: pytest.fail("no probe expected"))
    assert sysinfo.used_drive_letters() == {"C", "D", "Z"}


def test_used_drive_letters_callable_override():
    assert sysinfo.used_drive_letters(lambda: 1) == {"A"}


def test_used_drive_letters_probe_on_error(monkeypatch):
    probed = []

    def exists(path):
        probed.append(path)
        if path == "E:\\":
            raise OSError("probe failed")
        return path in ("C:\\", "Q:\\")

    def boom():
        raise OSError("GetLogicalDrives failed")

    monkeypatch.setattr(os.path, "exists", exists)
    assert sysinfo.used_drive_letters(boom) == {"C", "Q"}
    assert probed == [chr(c) + ":\\" for c in range(65, 91)]


def test_used_drive_letters_no_windll_probes(monkeypatch):
    monkeypatch.delattr(ctypes, "windll", raising=False)
    monkeypatch.setattr(os.path, "exists", lambda p: p == "C:\\")
    assert sysinfo.used_drive_letters() == {"C"}
