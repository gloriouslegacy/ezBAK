"""Unit tests for ezbak.winsys: proc (pass-through), fileattrs (pywin32 / ctypes), paths."""
import ctypes
import importlib
import importlib.util
import os
import subprocess
import sys
from pathlib import Path

import pytest

from ezbak.winsys import fileattrs, paths, proc

# Same module object as the characterization tests (no tests/__init__.py, so load by path).
_USER_TREE = "ezbak_test_user_tree"
if _USER_TREE not in sys.modules:
    _spec = importlib.util.spec_from_file_location(
        _USER_TREE, Path(__file__).resolve().parent / "fixtures" / "user_tree.py")
    sys.modules[_USER_TREE] = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(sys.modules[_USER_TREE])
user_tree = sys.modules[_USER_TREE]

H = 0x02
S = 0x04
R = 0x400


# ---------------------------------------------------------------- proc

class _Sentinel:
    pass


def test_proc_run_passes_argv_and_kwargs_unchanged(monkeypatch):
    calls = []
    result = _Sentinel()

    def fake_run(*args, **kwargs):
        calls.append((args, kwargs))
        return result

    monkeypatch.setattr(subprocess, "run", fake_run)
    argv = ["schtasks", "/Query", "/TN", "ezBAK_x"]
    kwargs = dict(capture_output=True, text=True, encoding="cp949", errors="ignore",
                  creationflags=0x08000000, shell=False)
    assert proc.run(argv, **kwargs) is result
    assert calls == [((argv,), kwargs)]
    assert calls[0][0][0] is argv


def test_proc_run_adds_no_defaults(monkeypatch):
    calls = []
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: calls.append((a, k)))
    proc.run("net use Z: /delete")
    assert calls == [(("net use Z: /delete",), {})]


def test_proc_stream_passes_argv_and_kwargs_unchanged(monkeypatch):
    calls = []
    result = _Sentinel()

    def fake_popen(*args, **kwargs):
        calls.append((args, kwargs))
        return result

    monkeypatch.setattr(subprocess, "Popen", fake_popen)
    argv = ["pnputil", "/add-driver", "x.inf", "/install"]
    kwargs = dict(stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
    assert proc.stream(argv, **kwargs) is result
    assert calls == [((argv,), kwargs)]


def test_proc_stream_adds_no_defaults(monkeypatch):
    calls = []
    monkeypatch.setattr(subprocess, "Popen", lambda *a, **k: calls.append((a, k)))
    proc.stream(["explorer", "C:\\"])
    assert calls == [((["explorer", "C:\\"],), {})]


# ---------------------------------------------------------------- fileattrs

def test_fileattrs_constants_match_winbase():
    assert fileattrs.FILE_ATTRIBUTE_READONLY == 0x01
    assert fileattrs.FILE_ATTRIBUTE_HIDDEN == 0x02
    assert fileattrs.FILE_ATTRIBUTE_SYSTEM == 0x04
    assert fileattrs.FILE_ATTRIBUTE_DIRECTORY == 0x10
    assert fileattrs.FILE_ATTRIBUTE_ARCHIVE == 0x20
    assert fileattrs.FILE_ATTRIBUTE_DEVICE == 0x40
    assert fileattrs.FILE_ATTRIBUTE_NORMAL == 0x80
    assert fileattrs.FILE_ATTRIBUTE_REPARSE_POINT == 0x400


@pytest.fixture
def attr_files(tmp_path):
    """plain file, hidden file, hidden+system dir, junction; attributes via SetFileAttributesW."""
    plain = tmp_path / "plain.txt"
    plain.write_text("x")
    hidden = tmp_path / "hidden.txt"
    hidden.write_text("x")
    hs_dir = tmp_path / "hs_dir"
    hs_dir.mkdir()
    target = tmp_path / "target"
    target.mkdir()
    link = tmp_path / "link"
    user_tree.make_junction(str(link), str(target))
    user_tree.set_attributes(str(hidden), H)
    user_tree.set_attributes(str(hs_dir), H | S)
    yield {"plain": str(plain), "hidden": str(hidden), "hs_dir": str(hs_dir),
           "link": str(link), "missing": str(tmp_path / "missing")}
    for p in (hidden, hs_dir):
        user_tree.set_attributes(str(p), 0x80 if p.is_file() else 0x10)


def _check_bits(get, files):
    assert not get(files["plain"]) & (H | S | R)
    assert get(files["hidden"]) & H
    assert not get(files["hidden"]) & S
    hs = get(files["hs_dir"])
    assert hs & H and hs & S and hs & 0x10
    assert get(files["link"]) & R


@pytest.mark.windows
def test_fileattrs_ctypes_path(attr_files, monkeypatch):
    monkeypatch.setattr(fileattrs, "HAS_PYWIN32", False)
    _check_bits(fileattrs.get_attributes, attr_files)
    with pytest.raises(OSError, match="GetFileAttributesW failed for"):
        fileattrs.get_attributes(attr_files["missing"])


@pytest.mark.windows
def test_fileattrs_pywin32_path(attr_files):
    pytest.importorskip("win32api")
    assert fileattrs.HAS_PYWIN32 is True
    _check_bits(fileattrs.get_attributes, attr_files)
    import pywintypes
    with pytest.raises(pywintypes.error):
        fileattrs.get_attributes(attr_files["missing"])


@pytest.fixture
def fileattrs_nopywin32():
    """Reload fileattrs with win32api/win32con hidden, as on a machine without pywin32."""
    with pytest.MonkeyPatch.context() as mp:
        mp.setitem(sys.modules, "win32api", None)
        mp.setitem(sys.modules, "win32con", None)
        mod = importlib.reload(fileattrs)
    yield mod
    importlib.reload(fileattrs)


@pytest.mark.windows
def test_fileattrs_reimport_without_pywin32(fileattrs_nopywin32, attr_files):
    mod = fileattrs_nopywin32
    assert mod.HAS_PYWIN32 is False
    assert mod.win32api is None
    _check_bits(mod.get_attributes, attr_files)


@pytest.mark.windows
def test_fileattrs_pywin32_and_ctypes_paths_agree(attr_files, monkeypatch):
    """The pywin32 branch and the ctypes branch return the same attribute words."""
    pytest.importorskip("win32api")
    for key in ("plain", "hidden", "hs_dir", "link"):
        path = attr_files[key]
        via_pywin32 = fileattrs.get_attributes(path)
        with monkeypatch.context() as mp:
            mp.setattr(fileattrs, "HAS_PYWIN32", False)
            assert fileattrs.get_attributes(path) == via_pywin32


# ---------------------------------------------------------------- paths

def test_users_root_matches_ezbak(ezbak_mod):
    assert paths.users_root() == os.path.join("C:", os.sep, "Users")
    assert paths.users_root() == ezbak_mod.USERS_ROOT


def test_appdata_dir_reads_env_at_call_time(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path / "a"))
    assert paths.appdata_dir() == str(tmp_path / "a")
    monkeypatch.setenv("APPDATA", str(tmp_path / "b"))
    assert paths.appdata_dir() == str(tmp_path / "b")
    monkeypatch.delenv("APPDATA")
    assert paths.appdata_dir() == ""


def test_resource_path_uses_meipass(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
    assert paths.resource_path("icon.ico") == os.path.join(str(tmp_path), "icon.ico")


def test_resource_path_falls_back_to_cwd(monkeypatch, tmp_path, ezbak_mod):
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    monkeypatch.chdir(tmp_path)
    assert paths.resource_path("icon.ico") == os.path.join(os.path.abspath("."), "icon.ico")
    assert paths.resource_path("icon.ico") == ezbak_mod.resource_path("icon.ico")


def test_list_users_applies_load_users_rules(tmp_path, monkeypatch):
    root = tmp_path / "Users"
    root.mkdir()
    for name in ("alice", "Bob", ".hidden", "All Users", "ALL USERS2", "Public"):
        (root / name).mkdir()
    (root / "desktop.ini").write_text("x")
    assert sorted(paths.list_users(str(root))) == sorted(["alice", "Bob", "ALL USERS2", "Public"])
    monkeypatch.setattr(paths, "users_root", lambda: str(root))
    assert sorted(paths.list_users()) == sorted(["alice", "Bob", "ALL USERS2", "Public"])


def test_list_users_missing_root_returns_empty(tmp_path):
    assert paths.list_users(str(tmp_path / "nope")) == []
