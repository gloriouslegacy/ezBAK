"""Characterization (golden) of hidden/system/skip judgment in the three current implementations.

- gui:      App.is_hidden                  (ezBAK.py)
- safe:     App.is_hidden_safe             (ezBAK.py)
- headless: ezbak.core.is_skipped(BackupPolicy.HEADLESS, check_filters=False) (engine path)

Attributes are set with ctypes SetFileAttributesW on entries under tmp_path.
Each implementation runs under two module variants: the normal import (pywin32 present)
and a copy imported with win32api/win32con unavailable (ctypes fallback branch).
Expected values are today's results; differences are kept as-is.
"""
import ctypes
import importlib.util
import os
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

# Same module object as test_backup_golden.py (no tests/__init__.py, so load by path).
_USER_TREE = "ezbak_test_user_tree"
if _USER_TREE not in sys.modules:
    _spec = importlib.util.spec_from_file_location(
        _USER_TREE, Path(__file__).resolve().parent.parent / "fixtures" / "user_tree.py")
    sys.modules[_USER_TREE] = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(sys.modules[_USER_TREE])
user_tree = sys.modules[_USER_TREE]

pytestmark = pytest.mark.windows

H = 0x02  # FILE_ATTRIBUTE_HIDDEN
S = 0x04  # FILE_ATTRIBUTE_SYSTEM

# (name, kind, attributes) — kind: file / dir / junction / missing
ENTRIES = [
    ("normal.txt", "file", 0),
    ("hidden.txt", "file", H),
    ("system.txt", "file", S),
    ("hidsys.txt", "file", H | S),
    ("hiddendir", "dir", H),
    ("systemdir", "dir", S),
    (".dotfile", "file", 0),
    ("thumbs.db", "file", 0),
    ("desktop.ini", "file", 0),
    ("$recycle.bin", "dir", 0),
    ("AppData", "dir", 0),
    ("Application Data", "dir", 0),
    ("Local Settings", "dir", 0),
    ("junction", "junction", 0),
    ("missing.txt", "missing", 0),
]

# Policies in order: (include_hidden, include_system)
POLICIES = [(False, False), (True, False), (False, True), (True, True)]

# EXPECTED[variant][name] = (gui, safe, headless); each string has one digit per policy
# in POLICIES order, "1" = skipped, "0" = kept.
EXPECTED = {
    "pywin32": {
        "normal.txt":       ("0000", "0000", "0000"),
        "hidden.txt":       ("1010", "1010", "1010"),
        "system.txt":       ("1100", "1100", "1100"),
        "hidsys.txt":       ("1110", "1110", "1110"),
        "hiddendir":        ("1010", "1010", "1010"),
        "systemdir":        ("1100", "1100", "1100"),
        ".dotfile":         ("0000", "0000", "1111"),
        "thumbs.db":        ("0000", "0000", "1111"),
        "desktop.ini":      ("0000", "0000", "1111"),
        "$recycle.bin":     ("0000", "0000", "1111"),
        "AppData":          ("1110", "0000", "1110"),
        "Application Data": ("1110", "0000", "1110"),
        "Local Settings":   ("1110", "0000", "1110"),
        "junction":         ("1111", "1111", "1111"),
        "missing.txt":      ("0000", "0000", "0000"),
    },
    "nopywin32": {
        "normal.txt":       ("0000", "0000", "0000"),
        "hidden.txt":       ("1010", "0000", "1010"),
        "system.txt":       ("1100", "0000", "1100"),
        "hidsys.txt":       ("1110", "0000", "1110"),
        "hiddendir":        ("1010", "0000", "1010"),
        "systemdir":        ("1100", "0000", "1100"),
        ".dotfile":         ("0000", "0000", "1111"),
        "thumbs.db":        ("0000", "0000", "1111"),
        "desktop.ini":      ("0000", "0000", "1111"),
        "$recycle.bin":     ("0000", "0000", "1111"),
        "AppData":          ("1110", "0000", "1110"),
        "Application Data": ("1110", "0000", "1110"),
        "Local Settings":   ("1110", "0000", "1110"),
        "junction":         ("1111", "0000", "1111"),
        "missing.txt":      ("0000", "0000", "0000"),
    },
}


class _Var:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


def _set_attributes(path, attrs):
    set_attrs = ctypes.windll.kernel32.SetFileAttributesW
    set_attrs.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
    set_attrs.restype = ctypes.c_int
    if not set_attrs(path, attrs):
        raise ctypes.WinError()


def _make_tree(root):
    """Create ENTRIES under root and return {name: absolute path}."""
    os.makedirs(root, exist_ok=True)
    target = os.path.join(root, "_junction_target")
    os.makedirs(target)
    paths = {}
    for name, kind, attrs in ENTRIES:
        path = os.path.join(root, name)
        if kind == "file":
            with open(path, "w", encoding="utf-8") as f:
                f.write("x")
        elif kind == "dir":
            os.makedirs(path)
        elif kind == "junction":
            user_tree.make_junction(path, target)
        if attrs:
            _set_attributes(path, attrs)
        paths[name] = path
    return paths


def _headless(mod, path, include_hidden, include_system):
    """Headless names/attributes judgment (filters separate), as the engine does it.

    Uses the module's own ezbak.core.is_skipped / BackupPolicy (mod.core_is_skipped is
    ezBAK's alias), so the nopywin32 copy judges with its fresh ezbak import.
    """
    flags = SimpleNamespace(include_hidden=include_hidden, include_system=include_system)
    return mod.core_is_skipped(path, flags, mod.BackupPolicy.HEADLESS,
                               check_filters=False) is not None


def _judge(mod, path, include_hidden, include_system):
    fake_app = SimpleNamespace(
        filters={"include": [], "exclude": []},
        hidden_mode_var=_Var("include" if include_hidden else "exclude"),
        system_mode_var=_Var("include" if include_system else "exclude"),
    )
    return (mod.App.is_hidden(fake_app, path),
            mod.App.is_hidden_safe(fake_app, path),
            _headless(mod, path, include_hidden, include_system))


def _table(mod, paths):
    out = {}
    for name, _kind, _attrs in ENTRIES:
        per_impl = [[], [], []]
        for inc_h, inc_s in POLICIES:
            for i, value in enumerate(_judge(mod, paths[name], inc_h, inc_s)):
                per_impl[i].append("1" if value else "0")
        out[name] = tuple("".join(bits) for bits in per_impl)
    return out


@pytest.fixture
def tree(tmp_path):
    return _make_tree(str(tmp_path / "user"))


def test_hidden_golden_pywin32(ezbak_mod, isolated_env, tree):
    if not ezbak_mod._have_pywin32:
        pytest.skip("pywin32 not installed")
    assert _table(ezbak_mod, tree) == EXPECTED["pywin32"]


def test_hidden_golden_nopywin32(ezbak_mod_nopywin32, isolated_env, tree):
    assert _table(ezbak_mod_nopywin32, tree) == EXPECTED["nopywin32"]


def test_attributes_were_applied(tree):
    """Guard: the fixture really produced hidden/system/reparse entries."""
    get_attrs = ctypes.windll.kernel32.GetFileAttributesW
    get_attrs.argtypes = [ctypes.c_wchar_p]
    get_attrs.restype = ctypes.c_uint32
    assert get_attrs(tree["hidden.txt"]) & H
    assert get_attrs(tree["system.txt"]) & S
    assert get_attrs(tree["hidsys.txt"]) & (H | S) == (H | S)
    assert get_attrs(tree["junction"]) & 0x400  # FILE_ATTRIBUTE_REPARSE_POINT


@pytest.mark.gui
def test_real_app_default_modes_and_toggle(ezbak_mod, isolated_env, gui_app, block_subprocess,
                                           tmp_path):
    """Moved from the old root test_is_hidden.py: a real App starts with both modes
    'exclude', and its Tk variables drive App.is_hidden on a hidden+system file."""
    path = str(tmp_path / "hidsys.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write("x")
    _set_attributes(path, H | S)

    app = gui_app(ezbak_mod)
    modes = [(app.hidden_mode_var.get(), app.system_mode_var.get())]
    results = [app.is_hidden(path)]
    app.hidden_mode_var.set("include")
    modes.append((app.hidden_mode_var.get(), app.system_mode_var.get()))
    results.append(app.is_hidden(path))
    app.system_mode_var.set("include")
    modes.append((app.hidden_mode_var.get(), app.system_mode_var.get()))
    results.append(app.is_hidden(path))

    assert modes == [("exclude", "exclude"), ("include", "exclude"), ("include", "include")]
    assert results == [True, True, False]
    assert block_subprocess == []
