"""Unit tests for ezbak.core.filters (is_skipped / matches_filter).

The input tables and expected values are those of the characterization tests
(tests/characterization/test_filters.py and test_hidden.py), run here directly
against core with the three policies:

- gui:      replace(BackupPolicy.GUI, skip_appdata_names=True, attr_ctypes_fallback=True)
            (old App.is_hidden)
- safe:     BackupPolicy.GUI       (old App.is_hidden_safe)
- headless: BackupPolicy.HEADLESS  (old core_run_backup.is_hidden_headless / matches_filter)

"no pywin32" is simulated by monkeypatching ezbak.winsys.fileattrs.HAS_PYWIN32,
which fileattrs reads on every call.
"""
import copy
import dataclasses
import importlib.util
import os
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

from ezbak.core import (DEFAULT_FILTERS, SKIP_FILTERED, SKIP_HIDDEN, BackupPolicy, FilterSet,
                        is_skipped, matches_filter)
from ezbak.winsys import fileattrs

pytestmark = pytest.mark.windows

_CHAR_DIR = Path(__file__).resolve().parent / "characterization"


def _load(name):
    key = f"ezbak_test_core_filters_{name}"
    if key not in sys.modules:
        spec = importlib.util.spec_from_file_location(key, _CHAR_DIR / f"{name}.py")
        sys.modules[key] = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(sys.modules[key])
    return sys.modules[key]


char_filters = _load("test_filters")
char_hidden = _load("test_hidden")

IS_HIDDEN_POLICY = dataclasses.replace(BackupPolicy.GUI, skip_appdata_names=True,
                                       attr_ctypes_fallback=True)


def _flags(include_hidden, include_system, filters=None):
    return SimpleNamespace(include_hidden=include_hidden, include_system=include_system,
                           filters=filters)


def _three(path, include_hidden, include_system, filters, is_dir=None, headless_filters=True):
    """(gui, safe, headless) as booleans; headless covers names/attributes + filters."""
    flags = _flags(include_hidden, include_system, filters)
    gui = is_skipped(path, flags, IS_HIDDEN_POLICY) is not None
    safe = is_skipped(path, flags, BackupPolicy.GUI) is not None
    headless = is_skipped(path, flags, BackupPolicy.HEADLESS, is_dir=is_dir,
                          check_filters=headless_filters) is not None
    return gui, safe, headless


# --- filter table -----------------------------------------------------------

@pytest.mark.parametrize("case_id,filters,rel,is_dir,expected", char_filters.CASES,
                         ids=[c[0] for c in char_filters.CASES])
def test_filter_table(tmp_path, case_id, filters, rel, is_dir, expected):
    path = char_filters._make_entry(tmp_path / "src", rel, is_dir)
    if filters == char_filters.DEFAULTS:
        filters = copy.deepcopy(DEFAULT_FILTERS)
    flags = _flags(True, True, filters)
    gui = is_skipped(path, flags, IS_HIDDEN_POLICY) is not None
    safe = is_skipped(path, flags, BackupPolicy.GUI) is not None
    # core_run_backup replaced filters=None with the defaults before matches_filter.
    headless = matches_filter(path, filters if filters is not None else DEFAULT_FILTERS, is_dir)
    assert (gui, safe, headless) == expected


@pytest.mark.parametrize("case_id,filters,rel,is_dir,expected", char_filters.CASES,
                         ids=[c[0] for c in char_filters.CASES])
def test_filter_table_headless_policy_none_default(tmp_path, case_id, filters, rel, is_dir,
                                                   expected):
    """is_skipped(HEADLESS) applies DEFAULT_FILTERS itself when filters is None."""
    path = char_filters._make_entry(tmp_path / "src", rel, is_dir)
    if filters == char_filters.DEFAULTS:
        filters = copy.deepcopy(DEFAULT_FILTERS)
    reason = is_skipped(path, _flags(True, True, filters), BackupPolicy.HEADLESS, is_dir=is_dir)
    base = os.path.basename(path).lower()
    builtin = base.startswith('.') or base in ("thumbs.db", "desktop.ini", "$recycle.bin")
    if builtin:
        assert reason == SKIP_HIDDEN
    else:
        assert (reason == SKIP_FILTERED) == expected[2]


def test_matches_filter_ignores_existence(tmp_path):
    path = str(tmp_path / "missing" / "a.txt")
    filters = char_filters.rules(char_filters.E, ("ext", "txt"))
    assert matches_filter(path, filters, False) is True
    flags = _flags(True, True, filters)
    assert is_skipped(path, flags, BackupPolicy.HEADLESS) == SKIP_FILTERED
    assert is_skipped(path, flags, BackupPolicy.GUI) is None
    assert is_skipped(path, flags, IS_HIDDEN_POLICY) is None


def test_matches_filter_accepts_filterset(tmp_path):
    path = char_filters._make_entry(tmp_path, "zqdocs/a.txt", False)
    fs = FilterSet.from_dict(char_filters.rules(char_filters.E, ("ext", "txt")))
    assert matches_filter(path, fs, False) is True
    assert matches_filter(path, None, False) is False


# --- hidden / system / reparse table ----------------------------------------

@pytest.fixture
def tree(tmp_path):
    return char_hidden._make_tree(str(tmp_path / "user"))


def _hidden_table(paths):
    out = {}
    for name, _kind, _attrs in char_hidden.ENTRIES:
        per_impl = [[], [], []]
        for inc_h, inc_s in char_hidden.POLICIES:
            empty = {"include": [], "exclude": []}
            values = _three(paths[name], inc_h, inc_s, empty, headless_filters=False)
            for i, value in enumerate(values):
                per_impl[i].append("1" if value else "0")
        out[name] = tuple("".join(bits) for bits in per_impl)
    return out


def test_hidden_table_pywin32(tree):
    if not fileattrs.HAS_PYWIN32:
        pytest.skip("pywin32 not installed")
    assert _hidden_table(tree) == char_hidden.EXPECTED["pywin32"]


def test_hidden_table_nopywin32(tree, monkeypatch):
    monkeypatch.setattr(fileattrs, "HAS_PYWIN32", False)
    assert _hidden_table(tree) == char_hidden.EXPECTED["nopywin32"]


def test_reason_hidden_wins_over_filter(tree):
    filters = char_filters.rules(char_filters.E, ("name", "hidden.txt"))
    flags = _flags(False, False, filters)
    assert is_skipped(tree["hidden.txt"], flags, BackupPolicy.HEADLESS) == SKIP_HIDDEN
    assert is_skipped(tree["normal.txt"], _flags(False, False, char_filters.rules(
        char_filters.E, ("name", "normal.txt"))), BackupPolicy.HEADLESS) == SKIP_FILTERED
    assert is_skipped(tree["normal.txt"], flags, BackupPolicy.HEADLESS) is None


@pytest.mark.parametrize("has_pywin32", [True, False])
def test_attribute_read_failure_falls_through_to_filters(tree, monkeypatch, has_pywin32):
    """Old ctypes-failure paths: is_hidden read NORMAL and went on to filters; Headless
    returned False from the hidden check and reported filters separately. Same result."""
    def failing(path):
        raise OSError("GetFileAttributesW failed")

    monkeypatch.setattr(fileattrs, "HAS_PYWIN32", has_pywin32)
    monkeypatch.setattr(fileattrs, "get_attributes", failing)
    hidden = tree["hidden.txt"]
    filters = char_filters.rules(char_filters.E, ("name", "hidden.txt"))
    flags = _flags(False, False, filters)
    assert is_skipped(hidden, _flags(False, False, {}), IS_HIDDEN_POLICY) is None
    assert is_skipped(hidden, flags, IS_HIDDEN_POLICY) == SKIP_FILTERED
    assert is_skipped(hidden, flags, BackupPolicy.GUI) == SKIP_FILTERED
    assert is_skipped(hidden, flags, BackupPolicy.HEADLESS, check_filters=False) is None
    assert is_skipped(hidden, flags, BackupPolicy.HEADLESS) == SKIP_FILTERED
