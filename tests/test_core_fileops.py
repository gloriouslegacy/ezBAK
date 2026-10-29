"""Unit tests for ezbak.core fileops (copy with progress, size total, read-only rmtree, multi-source copy)."""
import dataclasses
import os
import stat
import sys
from datetime import datetime, timedelta
from types import SimpleNamespace

import pytest

from ezbak.core import (BackupPolicy, compute_total_size, copy_file_with_progress, copy_sources,
                        handle_readonly_error, is_skipped, rmtree_readonly)


def _write(path, size):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"x" * size)
    return path


def _stepping_clock(step_seconds):
    """Clock that advances `step_seconds` per call (no real waiting)."""
    state = {"now": datetime(2025, 1, 1)}

    def clock():
        current = state["now"]
        state["now"] = current + timedelta(seconds=step_seconds)
        return current
    return clock


# --- copy_file_with_progress -------------------------------------------------

def test_copy_reports_progress_per_chunk(tmp_path):
    src = _write(tmp_path / "src.bin", 10)
    dst = tmp_path / "out" / "nested" / "dst.bin"
    deltas = []
    copied = copy_file_with_progress(str(src), str(dst), deltas.append, buffer_size=4)
    assert copied == 10
    assert deltas == [4, 4, 2]
    assert dst.read_bytes() == src.read_bytes()


def test_copy_error_goes_to_callback_and_returns_zero(tmp_path):
    errors = []
    copied = copy_file_with_progress(str(tmp_path / "missing"), str(tmp_path / "d"), on_error=errors.append)
    assert copied == 0
    assert len(errors) == 1 and isinstance(errors[0], OSError)


def test_copy_timeout_uses_injected_clock(tmp_path):
    src = _write(tmp_path / "big.bin", 12)
    dst = tmp_path / "dst.bin"
    timeouts, errors, deltas = [], [], []
    # start=0s, checks at 20s (ok, one chunk), 40s (> 30s timeout)
    copied = copy_file_with_progress(str(src), str(dst), deltas.append, buffer_size=4, timeout_seconds=30,
                                     clock=_stepping_clock(20), on_timeout=timeouts.append,
                                     on_error=errors.append)
    assert copied == 0
    assert deltas == [4]
    assert timeouts == [str(src)]
    assert isinstance(errors[0], TimeoutError)
    assert str(errors[0]) == f"File copy timeout: {src}"


def test_copy_without_timeout_ignores_clock(tmp_path):
    src = _write(tmp_path / "a.bin", 8)

    def clock():
        raise AssertionError("clock must not be read without timeout")
    assert copy_file_with_progress(str(src), str(tmp_path / "b.bin"), clock=clock) == 8


# --- compute_total_size -------------------------------------------------------

@pytest.fixture
def size_tree(tmp_path):
    root = tmp_path / "home"
    _write(root / "a.txt", 10)
    _write(root / "sub" / "b.txt", 20)
    _write(root / ".cfg", 5)
    _write(root / "thumbs.db", 7)
    _write(root / "AppData" / "x.txt", 11)
    _write(root / "c.tmp", 3)
    return root


_JOB = SimpleNamespace(include_hidden=False, include_system=False,
                       filters={"include": [], "exclude": [{"type": "ext", "pattern": "tmp"}]})
_POLICIES = {
    "gui_safe": (BackupPolicy.GUI, (53, 5)),
    "gui_is_hidden": (dataclasses.replace(BackupPolicy.GUI, skip_appdata_names=True,
                                          attr_ctypes_fallback=True), (42, 4)),
    "headless": (BackupPolicy.HEADLESS, (30, 2)),
}


@pytest.mark.parametrize("name", sorted(_POLICIES))
def test_total_size_reflects_policy_skips(size_tree, name):
    policy, expected = _POLICIES[name]

    def skip(path, is_dir):
        return is_skipped(path, _JOB, policy, is_dir=is_dir) is not None
    assert compute_total_size(str(size_tree), skip) == expected


def test_total_size_progress_every_n(size_tree):
    calls = []
    compute_total_size(str(size_tree), lambda p, d: False, on_progress=lambda n, t: calls.append(n),
                       progress_every=2)
    assert calls == [2, 4, 6]


def test_total_size_check_errors_lenient_vs_strict(size_tree):
    def skip(path, is_dir):
        raise ValueError("boom")
    file_errors = []
    # lenient: directories kept, every file check error reported, nothing counted
    assert compute_total_size(str(size_tree), skip, on_file_error=lambda fp, e: file_errors.append(fp)) == (0, 0)
    assert len(file_errors) == 6
    with pytest.raises(ValueError):
        compute_total_size(str(size_tree), skip, strict=True)
    caught = []
    assert compute_total_size(str(size_tree), skip, strict=True, on_error=caught.append) == (0, 0)
    assert isinstance(caught[0], ValueError)


# --- read-only rmtree ---------------------------------------------------------

@pytest.mark.windows
@pytest.mark.skipif(sys.platform != "win32", reason="read-only blocks deletion only on Windows")
def test_rmtree_readonly_removes_readonly_files(tmp_path):
    root = tmp_path / "old_backup"
    ro = _write(root / "deep" / "ro.txt", 3)
    _write(root / "plain.txt", 3)
    os.chmod(ro, stat.S_IREAD)
    logs = []
    rmtree_readonly(str(root), logs.append)
    assert not root.exists()
    assert logs == [f"INFO: Removed read-only and successfully deleted {ro}"]


def test_handle_readonly_error_messages(tmp_path):
    logs = []
    handle_readonly_error(lambda p: None, "p", (OSError, OSError("disk"), None), logs.append)
    assert logs == ["ERROR: Deletion failed for p: disk"]

    def failing(path):
        raise OSError("still locked")
    logs.clear()
    target = _write(tmp_path / "f.txt", 1)
    handle_readonly_error(failing, str(target), (PermissionError, PermissionError(), None), logs.append)
    assert logs == [f"ERROR: Could not delete {target} even after removing read-only flag: still locked"]


# --- copy_sources ---------------------------------------------------------------

def test_copy_sources_dirs_and_files(tmp_path):
    src_dir = tmp_path / "Docs"
    for i in range(11):
        _write(src_dir / "sub" / f"f{i}.txt", 1)
    _write(src_dir / "skip.tmp", 1)
    single = _write(tmp_path / "single.txt", 2)
    hidden_single = _write(tmp_path / "single.tmp", 2)
    dest = tmp_path / "dest"
    copies, counts = [], []

    def skip(path, is_dir):
        return path.endswith(".tmp")
    done = copy_sources([str(src_dir), str(single), str(hidden_single)], str(dest), skip,
                        on_copy=lambda s, d: copies.append(d), on_count=counts.append)
    assert done == 12
    assert counts == [10]
    assert (dest / "Docs" / "sub" / "f0.txt").exists()
    assert (dest / "single.txt").read_bytes() == b"xx"
    assert not (dest / "Docs" / "skip.tmp").exists()
    assert not (dest / "single.tmp").exists()
    assert len(copies) == 12


def test_copy_sources_uses_given_copier(tmp_path):
    single = _write(tmp_path / "one.txt", 1)
    seen = []
    done = copy_sources([str(single)], str(tmp_path / "d"), lambda p, d: False,
                        copy_one=lambda s, d: seen.append((s, d)))
    assert done == 1
    assert seen == [(str(single), os.path.join(str(tmp_path / "d"), "one.txt"))]
