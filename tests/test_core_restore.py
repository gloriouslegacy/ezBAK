"""ezbak.core restore: selection check, backup location (P3-2) and run_restore (P3-3)."""
import os

import pytest

from ezbak.core import (KIND_DIRECT, KIND_PARENT, RestorePlan, RestoreSelectionError,
                        locate_backup, run_restore, validate_selection)


def _never(candidates):
    raise AssertionError(f"choose_backup must not be called: {candidates}")


def _backup(root, name, with_user_data=True):
    path = root / name
    path.mkdir()
    if with_user_data:
        (path / "Documents").mkdir()
    return path


# --- validate_selection ---------------------------------------------------

def test_validate_log_file_message(tmp_path):
    p = str(tmp_path / "x_backup_1_restore.log")
    with pytest.raises(RestoreSelectionError) as ei:
        validate_selection(p)
    assert ei.value.message == (
        f"The selected path is a log file.\n\n'{p}'\n\nYou must select a backup folder")
    assert str(ei.value) == ei.value.message


def test_validate_log_suffix_is_case_insensitive_and_checked_first(tmp_path):
    p = tmp_path / "dir.LOG"
    p.mkdir()
    with pytest.raises(RestoreSelectionError, match="is a log file"):
        validate_selection(str(p))


def test_validate_file_message(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("x")
    with pytest.raises(RestoreSelectionError) as ei:
        validate_selection(str(p))
    assert ei.value.message == f"The selected path is a file. Please select a folder:\n{p}"


def test_validate_missing_message(tmp_path):
    p = str(tmp_path / "nope")
    with pytest.raises(RestoreSelectionError) as ei:
        validate_selection(p)
    assert ei.value.message == f"Backup folder not found:\n{p}"


def test_validate_folder_passes(tmp_path):
    assert validate_selection(str(tmp_path)) is None


# --- locate_backup: direct selection ---------------------------------------

def test_direct_selection(tmp_path):
    b = _backup(tmp_path, "alice_backup_2025-01-01_000000")
    plan = locate_backup(str(b), "bob", _never)
    assert plan == RestorePlan(KIND_DIRECT, str(tmp_path), b.name, "alice")
    assert plan.source_dir == str(b)


@pytest.mark.parametrize("folder", ["Desktop", "Documents", "Downloads", "Pictures", "Videos", "Music"])
def test_direct_any_user_folder_counts(tmp_path, folder):
    b = _backup(tmp_path, "u_backup_1", with_user_data=False)
    (b / folder).mkdir()
    assert locate_backup(str(b), "u", _never).kind == KIND_DIRECT


def test_direct_without_user_folders(tmp_path):
    b = _backup(tmp_path, "u_backup_1", with_user_data=False)
    (b / "AppData").mkdir()
    with pytest.raises(RestoreSelectionError) as ei:
        locate_backup(str(b), "u", _never)
    assert ei.value.message == (
        f"The selected folder appears to be named like a backup but does not contain user data:\n{b}\n\n"
        f"Please select either:\n"
        f"1. A parent folder containing backup folders, or\n"
        f"2. A valid backup folder with user data")


# --- locate_backup: parent selection ---------------------------------------

def test_parent_single_backup_needs_no_choice(tmp_path):
    tmp_path = tmp_path / "root"  # test name holds '_backup_'; keep the parent name neutral
    tmp_path.mkdir()
    _backup(tmp_path, "carol_backup_1")
    (tmp_path / "other").mkdir()
    (tmp_path / "x_backup_file.txt").write_text("not a folder")
    plan = locate_backup(str(tmp_path), "dave", _never)
    assert plan == RestorePlan(KIND_PARENT, str(tmp_path), "carol_backup_1", "carol")


def test_parent_choice_is_user_first_then_newest_first(tmp_path):
    for name in ("bob_backup_2025-03-01_090000", "alice_backup_2025-01-01_090000",
                 "zed_backup_2025-02-01_090000", "alice_backup_2025-04-01_090000"):
        _backup(tmp_path, name, with_user_data=False)
    seen = []

    def choose(candidates):
        seen.append(list(candidates))
        return "alice_backup_2025-01-01_090000"

    plan = locate_backup(str(tmp_path), "alice", choose)
    assert seen == [["alice_backup_2025-04-01_090000", "alice_backup_2025-01-01_090000",
                     "zed_backup_2025-02-01_090000", "bob_backup_2025-03-01_090000"]]
    assert plan == RestorePlan(KIND_PARENT, str(tmp_path),
                               "alice_backup_2025-01-01_090000", "alice")


def test_parent_choice_ignores_case_of_user_name(tmp_path):
    for name in ("alice_backup_2025-01-01_090000", "Alice_backup_2025-04-01_090000",
                 "ALICE_backup_2025-02-01_090000", "Bob_backup_2025-05-01_090000",
                 "bob_backup_2025-03-01_090000"):
        _backup(tmp_path, name, with_user_data=False)
    seen = []

    def choose(candidates):
        seen.append(list(candidates))
        return candidates[0]

    plan = locate_backup(str(tmp_path), "aLiCe", choose)
    assert seen == [["Alice_backup_2025-04-01_090000", "ALICE_backup_2025-02-01_090000",
                     "alice_backup_2025-01-01_090000",
                     "Bob_backup_2025-05-01_090000", "bob_backup_2025-03-01_090000"]]
    assert plan.folder_name == "Alice_backup_2025-04-01_090000"


@pytest.mark.parametrize("answer", [None, ""])
def test_parent_choice_cancelled(tmp_path, answer):
    _backup(tmp_path, "a_backup_1")
    _backup(tmp_path, "a_backup_2")
    assert locate_backup(str(tmp_path), "a", lambda c: answer) is None


def test_parent_without_backups(tmp_path):
    (tmp_path / "plain").mkdir()
    with pytest.raises(RestoreSelectionError) as ei:
        locate_backup(str(tmp_path), "a", _never)
    assert ei.value.message == (
        f"No backup folders found in:\n{tmp_path}\n\n"
        f"Please select either:\n"
        f"1. A parent folder containing backup folders (e.g., E:\\), or\n"
        f"2. A backup folder directly (e.g., test_backup_2025-12-04_091400)")


def test_parent_listdir_error(tmp_path, monkeypatch):
    def boom(path):
        raise PermissionError("denied")

    monkeypatch.setattr(os, "listdir", boom)
    with pytest.raises(RestoreSelectionError) as ei:
        locate_backup(str(tmp_path), "a", _never)
    assert ei.value.message == "Error reading backup folder:\ndenied"
    assert isinstance(ei.value.__cause__, PermissionError)


# --- run_restore ----------------------------------------------------------

class _Reporter:
    def __init__(self):
        self.calls = []

    def log(self, message):
        self.calls.append(('log', message))

    def status(self, message):
        self.calls.append(('status', message))

    def progress(self, value):
        self.calls.append(('progress', value))

    def progress_max(self, value):
        self.calls.append(('progress_max', value))


def _no_skip(path, is_dir=None):
    return False


def _plan(root, name, files=None):
    src = root / name
    src.mkdir()
    for rel, data in (files or {}).items():
        f = src / rel
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_bytes(data)
    return RestorePlan(KIND_PARENT, str(root), name, name.split("_backup_")[0])


def _run(plan, user, dest, *, skip=_no_skip, free=10 ** 12, confirm=None):
    lines, rep = [], _Reporter()
    confirm = confirm or (lambda *a: (_ for _ in ()).throw(AssertionError("no confirm")))
    result = run_restore(plan, user, str(dest), lines.append, rep, skip=skip,
                         free_space=lambda path: free, confirm_cross_user=confirm)
    return result, lines, rep.calls


def test_run_restore_cross_user_declined(tmp_path):
    plan = _plan(tmp_path, "bob_backup_1", {"a.txt": b"x"})
    asked = []
    result, lines, calls = _run(plan, "alice", tmp_path / "dest",
                                confirm=lambda *a: asked.append(a) or False)
    assert result.cancelled and result.files == 0
    assert asked == [("bob", "alice", "bob_backup_1")]
    assert lines == ["Cross-user restore detected: bob → alice",
                     "Cross-user restore cancelled by user"]
    assert calls == []
    assert not (tmp_path / "dest").exists()


def test_run_restore_copies_and_counts(tmp_path):
    plan = _plan(tmp_path, "Alice_backup_1",
                 {"a.txt": b"abc", "sub/b.txt": b"12345", "sub/h.txt": b"zz"})
    dest = tmp_path / "dest"
    result, lines, calls = _run(plan, "alice", dest,
                                skip=lambda p, is_dir=None: p.endswith("h.txt"))
    assert (result.cancelled, result.error, result.files, result.bytes, result.total) == (
        False, None, 2, 8, 8)
    assert (dest / "sub" / "b.txt").read_bytes() == b"12345"
    assert not (dest / "sub" / "h.txt").exists()
    assert calls[0] == ('progress_max', 8) and calls[-1] == ('progress', 8)
    assert lines[0] == f"Restoring from {plan.source_dir} to {dest}"
    assert lines[1] == "Total bytes to restore: 8"
    assert any(l.startswith("Skipping hidden file: ") for l in lines)
    assert lines[-2:] == ["Restore completed. Files restored: 2, Total size: 8.00 B",
                          "Restore successfully completed."]


def test_run_restore_size_error_falls_back_to_100mb(tmp_path):
    plan = _plan(tmp_path, "a_backup_1", {"d/x.txt": b"x"})

    def skip(path, is_dir=None):
        raise RuntimeError("attr")

    result, lines, calls = _run(plan, "a", tmp_path / "dest", skip=skip, free=0)
    assert result.total == 100 * 1024 * 1024
    assert "Total size calculation error: attr" in lines
    assert calls == [('progress_max', 100 * 1024 * 1024)]


def test_run_restore_not_enough_space(tmp_path):
    plan = _plan(tmp_path, "a_backup_1", {"x.bin": b"0" * 2048})
    result, lines, calls = _run(plan, "a", tmp_path / "dest", free=1024)
    assert result.error == ("Not enough free space to restore.\nRequired: 2.00 KB"
                            "\nAvailable: 1.00 KB")
    assert lines[-1] == "Insufficient free space. Required=2048, Free=1024"
    assert result.files == 0 and not (tmp_path / "dest").exists()


def test_run_restore_free_space_error_propagates(tmp_path):
    plan = _plan(tmp_path, "a_backup_1", {"x.txt": b"x"})

    def free_space(path):
        raise AttributeError("no get_free_space")

    with pytest.raises(AttributeError):
        run_restore(plan, "a", str(tmp_path / "dest"), [].append, _Reporter(),
                    skip=_no_skip, free_space=free_space, confirm_cross_user=None)


def test_run_restore_progress_message_every_50_files(tmp_path):
    plan = _plan(tmp_path, "a_backup_1", {f"f{i:03}.txt": b"x" for i in range(101)})
    result, lines, calls = _run(plan, "a", tmp_path / "dest")
    assert result.files == 101
    assert [c for c in calls if c[0] == 'log'] == [('log', "Restored 50 files..."),
                                                   ('log', "Restored 100 files...")]
    assert "50 files restored so far; bytes_copied=50" in lines
    assert "100 files restored so far; bytes_copied=100" in lines
