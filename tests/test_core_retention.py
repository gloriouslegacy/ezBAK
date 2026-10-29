"""ezbak.core retention: old backup / old log cleanup (P2-4)."""
import os
import time
from datetime import datetime

import pytest

from ezbak.core import (BackupPolicy, JobLine, JobLogFormat, cleanup_old_backups,
                        cleanup_old_logs)
from ezbak.storage import LocalFolderStore

DAY = 86400
GUI, HEADLESS = JobLogFormat.GUI, JobLogFormat.HEADLESS


def render(lines):
    return "".join(f"{ln.lead}[TS] {ln.message}{ln.end}" for ln in lines)


class FakeStore:
    def __init__(self, names, fail=(), messages=None):
        self.names = list(names)
        self.fail = set(fail)
        self.messages = messages or {}
        self.deleted = []

    def list_backups(self, user_name):
        return sorted(n for n in self.names if n.lower().startswith(f"{user_name.lower()}_backup_"))

    def path_of(self, name):
        return f"D:\\dest\\{name}"

    def delete_backup(self, name, on_log=None):
        for msg in self.messages.get(name, []):
            on_log(msg)
        if name in self.fail:
            raise OSError("locked")
        self.deleted.append(name)


NAMES = [f"u_backup_2026-01-0{i}_000000" for i in range(1, 5)]


def test_backups_gui_wording():
    out, store = [], FakeStore(NAMES)
    cleanup_old_backups(store, "u", 2, GUI, BackupPolicy.GUI, out.append)
    assert store.deleted == NAMES[:2]
    assert render(out) == (
        "[TS] \n[Old backup cleanup]\n"
        "[TS] Found 4 backups. Keeping the 2 most recent.\n"
        f"[TS]   Deleting old backup: D:\\dest\\{NAMES[0]}\n"
        f"[TS]   Deleting old backup: D:\\dest\\{NAMES[1]}\n"
        "[TS] Cleanup finished.\n")


def test_backups_headless_wording_plain_lines():
    out = []
    cleanup_old_backups(FakeStore(NAMES[:1]), "u", 2, HEADLESS, BackupPolicy.HEADLESS, out.append)
    assert render(out) == (
        "[TS] [Old backup cleanup]\n"
        "[TS] Found 1 backups. No old backups to clean up (retention policy: keep 2).\n")


@pytest.mark.parametrize("fmt, policy, expected", [
    (GUI, BackupPolicy.GUI, "[TS] \n[Old backup cleanup]\n"),
    (HEADLESS, BackupPolicy.HEADLESS, "[TS] [Old backup cleanup]\n"),
])
def test_backups_disabled(fmt, policy, expected):
    out, store = [], FakeStore(NAMES)
    cleanup_old_backups(store, "u", 0, fmt, policy, out.append)
    assert render(out) == expected + "[TS] Backup cleanup is disabled (retention count is 0 or less).\n"
    assert store.deleted == []


def test_backups_delete_messages_go_to_out():
    msg = "ERROR: Deletion failed for x: boom"
    out = []
    cleanup_old_backups(FakeStore(NAMES[:3], messages={NAMES[0]: [msg]}), "u", 2, GUI, BackupPolicy.GUI, out.append)
    assert JobLine(msg) in out


def test_backups_delete_failure_headless_continues():
    out, store = [], FakeStore(NAMES, fail={NAMES[0]})
    cleanup_old_backups(store, "u", 2, HEADLESS, BackupPolicy.HEADLESS, out.append)
    text = render(out)
    assert f"[TS]   ERROR: Could not delete D:\\dest\\{NAMES[0]}: locked\n" in text
    assert store.deleted == [NAMES[1]]
    assert text.endswith("[TS] Cleanup finished.\n")


def test_backups_delete_failure_gui_stops():
    out, store = [], FakeStore(NAMES, fail={NAMES[0]})
    cleanup_old_backups(store, "u", 2, GUI, BackupPolicy.GUI, out.append)
    assert store.deleted == []
    assert render(out).endswith("[TS] ERROR: An error occurred during backup cleanup: locked\n")


def test_backups_delete_isolation_comes_from_policy_not_format():
    out, store = [], FakeStore(NAMES, fail={NAMES[0]})
    cleanup_old_backups(store, "u", 2, HEADLESS, BackupPolicy.GUI, out.append)
    assert store.deleted == []
    assert "Could not delete" not in render(out)


def test_backups_other_user_ignored(tmp_path):
    for name in ["Alice_backup_2026-01-01_000000", "alice_backup_2026-01-02_000000",
                 "alice_backup_2026-01-03_000000", "bob_backup_2026-01-01_000000"]:
        (tmp_path / name).mkdir()
    (tmp_path / "alice_backup_2026-01-00_000000.log").write_text("x")
    out = []
    cleanup_old_backups(LocalFolderStore(str(tmp_path)), "alice", 2, GUI, BackupPolicy.GUI, out.append)
    assert sorted(os.listdir(tmp_path)) == [
        "alice_backup_2026-01-00_000000.log", "alice_backup_2026-01-02_000000",
        "alice_backup_2026-01-03_000000", "bob_backup_2026-01-01_000000"]
    assert JobLine("Found 3 backups. Keeping the 2 most recent.") in out


FIXED = datetime(2026, 10, 12, 9, 0, 0)


def run_logs(folder, days, fmt, policy):
    out = []
    cleanup_old_logs(str(folder), days, fmt, policy, out.append,
                     clock=lambda: FIXED, now_ts=time.time)
    return [ln.message for ln in out], out


def make_logs(tmp_path):
    old = tmp_path / "old.log"
    old.write_text("x")
    t = time.time() - 40 * DAY
    os.utime(old, (t, t))
    (tmp_path / "new.log").write_text("x")
    (tmp_path / "other.txt").write_text("x")
    return old


def test_logs_gui_and_headless_wording(tmp_path):
    make_logs(tmp_path)
    gui, _ = run_logs(tmp_path, 30, GUI, BackupPolicy.GUI)
    assert not (tmp_path / "old.log").exists()
    assert gui[0] == "\n[Old log cleanup]"
    assert gui[2] == "Current time: 2026-10-12 09:00:00"
    assert "Total files in folder: 3" in gui and "Log files found: 2" in gui
    assert "    ✓ Deleted: old.log" in gui
    assert gui[-1] == "Log cleanup completed. Deleted 1 files."

    make_logs(tmp_path)
    hl, lines = run_logs(tmp_path, 30, HEADLESS, BackupPolicy.HEADLESS)
    assert lines[0] == JobLine("[Old log cleanup]")
    assert not any(m.startswith("Total files") for m in hl)
    assert hl[-1] == "Log cleanup completed. Deleted 1 files."


def test_logs_disabled_and_missing_folder(tmp_path):
    gui, _ = run_logs(tmp_path, 0, GUI, BackupPolicy.GUI)
    assert gui == ["Log cleanup disabled (retention days <= 0)"]
    hl, _ = run_logs(tmp_path, 0, HEADLESS, BackupPolicy.HEADLESS)
    assert hl == ["[Old log cleanup]", "Log cleanup disabled (retention days <= 0)"]
    gui, _ = run_logs(tmp_path / "nope", 30, GUI, BackupPolicy.GUI)
    assert gui == [f"Log folder does not exist: {tmp_path / 'nope'}"]


def test_logs_no_log_files(tmp_path):
    (tmp_path / "a.txt").write_text("x")
    gui, _ = run_logs(tmp_path, 30, GUI, BackupPolicy.GUI)
    assert gui[-1] == "No .log files found in the folder."
    hl, _ = run_logs(tmp_path, 30, HEADLESS, BackupPolicy.HEADLESS)
    assert hl[-1] == "No old log files found to delete."


@pytest.mark.parametrize("policy, fmt, retried", [
    (BackupPolicy.GUI, GUI, True),
    (BackupPolicy.HEADLESS, HEADLESS, False),
])
def test_logs_retry_pass_flag(tmp_path, monkeypatch, policy, fmt, retried):
    old = make_logs(tmp_path)
    real_remove, calls = os.remove, []

    def flaky_remove(path):
        calls.append(path)
        if len(calls) == 1:
            raise PermissionError("busy")
        real_remove(path)

    monkeypatch.setattr(os, "remove", flaky_remove)
    msgs, _ = run_logs(tmp_path, 30, fmt, policy)
    assert "    ✗ Delete failed: busy" in msgs
    assert ("  ✓ Deleted: old.log" in msgs) is retried
    assert old.exists() is not retried
    assert msgs[-1] == ("Log cleanup completed. Deleted 1 files." if retried
                        else "No old log files found to delete.")
