"""ezbak.core engine: run_backup step order, policy flags and wording (P2-6)."""
import io
import os
from dataclasses import replace
from datetime import datetime

import pytest

from ezbak.core import (BackupJob, BackupPolicy, BackupResult, InsufficientSpaceError, JobLog,
                       JobLogFormat, LogItem, TargetResult, run_backup)

USER = "tester"
NOW = datetime(2026, 10, 12, 9, 30, 0)
NO_FILTERS = {"include": [], "exclude": []}


def clock():
    return NOW


class FakeReporter:
    def __init__(self):
        self.calls = []

    def log(self, message):
        self.calls.append(("log", message))

    def status(self, message):
        self.calls.append(("status", message))

    def progress(self, value):
        self.calls.append(("progress", value))

    def progress_max(self, value):
        self.calls.append(("progress_max", value))


class FakeStore:
    def __init__(self, free=10 ** 12):
        self.free = free
        self.listed = []

    def list_backups(self, user_name):
        self.listed.append(user_name)
        return []

    def path_of(self, name):
        return name

    def delete_backup(self, name, on_log=None):
        raise AssertionError("nothing to delete")

    def free_space(self):
        return self.free


class FakeTarget:
    name = "fake"

    def __init__(self):
        self.contexts = []

    def backup(self, ctx):
        self.contexts.append(ctx)
        ctx.out(("target ran", "", "\n"))
        return TargetResult(self.name)


class RaisingFiltersJob:
    """Job whose filter lookup fails, so every skip check raises."""

    def __init__(self, dest):
        self.user, self.dest = USER, dest
        self.include_hidden = self.include_system = True
        self.retention_count, self.log_retention_days = 2, 30

    @property
    def filters(self):
        raise RuntimeError("bad filters")


def _write(path, data=b"x"):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(data)


@pytest.fixture
def tree(tmp_path):
    """<tmp>/Users/tester with a plain file, an empty file, a dot file, a filtered
    file (.tmp) and a filtered directory (OneDrive); <tmp>/dest/bk is the prepared folder."""
    home = tmp_path / "Users" / USER
    _write(str(home / "a.txt"), b"hello")
    _write(str(home / "empty.txt"), b"")
    _write(str(home / ".dot"), b"dot")
    _write(str(home / "skip.tmp"), b"tmp")
    _write(str(home / "sub" / "b.txt"), b"world!")
    _write(str(home / "OneDrive" / "c.txt"), b"cloud")
    dest = tmp_path / "dest"
    backup = dest / "bk"
    backup.mkdir(parents=True)
    filters = {"include": [], "exclude": [{"type": "name", "pattern": "onedrive*"},
                                          {"type": "ext", "pattern": "tmp"}]}
    return dict(root=str(tmp_path / "Users"), home=str(home), dest=str(dest), backup=str(backup),
                filters=filters)


def _run(tree, fmt=JobLogFormat.HEADLESS, policy=BackupPolicy.HEADLESS, job=None,
         store=None, targets=(), reporter=None):
    job = job or BackupJob(user=USER, dest=tree["dest"], filters=tree["filters"])
    stream = io.StringIO()
    reporter = reporter or FakeReporter()
    result = run_backup(job, tree["backup"], JobLog(stream, clock), fmt, store or FakeStore(),
                        list(targets), reporter, policy, clock, users_root=tree["root"],
                        log_path="L.log", now_ts=lambda: NOW.timestamp())
    return result, stream.getvalue(), reporter


def test_headless_logs_hidden_and_filtered_separately(tree):
    result, text, _ = _run(tree)
    home = tree["home"]
    assert f"Skipping hidden file: {os.path.join(home, '.dot')}" in text
    assert f"Skipping filtered file: {os.path.join(home, 'skip.tmp')}" in text
    assert f"Skipping filtered directory: {os.path.join(home, 'OneDrive')}" in text
    assert "Processing directory" not in text
    assert os.path.isfile(os.path.join(tree["backup"], "sub", "b.txt"))
    assert not os.path.exists(os.path.join(tree["backup"], "OneDrive"))
    assert result == BackupResult(backup_path=tree["backup"], log_path="L.log",
                                  files=2, folders=2, bytes=11)


def test_gui_reports_filtered_as_hidden(tree):
    _result, text, _ = _run(tree, JobLogFormat.GUI, BackupPolicy.GUI)
    home = tree["home"]
    assert f"Skipping hidden directory: {os.path.join(home, 'OneDrive')}" in text
    assert f"Skipping hidden file: {os.path.join(home, 'skip.tmp')}" in text
    assert "Skipping filtered" not in text
    assert f"Processing directory: {home} (relative: .)" in text
    assert text.rstrip("\n").endswith("Backup successfully completed.")


@pytest.mark.parametrize("attempted,files,size", [(False, 2, "11.00 B"), (True, 3, "11.00 B")])
def test_count_attempted_copies(tree, attempted, files, size):
    policy = replace(BackupPolicy.HEADLESS, count_attempted_copies=attempted)
    result, text, _ = _run(tree, policy=policy)
    assert result.files == files
    assert f"Backup completed. Files backed up: {files}, Folders: 2, Total size: {size}" in text


def test_skip_errors_fail_files_without_flag(tree):
    job = RaisingFiltersJob(tree["dest"])
    result, text, _ = _run(tree, job=job)
    home = tree["home"]
    assert f"Error checking directory {os.path.join(home, 'sub')}: bad filters" in text
    assert f"ERROR: Failed to copy file {os.path.join(home, 'a.txt')}: bad filters" in text
    assert result.files == 0
    assert "Calculating total size... 0.00 B" in text


def test_skip_errors_include_with_flag(tree):
    policy = replace(BackupPolicy.HEADLESS, skip_errors_include=True)
    result, text, _ = _run(tree, policy=policy, job=RaisingFiltersJob(tree["dest"]))
    assert "Error checking directory" not in text and "Failed to copy file" not in text
    assert result.files == 4  # all but the dot file (name rule) and the empty file
    assert os.path.isfile(os.path.join(tree["backup"], "OneDrive", "c.txt"))


def test_critical_error_is_logged_and_raised(tree):
    fmt = replace(JobLogFormat.HEADLESS, completed=LogItem("{missing}"))
    stream = io.StringIO()
    job = BackupJob(user=USER, dest=tree["dest"], filters=tree["filters"])
    with pytest.raises(KeyError):
        run_backup(job, tree["backup"], JobLog(stream, clock), fmt, FakeStore(), [],
                   FakeReporter(), BackupPolicy.HEADLESS, clock, users_root=tree["root"])
    assert "CRITICAL ERROR in backup loop: 'missing'" in stream.getvalue()


def test_targets_and_retention_run_after_copy(tree):
    target, store = FakeTarget(), FakeStore()
    _result, text, _ = _run(tree, targets=[target], store=store)
    ctx = target.contexts[0]
    assert (ctx.backup_path, ctx.dest, ctx.fmt, ctx.policy) == (
        tree["backup"], tree["dest"], JobLogFormat.HEADLESS, BackupPolicy.HEADLESS)
    assert store.listed == [USER]
    order = [text.index(s) for s in ("Calculating total size", "Copying file", "Backup completed",
                                     "target ran", "[Old backup cleanup]", "[Old log cleanup]")]
    assert order == sorted(order)
    assert text.startswith(f"[{NOW:%Y-%m-%d %H:%M:%S}] Calculating total size... 11.00 B")


def test_headless_copy_error_goes_to_reporter_only(tree):
    os.makedirs(os.path.join(tree["backup"], "a.txt"))  # destination is a folder: copy fails
    result, text, reporter = _run(tree)
    logs = [m for kind, m in reporter.calls if kind == "log"]
    assert len(logs) == 1 and logs[0].startswith("ERROR copying file ")
    assert "ERROR: copying file" not in text
    assert result.files == 1


def test_gui_copy_error_goes_to_job_log(tree):
    os.makedirs(os.path.join(tree["backup"], "a.txt"))
    _result, text, reporter = _run(tree, JobLogFormat.GUI, BackupPolicy.GUI)
    assert "ERROR: copying file " in text
    assert ("status", "Starting backup...") in reporter.calls


def test_headless_reporter_gets_no_status(tree):
    _result, _text, reporter = _run(tree)
    assert {kind for kind, _m in reporter.calls} <= {"progress", "progress_max"}


def test_free_space_check(tree):
    stream = io.StringIO()
    job = BackupJob(user=USER, dest=tree["dest"], filters=NO_FILTERS)
    with pytest.raises(InsufficientSpaceError) as info:
        run_backup(job, tree["backup"], JobLog(stream, clock), JobLogFormat.GUI, FakeStore(free=1),
                   [], FakeReporter(), BackupPolicy.GUI, clock, users_root=tree["root"])
    assert info.value.free == 1
    assert "Insufficient free space for backup. Required=" in stream.getvalue()
    assert "Copying file" not in stream.getvalue()


# --- GUI policy path (P2-7) ---

class RaisingRetentionJob:
    """Job whose retention count (or log days) lookup fails at cleanup time."""

    def __init__(self, tree, field="retention_count"):
        self.user, self.dest, self.filters = USER, tree["dest"], tree["filters"]
        self.include_hidden = self.include_system = False
        self._field = field

    def _value(self, field, value):
        if field == self._field:
            raise RuntimeError(f"bad {field}")
        return value

    @property
    def retention_count(self):
        return self._value("retention_count", 2)

    @property
    def log_retention_days(self):
        return self._value("log_retention_days", 30)


def test_gui_skips_backup_cleanup_when_count_unreadable(tree):
    store = FakeStore()
    result, text, _ = _run(tree, JobLogFormat.GUI, BackupPolicy.GUI, store=store,
                           job=RaisingRetentionJob(tree))
    assert "ERROR: An error occurred during backup cleanup: bad retention_count" in text
    assert store.listed == [] and "[Old backup cleanup]" not in text
    assert "[Old log cleanup]" in text and "Cleanup failed" not in text
    assert text.rstrip("\n").endswith("Backup successfully completed.")
    assert result.critical_error is None


def test_gui_cleanup_failed_when_log_days_unreadable(tree):
    store = FakeStore()
    _result, text, _ = _run(tree, JobLogFormat.GUI, BackupPolicy.GUI, store=store,
                            job=RaisingRetentionJob(tree, "log_retention_days"))
    assert store.listed == [USER]  # backup cleanup ran first, as before
    assert "Cleanup failed: bad log_retention_days" in text
    assert "[Old log cleanup]" not in text
    assert text.rstrip("\n").endswith("Backup successfully completed.")


@pytest.mark.parametrize("field", ["retention_count", "log_retention_days"])
def test_headless_lets_retention_errors_propagate(tree, field):
    with pytest.raises(RuntimeError, match=f"bad {field}"):
        _run(tree, job=RaisingRetentionJob(tree, field))


def test_gui_returns_critical_error(tree):
    fmt = replace(JobLogFormat.GUI, completed=LogItem("{missing}"))
    target = FakeTarget()
    result, text, _ = _run(tree, fmt, BackupPolicy.GUI, targets=[target])
    assert isinstance(result.critical_error, KeyError)
    assert "CRITICAL ERROR in backup loop: 'missing'" in text
    assert target.contexts == [] and "Backup successfully completed." not in text


def test_gui_insufficient_space_raises_before_copy(tree):
    stream = io.StringIO()
    job = BackupJob(user=USER, dest=tree["dest"], filters=tree["filters"])
    reporter = FakeReporter()
    with pytest.raises(InsufficientSpaceError) as info:
        run_backup(job, tree["backup"], JobLog(stream, clock), JobLogFormat.GUI, FakeStore(free=3),
                   [], reporter, BackupPolicy.GUI, clock, users_root=tree["root"])
    assert (info.value.required, info.value.free) == (14, 3)
    assert ("status", "Checking available space...") in reporter.calls
    assert ("status", "Starting backup...") not in reporter.calls
    assert "Insufficient free space for backup. Required=14 Free=3" in stream.getvalue()


def test_gui_checks_space_after_size_error(tree, monkeypatch):
    real_walk, calls = os.walk, []

    def walk(*args, **kwargs):
        calls.append(args[0])
        if len(calls) == 1:  # the size step's walk fails at once
            raise OSError("size walk failed")
        return real_walk(*args, **kwargs)

    monkeypatch.setattr(os, "walk", walk)
    stream = io.StringIO()
    job = BackupJob(user=USER, dest=tree["dest"], filters=tree["filters"])
    with pytest.raises(InsufficientSpaceError):
        run_backup(job, tree["backup"], JobLog(stream, clock), JobLogFormat.GUI, FakeStore(free=-1),
                   [], FakeReporter(), BackupPolicy.GUI, clock, users_root=tree["root"])
    text = stream.getvalue()
    assert "Total bytes calculation error: size walk failed" in text
    assert "Calculating total size... 0.00 B" in text


def test_gui_progress_max_and_final_progress(tree):
    _result, _text, reporter = _run(tree, JobLogFormat.GUI, BackupPolicy.GUI)
    progress = [v for kind, v in reporter.calls if kind == "progress"]
    maxes = [v for kind, v in reporter.calls if kind == "progress_max"]
    assert maxes == [14]  # a.txt 5 + .dot 3 + sub/b.txt 6 (GUI keeps dot files)
    assert progress == sorted(progress) and progress[-1] == progress[-2] == 14
    first_progress = next(i for i, c in enumerate(reporter.calls) if c[0] == "progress")
    assert reporter.calls.index(("progress_max", 14)) < first_progress
