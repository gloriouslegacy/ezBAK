"""Unit tests of the GUI bridge (ezbak.gui public names, P2-7 carried over to P5-0d).

GuiReporter (queue tuples, progress throttle, final report), GuiLogStream, GuiFolderStore,
GuiJob and int_or, on stand-in App objects: no Tk, no real disk free space.
"""
import logging
import os
import queue
import threading
from datetime import datetime
from types import SimpleNamespace

import pytest

from ezbak.core import diag
from ezbak.gui import GuiFolderStore, GuiJob, GuiLogStream, GuiReporter, begin_job, int_or
from ezbak.storage import LocalFolderStore


class _Var:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class _BrokenVar:
    def get(self):
        raise RuntimeError("var unreadable")


def _drain(q):
    items = []
    while not q.empty():
        items.append(q.get_nowait())
    return items


# --- int_or ---

@pytest.mark.parametrize("value,expected", [
    ("5", 5), (7, 7), (3.9, 3), (" 4 ", 4), ("-1", -1),
    ("abc", 9), ("", 9), (None, 9), ("3.5", 9), ([1], 9),
])
def test_int_or(value, expected):
    assert int_or(value, 9) == expected


def test_int_or_does_not_hide_other_errors():
    class Odd:
        def __int__(self):
            raise RuntimeError("boom")

    with pytest.raises(RuntimeError):
        int_or(Odd(), 1)


# --- GuiJob ---

def _app_vars(count, days):
    return SimpleNamespace(retention_count_var=count, log_retention_days_var=days)


def test_gui_job_forwards_job_fields():
    job = SimpleNamespace(user="tester", dest="D:\\bk", include_hidden=False, filters=None)
    gui_job = GuiJob(_app_vars(_Var("3"), _Var("7")), job)
    assert (gui_job.user, gui_job.dest, gui_job.include_hidden, gui_job.filters) == (
        "tester", "D:\\bk", False, None)
    with pytest.raises(AttributeError):
        gui_job.no_such_field


def test_gui_job_reads_retention_when_asked():
    count, days = _Var("3"), _Var("7")
    gui_job = GuiJob(_app_vars(count, days), SimpleNamespace(retention_count=99, log_retention_days=99))
    assert (gui_job.retention_count, gui_job.log_retention_days) == (3, 7)
    count.value, days.value = "5", "60"  # changed after the job was made
    assert (gui_job.retention_count, gui_job.log_retention_days) == (5, 60)


def test_gui_job_non_number_falls_back():
    gui_job = GuiJob(_app_vars(_Var("abc"), _Var("")), SimpleNamespace())
    assert (gui_job.retention_count, gui_job.log_retention_days) == (2, 30)


def test_gui_job_read_error_propagates():
    gui_job = GuiJob(_app_vars(_BrokenVar(), _BrokenVar()), SimpleNamespace())
    with pytest.raises(RuntimeError, match="var unreadable"):
        gui_job.retention_count
    with pytest.raises(RuntimeError, match="var unreadable"):
        gui_job.log_retention_days


# --- GuiLogStream ---

class _File:
    def __init__(self, lock, fail=False):
        self.lock, self.fail = lock, fail
        self.events = []

    def write(self, text):
        if self.fail:
            raise OSError("disk full")
        self.events.append(("write", text, self.lock.locked()))

    def flush(self):
        self.events.append(("flush", self.lock.locked()))


def _log_app(fail=False, open_file=True):
    lock = threading.Lock()
    return SimpleNamespace(log_lock=lock, _log_file=_File(lock, fail) if open_file else None)


def test_log_stream_writes_and_flushes_each_line_under_lock():
    app = _log_app()
    stream = GuiLogStream(app)
    stream.write("first\n")
    stream.write("second\n")
    stream.flush()
    assert app._log_file.events == [("write", "first\n", True), ("flush", True),
                                    ("write", "second\n", True), ("flush", True)]
    assert not app.log_lock.locked()


def test_log_stream_without_open_file_writes_nothing():
    app = _log_app(open_file=False)
    GuiLogStream(app).write("ignored\n")
    assert app._log_file is None and not app.log_lock.locked()


def test_log_stream_write_error_is_logged_not_raised(caplog):
    app = _log_app(fail=True)
    with caplog.at_level(logging.DEBUG):
        GuiLogStream(app).write("secret line\n")
    messages = [r.getMessage() for r in caplog.records]
    assert messages == ["event=job_log_write_failed error=OSError"]
    assert not app.log_lock.locked()


def test_log_stream_reads_the_current_file():
    app = _log_app(open_file=False)
    stream = GuiLogStream(app)
    app._log_file = _File(app.log_lock)  # opened after the stream was made
    stream.write("late\n")
    assert app._log_file.events[0] == ("write", "late\n", True)


# --- GuiReporter ---

class _Clock:
    def __init__(self, t):
        self.t = t

    def __call__(self):
        return self.t


def _reporter_app(throttle=0.12, last=0.0):
    return SimpleNamespace(message_queue=queue.Queue(), progress_throttle_secs=throttle,
                           _last_progress_ts=last)


def test_reporter_queue_tuples():
    app = _reporter_app()
    rep = GuiReporter(app, now=_Clock(100.0))
    rep.log("a log")
    rep.status("a status")
    rep.progress_max(500)
    assert _drain(app.message_queue) == [("log", "a log"), ("update_status", "a status"),
                                         ("update_progress_max", 500)]


def test_reporter_progress_throttle():
    app = _reporter_app(throttle=0.12)
    clock = _Clock(100.0)
    rep = GuiReporter(app, now=clock)
    rep.progress(10)            # first growing value: sent, clock set
    clock.t = 100.05
    rep.progress(20)            # inside the throttle window: dropped
    clock.t = 100.12
    rep.progress(30)            # window passed (>=): sent
    clock.t = 100.13
    rep.progress(40)            # dropped
    assert _drain(app.message_queue) == [("update_progress", 10), ("update_progress", 30)]
    assert app._last_progress_ts == 100.12


def test_reporter_final_value_always_sent_without_moving_clock():
    app = _reporter_app(throttle=1e9, last=50.0)
    clock = _Clock(60.0)
    rep = GuiReporter(app, now=clock)
    rep.progress(10)            # growing, but throttled (60 - 50 < 1e9)
    rep.progress(10)            # same value: final report, always sent
    rep.progress(7)             # smaller: also final
    assert _drain(app.message_queue) == [("update_progress", 10), ("update_progress", 7)]
    assert app._last_progress_ts == 50.0


def test_reporter_zero_first_value_counts_as_final():
    app = _reporter_app(throttle=1e9, last=99.0)
    rep = GuiReporter(app, now=_Clock(100.0))
    rep.progress(0)
    assert _drain(app.message_queue) == [("update_progress", 0)]
    assert app._last_progress_ts == 99.0


def test_reporter_reads_throttle_from_app_each_time():
    app = _reporter_app(throttle=1e9)
    clock = _Clock(100.0)
    rep = GuiReporter(app, now=clock)
    rep.progress(1)             # 100 - 0 < 1e9: dropped
    app.progress_throttle_secs = 0.0
    rep.progress(2)
    assert _drain(app.message_queue) == [("update_progress", 2)]
    assert app._last_progress_ts == 100.0


# --- GuiFolderStore ---

def test_folder_store_free_space_from_app(tmp_path):
    calls = []
    app = SimpleNamespace(get_free_space_reliable=lambda path: calls.append(path) or 1234)
    store = GuiFolderStore(app, str(tmp_path), clock=lambda: datetime(2024, 1, 2, 3, 4, 5))
    assert store.free_space() == 1234
    assert calls == [str(tmp_path)]


def test_folder_store_is_a_local_folder_store(tmp_path):
    app = SimpleNamespace(get_free_space_reliable=lambda path: 0)
    store = GuiFolderStore(app, str(tmp_path), clock=lambda: datetime(2024, 1, 2, 3, 4, 5))
    assert isinstance(store, LocalFolderStore)
    assert store.dest == str(tmp_path)
    assert store.new_backup_path("tester") == os.path.join(
        str(tmp_path), "tester_backup_2024-01-02_030405")
    assert store.new_log_path("tester") == os.path.join(
        str(tmp_path), "tester_backup_2024-01-02_030405.log")
    (tmp_path / "tester_backup_2023-01-01_000000").mkdir()
    (tmp_path / "other_backup_2023-01-01_000000").mkdir()
    assert store.list_backups("tester") == ["tester_backup_2023-01-01_000000"]


def test_begin_job_binds_a_run_id_per_thread(caplog):
    """P6-3: each work thread binds its own run_id and writes one start line."""
    caplog.set_level(logging.INFO, logger="ezbak.gui.bridge")
    seen, before = {}, diag.current()

    def work():
        seen["run_id"] = begin_job("backup", job_log="D:\\x.log")
        seen["current"] = diag.current()
    thread = threading.Thread(target=work)
    thread.start()
    thread.join()
    assert seen["current"] == seen["run_id"]
    assert len(seen["run_id"]) == len("20261027-090507-abcd")
    assert diag.current() == before  # the test thread is untouched
    assert caplog.messages == ["event=job_start kind=backup job_log=D:\\x.log"]
