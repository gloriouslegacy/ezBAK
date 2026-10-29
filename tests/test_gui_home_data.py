"""U7-5: ezbak.gui.home_data — copy speed samples and the slow loads of the home cards.

A stand-in App (no Tk root): progress bar values, a recording Sparkline, viewable or not.
The work thread is run inline and its sources are replaced, so nothing real is read.
"""
from datetime import datetime
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.gui

pytest.importorskip("customtkinter")

from ezbak.gui import home_data  # noqa: E402


class Chart:
    def __init__(self):
        self.calls = []

    def push(self, value):
        self.calls.append(("push", value))

    def clear(self):
        self.calls.append(("clear",))


class FakeApp:
    def __init__(self, viewable=True):
        self.progress_bar = {"value": 0, "maximum": 100}
        self.home_speed_chart = Chart()
        self.home_data = home_data.HomeData()
        self.viewable = viewable

    def winfo_viewable(self):
        return int(self.viewable)


@pytest.fixture
def sources(monkeypatch):
    """Inline thread and fixed sources; returns the call record."""
    calls = []

    class InlineThread:
        def __init__(self, target, args, daemon):
            self.run = lambda: target(*args)

        def start(self):
            calls.append("thread")
            self.run()

    last = {"time": "2026-10-27 09:05:07", "user": "kim", "dest": "D:\\Backup",
            "folder": "D:\\Backup\\kim_backup_x", "mode": "gui"}
    monkeypatch.setattr(home_data.threading, "Thread", InlineThread)
    monkeypatch.setattr(home_data, "core_history", SimpleNamespace(last_backup=lambda: last))
    monkeypatch.setattr(home_data, "automation_schedule_config",
                        SimpleNamespace(load_all=lambda: [{"task_name": "n"}]))
    monkeypatch.setattr(home_data, "automation_scheduler", SimpleNamespace(
        next_scheduled=lambda configs, now: ("n", datetime(2026, 11, 1, 9, 0))))
    monkeypatch.setattr(home_data, "winsys_free_space", lambda path: calls.append(path) or 5 * 1024**3)

    def system_name():
        calls.append("system_name")
        return "BOARD"
    monkeypatch.setattr(home_data, "get_system_info", system_name)
    return calls


def tick_at(app, t):
    home_data.tick(app, now=lambda: t)


def test_nothing_happens_while_not_viewable(sources):
    app = FakeApp(viewable=False)
    app.progress_bar.update(value=500, maximum=1000)
    tick_at(app, 1.0)
    assert sources == [] and app.home_speed_chart.calls == [] and app.home_data.ticks == 0


def test_slow_load_fills_the_fields_and_reads_the_system_name_once(sources):
    app = FakeApp()
    for i in range(home_data.SLOW_EVERY + 1):
        tick_at(app, float(i))
    data = app.home_data
    assert data.last_backup["user"] == "kim"
    assert data.next_run == ("n", datetime(2026, 11, 1, 9, 0))
    assert data.free_space == ("D:", 5 * 1024**3)
    assert data.system_name == "BOARD" and not data.loading
    assert sources == ["thread", "D:\\Backup", "system_name", "thread", "D:\\Backup"]


def test_free_space_fallback_and_no_last_backup(sources, monkeypatch):
    monkeypatch.setattr(home_data, "winsys_free_space",
                        lambda path: home_data.winsys_sysinfo.FALLBACK_FREE_BYTES)
    assert home_data._free_space({"dest": "D:\\B"}) is None
    assert home_data._free_space(None) is None
    assert home_data._free_space({"dest": ""}) is None


def test_load_failure_keeps_values_and_clears_loading(sources, monkeypatch, caplog):
    def broken():
        raise OSError("disk")
    monkeypatch.setattr(home_data, "core_history", SimpleNamespace(last_backup=broken))
    app = FakeApp()
    tick_at(app, 0.0)
    assert app.home_data.last_backup is None and not app.home_data.loading
    assert "event=home_data_load failed" in caplog.text


def test_speed_samples_bytes_jobs_only(sources):
    app = FakeApp()
    bar, chart, data = app.progress_bar, app.home_speed_chart.calls, app.home_data
    bar.update(value=50, maximum=100)         # percentage job: no speed
    tick_at(app, 0.0)
    assert data.speed is None and chart == []
    bar.update(value=1000, maximum=10000)      # bytes job starts: new line
    tick_at(app, 1.0)
    bar["value"] = 3000
    tick_at(app, 1.5)
    assert data.speed == 4000.0 and chart == [("clear",), ("push", 4000.0)]
    bar["value"] = 500                          # value went back: a new job
    tick_at(app, 2.0)
    assert chart[-1] == ("clear",)
    bar["value"] = 10000                        # finished (value == maximum): idle
    tick_at(app, 2.5)
    assert data.speed is None
