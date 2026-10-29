"""Unit tests for ezbak.automation.headless (run_backup, HeadlessReporter, run), P3-8·9."""
import json
import logging
import os
import stat
from datetime import datetime

import pytest

from ezbak.automation import HeadlessReporter, headless
from ezbak.core import DEFAULT_FILTERS, BackupJob, BackupPolicy, JobLog, JobLogFormat
from ezbak.storage import LocalFolderStore

NOW = datetime(2024, 1, 2, 3, 4, 5)
NAME = "tester_backup_2024-01-02_030405"


def clock():
    return NOW


@pytest.fixture
def engine_calls(monkeypatch):
    """Replace the engine with a recorder; each call is (args, kwargs, folder_existed)."""
    calls = []

    def fake_engine(*args, **kwargs):
        calls.append((args, kwargs, os.path.isdir(args[1])))

    monkeypatch.setattr(headless, "engine_run_backup", fake_engine)
    return calls


@pytest.fixture
def users_root(tmp_path):
    root = tmp_path / "Users"
    (root / "tester").mkdir(parents=True)
    return str(root)


def test_run_backup_returns_paths_and_creates_folder(tmp_path, users_root, engine_calls):
    dest = tmp_path / "dest"
    dest.mkdir()
    backup_path, log_path = headless.run_backup("tester", str(dest), users_root=users_root,
                                                clock=clock)
    assert backup_path == os.path.join(str(dest), NAME)
    assert log_path == os.path.join(str(dest), NAME + ".log")
    assert os.path.isdir(backup_path)
    assert os.path.isfile(log_path)
    assert len(engine_calls) == 1 and engine_calls[0][2] is True


def test_same_name_folder_with_readonly_file_is_replaced(tmp_path, users_root, engine_calls):
    dest = tmp_path / "dest"
    old = dest / NAME / "sub"
    old.mkdir(parents=True)
    stale = old / "stale.txt"
    stale.write_text("old", encoding="utf-8")
    os.chmod(stale, stat.S_IREAD)
    backup_path, _ = headless.run_backup("tester", str(dest), users_root=users_root, clock=clock)
    assert os.path.isdir(backup_path)
    assert os.listdir(backup_path) == []


def test_missing_user_raises_after_mkdir_without_log(tmp_path, users_root, engine_calls):
    dest = tmp_path / "dest"
    dest.mkdir()
    with pytest.raises(FileNotFoundError) as exc:
        headless.run_backup("nobody", str(dest), users_root=users_root, clock=clock)
    user_home = os.path.join(users_root, "nobody")
    assert str(exc.value) == f"User folder '{user_home}' not found."
    assert os.listdir(dest) == ["nobody_backup_2024-01-02_030405"]  # folder made, no log
    assert engine_calls == []


def test_engine_arguments(tmp_path, users_root, engine_calls):
    dest = tmp_path / "dest"
    dest.mkdir()
    filters = {"exclude": [], "include": []}
    backup_path, log_path = headless.run_backup(
        "tester", str(dest), True, False, "logs", 5, 7, filters,
        users_root=users_root, clock=clock)
    args, kwargs, _ = engine_calls[0]
    job, path, job_log, fmt, store, targets, reporter, policy, clk = args
    assert job == BackupJob(user="tester", dest=str(dest), include_hidden=True,
                            include_system=False, retention_count=5, log_retention_days=7,
                            filters=filters, log_folder="logs")
    assert job.filters is filters
    assert path == backup_path
    assert isinstance(job_log, JobLog)
    assert fmt is JobLogFormat.HEADLESS
    assert isinstance(store, LocalFolderStore)
    assert [type(t).__name__ for t in targets] == ["Bookmarks"]
    assert isinstance(reporter, HeadlessReporter)
    assert type(reporter).__name__ == "_HeadlessReporter"  # recorded by the CLI golden
    assert policy is BackupPolicy.HEADLESS
    assert clk is clock
    assert kwargs["users_root"] == users_root
    assert kwargs["log_path"] == log_path
    assert isinstance(kwargs["now_ts"](), float)


def test_filters_none_uses_copy_of_defaults(tmp_path, users_root, engine_calls):
    dest = tmp_path / "dest"
    dest.mkdir()
    headless.run_backup("tester", str(dest), users_root=users_root, clock=clock)
    job = engine_calls[0][0][0]
    assert job.filters == DEFAULT_FILTERS and job.filters is not DEFAULT_FILTERS
    assert (job.include_hidden, job.include_system, job.log_folder,
            job.retention_count, job.log_retention_days) == (False, False, None, 2, 30)


def test_log_header_written_utf8(tmp_path, users_root, engine_calls):
    dest = tmp_path / "dest"
    dest.mkdir()
    _, log_path = headless.run_backup("tester", str(dest), users_root=users_root, clock=clock)
    with open(log_path, encoding="utf-8") as f:
        text = f.read()
    assert text.startswith("[2024-01-02 03:04:05] ")
    assert "tester" in text


def test_headless_reporter_output(capsys):
    reporter = HeadlessReporter()
    reporter.log("copy failed: x")
    reporter.status("ignored")
    reporter.progress(10)
    reporter.progress_max(100)
    assert capsys.readouterr().out == "copy failed: x\n"


def _args(**kw):
    from ezbak.automation import cli
    args = cli.parse([])
    for key, value in kw.items():
        setattr(args, key, value)
    return args


@pytest.fixture
def fake_run_backup(monkeypatch):
    """Replace headless.run_backup; set .error to make it raise."""
    calls = []

    def fake(*args, **kwargs):
        calls.append((args, kwargs))
        if fake.error is not None:
            raise fake.error
        return "B", "L"

    fake.error = None
    fake.calls = calls
    monkeypatch.setattr(headless, "run_backup", fake)
    return fake


@pytest.fixture(autouse=True)
def temp_appdata(tmp_path, monkeypatch):
    """headless.run sets up the diagnostic log under %APPDATA% (P6-2): keep it in tmp."""
    monkeypatch.setenv("APPDATA", str(tmp_path / "AppData"))
    return tmp_path / "AppData" / "ezBAK" / "logs" / "ezbak.log"


def _diag_lines(path):
    for handler in logging.getLogger("ezbak").handlers:
        handler.flush()
    return path.read_text(encoding="utf-8").splitlines()


def _write_config(tmp_path, config):
    path = tmp_path / "cfg.json"
    path.write_text(json.dumps(config), encoding="utf-8")
    return str(path)


def test_run_without_request_returns_none(fake_run_backup):
    assert headless.run(_args(), users_root="R", clock=clock) is None
    assert headless.run(_args(user="u"), users_root="R", clock=clock) is None
    assert fake_run_backup.calls == []


def test_run_config_success_returns_0(tmp_path, fake_run_backup, capsys):
    path = _write_config(tmp_path, {"user": "u", "dest": "d", "retention_count": "5"})
    assert headless.run(_args(schedule_config=path), users_root="R", clock=clock) == 0
    args, kwargs = fake_run_backup.calls[0]
    assert args == ("u", "d") and kwargs["retention_count"] == "5"
    assert kwargs["log_folder"] == "logs" and kwargs["users_root"] == "R"
    assert kwargs["clock"] is clock
    assert "[2024-01-02 03:04:05] Backup finished → B" in capsys.readouterr().out


def test_run_config_missing_file_returns_1(tmp_path, fake_run_backup, capsys):
    path = str(tmp_path / "none.json")
    assert headless.run(_args(schedule_config=path), users_root="R", clock=clock) == 1
    assert "ERROR: Configuration file not found" in capsys.readouterr().out
    assert fake_run_backup.calls == []


def test_run_config_error_returns_1_with_traceback(tmp_path, fake_run_backup, capsys):
    fake_run_backup.error = RuntimeError("boom")
    path = _write_config(tmp_path, {"user": "u", "dest": "d"})
    assert headless.run(_args(schedule_config=path), users_root="R", clock=clock) == 1
    out, err = capsys.readouterr()
    assert "ERROR during scheduled backup: boom" in out
    assert "Traceback" in err


def test_run_direct_returns_0_even_on_error(fake_run_backup, capsys):
    assert headless.run(_args(user="u", dest="d", filters='{"include": []}'),
                        users_root="R", clock=clock) == 0
    assert fake_run_backup.calls[0][1]["filters"] == {"include": []}
    fake_run_backup.error = RuntimeError("boom")
    assert headless.run(_args(user="u", dest="d"), users_root="R", clock=clock) == 0
    out, err = capsys.readouterr()
    assert "ERROR during headless backup: boom" in out and err == ""


def test_run_without_request_sets_up_no_diagnostic_log(fake_run_backup, temp_appdata):
    assert headless.run(_args(user="u"), users_root="R", clock=clock) is None
    assert not temp_appdata.exists()


def test_run_logs_the_failure_with_a_run_id(tmp_path, fake_run_backup, temp_appdata):
    """P6-2: a Headless run binds one run_id; the failure goes to ezbak.log with its
    traceback (stdout and the exit code stay as before)."""
    fake_run_backup.error = RuntimeError("boom")
    path = _write_config(tmp_path, {"user": "u", "dest": "d"})
    assert headless.run(_args(schedule_config=path), users_root="R", clock=clock) == 1
    lines = _diag_lines(temp_appdata)
    assert " ERROR ezbak.automation.headless [20240102-030405-" in lines[0]
    assert lines[0].endswith("] event=job_failed kind=backup mode=config")
    assert "RuntimeError: boom" in lines


def test_run_backup_writes_the_job_log_path(tmp_path, users_root, engine_calls, caplog):
    caplog.set_level(logging.INFO, logger="ezbak.automation.headless")
    _, log_path = headless.run_backup("tester", str(tmp_path / "dest"),
                                      users_root=users_root, clock=clock)
    assert f"event=job_start kind=backup mode=headless job_log={log_path}" in caplog.messages


def test_run_backup_records_the_last_backup(tmp_path, users_root, engine_calls, temp_appdata):
    """U7-2: a finished Headless backup is the home page's last backup (state.json)."""
    from ezbak.core import history
    dest = str(tmp_path / "dest")
    backup_path, _ = headless.run_backup("tester", dest, users_root=users_root, clock=clock)
    assert history.last_backup() == {"time": "2024-01-02 03:04:05", "user": "tester",
                                     "dest": dest, "folder": backup_path, "mode": "headless"}
