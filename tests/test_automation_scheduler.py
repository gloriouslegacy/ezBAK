"""Unit tests for automation.scheduler: pure functions, XML file (P3-6), schtasks run (P3-7)."""
import locale
import os
import sys
import types
from datetime import datetime
from xml.etree import ElementTree

import pytest

from ezbak.automation import ScheduleSpec, scheduler
from ezbak.winsys import proc

NOW = datetime(2026, 10, 12, 9, 30, 5)


def _xml(sc_type, **kw):
    values = dict(executable="C:\\Py\\python.exe", arguments='--schedule-config "C:\\c.json"',
                  formatted_time="07:15", sc_type=sc_type, run_as_user="PC\\홍길동", monthly_day=15)
    values.update(kw)
    return scheduler.build_task_xml(**values, now=NOW)


@pytest.mark.parametrize("sc_type, present, absent", [
    ("DAILY", "<DaysInterval>1</DaysInterval>", "<ScheduleByWeek>"),
    ("WEEKLY", "<Sunday />", "<ScheduleByDay>"),
    ("MONTHLY", "<Day>15</Day>", "<ScheduleByWeek>"),
])
def test_build_task_xml_trigger_by_type(sc_type, present, absent):
    xml = _xml(sc_type)
    assert xml.startswith('<?xml version="1.0" encoding="UTF-16"?>\n<Task version="1.2"')
    assert xml.endswith("  </Actions>\n</Task>")
    assert "<Date>2026-10-12T09:30:05</Date>" in xml
    assert "<StartBoundary>2026-10-13T07:15:00</StartBoundary>" in xml
    assert xml.count("PC\\홍길동") == 2  # Author and UserId
    assert present in xml and absent not in xml
    assert xml.count("<CalendarTrigger>") == 1


def test_build_task_xml_monthly_lists_all_months():
    xml = _xml("MONTHLY")
    for month in ("January", "June", "December"):
        assert f"<{month} />" in xml


def test_build_task_xml_unknown_type_has_no_trigger():
    xml = _xml("HOURLY")
    assert "<Triggers>\n  </Triggers>" in xml


def test_build_task_xml_escapes_values():
    # fix-f: element text escapes & < > only; quotes stay as they are.
    xml = _xml("DAILY", executable="C:\\A&B\\python.exe", arguments='"C:\\x<y>\\ezBAK.py"',
               run_as_user="D&C\\u<1>")
    assert "<Command>C:\\A&amp;B\\python.exe</Command>" in xml
    assert '<Arguments>"C:\\x&lt;y&gt;\\ezBAK.py"</Arguments>' in xml
    assert "<Author>D&amp;C\\u&lt;1&gt;</Author>" in xml
    assert "<UserId>D&amp;C\\u&lt;1&gt;</UserId>" in xml
    root = ElementTree.fromstring(xml.split("\n", 1)[1])
    ns = "{http://schemas.microsoft.com/windows/2004/02/mit/task}"
    assert root.find(f"{ns}Actions/{ns}Exec/{ns}Arguments").text == '"C:\\x<y>\\ezBAK.py"'


def test_build_task_xml_start_crosses_month_end():
    xml = scheduler.build_task_xml("e", "a", "23:59", "DAILY", "u", now=datetime(2024, 2, 29, 1, 0))
    assert "<StartBoundary>2024-03-01T23:59:00</StartBoundary>" in xml


@pytest.mark.parametrize("given, expected", [("7:05", "07:05"), (" 14:30 ", "14:30"), ("00:00", "00:00")])
def test_normalize_time(given, expected):
    assert scheduler.normalize_time(given) == expected


@pytest.mark.parametrize("given", ["24:00", "7", "ab:cd", ""])
def test_normalize_time_invalid(given):
    with pytest.raises(ValueError) as err:
        scheduler.normalize_time(given)
    assert str(err.value) == f"Invalid time format '{given}'. Please use HH:MM format (e.g., 14:30)"


@pytest.mark.parametrize("given, expected", [("Daily", "DAILY"), (" WEEKLY ", "WEEKLY"), ("monthly", "MONTHLY")])
def test_schedule_type(given, expected):
    assert scheduler.schedule_type(given) == expected


def test_schedule_type_invalid():
    with pytest.raises(ValueError) as err:
        scheduler.schedule_type("Hourly")
    assert str(err.value) == "Unsupported schedule type: Hourly. Use Daily, Weekly, or Monthly"


def test_task_arguments_not_frozen(monkeypatch):
    monkeypatch.delattr(sys, "frozen", raising=False)
    assert scheduler.task_arguments("C:\\app\\ezBAK.py", "C:\\c.json") == \
        '"C:\\app\\ezBAK.py" --schedule-config "C:\\c.json"'


def test_task_arguments_frozen(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    assert scheduler.task_arguments("C:\\app\\ezBAK.py", "C:\\c.json") == '--schedule-config "C:\\c.json"'


@pytest.mark.parametrize("environ, expected", [
    ({"USERNAME": "kim", "USERDOMAIN": "CORP"}, "CORP\\kim"),
    ({"USERNAME": "kim", "USERDOMAIN": "KIM"}, "kim"),
    ({"USERNAME": "kim", "COMPUTERNAME": "PC1"}, "PC1\\kim"),
    ({"USERNAME": "kim", "USERDOMAIN": ""}, "kim"),
    ({}, "Administrator"),
])
def test_run_as_user(environ, expected):
    assert scheduler.run_as_user(environ) == expected


def test_build_create_cmd():
    assert scheduler.build_create_cmd("ezBAK_D", "C:\\x.xml") == \
        ["schtasks", "/create", "/tn", "ezBAK_D", "/xml", "C:\\x.xml", "/f"]


@pytest.mark.parametrize("rc, stderr, stdout, expected", [
    (5, "", "", "schtasks failed with return code 5"),
    (1, None, "out", "schtasks failed with return code 1\n\n"
     "General error. Check if task name is valid and you have admin privileges.\n\nSystem Output: out"),
    (267, "ERROR: The task name is invalid time", "", "schtasks failed with return code 267\n\n"
     "Solution: Check time format (use HH:MM, e.g., 14:30)\n\n"
     "Invalid task name or path. Avoid special characters in task name.\n\n"
     "System Error: ERROR: The task name is invalid time"),
    (-2147216609, "Service is not running", "", "schtasks failed with return code -2147216609\n\n"
     "Solution: Task Scheduler service may not be running. Try: net start schedule\n\n"
     "Task Scheduler service is not available.\n\nSystem Error: Service is not running"),
])
def test_explain_error_table(rc, stderr, stdout, expected):
    assert scheduler.explain_error(rc, stderr, stdout) == expected


def test_explain_error_access_denied_adds_admin_info():
    msg = scheduler.explain_error(2, "ERROR: Access is denied.", "")
    assert msg == ("schtasks failed with return code 2\n\n"
                   "Solution: Run the application as Administrator\n\n"
                   "Access denied. Try running as administrator.\n\n"
                   "Additional Info:\n- Make sure ezBAK is running as Administrator"
                   "\n- The scheduled task will run with HIGHEST privilege level"
                   "\n- Check Windows Task Scheduler service is running\n\n"
                   "System Error: ERROR: Access is denied.")


@pytest.mark.parametrize("stderr, hint", [
    ("task already exists", "Choose a different task name"),
    ("no privilege", "Administrator privileges required for task creation"),
    ("user not found", "Check user account format"),
])
def test_explain_error_stderr_hints(stderr, hint):
    assert f"\n\nSolution: {hint}" in scheduler.explain_error(99, stderr, "")


def test_write_task_xml_utf16_bom_crlf(monkeypatch, tmp_path):
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    xml = _xml("DAILY")
    path = scheduler.write_task_xml("ezBAK_매일", xml)
    assert path == os.path.join(str(tmp_path), "AppData", "Local", "ezBAK", "schedules", "ezBAK_매일_task.xml")
    with open(path, "rb") as f:
        raw = f.read()
    assert raw[:2] == b"\xff\xfe"
    text = raw.decode("utf-16")
    assert "\r\n" in text and "\n" not in text.replace("\r\n", "")
    assert text.replace("\r\n", "\n") == xml


def test_write_task_xml_file_name_uses_safe_name(monkeypatch, tmp_path):
    # fix-f: same rule as the config JSON; characters invalid in file names become '_'.
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    path = scheduler.write_task_xml('a<b>c:d"e/f\\g|h?i*j&k', "x")
    assert os.path.basename(path) == "a_b_c_d_e_f_g_h_i_j&k_task.xml"
    assert os.path.dirname(path) == os.path.join(str(tmp_path), "AppData", "Local", "ezBAK", "schedules")
    with open(path, encoding="utf-16") as f:
        assert f.read() == "x"


def test_write_task_xml_home_read_at_call_time(monkeypatch, tmp_path):
    monkeypatch.setenv("USERPROFILE", str(tmp_path / "a"))
    first = scheduler.write_task_xml("t", "x")
    monkeypatch.setenv("USERPROFILE", str(tmp_path / "b"))
    second = scheduler.write_task_xml("t", "x")
    assert first.startswith(str(tmp_path / "a")) and second.startswith(str(tmp_path / "b"))


class _Run:
    """Stand-in for proc.run: records argv/kwargs, answers with a result or raises."""

    def __init__(self, returncode=0, stdout="", stderr="", raises=None):
        self.calls, self.raises = [], raises
        self.result = types.SimpleNamespace(returncode=returncode, stdout=stdout, stderr=stderr)

    def __call__(self, argv, **kwargs):
        self.calls.append((argv, kwargs))
        if self.raises is not None:
            raise self.raises
        return self.result


@pytest.fixture
def run(monkeypatch):
    def install(**kw):
        fake = _Run(**kw)
        monkeypatch.setattr(proc, "run", fake)
        return fake
    return install


@pytest.fixture
def create_env(monkeypatch, tmp_path, isolated_env):
    monkeypatch.setenv("USERPROFILE", str(tmp_path / "home"))
    monkeypatch.setenv("USERNAME", "tester")
    monkeypatch.setenv("USERDOMAIN", "PC")
    monkeypatch.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "<ENC>")
    monkeypatch.delattr(sys, "frozen", raising=False)
    return isolated_env


def _spec(**kw):
    values = dict(task_name="ezBAK_T", user="tester", dest="D:\\bk", schedule="Daily", time_str="9:05",
                  monthly_day=1, include_hidden=False, include_system=True, retention_count=3,
                  log_retention_days=15, filters=None)
    values.update(kw)
    return ScheduleSpec(**values)


def _create(log):
    return scheduler.create_task(_spec(), "C:\\app\\ezBAK.py", out=log.append, now=NOW,
                                 debug=lambda msg: None)


def test_create_task_success_argv_kwargs(run, create_env):
    fake, log = run(), []
    text = _create(log)
    argv, kwargs = fake.calls[0]
    xml_file = os.path.join(os.path.expanduser("~"), "AppData", "Local", "ezBAK", "schedules",
                            "ezBAK_T_task.xml")
    assert argv == ["schtasks", "/create", "/tn", "ezBAK_T", "/xml", xml_file, "/f"]
    assert kwargs == dict(check=False, shell=False, capture_output=True, text=True,
                          encoding="<ENC>", errors="replace", timeout=180)
    assert os.path.isfile(xml_file)
    assert (create_env["APPDATA"] / "ezBAK" / "ezBAK_Schedule" / "ezBAK_T.json").is_file()
    assert text.startswith("Scheduled task 'ezBAK_T' created successfully!\n\nSchedule: Daily at 09:05\n")
    assert text.endswith("Run As: PC\\tester")
    assert log[0] == "Time format validated: 9:05 -> 09:05"
    assert 'Task arguments: "C:\\app\\ezBAK.py" --schedule-config' in log[3]
    assert log[-1] == "Task creation successful: ezBAK_T (with admin privileges)"


def test_create_task_rc1_explains_error(run, create_env):
    run(returncode=1, stderr="ERROR: Access is denied.\n")
    log = []
    with pytest.raises(RuntimeError) as info:
        _create(log)
    assert str(info.value) == scheduler.explain_error(1, "ERROR: Access is denied.\n", "")
    assert "schtasks stderr: ERROR: Access is denied.\n" in log


def test_create_task_timeout(run, create_env):
    run(raises=proc.TimeoutExpired(["schtasks"], 180))
    log = []
    with pytest.raises(RuntimeError, match="^Task creation timed out. The system may be busy.$") as info:
        _create(log)
    assert isinstance(info.value.__context__, proc.TimeoutExpired)
    assert log[-1] == "schtasks command timed out"


def _config(env):
    return env["APPDATA"] / "ezBAK" / "ezBAK_Schedule" / "ezBAK_T.json"


def _xml_file():
    return os.path.join(os.path.expanduser("~"), "AppData", "Local", "ezBAK", "schedules",
                        "ezBAK_T_task.xml")


def test_create_task_rc1_removes_new_config_and_xml(run, create_env):
    run(returncode=1, stderr="ERROR: Access is denied.\n")
    with pytest.raises(RuntimeError):
        _create([])
    assert not _config(create_env).exists()
    assert not os.path.exists(_xml_file())


def test_create_task_rc1_restores_previous_config(run, create_env):
    run(returncode=2, stdout="INFO: odd\n")
    prev = b'{"task_name": "pre-existing"}'
    _config(create_env).parent.mkdir(parents=True)
    _config(create_env).write_bytes(prev)
    with pytest.raises(RuntimeError) as info:
        _create([])
    assert str(info.value) == scheduler.explain_error(2, "", "INFO: odd\n")
    assert _config(create_env).read_bytes() == prev
    assert not os.path.exists(_xml_file())


def test_create_task_xml_failure_restores_config(run, create_env, monkeypatch):
    fake = run()

    def fail(task_name, xml):
        raise OSError("disk full")
    monkeypatch.setattr(scheduler, "write_task_xml", fail)
    with pytest.raises(RuntimeError, match="^Failed to create task XML file: disk full$"):
        _create([])
    assert fake.calls == []
    assert not _config(create_env).exists()


def test_create_task_rollback_failure_warns_and_keeps_error(run, create_env, monkeypatch, caplog):
    run(returncode=1, stderr="ERROR: Access is denied.\n")

    def fail(task_name, prev):
        raise PermissionError("locked")
    monkeypatch.setattr(scheduler.schedule_config, "restore_raw", fail)
    with caplog.at_level("WARNING", logger=scheduler.__name__):
        with pytest.raises(RuntimeError) as info:
            _create([])
    assert str(info.value) == scheduler.explain_error(1, "ERROR: Access is denied.\n", "")
    assert info.value.__cause__ is None
    warnings = [r for r in caplog.records if r.levelname == "WARNING"]
    assert len(warnings) == 1 and "rollback_failed task=ezBAK_T error=PermissionError" in warnings[0].message


def test_create_task_timeout_keeps_config_and_xml(run, create_env):
    run(raises=proc.TimeoutExpired(["schtasks"], 180))
    with pytest.raises(RuntimeError, match="timed out"):
        _create([])
    assert _config(create_env).is_file()
    assert os.path.isfile(_xml_file())


def test_create_task_bad_time_runs_nothing(run, create_env):
    fake = run()
    with pytest.raises(ValueError):
        scheduler.create_task(_spec(time_str="25:00"), "x.py", out=[].append, now=NOW,
                              debug=lambda msg: None)
    assert fake.calls == []


def test_build_delete_cmd():
    assert scheduler.build_delete_cmd("A B") == ["schtasks", "/Delete", "/TN", "A B", "/F"]


def test_delete_task_success_keeps_shell_true(run, create_env):
    fake = run()
    assert scheduler.delete_task("ezBAK_T") is None
    assert fake.calls == [(["schtasks", "/Delete", "/TN", "ezBAK_T", "/F"],
                           dict(shell=True, capture_output=True, text=True))]


@pytest.mark.parametrize("stdout, stderr, message", [
    ("out\n", "ERROR: not found\n", "ERROR: not found\n"),
    ("out\n", "", "out\n"),
    ("", "", ""),
])
def test_delete_task_failure(run, create_env, stdout, stderr, message):
    run(returncode=1, stdout=stdout, stderr=stderr)
    with pytest.raises(RuntimeError) as info:
        scheduler.delete_task("ezBAK_T")
    assert str(info.value) == message


def _make_xml_file():
    os.makedirs(os.path.dirname(_xml_file()), exist_ok=True)
    with open(_xml_file(), "w", encoding="utf-16") as f:
        f.write("<Task/>")


def test_delete_task_success_removes_xml(run, create_env):
    run()
    _make_xml_file()
    assert scheduler.delete_task("ezBAK_T") is None
    assert not os.path.exists(_xml_file())


def test_delete_task_xml_path_uses_safe_name(run, create_env):
    assert scheduler.task_xml_path("a/b") == os.path.join(
        os.path.dirname(_xml_file()), f"{scheduler.schedule_config.safe_name('a/b')}_task.xml")


def test_delete_task_without_xml_is_quiet(run, create_env, caplog):
    run()
    with caplog.at_level("WARNING", logger=scheduler.__name__):
        assert scheduler.delete_task("ezBAK_T") is None
    assert not [r for r in caplog.records if r.levelname == "WARNING"]


def test_delete_task_failure_keeps_xml(run, create_env):
    run(returncode=1, stderr="ERROR: not found\n")
    _make_xml_file()
    with pytest.raises(RuntimeError):
        scheduler.delete_task("ezBAK_T")
    assert os.path.isfile(_xml_file())


def test_delete_task_xml_remove_failure_only_warns(run, create_env, monkeypatch, caplog):
    run()
    _make_xml_file()

    def fail(path):
        raise PermissionError("locked")
    monkeypatch.setattr(scheduler.os, "remove", fail)
    with caplog.at_level("WARNING", logger=scheduler.__name__):
        assert scheduler.delete_task("ezBAK_T") is None
    warnings = [r for r in caplog.records if r.levelname == "WARNING"]
    assert len(warnings) == 1
    assert "xml_remove_failed task=ezBAK_T error=PermissionError" in warnings[0].message


# --------------------------------------------------------------- next run (U7-1)

# 2026-10-27 is a Tuesday; created the day before, so the first run (created + 1 day) is today
RUN_NOW = datetime(2026, 10, 27, 10, 0)
CREATED = "2026-10-26 08:00:00"


def _cfg(schedule, time_str, monthly_day=1, created_at=CREATED, task_name="t"):
    return {"task_name": task_name, "schedule": schedule, "time_str": time_str,
            "monthly_day": monthly_day, "created_at": created_at}


@pytest.mark.parametrize("config, expected", [
    (_cfg("Daily", "11:30"), datetime(2026, 10, 27, 11, 30)),        # later today
    (_cfg("Daily", "09:00"), datetime(2026, 10, 28, 9, 0)),          # passed: tomorrow
    (_cfg("daily ", " 9:5"), datetime(2026, 10, 28, 9, 5)),          # same normalising as create_task
    (_cfg("Weekly", "09:00"), datetime(2026, 11, 1, 9, 0)),          # Sunday, as the task XML
    (_cfg("Monthly", "09:00", 27), datetime(2026, 11, 27, 9, 0)),    # today's run has passed
    (_cfg("Monthly", "12:00", 27), datetime(2026, 10, 27, 12, 0)),
    (_cfg("Monthly", "09:00", 31), datetime(2026, 10, 31, 9, 0)),
    (_cfg("Monthly", "09:00", 31, created_at="2026-10-31 10:00:00"), datetime(2026, 12, 31, 9, 0)),  # no 31 Nov
    (_cfg("Daily", "08:00", created_at="2026-10-27 09:00:00"), datetime(2026, 10, 28, 8, 0)),  # first run tomorrow
    (_cfg("Daily", "20:00", created_at="2026-10-27 09:00:00"), datetime(2026, 10, 28, 20, 0)),  # not today
    (_cfg("Daily", "11:00", created_at=None), datetime(2026, 10, 27, 11, 0)),   # no created_at: from now
    (_cfg("Daily", "11:00", created_at="yesterday"), datetime(2026, 10, 27, 11, 0)),
])
def test_next_run(config, expected):
    assert scheduler.next_run(config, RUN_NOW) == expected


@pytest.mark.parametrize("config", [
    _cfg("Hourly", "09:00"), _cfg("Daily", "25:00"), _cfg("Daily", ""),
    _cfg("Monthly", "09:00", "x"), _cfg("Monthly", "09:00", 32), {},
])
def test_next_run_unknown(config):
    assert scheduler.next_run(config, RUN_NOW) is None


def test_next_scheduled_earliest_and_skips_bad():
    configs = [_cfg("Weekly", "09:00", task_name="weekly"), "not a dict",
               _cfg("Daily", "11:30", task_name="daily"), _cfg("Hourly", "09:00", task_name="bad")]
    assert scheduler.next_scheduled(configs, RUN_NOW) == ("daily", datetime(2026, 10, 27, 11, 30))
    assert scheduler.next_scheduled([_cfg("Hourly", "09:00")], RUN_NOW) is None
    assert scheduler.next_scheduled([], RUN_NOW) is None
