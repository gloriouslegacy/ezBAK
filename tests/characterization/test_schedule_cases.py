"""Characterization (golden) of scheduled-task create/delete on the untouched ezBAK.py (P3-0b2).

App.create_scheduled_task and App.delete_scheduled_task run on the session App (gui_app)
with every external command answered by record_subprocess (success, rc 1, TimeoutExpired);
nothing is registered. %APPDATA%, %LOCALAPPDATA% and USERPROFILE (the task XML goes to
~\\AppData\\Local\\ezBAK\\schedules) point at the test's temp folder.

The schtasks argv is already pinned by test_commands_argv (argv_golden.json); here each case
records the rest: the files left under the temp env (schedule JSON text with its line ends,
task XML text), the returned message or raised error (type, text, cause type), the dialogs,
the message queue, the diagnostic prints and the write_detailed_log messages (recorded,
since no log file is open; traceback frames folded). Paths: temp root <TMP>, repo <REPO>,
python <PY>. The clock is fixed_clock (2024-01-02 03:04:05). Odd behavior is kept as-is.

test_gui_json_to_cli runs the JSON written by create_scheduled_task through the Headless
CLI (`--schedule-config`, which is what the task arguments use) with runpy (run_ezbak_main):
the stub engine must receive the JSON values unchanged. User `Public`: runpy cannot inject
USERS_ROOT.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_schedule_cases.py
"""
import contextlib
import io
import json
import locale
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = [pytest.mark.gui, pytest.mark.windows]

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
GOLDEN_PATH = REPO_ROOT / "tests" / "fixtures" / "schedule_cases_golden.json"
DIALOGS = ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel")
# The real user's XML folder, taken before any test patches USERPROFILE: it must not change.
REAL_SCHEDULES = Path(os.path.expanduser("~")) / "AppData" / "Local" / "ezBAK" / "schedules"
CUSTOM_FILTERS = {"include": [{"type": "Name", "pattern": "*.DOCX"}],
                  "exclude": [{"type": "NAME", "pattern": "~$*"}, {"type": "name", "pattern": "*.tmp"}]}
ACCESS_DENIED = "ERROR: Access is denied.\n"


def _create(task_name="ezBAK_Test", schedule="Daily", time_str="09:30", **kwargs):
    return {"op": "create", "args": dict(task_name=task_name, schedule=schedule,
                                         time_str=time_str, user="tester", **kwargs)}


# Each case: op create/delete, args, optional respond (record_subprocess rule kwargs),
# env (overrides, None removes), config (schedule JSON files written before the call).
CASES = {
    "create_daily_ok": _create(include_hidden=True, retention_count=3, log_retention_days=15,
                               filters=CUSTOM_FILTERS),
    "create_weekly_time_and_kind_normalized": _create(schedule=" WEEKLY ", time_str=" 9:5 "),
    "create_monthly_day31": _create(schedule="monthly", time_str="23:59", monthly_day=31,
                                    include_system=True),
    "create_run_as_user_without_domain": {**_create(), "env": {"USERDOMAIN": "tester"}},
    "create_run_as_user_computername": {**_create(), "env": {"USERDOMAIN": None,
                                                             "COMPUTERNAME": "PC01"}},
    "create_name_colon": _create(task_name="Daily:Backup"),
    "create_name_invalid_chars": _create(task_name='Nightly <A|B>?"*'),
    "create_name_ampersand": _create(task_name="A&B"),  # fix-f: <Arguments> escapes & as &amp;
    "create_bad_time": _create(time_str="25:00"),
    "create_bad_time_text": _create(time_str="nine"),
    "create_bad_kind": _create(schedule="Hourly"),
    "create_rc1_access_denied": {**_create(), "respond": {"returncode": 1, "stderr": ACCESS_DENIED}},
    "create_rc2_stdout_only": {**_create(), "respond": {"returncode": 2, "stdout": "INFO: odd\n"}},
    # fix-g: a failed create puts the config back (removed, or the previous text) and removes the XML
    "create_overwrite_rc1_keeps_previous": {**_create(), "config": ["ezBAK_Test.json"],
                                            "respond": {"returncode": 1, "stderr": ACCESS_DENIED}},
    "create_timeout": {**_create(), "respond": {"raises": "timeout"}},  # kept: may be registered (F3)
    "create_no_appdata": {**_create(), "env": {"APPDATA": None}},
    "delete_ok_config_sanitized": {"op": "delete", "args": {"task_name": "Old:Task"},
                                   "config": ["Old_Task.json"]},
    "delete_ok_no_config": {"op": "delete", "args": {"task_name": "ezBAK_Test"}},
    "delete_fail_rc1": {"op": "delete", "args": {"task_name": "ezBAK_Test"},
                        "config": ["ezBAK_Test.json"],
                        "respond": {"returncode": 1,
                                    "stderr": "ERROR: The system cannot find the file specified.\n"}},
    "delete_fail_rc1_no_output": {"op": "delete", "args": {"task_name": "ezBAK_Test"},
                                  "respond": {"returncode": 1}},
    "delete_empty_name": {"op": "delete", "args": {"task_name": ""}},
    "delete_no_appdata": {"op": "delete", "args": {"task_name": "ezBAK_Test"},
                          "env": {"APPDATA": None}},
    "create_then_delete": {"op": "create_then_delete", "args": _create()["args"]},
}


def _fold_traceback(text):
    """Keep the 'Traceback' header, chain lines and exception lines; frames become one line."""
    out, in_frames = [], False
    for line in text.splitlines():
        if line.endswith("Traceback (most recent call last):"):  # also "Full error traceback: ..."
            out.extend([line, "  <FRAMES>"])
            in_frames = True
        elif in_frames and line.startswith(" "):
            continue
        else:
            in_frames = False
            out.append(line)
    return out


def _normalizer(tmp_path):
    subs = [(str(tmp_path), "<TMP>"), (str(REPO_ROOT), "<REPO>"), (sys.executable, "<PY>")]

    def norm(value):
        if isinstance(value, str):
            for root, mark in subs:
                value = value.replace(root.replace("\\", "\\\\"), mark).replace(root, mark)
            return value
        if isinstance(value, list):
            return [norm(v) for v in value]
        if isinstance(value, dict):
            return {norm(k): norm(v) for k, v in value.items()}
        return value
    return norm


def _read_file(path):
    data = path.read_bytes()
    if path.suffix == ".xml":
        assert data[:2] == b"\xff\xfe"  # UTF-16 LE with BOM
        text = data.decode("utf-16")
    elif path.suffix == ".json":
        text = data.decode("utf-8")
    else:
        return {"size": len(data)}
    return {"crlf": "\r\n" in text, "lines": text.splitlines()}


def _tree(tmp_path):
    """Files under the temp env and home (the temp root itself holds dest/ and these)."""
    files = {}
    for base in ("env", "home"):
        for root, _dirs, names in os.walk(tmp_path / base):
            for name in names:
                path = Path(root) / name
                files[path.relative_to(tmp_path).as_posix()] = _read_file(path)
    return dict(sorted(files.items()))


def _real_schedules_state():
    if not REAL_SCHEDULES.is_dir():
        return None
    return sorted((p.name, p.stat().st_mtime_ns) for p in REAL_SCHEDULES.iterdir())


@pytest.fixture
def sched(ezbak_mod, tmp_path, monkeypatch, isolated_env, fixed_clock, record_subprocess, gui_app):
    """Session App with dialogs, detailed log and the clock patched; env in tmp_path."""
    # ezbak_mod is requested before fixed_clock so the session import is not made under the patch.
    mod = ezbak_mod
    fixed_clock.inject(mod)
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.setenv("USERNAME", "tester")
    monkeypatch.setenv("USERDOMAIN", "TESTPC")
    monkeypatch.setenv("COMPUTERNAME", "TESTPC")
    monkeypatch.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "<PREFERRED>")
    dialogs, detailed = [], []

    def answer(name):
        def dialog(*args, **kwargs):
            dialogs.append([name] + list(args[:2]))
        return dialog

    for name in DIALOGS:
        monkeypatch.setattr(mod.Win11Dialog, name, answer(name), raising=False)
        monkeypatch.setattr(mod.messagebox, name, answer(name), raising=False)
    app = gui_app(mod)
    app.write_detailed_log = detailed.append  # instance attribute: gui_app's reset removes it
    while not app.message_queue.empty():
        app.message_queue.get_nowait()
    real_before = _real_schedules_state()
    try:
        yield {"mod": mod, "app": app, "rec": record_subprocess, "mp": monkeypatch,
               "env": isolated_env, "dialogs": dialogs, "detailed": detailed}
    finally:
        del app.write_detailed_log
    assert _real_schedules_state() == real_before


def _apply(case, ctx, tmp_path):
    for key, value in case.get("env", {}).items():
        if value is None:
            ctx["mp"].delenv(key, raising=False)
        else:
            ctx["mp"].setenv(key, value)
    for name in case.get("config", []):
        config_dir = ctx["env"]["APPDATA"] / "ezBAK" / "ezBAK_Schedule"
        config_dir.mkdir(parents=True, exist_ok=True)
        (config_dir / name).write_text('{"task_name": "pre-existing"}', encoding="utf-8")
    respond = dict(case.get("respond", {}))
    if respond.get("raises") == "timeout":
        respond["raises"] = subprocess.TimeoutExpired(["schtasks"], 180)
    if respond:
        ctx["rec"].respond("schtasks", **respond)


def _call(case, app, tmp_path):
    args = dict(case["args"])
    if case["op"] == "delete":
        return app.delete_scheduled_task(**args)
    args["dest"] = str(tmp_path / "dest")
    result = app.create_scheduled_task(**args)
    if case["op"] == "create_then_delete":
        app.delete_scheduled_task(args["task_name"])
    return result


def _read_xml_on_create(ctx):
    """Read the task XML when schtasks /create runs (a failed create removes it, fix-g)."""
    xml_reads, recorded_run = {}, subprocess.run

    def run(*args, **kwargs):
        argv = args[0] if args else kwargs.get("args")
        if argv[:2] == ["schtasks", "/create"]:
            xml_path = Path(argv[argv.index("/xml") + 1])
            xml_reads[str(xml_path)] = _read_file(xml_path)
        return recorded_run(*args, **kwargs)
    ctx["mp"].setattr(subprocess, "run", run)
    return xml_reads


def _run_case(case, ctx, tmp_path):
    _apply(case, ctx, tmp_path)
    xml_reads = _read_xml_on_create(ctx)
    app, out = ctx["app"], io.StringIO()
    result = {"return": None, "error": None}
    try:
        with contextlib.redirect_stdout(out):
            value = _call(case, app, tmp_path)
        result["return"] = value.splitlines() if isinstance(value, str) else value
    except Exception as exc:  # recorded: the raised error is part of the behavior
        cause = exc.__cause__
        result["error"] = {"type": type(exc).__name__, "lines": str(exc).splitlines(),
                           "cause": None if cause is None else type(cause).__name__}
    queue = []
    while not app.message_queue.empty():
        queue.append(list(app.message_queue.get_nowait()))
    result.update({
        "schtasks_calls": [c["args"][0][1] for c in ctx["rec"].calls],
        "task_xml": xml_reads,
        "files": _tree(tmp_path),
        "dialogs": ctx["dialogs"],
        "queue": queue,
        "prints": out.getvalue().splitlines(),
        "detailed_log": [line for msg in ctx["detailed"] for line in _fold_traceback(msg)],
    })
    return _normalizer(tmp_path)(result)


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


def _check(name, result):
    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden[name] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[name]


@pytest.mark.parametrize("case", sorted(CASES))
def test_schedule_case(case, sched, tmp_path):
    result = _run_case(CASES[case], sched, tmp_path)
    for path in list(result["files"]) + list(result["task_xml"]):
        assert path.startswith(("env/", "home/", "<TMP>")), path
    _check(case, result)


@pytest.mark.skipif(not os.path.isdir(r"C:\Users\Public"),
                    reason=r"C:\Users\Public missing: the round trip backs up user 'Public'")
@pytest.mark.filterwarnings("ignore:invalid escape sequence:DeprecationWarning")
def test_gui_json_to_cli(ezbak_mod, run_ezbak_main, block_subprocess, sched, tmp_path):
    # Order: block_subprocess (via run_ezbak_main) before record_subprocess (via sched), so the
    # recorder answers schtasks; ezbak_mod before fixed_clock (session import not under the patch).
    """GUI-written schedule JSON -> runpy CLI --schedule-config: the engine gets the JSON values."""
    dest = tmp_path / "dest"
    dest.mkdir()
    app = sched["app"]
    with contextlib.redirect_stdout(io.StringIO()):
        app.create_scheduled_task("ezBAK_RoundTrip", "Weekly", "07:15", "Public", str(dest),
                                  include_hidden=True, include_system=True, retention_count=5,
                                  log_retention_days=9, filters=CUSTOM_FILTERS)
    config_path = sched["env"]["APPDATA"] / "ezBAK" / "ezBAK_Schedule" / "ezBAK_RoundTrip.json"
    config = json.loads(config_path.read_text(encoding="utf-8"))
    assert config["filters"]["exclude"][0]["type"] == "NAME"  # D7: written as received

    run = run_ezbak_main.run(["--schedule-config", str(config_path)])

    assert block_subprocess == []
    assert [c["args"][0][1] for c in sched["rec"].calls] == ["/create"]  # the CLI ran nothing
    assert (run.end, run.code) == ("SystemExit", 0)
    assert len(run.engine_calls) == 1
    job = run.engine_calls[0]["job"]
    for key in ("user", "dest", "include_hidden", "include_system", "retention_count",
                "log_retention_days", "filters"):
        assert job[key] == config[key], key
    norm = run_ezbak_main.normalize
    _check("roundtrip_gui_json_to_cli", {
        "config": json.loads(norm(json.dumps(config))),
        "end": run.end, "code": run.code,
        "stdout": norm(run.stdout).splitlines(),
        "stderr": norm(run.stderr).splitlines(),
        "engine_job": json.loads(norm(json.dumps(job))),
    })
