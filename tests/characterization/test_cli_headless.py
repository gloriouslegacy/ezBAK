"""Characterization (golden) of the untouched ezBAK.py __main__ block: the Headless CLI (P3-0b).

ezBAK.py is executed with runpy as __main__ (run_ezbak_main in conftest.py). The backup
engine ezbak.core.run_backup is a stub, so core_run_backup still prepares the backup
folder and writes the job log header, but nothing is copied. The user is `Public`
(C:\\Users\\Public): runpy executes the module afresh, so USERS_ROOT cannot be injected.
UAC relaunch (ShellExecuteW) and App() (Tk) are recorded, never performed.

Recorded per case: how the run ended (SystemExit code / os._exit code / App() / return),
stdout and stderr (temp root -> <TMP>, repo -> <REPO>, python -> <PY>, timestamps -> <TS>,
traceback frames folded into "  <FRAMES>"), the stub engine's arguments, the
ShellExecuteW calls, and the destination tree with the job log text.
Today's behavior is recorded as-is, odd or not.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_cli_headless.py
"""
import json
import os
from pathlib import Path

import pytest

pytestmark = [
    pytest.mark.windows,
    # runpy compiles ezBAK.py afresh: its docstring escape warning (line ~3363) is not ours
    pytest.mark.filterwarnings("ignore:invalid escape sequence:DeprecationWarning"),
    pytest.mark.skipif(not os.path.isdir(r"C:\Users\Public"),
                       reason=r"C:\Users\Public missing: the CLI golden backs up user 'Public'"),
]

GOLDEN_PATH = Path(__file__).resolve().parent.parent / "fixtures" / "cli_golden.json"
USER = "Public"
MISSING_USER = "NoSuchUser_ezbak_cli"
CUSTOM_FILTERS = {"include": [], "exclude": [{"type": "name", "pattern": "*.tmp"}]}
# A list instead of the {"include", "exclude"} dict: parsed, then the engine setup fails.
LIST_FILTERS = [{"type": "exclude", "pattern": "*.tmp"}]


def _config(tmp_path, dest, user=USER):
    return {
        "task_name": "ezBAK_Test", "user": user, "dest": str(dest), "schedule": "DAILY",
        "time_str": "09:00", "monthly_day": 1, "include_hidden": True, "include_system": False,
        "retention_count": 3, "log_retention_days": 15, "filters": CUSTOM_FILTERS,
        "created_at": "2024-01-01 00:00:00",
    }


def _write_config(tmp_path, content):
    path = tmp_path / "config" / "ezBAK_Test.json"
    path.parent.mkdir(exist_ok=True)
    text = content if isinstance(content, str) else json.dumps(content, indent=2)
    path.write_text(text, encoding="utf-8")
    return path


def case_config_ok(tmp_path, dest):
    return {"argv": ["--schedule-config", str(_write_config(tmp_path, _config(tmp_path, dest)))]}


def case_config_missing(tmp_path, dest):
    return {"argv": ["--schedule-config", str(tmp_path / "config" / "absent.json")]}


def case_config_bad_json(tmp_path, dest):
    return {"argv": ["--schedule-config", str(_write_config(tmp_path, '{"user": "Public",'))]}


def case_config_user_missing(tmp_path, dest):
    config = _config(tmp_path, dest, user=MISSING_USER)
    return {"argv": ["--schedule-config", str(_write_config(tmp_path, config))]}


def case_direct_ok(tmp_path, dest):
    return {"argv": ["--user", USER, "--dest", str(dest), "--include-system",
                     "--retention", "4", "--log-retention", "7", "--log-folder", "mylogs",
                     "--filters", json.dumps(CUSTOM_FILTERS)]}


def case_direct_engine_error(tmp_path, dest):
    return {"argv": ["--user", USER, "--dest", str(dest)],
            "engine_error": RuntimeError("engine failed (stub)")}


def case_direct_bad_filters(tmp_path, dest):
    return {"argv": ["--user", USER, "--dest", str(dest), "--filters", "{not json"]}


def case_direct_filters_list(tmp_path, dest):
    return {"argv": ["--user", USER, "--dest", str(dest), "--filters", json.dumps(LIST_FILTERS)]}


def case_auto_backup(tmp_path, dest):
    return {"argv": ["--auto-backup", "--user", USER, "--dest", str(dest)]}


def case_retention_not_int(tmp_path, dest):
    return {"argv": ["--user", USER, "--dest", str(dest), "--retention", "abc"]}


def case_user_only_not_admin(tmp_path, dest):
    return {"argv": ["--user", USER], "is_admin": 0}


def case_user_only_admin(tmp_path, dest):
    return {"argv": ["--user", USER], "is_admin": 1}


CASES = {name[len("case_"):]: fn for name, fn in globals().items() if name.startswith("case_")}


def _fold_tracebacks(text):
    """Keep the 'Traceback' header and the exception line; frames (file/line) become one line."""
    out, in_frames = [], False
    for line in text.splitlines():
        if line == "Traceback (most recent call last):":
            out.extend([line, "  <FRAMES>"])
            in_frames = True
        elif in_frames and line.startswith(" "):
            continue
        else:
            in_frames = False
            out.append(line)
    return out


def _tree(dest, runner):
    entries = {}
    for path in sorted(dest.rglob("*")):
        rel = runner.normalize(path.relative_to(dest).as_posix())
        if path.is_dir():
            entries[rel + "/"] = None
        else:
            text = path.read_text(encoding="utf-8")
            entries[rel] = [runner.normalize(line) for line in text.splitlines()]
    return entries


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


@pytest.mark.parametrize("case", sorted(CASES))
def test_cli_case(case, tmp_path, run_ezbak_main, block_subprocess):
    dest = tmp_path / "dest"
    dest.mkdir()
    opts = CASES[case](tmp_path, dest)
    run = run_ezbak_main.run(opts["argv"], engine_error=opts.get("engine_error"),
                             is_admin=opts.get("is_admin", 0))
    norm = run_ezbak_main.normalize
    result = {
        "argv": [norm(a) for a in opts["argv"]],
        "end": run.end,
        "code": run.code,
        "stdout": _fold_tracebacks(norm(run.stdout)),
        "stderr": _fold_tracebacks(norm(run.stderr)),
        "engine_calls": json.loads(norm(json.dumps(run.engine_calls))),
        "shell_execute": json.loads(norm(json.dumps(run.shell_calls))),
        "dest": _tree(dest, run_ezbak_main),
    }

    assert block_subprocess == []
    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden[case] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[case]


def test_harness_restores_global_state(run_ezbak_main, ezbak_mod, tmp_path):
    """After a runpy run the shared ezbak modules, the engine, sys.argv, os._exit,
    tkinter.Tk and ctypes shell32 are the ones from before (other tests share them)."""
    import ctypes
    import sys
    import tkinter
    import ezbak.core
    before = {
        "argv": list(sys.argv), "exit": os._exit, "tk_init": tkinter.Tk.__init__,
        "engine": ezbak.core.run_backup, "shell32": ctypes.windll.shell32,
        "modules": {n: m for n, m in sys.modules.items() if n in ("ezBAK", "ezbak") or n.startswith("ezbak.")},
    }
    with pytest.MonkeyPatch.context() as mp:
        run = type(run_ezbak_main)(mp, tmp_path).run(["--user", USER, "--dest", str(tmp_path)])
    assert run.end == "SystemExit" and len(run.engine_calls) == 1
    assert sys.argv == before["argv"] and os._exit is before["exit"]
    assert tkinter.Tk.__init__ is before["tk_init"]
    assert ezbak.core.run_backup is before["engine"] and ezbak_mod.core_engine_run_backup is before["engine"]
    assert ctypes.windll.shell32 is before["shell32"]
    after = {n: m for n, m in sys.modules.items() if n in ("ezBAK", "ezbak") or n.startswith("ezbak.")}
    assert after == before["modules"]
