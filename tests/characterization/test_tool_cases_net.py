"""Characterization (golden) of the app-list export and NAS cases (P4-0b), on the untouched ezBAK.py.

Covered: start_winget_export_thread -> run_winget_export (worker inline): folder cancelled,
primary export, basic export, `winget list --accept-source-agreements`, plain `winget list`
success and failure, everything raising, Popen FileNotFoundError, list TimeoutExpired.
open_nas_connect_dialog / open_nas_disconnect_dialog: cancel, success, rc != 0 (stderr,
stdout, code text), UNC without user, bad drive letter, and net use raising (after g-k the
except branch logs, shows one error dialog and returns; nothing leaves the method).

argv_golden.json already fixes the argv and kwargs; here the rest is fixed: the GUI queue in
order, the job log (<TS>, <TMP>, frames -> <FRAMES>) and files made (winget), the lines the
NAS code hands to write_detailed_log (no log file is open there, so they are dropped), the
dialogs (name, title, message; askstring answers are not recorded) and the command texts.

Security: the NAS cases use the dummy password "dummy-pw". It must not appear in caplog
(every logger at DEBUG), print output, the queue, dialogs, detailed-log lines or any file
under the temp root; in the golden it may appear only inside "commands" (argv), as in
argv_golden.json.

Nothing real runs: record_subprocess answers every command (no winget, no net use, no
network), os.startfile is recorded, the clock is fixed_clock, used drive letters are fixed.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_tool_cases_net.py
"""
import importlib.util
import json
import logging
import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = [pytest.mark.windows, pytest.mark.gui]

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "tool_net_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_tool_harness_net", FIXTURES / "tool_harness.py")
harness = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(harness)

NAS_UNC = "\\\\nas\\share"
NAS_USER = "nasuser"
NAS_PW = "dummy-pw"  # test value only, never a real secret


def _commands(rec):
    return [[c["api"], c["args"][0] if isinstance(c["args"][0], str) else " ".join(c["args"][0])]
            for c in rec.calls]


@pytest.fixture
def gui(ezbak_mod, tmp_path, monkeypatch, isolated_env, fixed_clock, record_subprocess,
        gui_app, capsys, caplog):
    # ezbak_mod is requested before fixed_clock so the session import is not made under the patch.
    fixed_clock.inject(ezbak_mod)
    app = gui_app(ezbak_mod)
    harness.inline_threads(ezbak_mod, monkeypatch)
    caplog.set_level(logging.DEBUG)
    for name in list(logging.root.manager.loggerDict):
        caplog.set_level(logging.DEBUG, logger=name)
    env = SimpleNamespace(mod=ezbak_mod, app=app, rec=record_subprocess, mp=monkeypatch,
                          tmp=tmp_path, dialogs=[], capsys=capsys, caplog=caplog,
                          norm=harness.normalizer(tmp_path, REPO_ROOT))
    harness.drain(app)
    capsys.readouterr()
    yield env
    app.close_log_file()
    app.reset_progress_bar()
    app.set_buttons_state("normal")
    harness.drain(app)


# ---------------------------------------------------------------- winget
# rules: (match, respond kwargs) in order; the first match answers (see SubprocessRecorder).

_PRIMARY_FAILS = [("--version", {"stdout": "v1.9.0\n"}),
                  ("--include-versions", {"returncode": 1, "stdout": "partial\n",
                                          "stderr": "export error\n"})]

WINGET = {
    "winget_folder_cancel": {"folder": False, "rules": []},
    "winget_primary_ok": {"rules": [("--version", {"stdout": "v1.9.0\n"}),
                                    ("--include-versions", {"stdout": "Exporting...\nDone\n"})]},
    "winget_basic_ok": {"rules": _PRIMARY_FAILS + [("winget export", {})]},
    "winget_list_ok": {"rules": _PRIMARY_FAILS + [
        ("winget export", {"returncode": 2, "stderr": "basic error\n"}),
        ("winget list --accept", {"stdout": "Name  Id\nApp   Vendor.App\n"})]},
    "winget_simple_ok": {"rules": _PRIMARY_FAILS + [
        ("winget export", {"returncode": 2}),
        ("winget list --accept", {"returncode": 3, "stderr": "list error\n"}),
        ("winget list", {"stdout": "Name Id\n"})]},
    "winget_simple_fail": {"rules": _PRIMARY_FAILS + [
        ("winget export", {"returncode": 2}),
        ("winget list --accept", {"returncode": 3}),
        ("winget list", {"returncode": 4, "stderr": "simple error\n"})]},
    "winget_all_raise": {"rules": _PRIMARY_FAILS + [
        ("winget export", {"raises": OSError("basic raised")}),
        ("winget list --accept", {"raises": OSError("list raised")}),
        ("winget list", {"raises": OSError("simple raised")})]},
    "winget_popen_not_found": {"rules": [
        ("--version", {"raises": FileNotFoundError(2, "The system cannot find the file specified")}),
        ("--include-versions", {"raises": FileNotFoundError(2, "The system cannot find the file specified")})]},
    "winget_list_timeout": {"rules": _PRIMARY_FAILS + [
        ("winget export", {"returncode": 2}),
        ("winget list --accept", {"raises": subprocess.TimeoutExpired(
            ["winget", "list", "--accept-source-agreements"], 300)}),
        ("winget list", {"stdout": "Name Id\n"})]},
}


@pytest.mark.parametrize("case", list(WINGET))
def test_winget_case(case, gui):
    spec = WINGET[case]
    apps = gui.tmp / "apps"
    apps.mkdir()
    for match, kwargs in spec["rules"]:
        gui.rec.respond(match, **kwargs)
    folder = None if spec.get("folder") is False else str(apps)
    harness.patch_dialogs(gui.mod, gui.mp, gui.dialogs, folder=folder)
    gui.app.start_winget_export_thread()
    gui.app.close_log_file()
    out = gui.capsys.readouterr().out.splitlines()
    harness.check_golden(GOLDEN_PATH, case, gui.norm({
        "stdout": out, "dialogs": gui.dialogs, "queue": harness.drain(gui.app),
        "tree": harness.walk(apps), "logs": harness.read_logs(apps),
        "commands": _commands(gui.rec)}))


# ---------------------------------------------------------------- NAS
# dialog: connect/disconnect; fill: Entry values (connect: UNC, drive, user, password),
# None -> Cancel; rules as above; yes: askyesno answers; strings: askstring answers;
# runs/errors: expected number of net use runs and showerror dialogs.

NAS = {
    "nas_connect_cancel": {"dialog": "connect", "fill": None},
    "nas_connect_user_ok": {"dialog": "connect", "fill": [NAS_UNC, "z", NAS_USER, NAS_PW],
                            "yes": [True]},
    "nas_connect_unc_no_user": {"dialog": "connect", "fill": [NAS_UNC, "y:", "", ""],
                                "yes": [False]},
    "nas_connect_rc_err": {"dialog": "connect", "fill": [NAS_UNC, "z", NAS_USER, NAS_PW],
                           "rules": [("net use", {"returncode": 2, "stdout": "out text\r\n",
                                                  "stderr": "System error 53 has occurred.\r\n"})]},
    "nas_connect_rc_out": {"dialog": "connect", "fill": [NAS_UNC, "z", NAS_USER, NAS_PW],
                           "rules": [("net use", {"returncode": 1,
                                                  "stdout": "The network path was not found.\r\n"})]},
    "nas_connect_rc_code": {"dialog": "connect", "fill": [NAS_UNC, "z", NAS_USER, NAS_PW],
                            "rules": [("net use", {"returncode": 2})]},
    "nas_connect_bad_drive": {"dialog": "connect", "fill": [NAS_UNC, "1", NAS_USER, NAS_PW]},
    "nas_connect_retry_ok": {"dialog": "connect", "fill": [NAS_UNC, "x", NAS_USER, NAS_PW],
                             "rules": [("net use", {"raises": OSError("net use failed"), "times": 2})]},
    "nas_connect_retry_raises": {"dialog": "connect", "fill": [NAS_UNC, "x", NAS_USER, NAS_PW],
                                 "rules": [("net use", {"raises": OSError("net use failed"), "times": 4})],
                                 "runs": 2, "errors": 1},
    "nas_disconnect_cancel": {"dialog": "disconnect", "fill": None},
    "nas_disconnect_drive_ok": {"dialog": "disconnect", "fill": ["z"]},
    "nas_disconnect_rc_err": {"dialog": "disconnect", "fill": [NAS_UNC],
                              "rules": [("net use", {"returncode": 2,
                                                     "stderr": "The network connection could not be found.\r\n"})]},
    "nas_disconnect_rc_code": {"dialog": "disconnect", "fill": ["z:"],
                               "rules": [("net use", {"returncode": 1})]},
    "nas_disconnect_bad_drive": {"dialog": "disconnect", "fill": ["zz"]},
    "nas_disconnect_retry_raises": {"dialog": "disconnect", "fill": ["z"],
                                    "rules": [("net use", {"raises": OSError("net use failed"), "times": 4})],
                                    "runs": 2, "errors": 1},
}


def _assert_no_secret(env, result, out, err, detailed):
    """The dummy password is nowhere but inside the recorded argv."""
    outside_argv = {k: v for k, v in result.items() if k != "commands"}
    places = {
        "golden outside argv": json.dumps(outside_argv, ensure_ascii=False),
        "caplog": env.caplog.text + "\n".join(str(r.args) for r in env.caplog.records),
        "print": out + err,
        "detailed log": "\n".join(detailed),
        "queue/dialogs": json.dumps([result["queue"], result["dialogs"]], ensure_ascii=False),
    }
    for place, text in places.items():
        assert NAS_PW not in text, place
    assert harness.files_containing(env.tmp, NAS_PW) == []


@pytest.mark.parametrize("case", list(NAS))
def test_nas_case(case, gui):
    spec = NAS[case]
    for match, kwargs in spec.get("rules", []):
        gui.rec.respond(match, **kwargs)
    harness.patch_dialogs(gui.mod, gui.mp, gui.dialogs, yes=spec.get("yes"),
                          strings=spec.get("strings"))
    gui.mp.setattr(gui.app, "get_used_drive_letters", lambda: {"C", "D"}, raising=False)
    if spec["fill"] is None:
        harness.fill_dialog(gui.mod, gui.app, gui.mp, [], press=1)
    else:
        harness.fill_dialog(gui.mod, gui.app, gui.mp, spec["fill"])
    detailed = []
    gui.mp.setattr(gui.app, "write_detailed_log", detailed.append, raising=False)
    method = getattr(gui.app, f"open_nas_{spec['dialog']}_dialog")
    raised = None
    method()  # g-k: no exception may leave the method
    captured = gui.capsys.readouterr()
    result = gui.norm({
        "raised": raised, "detailed": detailed, "dialogs": gui.dialogs,
        "queue": harness.drain(gui.app), "commands": _commands(gui.rec)})
    _assert_no_secret(gui, result, captured.out, captured.err, detailed)
    if "runs" in spec:
        assert sum("net use" in c[1] for c in result["commands"]) == spec["runs"]
        assert [d[0] for d in result["dialogs"]].count("showerror") == spec["errors"]
    harness.check_golden(GOLDEN_PATH, case, result)
