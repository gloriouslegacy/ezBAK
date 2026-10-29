"""Characterization (golden) of the NAS dialogs before the P5 move (P5-0c), split out of
test_gui_dialogs.py (P5-fix #12b).

Pinned in tests/fixtures/gui_dialogs_golden.json (shared with test_gui_dialogs.py;
check_golden rewrites only its own keys):
- nas_*: the NAS connect/disconnect dialog trees and their input checks (no net use).

The password is dummy-pw only and must not reach stdout. Dialogs are Toplevels under the
session App (gui_app), driven through gui_harness.modal_hook. Variants: dark/en (the App's
own) and light/ko (theme and translator swapped with gui_harness.override).

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_gui_dialogs_nas.py
"""
import contextlib
import importlib.util
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.gui

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "gui_dialogs_golden.json"


def _load(name, file):
    spec = importlib.util.spec_from_file_location(name, FIXTURES / file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


gh = _load("ezbak_test_gui_harness_dialogs_nas", "gui_harness.py")
th = _load("ezbak_test_tool_harness_dialogs_nas", "tool_harness.py")

VARIANTS = {"dark_en": ("dark", "en"), "light_ko": ("light", "ko")}
SECRET = "dummy-pw"


@pytest.fixture
def dlg(ezbak_mod, isolated_env, gui_app, block_subprocess, tmp_path):
    """Session App plus a normalizer; afterwards no dialog is left, the App's instance dict
    and filters are as before and no external command was attempted."""
    app = gui_app(ezbak_mod)
    with gh.app_unchanged(app):
        yield SimpleNamespace(mod=ezbak_mod, app=app, tmp=tmp_path, env=isolated_env,
                              norm=gh.path_normalizer(tmp_path), dialogs=[])
    assert block_subprocess == []


def golden(name, result):
    th.check_golden(GOLDEN_PATH, name, result)


@contextlib.contextmanager
def variant(s, name):
    theme, lang = VARIANTS[name]
    with gh.override(s.app, theme=s.mod.Win11Theme(initial_mode=theme),
                     translator=s.mod.Translator(initial_lang=lang)):
        yield


def record_dialogs(s, monkeypatch, yes=None, folder=None):
    th.patch_dialogs(s.mod, monkeypatch, s.dialogs, yes=yes, folder=folder)


# --- NAS ------------------------------------------------------------------------------

def _nas(s, which, handler):
    with gh.override(s.app, find_free_drive_letter=lambda: "Q"), \
            gh.modal_hook(handler) as events:
        getattr(s.app, f"open_nas_{which}_dialog")()
    return events


@pytest.mark.parametrize("name", list(VARIANTS))
@pytest.mark.parametrize("which", ["connect", "disconnect"])
def test_nas_tree(dlg, monkeypatch, which, name):
    record_dialogs(dlg, monkeypatch)
    seen = {}

    def handler(window):
        seen["snapshot"] = gh.dialog_snapshot(window)
        gh.invoke_binding(window, "<Escape>")

    with variant(dlg, name):
        events = _nas(dlg, which, handler)
    golden(f"nas_{which}_tree_{name}", {"snapshot": seen["snapshot"], "events": events,
                                        "queue": th.drain(dlg.app), "dialogs": dlg.dialogs})


NAS_INPUTS = {  # entries in order: connect (unc, drive, user, password), disconnect (target)
    "connect": [["", "Q", "", ""], ["//srv/share", "Q", "", ""], ["\\\\srv\\share", "Q", "user1", ""],
                ["\\\\srv\\share", "12", "", ""], ["\\\\srv\\share", " :", "", ""]],
    "disconnect": [[""], ["   "]],
}


@pytest.mark.parametrize("which", list(NAS_INPUTS))
def test_nas_input_checks(dlg, monkeypatch, capsys, which):
    """Inputs the OK button rejects; the dialog stays open each time, then Escape."""
    record_dialogs(dlg, monkeypatch)
    steps = []

    def handler(window):
        entries = gh.descendants(window, dlg.mod.tk.Entry)
        for values in NAS_INPUTS[which]:
            for entry, value in zip(entries, values):
                gh.type_into(entry, value)
            dlg.dialogs.clear()
            gh.button(window, "OK").invoke()
            steps.append([values, window.winfo_exists(), list(dlg.dialogs),
                          [str(e.cget("state")) for e in entries]])
        gh.invoke_binding(window, "<Escape>")

    events = _nas(dlg, which, handler)
    assert SECRET not in capsys.readouterr().out
    golden(f"nas_{which}_input_checks", {"steps": steps, "events": events,
                                         "queue": th.drain(dlg.app)})


def test_nas_map_toggles_drive_entry(dlg):  # P5-fix h-j: drive entry follows the map checkbox
    def handler(window):  # untick then tick; modal_hook closes the window afterwards
        check = gh.descendants(window, dlg.mod.tk.Checkbutton)[0]
        drive = gh.descendants(window, dlg.mod.tk.Entry)[1]
        assert [(check.invoke(), str(drive.cget("state")))[1] for _ in range(2)] == ["disabled", "normal"]
    _nas(dlg, "connect", handler)
