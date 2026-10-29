"""Characterization (golden) of the full browser-profile backup cases (P4-0c), on the untouched ezBAK.py.

Covered: start_browser_profiles_backup_thread -> run_browser_profiles_backup (worker inline):
no user selected (same warning as the regular backup), folder cancelled, the confirmation
answered 'No' (nothing made, no log), no source at all, Chrome + Edge User Data + Firefox Profiles + profiles.ini,
a junction inside User Data (skipped by is_reparse), not enough free space, shutil.disk_usage
failing (free space then comes from winsys.sysinfo's GetDiskFreeSpaceExW step, fixed here).
Each case runs twice: with pywin32 (win32api.GetFileAttributes) and without it
(ezbak_mod_nopywin32, the ctypes GetFileAttributesW branch).

Fixed here: the return value, the GUI queue in order (consecutive update_progress folded to
the last one, as in test_backup_cases), the job log (<TS>, <TMP>), the tree made under the
destination, the dialogs and print output.

The sources are the selected user's <USERS_ROOT>/<user>/AppData/{Local,Roaming} (g-p, P4-fix
F4=C), USERS_ROOT pointed at a temp tree by user_root; isolated_env keeps the App's own
%APPDATA% in temp too. No real user profile is read or written. Free space is never the real disk:
shutil.disk_usage is patched in every case. No external command runs (block_subprocess).

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_tool_cases_profiles.py
"""
import importlib.util
import os
import shutil
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = [pytest.mark.windows, pytest.mark.gui]

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "tool_profiles_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_tool_harness_profiles", FIXTURES / "tool_harness.py")
harness = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(harness)

_tree_spec = importlib.util.spec_from_file_location("ezbak_test_user_tree_profiles", FIXTURES / "user_tree.py")
user_tree = importlib.util.module_from_spec(_tree_spec)
_tree_spec.loader.exec_module(user_tree)

BIG_FREE = 10**12
USER = user_tree.USER_NAME


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _chrome(env):
    base = env["LOCALAPPDATA"] / "Google" / "Chrome" / "User Data"
    _write(base / "Local State", '{"chrome": 1}')
    _write(base / "Default" / "Bookmarks", '{"roots": {}}')
    _write(base / "Default" / "Login Data", "chrome-login")
    return base


def _all_three(env):
    _chrome(env)
    edge = env["LOCALAPPDATA"] / "Microsoft" / "Edge" / "User Data"
    _write(edge / "Local State", '{"edge": 1}')
    _write(edge / "Profile 1" / "Preferences", '{"p": 1}')
    firefox = env["APPDATA"] / "Mozilla" / "Firefox"
    _write(firefox / "Profiles" / "abcd.default-release" / "prefs.js", "user_pref(1);")
    _write(firefox / "Profiles" / "abcd.default-release" / "cookies.sqlite", "cookies")
    _write(firefox / "profiles.ini", "[Profile0]\nPath=Profiles/abcd.default-release\n")


def _junction(env):
    base = _chrome(env)
    outside = env["root"].parent / "outside"
    _write(outside / "secret.txt", "must not be copied")
    user_tree.make_junction(str(base / "Linked"), str(outside))


# setup(env) prepares the sources; folder False: the folder dialog is cancelled;
# user: the selected user ("" none); yes: askyesno answers (default [True], the confirmation);
# free: shutil.disk_usage free bytes, or "raise" (then sysinfo's API step answers api_free).
CASES = {
    "profiles_no_user": {"setup": _all_three, "user": ""},
    "profiles_folder_cancel": {"folder": False},
    "profiles_confirm_no": {"setup": _all_three, "yes": [False]},
    "profiles_no_source": {},
    "profiles_all_three": {"setup": _all_three},
    "profiles_junction_skipped": {"setup": _junction},
    "profiles_insufficient_space": {"setup": _all_three, "free": 10},
    "profiles_disk_usage_fail": {"setup": _all_three, "free": "raise", "api_free": BIG_FREE},
}


def _user_env(user_root):
    """The selected user's profile under the injected USERS_ROOT (sources of g-p)."""
    app_data = user_root.path / USER / "AppData"
    (app_data / "Local").mkdir(parents=True)
    (app_data / "Roaming").mkdir(parents=True)
    return {"root": user_root.path, "LOCALAPPDATA": app_data / "Local", "APPDATA": app_data / "Roaming"}


def _fold_progress(messages):
    folded = []
    for message in messages:
        if folded and message[0] == "update_progress" and folded[-1][0] == "update_progress":
            folded[-1] = message
        else:
            folded.append(message)
    return folded


def _patch_free_space(mod, monkeypatch, spec):
    free = spec.get("free", BIG_FREE)
    if free == "raise":
        def disk_usage(path):
            raise OSError(f"disk_usage failed for {path}")
    else:
        def disk_usage(path):
            return SimpleNamespace(free=free)
    monkeypatch.setattr(shutil, "disk_usage", disk_usage)
    if "api_free" in spec:
        # Patch the globals the function really reads: the nopywin32 copy of ezbak is popped
        # from sys.modules, so sys.modules[__module__] would be the other (pywin32) copy.
        monkeypatch.setitem(mod.winsys_free_space.__globals__, "_api_free_space",
                            lambda path: spec["api_free"])


@pytest.mark.parametrize("variant", ["pywin32", "nopywin32"])
@pytest.mark.parametrize("case", list(CASES))
def test_profiles_case(case, variant, ezbak_mod, ezbak_mod_nopywin32, tmp_path, monkeypatch,
                       isolated_env, user_root, fixed_clock, block_subprocess, gui_app, capsys):
    mod = {"pywin32": ezbak_mod, "nopywin32": ezbak_mod_nopywin32}[variant]
    if variant == "pywin32" and not mod._have_pywin32:
        pytest.skip("pywin32 not installed")
    spec = CASES[case]
    fixed_clock.inject(mod)
    user_root.inject(mod)
    env = _user_env(user_root)
    if "setup" in spec:
        spec["setup"](env)
    dest = tmp_path / "dest"
    dest.mkdir()
    dialogs = []
    harness.patch_dialogs(mod, monkeypatch, dialogs, yes=spec.get("yes", [True]),
                          folder=None if spec.get("folder") is False else str(dest))
    _patch_free_space(mod, monkeypatch, spec)
    app = gui_app(mod)
    app.progress_throttle_secs = 1e9  # restored by gui_app's reset
    app.user_var.set(spec.get("user", USER))  # Tk variables restored by gui_app's reset
    harness.inline_threads(mod, monkeypatch)
    harness.drain(app)
    capsys.readouterr()
    try:
        returned = app.start_browser_profiles_backup_thread()
    finally:
        app.close_log_file()
    out = capsys.readouterr().out.splitlines()
    norm = harness.normalizer(tmp_path, REPO_ROOT)
    result = norm({
        "returned": repr(returned), "stdout": out, "dialogs": dialogs,
        "queue": _fold_progress(harness.drain(app)),
        "tree": harness.walk(dest), "logs": harness.read_logs(dest)})
    assert block_subprocess == []
    assert os.path.isdir(env["LOCALAPPDATA"])  # sources stay in the temp tree
    if case in ("profiles_no_user", "profiles_confirm_no", "profiles_folder_cancel"):
        assert result["tree"] == [] and result["logs"] == {}  # nothing made, no job log
    harness.check_golden(GOLDEN_PATH, f"{case}-{variant}", result)
