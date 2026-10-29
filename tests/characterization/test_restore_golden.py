"""Characterization (golden) of a user-data restore from a temp backup folder (T0-6).

Path: App.start_restore_thread -> App.run_restore (real App from gui_app, window
withdrawn; dialogs and the folder picker patched, worker thread run inline).
There is no Headless restore: the CLI / scheduled task only backs up.

The backup folder is made by the real GUI backup (App.run_backup of the copy imported
without pywin32, the fullest output in backup_golden.json) from the user tree in
tests/fixtures/user_tree.py. The source profile is then moved away and an empty
<USERS_ROOT>/tester is the restore target. Both the target profile's AppData/Local and
the running account's %LOCALAPPDATA% hold an empty Chrome and Edge "User Data/Default"
folder: bookmark restore writes under the target profile (fix-d) and the
%LOCALAPPDATA% record shows the running account is left untouched.

Two ways to pick the folder are fixed:
- parent: the folder holding the backup folders (run_restore lists '*_backup_*')
- direct: the backup folder itself

Recorded per run: the restored tree, its file count, the restore job log
(<TS>, <TMP>, <AGE> replaced), the files under %LOCALAPPDATA% after restore and the
names of external commands started (in order). Expected values are today's results.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_restore_golden.py
"""
import importlib.util
import json
import os
import re
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = [pytest.mark.windows, pytest.mark.gui]

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "restore_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_user_tree_restore",
                                               FIXTURES / "user_tree.py")
user_tree = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(user_tree)

TS_RE = re.compile(r"\d{4}-\d{2}-\d{2}[ _]\d{2}:?\d{2}:?\d{2}")
AGE_RE = re.compile(r"Age: -?\d+\.\d+ days")
DIALOGS = ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel")


class _InlineThread:
    """Stand-in for threading.Thread that runs the target on start(), in this thread."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target, self._args, self._kwargs = target, args, kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


def _normalize(text, tmp_path):
    text = TS_RE.sub("<TS>", text.replace(str(tmp_path), "<TMP>"))
    return AGE_RE.sub("Age: <AGE> days", text)


def _walk(root):
    """Sorted relative entries under root ('dir/' for folders) and the file count."""
    entries, files = [], 0
    for dirpath, dirnames, filenames in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        prefix = "" if rel == "." else rel.replace(os.sep, "/") + "/"
        entries += [prefix + d + "/" for d in dirnames]
        entries += [prefix + f for f in filenames]
        files += len(filenames)
    return sorted(entries), files


def _patch_dialogs(mod, monkeypatch, picked_folder, dialogs):
    def answer(name):
        def dialog(*args, **kwargs):
            dialogs.append((name,) + tuple(args[:2]))
            return True
        return dialog

    for name in DIALOGS:
        monkeypatch.setattr(mod.Win11Dialog, name, answer(name), raising=False)
        monkeypatch.setattr(mod.messagebox, name, answer(name), raising=False)
    monkeypatch.setattr(mod.filedialog, "askdirectory", lambda **kwargs: str(picked_folder))
    monkeypatch.setattr(mod, "threading",
                        SimpleNamespace(Thread=_InlineThread, Lock=threading.Lock))


def _drain(app):
    messages = []
    while not app.message_queue.empty():
        messages.append(app.message_queue.get_nowait())
    return messages


def _make_backup(mod, dest, monkeypatch, gui_app):
    """Real GUI backup of <USERS_ROOT>/tester into dest; returns the backup folder."""
    dialogs = []
    _patch_dialogs(mod, monkeypatch, dest, dialogs)
    app = gui_app(mod)
    app.user_var.set(user_tree.USER_NAME)
    app.start_backup_thread()
    assert [m for m in _drain(app) if m[0] == "show_error"] == []
    assert dialogs == []
    folders = [p for p in dest.iterdir() if p.is_dir() and "_backup_" in p.name]
    assert len(folders) == 1, folders
    return folders[0]


def _run_restore(mod, picked_folder, monkeypatch, gui_app):
    dialogs = []
    _patch_dialogs(mod, monkeypatch, picked_folder, dialogs)
    app = gui_app(mod)
    app.user_var.set(user_tree.USER_NAME)
    app.start_restore_thread()
    messages = _drain(app)
    assert [m for m in messages if m[0] == "show_error"] == []
    # Only the overwrite confirmation; same user, so no cross-user prompt.
    assert [d[0] for d in dialogs] == ["askyesno"]


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


@pytest.mark.parametrize("variant", ["pywin32", "nopywin32"])
@pytest.mark.parametrize("selection", ["parent", "direct"])
def test_restore_golden(selection, variant, ezbak_mod, ezbak_mod_nopywin32, tmp_path,
                        monkeypatch, isolated_env, user_root, fixed_clock, block_subprocess,
                        gui_app):
    mod = {"pywin32": ezbak_mod, "nopywin32": ezbak_mod_nopywin32}[variant]
    if variant == "pywin32" and not mod._have_pywin32:
        pytest.skip("pywin32 not installed")
    backup_mod = ezbak_mod_nopywin32
    for m in {mod, backup_mod}:
        user_root.inject(m)
        fixed_clock.inject(m)

    # 1. Backup folder from the real GUI backup.
    home = Path(user_tree.build_user_tree(str(user_root.path / user_tree.USER_NAME)))
    monkeypatch.setenv("LOCALAPPDATA", str(home / "AppData" / "Local"))
    dest = tmp_path / "dest"
    dest.mkdir()
    backup_folder = _make_backup(backup_mod, dest, monkeypatch, gui_app)

    # 2. Target profile and a clean %LOCALAPPDATA%, both with browser profile folders.
    home.rename(tmp_path / "source_home")
    home.mkdir()
    local = isolated_env["LOCALAPPDATA"]
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    for browser in (("Google", "Chrome"), ("Microsoft", "Edge")):
        local.joinpath(*browser, "User Data", "Default").mkdir(parents=True)
        home.joinpath("AppData", "Local", *browser, "User Data", "Default").mkdir(parents=True)

    # 3. Restore.
    picked = dest if selection == "parent" else backup_folder
    _run_restore(mod, picked, monkeypatch, gui_app)

    tree, file_count = _walk(home)
    logs = sorted(p for p in dest.iterdir() if p.name.endswith(".log") and "_restore_" in p.name)
    assert len(logs) == 1, logs
    log_lines = _normalize(logs[0].read_text(encoding="utf-8"), tmp_path).splitlines()
    result = {
        "tree": tree,
        "file_count": file_count,
        "log": log_lines,
        "localappdata": _walk(local)[0],
        "commands": [name for name, _argv in block_subprocess],
    }

    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden.setdefault(selection, {})[variant] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[selection][variant]
