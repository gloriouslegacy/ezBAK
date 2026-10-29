"""Characterization (golden) of a full user-data backup into a temp folder.

- gui:      App.start_backup_thread -> App.run_backup (real App, window withdrawn;
            folder dialog patched, worker thread run inline)
- headless: core_run_backup (what the scheduled task calls)

Both run against the user tree in tests/fixtures/user_tree.py with default settings
(hidden/system excluded, default filters, keep 2 backups, keep logs 30 days), under
the normal import (pywin32) and the copy imported without pywin32.

Recorded per run: every folder/file created under the destination, the number of
files inside the backup folder, and the job log with timestamps replaced by <TS>,
the temp root by <TMP> and the log age by <AGE>. Expected values are today's
results; GUI/Headless differences are kept as-is.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_backup_golden.py
"""
import importlib.util
import json
import os
import re
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.windows

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "backup_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_user_tree", FIXTURES / "user_tree.py")
user_tree = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(user_tree)

TS_RE = re.compile(r"\d{4}-\d{2}-\d{2}[ _]\d{2}:?\d{2}:?\d{2}")
# Log-cleanup "Age" compares real time.time() with the log's mtime: prints 0.0 or -0.0.
AGE_RE = re.compile(r"Age: -?\d+\.\d+ days")


class _InlineThread:
    """Stand-in for threading.Thread that runs the target on start(), in this thread."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target, self._args, self._kwargs = target, args, kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


def _normalize(text, tmp_path):
    text = TS_RE.sub("<TS>", text.replace(str(tmp_path), "<TMP>"))
    return AGE_RE.sub("Age: <AGE> days", text)


def _snapshot(dest, tmp_path):
    entries, logs, file_count = [], [], 0
    for dirpath, dirnames, filenames in os.walk(dest):
        rel = os.path.relpath(dirpath, dest)
        prefix = "" if rel == "." else rel.replace(os.sep, "/") + "/"
        for d in dirnames:
            entries.append(prefix + d + "/")
        for f in filenames:
            entries.append(prefix + f)
            if prefix:
                file_count += 1
            elif f.endswith(".log"):
                logs.append(os.path.join(dirpath, f))
    assert len(logs) == 1, f"expected one job log in {dest}, found {logs}"
    with open(logs[0], encoding="utf-8") as f:
        log_lines = _normalize(f.read(), tmp_path).splitlines()
    return {
        "tree": sorted(_normalize(e, tmp_path) for e in entries),
        "file_count": file_count,
        "log": log_lines,
    }


def _run_gui(mod, dest, monkeypatch, gui_app):
    dialogs = []

    def record_dialog(*args, **kwargs):
        dialogs.append(args[:2])

    for name in ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel"):
        monkeypatch.setattr(mod.Win11Dialog, name, record_dialog, raising=False)
        monkeypatch.setattr(mod.messagebox, name, record_dialog, raising=False)
    monkeypatch.setattr(mod.filedialog, "askdirectory", lambda **kwargs: str(dest))

    app = gui_app(mod)
    app.user_var.set(user_tree.USER_NAME)
    monkeypatch.setattr(mod, "threading",
                        SimpleNamespace(Thread=_InlineThread, Lock=threading.Lock))
    app.start_backup_thread()
    messages = []
    while not app.message_queue.empty():
        messages.append(app.message_queue.get_nowait())
    assert dialogs == []
    assert [m for m in messages if m[0] == "show_error"] == []


def _run_headless(mod, dest, monkeypatch, gui_app):
    mod.core_run_backup(user_tree.USER_NAME, str(dest))


RUNNERS = {"gui": _run_gui, "headless": _run_headless}


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


@pytest.mark.parametrize("variant", ["pywin32", "nopywin32"])
@pytest.mark.parametrize("path", [pytest.param("gui", marks=pytest.mark.gui), "headless"])
def test_backup_golden(path, variant, ezbak_mod, ezbak_mod_nopywin32, tmp_path, monkeypatch,
                       isolated_env, user_root, fixed_clock, block_subprocess, request):
    # Both module fixtures are session-scoped, so they are imported before fixed_clock
    # patches datetime (an import under the patch would keep FixedDatetime for good).
    mod = {"pywin32": ezbak_mod, "nopywin32": ezbak_mod_nopywin32}[variant]
    if variant == "pywin32" and not mod._have_pywin32:
        pytest.skip("pywin32 not installed")
    user_root.inject(mod)
    fixed_clock.inject(mod)
    home = user_tree.build_user_tree(str(user_root.path / user_tree.USER_NAME))
    monkeypatch.setenv("LOCALAPPDATA", os.path.join(home, "AppData", "Local"))
    dest = tmp_path / "dest"
    dest.mkdir()

    # gui_app only for the GUI path, so Headless runs never touch the shared Tk root.
    gui_app = request.getfixturevalue("gui_app") if path == "gui" else None
    RUNNERS[path](mod, dest, monkeypatch, gui_app)
    result = _snapshot(dest, tmp_path)

    assert block_subprocess == []
    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden.setdefault(path, {})[variant] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[path][variant]
