"""Characterization (golden) of backup special cases (P2-0), on the untouched ezBAK.py.

Same harness as test_backup_golden.py (gui: App.start_backup_thread -> App.run_backup
with the worker run inline; headless: core_run_backup), but each case prepares the
destination or the environment first:

- old_backups:     three older backup folders of the same user (one holds a read-only
                   file) plus another user's folder; retention 2
- old_logs:        old and recent .log files (mtime set with os.utime) and an old .txt
- bookmarks:       Chrome and Edge bookmark files in the backed-up user's profile
                   (<user_root>/<USER>/AppData/Local), plus decoys with other content
                   in the running account's %LOCALAPPDATA% (P2-fix #1)
- same_name:       the backup folder name (fixed by fixed_clock) already exists and
                   holds a read-only file
- user_missing:    the user folder does not exist
- low_space (gui): get_free_space_reliable returns 0
- queue (gui):     default settings on the full user tree; only the queue is recorded

Recorded per run (temp root -> <TMP>, real-clock times -> <TS>, zero log age -> <AGE>,
localized [WinError n] text -> <MSG>):
the destination tree (read-only files marked " [ro]"), the job log
(None when absent), the exception raised (Headless), and the GUI queue with runs of
consecutive update_progress folded into one (the last value). The progress throttle
is raised so only the first in-loop progress message is sent, whatever the speed.
Today's behavior is recorded as-is, odd or not. Q1 exception P2-fix #2: the Headless
same-name delete now removes read-only files too (it used a handler with the wrong
arguments, p2-plan section 5 #2), so its golden holds only the new backup's files.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_backup_cases.py
"""
import importlib.util
import json
import os
import re
import stat
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.windows

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "backup_cases_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_user_tree_cases", FIXTURES / "user_tree.py")
user_tree = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(user_tree)

USER = user_tree.USER_NAME
FIXED_NAME = f"{USER}_backup_2024-01-02_030405"  # backup folder / log name under fixed_clock
# Only times taken from the real clock are replaced; fixed_clock makes the others stable,
# so seeded folder names (dates) stay readable in the tree and the log.
REAL_TS_RE = re.compile(r"(Modified|Cutoff time): \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")
# Age of a log written during the run compares real time.time() with its mtime: 0.0 or -0.0.
ZERO_AGE_RE = re.compile(r"Age: -?0\.0 days")
# Windows error text is localized ("Access is denied" / ...): keep the code, drop the text.
WINERROR_RE = re.compile(r"(\[WinError \d+\])[^:\[]*:")
FILE_ATTRIBUTE_READONLY = 0x01
DAY = 86400

SMALL_TREE = {"Documents/report.txt": "quarterly report\n", "Desktop/todo.txt": "todo\n"}


class _InlineThread:
    """Stand-in for threading.Thread that runs the target on start(), in this thread."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target, self._args, self._kwargs = target, args, kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


def _normalize(text, tmp_path):
    root = str(tmp_path)
    text = text.replace(root, "<TMP>").replace(root.replace("\\", "\\\\"), "<TMP>")  # repr() form
    text = WINERROR_RE.sub(r"\1 <MSG>:", REAL_TS_RE.sub(r"\1: <TS>", text))
    return ZERO_AGE_RE.sub("Age: <AGE> days", text)


def _write(path, content="x\n", readonly=False, age_days=None):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)
    if age_days is not None:
        mtime = time.time() - age_days * DAY
        os.utime(path, (mtime, mtime))
    if readonly:
        os.chmod(path, stat.S_IREAD)


def _small_home(user_root):
    home = user_root.path / USER
    for rel, content in SMALL_TREE.items():
        _write(str(home / rel), content)
    return home


# --- case setups: (tmp_path, dest, user_root, isolated_env, monkeypatch) -> options ---

def _setup_old_backups(tmp_path, dest, user_root, env, mp):
    _small_home(user_root)
    _write(str(dest / f"{USER}_backup_2023-01-01_000000" / "Documents" / "old.txt"), readonly=True)
    _write(str(dest / f"{USER}_backup_2023-06-01_000000" / "a.txt"))
    _write(str(dest / f"{USER}_backup_2023-12-01_000000" / "b.txt"))
    _write(str(dest / "other_backup_2020-01-01_000000" / "c.txt"))
    return {}


def _setup_old_logs(tmp_path, dest, user_root, env, mp):
    _small_home(user_root)
    _write(str(dest / f"{USER}_backup_2023-01-01_000000.log"), age_days=60)
    _write(str(dest / f"{USER}_backup_2023-12-31_000000.log"), age_days=1)
    _write(str(dest / "notes.txt"), age_days=60)
    return {}


def _setup_bookmarks(tmp_path, dest, user_root, env, mp):
    home = _small_home(user_root)
    chrome = ("Google", "Chrome", "User Data", "Default", "Bookmarks")
    edge = ("Microsoft", "Edge", "User Data", "Default", "Bookmarks")
    local = os.path.join(str(home), "AppData", "Local")
    _write(os.path.join(local, *chrome), '{"chrome": 1}\n')
    _write(os.path.join(local, *edge), '{"edge": 1}\n')
    # Decoys in the running account's %LOCALAPPDATA%: must not be backed up.
    decoy = str(env["LOCALAPPDATA"])
    _write(os.path.join(decoy, *chrome), '{"decoy": "chrome"}\n')
    _write(os.path.join(decoy, *edge), '{"decoy": "edge"}\n')
    return {"bookmark_files": ("Chrome_Bookmarks", "Edge_Bookmarks")}


def _setup_same_name(tmp_path, dest, user_root, env, mp):
    _small_home(user_root)
    _write(str(dest / FIXED_NAME / "stale.txt"), "stale\n", readonly=True)
    _write(str(dest / FIXED_NAME / "Documents" / "report.txt"), "older report\n")
    return {}


def _setup_user_missing(tmp_path, dest, user_root, env, mp):
    return {"user": "ghost"}


def _setup_low_space(tmp_path, dest, user_root, env, mp):
    _small_home(user_root)
    return {"free_space": 0}


def _setup_queue(tmp_path, dest, user_root, env, mp):
    home = user_tree.build_user_tree(str(user_root.path / USER))
    mp.setenv("LOCALAPPDATA", os.path.join(home, "AppData", "Local"))
    return {"queue_only": True}


SETUPS = {
    "old_backups": _setup_old_backups,
    "old_logs": _setup_old_logs,
    "bookmarks": _setup_bookmarks,
    "same_name": _setup_same_name,
    "user_missing": _setup_user_missing,
    "low_space": _setup_low_space,
    "queue": _setup_queue,
}
GUI_ONLY = {"low_space", "queue"}


# --- recording ---

def _fold_queue(messages, tmp_path):
    folded = []
    for task, value in messages:
        if isinstance(value, str):
            value = _normalize(value, tmp_path)
        if task == "update_progress" and folded and folded[-1][0] == "update_progress":
            folded[-1] = [task, value]
        else:
            folded.append([task, value])
    return folded


def _snapshot(dest, tmp_path):
    entries = []
    for dirpath, dirnames, filenames in os.walk(dest):
        rel = os.path.relpath(dirpath, dest)
        prefix = "" if rel == "." else rel.replace(os.sep, "/") + "/"
        entries.extend(prefix + d + "/" for d in dirnames)
        for f in filenames:
            attrs = os.stat(os.path.join(dirpath, f)).st_file_attributes
            entries.append(prefix + f + (" [ro]" if attrs & FILE_ATTRIBUTE_READONLY else ""))
    return sorted(_normalize(e, tmp_path) for e in entries)


def _job_log(dest, user, tmp_path):
    path = dest / f"{user}_backup_2024-01-02_030405.log"
    if not path.exists():
        return None
    return _normalize(path.read_text(encoding="utf-8"), tmp_path).splitlines()


# --- runners ---

def _run_gui(mod, dest, monkeypatch, gui_app, opts):
    dialogs = []

    def record_dialog(*args, **kwargs):
        dialogs.append(list(args[:2]))

    for name in ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel"):
        monkeypatch.setattr(mod.Win11Dialog, name, record_dialog, raising=False)
        monkeypatch.setattr(mod.messagebox, name, record_dialog, raising=False)
    monkeypatch.setattr(mod.filedialog, "askdirectory", lambda **kwargs: str(dest))
    if "free_space" in opts:
        monkeypatch.setattr(mod.App, "get_free_space_reliable",
                            lambda self, path: opts["free_space"])

    app = gui_app(mod)
    app.user_var.set(opts.get("user", USER))
    app.progress_throttle_secs = 1e9  # restored by gui_app's reset
    if "app_patch" in opts:
        opts["app_patch"](app, monkeypatch)
    monkeypatch.setattr(mod, "threading",
                        SimpleNamespace(Thread=_InlineThread, Lock=threading.Lock))
    app.start_backup_thread()
    messages = []
    while not app.message_queue.empty():
        messages.append(app.message_queue.get_nowait())
    return {"dialogs": dialogs, "queue": messages}


def _run_headless(mod, dest, monkeypatch, gui_app, opts):
    try:
        mod.core_run_backup(opts.get("user", USER), str(dest))
    except Exception as e:  # recorded as part of today's behavior
        return {"exception": [type(e).__name__, str(e)]}
    return {"exception": None}


RUNNERS = {"gui": _run_gui, "headless": _run_headless}


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


def _params():
    for case in SETUPS:
        for path in ("gui", "headless"):
            if path == "headless" and case in GUI_ONLY:
                continue
            marks = [pytest.mark.gui] if path == "gui" else []
            yield pytest.param(case, path, marks=marks, id=f"{case}-{path}")


@pytest.mark.parametrize("variant", ["pywin32", "nopywin32"])
@pytest.mark.parametrize("case,path", list(_params()))
def test_backup_case(case, path, variant, ezbak_mod, ezbak_mod_nopywin32, tmp_path, monkeypatch,
                     isolated_env, user_root, fixed_clock, block_subprocess, request):
    mod = {"pywin32": ezbak_mod, "nopywin32": ezbak_mod_nopywin32}[variant]
    if variant == "pywin32" and not mod._have_pywin32:
        pytest.skip("pywin32 not installed")
    user_root.inject(mod)
    fixed_clock.inject(mod)
    dest = tmp_path / "dest"
    dest.mkdir()
    opts = SETUPS[case](tmp_path, dest, user_root, isolated_env, monkeypatch)

    gui_app = request.getfixturevalue("gui_app") if path == "gui" else None
    run = RUNNERS[path](mod, dest, monkeypatch, gui_app, opts)
    result = {}
    if "queue" in run:
        result["queue"] = _fold_queue(run.pop("queue"), tmp_path)
    if not opts.get("queue_only"):
        result.update(run)
        if result.get("exception"):
            result["exception"] = [_normalize(t, tmp_path) for t in result["exception"]]
        result["tree"] = _snapshot(dest, tmp_path)
        result["log"] = _job_log(dest, opts.get("user", USER), tmp_path)
        for name in opts.get("bookmark_files", ()):
            copied = dest / name
            result.setdefault("bookmarks", {})[name] = (
                copied.read_text(encoding="utf-8") if copied.exists() else None)

    assert block_subprocess == []
    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden.setdefault(case, {}).setdefault(path, {})[variant] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[case][path][variant]


# --- GUI retention settings that cannot be read (P2-7 review; not part of the golden) ---
# Expected text and outcome come from the code before P2-7 (28d98f1 App._cleanup_old_backups
# and the cleanup block of App.run_backup): a non-number falls back to 2 / 30 and cleans up;
# a var whose get() fails logs "ERROR: An error occurred during backup cleanup: ..." and
# skips the backup cleanup, then "Cleanup failed: ..." and no log cleanup.

class _BrokenVar:
    def get(self):
        raise RuntimeError("var unreadable")


def _setup_old_backups_and_logs(tmp_path, dest, user_root, env, mp):
    _setup_old_backups(tmp_path, dest, user_root, env, mp)
    _write(str(dest / f"{USER}_backup_2023-01-01_000000.log"), age_days=60)
    return {}


def _run_retention_case(mod, tmp_path, monkeypatch, user_root, fixed_clock, gui_app, patch):
    user_root.inject(mod)
    fixed_clock.inject(mod)
    dest = tmp_path / "dest"
    dest.mkdir()
    opts = _setup_old_backups_and_logs(tmp_path, dest, user_root, None, monkeypatch)
    opts["app_patch"] = patch
    run = _run_gui(mod, dest, monkeypatch, gui_app, opts)
    return run, _snapshot(dest, tmp_path), "\n".join(_job_log(dest, USER, tmp_path))


@pytest.mark.gui
def test_gui_unreadable_retention_vars_skip_cleanup(ezbak_mod, tmp_path, monkeypatch, isolated_env,
                                                    user_root, fixed_clock, block_subprocess, gui_app):
    # Own MonkeyPatch, undone before gui_app's reset reads the vars again.
    broken = pytest.MonkeyPatch()

    def patch(app, mp):
        broken.setattr(app, "retention_count_var", _BrokenVar())
        broken.setattr(app, "log_retention_days_var", _BrokenVar())

    try:
        run, tree, log = _run_retention_case(ezbak_mod, tmp_path, monkeypatch, user_root,
                                             fixed_clock, gui_app, patch)
    finally:
        broken.undo()
    assert f"{USER}_backup_2023-01-01_000000/Documents/old.txt [ro]" in tree
    assert f"{USER}_backup_2023-06-01_000000/a.txt" in tree
    assert f"{USER}_backup_2023-01-01_000000.log" in tree
    assert "BackupKeep=2, LogKeep=30" in log  # header fallback (P2-1)
    assert "ERROR: An error occurred during backup cleanup: var unreadable" in log
    assert "Cleanup failed: var unreadable" in log
    assert "[Old backup cleanup]" not in log and "[Old log cleanup]" not in log
    assert log.endswith("Backup successfully completed.")
    folder = os.path.join("<TMP>", "dest", FIXED_NAME)
    assert ["open_folder", folder] in _fold_queue(run["queue"], tmp_path)
    assert block_subprocess == []


@pytest.mark.gui
def test_gui_non_numeric_retention_falls_back(ezbak_mod, tmp_path, monkeypatch, isolated_env,
                                              user_root, fixed_clock, block_subprocess, gui_app):
    def patch(app, mp):
        app.retention_count_var.set("abc")  # restored by gui_app's reset
        app.log_retention_days_var.set("x")

    _run, tree, log = _run_retention_case(ezbak_mod, tmp_path, monkeypatch, user_root, fixed_clock,
                                          gui_app, patch)
    assert "BackupKeep=abc, LogKeep=x" in log
    assert "Found 4 backups. Keeping the 2 most recent." in log
    assert not any(e.startswith(f"{USER}_backup_2023-0") for e in tree)  # two oldest deleted
    assert f"{USER}_backup_2023-01-01_000000.log" not in tree  # 60-day-old log deleted (30)
    assert "ERROR: An error occurred during backup cleanup" not in log
    assert "Cleanup failed" not in log
    assert block_subprocess == []
