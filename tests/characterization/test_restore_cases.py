"""Characterization (golden) of restore special cases (P3-0a), on the untouched ezBAK.py.

Path: App.start_restore_thread -> App.run_restore, the same harness as
test_restore_golden.py (real App from gui_app, dialogs and the folder picker patched,
worker thread run inline). Each case lays out hand-made backup folders under
<TMP>/dest and an empty target profile <USERS_ROOT>/tester first.

The worker waits for the user by polling App attributes (_folder_selection_result,
_cross_user_confirmed; P3 decision D2 keeps the polling). With the worker inline, the
queue is not processed while it runs, so the answer is put on the App before the start;
time.sleep is replaced by a stub that aborts the test if a wait would ever loop.

Recorded per case (temp root -> <TMP>, real-clock times -> <TS>, zero log age -> <AGE>,
localized [WinError n] text -> <MSG>): the dialogs called (name, title, message), the
GUI queue (runs of consecutive update_progress folded into the last value; the progress
throttle is raised), the restore target tree, the restore job logs under dest, the
files under %LOCALAPPDATA%, and answer attributes left unconsumed on the App.
Today's behavior is recorded as-is, odd or not (disk_usage_fail: shutil.disk_usage fails
and the restore falls back to winsys.sysinfo.free_space, whose GetDiskFreeSpaceExW step
is stubbed to a fixed large value so the case does not depend on the real disk; fix-c).
A restore stopped before copying (cross-user 'no', low space) removes its job log, so
cross_direct_no, cross_parent_no and low_space record no logs (fix-e).
Results do not depend on pywin32 here (no hidden files), so only the pywin32 copy runs.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_restore_cases.py
"""
import json
import os
import re
import sys
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = [pytest.mark.windows, pytest.mark.gui]

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "restore_cases_golden.json"

USER = "tester"
OTHER = "other"
DIALOGS = ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel")
ANSWER_ATTRS = ("_folder_selection_result", "_cross_user_confirmed")
REAL_TS_RE = re.compile(r"(Modified|Cutoff time): \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")
ZERO_AGE_RE = re.compile(r"Age: -?0\.0 days")
WINERROR_RE = re.compile(r"(\[WinError \d+\])[^:\[]*:")

SMALL_TREE = {"Documents/report.txt": "quarterly report\n", "Desktop/todo.txt": "todo\n"}
OWN_NEW = f"{USER}_backup_2023-12-01_000000"
OWN_OLD = f"{USER}_backup_2023-06-01_000000"
OTHER_BK = f"{OTHER}_backup_2023-09-01_000000"


class _PollLoop(BaseException):
    """Raised by the time.sleep stub: a wait loop found no answer (escapes run_restore)."""


class _InlineThread:
    """Stand-in for threading.Thread that runs the target on start(), in this thread."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target, self._args, self._kwargs = target, args, kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


def _normalize(value, tmp_path):
    if isinstance(value, (list, tuple)):
        return [_normalize(v, tmp_path) for v in value]
    if not isinstance(value, str):
        return value
    root = str(tmp_path)
    text = value.replace(root, "<TMP>").replace(root.replace("\\", "\\\\"), "<TMP>")
    text = WINERROR_RE.sub(r"\1 <MSG>:", REAL_TS_RE.sub(r"\1: <TS>", text))
    return ZERO_AGE_RE.sub("Age: <AGE> days", text)


def _write(path, content="x\n"):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


def _backup(dest, name, tree=None):
    folder = dest / name
    for rel, content in (tree or SMALL_TREE).items():
        _write(str(folder / rel), content)
    return folder


# --- case setups: (dest, home, local) -> options ---
# options: pick (folder picker result), user, overwrite (askyesno answer),
# answers ({attr: value} put on the App before the start), free (disk_usage free bytes),
# api_free (free bytes from sysinfo's GetDiskFreeSpaceExW step).

def _no_user(dest, home, local):
    _backup(dest, OWN_NEW)
    return {"pick": dest, "user": ""}


def _overwrite_no(dest, home, local):
    _backup(dest, OWN_NEW)
    return {"pick": dest, "overwrite": False}


def _picker_cancel(dest, home, local):
    _backup(dest, OWN_NEW)
    return {"pick": ""}


def _log_file(dest, home, local):
    _write(str(dest / f"{OWN_NEW}.log"))
    return {"pick": dest / f"{OWN_NEW}.log"}


def _plain_file(dest, home, local):
    _write(str(dest / "notes.txt"))
    return {"pick": dest / "notes.txt"}


def _missing_path(dest, home, local):
    return {"pick": dest / "nowhere"}


def _backup_without_user_folders(dest, home, local):
    _backup(dest, OWN_NEW, {"stray.txt": "stray\n"})
    return {"pick": dest / OWN_NEW}


def _parent_without_backups(dest, home, local):
    _write(str(dest / "misc" / "a.txt"))
    return {"pick": dest}


def _multi_select(dest, home, local):
    _backup(dest, OWN_NEW)
    _backup(dest, OWN_OLD, {"Documents/old.txt": "old\n"})
    _backup(dest, OTHER_BK)
    return {"pick": dest, "answers": {"_folder_selection_result": OWN_OLD}}


def _multi_cancel(dest, home, local):
    _backup(dest, OWN_NEW)
    _backup(dest, OTHER_BK)
    return {"pick": dest, "answers": {"_folder_selection_result": None}}


def _cross(direct, confirmed):
    def setup(dest, home, local):
        _backup(dest, OTHER_BK, {"Documents/theirs.txt": "theirs\n"})
        return {"pick": dest / OTHER_BK if direct else dest,
                "answers": {"_cross_user_confirmed": confirmed}}
    return setup


def _low_space(dest, home, local):
    _backup(dest, OWN_NEW)
    return {"pick": dest, "free": 0}


def _disk_usage_fail(dest, home, local):
    # The target profile does not exist: shutil.disk_usage raises, the fallback runs.
    _backup(dest, OWN_NEW)
    os.rmdir(home)
    return {"pick": dest, "api_free": 10**12}


def _no_bookmarks(dest, home, local):
    _backup(dest, OWN_NEW)
    return {"pick": dest}


def _many_files(dest, home, local):
    _backup(dest, OWN_NEW, {f"Documents/many/f{i:02d}.txt": f"{i}\n" for i in range(55)})
    return {"pick": dest / OWN_NEW}


def _success(dest, home, local):
    _backup(dest, OWN_NEW)
    _write(str(dest / "Chrome_Bookmarks"), '{"chrome": 1}\n')
    _write(str(dest / "Edge_Bookmarks"), '{"edge": 1}\n')
    # Browser folders in both places: bookmarks go under the target profile (fix-d),
    # the running account's %LOCALAPPDATA% stays untouched.
    for browser in (("Google", "Chrome"), ("Microsoft", "Edge")):
        local.joinpath(*browser, "User Data", "Default").mkdir(parents=True)
        home.joinpath("AppData", "Local", *browser, "User Data", "Default").mkdir(parents=True)
    return {"pick": dest}


SETUPS = {
    "no_user": _no_user,
    "overwrite_no": _overwrite_no,
    "picker_cancel": _picker_cancel,
    "log_file": _log_file,
    "plain_file": _plain_file,
    "missing_path": _missing_path,
    "backup_without_user_folders": _backup_without_user_folders,
    "parent_without_backups": _parent_without_backups,
    "multi_select": _multi_select,
    "multi_cancel": _multi_cancel,
    "cross_direct_yes": _cross(True, True),
    "cross_direct_no": _cross(True, False),
    "cross_parent_yes": _cross(False, True),
    "cross_parent_no": _cross(False, False),
    "low_space": _low_space,
    "disk_usage_fail": _disk_usage_fail,
    "no_bookmarks": _no_bookmarks,
    "many_files": _many_files,
    "success": _success,
}


# --- recording ---

def _fold_queue(messages, tmp_path):
    folded = []
    for task, value in messages:
        value = _normalize(value, tmp_path)
        if task == "update_progress" and folded and folded[-1][0] == "update_progress":
            folded[-1] = [task, value]
        else:
            folded.append([task, value])
    return folded


def _walk(root, tmp_path):
    entries = []
    for dirpath, dirnames, filenames in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        prefix = "" if rel == "." else rel.replace(os.sep, "/") + "/"
        entries += [prefix + d + "/" for d in dirnames]
        entries += [prefix + f for f in filenames]
    return sorted(_normalize(e, tmp_path) for e in entries)


def _restore_logs(dest, tmp_path):
    logs = {}
    for dirpath, _dirnames, filenames in os.walk(dest):
        for f in filenames:
            if f.endswith(".log") and "_restore_" in f:
                path = Path(dirpath) / f
                rel = path.relative_to(dest).as_posix()
                logs[rel] = _normalize(path.read_text(encoding="utf-8"), tmp_path).splitlines()
    return logs


def _run(mod, opts, monkeypatch, gui_app):
    dialogs = []

    def answer(name):
        def dialog(*args, **kwargs):
            dialogs.append([name] + list(args[:2]))
            return opts.get("overwrite", True) if name == "askyesno" else None
        return dialog

    def no_wait(_secs):
        raise _PollLoop("run_restore waited for an answer that was not given")

    for name in DIALOGS:
        monkeypatch.setattr(mod.Win11Dialog, name, answer(name), raising=False)
        monkeypatch.setattr(mod.messagebox, name, answer(name), raising=False)
    monkeypatch.setattr(mod.filedialog, "askdirectory", lambda **kwargs: str(opts["pick"]))
    monkeypatch.setattr(mod, "threading",
                        SimpleNamespace(Thread=_InlineThread, Lock=threading.Lock))
    monkeypatch.setattr(mod.time, "sleep", no_wait)
    if "free" in opts:
        monkeypatch.setattr(mod.shutil, "disk_usage",
                            lambda path: SimpleNamespace(free=opts["free"]))
    if "api_free" in opts:
        sysinfo = sys.modules[mod.winsys_free_space.__module__]
        monkeypatch.setattr(sysinfo, "_api_free_space", lambda path: opts["api_free"])

    app = gui_app(mod)
    app.user_var.set(opts.get("user", USER))
    app.progress_throttle_secs = 1e9  # restored by gui_app's reset
    try:
        for attr, value in opts.get("answers", {}).items():
            setattr(app, attr, value)
        app.start_restore_thread()
        left = sorted(a for a in ANSWER_ATTRS if hasattr(app, a))
    finally:
        for attr in ANSWER_ATTRS:  # underscore names survive gui_app's reset
            if hasattr(app, attr):
                delattr(app, attr)
    messages = []
    while not app.message_queue.empty():
        messages.append(app.message_queue.get_nowait())
    return dialogs, messages, left


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


@pytest.mark.parametrize("case", list(SETUPS))
def test_restore_case(case, ezbak_mod, tmp_path, monkeypatch, isolated_env, user_root,
                      fixed_clock, block_subprocess, gui_app):
    mod = ezbak_mod
    if not mod._have_pywin32:
        pytest.skip("pywin32 not installed")
    user_root.inject(mod)
    fixed_clock.inject(mod)
    dest = tmp_path / "dest"
    dest.mkdir()
    home = user_root.path / USER
    home.mkdir(parents=True)
    local = isolated_env["LOCALAPPDATA"]
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    opts = SETUPS[case](dest, home, local)

    dialogs, messages, left = _run(mod, opts, monkeypatch, gui_app)
    result = {
        "dialogs": _normalize(dialogs, tmp_path),
        "queue": _fold_queue(messages, tmp_path),
        "target": _walk(home, tmp_path) if home.exists() else None,
        "logs": _restore_logs(dest, tmp_path),
        "localappdata": _walk(local, tmp_path),
        "answers_left": left,
    }

    assert block_subprocess == []
    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden[case] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[case]
