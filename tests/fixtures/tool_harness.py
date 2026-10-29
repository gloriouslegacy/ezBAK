"""Shared setup for the tool case goldens (tests/characterization/test_tool_cases*.py, P4-0).

Helpers copied from test_commands_argv.py / test_schedule_cases.py (those files stay as
they are): InlineThread, fill_dialog, fold_traceback. New here: a fake winreg (patched
onto the real module, so the registry is never opened), dialog recorder with answers,
queue drain, folder walk, normalizer and the golden check.

Loaded with importlib (tests/fixtures is not a package), like cli_harness.py.
"""
import json
import os
import re
import sys
import threading
import time
import winreg
from types import SimpleNamespace

DIALOGS = ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel")
LOG_TS_RE = re.compile(r"\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]")


class InlineThread:
    """Stand-in for threading.Thread that runs the target on start(), in this thread."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target, self._args, self._kwargs = target, args, kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


def inline_threads(mod, monkeypatch):
    monkeypatch.setattr(mod, "threading", SimpleNamespace(Thread=InlineThread, Lock=threading.Lock))


def fold_traceback(text):
    """Keep the 'Traceback' header, chain lines and exception lines; frames become one line."""
    out, in_frames = [], False
    for line in text.splitlines():
        if line.endswith("Traceback (most recent call last):"):
            out.extend([line, "  <FRAMES>"])
            in_frames = True
        elif in_frames and line.startswith(" "):
            continue
        else:
            in_frames = False
            out.append(line)
    return out


def normalizer(tmp_path, repo_root):
    """value -> value with temp root <TMP>, repo <REPO>, python <PY>, log stamps <TS>."""
    subs = [(str(tmp_path), "<TMP>"), (str(repo_root), "<REPO>"), (sys.executable, "<PY>")]

    def norm(value):
        if isinstance(value, str):
            for old, new in subs:
                value = value.replace(old, new)
            return LOG_TS_RE.sub("[<TS>]", value)
        if isinstance(value, (list, tuple)):
            return [norm(v) for v in value]
        if isinstance(value, dict):
            return {k: norm(v) for k, v in value.items()}
        return value

    return norm


class FakeWinreg:
    """Replaces winreg.OpenKey / QueryValueEx / CloseKey (attributes of the real module).

    open_error: exception OpenKey raises (None: returns a fake handle).
    values: {value name: value}; a missing name raises FileNotFoundError like winreg.
    """

    def __init__(self, monkeypatch, open_error=None, values=None):
        self.calls = []
        self._open_error = open_error
        self._values = values or {}
        monkeypatch.setattr(winreg, "OpenKey", self._open)
        monkeypatch.setattr(winreg, "QueryValueEx", self._query)
        monkeypatch.setattr(winreg, "CloseKey", self._close)

    def _open(self, key, sub_key, reserved=0, access=winreg.KEY_READ):
        self.calls.append(["OpenKey", "HKLM" if key == winreg.HKEY_LOCAL_MACHINE else key,
                           sub_key, reserved, "KEY_READ" if access == winreg.KEY_READ else access])
        if self._open_error is not None:
            raise self._open_error
        return "<handle>"

    def _query(self, handle, name):
        self.calls.append(["QueryValueEx", name])
        if name not in self._values:
            raise FileNotFoundError(2, "The system cannot find the file specified")
        return self._values[name], winreg.REG_SZ

    def _close(self, handle):
        self.calls.append(["CloseKey"])


def blocked_registry(monkeypatch):
    return FakeWinreg(monkeypatch, open_error=OSError("registry blocked in tests"))


def registry_value(monkeypatch, product):
    return FakeWinreg(monkeypatch, values={"BaseBoardProduct": product})


def patch_dialogs(mod, monkeypatch, dialogs, yes=None, folder=None, strings=None):
    """Record [name, title, message] of every dialog; askyesno/askokcancel pop `yes`
    (default False); simpledialog.askstring pops `strings` (default None, the typed value
    is not recorded); filedialog.askdirectory returns `folder`. Questions a worker sends
    to the Tk thread ('ask_yes_no', g-e) are answered the same way (main_thread_turns)."""
    answers = list(yes or [])
    typed = list(strings or [])

    def askstring(*args, **kwargs):
        dialogs.append(["askstring"] + [str(a) for a in args[:2]])
        return typed.pop(0) if typed else None

    def recorder(name):
        def dialog(*args, **kwargs):
            dialogs.append([name] + [str(a) for a in args[:2]])
            if name in ("askyesno", "askokcancel"):
                return answers.pop(0) if answers else False
            return None
        return dialog

    for name in DIALOGS:
        monkeypatch.setattr(mod.Win11Dialog, name, recorder(name), raising=False)
        monkeypatch.setattr(mod.messagebox, name, recorder(name), raising=False)
    monkeypatch.setattr(mod.simpledialog, "askstring", askstring)
    monkeypatch.setattr(mod.filedialog, "askdirectory", lambda **kwargs: folder)
    main_thread_turns(mod, monkeypatch)


class NoAnswer(BaseException):
    """Raised by the poll stub: a worker waits for an answer no queued question gives."""


def main_thread_turns(mod, monkeypatch):
    """With inline workers nobody runs process_queue while App._poll_answer waits, so the
    bridge's time.sleep becomes one Tk-thread turn: the oldest unhandled 'ask_yes_no' in the
    queue goes through App._handle_queue_message (the recorded askyesno answers it). The
    message stays queued, so goldens still show it. No such message: NoAnswer (fail fast)."""
    handled = set()

    def turn(_secs):
        app = sys._getframe(1).f_locals["self"]  # the waiting App (caller: _poll_answer)
        for item in list(app.message_queue.queue):
            if item[0] == "ask_yes_no" and id(item) not in handled:
                handled.add(id(item))
                app._handle_queue_message(*item)
                return
        raise NoAnswer("worker is waiting but no 'ask_yes_no' message is queued")

    fake_time = SimpleNamespace(time=time.time, sleep=turn)
    monkeypatch.setitem(mod.App._poll_answer.__globals__, "time", fake_time)


def fill_dialog(mod, app, monkeypatch, values, press=0):
    """Patch app.wait_window: type `values` into the dialog's Entry widgets, press button
    `press` (0 OK, 1 Cancel). A dialog the button left open (failed check) is then closed,
    as the user would close it."""
    buttons = []
    original = mod.Win11Dialog._create_dialog_button

    def capture(parent, text, command, theme, is_primary=False):
        buttons.append(command)
        return original(parent, text, command, theme, is_primary=is_primary)

    monkeypatch.setattr(mod.Win11Dialog, "_create_dialog_button", staticmethod(capture))

    def entries(widget):
        for child in widget.winfo_children():
            if isinstance(child, mod.tk.Entry):
                yield child
            yield from entries(child)

    def wait_window(dlg):
        for entry, value in zip(list(entries(dlg)), values):
            if value is not None:
                entry.delete(0, "end")
                entry.insert(0, value)
        buttons[press]()  # 0 on_ok, 1 on_cancel
        if dlg.winfo_exists():
            dlg.destroy()

    monkeypatch.setattr(app, "wait_window", wait_window, raising=False)


def drain(app):
    messages = []
    while not app.message_queue.empty():
        task, value = app.message_queue.get_nowait()
        messages.append([task, value])
    return messages


def walk(root):
    """Sorted relative entries under root ('dir/' for folders)."""
    entries = []
    for dirpath, dirnames, filenames in os.walk(root):
        rel = os.path.relpath(dirpath, root)
        prefix = "" if rel == "." else rel.replace(os.sep, "/") + "/"
        entries += [prefix + d + "/" for d in dirnames]
        entries += [prefix + f for f in filenames]
    return sorted(entries)


def read_logs(root):
    """{relative path: lines} of every .log file under root (traceback frames folded)."""
    logs = {}
    for dirpath, _dirnames, filenames in os.walk(root):
        for name in filenames:
            if name.endswith(".log"):
                path = os.path.join(dirpath, name)
                rel = os.path.relpath(path, root).replace(os.sep, "/")
                with open(path, encoding="utf-8") as f:
                    logs[rel] = fold_traceback(f.read())
    return logs


def files_containing(root, secret):
    """Relative paths of files under root whose bytes contain `secret` (expected: none)."""
    found = []
    for dirpath, _dirnames, filenames in os.walk(root):
        for name in filenames:
            path = os.path.join(dirpath, name)
            with open(path, "rb") as f:
                if secret.encode("utf-8") in f.read():
                    found.append(os.path.relpath(path, root))
    return found


def check_golden(path, name, result):
    """Compare result with golden[name]; EZBAK_REGEN_GOLDEN=1 rewrites the entry first."""
    def load():
        return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}

    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = load()
        golden[name] = result
        path.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True) + "\n",
                        encoding="utf-8")
    assert result == load()[name]
