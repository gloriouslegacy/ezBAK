"""Characterization (golden) of every external command ezBAK.py starts.

Call sites in ezBAK.py (line numbers at T0-3; 30 sites, all covered):

| #  | line | function                         | command                                    | scenario(s)                     |
|----|------|----------------------------------|--------------------------------------------|---------------------------------|
| 1  | 1125 | get_system_info                  | wmic baseboard get product /value          | sysinfo_fallback_chain, driver_* |
| 2  | 1141 | get_system_info                  | wmic computersystem get model /value       | sysinfo_fallback_chain          |
| 3  | 1157 | get_system_info                  | wmic computersystem get manufacturer       | sysinfo_fallback_chain          |
| 4  | -    | show_notification (del. P3-10)   | - (D6: no caller, golden dropped)          | -                               |
| 5  | 2236 | App.open_github_link             | webbrowser.open_new(<releases url>)        | github_link                     |
| 6  | 2511 | App.process_queue 'open_folder'  | os.startfile(<folder>)                     | queue_open_folder               |
| 7  | 4337 | App.create_scheduled_task        | schtasks /create /tn /xml <file> /f (+XML) | schedule_create_*               |
| 8  | 4547 | App.delete_scheduled_task        | schtasks /Delete /TN <name> /F             | schedule_delete                 |
| 9  | 4888 | App.open_nas_connect_dialog      | net use [drv] unc [/user:u pw] /persistent | nas_connect_*                   |
| 10 | 4890 |   same, retry without encoding   | net use ...                                | nas_connect_retry_path          |
| 11 | 4909 |   same, open share               | os.startfile(drive or unc)                 | nas_connect_drive_user          |
| 12 | -    |   except-branch re-prompt (g-k)  | - (deleted; except shows an error only)    | nas_connect_retry_path          |
| 13 | -    |   except-branch retry (g-k)      | - (deleted)                                | nas_connect_retry_path          |
| 14 | -    |   except-branch open share (g-k) | - (deleted)                                | nas_connect_retry_path          |
| 15 | 5086 | App.open_nas_disconnect_dialog   | net use <target> /delete /y                | nas_disconnect_*                |
| 16 | 5088 |   same, retry without encoding   | net use <target> /delete /y                | nas_disconnect_retry_path       |
| 17 | -    |   except-branch (g-k)            | - (deleted; except shows an error only)    | nas_disconnect_retry_path       |
| 18 | -    |   except-branch retry (g-k)      | - (deleted)                                | nas_disconnect_retry_path       |
| 19 | 5180 | App.run_winget_export            | winget --version                           | winget_*                        |
| 20 | 5191 |   same (Popen)                   | winget export --output ... --include-...   | winget_*                        |
| 21 | 5218 |   fallback 1                     | winget export --output <path>              | winget_fallbacks                |
| 22 | 5244 |   fallback 2                     | winget list --accept-source-agreements     | winget_fallbacks                |
| 23 | 5270 |   fallback 3                     | winget list                                | winget_fallbacks                |
| 24 | 5642 | App.run_driver_backup (Popen)    | pnputil.exe /export-driver * "<dir>"       | driver_backup                   |
| 25 | 5817 | App.run_driver_restore (Popen)   | pnputil.exe /add-driver "<dir>\\*.inf" ... | driver_restore                  |
| 26 | 5870 | App.open_device_manager          | os.startfile(%WINDIR%\\system32\\devmgmt.msc) | devmgr_startfile             |
| 27 | 5876 |   fallback (Popen)               | mmc devmgmt.msc                            | devmgr_mmc                      |
| 28 | 5884 |   fallback (Popen, shell)        | start devmgmt.msc                          | devmgr_shell_start              |
| 29 | 5905 | App.open_file_explorer           | os.startfile(<home>)                       | explorer                        |
| 30 | 8197 | __main__ (not frozen, not admin) | ShellExecuteW(None, "runas", py, argv, ..) | relaunch_runas                  |

There is no "schtasks /query" call: schedules are listed from the JSON config files.

Every call is answered by record_subprocess (nothing is started). Recorded per call: the API,
positional args (argv) and keyword args (shell, encoding, errors, creationflags, ...).
Normalized: temp root -> <TMP>, sys.executable -> <PY>, repo root -> <REPO>; the
locale's preferred encoding is fixed to "<PREFERRED>" while a scenario runs; time is fixed
by fixed_clock. NAS credentials are the dummy value "dummy-pw".

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_commands_argv.py
"""
import json
import locale
import os
import runpy
import sys
import threading
import winreg
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.windows

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
EZBAK_PATH = REPO_ROOT / "ezBAK.py"
GOLDEN_PATH = Path(__file__).resolve().parent.parent / "fixtures" / "argv_golden.json"
NAS_UNC = "\\\\nas\\share"
NAS_USER = "nasuser"
NAS_PW = "dummy-pw"


class _InlineThread:
    """Stand-in for threading.Thread that runs the target on start(), in this thread."""

    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target, self._args, self._kwargs = target, args, kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


class _ExitCalled(BaseException):
    """Raised by the patched os._exit (the __main__ block calls it after sys.exit)."""


def _normalize(obj, subs):
    if isinstance(obj, str):
        for old, new in subs:
            obj = obj.replace(old, new)
        return obj
    if isinstance(obj, (list, tuple)):
        return [_normalize(v, subs) for v in obj]
    if isinstance(obj, dict):
        return {k: _normalize(v, subs) for k, v in obj.items()}
    return obj


def _block_registry(monkeypatch):
    def blocked(*args, **kwargs):
        raise OSError("registry blocked in tests")
    monkeypatch.setattr(winreg, "OpenKey", blocked)


# ---------------------------------------------------------------- scenarios without App

def _sysinfo_fallback_chain(ctx):
    _block_registry(ctx.mp)
    ctx.rec.respond("baseboard", stdout="Product=\r\n")
    ctx.rec.respond("model", stdout="Model=System Product Name\r\n")
    ctx.rec.respond("manufacturer", stdout="Manufacturer=ACME Corp\r\n")
    return {"result": ctx.mod.get_system_info()}


def _relaunch_runas(ctx):
    ctx.rec.fake_shell32(is_admin=0, shell_execute_ret=42)

    def fake_exit(code):
        raise _ExitCalled(code)

    ctx.mp.setattr(os, "_exit", fake_exit)
    ctx.mp.setattr(sys, "argv", [str(EZBAK_PATH), "--lang", "a b"])
    with pytest.raises(_ExitCalled):
        runpy.run_path(str(EZBAK_PATH), run_name="__main__")


# ---------------------------------------------------------------- scenarios with App

def _github_link(ctx):
    ctx.app.open_github_link(None)


def _queue_open_folder(ctx):
    _drain(ctx.app)
    ctx.app.message_queue.put(("open_folder", str(ctx.tmp / "out")))
    ctx.app.process_queue()


def _schedule_create(schedule, monthly_day=1):
    def scenario(ctx):
        dest = ctx.tmp / "dest"
        msg = ctx.app.create_scheduled_task("ezBAK_Test", schedule, "9:05", "tester", str(dest),
                                            monthly_day=monthly_day)
        argv = next(c["args"][0] for c in ctx.rec.calls if c["args"][0][0] == "schtasks")
        xml_file = argv[argv.index("/xml") + 1]
        with open(xml_file, encoding="utf-16") as f:
            xml_lines = f.read().splitlines()
        return {"message": msg.splitlines(), "task_xml": xml_lines}
    return scenario


def _schedule_delete(ctx):
    ctx.app.delete_scheduled_task("ezBAK_Test")


def _fill_dialog(ctx, values):
    """Patch app.wait_window: type `values` into the dialog's Entry widgets, press OK."""
    buttons = []
    original = ctx.mod.Win11Dialog._create_dialog_button

    def capture(parent, text, command, theme, is_primary=False):
        buttons.append(command)
        return original(parent, text, command, theme, is_primary=is_primary)

    ctx.mp.setattr(ctx.mod.Win11Dialog, "_create_dialog_button", staticmethod(capture))

    def entries(widget):
        for child in widget.winfo_children():
            if isinstance(child, ctx.mod.tk.Entry):
                yield child
            yield from entries(child)

    def wait_window(dlg):
        for entry, value in zip(list(entries(dlg)), values):
            if value is not None:
                entry.delete(0, "end")
                entry.insert(0, value)
        buttons[0]()  # on_ok

    ctx.mp.setattr(ctx.app, "wait_window", wait_window, raising=False)


def _nas_connect_drive_user(ctx):
    _fill_dialog(ctx, [NAS_UNC, "z", NAS_USER, NAS_PW])
    ctx.yes += [True]  # open share now
    ctx.app.open_nas_connect_dialog()


def _nas_connect_retry_path(ctx):
    # Every run raises: covers the no-encoding retry; the except branch only reports (g-k).
    ctx.rec.respond("net use", raises=OSError("net use failed"), times=3)
    _fill_dialog(ctx, [NAS_UNC, "x", "", ""])
    ctx.app.open_nas_connect_dialog()


def _nas_disconnect_unc(ctx):
    _fill_dialog(ctx, [NAS_UNC])
    ctx.app.open_nas_disconnect_dialog()


def _nas_disconnect_retry_path(ctx):
    ctx.rec.respond("net use", raises=OSError("net use failed"), times=3)
    _fill_dialog(ctx, ["z"])
    ctx.app.open_nas_disconnect_dialog()


def _winget(fail):
    def scenario(ctx):
        ctx.rec.respond("--version", stdout="v1.9.0\n")
        if fail:
            ctx.rec.respond("--include-versions", returncode=1, stderr="export error\n")
            ctx.rec.respond("winget export", returncode=1)
            ctx.rec.respond("winget list --accept", returncode=1)
            ctx.rec.respond("winget list", stdout="Name Id\n")
        else:
            ctx.rec.respond("--include-versions", stdout="exported\n")
        out = ctx.tmp / "apps"
        out.mkdir()
        ctx.folder = str(out)
        ctx.app.start_winget_export_thread()
        ctx.app.close_log_file()
    return scenario


def _driver_backup(ctx):
    _block_registry(ctx.mp)
    ctx.rec.respond("baseboard", stdout="Product=TESTBOARD\r\n")
    ctx.rec.respond("pnputil", stdout="Exported 1 driver\n")
    out = ctx.tmp / "drivers"
    out.mkdir()
    ctx.folder = str(out)
    ctx.app.start_driver_backup_thread()


def _driver_restore(ctx):
    _block_registry(ctx.mp)
    ctx.rec.respond("baseboard", stdout="Product=TESTBOARD\r\n")
    ctx.rec.respond("pnputil", stdout="Added 1 driver\n")
    folder = ctx.tmp / "drivers" / "TESTBOARD_drivers_backup_2024-01-02_030405"
    folder.mkdir(parents=True)
    ctx.folder = str(folder)
    ctx.app.start_driver_restore_thread()


def _devmgr(variant):
    def scenario(ctx):
        windir = ctx.tmp / "Windows"
        (windir / "system32").mkdir(parents=True)
        if variant == "startfile":
            (windir / "system32" / "devmgmt.msc").write_text("", encoding="utf-8")
        if variant == "shell_start":
            ctx.rec.respond("mmc", raises=OSError("mmc missing"))
        ctx.mp.setenv("WINDIR", str(windir))
        ctx.app.open_device_manager()
    return scenario


def _explorer(ctx):
    ctx.app.open_file_explorer()


PLAIN = {
    "sysinfo_fallback_chain": _sysinfo_fallback_chain,
    "relaunch_runas": _relaunch_runas,
}

GUI = {
    "github_link": _github_link,
    "queue_open_folder": _queue_open_folder,
    "schedule_create_daily": _schedule_create("Daily"),
    "schedule_create_weekly": _schedule_create("Weekly"),
    "schedule_create_monthly": _schedule_create("Monthly", monthly_day=15),
    "schedule_delete": _schedule_delete,
    "nas_connect_drive_user": _nas_connect_drive_user,
    "nas_connect_retry_path": _nas_connect_retry_path,
    "nas_disconnect_unc": _nas_disconnect_unc,
    "nas_disconnect_retry_path": _nas_disconnect_retry_path,
    "winget_export_ok": _winget(fail=False),
    "winget_fallbacks": _winget(fail=True),
    "driver_backup": _driver_backup,
    "driver_restore": _driver_restore,
    "devmgr_startfile": _devmgr("startfile"),
    "devmgr_mmc": _devmgr("mmc"),
    "devmgr_shell_start": _devmgr("shell_start"),
    "explorer": _explorer,
}


# ---------------------------------------------------------------- harness

def _drain(app):
    while not app.message_queue.empty():
        app.message_queue.get_nowait()


def _patch_dialogs(mod, ctx, monkeypatch):
    def recorder(name):
        def dialog(*args, **kwargs):
            ctx.dialogs.append(name)
            if name in ("askyesno", "askokcancel"):
                return ctx.yes.pop(0) if ctx.yes else False
            return None
        return dialog

    for name in ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel"):
        monkeypatch.setattr(mod.Win11Dialog, name, recorder(name), raising=False)
        monkeypatch.setattr(mod.messagebox, name, recorder(name), raising=False)

    def askstring(*args, **kwargs):
        ctx.dialogs.append("askstring")
        return ctx.strings.pop(0) if ctx.strings else None

    monkeypatch.setattr(mod.simpledialog, "askstring", askstring)
    monkeypatch.setattr(mod.filedialog, "askdirectory", lambda **kwargs: ctx.folder)


def _load_golden():
    if GOLDEN_PATH.exists():
        return json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
    return {}


def _check(name, ctx, extra):
    subs = [(str(ctx.tmp), "<TMP>"), (sys.executable, "<PY>"), (str(REPO_ROOT), "<REPO>")]
    result = _normalize({"calls": ctx.rec.calls, "dialogs": ctx.dialogs, **(extra or {})}, subs)
    if os.environ.get("EZBAK_REGEN_GOLDEN") == "1":
        golden = _load_golden()
        golden[name] = result
        GOLDEN_PATH.write_text(json.dumps(golden, indent=2, ensure_ascii=False, sort_keys=True)
                               + "\n", encoding="utf-8")
    assert result == _load_golden()[name]


@pytest.fixture
def ctx(ezbak_mod, tmp_path, monkeypatch, isolated_env, fixed_clock, record_subprocess):
    # ezbak_mod is requested before fixed_clock so the session import is not made under the patch.
    fixed_clock.inject(ezbak_mod)
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.setenv("USERNAME", "tester")
    monkeypatch.setenv("USERDOMAIN", "TESTPC")
    return SimpleNamespace(mod=ezbak_mod, app=None, rec=record_subprocess, mp=monkeypatch,
                           tmp=tmp_path, dialogs=[], yes=[], strings=[], folder=None)


@pytest.mark.parametrize("name", sorted(PLAIN))
def test_commands_argv_plain(name, ctx):
    ctx.mp.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "<PREFERRED>")
    extra = PLAIN[name](ctx)
    _check(name, ctx, extra)


@pytest.fixture(scope="module")
def shared_app(gui_app_session):
    """The session App (see conftest gui_app), reset once; the GUI scenarios of this
    module then share its state in order, as with a fresh App made for the module."""
    gui_app_session.reset()
    yield gui_app_session.app
    gui_app_session.reset()


@pytest.mark.gui
@pytest.mark.parametrize("name", sorted(GUI))
def test_commands_argv_gui(name, shared_app, ctx):
    mod, app = ctx.mod, shared_app
    _patch_dialogs(mod, ctx, ctx.mp)
    ctx.app = app
    ctx.mp.setattr(mod, "threading", SimpleNamespace(Thread=_InlineThread, Lock=threading.Lock))
    ctx.mp.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "<PREFERRED>")
    _drain(app)
    ctx.rec.calls.clear()
    try:
        extra = GUI[name](ctx)
    finally:
        app.close_log_file()
        _drain(app)
    _check(name, ctx, extra)
