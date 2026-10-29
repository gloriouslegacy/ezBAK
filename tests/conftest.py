"""Shared fixtures for ezBAK tests (T0-1 scope).

- ezbak_mod: the untouched root ezBAK.py, seen through the test facade
  (tests/fixtures/ezbak_facade.py: root -> ezbak.gui public names -> alias table).
- ezbak_mod_nopywin32: a second copy imported while pywin32 is hidden,
  so its ctypes fallback branch is defined exactly as on a machine without pywin32.
- isolated_env: temporary %APPDATA% / %LOCALAPPDATA%.
- user_root: temp folder standing in for the C: Users folder (patched into ezBAK.USERS_ROOT).
- fixed_clock: datetime.now() returns FIXED_NOW (module-level and local imports).
- block_subprocess: every subprocess / os.system / os.startfile call is recorded
  and fails, so no external command can run.
- record_subprocess: every external call (subprocess, os.system/startfile, webbrowser,
  optionally ctypes shell32.ShellExecuteW) is recorded with its arguments and answered
  with a fake result chosen by the test; nothing is started.
- gui_app: the one withdrawn App (Tk root) of the session, reset for each test.
- run_ezbak_main: runs the untouched ezBAK.py as __main__ (runpy) with the backup
  engine stubbed, no UAC relaunch, no Tk; returns what it printed, its exit and calls.
"""
import copy
import ctypes
import datetime as _dt
import io
import importlib
import logging
import importlib.util
import os
import subprocess
import sys
import types
import webbrowser
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
EZBAK_PATH = REPO_ROOT / "ezBAK.py"

if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

_facade_spec = importlib.util.spec_from_file_location(
    "ezbak_test_facade", REPO_ROOT / "tests" / "fixtures" / "ezbak_facade.py")
ezbak_facade = importlib.util.module_from_spec(_facade_spec)
_facade_spec.loader.exec_module(ezbak_facade)


@pytest.fixture(autouse=True)
def no_diag_handler_left():
    """A test that runs an entry point (headless.run, gui.run) gets the diagnostic file
    handler of ezbak.core.diag.setup (P6); remove it and restore the `ezbak` logger
    level afterwards, so no file stays open and caplog sees DEBUG in later tests.
    Done on the logging objects, so it covers the nopywin32 copy of diag as well."""
    yield
    logger = logging.getLogger("ezbak")
    for handler in [h for h in logger.handlers if getattr(h, "_ezbak_diag", False)]:
        logger.removeHandler(handler)
        logger.setLevel(handler._ezbak_prev_level)
        handler.close()


@pytest.fixture(scope="session")
def ezbak_mod():
    """Root ezBAK.py (has a __main__ guard; import creates no Tk, UAC or schtasks),
    wrapped in the test facade (tests/fixtures/ezbak_facade.py, P5 D2)."""
    mod = importlib.import_module("ezBAK")
    assert Path(mod.__file__).resolve() == EZBAK_PATH
    return ezbak_facade.EzbakFacade(mod)


@pytest.fixture(scope="session")
def ezbak_mod_nopywin32():
    """Import ezBAK.py again under another name with win32api/win32con unavailable.

    The ezbak package is imported afresh for this copy too, so its
    ezbak.winsys.fileattrs (used by ezbak.core.filters) also has HAS_PYWIN32
    False. The normal ezbak modules are restored in sys.modules afterwards.
    """
    importlib.import_module("ezBAK")  # make sure the normal ezbak modules are loaded first
    with pytest.MonkeyPatch.context() as mp:
        mp.setitem(sys.modules, "win32api", None)
        mp.setitem(sys.modules, "win32con", None)
        normal = [n for n in sys.modules if n == "ezbak" or n.startswith("ezbak.")]
        for name in normal:
            mp.delitem(sys.modules, name)
        spec = importlib.util.spec_from_file_location("ezBAK_nopywin32", EZBAK_PATH)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        fresh = ezbak_facade.capture_ezbak_modules()  # before they leave sys.modules
        fresh_fileattrs = fresh["ezbak.winsys.fileattrs"]
        extra = [n for n in fresh if n not in normal]
    for name in extra:  # submodules only the fresh import loaded
        sys.modules.pop(name, None)
    facade = ezbak_facade.EzbakFacade(mod, fresh)
    assert facade._have_pywin32 is False
    assert fresh_fileattrs.HAS_PYWIN32 is False
    assert sys.modules["ezbak.winsys.fileattrs"] is not fresh_fileattrs
    return facade


@pytest.fixture
def isolated_env(tmp_path, monkeypatch):
    """Point %APPDATA% and %LOCALAPPDATA% at empty temp folders."""
    appdata = tmp_path / "env" / "AppData" / "Roaming"
    localappdata = tmp_path / "env" / "AppData" / "Local"
    appdata.mkdir(parents=True)
    localappdata.mkdir(parents=True)
    monkeypatch.setenv("APPDATA", str(appdata))
    monkeypatch.setenv("LOCALAPPDATA", str(localappdata))
    return {"APPDATA": appdata, "LOCALAPPDATA": localappdata}


@pytest.fixture
def user_root(tmp_path, monkeypatch):
    """Create <tmp>/Users and point ezBAK.USERS_ROOT at it via inject(mod)."""
    root = tmp_path / "Users"
    root.mkdir()

    def inject(mod):
        monkeypatch.setattr(mod, "USERS_ROOT", str(root))

    return types.SimpleNamespace(path=root, inject=inject)


FIXED_NOW = (2024, 1, 2, 3, 4, 5)


class FixedDatetime(_dt.datetime):
    """datetime whose now() is always FIXED_NOW (other methods unchanged)."""

    @classmethod
    def now(cls, tz=None):
        value = cls(*FIXED_NOW)
        return value if tz is None else value.replace(tzinfo=tz)


@pytest.fixture
def fixed_clock(monkeypatch):
    """Freeze datetime.now().

    Patches datetime.datetime (covers `from datetime import datetime` inside functions,
    e.g. core_run_backup) and, through inject(mod), the module's own `datetime` name.
    time.time() is left real (copy timeouts and log-age checks use it).
    """
    monkeypatch.setattr(_dt, "datetime", FixedDatetime)

    def inject(mod):
        monkeypatch.setattr(mod, "datetime", FixedDatetime)

    return types.SimpleNamespace(now=FixedDatetime.now(), inject=inject)


class BlockedSubprocess(RuntimeError):
    """Raised when code under test tries to start an external process."""


@pytest.fixture
def block_subprocess(monkeypatch):
    """Record and fail every attempt to run an external command; returns the call list."""
    calls = []

    def make(name):
        def blocked(*args, **kwargs):
            argv = args[0] if args else kwargs.get("args")
            calls.append((name, argv))
            raise BlockedSubprocess(f"{name} blocked in tests: {argv!r}")
        return blocked

    for name in ("run", "Popen", "call", "check_call", "check_output",
                 "getoutput", "getstatusoutput"):
        monkeypatch.setattr(subprocess, name, make(f"subprocess.{name}"))
    monkeypatch.setattr(os, "system", make("os.system"))
    monkeypatch.setattr(os, "startfile", make("os.startfile"), raising=False)
    return calls


_STREAM_NAMES = {subprocess.PIPE: "PIPE", subprocess.DEVNULL: "DEVNULL", subprocess.STDOUT: "STDOUT"}


def _plain_kwargs(kwargs):
    """kwargs as JSON-friendly values (PIPE/DEVNULL/STDOUT constants become their names)."""
    out = {}
    for key, value in sorted(kwargs.items()):
        if key in ("stdin", "stdout", "stderr") and value in _STREAM_NAMES:
            value = _STREAM_NAMES[value]
        out[key] = value
    return out


def _command_text(argv):
    return argv if isinstance(argv, str) else " ".join(str(a) for a in (argv or ()))


class _FakePopen:
    """Finished process stand-in: poll()/wait() return the fake returncode at once."""

    def __init__(self, argv, kwargs, result):
        self.args = argv
        self.pid = 0
        self.returncode = result["returncode"]
        self.stdin = None
        self.stdout = io.StringIO(result["stdout"]) if kwargs.get("stdout") == subprocess.PIPE else None
        self.stderr = io.StringIO(result["stderr"]) if kwargs.get("stderr") == subprocess.PIPE else None

    def poll(self):
        return self.returncode

    def wait(self, timeout=None):
        return self.returncode

    def communicate(self, input=None, timeout=None):
        return (self.stdout.read() if self.stdout else None,
                self.stderr.read() if self.stderr else None)

    def kill(self):
        pass

    terminate = kill

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class SubprocessRecorder:
    """Call log plus fake answers. Rules are checked in order; the first whose `match`
    is a substring of the command text answers (default: returncode 0, empty output)."""

    def __init__(self, monkeypatch):
        self._mp = monkeypatch
        self.calls = []
        self._rules = []

    def respond(self, match, returncode=0, stdout="", stderr="", raises=None, times=None):
        """Answer calls containing `match`; `raises` is raised instead; `times` limits uses."""
        self._rules.append({"match": match, "returncode": returncode, "stdout": stdout,
                            "stderr": stderr, "raises": raises, "times": times})

    def _answer(self, argv):
        text = _command_text(argv)
        for rule in self._rules:
            if rule["match"] in text:
                if rule["times"] is not None:
                    rule["times"] -= 1
                    if rule["times"] <= 0:
                        self._rules.remove(rule)
                if rule["raises"] is not None:
                    raise rule["raises"]
                return rule
        return {"returncode": 0, "stdout": "", "stderr": ""}

    def record(self, api, args, kwargs):
        self.calls.append({"api": api, "args": list(args), "kwargs": _plain_kwargs(kwargs)})

    def _run(self, *args, **kwargs):
        argv = args[0] if args else kwargs.get("args")
        self.record("subprocess.run", args, kwargs)
        result = self._answer(argv)
        return subprocess.CompletedProcess(argv, result["returncode"], result["stdout"], result["stderr"])

    def _popen(self, *args, **kwargs):
        argv = args[0] if args else kwargs.get("args")
        self.record("subprocess.Popen", args, kwargs)
        return _FakePopen(argv, kwargs, self._answer(argv))

    def _simple(self, api, value_of):
        def fake(*args, **kwargs):
            self.record(api, args, kwargs)
            return value_of(self._answer(args[0] if args else None))
        return fake

    def fake_shell32(self, is_admin=0, shell_execute_ret=42):
        """Replace ctypes.windll.shell32: ShellExecuteW is recorded, IsUserAnAdmin is fixed."""
        recorder = self

        class _Shell32:
            def IsUserAnAdmin(self):
                return is_admin

            def ShellExecuteW(self, *args):
                recorder.record("ShellExecuteW", args, {})
                return shell_execute_ret

        self._mp.setattr(ctypes.windll, "shell32", _Shell32(), raising=False)

    def install(self):
        mp = self._mp
        mp.setattr(subprocess, "run", self._run)
        mp.setattr(subprocess, "Popen", self._popen)
        for name in ("call", "check_call"):
            mp.setattr(subprocess, name, self._simple(f"subprocess.{name}", lambda r: r["returncode"]))
        mp.setattr(subprocess, "check_output", self._simple("subprocess.check_output", lambda r: r["stdout"]))
        mp.setattr(subprocess, "getoutput", self._simple("subprocess.getoutput", lambda r: r["stdout"]))
        mp.setattr(subprocess, "getstatusoutput", self._simple(
            "subprocess.getstatusoutput", lambda r: (r["returncode"], r["stdout"])))
        mp.setattr(os, "system", self._simple("os.system", lambda r: r["returncode"]))
        mp.setattr(os, "startfile", self._simple("os.startfile", lambda r: None), raising=False)
        for name in ("open", "open_new", "open_new_tab"):
            mp.setattr(webbrowser, name, self._simple(f"webbrowser.{name}", lambda r: True))
        return self


@pytest.fixture
def record_subprocess(monkeypatch):
    """Record every external call and answer it with a fake result (see SubprocessRecorder)."""
    return SubprocessRecorder(monkeypatch).install()


_SCALARS = (type(None), bool, int, float, str)


def _app_state(mod, app):
    """Snapshot of what a test can change on an App: Tk variables, filters, scalars."""
    state = {"variables": {}, "scalars": {}, "keys": set(vars(app))}
    for name, value in vars(app).items():
        if isinstance(value, mod.tk.Variable):
            state["variables"][name] = value.get()
        elif isinstance(value, _SCALARS):
            state["scalars"][name] = value
    state["filters"] = copy.deepcopy(app.filters)
    return state


def _reset_app(app, state):
    app.close_log_file()
    for name in set(vars(app)) - state["keys"]:
        if not name.startswith("_"):  # leading underscore: tkinter internals, keep
            delattr(app, name)
    for name, value in state["variables"].items():
        var = getattr(app, name)
        if var.get() != value:
            var.set(value)
    for name, value in state["scalars"].items():
        setattr(app, name, value)
    app.filters = copy.deepcopy(state["filters"])
    while not app.message_queue.empty():
        app.message_queue.get_nowait()


@pytest.fixture(scope="session")
def gui_app_session(ezbak_mod, tmp_path_factory):
    """The single App of the session. Tk() is created only once per process: after
    junction creation in-process, re-creating Tk failed now and then (tcl files not found)."""
    env = tmp_path_factory.mktemp("gui_app_env")
    with pytest.MonkeyPatch.context() as mp:
        for var in ("APPDATA", "LOCALAPPDATA"):
            (env / var).mkdir()
            mp.setenv(var, str(env / var))
        for name in ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel"):
            mp.setattr(ezbak_mod.Win11Dialog, name, lambda *a, **k: None, raising=False)
            mp.setattr(ezbak_mod.messagebox, name, lambda *a, **k: None, raising=False)
        app = ezbak_mod.App(entry_script=str(EZBAK_PATH))  # as ezBAK.py's gui.run passes
    app.withdraw()
    state = _app_state(ezbak_mod, app)
    session = types.SimpleNamespace(app=app, base=ezbak_mod.App,
                                    reset=lambda: _reset_app(app, state))
    try:
        yield session
    finally:
        app.__class__ = ezbak_mod.App
        app.destroy()


@pytest.fixture
def gui_app(gui_app_session):
    """Return bind(mod) -> the session App, reset, running `mod`'s App methods.

    bind(ezbak_mod_nopywin32) switches the instance's class to that module's App, so its
    methods (and the module globals they read) are used on the same Tk root.
    """
    session = gui_app_session

    def bind(mod):
        session.app.__class__ = mod.App
        return session.app

    session.app.__class__ = session.base
    session.reset()
    try:
        yield bind
    finally:
        session.app.__class__ = session.base
        session.reset()


# --- runpy harness for the __main__ block (Headless CLI, P3-0b): tests/fixtures/cli_harness.py ---

_spec = importlib.util.spec_from_file_location("ezbak_test_cli_harness",
                                               REPO_ROOT / "tests" / "fixtures" / "cli_harness.py")
cli_harness = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cli_harness)
HardExit = cli_harness.HardExit
AppStarted = cli_harness.AppStarted
describe_value = cli_harness.describe_value


class MainRunner(cli_harness.MainRunner):
    """cli_harness.MainRunner answering ShellExecuteW through SubprocessRecorder."""

    def __init__(self, monkeypatch, tmp_path):
        super().__init__(monkeypatch, tmp_path, SubprocessRecorder)


@pytest.fixture
def run_ezbak_main(monkeypatch, tmp_path, isolated_env, fixed_clock, block_subprocess):
    """MainRunner for this test: runner.run([...args]) -> SimpleNamespace(end, code, stdout,
    stderr, engine_calls, shell_calls). sys.argv, os._exit, tkinter.Tk, cwd and the engine
    stubs are restored by monkeypatch; external commands stay blocked."""
    return MainRunner(monkeypatch, tmp_path)
