"""runpy harness for the __main__ block of ezBAK.py (Headless CLI, P3-0b; thin entry P5-16b).

Used by the conftest.py fixture run_ezbak_main (loaded there by file path, like
user_tree.py). The subprocess recorder class is passed in by conftest, so this file
does not import conftest.
"""
import contextlib
import dataclasses
import io
import os
import re
import runpy
import sys
import types
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
EZBAK_PATH = REPO_ROOT / "ezBAK.py"

_SCALARS = (type(None), bool, int, float, str)


class HardExit(BaseException):
    """os._exit(code) under run_ezbak_main (the real one would end pytest)."""


class AppStarted(BaseException):
    """App() reached under run_ezbak_main: tkinter.Tk.__init__ is replaced by this raise."""


_TS_RE = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")


def describe_value(value):
    """JSON-friendly view of an engine argument (class constants by name, dataclasses as dicts)."""
    if isinstance(value, _SCALARS):
        return value
    if isinstance(value, (list, tuple)):
        return [describe_value(v) for v in value]
    if isinstance(value, dict):
        return {str(k): describe_value(v) for k, v in value.items()}
    cls = type(value)
    for name in dir(cls):
        if not name.startswith("_") and getattr(cls, name, None) is value:
            return f"{cls.__name__}.{name}"
    if dataclasses.is_dataclass(value):
        return describe_value(dataclasses.asdict(value))
    if callable(value) and not isinstance(value, type) and getattr(value, "__name__", "") == "<lambda>":
        return "<lambda>"
    return f"<{cls.__name__}>"


class MainRunner:
    """run(argv) executes ezBAK.py as __main__; normalize(text) masks <TMP>/<REPO>/<PY>/<TS>."""

    def __init__(self, monkeypatch, tmp_path, recorder_cls):
        self._mp = monkeypatch
        self.tmp_path = tmp_path
        self._recorder_cls = recorder_cls

    def normalize(self, text):
        for root, mark in ((str(self.tmp_path), "<TMP>"), (str(REPO_ROOT), "<REPO>"),
                           (sys.executable, "<PY>")):
            text = text.replace(root.replace("\\", "\\\\"), mark).replace(root, mark)
        return _TS_RE.sub("<TS>", text)

    def _stub_engine(self, calls, engine_error):
        import ezbak.core
        original = ezbak.core.run_backup

        def stub(*args, **kwargs):
            names = ("job", "backup_path", "job_log", "fmt", "store", "targets",
                     "reporter", "policy", "clock")
            call = {n: describe_value(v) for n, v in zip(names, args)}
            call.update({k: describe_value(v) for k, v in kwargs.items()})
            if len(args) > 8 and callable(args[8]):
                call["clock"] = str(args[8]())
            calls.append(call)
            if engine_error is not None:
                raise engine_error

        for name, module in list(sys.modules.items()):
            if module is not None and (name == "ezbak" or name.startswith("ezbak.")):
                for attr, value in list(vars(module).items()):
                    if value is original:
                        self._mp.setattr(module, attr, stub)

    def run(self, argv, *, engine_error=None, is_admin=0, shell_execute_ret=42):
        """Run with sys.argv = [ezBAK.py, *argv]. engine_error: exception the stubbed
        ezbak.core.run_backup raises (None: returns normally)."""
        import tkinter
        mp = self._mp
        engine_calls = []
        shell = self._recorder_cls(mp)
        shell.fake_shell32(is_admin=is_admin, shell_execute_ret=shell_execute_ret)
        self._stub_engine(engine_calls, engine_error)

        def hard_exit(code):
            raise HardExit(code)

        def no_tk(self, *args, **kwargs):
            raise AppStarted()

        mp.setattr(os, "_exit", hard_exit)
        mp.setattr(tkinter.Tk, "__init__", no_tk)
        mp.delattr(sys, "frozen", raising=False)
        mp.setattr(sys, "argv", [str(EZBAK_PATH), *argv])
        mp.chdir(self.tmp_path)
        before = set(sys.modules)
        main_before = sys.modules.get("__main__")
        out, err = io.StringIO(), io.StringIO()
        end = {"end": "return", "code": None}
        try:
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                runpy.run_path(str(EZBAK_PATH), run_name="__main__")
        except SystemExit as exc:
            end = {"end": "SystemExit", "code": exc.code}
        except HardExit as exc:
            end = {"end": "os._exit", "code": exc.args[0]}
        except AppStarted:
            end = {"end": "App()", "code": None}
        finally:
            for name in set(sys.modules) - before:
                if name == "ezBAK" or name == "ezbak" or name.startswith("ezbak."):
                    sys.modules.pop(name, None)
            assert sys.modules.get("__main__") is main_before
        return types.SimpleNamespace(
            stdout=out.getvalue(), stderr=err.getvalue(), engine_calls=engine_calls,
            shell_calls=[c for c in shell.calls if c["api"] == "ShellExecuteW"], **end)
