"""Characterization (golden) of the GUI copy and free-space check cases (P5-0d), before the P5 move.

Covered (worker threads run inline):
- copy:  App.copy_files -> run_copy_thread: source dialog cancelled / failing, destination
         cancelled, the full user tree plus a single file, hidden+system included, a missing
         source, not enough free space, shutil.disk_usage failing (free space then comes
         from winsys.sysinfo's GetDiskFreeSpaceExW step, fixed here).
- space: App.check_space -> run_space_calculation, then show_space_check_result for the
         queued 'space_check_result' (as process_queue would): the same cancels, enough /
         not enough space, hidden+system included, disk_usage failing, and a skip check
         that raises (strict size walk -> "Space calculation error").
Each case runs with pywin32 and without it (ezbak_mod_nopywin32, ctypes attributes).

Fixed here: the return value, print output, dialogs, the GUI queue in order (consecutive
update_progress folded to the last one), the destination tree and its job logs. The source
dialog (SelectSourcesDialog) is replaced by a stub that returns the case's paths; its own
golden is P5-0c. Sources are temporary trees (fixtures/user_tree.py), never real user
folders. Free space is never the real disk: shutil.disk_usage is patched in every case.
Paths -> <TMP>, log stamps -> [<TS>], localized [WinError n] text -> <MSG>, traceback
frames -> <FRAMES>. Without pywin32 the size walk of the copy (is_hidden_safe, no ctypes
fallback) counts hidden files too, so the two variants differ there; both are fixed.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_gui_cases.py
"""
import importlib.util
import os
import re
import shutil
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = [pytest.mark.windows, pytest.mark.gui]

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "gui_cases_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_tool_harness_gui_cases", FIXTURES / "tool_harness.py")
harness = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(harness)

_tree_spec = importlib.util.spec_from_file_location("ezbak_test_user_tree_gui_cases", FIXTURES / "user_tree.py")
user_tree = importlib.util.module_from_spec(_tree_spec)
_tree_spec.loader.exec_module(user_tree)

BIG_FREE = 10**12
# Windows error text is localized: keep the code, drop the text (as test_backup_cases).
WINERROR_RE = re.compile(r"(\[WinError \d+\])[^:\[]*:")


# --- sources: (src root) -> list of selected paths ---

def _single(src):
    path = src / "single.txt"
    path.write_text("single file\n", encoding="utf-8")
    return str(path)


def _hidden_single(src):
    path = src / "hidden_single.txt"
    path.write_text("hidden single\n", encoding="utf-8")
    user_tree.set_attributes(str(path), user_tree.H)
    return str(path)


def _home(src):
    return user_tree.build_user_tree(str(src / user_tree.USER_NAME))


def _tree_and_file(src):
    return [_home(src), _single(src)]


def _tree_hidden_missing(src):
    return [_home(src), _single(src), _hidden_single(src), str(src / "missing")]


def _missing_first(src):
    return [str(src / "missing"), _single(src)]


def _single_only(src):
    return [_single(src)]


def _raise_on_documents(app, monkeypatch):
    original = app.is_hidden

    def is_hidden(path):
        if os.path.basename(path) == "Documents":
            raise PermissionError(f"check failed for {os.path.basename(path)}")
        return original(path)

    monkeypatch.setattr(app, "is_hidden", is_hidden, raising=False)


# sources: callable(src) -> selected paths, or "raise" (the dialog fails);
# folder False: the destination dialog is cancelled; include: hidden+system included;
# free: shutil.disk_usage free bytes, or "raise" (then sysinfo's API step answers api_free).
CASES = {
    "copy_sources_cancel": {"op": "copy", "sources": lambda src: []},
    "copy_sources_dialog_error": {"op": "copy", "sources": "raise"},
    "copy_dest_cancel": {"op": "copy", "sources": _single_only, "folder": False},
    "copy_tree": {"op": "copy", "sources": _tree_and_file},
    "copy_include_hidden_system": {"op": "copy", "sources": _tree_hidden_missing, "include": True},
    "copy_missing_source": {"op": "copy", "sources": _missing_first},
    "copy_insufficient_space": {"op": "copy", "sources": _tree_and_file, "free": 10},
    "copy_disk_usage_fail": {"op": "copy", "sources": _single_only, "free": "raise",
                             "api_free": BIG_FREE},
    "space_sources_cancel": {"op": "space", "sources": lambda src: []},
    "space_sources_dialog_error": {"op": "space", "sources": "raise"},
    "space_dest_cancel": {"op": "space", "sources": _single_only, "folder": False},
    "space_sufficient": {"op": "space", "sources": _tree_hidden_missing},
    "space_include_hidden_system": {"op": "space", "sources": _tree_hidden_missing, "include": True},
    "space_insufficient": {"op": "space", "sources": _tree_and_file, "free": 10},
    "space_disk_usage_fail": {"op": "space", "sources": _tree_and_file, "free": "raise",
                              "api_free": BIG_FREE},
    "space_check_error": {"op": "space", "sources": _tree_and_file,
                          "app_patch": _raise_on_documents},
}


def _has_pywin32(mod):
    try:
        return bool(mod._have_pywin32)
    except AttributeError:  # alias moved by P5-0a: same answer from the import itself
        return importlib.util.find_spec("win32api") is not None


def _fold_progress(messages):
    folded = []
    for message in messages:
        if folded and message[0] == "update_progress" and folded[-1][0] == "update_progress":
            folded[-1] = message
        else:
            folded.append(message)
    return folded


def _normalizer(tmp_path):
    base = harness.normalizer(tmp_path, REPO_ROOT)
    repr_root = str(tmp_path).replace("\\", "\\\\")

    def norm(value):
        if isinstance(value, str):
            value = WINERROR_RE.sub(r"\1 <MSG>:", value.replace(repr_root, "<TMP>"))
        if isinstance(value, (list, tuple)):
            return [norm(v) for v in value]
        if isinstance(value, dict):
            return {k: norm(v) for k, v in value.items()}
        return base(value)

    return norm


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
        # The globals of the function itself: ezbak_mod_nopywin32 holds its own copy of
        # winsys.sysinfo, which is no longer in sys.modules.
        monkeypatch.setitem(mod.winsys_free_space.__globals__, "_api_free_space",
                            lambda path: spec["api_free"])


def _patch_sources_dialog(mod, monkeypatch, selected, titles):
    class SourcesStub:
        """SelectSourcesDialog stand-in: records the title, returns the case's paths."""

        def __init__(self, parent, is_hidden_fn=None, title=None, **kwargs):
            titles.append(title)
            if selected == "raise":
                raise RuntimeError("sources dialog failed")
            self.selected_paths = list(selected)

    monkeypatch.setattr(mod, "SelectSourcesDialog", SourcesStub)


def _space_results(app):
    """process_queue's 'space_check_result' branch, run for each queued result; what it
    queues follows as ["<shown>", messages] (its dialog goes to "dialogs")."""
    out = []
    for task, value in harness.drain(app):
        out.append([task, value])
        if task == "space_check_result":
            app.show_space_check_result(value)
            out.append(["<shown>", harness.drain(app)])
    return out


COPY_DONE = ("copy_tree", "copy_include_hidden_system", "copy_missing_source", "copy_disk_usage_fail")


def _check_copy_total(case, spec, queue):
    """fix h-l/h-m/h-n: a finished copy fills exactly the precomputed total (last update_progress
    in the queue == update_progress_max); short space still stops the copy."""
    totals = [value for task, value in queue if task == "update_progress_max"]
    done = [i for i, (task, value) in enumerate(queue)
            if task == "log" and str(value).startswith("Copy complete")]
    if case in COPY_DONE:
        assert len(totals) == 1 and len(done) == 1, queue
        last = [value for task, value in queue if task == "update_progress"]
        assert last and last[-1] == totals[0], (totals, last)
    elif case == "copy_insufficient_space":
        assert not done and len(totals) == 1 and totals[0] > spec["free"], queue
        assert any(task == "show_error" and value.startswith("Not enough free space")
                   for task, value in queue), queue


@pytest.mark.parametrize("variant", ["pywin32", "nopywin32"])
@pytest.mark.parametrize("case", list(CASES))
def test_gui_case(case, variant, ezbak_mod, ezbak_mod_nopywin32, tmp_path, monkeypatch,
                  isolated_env, fixed_clock, block_subprocess, gui_app, capsys):
    mod = {"pywin32": ezbak_mod, "nopywin32": ezbak_mod_nopywin32}[variant]
    if variant == "pywin32" and not _has_pywin32(mod):
        pytest.skip("pywin32 not installed")
    spec = CASES[case]
    fixed_clock.inject(mod)
    src = tmp_path / "src"
    src.mkdir()
    selected = spec["sources"] if spec["sources"] == "raise" else spec["sources"](src)
    dest = tmp_path / "dest"
    dest.mkdir()
    dialogs, titles = [], []
    harness.patch_dialogs(mod, monkeypatch, dialogs,
                          folder=None if spec.get("folder") is False else str(dest))
    _patch_sources_dialog(mod, monkeypatch, selected, titles)
    _patch_free_space(mod, monkeypatch, spec)
    app = gui_app(mod)
    monkeypatch.setattr(app, "wait_window", lambda dlg: None, raising=False)
    app.progress_throttle_secs = 1e9  # restored by gui_app's reset
    if spec.get("include"):
        app.hidden_mode_var.set("include")  # restored by gui_app's reset
        app.system_mode_var.set("include")
    if "app_patch" in spec:
        spec["app_patch"](app, monkeypatch)
    harness.inline_threads(mod, monkeypatch)
    harness.drain(app)
    capsys.readouterr()
    try:
        if spec["op"] == "copy":
            returned = app.copy_files()
            queue = _fold_progress(harness.drain(app))
            _check_copy_total(case, spec, queue)
        else:
            returned = app.check_space()
            queue = _space_results(app)
    finally:
        app.close_log_file()
        app.progress_bar.stop()
        app.progress_bar["mode"] = "determinate"
    out = harness.fold_traceback(capsys.readouterr().out)  # frames: <FRAMES>
    norm = _normalizer(tmp_path)
    result = norm({
        "returned": repr(returned), "stdout": out, "source_dialog_titles": titles,
        "dialogs": dialogs, "queue": queue,
        "tree": harness.walk(dest), "logs": harness.read_logs(dest)})
    assert block_subprocess == []
    harness.check_golden(GOLDEN_PATH, f"{case}-{variant}", result)
