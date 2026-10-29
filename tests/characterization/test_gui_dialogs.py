"""Characterization (golden) of the dialogs before the P5 move (P5-0c).

Pinned in tests/fixtures/gui_dialogs_golden.json:
- win11_*: Win11Dialog showinfo/showwarning/showerror/askyesno (tree, answer, grab/wait);
- filter_*: FilterManagerDialog tree, the add-rule dialog, _add_rule/_remove/_clear/_save
  results, key bindings, and App.open_filter_manager save/cancel;
- sources_*: SelectSourcesDialog on a temp user tree (drives fixed to C: and D:), the
  check/expand steps and the _on_ok result;
- schedule_*: ScheduleBackupDialog tree and bindings, and App.schedule_backup with the
  dialog's on_create (task_data) / on_delete / close (App methods recorded, no schtasks);
- (nas_*: the NAS dialogs moved to test_gui_dialogs_nas.py, same golden file);
- folder_select_*: App.show_folder_selection_dialog tree and result.

Dialogs are Toplevels under the session App (gui_app). gui_harness.modal_hook replaces
wait_window (and grab_set/release) so a test drives the dialog where the code would wait.
No geometry, no real drives or user folders; temp paths are <TMP>. Variants: dark/en
(the App's own) and light/ko (theme and translator swapped with gui_harness.override).

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_gui_dialogs.py
"""
import contextlib
import copy
import importlib.util
import json
import os
import re
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.gui

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "gui_dialogs_golden.json"


def _load(name, file):
    spec = importlib.util.spec_from_file_location(name, FIXTURES / file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


gh = _load("ezbak_test_gui_harness_dialogs", "gui_harness.py")
th = _load("ezbak_test_tool_harness_dialogs", "tool_harness.py")
user_tree = _load("ezbak_test_user_tree_dialogs", "user_tree.py")

VARIANTS = {"dark_en": ("dark", "en"), "light_ko": ("light", "ko")}
DRIVE_RE = re.compile(r"^[A-Z]:\\$")
FAKE_DRIVES = ("C:\\", "D:\\")


@pytest.fixture
def dlg(ezbak_mod, isolated_env, gui_app, block_subprocess, tmp_path):
    """Session App plus a normalizer; afterwards no dialog is left, the App's instance dict
    and filters are as before and no external command was attempted."""
    app = gui_app(ezbak_mod)
    with gh.app_unchanged(app):
        yield SimpleNamespace(mod=ezbak_mod, app=app, tmp=tmp_path, env=isolated_env,
                              norm=gh.path_normalizer(tmp_path), dialogs=[])
    assert block_subprocess == []


def golden(name, result):
    th.check_golden(GOLDEN_PATH, name, result)


@contextlib.contextmanager
def variant(s, name):
    theme, lang = VARIANTS[name]
    with gh.override(s.app, theme=s.mod.Win11Theme(initial_mode=theme),
                     translator=s.mod.Translator(initial_lang=lang)):
        yield


def record_dialogs(s, monkeypatch, yes=None, folder=None):
    th.patch_dialogs(s.mod, monkeypatch, s.dialogs, yes=yes, folder=folder)


# --- Win11Dialog ----------------------------------------------------------------------

def _win11(s, kind, name, press, **kwargs):
    seen = {}

    def handler(window):
        seen["snapshot"] = gh.dialog_snapshot(window)
        if press:
            gh.button(window, press).invoke()

    theme = translator = None
    if name != "dark_en":
        theme = s.mod.Win11Theme(initial_mode=VARIANTS[name][0])
        translator = s.mod.Translator(initial_lang=VARIANTS[name][1])
    with gh.modal_hook(handler) as events:
        returned = getattr(s.mod.Win11Dialog, kind)("Title", "Line one\nLine two",
                                                     parent=s.app, theme=theme,
                                                     translator=translator, **kwargs)
    return {"returned": returned, "events": events, "snapshot": seen["snapshot"]}


@pytest.mark.parametrize("name", list(VARIANTS))
@pytest.mark.parametrize("kind", ["showinfo", "showwarning", "showerror", "showinfo_noicon"])
def test_win11_message(dlg, kind, name):
    ok = "OK" if name == "dark_en" else "확인"
    extra = {"show_icon": False} if kind == "showinfo_noicon" else {}
    golden(f"win11_{kind}_{name}", _win11(dlg, kind.replace("_noicon", ""), name, ok, **extra))


@pytest.mark.parametrize("name", list(VARIANTS))
def test_win11_askyesno(dlg, name):
    labels = {"dark_en": ("Yes", "No"), "light_ko": ("예", "아니오")}[name]
    result = _win11(dlg, "askyesno", name, labels[0])
    for press in (labels[1], None):  # No, then closing the window
        other = _win11(dlg, "askyesno", name, press)
        assert other["snapshot"] == result["snapshot"]
        result[f"returned_{'no' if press else 'closed'}"] = other["returned"]
    golden(f"win11_askyesno_{name}", result)


# --- Filter manager -------------------------------------------------------------------

START_FILTERS = {"include": [{"type": "ext", "pattern": ".docx"}],
                 "exclude": [{"type": "name", "pattern": "Thumbs.db"},
                             {"type": "path", "pattern": "AppData\\Temp"},
                             {"type": "bogus", "pattern": "y"}, {"pattern": "no-type"}]}


def _lists(window):
    return {"inc": list(window.inc_list.get(0, "end")), "exc": list(window.exc_list.get(0, "end")),
            "inc_rules": copy.deepcopy(window.inc), "exc_rules": copy.deepcopy(window.exc)}


def _add_dialog(window):
    (child,) = gh.toplevels(window)
    return child


@pytest.mark.parametrize("name", list(VARIANTS))
def test_filter_tree(dlg, name):
    with variant(dlg, name), gh.modal_hook(lambda w: None) as events:
        window = dlg.mod.FilterManagerDialog(dlg.app, current_filters=START_FILTERS)
        result = {"snapshot": gh.dialog_snapshot(window), "lists": _lists(window)}
        for kind in ("include", "exclude"):
            window._add_rule(kind)
            child = _add_dialog(window)
            result[f"add_{kind}"] = gh.dialog_snapshot(child)
            child.destroy()
        result["bindings"] = gh.window_bindings(window)
        window.destroy()
    result["events"] = events
    golden(f"filter_tree_{name}", result)


def test_filter_edit_and_save(dlg, monkeypatch):
    record_dialogs(dlg, monkeypatch)
    tk = dlg.mod.tk
    steps = []
    with gh.modal_hook(lambda w: None):
        window = dlg.mod.FilterManagerDialog(dlg.app, current_filters=START_FILTERS)
        window._add_rule("include")
        child = _add_dialog(window)
        examples = {}
        for radio in gh.descendants(child, tk.Radiobutton):
            radio.invoke()
            examples[radio.cget("value")] = gh.descendants(child, tk.Label)[-1].cget("text")
        gh.type_into(gh.descendants(child, tk.Entry)[0], "  Thumbs.db  ")
        gh.button(child, "Add Rule").invoke()
        steps.append(["add include name", child.winfo_exists(), _lists(window)])
        window._add_rule("exclude")
        child = _add_dialog(window)
        gh.button(child, "Add Rule").invoke()  # empty pattern: warning, stays open
        steps.append(["add empty", child.winfo_exists(), list(dlg.dialogs)])
        gh.invoke_binding(child, "<Escape>")
        steps.append(["escape", child.winfo_exists(), _lists(window)])
        window._add_rule("exclude")
        child = _add_dialog(window)
        gh.descendants(child, tk.Radiobutton)[2].invoke()
        gh.type_into(gh.descendants(child, tk.Entry)[0], "C:\\Temp")
        gh.invoke_binding(child, "<Return>")
        steps.append(["add exclude path (Return)", child.winfo_exists(), _lists(window)])
        window.inc_list.selection_set(0)
        window._remove_selected("include")
        window.exc_list.selection_set(1, 2)
        window._remove_selected("exclude")
        steps.append(["remove selected", _lists(window)])
        window._clear("include")
        steps.append(["clear include", _lists(window)])
        window._save()
        result = {"examples": examples, "steps": steps, "saved": window.result,
                  "open_after_save": window.winfo_exists()}
        window = dlg.mod.FilterManagerDialog(dlg.app, current_filters=None)
        result["none_filters"] = _lists(window)
        window._cancel()
        result["cancel_result"] = window.result
    golden("filter_edit_and_save", result)


@pytest.mark.parametrize("action", ["save", "cancel"])
def test_filter_from_app(dlg, action):
    app = dlg.app
    saved_filters = copy.deepcopy(app.filters)

    def handler(window):
        window._add_rule("include")
        child = _add_dialog(window)
        gh.type_into(gh.descendants(child, dlg.mod.tk.Entry)[0], ".pdf")
        gh.invoke_binding(child, "<Return>")
        gh.invoke_binding(window, "<Control-s>" if action == "save" else "<Escape>")

    with gh.modal_hook(handler) as events:
        app.open_filter_manager()
    settings = Path(dlg.env["APPDATA"]) / "ezbak" / "settings.json"
    saved = json.loads(settings.read_text(encoding="utf-8"))["filters"] if settings.exists() else None
    result = {"events": events, "filters": copy.deepcopy(app.filters), "queue": th.drain(app),
              "settings_filters": saved}
    app.filters = saved_filters
    golden(f"filter_from_app_{action}", result)


# --- Select sources -------------------------------------------------------------------

@pytest.fixture
def fake_drives(monkeypatch, ezbak_mod):
    real = os.path.exists
    monkeypatch.setattr(ezbak_mod.os.path, "exists",
                        lambda p: p in FAKE_DRIVES if DRIVE_RE.match(str(p)) else real(p))


def _items(window, s, item=""):
    out = []
    for child in window.tree.get_children(item):
        node = {"text": s.norm(window.tree.item(child, "text")),
                "open": bool(window.tree.item(child, "open")), "checked": child in window._checked}
        if child in window._item_path:
            node["path"] = s.norm(window._item_path[child])
        kids = _items(window, s, child)
        if kids:
            node["children"] = kids
        out.append(node)
    return out


def _find(window, text_end, item=""):
    for child in window.tree.get_children(item):
        if window.tree.item(child, "text").endswith(text_end):
            return child
    raise AssertionError(text_end)


def _sources(s, monkeypatch):
    home = str(s.tmp / "src")
    if not os.path.isdir(home):
        user_tree.build_user_tree(home)
    record_dialogs(s, monkeypatch, folder=home.replace("\\", "/"))
    return s.mod.SelectSourcesDialog(s.app, is_hidden_fn=s.app.is_hidden,
                                     title=s.app.translator.get("select_sources"))


@pytest.mark.windows
@pytest.mark.parametrize("name", list(VARIANTS))
def test_sources_tree(dlg, monkeypatch, fake_drives, name):
    with variant(dlg, name), gh.modal_hook(lambda w: None) as events:
        window = _sources(dlg, monkeypatch)
        result = {"snapshot": gh.dialog_snapshot(window), "items": _items(window, dlg),
                  "bindings": gh.window_bindings(window)}
        window.destroy()
    result["events"] = events
    golden(f"sources_tree_{name}", result)


@pytest.mark.windows
def test_sources_select_and_ok(dlg, monkeypatch, fake_drives):
    steps = []
    with gh.modal_hook(lambda w: None):
        window = _sources(dlg, monkeypatch)

        def step(label, *calls):
            for call in calls:
                call()
            steps.append([label, window.status_label.cget("text"), _items(window, dlg)])

        def open_(item):
            window.tree.focus(item)
            window._on_open(None)

        def toggle(item):
            window.tree.focus(item)
            window._toggle_selected()

        step("refresh", window._refresh_tree)
        step("add root", window._choose_root)
        root = _find(window, "src")
        step("open root", lambda: open_(root))
        docs = _find(window, "Documents", root)
        step("open Documents, check it", lambda: open_(docs), lambda: toggle(docs))
        step("open checked sub", lambda: open_(_find(window, "sub", docs)))
        step("uncheck report.txt", lambda: toggle(_find(window, "report.txt", docs)))
        step("check Desktop (closed)", lambda: toggle(_find(window, "Desktop", root)))
        step("select all", window._select_all)
        step("clear all", window._clear_all)
        step("check Pictures and Documents/notes.md",
             lambda: toggle(_find(window, "Pictures", root)),
             lambda: toggle(_find(window, "notes.md", docs)))
        window._on_ok()
        # _on_ok returns selected_paths in tree order (#13); the golden key keeps its sorted form
        result = {"steps": steps, "selected_sorted": sorted(dlg.norm(window.selected_paths)),
                  "open_after_ok": window.winfo_exists(), "dialogs": dlg.norm(dlg.dialogs)}
        window = _sources(dlg, monkeypatch)
        window._toggle_item(_find(window, "D:\\"))
        window._on_cancel()
        result["cancel"] = [window.selected_paths, window.winfo_exists()]
    golden("sources_select_and_ok", result)


@pytest.mark.windows
def test_sources_ok_tree_order_and_separator(dlg, monkeypatch, fake_drives):
    """h-k: _on_ok returns paths in tree order (not check order) with '\\' only."""
    with gh.modal_hook(lambda w: None):
        window = _sources(dlg, monkeypatch)
        window._choose_root()
        root = _find(window, "src")
        window.tree.focus(root)
        window._on_open(None)
        docs = _find(window, "Documents", root)
        window.tree.focus(docs)
        window._on_open(None)
        # checked in an order that differs from the tree order
        for item in (_find(window, "Pictures", root), _find(window, "notes.md", docs),
                     _find(window, "D:\\"), _find(window, "Desktop", root)):
            window._toggle_item(item)
        window._on_ok()
    home = str(dlg.tmp / "src")
    assert window.selected_paths == ["D:\\", os.path.join(home, "Desktop"),
                                     os.path.join(home, "Documents", "notes.md"),
                                     os.path.join(home, "Pictures")]
    assert not any("/" in p for p in window.selected_paths)


# --- Schedule -------------------------------------------------------------------------

@pytest.mark.parametrize("name", list(VARIANTS))
def test_schedule_tree(dlg, name):
    dlg.app.user_var.set("alice")
    with variant(dlg, name), gh.modal_hook(lambda w: None) as events:
        window = dlg.mod.ScheduleBackupDialog(dlg.app)
        result = {"snapshot": gh.dialog_snapshot(window), "filters": window.filters}
        window.schedule_var.set("Monthly")
        result["monthly_combo_state"] = str(window.monthly_day_combo.cget("state"))
        window.schedule_var.set("Daily")
        result["daily_combo_state"] = str(window.monthly_day_combo.cget("state"))
        result["bindings"] = gh.window_bindings(window)
        window.destroy()
    result["events"] = events
    golden(f"schedule_tree_{name}", result)


SCHEDULE_CASES = {
    "create_weekly": {"task": "Nightly", "dest": "dest", "schedule": "Weekly", "hour": "23",
                      "minute": "30", "hidden": True, "retention": "5", "days": "7"},
    "create_monthly_untouched": {"dest": "dest", "schedule": "Monthly", "day": "15"},
    "create_name_spaces": {"task": "  Spaced  ", "dest": "dest", "system": True},
    "create_no_name": {"task": "  ", "dest": "dest"},
    "create_no_dest": {"task": "T"},
    "create_missing_dest_no": {"task": "T", "dest": "missing", "yes": [False]},
    "create_missing_dest_yes": {"task": "T", "dest": "missing", "yes": [True]},
    "create_bad_time": {"task": "T", "dest": "dest", "hour": "25"},
    "create_bad_retention": {"task": "T", "dest": "dest", "retention": "x"},
    "delete": {"task": " Nightly ", "method": "on_delete"},
    "delete_no_name": {"task": "", "method": "on_delete"},
    "close": {"method": "on_close"},
}
SCHEDULE_FIELDS = {"task": "task_var", "schedule": "schedule_var", "hour": "hour_var",
                   "minute": "minute_var", "day": "monthly_day_var", "hidden": "hidden_var",
                   "system": "system_var", "retention": "retention_count_var",
                   "days": "log_retention_days_var"}


@pytest.mark.parametrize("case", list(SCHEDULE_CASES))
def test_schedule_case(dlg, monkeypatch, case):
    spec = SCHEDULE_CASES[case]
    (dlg.tmp / "dest").mkdir()
    record_dialogs(dlg, monkeypatch, yes=spec.get("yes"))
    th.inline_threads(dlg.mod, monkeypatch)
    dlg.app.user_var.set("alice")
    seen = {}

    def handler(window):
        seen["window"] = window
        for key, attr in SCHEDULE_FIELDS.items():
            if key in spec:
                getattr(window, attr).set(spec[key])
        if "dest" in spec:
            window.dest_var.set(str(dlg.tmp / spec["dest"]))
        getattr(window, spec.get("method", "on_create"))()
        seen["open_after"] = window.winfo_exists()

    calls = []
    with gh.recording(dlg.app, "app", calls, keep=("schedule_backup",)), \
            gh.modal_hook(handler) as events:
        dlg.app.schedule_backup()
    window = seen["window"]
    golden(f"schedule_{case}", dlg.norm({
        "action": window.action, "result": window.result, "task_name": window.task_name,
        "open_after": seen["open_after"], "app_calls": calls, "dialogs": dlg.dialogs,
        "queue": th.drain(dlg.app), "events": events}))


# --- Restore folder selection ---------------------------------------------------------

FOLDERS = ["alice_20240102_030405", "alice_20231231_235959", "bob_20230101_000000"]


def _select(s, folders, handler):
    with gh.modal_hook(handler) as events:
        s.app.show_folder_selection_dialog(list(folders))
    return events, vars(s.app).pop("_folder_selection_result")


@pytest.mark.parametrize("name", list(VARIANTS))
def test_folder_select_tree(dlg, name):
    seen = {}

    def handler(window):
        seen["snapshot"] = gh.dialog_snapshot(window)
        (listbox,) = gh.descendants(window, dlg.mod.tk.Listbox)
        seen["list"] = [list(listbox.get(0, "end")), list(listbox.curselection())]

    with variant(dlg, name):
        events, chosen = _select(dlg, FOLDERS, handler)
    golden(f"folder_select_tree_{name}", {**seen, "events": events, "closed_result": chosen})


FOLDER_CASES = {"ok_default": (None, "OK"), "ok_third": (2, "OK"), "double_click_second": (1, None),
                "cancel": (1, "Cancel"), "none_selected_ok": ("clear", "OK")}


@pytest.mark.parametrize("case", list(FOLDER_CASES))
def test_folder_select_case(dlg, case):
    index, press = FOLDER_CASES[case]

    def handler(window):
        (listbox,) = gh.descendants(window, dlg.mod.tk.Listbox)
        if index is not None:
            listbox.selection_clear(0, "end")
            if index != "clear":
                listbox.selection_set(index)
        if press:
            gh.button(window, press).invoke()
        else:
            gh.invoke_binding(listbox, "<Double-Button-1>")

    events, chosen = _select(dlg, FOLDERS, handler)
    golden(f"folder_select_{case}", {"result": chosen, "events": events})


def test_folder_select_empty(dlg):
    events, chosen = _select(dlg, [], lambda w: gh.button(w, "OK").invoke())
    golden("folder_select_empty", {"result": chosen, "events": events})
