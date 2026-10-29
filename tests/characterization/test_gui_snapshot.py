"""Characterization (golden) of the main App window before the P5 move (P5-0b, D6).

Pinned in tests/fixtures/gui_snapshot_golden.json:
- tree_<lang>_<theme>: the window as App.__init__ builds it when settings.json says
  that language and theme (create_widgets + update_ui_texts on a fresh host, see
  gui_harness.make_host), with the user list from a temp USERS_ROOT (menu: None since U5-7);
- bindings: every key binding of the App and every settings/help page button (the former
  menu commands, U5-7), as the App or
  KeyboardShortcuts method each one calls (methods replaced by recorders);
- theme_switch / language_switch: what switch_theme / switch_language change on the
  live App (root colour, progress style, sound toggle, texts, menu, saved settings, dialog);
- translations: Translator.TRANSLATIONS in key order, and the get() fallback;
- settings_*: settings.json written by save_settings and read by load_settings.

Widget snapshots hold no pixel geometry and no path names (children in creation order);
clock texts are <TS>. settings.json lives in isolated_env's temp %APPDATA% only.
Each test puts the session App's theme, language, variables and filters back.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_gui_snapshot.py
"""
import copy
import importlib.util
import json
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.gui

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "gui_snapshot_golden.json"


def _load(name, file):
    spec = importlib.util.spec_from_file_location(name, FIXTURES / file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


gh = _load("ezbak_test_gui_harness", "gui_harness.py")
check_golden = _load("ezbak_test_tool_harness_gui", "tool_harness.py").check_golden

USERS = ("alice", "bob", "Public")
PAGE_BUTTONS = ("theme_light_btn", "theme_dark_btn", "lang_en_btn", "lang_ko_btn",
                "help_shortcuts_btn", "help_github_btn")
# Texts that follow running state: the status line, the chips (progress bar and filters,
# on a timer) and the home user texts (the user list, loaded from the real C:\Users when
# the session App was made). Masked in tree comparisons, left out of language_switch.
LIVE_TEXTS = ("status_label", "nav_status_label", "home_filter_chip", "home_user_label",
              "home_user_count_label")
STYLE = "green.Horizontal.TProgressbar"


def golden(name, result):
    check_golden(GOLDEN_PATH, name, result)


@pytest.fixture
def snap(ezbak_mod, isolated_env, gui_app, user_root, monkeypatch):
    """Session App (dark/en) with dialogs recorded; theme, language, Tk variables, filters
    and the status text are put back before the temp %APPDATA% goes away."""
    mod = ezbak_mod
    app = gui_app(mod)
    for name in USERS:
        (user_root.path / name).mkdir()
    user_root.inject(mod)
    dialogs = []

    def recorder(kind):
        def dialog(title, message, *args, **kwargs):
            dialogs.append([kind, str(title), str(message)])
            return False
        return dialog

    for kind in ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel"):
        monkeypatch.setattr(mod.Win11Dialog, kind, recorder(kind), raising=False)
        monkeypatch.setattr(mod.messagebox, kind, recorder(kind), raising=False)
    theme = (app.theme.current_mode, app.theme.colors)
    lang = app.translator.current_lang
    variables = {n: v.get() for n, v in vars(app).items() if isinstance(v, mod.tk.Variable)}
    filters = copy.deepcopy(app.filters)
    status = app.status_label.cget("text")
    try:
        yield SimpleNamespace(mod=mod, app=app, dialogs=dialogs, env=isolated_env)
    finally:
        app.theme.current_mode, app.theme.colors = theme
        app.translator.current_lang = lang
        for name, value in variables.items():
            if getattr(app, name).get() != value:
                getattr(app, name).set(value)
        app.filters = filters
        app.apply_theme_to_all_widgets()
        app.update_ui_texts()
        app.status_label.config(text=status)


def _settings_file(s):
    return Path(s.env["APPDATA"]) / "ezbak" / "settings.json"


def _settings_text(s):
    path = _settings_file(s)
    return path.read_text(encoding="utf-8") if path.exists() else None


def _host_window(s, theme, lang):
    """Snapshot of a fresh App window build in theme/lang (DEBUG prints must not appear)."""
    host = gh.make_host(s.mod, s.app, theme, lang)
    try:
        host.create_widgets()
        host.update_ui_texts()
        namer = gh.Namer(host)
        return {"title": host.title(),
                "menu": gh.menubar(host, namer),
                "users": gh.plain(host.user_combo.cget("values")),
                "user": host.user_var.get(),
                "tree": gh.snapshot_tree(host, namer)}
    finally:
        host.destroy()


@pytest.mark.parametrize("lang", ["en", "ko"])
@pytest.mark.parametrize("theme", ["dark", "light"])
def test_window_tree(snap, capsys, theme, lang):
    result = _host_window(snap, theme, lang)
    assert "DEBUG" not in capsys.readouterr().out
    assert _settings_text(snap) is None  # building the window saves nothing
    golden(f"tree_{lang}_{theme}", result)


def _first_difference(a, b, path="tree"):
    """Path and values of the first place where two snapshots differ (failure message)."""
    if isinstance(a, dict) and isinstance(b, dict):
        for key in sorted(set(a) | set(b), key=str):
            if a.get(key) != b.get(key):
                return _first_difference(a.get(key), b.get(key), f"{path}.{key}")
    elif isinstance(a, list) and isinstance(b, list) and len(a) == len(b):
        for i, (x, y) in enumerate(zip(a, b)):
            if x != y:
                return _first_difference(x, y, f"{path}[{i}]")
    return f"{path}: live={a!r:.200} fresh={b!r:.200}"


def test_live_app_matches_fresh_build(snap):
    """The session App (created dark/en by App.__init__) has the tree of a fresh build:
    proves make_host builds what __init__ builds. Running state is masked."""
    fresh = _host_window(snap, "dark", "en")
    namer = gh.Namer(snap.app)
    live = gh.snapshot_tree(snap.app, namer)
    live, fresh_tree = gh.mask_live(live, LIVE_TEXTS), gh.mask_live(fresh["tree"], LIVE_TEXTS)
    assert live == fresh_tree, _first_difference(live, fresh_tree)
    assert gh.menubar(snap.app, namer) == fresh["menu"]
    assert snap.app.title() == fresh["title"]


def test_bindings_and_menu_commands(snap):
    app = snap.app
    calls = []
    keys, buttons = {}, {}
    before = (dict(vars(app)), dict(vars(app.shortcuts)))
    shortcuts = app.shortcuts
    with gh.recording(app, "app", calls), \
            gh.recording(shortcuts, "shortcuts", calls, keep=("safe_execute",)):
        for seq in sorted(gh.own_binding_sequences(app)):
            gh.invoke_binding(app, seq)
            keys[seq], calls[:] = list(calls), []
        for name in PAGE_BUTTONS:  # the former menu commands (U5-7, D1=B)
            getattr(app, name).invoke()
            buttons[name], calls[:] = list(calls), []
    assert snap.dialogs == []
    assert (dict(vars(app)), dict(vars(app.shortcuts))) == before
    assert str(app.cget("menu")) == ""
    golden("bindings", {"keys": keys, "page_buttons": buttons})


def _theme_state(s):
    app = s.app
    return {"mode": app.theme.current_mode,
            "root_bg": str(app.cget("bg")),
            "progress_style": gh.plain(app.style.configure(STYLE)),
            "toggle_bg": str(app.sound_toggle.cget("bg")),
            "toggle_items": gh.canvas_items(app.sound_toggle),
            "settings": _settings_text(s),
            "dialogs": list(s.dialogs)}


def test_theme_switch(snap):
    steps = []
    for mode in ("dark", "light", "light", "dark"):
        snap.dialogs.clear()
        snap.app.switch_theme(mode)
        steps.append([mode, _theme_state(snap)])
    golden("theme_switch", steps)


def _texts(s):
    """{attr: text} of the App's widget attributes (tk or CTk) that have a text option,
    LIVE_TEXTS except status_label (set by the test)."""
    out = {}
    for name, value in sorted(vars(s.app).items()):
        if isinstance(value, s.mod.tk.Misc) and (name == "status_label" or name not in LIVE_TEXTS):
            text = gh.widget_text(value)
            if text is not None:
                out[name] = gh.plain(text)
    return out


def test_language_switch(snap):
    app = snap.app
    app.status_label.config(text=app.translator.get("select_user_begin"))
    steps = []
    for lang in ("ko", "xx", "en"):
        snap.dialogs.clear()
        returned = app.switch_language(lang)
        steps.append([lang, {"returned": returned, "lang": app.translator.current_lang,
                             "title": app.title(), "texts": _texts(snap),
                             "menu": gh.menubar(app, gh.Namer(app)),
                             "settings": _settings_text(snap), "dialogs": list(snap.dialogs)}])
    golden("language_switch", steps)


def test_translations(ezbak_mod):
    table = ezbak_mod.Translator.TRANSLATIONS
    translator = ezbak_mod.Translator()
    fallback = [translator.get("no_such_key"), translator.get("no_such_key", "dflt"),
                translator.set_language("xx"), translator.current_lang]
    translator.set_language("ko")
    fallback += [translator.get("no_such_key"), translator.get("app_title")]
    golden("translations", {"langs": list(table),
                            "tables": {lang: [[k, v] for k, v in entries.items()]
                                       for lang, entries in table.items()},
                            "fallback": fallback})


def _state(app):
    return {"hidden": app.hidden_mode_var.get(), "system": app.system_mode_var.get(),
            "retention_count": app.retention_count_var.get(),
            "log_retention_days": app.log_retention_days_var.get(),
            "sound_enabled": app.sound_enabled_var.get(), "filters": app.filters,
            "theme": app.theme.current_mode, "theme_bg": app.theme.get("bg"),
            "language": app.translator.current_lang}


def test_settings_save_and_round_trip(snap):
    app = snap.app
    path = app.settings_file_path()
    assert Path(path) == _settings_file(snap)
    app.hidden_mode_var.set("include")
    app.retention_count_var.set("5")
    app.log_retention_days_var.set("7")
    app.sound_enabled_var.set(False)
    app.filters = {"include": [{"type": "ext", "pattern": ".docx"}],
                   "exclude": [{"type": "name", "pattern": "Thumbs.db"},
                               {"type": "path", "pattern": "AppData\\Temp"}]}
    app.theme.current_mode, app.theme.colors = "light", app.theme.LIGHT
    app.translator.current_lang = "ko"
    app.save_settings()
    text = _settings_text(snap)
    saved = _state(app)
    app.hidden_mode_var.set("exclude")
    app.retention_count_var.set("2")
    app.filters = {"include": [], "exclude": []}
    app.theme.current_mode, app.theme.colors = "dark", app.theme.DARK
    app.translator.current_lang = "en"
    # the variable traces saved the reverted values; put the first save back, then load it
    _settings_file(snap).write_text(text, encoding="utf-8")
    app.load_settings()
    assert _state(app) == saved
    golden("settings_save", {"path": os.path.relpath(path, snap.env["APPDATA"]),
                             "text": text, "state": saved})


LOAD_CASES = {
    "full": json.dumps({
        "hidden": "include", "system": "include", "retention_count": 3,
        "log_retention_days": "14", "sound_enabled": False, "theme": "light",
        "language": "ko",
        "filters": {"include": [{"type": " EXT ", "pattern": " .pdf "},
                                {"type": "bogus", "pattern": "x"},
                                {"type": "name", "pattern": "  "}, "not-a-dict"],
                    "exclude": [{"type": "Path", "pattern": "Temp"}, {"pattern": "y"}]}}),
    "empty_object": "{}",
    "odd_values": json.dumps({"theme": "blue", "language": "fr", "filters": ["x"],
                              "retention_count": None}),
    "not_json": "{not json",
    "not_object": "[1, 2]",
    "missing": None,
}


@pytest.mark.parametrize("case", list(LOAD_CASES))
def test_settings_load(snap, capsys, case):
    path = _settings_file(snap)
    path.parent.mkdir(parents=True, exist_ok=True)
    if LOAD_CASES[case] is not None:
        path.write_text(LOAD_CASES[case], encoding="utf-8")
    capsys.readouterr()
    snap.app.load_settings()
    golden(f"settings_load_{case}", {"state": _state(snap.app),
                                     "stdout": capsys.readouterr().out.splitlines(),
                                     "file_after": _settings_text(snap)})
