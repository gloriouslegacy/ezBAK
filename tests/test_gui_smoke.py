"""GUI smoke tests (T0-4b): App without a visible window, main widgets, queue messages.

App.process_queue handles 14 message kinds (13 + ask_yes_no, g-e). Each test puts one message, calls
process_queue once and checks what the current code shows (progress, log text,
status bar, button state, dialog calls). Dialogs are replaced, after() does not
re-schedule, external commands are blocked (block_subprocess).
"""
import inspect
import logging
import threading
import time
import types

import pytest

pytestmark = pytest.mark.gui

QUEUE_KINDS = (
    "log", "update_progress", "space_check_result", "update_progress_max",
    "update_status", "enable_buttons", "show_error", "open_folder",
    "stop_progress", "play_sound", "reset_progress", "show_cross_user_warning",
    "select_folder", "ask_yes_no",
)

ACTION_BUTTONS = (
    "backup_btn", "restore_btn", "copy_btn", "driver_backup_btn", "driver_restore_btn",
    "devmgr_btn", "file_explorer_btn", "save_log_btn", "check_space_btn", "schedule_btn",
    "winget_export_btn", "filters_btn", "nas_connect_btn", "nas_disconnect_btn",
    "browser_profiles_btn",
)


class DialogRecorder:
    """Stands in for Win11Dialog's modal helpers; records (name, title, message)."""

    def __init__(self, answer=True):
        self.calls = []
        self.answer = answer

    def make(self, name):
        def fake(title, message, *args, **kwargs):
            self.calls.append((name, title, message))
            return self.answer
        return fake


@pytest.fixture
def smoke(ezbak_mod, gui_app, block_subprocess, monkeypatch):
    """Session App with dialogs replaced and after() recorded instead of scheduled."""
    app = gui_app(ezbak_mod)
    dialogs = DialogRecorder()
    for name in ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel"):
        monkeypatch.setattr(ezbak_mod.Win11Dialog, name, dialogs.make(name), raising=False)
        monkeypatch.setattr(ezbak_mod.messagebox, name, dialogs.make(name), raising=False)
    scheduled = []
    app.after = lambda ms, func=None, *args: scheduled.append((ms, func)) or "after#test"
    try:
        yield types.SimpleNamespace(app=app, mod=ezbak_mod, dialogs=dialogs,
                                    scheduled=scheduled, blocked=block_subprocess)
    finally:
        del app.after
        app.reset_progress_bar()
        app.set_buttons_state("normal")


def run_one(s, task, value):
    """Queue one message and process the queue once; the queue must end empty."""
    s.app.message_queue.put((task, value))
    s.app.process_queue()
    assert s.app.message_queue.empty()
    assert s.scheduled == [(100, s.app.process_queue)]


def log_lines(app):
    return app.log_text.get("1.0", "end").rstrip("\n").splitlines()


def status(app):
    return app.status_label.cget("text")


# --- creation and widgets ------------------------------------------------------------

def test_app_created_without_visible_window(smoke):
    app = smoke.app
    assert isinstance(app, smoke.mod.tk.Tk)
    assert app.state() == "withdrawn"
    assert app.winfo_viewable() == 0


PAGES = ("home", "backup", "restore", "schedule", "tools", "logs", "settings", "help")


def test_show_page_switches_pages_and_sidebar(smoke):
    """U4-2: one page packed at a time, its sidebar entry selected, title from nav_<page>."""
    app = smoke.app
    assert tuple(app.pages) == PAGES and tuple(app.nav_items) == PAGES
    assert app.current_page == "home"
    try:
        for page in PAGES:
            app.nav_items[page].invoke()
            shown = [name for name, frame in app.pages.items() if frame.winfo_manager()]
            assert shown == [page]
            assert [n for n, item in app.nav_items.items() if item.selected] == [page]
            assert app.page_title_label.cget("text") == app.translator.get(f"nav_{page}")
            log_shown_in = app.log_card.pack_info()["in"] if app.log_card.winfo_manager() else None
            assert log_shown_in == (app.pages[page] if page in ("home", "logs") else None)
    finally:
        app.show_page("home")
    assert app.backup_btn.winfo_manager() == "grid"  # widgets survive page switches


def test_nav_status_chip_follows_progress(smoke):
    app = smoke.app
    app.progress_bar.config(maximum=200, value=84)
    assert app._nav_status_text() == app.translator.get("status_working").format(42)
    app.progress_bar.config(value=0)
    assert app._nav_status_text() == app.translator.get("status_idle")


def test_home_page_values_and_start_button(smoke, isolated_env):
    """U5-1: the home page shows the App's options and starts a backup like Ctrl+B.
    (The option variables save settings.json when written: temp %APPDATA%, put back.)"""
    app = smoke.app
    t = app.translator.get
    assert app.home_backup_btn.cget("text") == t("home_start_backup")
    assert app.home_backup_btn.cget("command") == app.start_backup_thread
    before = (app.retention_count_var.get(), app.hidden_mode_var.get())
    try:
        app.retention_count_var.set("5")
        assert app.home_retention_tile.get_value() == "5"
        app.hidden_mode_var.set("include")
        assert app.home_hidden_label.cget("text").endswith(t("home_included"))
    finally:
        app.retention_count_var.set(before[0])
        app.hidden_mode_var.set(before[1])
    rules = len(app.filters["include"]) + len(app.filters["exclude"])
    assert app.home_filter_chip.cget("text") == t("home_filters").format(rules)
    app.set_buttons_state("disabled")
    assert app.home_backup_btn.cget("state") == "disabled"
    app.set_buttons_state("normal")
    assert app.home_backup_btn.cget("state") == "normal"


def test_main_widgets_exist(smoke):
    app, tk, ttk = smoke.app, smoke.mod.tk, smoke.mod.ttk
    assert app.backup_btn.cget("text")
    assert app.restore_btn.cget("text")
    for name in ACTION_BUTTONS:
        assert isinstance(getattr(app, name), tk.Widget), name
    assert isinstance(app.user_combo, ttk.Combobox)
    assert isinstance(app.user_var, tk.StringVar)
    assert str(app.user_combo.cget("textvariable")) == str(app.user_var)
    assert isinstance(app.progress_bar, ttk.Progressbar)
    assert isinstance(app.status_label, tk.Label)
    assert isinstance(app.log_text, tk.Text)


def test_no_menu_bar_commands_on_settings_and_help_pages(smoke):
    """U5-7 (D1=B): no menu bar; its View/Language/Help commands are page buttons."""
    app = smoke.app
    assert str(app.cget("menu")) == ""
    t = app.translator.get
    assert app.theme_light_btn.cget("text") == t("theme_light")
    assert app.lang_ko_btn.cget("text") == "한국어"
    assert app.help_shortcuts_btn.cget("text") == t("help_shortcuts")
    assert app.sound_toggle.master.master is app.settings_cards["sound"].body


# --- queue messages (14 kinds) -------------------------------------------------------

def test_queue_kind_list_matches_process_queue(smoke):
    # P5-8b: process_queue hands each message to _handle_queue_message, which passes the
    # kinds it does not know on to _handle_queue_dialog_message; read all three.
    app_cls = smoke.mod.App
    source = "".join(inspect.getsource(getattr(app_cls, name)) for name in (
        "process_queue", "_handle_queue_message", "_handle_queue_dialog_message"))
    assert "self._handle_queue_message(task, value)" in source
    found = tuple(k for k in QUEUE_KINDS if f"task == '{k}'" in source)
    assert found == QUEUE_KINDS
    assert source.count("task == '") == len(QUEUE_KINDS)


def test_log(smoke):
    app = smoke.app
    app._last_ui_log_ts = 0.0
    run_one(smoke, "log", "smoke log line")
    assert log_lines(app)[-1].endswith("] smoke log line")
    assert str(app.log_text.cget("state")) == "disabled"


def test_log_throttled_within_interval(smoke):
    app = smoke.app
    app._last_ui_log_ts = 0.0
    app.message_queue.put(("log", "first"))
    app.message_queue.put(("log", "second"))
    app.process_queue()
    lines = log_lines(app)
    assert lines[-1].endswith("] first")  # second arrives within 0.08 s and is dropped
    assert not any(line.endswith("] second") for line in lines)


def test_update_progress_max(smoke):
    app = smoke.app
    run_one(smoke, "update_progress_max", 400)
    assert float(app.progress_bar.cget("maximum")) == 400
    assert status(app) == "Progress: 0%"


def test_update_progress(smoke):
    app = smoke.app
    app.progress_bar["maximum"] = 400
    run_one(smoke, "update_progress", 100)
    assert float(app.progress_bar.cget("value")) == 100
    assert status(app) == "Progress: 25%"


def test_update_status(smoke):
    run_one(smoke, "update_status", "Backup complete!")
    assert status(smoke.app) == "Backup complete!"


def test_enable_buttons(smoke, caplog):
    caplog.set_level(logging.DEBUG, logger="ezbak.gui.bridge")
    app = smoke.app
    app.set_buttons_state("disabled")
    run_one(smoke, "enable_buttons", None)
    for name in ACTION_BUTTONS:
        assert str(getattr(app, name).cget("state")) == "normal", name
    assert "Buttons enabled" in caplog.messages  # diagnostic log since P6-5c (was a DEBUG print)


def test_show_error(smoke):
    run_one(smoke, "show_error", "disk failed")
    assert smoke.dialogs.calls == [("showerror", smoke.app.translator.get("error"), "disk failed")]


def test_open_folder(smoke, tmp_path):
    run_one(smoke, "open_folder", str(tmp_path))  # startfile failure is swallowed
    assert smoke.blocked == [("os.startfile", str(tmp_path))]


def test_stop_progress(smoke):
    app = smoke.app
    app.progress_bar["mode"] = "indeterminate"
    run_one(smoke, "stop_progress", None)
    assert str(app.progress_bar.cget("mode")) == "determinate"


def test_reset_progress(smoke):
    app = smoke.app
    app.progress_bar["maximum"] = 400
    app.progress_bar["value"] = 50
    app.status_label.config(text="busy")
    run_one(smoke, "reset_progress", None)
    assert float(app.progress_bar.cget("value")) == 0
    assert float(app.progress_bar.cget("maximum")) == 100
    assert str(app.progress_bar.cget("mode")) == "determinate"
    assert status(app) == "Ready"


@pytest.fixture
def sync_thread(smoke, monkeypatch):
    """Run threads started by ezBAK at start() (only the module's own `threading` name)."""
    class SyncThread:
        def __init__(self, target=None, daemon=None, **kwargs):
            self._target = target

        def start(self):
            self._target()

    monkeypatch.setattr(smoke.mod, "threading", types.SimpleNamespace(Thread=SyncThread))


@pytest.mark.parametrize("enabled, expected", [(True, ["MB_ICONHAND"]), (False, [])])
def test_play_sound(smoke, sync_thread, monkeypatch, enabled, expected):
    import winsound
    beeps = []
    monkeypatch.setattr(winsound, "MessageBeep", beeps.append)
    smoke.app.sound_enabled_var.set(enabled)
    run_one(smoke, "play_sound", "error")
    assert beeps == [getattr(winsound, name) for name in expected]


@pytest.mark.parametrize("required, status_text, dialog", [
    (100, "Space check completed - OK", "showinfo"),
    (10_000, "Space check completed - Insufficient", "showwarning"),
])
def test_space_check_result(smoke, required, status_text, dialog):
    app = smoke.app
    app._last_ui_log_ts = 0.0
    result = {"required": required, "available": 1000, "file_count": 3, "sources_count": 2}
    run_one(smoke, "space_check_result", result)
    # the follow-up log/status messages are handled in the same process_queue call
    assert [c[0] for c in smoke.dialogs.calls] == [dialog]
    assert "Files: 3 items" in smoke.dialogs.calls[0][2]
    assert status(app) == status_text
    assert log_lines(app)[-1].split("] ", 1)[1].startswith("Space ")


@pytest.mark.parametrize("answer", [True, False])
def test_show_cross_user_warning(smoke, answer):
    smoke.dialogs.answer = answer
    run_one(smoke, "show_cross_user_warning", ("alice", "bob", "Backup_alice_1"))
    assert [c[0] for c in smoke.dialogs.calls] == ["askyesno"]
    assert "alice" in smoke.dialogs.calls[0][2] and "bob" in smoke.dialogs.calls[0][2]
    assert smoke.app._cross_user_confirmed is answer


@pytest.mark.parametrize("answer", [True, False])
def test_ask_yes_no(smoke, answer):
    smoke.dialogs.answer = answer
    run_one(smoke, "ask_yes_no", ("Title", "Question?"))
    assert smoke.dialogs.calls == [("askyesno", "Title", "Question?")]
    assert smoke.app._yes_no_answer is answer
    del smoke.app._yes_no_answer


@pytest.mark.parametrize("answer", [True, False])
def test_wait_yes_no_dialog_on_main_thread(smoke, monkeypatch, answer):
    # g-e: a worker's _wait_yes_no queues the question; process_queue on the main thread
    # shows the dialog and the worker's 0.1 s poll returns that answer.
    seen = []

    def askyesno(title, message, *args, **kwargs):
        seen.append((title, message, threading.current_thread() is threading.main_thread()))
        return answer

    monkeypatch.setattr(smoke.mod.Win11Dialog, "askyesno", askyesno)
    result = {}
    worker = threading.Thread(
        target=lambda: result.update(answer=smoke.app._wait_yes_no("Title", "Question?")))
    worker.start()
    deadline = time.monotonic() + 5
    while worker.is_alive() and time.monotonic() < deadline:
        smoke.app.process_queue()
        worker.join(0.05)
    assert not worker.is_alive()
    assert result == {"answer": answer}
    assert seen == [("Title", "Question?", True)]
    assert not hasattr(smoke.app, "_yes_no_answer")


def test_select_folder(smoke, monkeypatch):
    tk = smoke.mod.tk
    ok_text = smoke.app.translator.get("dialog_ok")
    seen = {}

    def fake_wait_window(self, window=None):
        """Instead of waiting for the user, press OK (first folder is preselected)."""
        listbox = [w for w in _descendants(self) if isinstance(w, tk.Listbox)][0]
        seen["items"] = list(listbox.get(0, "end"))
        button = [w for w in _descendants(self)
                  if _is_button(w, tk) and w.cget("text") == ok_text][0]
        button.invoke()

    monkeypatch.setattr(tk.Toplevel, "grab_set", lambda self: None)
    monkeypatch.setattr(tk.Toplevel, "wait_window", fake_wait_window)
    folders = ["Backup_a_2", "Backup_a_1"]
    run_one(smoke, "select_folder", folders)
    assert seen["items"] == folders
    assert smoke.app._folder_selection_result == "Backup_a_2"
    assert not [w for w in smoke.app.winfo_children() if isinstance(w, tk.Toplevel)]


def _is_button(widget, tk):
    """tk.Button or CTkButton (U2-2: dialogs may move to CustomTkinter, ADR 0004)."""
    if isinstance(widget, tk.Button):
        return True
    return any(klass.__name__ == "CTkButton" and klass.__module__.startswith("customtkinter")
               for klass in type(widget).__mro__)


def _descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from _descendants(child)
