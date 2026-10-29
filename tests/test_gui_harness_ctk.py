"""gui_harness on CustomTkinter widgets (U2-2, ADR 0004).

A small CTk tree is built under the session App and read back the way the window and
dialog goldens read widgets: CTk options through cget, the drawing parts (canvas,
labels) left out, children the code adds kept in creation order, CTk methods not taken
for the App's own methods, CTk buttons found by text. The widgets are destroyed again.
"""
import importlib.util
from pathlib import Path

import pytest

pytestmark = pytest.mark.gui

ctk = pytest.importorskip("customtkinter")

FIXTURES = Path(__file__).resolve().parent / "fixtures"
_spec = importlib.util.spec_from_file_location("ezbak_test_gui_harness_ctk",
                                               FIXTURES / "gui_harness.py")
gh = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gh)


def on_save():
    """Stands in for an App method given as a CTk command."""


class Owner:
    """Holds widget attributes like App does, so Namer gives them names."""


@pytest.fixture
def ctk_tree(ezbak_mod, gui_app):
    app = gui_app(ezbak_mod)
    tk = ezbak_mod.tk
    owner = Owner()
    # A CTk widget whose master is a plain tk widget replaces the master's config and
    # configure on the instance: a throwaway tk.Frame keeps them off the session App.
    holder = tk.Frame(app)
    owner.card = ctk.CTkFrame(holder, corner_radius=12, fg_color=("#FFFFFF", "#2B2B2B"))
    owner.save_btn = ctk.CTkButton(owner.card, text="Save", command=on_save, width=120)
    owner.note = tk.Label(owner.card, text="plain tk child")
    owner.save_btn.pack(side="left", padx=8)
    owner.note.pack(side="left")
    try:
        yield owner
    finally:
        holder.destroy()
    assert "configure" not in vars(app) and "config" not in vars(app)


def test_ctk_node_options_and_children(ctk_tree):
    node = gh.widget_node(ctk_tree.card, gh.Namer(ctk_tree))
    assert node["py"] == "CTkFrame"
    assert node["attr"] == "card"
    assert node["options"]["corner_radius"] == 12
    assert node["options"]["fg_color"] == ["#FFFFFF", "#2B2B2B"]
    assert node["pack_propagate"] is True
    assert node["bindings"] == {"<Configure>": "_update_dimensions_event"}
    assert "grid_propagate" not in node  # the frame's canvas is placed, not a slave
    assert [child["py"] for child in node["children"]] == ["CTkButton", "Label"]
    button = node["children"][0]
    assert button["attr"] == "save_btn"
    assert button["options"]["text"] == "Save"
    assert button["options"]["command"] == "on_save"
    assert button["options"]["width"] == 120
    assert len(button["options"]["font"]) == 3  # [family, size, weight]
    assert "children" not in button  # canvas and text label are drawing parts
    assert "grid_propagate" not in button
    assert button["layout"]["side"] == "left"
    assert node["children"][1]["options"]["text"] == "plain tk child"


def test_ctk_snapshot_is_stable(ctk_tree):
    first = gh.snapshot_tree(ctk_tree.card, gh.Namer(ctk_tree))
    assert gh.snapshot_tree(ctk_tree.card, gh.Namer(ctk_tree)) == first


def test_widget_text_and_button_lookup(ctk_tree):
    assert gh.widget_text(ctk_tree.save_btn) == "Save"
    assert gh.widget_text(ctk_tree.note) == "plain tk child"
    assert gh.widget_text(ctk_tree.card) is None
    assert gh.button(ctk_tree.card, "Save") is ctk_tree.save_btn
    assert gh.is_ctk(ctk_tree.save_btn) and not gh.is_ctk(ctk_tree.note)


def test_own_members_leave_out_customtkinter():
    class Mixin:
        def start_backup_thread(self):
            pass

    members = gh.own_members(type("Host", (Mixin, ctk.CTkFrame), {}))
    assert "start_backup_thread" in members
    assert "configure" not in members and "cget" not in members and "_draw" not in members
