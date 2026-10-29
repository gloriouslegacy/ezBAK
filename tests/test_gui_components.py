"""CustomTkinter shared widgets of ezbak.gui (U3-3, ui-plan §4): Card, StatTile, Chip,
SidebarItem, RingButton.

Built under the session App inside a throwaway tk.Frame (a CTk widget replaces the
config/configure of a plain tk master, see test_gui_harness_ctk.py) and destroyed again.
Colours are compared with the theme token values (light, dark).
"""
import importlib.util
import tkinter
from pathlib import Path

import pytest

pytestmark = pytest.mark.gui

pytest.importorskip("customtkinter")

from ezbak.gui import Card, Chip, RingButton, SidebarItem, Sparkline, StatTile  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
_spec = importlib.util.spec_from_file_location("ezbak_test_gui_harness_components",
                                               FIXTURES / "gui_harness.py")
gh = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gh)

CARD = ["#FFFFFF", "#2B2B2B"]
ACCENT_SOFT = ["#E1F1FF", "#003D71"]


@pytest.fixture
def holder(ezbak_mod, gui_app):
    app = gui_app(ezbak_mod)
    frame = ezbak_mod.tk.Frame(app)
    try:
        yield frame
    finally:
        frame.destroy()
    assert "configure" not in vars(app) and "config" not in vars(app)


class Calls:
    def __init__(self):
        self.count = 0

    def __call__(self):
        self.count += 1
        return "ran"


def ring_items(ring):
    canvas = ring._canvas
    return [(canvas.type(item), canvas.itemcget(item, "text") if canvas.type(item) == "text" else None)
            for item in canvas.find_withtag("ring")]


def test_card_with_and_without_title(holder):
    card = Card(holder, title="System")
    assert gh.plain(card.cget("fg_color")) == CARD
    assert card.cget("corner_radius") == 12
    assert card.cget("border_width") == 1
    assert card.title_label.cget("text") == "System"
    node = gh.widget_node(card, gh.Namer())
    assert [child["py"] for child in node["children"]] == ["CTkLabel", "CTkFrame"]
    plain = Card(holder)
    assert plain.title_label is None
    assert plain.winfo_children() == [plain.body]


def test_stat_tile_value(holder):
    tile = StatTile(holder, caption="Backups to keep", value="2")
    assert tile.caption_label.cget("text") == "Backups to keep"
    assert tile.get_value() == "2"
    tile.set_value("5")
    assert tile.get_value() == "5"


def test_chip_is_a_pill_button(holder):
    calls = Calls()
    chip = Chip(holder, text="3 filters", command=calls)
    assert chip.cget("height") == 28
    assert chip.cget("corner_radius") == 14
    assert gh.plain(chip.cget("fg_color")) == ACCENT_SOFT
    chip.invoke()
    assert calls.count == 1
    assert gh.button(holder, "3 filters") is chip


def test_sidebar_item_select_invoke_and_state(holder):
    calls = Calls()
    item = SidebarItem(holder, icon="", text="Home", command=calls)
    assert item.cget("fg_color") == "transparent"
    assert item.text_label.cget("text") == "Home"
    assert item.icon_label.cget("text") == ""
    item.set_selected(True)
    assert gh.plain(item.cget("fg_color")) == ACCENT_SOFT
    item.set_selected(False)
    assert item.cget("fg_color") == "transparent"
    assert item.invoke() == "ran"
    item._on_click()
    assert calls.count == 2
    item.configure(state="disabled")
    assert item.invoke() is None and calls.count == 2
    item.configure(state="normal", text="Start")
    assert item.cget("text") == "Start" and item.text_label.cget("text") == "Start"
    assert gh.widget_text(item) == "Start"


def test_ring_button_idle_progress_and_disabled(holder):
    calls = Calls()
    ring = RingButton(holder, text="Start backup", command=calls)
    assert ring.cget("width") == 120 and ring.cget("height") == 120
    assert ring_items(ring) == [("oval", None), ("text", "Start backup")]
    ring.set_progress(0.42)
    assert ring.get_progress() == 0.42
    assert ring_items(ring) == [("oval", None), ("arc", None), ("text", "42%")]
    arc = ring._canvas.find_withtag("ring")[1]
    assert float(ring._canvas.itemcget(arc, "extent")) == pytest.approx(-359.9 * 0.42, abs=0.01)
    ring.set_progress(0)
    assert ring_items(ring) == [("oval", None), ("text", "0%")]
    ring.set_progress(2)
    assert ring.get_progress() == 1.0
    ring.set_progress(None)
    assert ring_items(ring) == [("oval", None), ("text", "Start backup")]
    assert ring.invoke() == "ran"
    ring.configure(state="disabled")
    assert ring.invoke() is None and calls.count == 1


def test_ring_button_redraws_after_ctk_redraw(holder):
    """CTk redraws a frame (scaling, appearance mode) through _draw: the ring comes back."""
    ring = RingButton(holder, text="Go")
    ring.set_progress(0.5)
    ring._draw()
    assert ring_items(ring) == [("oval", None), ("arc", None), ("text", "50%")]


def test_snapshot_of_button_like_frames(holder):
    """The harness records text/command/state of the button-like frames, not the canvas."""
    def start_backup_thread():
        pass

    ring = RingButton(holder, text="Start", command=start_backup_thread)
    item = SidebarItem(holder, icon="", text="Backup")
    ring_node = gh.widget_node(ring, gh.Namer())
    assert ring_node["py"] == "RingButton"
    assert ring_node["options"]["text"] == "Start"
    assert ring_node["options"]["command"] == "start_backup_thread"
    assert ring_node["options"]["state"] == "normal"
    assert "children" not in ring_node
    item_node = gh.widget_node(item, gh.Namer())
    assert [child["py"] for child in item_node["children"]] == ["CTkLabel", "CTkLabel"]


def spark_lines(spark):
    # CTkCanvas.coords is a setter for its own shapes (returns None): read with tkinter's
    canvas = spark._canvas
    return [tkinter.Canvas.coords(canvas, item) for item in canvas.find_withtag("spark")]


def test_sparkline_draws_base_and_scaled_line(holder):
    spark = Sparkline(holder, width=100, height=50, samples=5)
    scale = spark._apply_widget_scaling(1)
    (base,) = spark_lines(spark)  # empty: only the base line
    assert base[1] == base[3] == pytest.approx((50 - 4) * scale)
    spark.push(10)
    assert len(spark_lines(spark)) == 1  # one sample is not a line yet
    spark.push(20)
    spark.push(-5)  # negative counts as 0
    assert spark.values() == [10.0, 20.0, 0.0]
    base, line = spark_lines(spark)
    xs, ys = line[0::2], line[1::2]
    step = 100 * scale / 4
    assert xs == pytest.approx([2 * step, 3 * step, 4 * step])  # newest at the right edge
    top, bottom = 4 * scale, (50 - 4) * scale
    assert ys == pytest.approx([bottom - (bottom - top) / 2, top, bottom])  # scaled to the max


def test_sparkline_keeps_the_last_samples_and_clears(holder):
    spark = Sparkline(holder, samples=3)
    for value in range(6):
        spark.push(value)
    assert spark.values() == [3.0, 4.0, 5.0]
    spark.clear()
    assert spark.values() == [] and len(spark_lines(spark)) == 1


def test_sparkline_redraws_with_the_frame(holder):
    spark = Sparkline(holder, samples=4)
    spark.push(1)
    spark.push(2)
    spark._draw()  # CTk redraw (scaling, appearance mode) deletes and redraws its canvas
    assert len(spark_lines(spark)) == 2
