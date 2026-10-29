"""U8: DPI factor helpers of ezbak.gui (dpi.py, ui-plan §9).

CTk's ScalingTracker is patched to a fixed monitor factor, so the results do not depend
on the machine's display. No Tk root is needed: the widget is a stand-in.
"""
import pytest

pytestmark = pytest.mark.gui

ctk = pytest.importorskip("customtkinter")

from ezbak.gui import dpi  # noqa: E402


class FakeTk:
    def __init__(self):
        self.calls = []

    def call(self, *args):
        self.calls.append(args)


class FakeWidget:
    def __init__(self, screen=(1920, 1080)):
        self.screen = screen
        self.tk = FakeTk()

    def _root(self):
        return self

    def winfo_screenwidth(self):
        return self.screen[0]

    def winfo_screenheight(self):
        return self.screen[1]


@pytest.fixture(autouse=True)
def keep_font_factor(monkeypatch):
    """sync_tk_scaling stores the factor in the module: restore it for the session App."""
    monkeypatch.setattr(dpi, "_font_factor", dpi._font_factor)


@pytest.fixture
def monitor(monkeypatch):
    def set_factor(value):
        seen = []

        def fake(window):
            seen.append(window)
            if isinstance(value, Exception):
                raise value
            return value
        monkeypatch.setattr(ctk.ScalingTracker, "get_window_dpi_scaling", staticmethod(fake))
        return seen
    return set_factor


def test_factor_reads_the_root_window(monitor):
    widget = FakeWidget()
    seen = monitor(1.25)
    assert dpi.factor(widget) == 1.25
    assert seen == [widget]


def test_factor_is_one_when_ctk_fails(monitor):
    monitor(OSError("GetDpiForMonitor"))
    assert dpi.factor(FakeWidget()) == 1.0


@pytest.mark.parametrize("scale, expected", [(1, 449), (1.25, 561), (1.5, 674), (2, 898)])
def test_px(monitor, scale, expected):
    monitor(scale)
    assert dpi.px(FakeWidget(), 449) == expected


def test_px_keeps_the_value_at_one(monitor):
    monitor(1.0)
    assert dpi.px(FakeWidget(), 360) == 360
    assert isinstance(dpi.px(FakeWidget(), 360), int)


@pytest.mark.parametrize("scale, expected", [(1, "750x520"), (1.5, "1125x780")])
def test_geometry_without_position(monitor, scale, expected):
    monitor(scale)
    assert dpi.geometry(FakeWidget(), 750, 520) == expected


def test_geometry_keeps_the_position(monitor):
    monitor(1.5)
    assert dpi.geometry(FakeWidget(), 546, 380, 70, 80) == "819x570+70+80"


@pytest.mark.parametrize("scale, expected", [(1, "450x340+735+370"), (1.5, "675x510+623+285")])
def test_centered_geometry(monitor, scale, expected):
    monitor(scale)
    assert dpi.centered_geometry(FakeWidget(), 450, 340) == expected


def test_sync_tk_scaling_leaves_tk_alone_at_one(monitor):
    root = FakeWidget()
    monitor(1)
    assert dpi.sync_tk_scaling(root) == 1
    assert root.tk.calls == []


def test_sync_tk_scaling_sets_points_to_pixels(monitor):
    root = FakeWidget()
    monitor(1.5)
    assert dpi.sync_tk_scaling(root) == 1.5
    assert root.tk.calls == [("tk", "scaling", 2.0)]


def test_font_px_follows_the_synced_root(monitor):
    monitor(1)
    dpi.sync_tk_scaling(FakeWidget())
    assert dpi.font_px(14) == 14
    monitor(1.5)
    dpi.sync_tk_scaling(FakeWidget())
    assert dpi.font_px(14) == 21
    monitor(1)
    dpi.sync_tk_scaling(FakeWidget())
    assert dpi.font_px(14) == 14
