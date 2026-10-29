"""U2-3: a ctk.CTk root does what App.__init__ needs (ADR 0004, ui-plan U2-3·D8).

tests/fixtures/ctk_root_probe.py runs in a child Python process (one Tk root per
process: the session App already holds this one). Kept as a regression test for U4-1,
where App's base class becomes ctk.CTk.
"""
import json
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.gui

pytest.importorskip("customtkinter")

PROBE = Path(__file__).resolve().parent / "fixtures" / "ctk_root_probe.py"


def run_probe(dpi_mode):
    proc = subprocess.run([sys.executable, str(PROBE), dpi_mode], capture_output=True,
                          text=True, timeout=120)
    assert proc.returncode == 0, proc.stderr
    lines = proc.stdout.splitlines()
    probe_lines = [line for line in lines if line.startswith("PROBE ")]
    assert len(probe_lines) == 1, proc.stdout
    others = [line for line in lines if not line.startswith("PROBE ")]
    return json.loads(probe_lines[0][len("PROBE "):]), others, proc.stderr


@pytest.fixture(scope="module")
def probe_off():
    return run_probe("off")


def test_ctk_root_supports_app_init(probe_off):
    result, stray_output, _ = probe_off
    assert stray_output == []  # CTk prints its warnings (unknown option, titlebar error)
    assert result["ctk_version"] == "5.2.2"
    assert result["state_after_init"] == "withdrawn"
    assert result["state_after_update"] == "withdrawn"  # withdraw before the window exists holds
    assert result["state_after_deiconify"] == "normal"
    assert result["iconbitmap_missing"] == "TclError"  # App falls back to the second icon
    # CTk.configure drops bg silently (not one of its tk options): App._init_window's
    # configure(bg=) must become configure(fg_color=) in U4-1.
    assert result["bg_after_configure_bg"] != "#202020"
    assert result["bg_after_configure_fg_color"] == "#202020"
    assert result["fg_color"] == "#202020"
    assert result["ttk_theme"] == "clam"
    assert result["var"] == "exclude"
    assert result["menu_via_config"] is True
    assert result["menu_via_configure"] is True
    assert result["class_switch"] == "second"
    assert result["class_switch_back"] == "first"


def test_dpi_awareness_off_keeps_the_process_unaware(probe_off):
    """D8 switch (App used it until U8-2): the process stays DPI unaware, CTk scaling is 1."""
    result, _, _ = probe_off
    assert result["window_scaling"] == 1
    assert result["widget_scaling"] == 1
    assert result["geometry_size"] == "800x600"
    if sys.platform.startswith("win") and result["dpi_awareness"] is not None:
        assert result["dpi_awareness"] == 0


@pytest.mark.skipif(not sys.platform.startswith("win"), reason="DPI awareness is a Windows setting")
def test_dpi_awareness_default_makes_the_process_aware():
    """Without D8's switch CTk makes the process DPI aware (App from U8-3)."""
    result, _, _ = run_probe("default")
    if result["dpi_awareness"] is None:
        pytest.skip("GetThreadDpiAwarenessContext not available")
    assert result["dpi_awareness"] != 0
