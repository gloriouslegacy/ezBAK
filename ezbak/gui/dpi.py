"""DPI scaling of the plain tk parts of the window (U8, ui-plan §5·§9, decision D8).

CTk widgets, CTk.geometry() and CTk.minsize() scale themselves. The tk.Toplevel dialogs,
wraplength values, the sound ToggleSwitch canvas and the Treeview row height are plain
pixels, and Tk fixes its point-to-pixel scale (`tk scaling`) before CTk makes the process
DPI aware. These helpers multiply such pixel values by the DPI factor of the App root.

The factor is 1 while CTk's automatic DPI awareness is off, off Windows, and at 100 %
(96 DPI), so then every helper returns its input unchanged.
"""
import logging

import customtkinter as ctk

_log = logging.getLogger(__name__)

POINTS_PER_INCH = 72
DPI_100 = 96

# Factor of the App root as set by sync_tk_scaling (1 until then): tk_font sizes are
# pixels (negative), which `tk scaling` does not touch, so they take it from here.
_font_factor = 1.0


def factor(widget):
    """DPI factor of the monitor the App root (widget's Tk root) is on: 1.0, 1.25, 1.5 ..."""
    try:
        return float(ctk.ScalingTracker.get_window_dpi_scaling(widget._root()))
    except Exception:
        _log.debug("event=dpi_factor failed", exc_info=True)
        return 1.0


def px(widget, value):
    """`value` pixels at 100 % as pixels at the widget's DPI factor (int, unchanged at 1)."""
    scale = factor(widget)
    return value if scale == 1 else round(value * scale)


def geometry(widget, width, height, x=None, y=None):
    """Tk geometry string with width and height scaled; x and y are screen pixels as given."""
    spec = f"{px(widget, width)}x{px(widget, height)}"
    return spec if x is None else f"{spec}+{x}+{y}"


def centered_geometry(widget, width, height):
    """Scaled geometry centred on the screen (the dialogs' `screenwidth // 2 - width // 2`)."""
    width, height = px(widget, width), px(widget, height)
    x = (widget.winfo_screenwidth() // 2) - (width // 2)
    y = (widget.winfo_screenheight() // 2) - (height // 2)
    return f"{width}x{height}+{x}+{y}"


def font_px(size):
    """Pixel font size at 100 % as pixels at the App root's factor (see sync_tk_scaling)."""
    return size if _font_factor == 1 else round(size * _font_factor)


def sync_tk_scaling(root):
    """Set Tk's point-to-pixel scale to the DPI factor (tk pt fonts follow the monitor)
    and remember the factor for the pixel fonts of font_px.

    Only when the factor is not 1: at 100 % Tk keeps the value it chose itself.
    """
    global _font_factor
    scale = factor(root)
    _font_factor = scale
    if scale != 1:
        root.tk.call("tk", "scaling", scale * DPI_100 / POINTS_PER_INCH)
    return scale
