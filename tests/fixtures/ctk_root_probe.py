"""Probe of a CustomTkinter root window, run in its own process (U2-3, ADR 0004).

tests/test_gui_ctk_root.py runs this with sys.executable: the test process already has
its one Tk root (the session App), and a second root in-process is what tests/CLAUDE.md
forbids. argv[1] is "off" (deactivate_automatic_dpi_awareness before the root, decision
D8) or "default". Prints one line `PROBE <json>` with what App.__init__ relies on when
its base class becomes ctk.CTk: withdraw before the window exists, configure(bg=)
versus configure(fg_color=),
config(menu=), geometry, iconbitmap failure, ttk style, Tk variables, __class__ switch,
and the process DPI awareness.
"""
import json
import sys
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk


def dpi_awareness():
    """0 unaware, 1 system aware, 2 per-monitor aware; None off Windows / old Windows."""
    if not sys.platform.startswith("win"):
        return None
    import ctypes
    user32 = ctypes.windll.user32
    try:
        user32.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        user32.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        return user32.GetAwarenessFromDpiAwarenessContext(user32.GetThreadDpiAwarenessContext())
    except AttributeError:
        return None


class FirstMixin:
    def start_backup_thread(self):
        return "first"


class SecondMixin:
    def start_backup_thread(self):
        return "second"


class Root(FirstMixin, ctk.CTk):
    def __init__(self):
        super().__init__()
        self.withdraw()  # App.__init__: hide until fully initialized


class OtherRoot(SecondMixin, ctk.CTk):
    """Stands in for the nopywin32 module's App (gui_app switches __class__)."""


def probe(dpi_mode):
    if dpi_mode == "off":
        ctk.deactivate_automatic_dpi_awareness()
    out = {"ctk_version": ctk.__version__}
    root = Root()
    out["state_after_init"] = root.state()
    root.title("ezBAK probe")
    try:
        root.iconbitmap("./no/such/icon.ico")
        out["iconbitmap_missing"] = "no error"
    except tk.TclError:
        out["iconbitmap_missing"] = "TclError"
    root.geometry("800x600")  # smaller than any runner screen
    root.configure(bg="#202020")  # what App._init_window does today
    out["bg_after_configure_bg"] = str(root.cget("bg"))
    root.configure(fg_color="#202020")  # the CTk way
    out["bg_after_configure_fg_color"] = str(root.cget("bg"))
    out["fg_color"] = root.cget("fg_color")
    style = ttk.Style(root)
    style.theme_use("clam")
    out["ttk_theme"] = style.theme_use()
    var = tk.StringVar(value="exclude")
    out["var"] = var.get()
    menubar = tk.Menu(root)
    menubar.add_cascade(label="Help", menu=tk.Menu(menubar, tearoff=0))
    root.config(menu=menubar)
    out["menu_via_config"] = str(root.cget("menu")) == str(menubar)
    other = tk.Menu(root)
    root.configure(menu=other)
    out["menu_via_configure"] = str(root.cget("menu")) == str(other)
    root.config(menu=menubar)
    root.update()
    out["state_after_update"] = root.state()
    out["window_scaling"] = ctk.ScalingTracker.get_window_scaling(root)
    out["widget_scaling"] = ctk.ScalingTracker.get_widget_scaling(root)
    out["dpi_awareness"] = dpi_awareness()
    root.__class__ = OtherRoot
    out["class_switch"] = root.start_backup_thread()
    root.__class__ = Root
    out["class_switch_back"] = root.start_backup_thread()
    root.deiconify()
    root.update()
    out["state_after_deiconify"] = root.state()
    root.config(menu="")  # a Windows menu bar takes its height out of the client area
    root.geometry("800x600")
    root.update()
    out["geometry_size"] = root.geometry().split("+")[0]  # read while mapped
    root.destroy()
    return out


if __name__ == "__main__":
    result = probe(sys.argv[1])
    print("PROBE " + json.dumps(result, sort_keys=True), flush=True)
