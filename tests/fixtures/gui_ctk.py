"""CustomTkinter helpers of gui_harness.py (U2-2, ADR 0004).

A CTk widget's configure() returns nothing, so its snapshot options are read with cget
(CTK_OPTIONS). The canvas, labels and frames it keeps in its own private attributes are
drawing parts: they are not recorded and do not count as children or slaves.

Loaded by gui_harness.py with importlib (tests/fixtures is not a package).
"""
import tkinter as tk

try:
    import customtkinter as ctk
except ImportError:  # not installed: no CTk widget can exist either
    ctk = None

# Options read with cget from a CTk widget (those the widget does not support are
# skipped). Colours stay as given: a (light, dark) pair or one colour.
CTK_OPTIONS = ("text", "fg_color", "hover_color", "text_color", "border_color",
               "border_width", "corner_radius", "width", "height", "state", "command",
               "font", "anchor", "compound")


def is_ctk_class(klass):
    """True for a class defined in customtkinter."""
    return klass.__module__.split(".")[0] == "customtkinter"


def is_ctk(widget):
    """True for a CustomTkinter widget or window (a class in its MRO is from customtkinter)."""
    return any(is_ctk_class(klass) for klass in type(widget).__mro__)


def ctk_parts(widget):
    """Path names of the drawing parts of a CTk widget: the Tk widgets it keeps in its own
    private attributes (_canvas, _text_label, _label, ...). Public attributes are children
    our own components create on purpose (Card.body, SidebarItem.text_label) and stay."""
    if not is_ctk(widget):
        return set()
    return {str(value) for name, value in vars(widget).items()
            if name.startswith("_") and isinstance(value, tk.Misc)}


def ctk_options(widget, namer, plain):
    """{option: value} of CTK_OPTIONS the CTk widget supports; command -> function name,
    variable -> owner attribute name, CTkFont -> [family, size, weight]."""
    out = {}
    for name in CTK_OPTIONS:
        try:
            value = widget.cget(name)
        except (ValueError, AttributeError, tk.TclError):
            continue
        if name == "command":
            value = getattr(value, "__name__", None) if value is not None else None
        elif ctk is not None and isinstance(value, ctk.CTkFont):
            value = [value.cget("family"), value.cget("size"), value.cget("weight")]
        out[name] = plain(value)
    for name in ("variable", "textvariable"):
        try:
            value = widget.cget(name)
        except (ValueError, AttributeError, tk.TclError):
            continue
        if isinstance(value, tk.Variable):
            out[name] = namer.variable(str(value))
    return out


def widget_text(widget):
    """The text option of a tk or CTk widget, or None when it has none."""
    if is_ctk(widget):
        try:
            return widget.cget("text")
        except (ValueError, AttributeError):
            return None
    return widget.cget("text") if "text" in widget.keys() else None


def is_button(widget):
    """tk.Button or CTkButton."""
    return isinstance(widget, tk.Button) or (ctk is not None and isinstance(widget, ctk.CTkButton))
