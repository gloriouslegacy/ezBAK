"""Shared CustomTkinter widgets of the redesigned window (U3-3, ui-plan §4, ADR 0004).

Card, StatTile, Chip, SidebarItem, RingButton and Sparkline (U7-4) take their colours, radii and fonts
from theme.TOKENS / RADIUS / font_spec, so a page only places them. Colours are
(light, dark) pairs; CTk picks one from its appearance mode. Nothing uses these yet
(the window switches to CTk in U4).

SidebarItem and RingButton are CTk frames that act like buttons: invoke() runs the
command unless the state is "disabled", and configure/cget accept text, command and
state like a CTkButton (RingButton draws on the frame's own canvas, so its drawing is
redone whenever CTk redraws the frame, e.g. on an appearance mode change).
"""
from collections import deque

import customtkinter as ctk

from .theme import (ICON_FONT_FAMILY, PAD_CARD, RADIUS, RING_DIAMETER, TOKENS,
                    font_spec)

_RING_TAG = "ring"
_LINE_TAG = "spark"
_LINE_WIDTH = 2
_LINE_PAD = 4
_RING_WIDTH = 10
_RING_PAD = 6


class Card(ctk.CTkFrame):
    """Rounded surface with an optional section title; children go into `body`."""

    def __init__(self, master, title=None, language="en", **kwargs):
        kwargs.setdefault("fg_color", TOKENS["card"])
        kwargs.setdefault("corner_radius", RADIUS["card"])
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", TOKENS["border"])
        super().__init__(master, **kwargs)
        self.title_label = None
        if title is not None:
            self.title_label = ctk.CTkLabel(self, text=title, anchor="w",
                                            font=font_spec("section", language),
                                            text_color=TOKENS["fg"])
            self.title_label.pack(fill="x", padx=PAD_CARD, pady=(PAD_CARD, 8))
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=PAD_CARD,
                       pady=(0 if title is not None else PAD_CARD, PAD_CARD))


class StatTile(Card):
    """Card with a muted caption above a large value (home metrics)."""

    def __init__(self, master, caption, value="", language="en", **kwargs):
        super().__init__(master, language=language, **kwargs)
        self.caption_label = ctk.CTkLabel(self.body, text=caption, anchor="w",
                                          font=font_spec("caption", language),
                                          text_color=TOKENS["fg_muted"])
        self.caption_label.pack(fill="x")
        self.value_label = ctk.CTkLabel(self.body, text=value, anchor="w",
                                        font=font_spec("title", language),
                                        text_color=TOKENS["fg"])
        self.value_label.pack(fill="x")

    def set_value(self, text):
        self.value_label.configure(text=text)

    def get_value(self):
        return self.value_label.cget("text")


class Chip(ctk.CTkButton):
    """Small pill button (e.g. "3 filters" on the home card)."""

    HEIGHT = 28

    def __init__(self, master, text, command=None, language="en", **kwargs):
        kwargs.setdefault("height", self.HEIGHT)
        kwargs.setdefault("width", 56)  # grows with the text
        kwargs.setdefault("corner_radius", self.HEIGHT // 2)
        kwargs.setdefault("fg_color", TOKENS["accent_soft"])
        kwargs.setdefault("hover_color", TOKENS["card_hover"])
        kwargs.setdefault("text_color", TOKENS["accent"])
        kwargs.setdefault("font", font_spec("caption", language))
        super().__init__(master, text=text, command=command, **kwargs)


class _Clickable:
    """configure/cget/invoke of text, command and state for the button-like frames.
    The subclass sets _text, _command, _state before CTkFrame.__init__ and implements
    _apply_look()."""

    _OWN_OPTIONS = ("text", "command", "state")

    def configure(self, require_redraw=False, **kwargs):
        changed = False
        for name in self._OWN_OPTIONS:
            if name in kwargs:
                setattr(self, "_" + name, kwargs.pop(name))
                changed = True
        super().configure(require_redraw=require_redraw, **kwargs)
        if changed:
            self._apply_look()

    def cget(self, attribute_name):
        if attribute_name in self._OWN_OPTIONS:
            return getattr(self, "_" + attribute_name)
        return super().cget(attribute_name)

    def invoke(self):
        """Run the command as a click would (nothing when disabled or without command)."""
        if self._state != "disabled" and self._command is not None:
            return self._command()
        return None

    def _on_click(self, event=None):
        self.invoke()


class SidebarItem(_Clickable, ctk.CTkFrame):
    """Sidebar menu entry: Segoe MDL2 glyph + label; selected entries use accent_soft."""

    HEIGHT = 40

    def __init__(self, master, icon, text, command=None, language="en", **kwargs):
        self._text, self._command, self._state = text, command, "normal"
        self.selected = False
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("corner_radius", RADIUS["sidebar_item"])
        kwargs.setdefault("height", self.HEIGHT)
        super().__init__(master, **kwargs)
        self.icon_label = ctk.CTkLabel(self, text=icon, width=24,
                                       font=(ICON_FONT_FAMILY, 16, "normal"))
        self.icon_label.pack(side="left", padx=(12, 8), pady=8)
        self.text_label = ctk.CTkLabel(self, text=text, anchor="w",
                                       font=font_spec("body", language))
        self.text_label.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=8)
        for widget in (self, self.icon_label, self.text_label):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
        self._apply_look()

    def set_selected(self, selected):
        self.selected = bool(selected)
        self._apply_look()

    def _apply_look(self, hover=False):
        if self.selected:
            fg, text = TOKENS["accent_soft"], TOKENS["accent"]
        else:
            fg = TOKENS["card_hover"] if hover else "transparent"
            text = TOKENS["fg_muted"] if self._state == "disabled" else TOKENS["fg"]
        ctk.CTkFrame.configure(self, fg_color=fg)
        self.icon_label.configure(text_color=text)
        self.text_label.configure(text=self._text, text_color=text)

    def _on_enter(self, event=None):
        if not self.selected and self._state != "disabled":
            self._apply_look(hover=True)

    def _on_leave(self, event=None):
        self._apply_look()


class RingButton(_Clickable, ctk.CTkFrame):
    """Round start button (home). Idle: a filled accent circle with the text. While a job
    runs, set_progress(0..1) draws a track ring, the done arc and the percentage;
    set_progress(None) goes back to idle."""

    def __init__(self, master, text, command=None, diameter=RING_DIAMETER, language="en",
                 **kwargs):
        self._text, self._command, self._state = text, command, "normal"
        self._progress = None
        self._language = language
        self._diameter = diameter
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("corner_radius", 0)
        super().__init__(master, width=diameter, height=diameter, **kwargs)
        self.bind("<Button-1>", self._on_click)

    def set_progress(self, fraction):
        self._progress = None if fraction is None else max(0.0, min(1.0, float(fraction)))
        self._apply_look()

    def get_progress(self):
        return self._progress

    def _apply_look(self):
        self._draw_ring()

    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates)
        self._draw_ring()

    def _draw_ring(self):
        canvas = getattr(self, "_canvas", None)
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete(_RING_TAG)
        size = self._apply_widget_scaling(self._diameter)
        pad, width = self._apply_widget_scaling(_RING_PAD), self._apply_widget_scaling(_RING_WIDTH)
        box = (pad, pad, size - pad, size - pad)
        color = self._apply_appearance_mode
        if self._progress is None:
            fill = TOKENS["track"] if self._state == "disabled" else TOKENS["accent"]
            canvas.create_oval(*box, fill=color(fill), outline="", tags=_RING_TAG)
            self._draw_label(self._text, "#FFFFFF", "body_bold")
            return
        canvas.create_oval(*box, outline=color(TOKENS["track"]), width=width, tags=_RING_TAG)
        if self._progress > 0:
            canvas.create_arc(*box, start=90, extent=-359.9 * self._progress, style="arc",
                              outline=color(TOKENS["accent"]), width=width, tags=_RING_TAG)
        self._draw_label(f"{int(self._progress * 100)}%", color(TOKENS["fg"]), "display")

    def _draw_label(self, text, fill, role):
        family, size, weight = font_spec(role, self._language)
        center = self._apply_widget_scaling(self._diameter) / 2
        self._canvas.create_text(center, center, text=text, fill=fill,
                                 font=(family, -round(self._apply_widget_scaling(size)), weight),
                                 tags=_RING_TAG)


class Sparkline(ctk.CTkFrame):
    """Small line chart of the last `samples` values (home copy speed, U7-4).

    push(value) adds a sample (negative counts as 0) and redraws; clear() empties it.
    The line is scaled to the largest sample shown, on a muted base line, and is drawn on
    the frame's own canvas like RingButton (redrawn when CTk redraws the frame).
    """

    SAMPLES = 60  # 30 s at the 0.5 s status timer (decision D4 A)

    def __init__(self, master, width=280, height=56, samples=SAMPLES, **kwargs):
        self._values = deque(maxlen=samples)
        self._samples = samples
        self._size = (width, height)
        kwargs.setdefault("fg_color", "transparent")
        kwargs.setdefault("corner_radius", 0)
        super().__init__(master, width=width, height=height, **kwargs)

    def push(self, value):
        self._values.append(max(0.0, float(value)))
        self._draw_line()

    def clear(self):
        self._values.clear()
        self._draw_line()

    def values(self):
        return list(self._values)

    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates)
        self._draw_line()

    def _draw_line(self):
        canvas = getattr(self, "_canvas", None)
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete(_LINE_TAG)
        width, height = (self._apply_widget_scaling(v) for v in self._size)
        pad = self._apply_widget_scaling(_LINE_PAD)
        color = self._apply_appearance_mode
        canvas.create_line(0, height - pad, width, height - pad, fill=color(TOKENS["track"]),
                           tags=_LINE_TAG)
        if len(self._values) < 2:
            return
        top = max(self._values) or 1.0
        step = width / (self._samples - 1)
        offset = self._samples - len(self._values)  # newest sample at the right edge
        points = []
        for i, value in enumerate(self._values):
            points += [(offset + i) * step, height - pad - value / top * (height - 2 * pad)]
        canvas.create_line(*points, fill=color(TOKENS["accent"]),
                           width=self._apply_widget_scaling(_LINE_WIDTH), tags=_LINE_TAG)
