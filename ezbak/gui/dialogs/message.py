"""Windows 11 style message dialogs (moved from ezBAK.py in P5-3).

Colours and fonts come from the design tokens (theme.token_colors / tk_font, U6-1, ADR
0004) for the theme's light or dark mode; layout, texts, buttons and results are as before.

Tests patch `Win11Dialog` by name; the internal calls below go through the
module-level name so a patched class is used everywhere.
"""
import tkinter as tk

from ezbak.winsys.paths import resource_path

from .. import dpi
from ..i18n import Translator
from ..theme import Win11Theme, tk_font, token_colors


def _open_askyesno_window(title, message, parent, theme):
    """Create, decorate and centre the askyesno Toplevel (first part of Win11Dialog.askyesno)."""
    colors = token_colors(theme)
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    # Set titlebar icon
    try:
        dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
    except Exception:
        pass  # titlebar icon is decoration only; a missing .ico must not block the dialog
    dlg.configure(bg=colors['card'])
    dlg.resizable(False, False)
    dlg.transient(parent)
    dlg.grab_set()

    # Calculate required height based on message length
    # Estimate: ~15 pixels per line, plus padding
    line_count = message.count('\n') + 1
    text_height = max(100, line_count * 20)  # Minimum 100, ~20 pixels per line
    total_height = text_height + 120  # Add space for buttons and padding

    # Center the dialog
    dlg.update_idletasks()
    width = 450
    height = min(600, total_height)  # Maximum 600 pixels
    dlg.geometry(dpi.centered_geometry(dlg, width, height))  # px at 100 % (U8)
    return dlg


class Win11Dialog:
    """Windows 11 style modern dialogs to replace default messageboxes"""

    @staticmethod
    def _create_dialog_button(parent, text, command, theme, is_primary=False):
        """Create a Windows 11 style dialog button with modern design"""
        colors = token_colors(theme)
        # Windows 11 style: Primary buttons have accent color, secondary have subtle background
        if is_primary:
            btn_bg = colors['accent']
            btn_fg = '#FFFFFF' if theme.current_mode == 'dark' else '#FFFFFF'
            hover_bg = colors['accent_hover']
            hover_fg = '#FFFFFF'
            pressed_bg = colors['accent_hover']
        else:
            btn_bg = colors['card']
            btn_fg = colors['fg']
            hover_bg = colors['card_hover']
            hover_fg = colors['fg']
            pressed_bg = colors['border']

        border_color = colors['border']

        btn = tk.Button(parent, text=text,
                       bg=btn_bg,
                       fg=btn_fg,
                       font=tk_font('body'),
                       relief="flat",
                       bd=0,
                       highlightthickness=1,
                       highlightbackground=border_color,
                       highlightcolor=border_color,
                       borderwidth=1,
                       width=10,
                       height=1,
                       command=command,
                       cursor="hand2",
                       padx=12,
                       pady=6)

        # Windows 11 style hover and press effects
        def on_enter(e):
            e.widget.config(bg=hover_bg, fg=hover_fg, relief="flat")

        def on_leave(e):
            e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

        def on_press(e):
            e.widget.config(bg=pressed_bg, relief="flat")

        def on_release(e):
            e.widget.config(bg=hover_bg, relief="flat")

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)

        return btn

    @staticmethod
    def showinfo(title, message, parent=None, theme=None, translator=None, show_icon=True):
        """Show information dialog"""
        if theme is None:
            theme = Win11Theme('dark')
        if translator is None:
            translator = Translator('en')
        colors = token_colors(theme)

        dlg = tk.Toplevel(parent)
        dlg.title(title)
        # Set titlebar icon
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # titlebar icon is decoration only; a missing .ico must not block the dialog
        dlg.configure(bg=colors['card'])
        dlg.resizable(False, False)
        dlg.transient(parent)
        dlg.grab_set()

        # Center the dialog
        dlg.update_idletasks()
        width = 490
        height = 380  # Increased to accommodate longer messages like schedule creation confirmation
        dlg.geometry(dpi.centered_geometry(dlg, width, height))  # px at 100 % (U8)

        # Message area
        msg_frame = tk.Frame(dlg, bg=colors['card'])
        msg_frame.pack(fill='both', expand=True, padx=24, pady=20)

        # Icon and message in horizontal layout
        content_frame = tk.Frame(msg_frame, bg=colors['card'])
        content_frame.pack(fill='both', expand=True)

        # Add info icon (optional)
        if show_icon:
            icon_label = tk.Label(content_frame, text="ℹ️",
                                 bg=colors['card'],
                                 fg=colors['accent'],
                                 font=("Segoe UI", 24))
            icon_label.pack(side='left', padx=(0, 12))

        tk.Label(content_frame, text=message,
                bg=colors['card'],
                fg=colors['fg'],
                font=tk_font('body', translator.current_lang),
                wraplength=dpi.px(dlg, 360),
                justify='left').pack(side='left', fill='both', expand=True)

        # Button area
        btn_frame = tk.Frame(dlg, bg=colors['card'])
        btn_frame.pack(fill='x', padx=24, pady=(0, 20))

        Win11Dialog._create_dialog_button(btn_frame, translator.get('dialog_ok'), dlg.destroy, theme, is_primary=True).pack(side='right')

        dlg.wait_window()

    @staticmethod
    def showwarning(title, message, parent=None, theme=None, translator=None):
        """Show warning dialog"""
        Win11Dialog.showinfo(title, message, parent, theme, translator)

    @staticmethod
    def showerror(title, message, parent=None, theme=None, translator=None):
        """Show error dialog"""
        Win11Dialog.showinfo(title, message, parent, theme, translator)

    @staticmethod
    def askyesno(title, message, parent=None, theme=None, translator=None):
        """Show yes/no dialog"""
        if theme is None:
            theme = Win11Theme('dark')
        if translator is None:
            translator = Translator('en')
        colors = token_colors(theme)

        result = [False]

        dlg = _open_askyesno_window(title, message, parent, theme)

        # Message area
        msg_frame = tk.Frame(dlg, bg=colors['card'])
        msg_frame.pack(fill='both', expand=True, padx=24, pady=20)

        tk.Label(msg_frame, text=message,
                bg=colors['card'],
                fg=colors['fg'],
                font=tk_font('body', translator.current_lang),
                wraplength=dpi.px(dlg, 400),
                justify='left').pack()

        # Button area
        btn_frame = tk.Frame(dlg, bg=colors['card'])
        btn_frame.pack(fill='x', padx=24, pady=(0, 20))

        def on_yes():
            result[0] = True
            dlg.destroy()

        def on_no():
            result[0] = False
            dlg.destroy()

        Win11Dialog._create_dialog_button(btn_frame, translator.get('dialog_no'), on_no, theme, is_primary=False).pack(side='right', padx=(8, 0))
        Win11Dialog._create_dialog_button(btn_frame, translator.get('dialog_yes'), on_yes, theme, is_primary=True).pack(side='right')

        dlg.wait_window()
        return result[0]
