"""Windows 11 style colour theme (moved from ezBAK.py in P5-2).

ThemeMixin (P5-9) holds App's _configure_progressbar_style, apply_theme_to_all_widgets
and switch_theme; it is mixed into App before ctk.CTk. `dialogs` is imported as a
module (not `from .dialogs import Win11Dialog`) because dialogs/message.py imports
this module: Win11Dialog is looked up when the method runs.
"""
import logging

import customtkinter as ctk

from . import dialogs, dpi

_log = logging.getLogger(__name__)


class Win11Theme:
    """Windows 11 style theme with light and dark mode support.

    U6-6 (ADR 0004): the keys that have a design token take its value (TOKENS below:
    bg, card, card_hover, fg, fg_muted, accent, accent_soft, border, success, danger,
    warning), so the tk widgets that read Win11Theme match the CTk ones. Keys without a
    token (hover/pressed shades, divider, shadow, disabled, scrollbar) keep their values.
    """

    # Light Theme Colors (Windows 11 Light)
    LIGHT = {
        'bg': '#F3F4F6',              # Main background
        'bg_secondary': '#F5F7FA',     # Secondary background
        'bg_elevated': '#FFFFFF',      # Elevated surfaces
        'fg': '#1F2328',              # Text color
        'fg_secondary': '#6B7280',    # Secondary text
        'accent': '#0078D4',          # Accent color (blue)
        'accent_hover': '#106EBE',    # Accent hover
        'accent_light': '#E1F1FF',    # Light accent background
        'success': '#107C10',         # Success/green
        'success_hover': '#0E6B0E',
        'danger': '#D13438',          # Danger/red
        'danger_hover': '#B32B2F',
        'warning': '#F7630C',         # Warning/orange
        'info': '#0078D4',            # Info/blue
        'border': '#E5E7EB',          # Border color
        'divider': '#EDEBE9',         # Divider
        'shadow': '#00000015',        # Shadow
        'disabled': '#A19F9D',        # Disabled state
        'hover': '#F5F7FA',           # Hover state
        'pressed': '#EBEBEB',         # Pressed state
        'text_bg': '#FFFFFF',         # Text area background
        'btn_bg': '#0078D4',          # Button background
        'btn_fg': '#ffffff',          # Button text
        'scrollbar_bg': '#C8C6C4',    # Scrollbar background
    }

    # Dark Theme Colors (Windows 11 Style)
    DARK = {
        'bg': '#202020',              # Main background - Windows 11 dark
        'bg_secondary': '#323232',     # Secondary background - Windows 11 elevated
        'bg_elevated': '#2B2B2B',      # Elevated surfaces
        'fg': '#FFFFFF',              # Text color - Pure white
        'fg_secondary': '#B3B3B3',    # Secondary text - Light gray
        'accent': '#0078D4',          # Accent color - Windows 11 blue
        'accent_hover': '#1A86D9',    # Accent hover - lighter blue on dark
        'accent_light': '#003D71',    # Dark accent background
        'success': '#0F7B0F',         # Success/green - Windows 11 green
        'success_hover': '#0D6B0D',
        'danger': '#C42B1C',          # Danger/red - Windows 11 red
        'danger_hover': '#A82315',
        'warning': '#F7630C',         # Warning/orange - Windows 11 orange
        'info': '#0078D4',            # Info/blue - Windows 11 info
        'border': '#3D3D3D',          # Border color - Subtle gray
        'divider': '#333333',         # Divider
        'shadow': '#00000040',        # Shadow
        'disabled': '#6d6d6d',        # Disabled state
        'hover': '#323232',           # Hover state
        'pressed': '#1a1a1a',         # Pressed state
        'text_bg': '#2B2B2B',         # Text area background
        'btn_bg': '#0078D4',          # Button background
        'btn_fg': '#ffffff',          # Button text
        'scrollbar_bg': '#5d5d5d',    # Scrollbar background
    }

    def __init__(self, initial_mode='dark'):
        self.current_mode = initial_mode
        self.colors = self.DARK if initial_mode == 'dark' else self.LIGHT

    def get(self, color_name):
        """Get color value by name"""
        return self.colors.get(color_name, '#000000')


# Design tokens for the CustomTkinter UI (U3-2, docs/nav-drafts/ui-plan.md §4).
# Colours are (light, dark) pairs as CTk expects; Win11Theme.LIGHT/DARK stay as they
# are until U6-6. Nothing uses these yet.
TOKENS = {
    'bg': ('#F3F4F6', '#202020'),            # window / content background
    'sidebar': ('#FFFFFF', '#1B1B1B'),
    'card': ('#FFFFFF', '#2B2B2B'),
    'card_hover': ('#F5F7FA', '#323232'),
    'border': ('#E5E7EB', '#3D3D3D'),
    'fg': ('#1F2328', '#FFFFFF'),
    'fg_muted': ('#6B7280', '#B3B3B3'),
    'accent': ('#0078D4', '#0078D4'),
    'accent_hover': ('#106EBE', '#1A86D9'),
    'accent_soft': ('#E1F1FF', '#003D71'),   # chips, selected sidebar item
    'success': ('#107C10', '#0F7B0F'),
    'danger': ('#D13438', '#C42B1C'),
    'warning': ('#F7630C', '#F7630C'),
    'track': ('#E5E7EB', '#484F58'),         # progress trough, switch off
}

SPACING = (4, 8, 12, 16, 24, 32)
PAD_CONTENT, PAD_CARD, GAP_CARD, GAP_BUTTON = 24, 16, 12, 8
RADIUS = {'card': 12, 'button': 8, 'input': 8, 'sidebar_item': 8}
RING_DIAMETER = 120
SIDEBAR_WIDTH = 220
WINDOW_SIZE = (1000, 700)
WINDOW_MIN_SIZE = (900, 600)

# Font family per UI language (decision D2 A); sizes are CTk pixels.
FONT_FAMILY = {'en': 'Segoe UI', 'ko': 'Malgun Gothic'}
LOG_FONT_FAMILY = 'Consolas'
FONT_SIZES = {
    'caption': (11, 'normal'),
    'body': (13, 'normal'),
    'body_bold': (13, 'bold'),
    'section': (16, 'bold'),
    'title': (20, 'bold'),
    'display': (28, 'bold'),
}

# Segoe MDL2 Assets glyphs for the sidebar (decision D10 A).
ICON_FONT_FAMILY = 'Segoe MDL2 Assets'
ICONS = {
    'home': '\uE80F',
    'backup': '\uE898',
    'restore': '\uE777',
    'schedule': '\uE787',
    'tools': '\uE90F',
    'logs': '\uE7C3',
    'settings': '\uE713',
    'help': '\uE897',
}


def font_spec(role, language='en'):
    """(family, size, weight) for a FONT_SIZES role in the given UI language."""
    size, weight = FONT_SIZES[role]
    return (FONT_FAMILY.get(language, FONT_FAMILY['en']), size, weight)


def tk_font(role, language='en'):
    """font_spec for plain tk widgets: a negative size is pixels, as CTk sizes are.
    Pixels at 100 %, multiplied by the App root's DPI factor (dpi.font_px, U8)."""
    family, size, weight = font_spec(role, language)
    return (family, -dpi.font_px(size), weight)


def token_colors(theme):
    """{token: colour} of TOKENS for the theme's mode (light or dark), for tk widgets that
    cannot take a (light, dark) pair: the dialogs from U6 on (Win11Theme stays for the
    rest until U6-6)."""
    index = 1 if theme.current_mode == 'dark' else 0
    return {name: pair[index] for name, pair in TOKENS.items()}


class ThemeMixin:
    def _configure_progressbar_style(self):
        """Configure professional progress bar style based on current theme"""
        if self.theme.current_mode == 'dark':
            # Dark theme: vibrant accent with subtle glow effect
            self.style.configure("green.Horizontal.TProgressbar",
                               troughcolor=self.theme.get('bg_elevated'),
                               background=self.theme.get('accent'),
                               bordercolor=self.theme.get('border'),
                               lightcolor=self.theme.get('accent_hover'),
                               darkcolor=self.theme.get('accent'),
                               borderwidth=0,
                               thickness=20)
        else:
            # Light theme: clean modern look
            self.style.configure("green.Horizontal.TProgressbar",
                               troughcolor=self.theme.get('bg_elevated'),
                               background=self.theme.get('accent'),
                               bordercolor=self.theme.get('border'),
                               lightcolor=self.theme.get('accent_light'),
                               darkcolor=self.theme.get('accent'),
                               borderwidth=0,
                               thickness=20)

    def _sync_ctk_appearance(self):
        """CustomTkinter appearance mode (global: CTk widgets and the title bar) = theme."""
        ctk.set_appearance_mode(self.theme.current_mode)

    def apply_theme_to_all_widgets(self):
        """Apply current theme to all widgets (requires app restart for full effect)"""
        # This is a simplified version - full theme switching would require rebuilding the UI
        self.configure(fg_color=self.theme.get('bg'))  # CTk root: fg_color, not bg
        self._sync_ctk_appearance()

        # Update progress bar style
        self._configure_progressbar_style()

        # Update sound toggle theme
        if hasattr(self, 'sound_toggle'):
            self.sound_toggle.update_theme(self.theme)

        try:
            self.save_settings()  # Save the theme preference
        except Exception:
            # settings.json not writable: the choice applies to this session only
            _log.debug("event=settings_save_failed source=theme", exc_info=True)

    def switch_theme(self, mode):
        """Switch between light and dark themes"""
        if mode != self.theme.current_mode:
            self.theme.current_mode = mode
            self.theme.colors = self.theme.DARK if mode == 'dark' else self.theme.LIGHT
            self.apply_theme_to_all_widgets()
            mode_text = self.translator.get('theme_dark') if mode == 'dark' else self.translator.get('theme_light')
            dialogs.Win11Dialog.showinfo(self.translator.get('theme_changed'),
                                         self.translator.get('theme_changed_msg').format(mode_text),
                                         parent=self, theme=self.theme, translator=self.translator, show_icon=False)
