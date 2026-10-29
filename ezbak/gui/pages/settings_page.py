"""Settings and help pages (U5-7, ui-plan §2, decision D1=B: the menu bar is gone).

What the View, Language and Help menus did is here, with the same calls:
- settings: Theme card (theme_light_btn / theme_dark_btn -> switch_theme), Language card
  (lang_en_btn / lang_ko_btn -> switch_language), Sound card (the layout puts
  sound_label and sound_toggle into its body);
- help: help_shortcuts_btn -> show_keyboard_shortcuts_help (F1 dialog),
  help_github_btn -> open_github_link(None), and the keyboard_shortcuts_content table.
Switching theme or language still shows the "restart for full effect" dialog.
update_texts(app) sets the titles and button texts in the current language.
"""
import customtkinter as ctk

from ..components import Card
from ..theme import GAP_CARD, TOKENS, font_spec

# card name: title key
_CARDS = {'theme': 'menu_theme', 'language': 'menu_language', 'sound': 'sound', 'help': 'nav_help'}
# button attribute: (card, text key or fixed text, column, command(app))
_BUTTONS = {
    'theme_light_btn': ('theme', 'theme_light', 0, lambda app: app.switch_theme('light')),
    'theme_dark_btn': ('theme', 'theme_dark', 1, lambda app: app.switch_theme('dark')),
    'lang_en_btn': ('language', None, 0, lambda app: app.switch_language('en')),
    'lang_ko_btn': ('language', None, 1, lambda app: app.switch_language('ko')),
    'help_shortcuts_btn': ('help', 'help_shortcuts', 0, lambda app: app.show_keyboard_shortcuts_help()),
    'help_github_btn': ('help', 'help_github', 1, lambda app: app.open_github_link(None)),
}
_FIXED_TEXTS = {'lang_en_btn': "English", 'lang_ko_btn': "한국어"}  # as the Language menu


def _text(app, attr):
    key = _BUTTONS[attr][1]
    return app.translator.get(key) if key else _FIXED_TEXTS[attr]


def build_cards(app):
    """Cards of the settings and help pages; returns {name: body frame}."""
    lang = app.translator.current_lang
    app.settings_cards, bodies = {}, {}
    for name, key in _CARDS.items():
        page = app.pages['help' if name == 'help' else 'settings']
        card = Card(page, title=app.translator.get(key), language=lang,
                    fg_color=app.theme.get('bg_elevated'), border_color=TOKENS['border'])
        card.pack(fill="x", pady=(0, GAP_CARD))
        app.settings_cards[name] = card
        bodies[name] = card.body
    rows = {}
    for attr, (card, _key, column, command) in _BUTTONS.items():
        if card not in rows:
            rows[card] = ctk.CTkFrame(bodies[card], fg_color="transparent")
            rows[card].pack(fill="x")
        button = ctk.CTkButton(rows[card], text=_text(app, attr), height=32,
                               font=font_spec('body', lang),
                               command=lambda c=command: c(app))
        button.grid(row=0, column=column, padx=(0, 8), sticky="w")
        setattr(app, attr, button)
    app.help_table_label = ctk.CTkLabel(bodies['help'], justify="left", anchor="w",
                                        text=app.translator.get('keyboard_shortcuts_content'),
                                        font=("Consolas", 12), text_color=TOKENS['fg'])
    app.help_table_label.pack(fill="x", pady=(12, 0))
    return bodies


def update_texts(app):
    for name, card in app.settings_cards.items():
        card.title_label.configure(text=app.translator.get(_CARDS[name]))
    for attr in _BUTTONS:
        getattr(app, attr).configure(text=_text(app, attr))
    app.help_table_label.configure(text=app.translator.get('keyboard_shortcuts_content'))
