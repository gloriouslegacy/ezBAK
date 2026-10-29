"""Page cards of the backup, restore, scheduled backup and tools pages (U5-2~5, ui-plan §2).

build_cards puts three titled Cards on the backup page, one each on the restore and
scheduled backup pages and three on the tools page, and returns their body frames; the layout sections then create the existing widgets in
them (same attributes, commands and variables):
- backup: backup_btn, filters_btn and the sound toggle;
- options: hidden/system radios, backups to keep, log days;
- extra: driver_backup_btn, browser_profiles_btn, winget_export_btn;
- restore (restore page): restore_btn, driver_restore_btn, between two hint labels
  (restore_hint_label above, restore_drivers_hint_label below; U5-3). The overwrite
  question and the folder choice stay dialogs (start_restore_thread);
- schedule (scheduled backup page): schedule_hint_label, then schedule_btn (U5-4). Create
  and delete stay in ScheduleBackupDialog;
- tools page (U5-5): tools_disk (check_space_btn, copy_btn), tools_system (devmgr_btn,
  file_explorer_btn; save_log_btn moved to the Activity Log card in U5-6), tools_nas
  (nas_connect_btn, nas_disconnect_btn).
The cards use the Win11Theme 'bg_elevated' colour so the tk widgets in them match.
update_texts(app) sets the card titles and hints in the current language.
"""
import customtkinter as ctk

from ..components import Card
from ..theme import GAP_CARD, TOKENS, font_spec

_HINTS = {'restore_hint_label': 'restore_hint', 'restore_drivers_hint_label': 'restore_drivers_hint',
          'schedule_hint_label': 'schedule_hint'}
# hint label: (card, pack side)
_HINT_PLACES = {'restore_hint_label': ('restore', 'top'),
                'restore_drivers_hint_label': ('restore', 'bottom'),
                'schedule_hint_label': ('schedule', 'top')}
HINT_WRAP = 640

_TITLES = {'backup': 'backup_main', 'options': 'backup_options', 'extra': 'backup_extra',
           'restore': 'restore_main', 'schedule': 'schedule_main',
           'tools_disk': 'tools_disk', 'tools_system': 'tools_system', 'tools_nas': 'tools_nas'}


def _card(app, parent, name):
    card = Card(parent, title=app.translator.get(_TITLES[name]),
                language=app.translator.current_lang,
                fg_color=app.theme.get('bg_elevated'), border_color=TOKENS['border'])
    card.pack(fill="x", pady=(0, GAP_CARD))
    app.page_cards[name] = card
    return card.body


def build_cards(app):
    """Cards of the backup and restore pages; returns {name: body frame}."""
    app.page_cards = {}
    bodies = {name: _card(app, app.pages['backup'], name)
              for name in ('backup', 'options', 'extra')}
    bodies['restore'] = _card(app, app.pages['restore'], 'restore')
    bodies['schedule'] = _card(app, app.pages['schedule'], 'schedule')
    for name in ('tools_disk', 'tools_system', 'tools_nas'):
        bodies[name] = _card(app, app.pages['tools'], name)
    for attr, (card, side) in _HINT_PLACES.items():
        label = ctk.CTkLabel(bodies[card], text=app.translator.get(_HINTS[attr]),
                             anchor="w", justify="left", wraplength=HINT_WRAP,
                             font=font_spec('body', app.translator.current_lang),
                             text_color=TOKENS['fg_muted'])
        label.pack(side=side, fill="x", pady=(0, 8) if side == 'top' else (8, 0))
        setattr(app, attr, label)
    return bodies


def update_texts(app):
    for name, card in app.page_cards.items():
        card.title_label.configure(text=app.translator.get(_TITLES[name]))
    for attr, key in _HINTS.items():
        getattr(app, attr).configure(text=app.translator.get(key))
