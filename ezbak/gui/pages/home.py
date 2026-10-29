"""Home page (U5-1, ui-plan §2; data cards U7-5, ui-plan §10).

- status card: RingButton home_backup_btn (= start_backup_thread; shows the progress
  bar's value as a ring while a job runs), the selected user and a filter chip
  (include + exclude rule count, opens the filter manager);
- two StatTiles: backups to keep, log retention days;
- system card: system name (U7), number of users in the list, hidden / system files;
- U7 data cards: free space tile (last backup's drive), copy speed card with a
  Sparkline, last backup card and next scheduled backup card. Their values come from
  home_data.HomeData (filled by home_data.tick on the status timer, a work thread for
  file reads); until then they show placeholders.
The Activity Log card is shared with the logs page (nav.show_page, decision D9).
Texts come from the translator; refresh(app) rewrites them and is called on build, when
an option variable changes, from update_ui_texts and from the status chip timer.
"""
import os
import tkinter as tk

import customtkinter as ctk

from ezbak.core import format_bytes

from ..components import Card, Chip, RingButton, Sparkline, StatTile
from ..theme import GAP_CARD, TOKENS, font_spec
from .home_data import HomeData


def build(app, parent):
    """Widgets of the home page in parent (the page frame); kept on app as home_*."""
    lang = app.translator.current_lang
    top = tk.Frame(parent, bg=app.theme.get('bg'))
    top.pack(fill="x", pady=(0, GAP_CARD))
    status = Card(top, language=lang)
    status.pack(fill="x")
    app.home_backup_btn = RingButton(status.body, text="", command=app.start_backup_thread,
                                     language=lang)
    app.home_backup_btn.pack(side="left", padx=(0, 24))
    info = ctk.CTkFrame(status.body, fg_color="transparent")
    info.pack(side="left", fill="both", expand=True)
    app.home_user_caption = ctk.CTkLabel(info, text="", anchor="w",
                                         font=font_spec('caption', lang),
                                         text_color=TOKENS['fg_muted'])
    app.home_user_caption.pack(fill="x", pady=(16, 0))
    app.home_user_label = ctk.CTkLabel(info, text="", anchor="w", font=font_spec('title', lang),
                                       text_color=TOKENS['fg'])
    app.home_user_label.pack(fill="x")
    app.home_filter_chip = Chip(info, text="", command=app.open_filter_manager, language=lang)
    app.home_filter_chip.pack(anchor="w", pady=(8, 0))

    stats = tk.Frame(parent, bg=app.theme.get('bg'))
    stats.pack(fill="x", pady=(0, GAP_CARD))
    app.home_retention_tile = StatTile(stats, caption="", language=lang)
    app.home_retention_tile.pack(side="left", fill="x", expand=True, padx=(0, GAP_CARD // 2))
    app.home_log_days_tile = StatTile(stats, caption="", language=lang)
    app.home_log_days_tile.pack(side="left", fill="x", expand=True, padx=(GAP_CARD // 2,) * 2)
    app.home_free_tile = StatTile(stats, caption="", language=lang)
    app.home_free_tile.pack(side="left", fill="x", expand=True, padx=(GAP_CARD // 2, 0))

    _build_data_cards(app, parent, lang)

    app.home_system_card = Card(parent, title="", language=lang)
    app.home_system_card.pack(fill="x", pady=(0, GAP_CARD))
    _labels(app, app.home_system_card.body, lang, 'home_system_name_label',
            'home_user_count_label', 'home_hidden_label', 'home_system_label')

    for var in (app.user_var, app.hidden_mode_var, app.system_mode_var,
                app.retention_count_var, app.log_retention_days_var):
        var.trace_add('write', lambda *args: refresh(app))
    refresh(app)


def _labels(app, parent, lang, *attrs):
    for attr in attrs:
        label = ctk.CTkLabel(parent, text="", anchor="w", font=font_spec('body', lang),
                             text_color=TOKENS['fg'])
        label.pack(fill="x")
        setattr(app, attr, label)


def _build_data_cards(app, parent, lang):
    """U7-5: copy speed card, then the last backup and next schedule cards side by side."""
    app.home_data = HomeData()
    app.home_speed_card = Card(parent, title="", language=lang)
    app.home_speed_card.pack(fill="x", pady=(0, GAP_CARD))
    _labels(app, app.home_speed_card.body, lang, 'home_speed_label')
    app.home_speed_chart = Sparkline(app.home_speed_card.body)
    app.home_speed_chart.pack(fill="x", pady=(8, 0))

    row = tk.Frame(parent, bg=app.theme.get('bg'))
    row.pack(fill="x", pady=(0, GAP_CARD))
    app.home_last_card = Card(row, title="", language=lang)
    app.home_last_card.pack(side="left", fill="both", expand=True, padx=(0, GAP_CARD // 2))
    _labels(app, app.home_last_card.body, lang, 'home_last_time_label', 'home_last_detail_label',
            'home_last_folder_label')
    app.home_next_card = Card(row, title="", language=lang)
    app.home_next_card.pack(side="left", fill="both", expand=True, padx=(GAP_CARD // 2, 0))
    _labels(app, app.home_next_card.body, lang, 'home_next_name_label', 'home_next_time_label')


def _refresh_data_cards(app):
    """U7-5 texts from app.home_data (placeholders until home_data.tick has loaded them)."""
    t, data = app.translator.get, app.home_data
    _set(app.home_free_tile.caption_label, t('home_free_space'))
    _set(app.home_free_tile.value_label,
         f"{format_bytes(data.free_space[1])} ({data.free_space[0]})" if data.free_space else "-")
    _set(app.home_speed_card.title_label, t('home_speed_card'))
    _set(app.home_speed_label,
         f"{format_bytes(data.speed)}/s" if data.speed is not None else t('home_speed_idle'))
    _set(app.home_last_card.title_label, t('home_last_backup_card'))
    last = data.last_backup
    if last:
        mode = t('home_mode_headless' if last.get('mode') == 'headless' else 'home_mode_gui')
        detail = f"{last.get('user', '-')} · {mode}"
        folder = os.path.basename(str(last.get('folder') or '')) or '-'
    else:
        detail, folder = "", ""
    _set(app.home_last_time_label, last['time'] if last else t('home_last_backup_none'))
    _set(app.home_last_detail_label, detail)
    _set(app.home_last_folder_label, folder)
    _set(app.home_next_card.title_label, t('home_next_schedule_card'))
    if data.next_run:
        name, when = data.next_run
        _set(app.home_next_name_label, name or '-')
        _set(app.home_next_time_label, when.strftime('%Y-%m-%d %H:%M'))
    else:
        _set(app.home_next_name_label, t('home_next_schedule_none'))
        _set(app.home_next_time_label, "")
    _set(app.home_system_name_label,
         f"{t('home_system_name')}: {data.system_name or t('home_loading')}")


def _included(app, var):
    return app.translator.get('home_included' if var.get() == 'include' else 'home_excluded')


def _set(widget, text):
    if widget.cget("text") != text:  # no redraw when nothing changed (timer)
        widget.configure(text=text)


def refresh(app):
    """All home texts and values from the App's current state and language."""
    t = app.translator.get
    _set(app.home_backup_btn, t('home_start_backup'))
    _set(app.home_user_caption, t('home_selected_user'))
    _set(app.home_user_label, app.user_var.get() or "-")
    rules = len(app.filters.get('include', [])) + len(app.filters.get('exclude', []))
    _set(app.home_filter_chip, t('home_filters').format(rules))
    _set(app.home_retention_tile.caption_label, t('home_retention'))
    _set(app.home_retention_tile.value_label, app.retention_count_var.get())
    _set(app.home_log_days_tile.caption_label, t('home_log_days'))
    _set(app.home_log_days_tile.value_label, app.log_retention_days_var.get())
    _set(app.home_system_card.title_label, t('home_system_card'))
    _set(app.home_user_count_label,
         f"{t('home_user_count')}: {len(app.user_combo['values'] or ())}")
    _set(app.home_hidden_label, f"{t('home_hidden_files')}: {_included(app, app.hidden_mode_var)}")
    _set(app.home_system_label, f"{t('home_system_files')}: {_included(app, app.system_mode_var)}")
    _refresh_data_cards(app)


def show_progress(app, fraction):
    """Ring of the home start button: None when idle, else 0..1 (status chip timer)."""
    if app.home_backup_btn.get_progress() != fraction:
        app.home_backup_btn.set_progress(fraction)
