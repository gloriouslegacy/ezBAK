"""Window shell of the redesigned App: sidebar and page switching (U4-2, ui-plan §1).

NavMixin is mixed into App through LayoutMixin (app.py unchanged). _build_shell makes
the left sidebar (logo, status chip, eight SidebarItems, footer) and the content area:
top bar, one tk.Frame per page (all built at start, show_page only packs one and
pack_forgets the shown one, so widgets, bindings and set_buttons_state targets stay),
and the bottom row that holds the progress card on every page.
The status chip reads progress_bar every 0.5 s (bridge and queue unchanged); the same
timer refreshes the home page and its ring (U5-1).
The Activity Log card is a child of page_area and is packed into the home or the logs
page by show_page (pack in_=, one widget for both: decision D9).
"""
import tkinter as tk

import customtkinter as ctk

from .components import SidebarItem
from .pages import backup as backup_page
from .pages import home as home_page
from .pages import home_data
from .pages import settings_page
from .theme import ICONS, SIDEBAR_WIDTH, TOKENS, font_spec

PAGES = ("home", "backup", "restore", "schedule", "tools", "logs", "settings", "help")
DEFAULT_PAGE = "home"
LOG_PAGES = ("home", "logs")  # pages that show the shared Activity Log card (D9)
STATUS_REFRESH_MS = 500


class NavMixin:
    def _build_shell(self):
        """Sidebar on the left, content (top bar, pages, bottom row) on the right."""
        self._build_sidebar()
        content = tk.Frame(self, bg=self.theme.get('bg'))
        content.pack(side="left", fill="both", expand=True)
        self.topbar = tk.Frame(content, bg=self.theme.get('bg'))
        self.topbar.pack(fill="x")
        self.page_title_label = tk.Label(self.topbar, text="",
                                         font=("Segoe UI", 16, "bold"),
                                         bg=self.theme.get('bg'), fg=self.theme.get('fg'))
        self.page_title_label.pack(side="bottom", anchor="w", padx=24, pady=(12, 0))
        self.bottom_bar = tk.Frame(content, bg=self.theme.get('bg'), padx=24)
        self.bottom_bar.pack(side="bottom", fill="x")
        self.page_area = tk.Frame(content, bg=self.theme.get('bg'))
        self.page_area.pack(fill="both", expand=True)
        self.pages = {page: tk.Frame(self.page_area, bg=self.theme.get('bg'), padx=24, pady=20)
                      for page in PAGES}
        self.current_page = None

    def _build_sidebar(self):
        lang = self.translator.current_lang
        self.sidebar = ctk.CTkFrame(self, width=SIDEBAR_WIDTH, corner_radius=0,
                                    fg_color=TOKENS['sidebar'])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        ctk.CTkLabel(self.sidebar, text="ezBAK", anchor="w", font=font_spec('title', lang),
                     text_color=TOKENS['fg']).pack(fill="x", padx=20, pady=(20, 4))
        self.nav_status_label = ctk.CTkLabel(self.sidebar, text=self.translator.get('status_idle'),
                                             font=font_spec('caption', lang), height=24,
                                             corner_radius=12, fg_color=TOKENS['accent_soft'],
                                             text_color=TOKENS['accent'])
        self.nav_status_label.pack(anchor="w", padx=20, pady=(0, 16))
        self.nav_items = {}
        for page in PAGES:
            item = SidebarItem(self.sidebar, icon=ICONS[page],
                               text=self.translator.get(f'nav_{page}'),
                               command=lambda p=page: self.show_page(p), language=lang)
            item.pack(fill="x", padx=12, pady=2)
            self.nav_items[page] = item

    def show_page(self, page):
        """Show one page (PAGES), mark its sidebar entry and set the top bar title."""
        if self.current_page is not None:
            self.pages[self.current_page].pack_forget()
        self.pages[page].pack(fill="both", expand=True)
        self.current_page = page
        log_card = getattr(self, 'log_card', None)
        if log_card is not None:
            if page in LOG_PAGES:
                log_card.pack(in_=self.pages[page], fill="both", expand=True, pady=(0, 12))
            else:
                log_card.pack_forget()
        for name, item in self.nav_items.items():
            item.set_selected(name == page)
        self.page_title_label.config(text=self.translator.get(f'nav_{page}'))

    def _update_nav_texts(self):
        """Sidebar, title and status chip texts in the current language (update_ui_texts)."""
        for page, item in self.nav_items.items():
            item.configure(text=self.translator.get(f'nav_{page}'))
        if self.current_page is not None:
            self.page_title_label.config(text=self.translator.get(f'nav_{self.current_page}'))
        self.nav_status_label.configure(text=self._nav_status_text())
        if hasattr(self, 'home_backup_btn'):
            home_page.refresh(self)
        if hasattr(self, 'page_cards'):
            backup_page.update_texts(self)
        if hasattr(self, 'settings_cards'):
            settings_page.update_texts(self)
        if hasattr(self, 'log_title_label'):
            self.log_title_label.config(text=self.translator.get('activity_log'))

    def _progress_fraction(self):
        """0..1 while a job shows progress, None when idle."""
        value = float(self.progress_bar['value'])
        maximum = float(self.progress_bar['maximum']) or 100.0
        return value / maximum if 0 < value < maximum else None

    def _nav_status_text(self):
        fraction = self._progress_fraction()
        if fraction is not None:
            return self.translator.get('status_working').format(int(fraction * 100))
        return self.translator.get('status_idle')

    def _refresh_nav_status(self):
        """Status chip, home ring and home values, every STATUS_REFRESH_MS (read only)."""
        self.nav_status_label.configure(text=self._nav_status_text())
        if hasattr(self, 'home_backup_btn'):
            home_page.show_progress(self, self._progress_fraction())
            home_data.tick(self)  # U7-5: speed sample, slow loads (only when viewable)
            home_page.refresh(self)
        self.after(STATUS_REFRESH_MS, self._refresh_nav_status)
