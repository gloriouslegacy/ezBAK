"""Main window layout of App (moved from ezBAK.py in P5-9).

LayoutMixin is mixed into App before ctk.CTk. create_widgets builds the window shell
(nav.NavMixin: sidebar, top bar, pages, bottom row; U4-2) and then runs four section
methods (header/options, backup, tools, progress/log) holding the original statements
in the original order. U4-2 only changed their parents (ui-plan §2.1 U4 column):
header clock and user card -> top bar, progress card -> bottom row, Activity Log ->
logs page, footer -> sidebar. From U5-2 the buttons and options go into the page cards
of pages/backup.py (backup, options, additional backups, restore, schedule, and the
three tools cards; save_log_btn sits on the Activity Log card since U5-6), and every
button has one of three kinds (primary, neutral, danger). Locals shared between
sections (main_frame, create_button) are passed explicitly; widgets kept on self use
the same attribute names as before.
Module-level names (tk, ttk, ToggleSwitch) match ezBAK.py.
"""
import logging
import tkinter as tk
from tkinter import ttk

from . import dpi
from .nav import DEFAULT_PAGE, NavMixin
from .pages import backup as backup_page
from .pages import home as home_page
from .pages import settings_page
from .widgets import ToggleSwitch

_log = logging.getLogger(__name__)


class LayoutMixin(NavMixin):
    def create_widgets(self):
        self._build_shell()
        main_frame = self._build_header_section()
        create_button = self._build_backup_section(main_frame)
        self._build_tools_section(main_frame, create_button)
        self._build_progress_log_section(main_frame, create_button)
        home_page.build(self, self.pages['home'])
        self.show_page(DEFAULT_PAGE)
        self._refresh_nav_status()

        # Note: update_ui_texts() is not called here to prevent duplicate Notice in Activity Log
        # The initial UI texts are already set during widget creation and in __init__
        # update_ui_texts() will be called when language is changed via switch_language()

    def _build_header_section(self):
        """Header clock (top bar), the page cards and the options row (backup page); returns
        the backup page, which the next sections call main_frame. (No menu bar since U5-7:
        its commands are on the settings and help pages, D1=B.)"""
        # Modern Windows 11 Header with theme colors - Minimal design with time only
        header_frame = tk.Frame(self.topbar, bg=self.theme.get('bg_secondary'), height=40)
        header_frame.pack(fill="x", pady=0)
        header_frame.pack_propagate(False)

        # Center the time in the header
        header_center = tk.Frame(header_frame, bg=self.theme.get('bg_secondary'))
        header_center.pack(expand=True, fill='both')

        self.header_time_label = tk.Label(header_center, text="",
                                         font=("Segoe UI", 11),
                                         bg=self.theme.get('bg_secondary'),
                                         fg=self.theme.get('fg'))
        self.header_time_label.pack(expand=True)
        self._update_header_time()

        # Main content frame: the backup page (U4-2) with its cards (U5-2)
        main_frame = self.pages['backup']
        self.page_bodies = backup_page.build_cards(self)
        self.page_bodies.update(settings_page.build_cards(self))

        # Options row in the options card
        options_row = tk.Frame(self.page_bodies['options'], bg=self.theme.get('bg_elevated'),
                               relief="flat", bd=0)
        options_row.pack(fill="x")

        # F1: two rows so System fits at the default width (retention | hidden/system)
        opt_frame = tk.Frame(options_row, bg=self.theme.get('bg_elevated'))
        opt_frame.pack(anchor="w")
        opt_frame2 = tk.Frame(options_row, bg=self.theme.get('bg_elevated'))
        opt_frame2.pack(anchor="w", pady=(6, 0))

        modern_font = ("Segoe UI", 9)

        # Retention count (left side)
        retention_label = tk.Label(opt_frame, text="Backups to Keep:",
                                  bg=self.theme.get('bg_elevated'),
                                  fg=self.theme.get('fg'),
                                  font=("Segoe UI", 9, "bold"))
        retention_label.pack(side="left", padx=(8, 4))

        retention_spinbox = tk.Spinbox(opt_frame, from_=0, to=99,
                                      textvariable=self.retention_count_var,
                                      width=5, font=modern_font,
                                      bg=self.theme.get('bg_elevated'),
                                      fg=self.theme.get('fg'),
                                      buttonbackground=self.theme.get('accent'),
                                      relief="solid", bd=1)
        retention_spinbox.pack(side="left", padx=(0, 2))

        retention_info_label = tk.Label(opt_frame, text="(0=all)",
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg_secondary'),
                                       font=modern_font)
        retention_info_label.pack(side="left", padx=(0, 12))

        sep = ttk.Separator(opt_frame, orient='vertical')
        sep.pack(side="left", fill='y', padx=10, pady=4)

        # Log retention
        log_retention_label = tk.Label(opt_frame, text="Log Days:",
                                      bg=self.theme.get('bg_elevated'),
                                      fg=self.theme.get('fg'),
                                      font=("Segoe UI", 9, "bold"))
        log_retention_label.pack(side="left", padx=(0, 4))

        log_retention_spinbox = tk.Spinbox(opt_frame, from_=0, to=365,
                                          textvariable=self.log_retention_days_var,
                                          width=5, font=modern_font,
                                          bg=self.theme.get('bg_elevated'),
                                          fg=self.theme.get('fg'),
                                          buttonbackground=self.theme.get('accent'),
                                          relief="solid", bd=1)
        log_retention_spinbox.pack(side="left", padx=(0, 2))

        log_retention_info_label = tk.Label(opt_frame, text="(0=off)",
                                           bg=self.theme.get('bg_elevated'),
                                           fg=self.theme.get('fg_secondary'),
                                           font=modern_font)
        log_retention_info_label.pack(side="left", padx=(0, 12))

        # Hidden files option (second row)
        hidden_label = tk.Label(opt_frame2, text="Hidden:",
                               bg=self.theme.get('bg_elevated'),
                               fg=self.theme.get('fg'),
                               font=("Segoe UI", 9, "bold"))
        hidden_label.pack(side="left", padx=(8, 6))

        hidden_include = tk.Radiobutton(opt_frame2, text="Include",
                                       variable=self.hidden_mode_var, value='include',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        hidden_include.pack(side="left", padx=(0, 4))

        hidden_exclude = tk.Radiobutton(opt_frame2, text="Exclude",
                                       variable=self.hidden_mode_var, value='exclude',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        hidden_exclude.pack(side="left", padx=(0, 12))

        sep3 = ttk.Separator(opt_frame2, orient='vertical')
        sep3.pack(side="left", fill='y', padx=10, pady=4)

        # System files option
        system_label = tk.Label(opt_frame2, text="System:",
                               bg=self.theme.get('bg_elevated'),
                               fg=self.theme.get('fg'),
                               font=("Segoe UI", 9, "bold"))
        system_label.pack(side="left", padx=(0, 6))

        system_include = tk.Radiobutton(opt_frame2, text="Include",
                                       variable=self.system_mode_var, value='include',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        system_include.pack(side="left", padx=(0, 4))

        system_exclude = tk.Radiobutton(opt_frame2, text="Exclude",
                                       variable=self.system_mode_var, value='exclude',
                                       bg=self.theme.get('bg_elevated'),
                                       fg=self.theme.get('fg'),
                                       selectcolor=self.theme.get('bg_elevated'),
                                       activebackground=self.theme.get('bg_elevated'),
                                       activeforeground=self.theme.get('accent'),
                                       font=modern_font, bd=0, relief="flat")
        system_exclude.pack(side="left", padx=(0, 8))

        return main_frame

    def _build_backup_section(self, main_frame):
        """Settings traces, user selection, sound toggle and the backup/restore/additional
        backup buttons (U5-2 cards); returns create_button."""
        # Persist settings
        try:
            self.hidden_mode_var.trace_add('write', lambda *args: self.save_settings())
            self.system_mode_var.trace_add('write', lambda *args: self.save_settings())
            self.retention_count_var.trace_add('write', lambda *args: self.save_settings())
            self.log_retention_days_var.trace_add('write', lambda *args: self.save_settings())
        except Exception:
            # trace_add unavailable: options just are not auto-saved
            _log.debug("event=settings_autosave_trace_failed", exc_info=True)

        # User selection with modern card style
        user_card = tk.Frame(self.topbar, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        user_card.pack(fill="x", padx=24, pady=(12, 0))

        user_inner = tk.Frame(user_card, bg=self.theme.get('bg_elevated'))
        user_inner.pack(fill="x", padx=16, pady=12)

        tk.Label(user_inner, text="Select User:",
                font=("Segoe UI", 11, "bold"),
                background=self.theme.get('bg_elevated'),
                foreground=self.theme.get('fg')).pack(side="left", padx=(0, 12))

        self.user_var = tk.StringVar()
        self.user_combo = ttk.Combobox(user_inner, textvariable=self.user_var,
                                      state="readonly", width=35, font=("Segoe UI", 10))
        self.user_combo.pack(side="left", fill="x", expand=True)
        self.load_users()

        # Sound toggle: settings page, sound card (U5-7)
        sound_frame = tk.Frame(self.page_bodies['sound'], bg=self.theme.get('bg_elevated'))
        sound_frame.pack(side="left")

        # Sound label
        self.sound_label = tk.Label(
            sound_frame,
            text=self.translator.get('sound'),
            bg=self.theme.get('bg_elevated'),
            fg=self.theme.get('fg'),
            font=("Segoe UI", 10)
        )
        self.sound_label.pack(side="left", padx=(0, 8))

        # Sound toggle switch
        self.sound_toggle = ToggleSwitch(
            sound_frame,
            width=dpi.px(self, 50),  # canvas pixels at 100 % (U8)
            height=dpi.px(self, 24),
            variable=self.sound_enabled_var,
            command=self._on_sound_toggle,
            theme=self.theme
        )
        self.sound_toggle.pack(side="left")

        main_buttons_frame = tk.Frame(self.page_bodies['backup'], bg=self.theme.get('bg_elevated'))
        main_buttons_frame.pack(side="left", fill="x", expand=True)
        restore_buttons_frame = tk.Frame(self.page_bodies['restore'], bg=self.theme.get('bg_elevated'))
        restore_buttons_frame.pack(fill="x")
        extra_buttons_frame = tk.Frame(self.page_bodies['extra'], bg=self.theme.get('bg_elevated'))
        extra_buttons_frame.pack(fill="x")
        for frame, columns in ((main_buttons_frame, 2), (restore_buttons_frame, 2),
                               (extra_buttons_frame, 3)):
            for i in range(columns):
                frame.columnconfigure(i, weight=1, uniform="button")

        btn_font = ("Segoe UI", 9)
        btn_height = 1

        # Create button helper function - Windows 11 style
        # Three button kinds (U5-2): primary action, neutral, danger
        kinds = {'primary': ('accent', 'accent_hover', None),
                 'neutral': ('bg_secondary', 'border', 'fg'),
                 'danger': ('danger', 'danger_hover', None)}

        def create_button(parent, text, kind, command, row, col):
            """Windows 11 style modern button with smooth transitions"""
            bg_key, hover_key, fg_key = kinds[kind]
            btn_bg = self.theme.get(bg_key)
            btn_fg = self.theme.get(fg_key) if fg_key else '#FFFFFF'
            btn_hover_bg = self.theme.get(hover_key)
            btn_hover_fg = btn_fg
            btn_pressed_bg = btn_hover_bg
            btn_border = self.theme.get('border')

            btn = tk.Button(parent, text=text,
                          bg=btn_bg,
                          fg=btn_fg,
                          font=btn_font,
                          relief="flat",
                          bd=0,
                          highlightthickness=1,
                          borderwidth=1,
                          height=btn_height,
                          command=command,
                          cursor="hand2",
                          padx=12,
                          pady=6)
            btn.config(highlightbackground=btn_border, highlightcolor=btn_border)
            btn.grid(row=row, column=col, padx=4, pady=4, sticky="ew")

            # Windows 11 style hover, press and release effects
            def on_enter(e):
                e.widget.config(bg=btn_hover_bg, fg=btn_hover_fg, relief="flat")

            def on_leave(e):
                e.widget.config(bg=btn_bg, fg=btn_fg, relief="flat")

            def on_press(e):
                e.widget.config(bg=btn_pressed_bg, relief="flat")

            def on_release(e):
                e.widget.config(bg=btn_hover_bg, relief="flat")

            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            btn.bind("<ButtonPress-1>", on_press)
            btn.bind("<ButtonRelease-1>", on_release)

            return btn

        # Main operation buttons: backup page, restore page, additional backups
        self.backup_btn = create_button(main_buttons_frame, "Backup Data", 'primary',
                                        self.start_backup_thread, 0, 0)

        self.restore_btn = create_button(restore_buttons_frame, "Restore Data", 'primary',
                                         self.start_restore_thread, 0, 0)

        self.filters_btn = create_button(main_buttons_frame, "Filters", 'neutral',
                                         self.open_filter_manager, 0, 1)

        self.driver_backup_btn = create_button(extra_buttons_frame, "Backup Drivers", 'neutral',
                                               self.start_driver_backup_thread, 0, 0)

        self.driver_restore_btn = create_button(restore_buttons_frame, "Restore Drivers", 'neutral',
                                                self.start_driver_restore_thread, 0, 1)
        self._extra_buttons_frame = extra_buttons_frame

        return create_button

    def _build_tools_section(self, main_frame, create_button):
        """Buttons of the tools page cards (U5-5), the additional backups (backup page) and
        the schedule button (schedule page)."""
        def buttons_frame(card, columns):
            frame = tk.Frame(self.page_bodies[card], bg=self.theme.get('bg_elevated'))
            frame.pack(fill="x")
            for i in range(columns):
                frame.columnconfigure(i, weight=1, uniform="button")
            return frame

        disk_frame = buttons_frame('tools_disk', 2)
        system_frame = buttons_frame('tools_system', 2)
        nas_frame = buttons_frame('tools_nas', 2)

        # Additional backups (backup page, U5-2)
        self.browser_profiles_btn = create_button(self._extra_buttons_frame, "Browser", 'neutral',
                                                  self.start_browser_profiles_backup_thread, 0, 1)

        self.winget_export_btn = create_button(self._extra_buttons_frame, "Export Apps", 'neutral',
                                               self.start_winget_export_thread, 0, 2)

        # Tools page: space and copy
        self.check_space_btn = create_button(disk_frame, "Check Space", 'neutral',
                                             self.check_space, 0, 0)

        self.copy_btn = create_button(disk_frame, "Copy Data", 'neutral',
                                      self.copy_files, 0, 1)

        # Tools page: system and files
        self.devmgr_btn = create_button(system_frame, "Device Mgr", 'neutral',
                                        self.open_device_manager, 0, 0)

        self.file_explorer_btn = create_button(system_frame, "Explorer", 'neutral',
                                               self.open_file_explorer, 0, 1)

        # Tools page: network drive
        self.nas_connect_btn = create_button(nas_frame, "Connect NAS", 'neutral',
                                             self.open_nas_connect_dialog, 0, 0)

        self.nas_disconnect_btn = create_button(nas_frame, "Disconnect", 'danger',
                                                self.open_nas_disconnect_dialog, 0, 1)

        # Scheduled backup page (U5-4)
        schedule_buttons_frame = tk.Frame(self.page_bodies['schedule'], bg=self.theme.get('bg_elevated'))
        schedule_buttons_frame.pack(fill="x")
        schedule_buttons_frame.columnconfigure(0, weight=1, uniform="button")
        schedule_buttons_frame.columnconfigure(1, weight=1, uniform="button")
        self.schedule_btn = create_button(schedule_buttons_frame, "Schedule", 'primary',
                                          self.schedule_backup, 0, 0)

    def _build_progress_log_section(self, main_frame, create_button):
        """Progress bar, status label, Activity Log (with save_log_btn, U5-6) and footer."""
        # Progress section
        progress_card = tk.Frame(self.bottom_bar, bg=self.theme.get('bg_elevated'), relief="flat", bd=0)
        progress_card.pack(fill="x", pady=(0, 12))

        progress_inner = tk.Frame(progress_card, bg=self.theme.get('bg_elevated'))
        progress_inner.pack(fill="x", padx=16, pady=12)

        self.progress_bar = ttk.Progressbar(progress_inner, orient="horizontal",
                                           mode="determinate",
                                           style="green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", expand=True, pady=(0, 8))

        self.status_label = tk.Label(progress_inner,
                                     text=self.translator.get('select_user_begin'),
                                     font=("Segoe UI", 10),
                                     background=self.theme.get('bg_elevated'),
                                     foreground=self.theme.get('fg_secondary'))
        self.status_label.pack()

        # Activity Log Card
        # One Activity Log card for the home and logs pages (D9): child of page_area,
        # packed into the shown page by show_page
        log_card = self.log_card = tk.Frame(self.page_area, bg=self.theme.get('bg_elevated'),
                                            relief="flat", bd=0)

        log_header = tk.Frame(log_card, bg=self.theme.get('bg_elevated'))
        log_header.pack(fill="x", padx=16, pady=(12, 8))

        self.log_title_label = tk.Label(log_header, text=self.translator.get('activity_log'),
                                        bg=self.theme.get('bg_elevated'),
                                        fg=self.theme.get('fg'),
                                        font=("Segoe UI", 12, "bold"))
        self.log_title_label.pack(side="left")

        # Save Log on the log card (U5-6): shown with the log on the home and logs pages
        log_buttons = tk.Frame(log_header, bg=self.theme.get('bg_elevated'))
        log_buttons.pack(side="right")
        self.save_log_btn = create_button(log_buttons, "Save Log", 'neutral', self.save_log, 0, 0)

        log_inner_frame = tk.Frame(log_card, bg=self.theme.get('bg_elevated'))
        log_inner_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        scrollbar = tk.Scrollbar(log_inner_frame, width=14,
                                bg=self.theme.get('bg_elevated'),
                                troughcolor=self.theme.get('bg'),
                                activebackground=self.theme.get('accent'))
        scrollbar.pack(side="right", fill="y")

        self.log_text = tk.Text(log_inner_frame, height=7,
                               bg=self.theme.get('bg_secondary'),
                               fg=self.theme.get('fg'),
                               relief="flat", bd=0,
                               font=("Consolas", 9),
                               state="disabled",
                               yscrollcommand=scrollbar.set,
                               padx=8, pady=6)
        self.log_text.tag_configure('bold', font=("Consolas", 9, "bold"))
        self.log_text.pack(fill="both", expand=True, side="left")

        scrollbar.config(command=self.log_text.yview)

        # Footer with theme colors - Minimal design
        footer_frame = tk.Frame(self.sidebar, bg=self.theme.get('bg'))
        footer_frame.pack(side="bottom", fill="x", padx=12, pady=12)

        shortcut_hint = tk.Label(footer_frame, text="Press F1 for keyboard shortcuts",
                                font=("Segoe UI", 8),
                                fg=self.theme.get('fg_secondary'),
                                bg=self.theme.get('bg'))
        shortcut_hint.pack(side="right")
