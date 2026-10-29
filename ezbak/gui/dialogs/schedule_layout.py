"""Widget layout of ScheduleBackupDialog (moved from ezBAK.py in P5-7).

Mixed into ScheduleBackupDialog. `_create_widgets` (210 lines before) runs the same
statements in the same order, split into two layout sections: task name, destination,
schedule and time first, then attributes, retention, filters and buttons.
Colours (self._colors) and fonts from the design tokens (U6-4).
"""
import logging
import tkinter as tk
from tkinter import ttk

from ..theme import tk_font

_log = logging.getLogger(__name__)


class ScheduleLayoutMixin:
    def _create_widgets(self):
        """Create UI widgets with Windows 11 styling"""
        try:
            # Safely get user name
            user_name = self._parent_user_name()

            # Modern styling
            label_font = tk_font('body_bold')
            entry_font = tk_font('body')

            row = self._layout_schedule_top(user_name, label_font, entry_font)
            self._layout_schedule_bottom(row, label_font, entry_font)

        except Exception as widget_error:
            _log.debug(f"Widget creation failed: {widget_error}")
            raise widget_error

    def _layout_schedule_top(self, user_name, label_font, entry_font):
        """Task name, destination, schedule and time rows; returns the last row."""
        # Task Name
        row = 0
        tk.Label(self, text="Task Name", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )
        self.task_var = tk.StringVar(value=f"ezBAK_Backup_{user_name}")
        tk.Entry(self, textvariable=self.task_var, width=40, font=entry_font,
                bg=self._colors['card'], fg=self._colors['fg'],
                relief="solid", bd=1).grid(
            row=row, column=1, columnspan=2, sticky="ew", padx=15, pady=8
        )

        # Destination Folder
        row += 1
        tk.Label(self, text="Destination Folder", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )
        self.dest_var = tk.StringVar()
        tk.Entry(self, textvariable=self.dest_var, width=30, font=entry_font,
                bg=self._colors['card'], fg=self._colors['fg'],
                relief="solid", bd=1).grid(
            row=row, column=1, sticky="ew", padx=15, pady=8
        )
        tk.Button(
            self, text=self.add_shortcut_text("Browse...", "Ctrl+B"),
            command=self._browse_destination,
            bg=self._colors['card'], fg=self._colors['fg'],
            font=tk_font('body'), relief="flat", bd=0,
            cursor="hand2"
        ).grid(row=row, column=2, padx=15, pady=8, sticky="ew")

        # Schedule
        row += 1
        tk.Label(self, text=self.add_shortcut_text("Schedule", "Alt+1/2/3"),
                bg=self._colors['bg'], fg=self._colors['fg'],
                font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )

        schedule_frame = tk.Frame(self, bg=self._colors['bg'])
        schedule_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=15, pady=8)

        self.schedule_var = tk.StringVar(value="Daily")
        schedule_combo = ttk.Combobox(
            schedule_frame, textvariable=self.schedule_var,
            values=["Daily", "Weekly", "Monthly"], state="readonly",
            width=18, font=entry_font
        )
        schedule_combo.pack(side="left")

        # Day selector for Monthly schedule (1-30)
        tk.Label(schedule_frame, text="  Day:", bg=self._colors['bg'],
                fg=self._colors['fg'], font=tk_font('body')).pack(side="left", padx=(10, 5))

        self.monthly_day_var = tk.StringVar(value="1")
        self.monthly_day_combo = ttk.Combobox(
            schedule_frame, textvariable=self.monthly_day_var,
            values=[str(i) for i in range(1, 31)],
            state="readonly",
            width=4, font=entry_font
        )
        self.monthly_day_combo.pack(side="left")

        # Update visibility based on schedule type
        def on_schedule_change(*args):
            if self.schedule_var.get() == "Monthly":
                self.monthly_day_combo.config(state="readonly")
            else:
                self.monthly_day_combo.config(state="disabled")

        self.schedule_var.trace_add('write', on_schedule_change)
        on_schedule_change()  # Initialize visibility

        # Time
        row += 1
        tk.Label(self, text="Time", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )

        time_frame = tk.Frame(self, bg=self._colors['bg'])
        time_frame.grid(row=row, column=1, sticky="w", padx=15, pady=8)

        # Hour combobox (00-23)
        self.hour_var = tk.StringVar(value="02")
        hour_combo = ttk.Combobox(
            time_frame, textvariable=self.hour_var,
            values=[f"{i:02d}" for i in range(24)],
            state="readonly",
            width=4, font=entry_font
        )
        hour_combo.pack(side="left")

        tk.Label(time_frame, text=":", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).pack(side="left", padx=5)

        # Minute combobox (00-59)
        self.minute_var = tk.StringVar(value="00")
        minute_combo = ttk.Combobox(
            time_frame, textvariable=self.minute_var,
            values=[f"{i:02d}" for i in range(60)],
            state="readonly",
            width=4, font=entry_font
        )
        minute_combo.pack(side="left")
        return row

    def _layout_schedule_bottom(self, row, label_font, entry_font):
        """Attributes, retention, filters and button rows after the time row."""
        # Attributes
        row += 1
        tk.Label(self, text="Attributes", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )

        attr_frame = tk.Frame(self, bg=self._colors['bg'])
        attr_frame.grid(row=row, column=1, columnspan=2, sticky="w", padx=15, pady=8)

        self.hidden_var = tk.BooleanVar(value=False)
        self.system_var = tk.BooleanVar(value=False)

        tk.Checkbutton(attr_frame, text="Include Hidden", variable=self.hidden_var,
                      bg=self._colors['bg'], fg=self._colors['fg'],
                      selectcolor=self._colors['bg'],
                      font=tk_font('body'), relief="flat", bd=0).pack(side="left")

        tk.Checkbutton(attr_frame, text="Include System", variable=self.system_var,
                      bg=self._colors['bg'], fg=self._colors['fg'],
                      selectcolor=self._colors['bg'],
                      font=tk_font('body'), relief="flat", bd=0).pack(side="left", padx=(15, 0))

        # Backups to Keep
        row += 1
        tk.Label(self, text="Backups to Keep (0=all):", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )
        self.retention_count_var = tk.StringVar(value="2")
        tk.Spinbox(self, from_=0, to=99, textvariable=self.retention_count_var,
                  width=10, font=entry_font,
                  bg=self._colors['card'], fg=self._colors['fg'],
                  buttonbackground=self._colors['accent'],
                  relief="solid", bd=1).grid(
            row=row, column=1, sticky="w", padx=15, pady=8
        )

        # Logs to Keep
        row += 1
        tk.Label(self, text="Logs to Keep (days, 0=off):", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )
        self.log_retention_days_var = tk.StringVar(value="30")
        tk.Spinbox(self, from_=0, to=365, textvariable=self.log_retention_days_var,
                  width=10, font=entry_font,
                  bg=self._colors['card'], fg=self._colors['fg'],
                  buttonbackground=self._colors['accent'],
                  relief="solid", bd=1).grid(
            row=row, column=1, sticky="w", padx=15, pady=8
        )

        # Filters
        row += 1
        tk.Label(self, text="Filters:", bg=self._colors['bg'],
                fg=self._colors['fg'], font=label_font).grid(
            row=row, column=0, sticky="w", padx=15, pady=8
        )
        tk.Button(
            self, text=self.add_shortcut_text("Manage Filters...", "Ctrl+F"),
            command=self.open_filter_manager,
            bg=self._colors['card'], fg=self._colors['fg'],
            font=tk_font('body'), relief="flat", bd=0,
            cursor="hand2"
        ).grid(row=row, column=1, sticky="w", padx=15, pady=8)

        # Button area with modern styling
        row += 1
        btn_frame = tk.Frame(self, bg=self._colors['bg'])
        btn_frame.grid(row=row, column=0, columnspan=3, sticky="e", padx=15, pady=20)

        # Windows 11 style buttons
        self._create_win11_button(btn_frame, self.add_shortcut_text("Create", "Ctrl+C"),
                                 self.on_create, is_primary=True).pack(side="right", padx=5)
        self._create_win11_button(btn_frame, self.add_shortcut_text("Delete", "Ctrl+D"),
                                 self.on_delete).pack(side="right", padx=5)
        self._create_win11_button(btn_frame, self.add_shortcut_text("Close", "Esc"),
                                 self.on_close).pack(side="right", padx=5)

        # Configure grid weights
        self.columnconfigure(1, weight=1)
