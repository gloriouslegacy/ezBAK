"""Custom Tk widgets (moved from ezBAK.py in P5-2).

create_win11_button is the one copy of the dialogs' _create_win11_button (P5-5).
"""
import tkinter as tk


class ToggleSwitch(tk.Canvas):
    """Modern toggle switch widget with Windows 11 style"""

    def __init__(self, parent, width=50, height=24, variable=None, command=None, theme=None, **kwargs):
        """
        Initialize toggle switch

        Args:
            parent: Parent widget
            width: Switch width (default 50)
            height: Switch height (default 24)
            variable: BooleanVar to bind to
            command: Callback function when toggled
            theme: Win11Theme instance for colors
        """
        super().__init__(parent, width=width, height=height,
                        bg=theme.get('bg_elevated') if theme else '#f5f5f5',
                        highlightthickness=0, **kwargs)

        self.width = width
        self.height = height
        self.variable = variable
        self.command = command
        self.theme = theme

        # Track initial state
        self.is_on = variable.get() if variable else False

        # Colors
        self.update_colors()

        # Draw switch
        self.draw_switch()

        # Bind click event
        self.bind("<Button-1>", self.toggle)

        # Bind variable trace if provided
        if self.variable:
            self.variable.trace_add('write', self._on_variable_change)

    def update_colors(self):
        """Update colors based on theme"""
        if self.theme:
            if self.theme.current_mode == 'dark':
                self.color_on = self.theme.get('accent')  # Blue when on
                self.color_off = '#484f58'  # Gray when off
                self.knob_color = '#ffffff'  # White knob
            else:
                self.color_on = self.theme.get('accent')  # Blue when on
                self.color_off = '#c8c6c4'  # Gray when off
                self.knob_color = '#ffffff'  # White knob
        else:
            self.color_on = '#0078D4'
            self.color_off = '#c8c6c4'
            self.knob_color = '#ffffff'

    def draw_switch(self):
        """Draw the toggle switch"""
        self.delete("all")

        # Calculate positions
        radius = self.height / 2
        knob_radius = radius - 3

        # Draw background track
        bg_color = self.color_on if self.is_on else self.color_off

        # Draw rounded rectangle for track
        self.create_oval(0, 0, self.height, self.height,
                        fill=bg_color, outline="")
        self.create_oval(self.width - self.height, 0,
                        self.width, self.height,
                        fill=bg_color, outline="")
        self.create_rectangle(radius, 0, self.width - radius, self.height,
                             fill=bg_color, outline="")

        # Draw knob
        if self.is_on:
            knob_x = self.width - radius
        else:
            knob_x = radius

        self.create_oval(knob_x - knob_radius, radius - knob_radius,
                        knob_x + knob_radius, radius + knob_radius,
                        fill=self.knob_color, outline="")

    def toggle(self, event=None):
        """Toggle the switch"""
        self.is_on = not self.is_on

        # Update variable if provided
        if self.variable:
            self.variable.set(self.is_on)

        # Redraw
        self.draw_switch()

        # Call command if provided
        if self.command:
            self.command()

    def _on_variable_change(self, *args):
        """Called when the bound variable changes"""
        if self.variable:
            new_state = self.variable.get()
            if new_state != self.is_on:
                self.is_on = new_state
                self.draw_switch()

    def update_theme(self, theme):
        """Update theme colors"""
        self.theme = theme
        self.update_colors()
        self.config(bg=theme.get('bg_elevated'))
        self.draw_switch()


def create_win11_button(theme, parent, text, command, is_primary=False, width=12):
    """Create a Windows 11 style button for dialogs (shared by the dialogs, P5-5 D5).
    Colours and font from the design tokens for the theme's mode (U6-1)."""
    from .theme import tk_font, token_colors  # at call time: theme imports dialogs, which import this module
    colors = token_colors(theme)
    if is_primary:
        btn_bg = colors['accent']
        btn_fg = '#FFFFFF'
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
                   command=command,
                   bg=btn_bg,
                   fg=btn_fg,
                   font=tk_font('body'),
                   relief="flat",
                   bd=0,
                   highlightthickness=1,
                   highlightbackground=border_color,
                   highlightcolor=border_color,
                   borderwidth=1,
                   width=width,
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
