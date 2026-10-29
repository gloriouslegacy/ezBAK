"""Filter manager dialog (moved from ezBAK.py in P5-5).

Colours and fonts come from the design tokens for the theme's mode (U6-2, ADR 0004);
widgets, texts, keys and results are as before.

Tests patch `Win11Dialog` by name; the calls below go through the module-level
name so a patched class is used here too.
"""
import tkinter as tk

from ezbak.winsys.paths import resource_path

from .. import dpi
from ..shortcuts import DialogShortcuts
from ..theme import tk_font, token_colors
from ..widgets import create_win11_button
from .message import Win11Dialog


def _font(dlg, role):
    """Token font (theme.tk_font) in the dialog's language. A module function, not a
    method: the tests' recording replaces dialog methods while bindings run."""
    return tk_font(role, dlg.translator.current_lang)


class FilterManagerDialog(tk.Toplevel, DialogShortcuts):
    def __init__(self, master, current_filters=None):
        tk.Toplevel.__init__(self, master)
        self.parent = master
        self.theme = master.theme  # Use parent's theme
        self.translator = master.translator  # Use parent's translator
        self._colors = token_colors(self.theme)  # design tokens for the theme's mode (U6-2)
        self.title(self.translator.get('filter_manager'))
        try:
            self.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # titlebar icon is decoration only; a missing .ico must not block the dialog
        self.configure(bg=self._colors['bg'])
        self.geometry(dpi.geometry(self, 750, 520))  # px at 100 % (U8)
        self.transient(master)
        try:
            self.grab_set()
        except Exception:
            pass  # grab can fail while another window holds it; the dialog still works modeless

        self.result = None
        cf = current_filters or {'include': [], 'exclude': []}
        # deep copy simple
        self.inc = [dict(type=i.get('type'), pattern=i.get('pattern')) for i in (cf.get('include') or [])]
        self.exc = [dict(type=i.get('type'), pattern=i.get('pattern')) for i in (cf.get('exclude') or [])]

        self._create_filter_widgets()
        self._setup_filter_shortcuts()
        self._refresh()

    def _create_win11_button(self, parent, text, command, is_primary=False, width=12):
        """Create a Windows 11 style button for dialogs"""
        return create_win11_button(self.theme, parent, text, command, is_primary, width)

    def _setup_filter_shortcuts(self):
        """Setup shortcuts specific to filter dialog"""
        # Apply basic dialog shortcuts
        self.setup_dialog_shortcuts()

        # Filter-related shortcuts
        self.bind('<Control-n>', lambda e: self._add_rule('include'))  # New include rule
        self.bind('<Control-e>', lambda e: self._add_rule('exclude'))  # New exclude rule
        self.bind('<Delete>', lambda e: self._remove_selected_active())
        self.bind('<Control-s>', lambda e: self._save())
        self.bind('<Control-r>', lambda e: self._clear_active())  # Remove all from active list
        self.bind('<F1>', lambda e: self._show_filter_help())

        # Manage list focus
        self.bind('<Tab>', lambda e: self._cycle_focus())
        
    def _cycle_focus(self):
        """Move focus between lists with Tab key"""
        try:
            focused = self.focus_get()
            if focused == self.inc_list:
                self.exc_list.focus_set()
            else:
                self.inc_list.focus_set()
        except Exception:
            self.inc_list.focus_set()

    def _create_filter_widgets(self):
        """Create filter widgets (including shortcut notations)"""
        # Top description
        info_frame = tk.Frame(self, bg=self._colors['card'])
        info_frame.pack(fill='x', padx=12, pady=(12,6))

        info_text = f"{self.translator.get('filter_info_include')}\n{self.translator.get('filter_info_exclude')}"
        tk.Label(info_frame, text=info_text, bg=self._colors['card'], fg=self._colors['fg_muted'],
                font=_font(self, 'body'), justify='left').pack(anchor='w')

        # Layout
        top = tk.Frame(self, bg=self._colors['card'])
        top.pack(fill='both', expand=True, padx=12, pady=6)

        # Include panel (add shortcut notation)
        inc_frame = tk.LabelFrame(top, text=self.add_shortcut_text(self.translator.get('include_rules'), "Ctrl+N"),
                                 bg=self._colors['card'], fg=self._colors['fg'], font=_font(self, 'body_bold'))
        inc_frame.pack(side='left', fill='both', expand=True, padx=(0,6))

        self.inc_list = tk.Listbox(inc_frame, height=14, font=("Consolas", 9),
                                  bg=self._colors['bg'], fg=self._colors['fg'], selectbackground=self._colors['success'])
        self.inc_list.pack(fill='both', expand=True, padx=6, pady=6)

        btn_row_inc = tk.Frame(inc_frame, bg=self._colors['card'])
        btn_row_inc.pack(fill='x', padx=6, pady=(0,6))

        self._create_win11_button(btn_row_inc, self.translator.get('add'),
                                 lambda: self._add_rule('include'), width=8).pack(side='left', padx=2)
        self._create_win11_button(btn_row_inc, self.add_shortcut_text(self.translator.get('remove'), "Del"),
                                 lambda: self._remove_selected('include'), width=12).pack(side='left', padx=2)
        self._create_win11_button(btn_row_inc, self.translator.get('clear_all'),
                                 lambda: self._clear('include'), width=10).pack(side='left', padx=2)

        # Exclude panel (add shortcut notation)
        exc_frame = tk.LabelFrame(top, text=self.add_shortcut_text(self.translator.get('exclude_rules'), "Ctrl+E"),
                                 bg=self._colors['card'], fg=self._colors['fg'], font=_font(self, 'body_bold'))
        exc_frame.pack(side='left', fill='both', expand=True, padx=(6,0))

        self.exc_list = tk.Listbox(exc_frame, height=14, font=("Consolas", 9),
                                  bg=self._colors['bg'], fg=self._colors['fg'], selectbackground=self._colors['danger'])
        self.exc_list.pack(fill='both', expand=True, padx=6, pady=6)

        btn_row_exc = tk.Frame(exc_frame, bg=self._colors['card'])
        btn_row_exc.pack(fill='x', padx=6, pady=(0,6))

        self._create_win11_button(btn_row_exc, self.translator.get('add'),
                                 lambda: self._add_rule('exclude'), width=8).pack(side='left', padx=2)
        self._create_win11_button(btn_row_exc, self.add_shortcut_text(self.translator.get('remove'), "Del"),
                                 lambda: self._remove_selected('exclude'), width=12).pack(side='left', padx=2)
        self._create_win11_button(btn_row_exc, self.translator.get('clear_all'),
                                 lambda: self._clear('exclude'), width=10).pack(side='left', padx=2)

        # Bottom shortcut guide
        shortcut_frame = tk.Frame(self, bg=self._colors['card'])
        shortcut_frame.pack(fill='x', padx=12, pady=(6,0))

        shortcut_text = "F1: Help  •  Tab: Switch Lists  •  Del: Remove Selected  •  Ctrl+R: Clear Active List"
        tk.Label(shortcut_frame, text=shortcut_text, bg=self._colors['card'], fg=self._colors['fg_muted'],
                font=_font(self, 'caption')).pack(anchor='w')

        # Bottom actions (add shortcut notation)
        actions = tk.Frame(self, bg=self._colors['card'])
        actions.pack(fill='x', padx=12, pady=12)

        self._create_win11_button(actions, self.add_shortcut_text(self.translator.get('help'), "F1"),
                                 self._show_filter_help, width=10).pack(side='left')

        self._create_win11_button(actions, self.add_shortcut_text(self.translator.get('save'), "Ctrl+S"),
                                 self._save, is_primary=True, width=12).pack(side='right')
        self._create_win11_button(actions, self.add_shortcut_text(self.translator.get('dialog_cancel'), "Esc"),
                                 self._cancel, width=12).pack(side='right', padx=(0,8))

    def _rule_to_text(self, r):
        """Convert rule to display text"""
        try:
            rule_type = r.get('type', 'unknown')
            pattern = r.get('pattern', '')
            
            # Add icon by type
            type_icons = {
                'ext': '📄',
                'name': '📁', 
                'path': '🗂️'
            }
            icon = type_icons.get(rule_type, '❓')
            
            return f"{icon} [{rule_type.upper()}] {pattern}"
        except Exception:
            return str(r)

    def _refresh(self):
        """Refresh list"""
        # Include List
        self.inc_list.delete(0, tk.END)
        for r in self.inc:
            self.inc_list.insert(tk.END, self._rule_to_text(r))
        
        # Exclude List
        self.exc_list.delete(0, tk.END)
        for r in self.exc:
            self.exc_list.insert(tk.END, self._rule_to_text(r))

    def _remove_selected(self, kind):
        """Remove selected items"""
        if kind == 'include':
            idxs = list(self.inc_list.curselection())
            for i in reversed(idxs):
                if 0 <= i < len(self.inc):
                    del self.inc[i]
        else:
            idxs = list(self.exc_list.curselection())
            for i in reversed(idxs):
                if 0 <= i < len(self.exc):
                    del self.exc[i]
        self._refresh()

    def _remove_selected_active(self):
        """Remove selected items from currently active list"""
        try:
            # Check focused widget
            focused = self.focus_get()
            if focused == self.inc_list:
                self._remove_selected('include')
            elif focused == self.exc_list:
                self._remove_selected('exclude')
        except Exception:
            pass  # focus lookup can fail during teardown; Del then does nothing

    def _clear(self, kind):
        """Clear all"""
        if kind == 'include':
            self.inc = []
        else:
            self.exc = []
        self._refresh()

    def _clear_active(self):
        """Clear all from currently active list"""
        try:
            focused = self.focus_get()
            if focused == self.inc_list:
                self._clear('include')
            elif focused == self.exc_list:
                self._clear('exclude')
        except Exception:
            pass  # focus lookup can fail during teardown; Ctrl+R then does nothing

    def _add_rule(self, kind):
        """Add rule dialog"""
        dlg = self._open_rule_window(kind)
        vtype = self._build_rule_type(dlg)
        vpat, ent = self._build_rule_pattern(dlg, vtype)
        self._build_rule_buttons(dlg, kind, vtype, vpat)

        try:
            ent.focus_set()
        except Exception:
            pass  # initial focus is a convenience only

    def _open_rule_window(self, kind):
        """Create the add-rule Toplevel (first part of _add_rule)."""
        dlg = tk.Toplevel(self)
        dlg.title(self.translator.get('add_include_rule' if kind == 'include' else 'add_exclude_rule'))
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # titlebar icon is decoration only; a missing .ico must not block the dialog
        dlg.configure(bg=self._colors['card'])
        dlg.geometry(dpi.geometry(dlg, 480, 320))  # px at 100 % (U8)
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass  # grab can fail while another window holds it; the dialog still works modeless
        return dlg

    def _build_rule_type(self, dlg):
        """Rule type radio buttons of the add-rule dialog; returns the type variable."""
        # Select rule type
        tk.Label(dlg, text=self.translator.get('rule_type'), bg=self._colors['card'], fg=self._colors['fg'],
                font=_font(self, 'body_bold')).pack(anchor='w', padx=15, pady=(15,5))

        vtype = tk.StringVar(value='ext')
        tframe = tk.Frame(dlg, bg=self._colors['card'])
        tframe.pack(anchor='w', padx=15, fill='x')

        type_options = [
            ('ext', f'📄 {self.translator.get("extension")}', self.translator.get('extension_desc')),
            ('name', f'📁 {self.translator.get("file_folder_name")}', self.translator.get('name_desc')),
            ('path', f'🗂️ {self.translator.get("full_path")}', self.translator.get('path_desc'))
        ]

        for value, label, desc in type_options:
            rb_frame = tk.Frame(tframe, bg=self._colors['card'])
            rb_frame.pack(anchor='w', pady=2)

            tk.Radiobutton(rb_frame, text=label, variable=vtype, value=value,
                          bg=self._colors['card'], fg=self._colors['fg'], selectcolor=self._colors['card'],
                          font=_font(self, 'body_bold')).pack(side='left')
            tk.Label(rb_frame, text=f"  {desc}", bg=self._colors['card'], fg=self._colors['fg_muted'],
                    font=_font(self, 'caption')).pack(side='left')
        return vtype

    def _build_rule_pattern(self, dlg, vtype):
        """Pattern entry and example hint of the add-rule dialog; returns (variable, entry)."""
        # Pattern input
        tk.Label(dlg, text=self.translator.get('pattern'), bg=self._colors['card'], fg=self._colors['fg'],
                font=_font(self, 'body_bold')).pack(anchor='w', padx=15, pady=(15,5))

        vpat = tk.StringVar()
        pat_frame = tk.Frame(dlg, bg=self._colors['card'])
        pat_frame.pack(fill='x', padx=15)

        ent = tk.Entry(pat_frame, textvariable=vpat, width=50, font=_font(self, 'body'),
                      bg=self._colors['bg'], fg=self._colors['fg'], insertbackground=self._colors['fg'])
        ent.pack(fill='x')

        # Example text
        example_frame = tk.Frame(dlg, bg=self._colors['card'])
        example_frame.pack(fill='x', padx=15, pady=(5,0))

        def update_example(*args):
            try:
                rule_type = vtype.get()
                examples = {
                    'ext': self.translator.get('examples_ext'),
                    'name': self.translator.get('examples_name'),
                    'path': self.translator.get('examples_path')
                }
                example_label.config(text=examples.get(rule_type, ''))
            except Exception:
                pass  # example hint is decoration only

        example_label = tk.Label(example_frame, text="", bg=self._colors['card'], fg=self._colors['fg_muted'],
                               font=_font(self, 'caption'))
        example_label.pack(anchor='w')

        vtype.trace('w', update_example)
        update_example()  # Initial display
        return vpat, ent

    def _build_rule_buttons(self, dlg, kind, vtype, vpat):
        """Add/Cancel buttons and key bindings of the add-rule dialog."""
        # Buttons
        btns = tk.Frame(dlg, bg=self._colors['card'])
        btns.pack(fill='x', padx=15, pady=15)

        def _ok():
            p = vpat.get().strip()
            if not p:
                Win11Dialog.showwarning(self.translator.get('warning'), self.translator.get('please_enter_pattern'),
                                      parent=dlg, theme=self.theme, translator=self.translator)
                return

            rule_type = vtype.get().strip().lower()
            rule = {'type': rule_type, 'pattern': p}

            if kind == 'include':
                self.inc.append(rule)
            else:
                self.exc.append(rule)

            self._refresh()

            try:
                dlg.grab_release()
            except Exception:
                pass  # window may already have released its grab
            dlg.destroy()

        def _cancel():
            try:
                dlg.grab_release()
            except Exception:
                pass  # window may already have released its grab
            dlg.destroy()

        Win11Dialog._create_dialog_button(btns, self.translator.get('add_rule'),
                                          _ok, self.theme, is_primary=True).pack(side='right')
        Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_cancel'),
                                          _cancel, self.theme).pack(side='right', padx=(0,8))

        # Shortcuts setting
        dlg.bind('<Return>', lambda e: _ok())
        dlg.bind('<Escape>', lambda e: _cancel())

    def _show_filter_help(self):
        """Show filter help"""
        help_text = self.translator.get('filter_manager_help_content')

        help_dlg = tk.Toplevel(self)
        help_dlg.title(self.translator.get('filter_manager_help_title'))
        help_dlg.configure(bg=self._colors['bg'])
        help_dlg.geometry(dpi.geometry(help_dlg, 600, 610))  # px at 100 % (U8)
        help_dlg.transient(self)
        try:
            help_dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # titlebar icon is decoration only; a missing .ico must not block the dialog
        try:
            help_dlg.grab_set()
        except Exception:
            pass  # grab can fail while another window holds it; the help still works modeless

        # Scrollable text area
        text_frame = tk.Frame(help_dlg, bg=self._colors['bg'])
        text_frame.pack(fill="both", expand=True, padx=15, pady=15)

        scrollbar = tk.Scrollbar(text_frame, width=20,
                               bg=self._colors['border'],
                               troughcolor=self._colors['bg'],
                               activebackground=self._colors['accent'])
        scrollbar.pack(side="right", fill="y")

        text_widget = tk.Text(text_frame,
                             font=("Consolas", 9),
                             bg=self._colors['card'],
                             fg=self._colors['fg'],
                             relief="flat",
                             wrap="word",
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)

        scrollbar.configure(command=text_widget.yview)

        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        # Close Button
        btn_frame = tk.Frame(help_dlg, bg=self._colors['bg'])
        btn_frame.pack(pady=(0,15))
        Win11Dialog._create_dialog_button(btn_frame, self.translator.get('help_close_esc'),
                                          help_dlg.destroy, self.theme, is_primary=True).pack()

        help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
        help_dlg.focus_set()

    def _save(self):
        """Save filters"""
        self.result = {'include': self.inc, 'exclude': self.exc}
        try:
            self.grab_release()
        except Exception:
            pass  # window may already have released its grab
        self.destroy()

    def _cancel(self):
        """Cancel"""
        self.result = None
        try:
            self.grab_release()
        except Exception:
            pass  # window may already have released its grab
        self.destroy()

    def on_cancel(self):
        """Cancel handler for ESC key"""
        self._cancel()
        
    def on_ok(self):
        """OK handler for Enter key"""
        self._save()
