"""Restore question dialogs of App (moved from ezBAK.py in P5-13).

RestoreSelectMixin is mixed into App before ctk.CTk. Both methods run on the Tk thread
(QueueMixin hands them 'select_folder' / 'show_cross_user_warning') and leave the
answer in the App attribute the work thread polls (_folder_selection_result,
_cross_user_confirmed; P5 decision D4). Module-level names (tk, resource_path,
Win11Dialog) match ezBAK.py so facade patches reach this module too.
Colours and fonts from the design tokens for the theme's mode (U6-5, ADR 0004).
"""
import tkinter as tk

from ezbak.winsys.paths import resource_path

from .. import dpi
from ..theme import tk_font, token_colors
from .message import Win11Dialog


def _c(app, name):
    """Design token colour for the App's theme mode (U6-5). A module function, not an
    App method: the tests' recording replaces App methods."""
    return token_colors(app.theme)[name]


class RestoreSelectMixin:
    """show_cross_user_warning_dialog and show_folder_selection_dialog of App."""

    def show_cross_user_warning_dialog(self, backup_user, target_user, folder_name):
        """Show warning dialog for cross-user restore"""
        message = self.translator.get('cross_user_restore_warning').format(backup_user, target_user, folder_name)
        response = Win11Dialog.askyesno(
            self.translator.get('warning'),
            message,
            parent=self,
            theme=self.theme,
            translator=self.translator
        )
        self._cross_user_confirmed = response

    def show_folder_selection_dialog(self, backup_folders):
        """Show dialog to select backup folder from list"""
        dialog = tk.Toplevel(self)
        dialog.title(self.translator.get('select_backup_folder'))

        # Set titlebar icon
        try:
            dialog.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # titlebar icon is cosmetic (missing .ico or Tk without icon support)

        dialog.transient(self)
        dialog.grab_set()

        # Center dialog
        dialog.update_idletasks()
        width = 500
        height = 400
        dialog.geometry(dpi.centered_geometry(dialog, width, height))  # px at 100 % (U8)

        # Apply theme
        dialog.configure(bg=_c(self, 'bg'))

        # Title label
        title_label = tk.Label(
            dialog,
            text=self.translator.get('select_backup_folder_msg'),
            font=tk_font('section', self.translator.current_lang),
            bg=_c(self, 'bg'),
            fg=_c(self, 'fg')
        )
        title_label.pack(pady=(20, 10))

        # Listbox frame
        list_frame = tk.Frame(dialog, bg=_c(self, 'bg'))
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        # Listbox
        listbox = tk.Listbox(
            list_frame,
            font=("Consolas", 9),
            bg=_c(self, 'card'),
            fg=_c(self, 'fg'),
            selectbackground=_c(self, 'accent'),
            selectforeground="white",
            yscrollcommand=scrollbar.set,
            height=15
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        # Add backup folders to listbox
        for folder in backup_folders:
            listbox.insert(tk.END, folder)

        # Select first item by default
        if backup_folders:
            listbox.selection_set(0)
            listbox.activate(0)

        selected_folder = [None]

        def on_ok():
            selection = listbox.curselection()
            if selection:
                selected_folder[0] = backup_folders[selection[0]]
            dialog.destroy()

        def on_cancel():
            selected_folder[0] = None
            dialog.destroy()

        def on_double_click(event):
            on_ok()

        listbox.bind('<Double-Button-1>', on_double_click)

        # Buttons frame
        btn_frame = tk.Frame(dialog, bg=_c(self, 'bg'))
        btn_frame.pack(pady=(0, 20))

        ok_btn = tk.Button(
            btn_frame,
            text=self.translator.get('dialog_ok'),
            command=on_ok,
            font=tk_font('body', self.translator.current_lang),
            bg=_c(self, 'accent'),
            fg="white",
            width=10,
            relief="flat",
            cursor="hand2"
        )
        ok_btn.pack(side="left", padx=5)

        cancel_btn = tk.Button(
            btn_frame,
            text=self.translator.get('dialog_cancel'),
            command=on_cancel,
            font=tk_font('body', self.translator.current_lang),
            bg=_c(self, 'card'),
            fg=_c(self, 'fg'),
            width=10,
            relief="flat",
            cursor="hand2"
        )
        cancel_btn.pack(side="left", padx=5)

        # Wait for dialog to close
        dialog.wait_window()

        # Set result
        self._folder_selection_result = selected_folder[0]
