"""Keyboard shortcuts for the main window and dialogs (moved from ezBAK.py in P5-4).

The odd bindings (e.g. <Control-1>..<Control-9>) are kept as is; they are
recorded as defect candidates in docs/roadmap.md.
"""
import logging
import tkinter as tk

from ezbak.winsys.paths import resource_path

from . import dpi
from .dialogs import Win11Dialog
from .theme import token_colors

_log = logging.getLogger(__name__)


class DialogShortcuts:
    """Common shortcut key functions for dialogs"""
    
    def setup_dialog_shortcuts(self):
        """Common Shortcut Key Features for Dialogs"""
        # Basic dialog shortcuts
        self.bind('<Escape>', lambda e: self.on_cancel() if hasattr(self, 'on_cancel') else self.destroy())
        self.bind('<Return>', lambda e: self.on_ok() if hasattr(self, 'on_ok') else None)
        self.bind('<Alt-F4>', lambda e: self.destroy())
        
        # Set focus (to receive keyboard events)
        self.focus_set()
        
    def add_shortcut_text(self, text, shortcut):
        """Add shortcut notation to text"""
        return f"{text} ({shortcut})"

class KeyboardShortcuts:
    """Keyboard shortcuts manager class for ezBAK"""
    
    def __init__(self, app):
        self.app = app
        self.setup_shortcuts()
        self.setup_help_dialog()
    
    def setup_shortcuts(self):
        """Register keyboard shortcuts"""
        
        # Main feature shortcuts
        self.app.bind('<Control-b>', lambda e: self.safe_execute(self.app.start_backup_thread))
        self.app.bind('<Control-r>', lambda e: self.safe_execute(self.app.start_restore_thread))
        self.app.bind('<Control-d>', lambda e: self.safe_execute(self.app.start_driver_backup_thread))
        self.app.bind('<Control-Shift-D>', lambda e: self.safe_execute(self.app.start_driver_restore_thread))
        
        # File management
        self.app.bind('<Control-c>', lambda e: self.safe_execute(self.app.copy_files))
        self.app.bind('<Control-o>', lambda e: self.safe_execute(self.app.open_file_explorer))
        self.app.bind('<Control-s>', lambda e: self.safe_execute(self.app.save_log))
        
        # Tools and utilities
        self.app.bind('<Control-m>', lambda e: self.safe_execute(self.app.open_device_manager))
        self.app.bind('<Control-k>', lambda e: self.safe_execute(self.app.check_space))
        self.app.bind('<Control-f>', lambda e: self.safe_execute(self.app.open_filter_manager))
        self.app.bind('<Control-t>', lambda e: self.safe_execute(self.app.schedule_backup))
        
        # Browser and app related
        self.app.bind('<Control-p>', lambda e: self.safe_execute(self.app.start_browser_profiles_backup_thread))
        self.app.bind('<Control-w>', lambda e: self.safe_execute(self.app.start_winget_export_thread))
        
        # NAS connection
        self.app.bind('<Control-n>', lambda e: self.safe_execute(self.app.open_nas_connect_dialog))
        self.app.bind('<Control-Shift-N>', lambda e: self.safe_execute(self.app.open_nas_disconnect_dialog))
        
        # General shortcuts
        self.app.bind('<F1>', lambda e: self.show_help())
        self.app.bind('<F5>', lambda e: self.refresh_user_list())
        self.app.bind('<Escape>', lambda e: self.cancel_operation())
        self.app.bind('<Control-q>', lambda e: self.safe_exit())
        
        # User selection shortcuts (number keys)
        for i in range(1, 10):
            self.app.bind(f'<Control-Key-{i}>', lambda e, idx=i-1: self.select_user_by_index(idx))
        
        # Focus shortcuts
        self.app.focus_set()  
    
    def safe_execute(self, function):
        """Execute function (with error handling)"""
        try:
            if callable(function):
                function()
            else:
                self.app.message_queue.put(('log', f"Function {function} is not callable"))
        except Exception as e:
            self.app.message_queue.put(('log', f"Shortcut execution error: {e}"))
            _log.debug(f"Shortcut error: {e}")
    
    def show_help(self):
        """Show shortcuts help (F1)"""
        help_text = self.app.translator.get('keyboard_shortcuts_content')
        colors = token_colors(self.app.theme)  # design tokens for the theme's mode (U6-1)

        # Custom help dialog
        help_dialog = tk.Toplevel(self.app)
        help_dialog.title(self.app.translator.get('keyboard_shortcuts_title'))
        help_dialog.configure(bg=colors['bg'])
        help_dialog.geometry(dpi.geometry(help_dialog, 449, 748))  # px at 100 % (U8)
        help_dialog.transient(self.app)
        help_dialog.grab_set()

        try:
            help_dialog.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # widget decoration: a missing icon must not block the help dialog

        # Text area
        text_frame = tk.Frame(help_dialog, bg=colors['bg'])
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create the scrollbar first
        scrollbar = tk.Scrollbar(text_frame,
                                width=20,
                                bg=colors['border'],
                                troughcolor=colors['bg'],
                                activebackground=colors['accent'],
                                relief="flat")
        scrollbar.pack(side="right", fill="y")

        text_widget = tk.Text(text_frame,
                             font=("Consolas", 10),
                             bg=colors['card'],
                             fg=colors['fg'],
                             relief="flat",
                             wrap="word",
                             yscrollcommand=scrollbar.set)
        text_widget.pack(side="left", fill="both", expand=True)

        # Connect the scrollbar and the text widget
        scrollbar.configure(command=text_widget.yview)

        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        # Close button
        btn_frame = tk.Frame(help_dialog, bg=colors['bg'])
        btn_frame.pack(pady=10)
        close_btn = Win11Dialog._create_dialog_button(btn_frame,
                                                      self.app.translator.get('help_exit_esc'),
                                                      help_dialog.destroy,
                                                      self.app.theme,
                                                      is_primary=True)
        close_btn.pack()

        # Close with ESC
        help_dialog.bind('<Escape>', lambda e: help_dialog.destroy())
        help_dialog.focus_set()
    
    def refresh_user_list(self):
        """Refresh user list (F5)"""
        try:
            self.app.load_users()
            self.app.message_queue.put(('log', "Refreshed user list."))
        except Exception as e:
            self.app.message_queue.put(('log', f"Error refreshing user list: {e}"))

    def select_user_by_index(self, index):
        """Quick select user by number key (Ctrl + 1~9)"""
        try:
            users = self.app.user_combo['values']
            if 0 <= index < len(users):
                self.app.user_combo.current(index)
                self.app.message_queue.put(('log', f"Selected user: {users[index]}"))
            else:
                self.app.message_queue.put(('log', f"User index {index+1} is out of range."))
        except Exception as e:
            self.app.message_queue.put(('log', f"Error selecting user: {e}"))

    def cancel_operation(self):
        """Cancel current operation (ESC)"""
        try:
            # Attempt to cancel if an operation is in progress
            if hasattr(self.app, 'progress_bar') and self.app.progress_bar['mode'] == 'indeterminate':
                self.app.message_queue.put(('log', "Cancellation requested"))
                # The actual cancellation logic needs to be implemented in each task thread
        except Exception as e:
            _log.debug(f"Cancel operation error: {e}")
    
    def safe_exit(self):
        """Exit program  (Ctrl + Q)"""
        try:
            if Win11Dialog.askyesno(self.app.translator.get('confirm_exit'),
                                   self.app.translator.get('confirm_exit_msg'),
                                   parent=self.app, theme=self.app.theme, translator=self.app.translator):
                try:
                    self.app.save_settings()
                    self.app.close_log_file()
                except Exception:
                    # exit continues even if saving fails (as before)
                    _log.debug("exit-time save failed", exc_info=True)
                self.app.quit()
                self.app.destroy()
        except Exception as e:
            _log.debug(f"Safe exit error: {e}")
    
    def setup_help_dialog(self):
        """Display shortcuts hint in status bar or menu"""
        try:
            if hasattr(self.app, 'status_label'):
                original_text = self.app.status_label.cget("text")
                hint_text = f"{original_text}"
                self.app.status_label.configure(text=hint_text)
        except Exception:
            pass  # widget decoration: the status bar hint is optional
