"""Source selection dialog for Copy Data (moved from ezBAK.py in P5-6).

Tests patch `Win11Dialog` and `filedialog.askdirectory`; the calls below go through
the module-level names so patched objects are used here too. The tree check/expand
helpers live in sources_tree.py.
Colours and fonts come from the design tokens for the theme's mode; the folder tree
(ttk.Treeview, no CTk equivalent) gets a ttk style in the same colours (U6-3, ADR 0004).
Widgets, texts, keys, check marks and the returned paths are as before.
"""
import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import font as tkfont

from ezbak.winsys.paths import resource_path

from .. import dpi
from ..shortcuts import DialogShortcuts
from ..theme import tk_font, token_colors
from ..widgets import create_win11_button
from .message import Win11Dialog
from .sources_tree import SourcesTreeMixin

_log = logging.getLogger(__name__)


def _tree_order(tree, item=""):
    """Item ids under `item` in tree order (depth-first, parents before children)."""
    for child in tree.get_children(item):
        yield child
        yield from _tree_order(tree, child)


TREE_STYLE = "Sources.Treeview"


def _style_tree(dlg):
    """ttk style of the folder tree in the token colours (CTk has no tree widget, U6-3).
    A module function, not a method: the tests' recording replaces dialog methods."""
    colors = dlg._colors
    style = ttk.Style(dlg)
    style.configure(TREE_STYLE, background=colors['bg'], fieldbackground=colors['bg'],
                    foreground=colors['fg'], bordercolor=colors['border'], font=tk_font('body'))
    style.map(TREE_STYLE, background=[('selected', colors['accent'])],
              foreground=[('selected', '#FFFFFF')])
    style.configure(TREE_STYLE + ".Heading", background=colors['card'], foreground=colors['fg'])
    if dpi.factor(dlg) != 1:
        # Row height is plain pixels; at 125 %+ the scaled font needs taller rows (U8)
        linespace = tkfont.Font(root=dlg, font=tk_font('body')).metrics('linespace')
        style.configure(TREE_STYLE, rowheight=max(dpi.px(dlg, 20), linespace + dpi.px(dlg, 4)))
    return TREE_STYLE


class SelectSourcesDialog(SourcesTreeMixin, tk.Toplevel, DialogShortcuts):
    def __init__(self, master, is_hidden_fn=None, title="Select Sources"):
        try:
            super().__init__(master)
            self.theme = master.theme  # Use parent's theme
            self._colors = token_colors(self.theme)  # design tokens for the theme's mode (U6-3)
            _log.debug(f"SelectSourcesDialog initialized with title: {title}")

            self.title(title)
            try:
                self.iconbitmap(resource_path('./icon/ezbak.ico'))
            except Exception:
                pass  # titlebar icon is decoration only; a missing .ico must not block the dialog
            self.configure(bg=self._colors['bg'])
            self.geometry(dpi.geometry(self, 780, 580))  # px at 100 % (U8)
            self.transient(master)
            try:
                self.grab_set()
            except Exception:
                pass  # grab can fail while another window holds it; the dialog still works modeless

            self.is_hidden_fn = is_hidden_fn or (lambda p: False)
            self.selected_paths = []

            self._checked = set()       # item ids that are checked
            self._item_path = {}        # item id -> full filesystem path

            _log.debug("Creating widgets...")
            self._create_source_widgets()
            
            _log.debug("Setting up shortcuts...")
            self._setup_source_shortcuts()
            
            _log.debug("Adding drive roots...")
            self._add_drive_roots()
            
            _log.debug("SelectSourcesDialog initialization complete")
            
        except Exception as e:
            _log.debug(f"SelectSourcesDialog init failed: {e}")
            import traceback
            _log.debug(f"Traceback: {traceback.format_exc()}")
            raise

    def _setup_source_shortcuts(self):
        """Set shortcuts specific to the source selection dialog"""
        try:
            # Apply basic dialog shortcuts
            self.setup_dialog_shortcuts()
            
            # Source selection related shortcuts
            self.bind('<Control-a>', lambda e: self._select_all())
            self.bind('<Control-n>', lambda e: self._clear_all())  # None (clear all)
            self.bind('<Control-r>', lambda e: self._choose_root())  # Root folder
            self.bind('<space>', lambda e: self._toggle_selected())
            self.bind('<Control-e>', lambda e: self._expand_selected())
            self.bind('<Control-c>', lambda e: self._collapse_selected())
            self.bind('<F5>', lambda e: self._refresh_tree())
            self.bind('<F1>', lambda e: self._show_source_help())
            
            _log.debug("Shortcuts setup complete")
        except Exception as e:
            _log.debug(f"_setup_source_shortcuts failed: {e}")

    def setup_dialog_shortcuts(self):
        """Set basic shortcuts for the dialog (implemented directly instead of via DialogShortcuts)"""
        try:
            # Basic dialog shortcuts
            self.bind('<Escape>', lambda e: self.on_cancel() if hasattr(self, 'on_cancel') else self.destroy())
            self.bind('<Return>', lambda e: self.on_ok() if hasattr(self, 'on_ok') else None)
            self.bind('<Alt-F4>', lambda e: self.destroy())
            
            # Set focus (to receive keyboard events)
            self.focus_set()
        except Exception as e:
            _log.debug(f"setup_dialog_shortcuts failed: {e}")
        
    def _create_win11_button(self, parent, text, command, is_primary=False, width=12):
        """Create a Windows 11 style button for dialogs"""
        return create_win11_button(self.theme, parent, text, command, is_primary, width)

    def _create_source_widgets(self):
        """Create source selection widgets"""
        try:
            _log.debug("Creating info frame...")
            # Top description
            info_frame = tk.Frame(self, bg=self._colors['card'])
            info_frame.pack(fill="x", padx=15, pady=(15, 10))

            tk.Label(info_frame, text="Select folders and files to copy",
                     bg=self._colors['card'], fg=self._colors['fg'], font=tk_font('section')).pack(anchor='w')

            tk.Label(info_frame, text="Use checkboxes or Space key to select items. Selected folders include all contents.",
                     bg=self._colors['card'], fg=self._colors['fg_muted'], font=tk_font('body')).pack(anchor='w', pady=(2,0))

            _log.debug("Creating top controls...")
            # Top controls
            top_controls = tk.Frame(self, bg=self._colors['card'])
            top_controls.pack(fill="x", padx=15, pady=(0, 10))

            self.add_root_btn = self._create_win11_button(top_controls,
                                                         self.add_shortcut_text("Add Root Folder...", "Ctrl+R"),
                                                         self._choose_root, width=20)
            self.add_root_btn.pack(side="left", padx=2)

            self.sel_all_btn = self._create_win11_button(top_controls,
                                                         self.add_shortcut_text("Select All", "Ctrl+A"),
                                                         self._select_all, width=15)
            self.sel_all_btn.pack(side="left", padx=2)

            self.clear_all_btn = self._create_win11_button(top_controls,
                                                           self.add_shortcut_text("Clear All", "Ctrl+N"),
                                                           self._clear_all, width=15)
            self.clear_all_btn.pack(side="left", padx=2)

            _log.debug("Creating tree area...")
            # Tree area
            tree_frame = tk.Frame(self, bg=self._colors['card'])
            tree_frame.pack(fill="both", expand=True, padx=15, pady=5)

            self.tree = ttk.Treeview(tree_frame, columns=("fullpath",), displaycolumns=(),
                                     style=_style_tree(self))
            vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self.tree.pack(side="left", fill="both", expand=True)

            self.tree.bind("<<TreeviewOpen>>", self._on_open)
            self.tree.bind("<Button-1>", self._on_click)

            _log.debug("Creating status frame...")
            # Bottom shortcut guide
            status_frame = tk.Frame(self, bg=self._colors['card'])
            status_frame.pack(fill='x', padx=15, pady=(5,0))

            self.status_label = tk.Label(status_frame, text="Ready - Use mouse or keyboard to select items",
                                        bg=self._colors['card'], fg=self._colors['fg_muted'], font=tk_font('body'))
            self.status_label.pack(side='left')

            _log.debug("Creating action buttons...")
            # Action buttons
            action_frame = tk.Frame(self, bg=self._colors['card'])
            action_frame.pack(fill="x", padx=15, pady=15)

            ok_btn = self._create_win11_button(action_frame,
                                              self.add_shortcut_text("OK", "Enter"),
                                              self._on_ok, is_primary=True, width=12)
            ok_btn.pack(side="right")

            cancel_btn = self._create_win11_button(action_frame,
                                                   self.add_shortcut_text("Cancel", "Esc"),
                                                   self._on_cancel, width=12)
            cancel_btn.pack(side="right", padx=(0, 8))
            
            _log.debug("Widget creation complete")
            
        except Exception as e:
            _log.debug(f"_create_source_widgets failed: {e}")
            import traceback
            _log.debug(f"Traceback: {traceback.format_exc()}")
            raise

    def _refresh_tree(self):
        """Refresh tree (F5)"""
        try:
            _log.debug("Refreshing tree...")
            
            # Initialize tree
            for item in self.tree.get_children(""):
                self.tree.delete(item)
            self._checked.clear()
            self._item_path.clear()
            
            # Re-add drives
            self._add_drive_roots()
            
            self._update_status("Tree refreshed")
            _log.debug("Tree refresh complete")
        except Exception as e:
            _log.debug(f"_refresh_tree failed: {e}")
            self._update_status(f"Refresh failed: {e}")

    def _update_status(self, message=None):
        """Update status bar"""
        try:
            if message:
                self.status_label.config(text=message)
            else:
                count = len(self._checked)
                if count == 0:
                    self.status_label.config(text="No items selected")
                elif count == 1:
                    self.status_label.config(text="1 item selected")
                else:
                    self.status_label.config(text=f"{count} items selected")
        except Exception as e:
            _log.debug(f"_update_status failed: {e}")

    def _show_source_help(self):
        """Show source selection help"""
        help_text = """Copy Data - Source Selection Help

KEYBOARD SHORTCUTS:
Ctrl + A         Select All Items
Ctrl + N         Clear All Selections  
Ctrl + R         Add Root Folder
Space            Toggle Selected Item
F5               Refresh Tree
F1               Show This Help
Enter            OK (Confirm Selection)
Esc              Cancel

SELECTION BEHAVIOR:
• [✓] Checked items will be included in the copy operation
• [ ] Unchecked items will be skipped
• Selecting a folder includes all its contents
• Only items without checked parents are included (avoids duplicates)

TIPS:
• Use "Add Root Folder" to browse for specific directories
• Check parent folders to include entire directory trees
• Use F5 to refresh if folder contents have changed
        """
        
        try:
            help_dlg = tk.Toplevel(self)
            help_dlg.title("Source Selection Help")
            help_dlg.configure(bg=self._colors['card'])
            help_dlg.geometry(dpi.geometry(help_dlg, 500, 400))  # px at 100 % (U8)
            help_dlg.transient(self)
            help_dlg.grab_set()

            text_widget = tk.Text(help_dlg,
                                 font=("Consolas", 10),
                                 bg=self._colors['bg'],
                                 fg=self._colors['fg'],
                                 relief="flat",
                                 wrap="word")
            text_widget.pack(fill="both", expand=True, padx=15, pady=15)

            text_widget.insert("1.0", help_text)
            text_widget.configure(state="disabled")

            btn_frame = tk.Frame(help_dlg, bg=self._colors['bg'])
            btn_frame.pack(pady=(0,15))
            close_btn = Win11Dialog._create_dialog_button(btn_frame,
                                                          "Close (Esc)",
                                                          help_dlg.destroy,
                                                          self.theme,
                                                          is_primary=True)
            close_btn.pack()
            
            help_dlg.bind('<Escape>', lambda e: help_dlg.destroy())
            help_dlg.focus_set()
            
        except Exception as e:
            _log.debug(f"_show_source_help failed: {e}")

    def _choose_root(self):
        """ Select root folder"""
        try:
            _log.debug("Choosing root folder...")
            path = filedialog.askdirectory(title="Select Root Folder", parent=self)
            if path:
                path = os.path.normpath(path)  # askdirectory returns '/'; keep '\' like the children (h-k)
                _log.debug(f"Selected path: {path}")
                item = self.tree.insert("", "end", text=self._format(path, False), open=False)
                self._item_path[item] = path
                self.tree.insert(item, "end", text="...")
                self._update_status(f"Added root: {os.path.basename(path)}")
                _log.debug("Root folder added successfully")
            else:
                _log.debug("No path selected")
        except Exception as e:
            _log.debug(f"_choose_root failed: {e}")
            self._update_status(f"Failed to add root: {e}")

    def _on_ok(self):
        """OK button"""
        try:
            _log.debug("_on_ok called")
            paths = []
            # Only include items that don't have a checked ancestor (avoid duplicates).
            # Walk the tree depth-first so the result follows the tree order (h-k).
            for item in _tree_order(self.tree):
                if item not in self._checked:
                    continue
                p = self._item_path.get(item)
                if not p:
                    continue
                # skip if any ancestor checked
                parent = self.tree.parent(item)
                skip = False
                while parent:
                    if parent in self._checked:
                        skip = True
                        break
                    parent = self.tree.parent(parent)
                if not skip:
                    paths.append(p)
            
            self.selected_paths = paths
            _log.debug(f"Selected {len(paths)} paths")
            for path in paths:
                _log.debug(f"Selected path: {path}")
            
            try:
                self.grab_release()
            except Exception:
                pass  # releasing a grab that is already gone is harmless; the dialog closes anyway
            self.destroy()
        except Exception as e:
            _log.debug(f"_on_ok failed: {e}")
            import traceback
            _log.debug(f"Traceback: {traceback.format_exc()}")

    def _on_cancel(self):
        """Cancel button"""
        try:
            _log.debug("_on_cancel called")
            self.selected_paths = []
            try:
                self.grab_release()
            except Exception:
                pass  # releasing a grab that is already gone is harmless; the dialog closes anyway
            self.destroy()
        except Exception as e:
            _log.debug(f"_on_cancel failed: {e}")

    def on_cancel(self):
        """Cancel handler for ESC key"""
        self._on_cancel()
        
    def on_ok(self):
        """OK handler for Enter key"""
        self._on_ok()
