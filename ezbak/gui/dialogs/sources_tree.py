"""Tree helpers of SelectSourcesDialog (moved from ezBAK.py in P5-6).

Mixed into SelectSourcesDialog: drive roots, lazy loading, check/expand toggles and
recursive check propagation. Uses `self.tree`, `self._checked`, `self._item_path`,
`self.is_hidden_fn` and `self._update_status` of the dialog.
"""
import logging
import os

_log = logging.getLogger(__name__)


class SourcesTreeMixin:
    def _format(self, name, checked):
        """Format as checkbox"""
        checkbox = "[✓]" if checked else "[ ]"
        return f"{checkbox} {name}"

    def _toggle_selected(self):
        """Toggle currently selected item (for Space key)"""
        try:
            item = self.tree.focus()
            if item:
                self._toggle_item(item)
                self._update_status()
        except Exception as e:
            _log.debug(f"_toggle_selected failed: {e}")

    def _expand_selected(self):
        """Expand selected item"""
        try:
            item = self.tree.focus()
            if item:
                self.tree.item(item, open=True)
        except Exception as e:
            _log.debug(f"_expand_selected failed: {e}")

    def _collapse_selected(self):
        """Collapse selected item"""
        try:
            item = self.tree.focus()
            if item:
                self.tree.item(item, open=False)
        except Exception as e:
            _log.debug(f"_collapse_selected failed: {e}")

    def _add_drive_roots(self):
        """Add drive roots to tree"""
        try:
            _log.debug("Adding drive roots...")
            drives = []
            
            # Detect drives with simple method
            for code in range(ord('C'), ord('Z') + 1):
                drive = f"{chr(code)}:\\"
                if os.path.exists(drive):
                    drives.append(drive)
            
            # Add drives to the tree
            for drive in sorted(drives):
                try:
                    item = self.tree.insert("", "end", text=self._format(drive, False), open=False)
                    self._item_path[item] = drive
                    # dummy child for lazy load
                    self.tree.insert(item, "end", text="...")
                    _log.debug(f"Added drive: {drive}")
                except Exception as e:
                    _log.debug(f"Failed to add drive {drive}: {e}")
                    
            _log.debug(f"Added {len(drives)} drives")
        except Exception as e:
            _log.debug(f"_add_drive_roots failed: {e}")

    def _on_open(self, event):
        """Tree item open event"""
        try:
            item = self.tree.focus()
            if not item:
                return
            children = self.tree.get_children(item)
            if len(children) == 1 and self.tree.item(children[0], "text") == "...":
                # populate children lazily
                self.tree.delete(children[0])
                path = self._item_path.get(item)
                parent_is_checked = item in self._checked

                if not path:
                    return
                
                entries = []
                with os.scandir(path) as it:
                    for entry in it:
                        full = entry.path
                        try:
                            if self.is_hidden_fn(full):
                                continue
                        except Exception:
                            # a failing hidden check shows the entry rather than hiding it
                            _log.debug("event=sources_hidden_check_failed path=%s", full, exc_info=True)
                        is_dir = False
                        try:
                            is_dir = entry.is_dir(follow_symlinks=False)
                        except Exception:
                            is_dir = False
                        entries.append((entry.name, is_dir, full))
                
                # sort: directories first, then files
                for name, is_dir, full in sorted(entries, key=lambda t: (not t[1], t[0].lower())):
                    child = self.tree.insert(item, "end", text=self._format(name, parent_is_checked), open=False)
                    self._item_path[child] = full
                    if parent_is_checked:
                        self._checked.add(child)
                    if is_dir:
                        self.tree.insert(child, "end", text="...")
        except Exception as e:
            _log.debug(f"_on_open failed: {e}")

    def _on_click(self, event):
        """Tree item click event"""
        try:
            item = self.tree.identify_row(event.y)
            if not item:
                return
            try:
                elem = self.tree.identify("element", event.x, event.y)
            except Exception:
                elem = ""
            elem_str = str(elem).lower()
            # Only toggle when clicking on the text area
            if "text" not in elem_str:
                return
            self._toggle_item(item)
            self._update_status()
        except Exception as e:
            _log.debug(f"_on_click failed: {e}")

    def _toggle_item(self, item):
        """Toggle item check state"""
        try:
            text = self.tree.item(item, "text")
            if not text:
                return
            checked = item in self._checked
            new_checked = not checked
            
            # Extract only name part from text (remove checkbox part)
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]  # "[✓] " or "[ ] " remove
            else:
                name = text
            
            self.tree.item(item, text=self._format(name, new_checked))
            
            if new_checked:
                self._checked.add(item)
            else:
                self._checked.discard(item)
                
            # propagate to descendants
            for child in self.tree.get_children(item):
                if self.tree.item(child, "text") == "...":
                    continue
                self._apply_to_descendants(child, new_checked)
        except Exception as e:
            _log.debug(f"_toggle_item failed: {e}")

    def _apply_to_descendants(self, item, checked):
        """Apply check state to descendants"""
        try:
            text = self.tree.item(item, "text")
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]
            else:
                name = text
            
            self.tree.item(item, text=self._format(name, checked))
            
            if checked:
                self._checked.add(item)
            else:
                self._checked.discard(item)
                
            for child in self.tree.get_children(item):
                if self.tree.item(child, "text") == "...":
                    continue
                self._apply_to_descendants(child, checked)
        except Exception as e:
            _log.debug(f"_apply_to_descendants failed: {e}")

    def _select_all(self):
        """Select all items"""
        try:
            for item in self.tree.get_children(""):
                self._set_checked_recursive(item, True)
            self._update_status()
        except Exception as e:
            _log.debug(f"_select_all failed: {e}")

    def _clear_all(self):
        """Clear all selections"""
        try:
            for item in self.tree.get_children(""):
                self._set_checked_recursive(item, False)
            self._update_status()
        except Exception as e:
            _log.debug(f"_clear_all failed: {e}")

    def _set_checked_recursive(self, item, checked):
        """ Set check state recursively"""
        try:
            text = self.tree.item(item, "text")
            if text.startswith("[✓]") or text.startswith("[ ]"):
                name = text[4:]
            else:
                name = text
            
            self.tree.item(item, text=self._format(name, checked))
            
            if checked:
                self._checked.add(item)
            else:
                self._checked.discard(item)
                
            for child in self.tree.get_children(item):
                if self.tree.item(child, "text") == "...":
                    continue
                self._set_checked_recursive(child, checked)
        except Exception as e:
            _log.debug(f"_set_checked_recursive failed: {e}")
