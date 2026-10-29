"""Translation system (moved from ezBAK.py in P5-1).

LanguageMixin (P5-9) holds App's update_ui_texts and switch_language; it is mixed
into App before tk.Tk. `dialogs` is imported as a module because dialogs/message.py
imports this module: Win11Dialog is looked up when the method runs.
"""
import logging

from . import dialogs
from .strings_en import EN
from .strings_ko import KO

_log = logging.getLogger(__name__)


class Translator:
    """Simple translation system for multi-language support"""

    TRANSLATIONS = {
        'en': EN,
        'ko': KO,
    }

    def __init__(self, initial_lang='en'):
        self.current_lang = initial_lang

    def set_language(self, lang_code):
        """Set current language"""
        if lang_code in self.TRANSLATIONS:
            self.current_lang = lang_code
            return True
        return False

    def get(self, key, default=None):
        """Get translated text"""
        trans = self.TRANSLATIONS.get(self.current_lang, {})
        return trans.get(key, default or key)


class LanguageMixin:
    def update_ui_texts(self):
        """Update all UI text labels with current language"""
        try:
            # Update window title
            self.title(self.translator.get('app_title'))

            # Check if Korean language (no emojis for Korean)
            is_korean = self.translator.current_lang == 'ko'

            self._update_button_texts()
            if hasattr(self, 'nav_items'):
                self._update_nav_texts()

            # Update status label
            if hasattr(self, 'status_label'):
                current_status = self.status_label.cget('text')
                if 'Select User' in current_status or '사용자 선택' in current_status:
                    self.status_label.config(text=self.translator.get('select_user_begin'))

            # Update sound label (no emojis for all languages)
            if hasattr(self, 'sound_label'):
                sound_text = self.translator.get('sound')
                self.sound_label.config(text=sound_text)

            # Note: Activity Log notice is only set once during initialization (lines 1183-1196)
            # and should not be updated when language changes to preserve existing log history

        except Exception as e:
            _log.debug(f"Error updating UI texts: {e}")

    def _update_button_texts(self):
        """Main and tools button texts (part of update_ui_texts, same order; P5-9 split)."""
        # Update button texts (no emojis for all languages)
        if hasattr(self, 'backup_btn'):
            text = self.translator.get('backup_data')
            self.backup_btn.config(text=text)
        if hasattr(self, 'restore_btn'):
            text = self.translator.get('restore_data')
            self.restore_btn.config(text=text)
        if hasattr(self, 'filters_btn'):
            text = self.translator.get('filters')
            self.filters_btn.config(text=text)
        if hasattr(self, 'driver_backup_btn'):
            text = self.translator.get('backup_drivers')
            self.driver_backup_btn.config(text=text)
        if hasattr(self, 'driver_restore_btn'):
            text = self.translator.get('restore_drivers')
            self.driver_restore_btn.config(text=text)

        # Update tools buttons (no emojis for all languages)
        if hasattr(self, 'browser_profiles_btn'):
            text = self.translator.get('browser')
            self.browser_profiles_btn.config(text=text)
        if hasattr(self, 'check_space_btn'):
            text = self.translator.get('check_space')
            self.check_space_btn.config(text=text)
        if hasattr(self, 'save_log_btn'):
            text = self.translator.get('save_log')
            self.save_log_btn.config(text=text)
        if hasattr(self, 'copy_btn'):
            text = self.translator.get('copy_data')
            self.copy_btn.config(text=text)
        if hasattr(self, 'devmgr_btn'):
            text = self.translator.get('device_mgr')
            self.devmgr_btn.config(text=text)
        if hasattr(self, 'schedule_btn'):
            text = self.translator.get('schedule')
            self.schedule_btn.config(text=text)
        if hasattr(self, 'winget_export_btn'):
            text = self.translator.get('export_apps')
            self.winget_export_btn.config(text=text)
        if hasattr(self, 'file_explorer_btn'):
            text = self.translator.get('explorer')
            self.file_explorer_btn.config(text=text)
        if hasattr(self, 'nas_connect_btn'):
            text = self.translator.get('connect_nas')
            self.nas_connect_btn.config(text=text)
        if hasattr(self, 'nas_disconnect_btn'):
            text = self.translator.get('disconnect')
            self.nas_disconnect_btn.config(text=text)

    def switch_language(self, lang_code):
        """Switch application language"""
        if self.translator.set_language(lang_code):
            self.update_ui_texts()
            try:
                self.save_settings()  # Save the language preference
            except Exception:
                # settings.json not writable: the choice applies to this session only
                _log.debug("event=settings_save_failed source=language", exc_info=True)
            dialogs.Win11Dialog.showinfo(self.translator.get('language_changed'),
                                         self.translator.get('language_changed_msg'),
                                         parent=self, theme=self.theme, translator=self.translator, show_icon=False)
