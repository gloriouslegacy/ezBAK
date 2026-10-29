"""settings.json persistence of App (moved from ezBAK.py in P5-8a).

SettingsMixin is mixed into App before tk.Tk; the methods use the same `self`
attributes as before. copy_default_filters is the single copy of the default
filter helper; ezBAK.py and dialogs/schedule.py import it under the same name.
Module-level names (os, copy, DEFAULT_FILTERS) match ezBAK.py so facade patches
reach this module too. The file path, key order and format stay unchanged.
"""
import copy
import logging
import os

from ezbak.core import DEFAULT_FILTERS

_log = logging.getLogger(__name__)


def copy_default_filters():
    """Return a deep copy of DEFAULT_FILTERS so callers can mutate safely."""
    return copy.deepcopy(DEFAULT_FILTERS)


class SettingsMixin:
    # Settings persistence
    def settings_file_path(self):
        appdata = os.environ.get('APPDATA') or os.path.expanduser('~')
        cfg_dir = os.path.join(appdata, 'ezbak')
        os.makedirs(cfg_dir, exist_ok=True)
        return os.path.join(cfg_dir, 'settings.json')

    def load_settings(self):
        import json
        path = self.settings_file_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # legacy keys
                if 'hidden' in data:
                    self.hidden_mode_var.set(data.get('hidden', 'exclude'))
                if 'system' in data:
                    self.system_mode_var.set(data.get('system', 'exclude'))
                if 'retention_count' in data:
                    self.retention_count_var.set(str(data.get('retention_count', '2')))
                if 'log_retention_days' in data:
                    self.log_retention_days_var.set(str(data.get('log_retention_days', '30')))

                # Sound settings load
                if 'sound_enabled' in data:
                    sound_enabled = data.get('sound_enabled', True)
                    self.sound_enabled_var.set(sound_enabled)
                # new filter structure
                fobj = data.get('filters', {}) if isinstance(data, dict) else {}
                inc = fobj.get('include', []) if isinstance(fobj, dict) else []
                exc = fobj.get('exclude', []) if isinstance(fobj, dict) else []

                def _sanitize(lst):
                    out = []
                    for it in lst or []:
                        try:
                            t = (it.get('type') or '').strip().lower()
                            p = (it.get('pattern') or '').strip()
                            if t in ('ext','name','path') and p:
                                out.append({'type': t, 'pattern': p})
                        except Exception:
                            continue
                    return out
                self.filters = {'include': _sanitize(inc), 'exclude': _sanitize(exc)}

                # Load theme and language settings
                if 'theme' in data:
                    theme_mode = data.get('theme', 'dark')
                    self.theme.current_mode = theme_mode
                    self.theme.colors = self.theme.DARK if theme_mode == 'dark' else self.theme.LIGHT

                if 'language' in data:
                    lang_code = data.get('language', 'en')
                    self.translator.current_lang = lang_code
                    # Note: UI texts will be updated after widgets are created in create_widgets()

            except Exception as e:
                _log.warning(f"Settings load error: {e}")

    def save_settings(self):
        import json
        path = self.settings_file_path()
        try:
            data = {
                'hidden': self.hidden_mode_var.get(),
                'system': self.system_mode_var.get(),
                'retention_count': self.retention_count_var.get(),
                'log_retention_days': self.log_retention_days_var.get(),
                'sound_enabled': self.sound_enabled_var.get(),
                'filters': self.filters,
                'theme': self.theme.current_mode,
                'language': self.translator.current_lang,
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f)
        except Exception as e:
            _log.warning(f"Settings save error: {e}")
