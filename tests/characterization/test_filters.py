"""Characterization (golden) of filter matching in the three current implementations.

- gui:      App.is_hidden            (ezBAK.py, GUI backup / size / source dialog path)
- safe:     App.is_hidden_safe       (ezBAK.py, GUI safe backup path with timeout)
- headless: ezbak.core.matches_filter with the filters core_run_backup passes
            (None -> copy_default_filters(); no existence check)

Expected values are what the code returns today. Differences between implementations
are kept as-is (do not "fix" them here). Change only with an approved Q1 exception.
"""
import os
from types import SimpleNamespace

import pytest

from ezbak.core import matches_filter

pytestmark = pytest.mark.windows

E = "exclude"
I = "include"


def rules(kind, *pairs):
    """rules(E, ('ext', 'txt')) -> {'include': [], 'exclude': [{'type': 'ext', 'pattern': 'txt'}]}"""
    out = {"include": [], "exclude": []}
    out[kind] = [{"type": t, "pattern": p} for t, p in pairs]
    return out


DEFAULTS = "DEFAULTS"  # placeholder: replaced with ezBAK.copy_default_filters()

# (id, filters, relative path, is_dir, expected (gui, safe, headless))
# True = skipped (excluded from backup), False = kept
CASES = [
    ("no_rules", {"include": [], "exclude": []}, "zqdocs/a.txt", False, (False, False, False)),
    ("ext_plain", rules(E, ("ext", "txt")), "zqdocs/a.txt", False, (True, True, True)),
    ("ext_star_dot_upper", rules(E, ("ext", "*.TXT")), "zqdocs/a.txt", False, (True, True, True)),
    ("ext_leading_dot", rules(E, ("ext", ".txt")), "zqdocs/a.txt", False, (True, True, True)),
    ("ext_upper_file", rules(E, ("ext", "txt")), "zqdocs/B.TXT", False, (True, True, True)),
    ("ext_ignored_for_dir", rules(E, ("ext", "txt")), "zqdocs/folder.txt", True, (False, False, False)),
    ("ext_prefix_no_match", rules(E, ("ext", "tx")), "zqdocs/a.txt", False, (False, False, False)),
    ("ext_double_suffix", rules(E, ("ext", "*.tar.gz")), "zqdocs/x.tar.gz", False, (False, False, False)),
    ("ext_last_suffix", rules(E, ("ext", "gz")), "zqdocs/x.tar.gz", False, (True, True, True)),
    ("ext_no_extension", rules(E, ("ext", "txt")), "zqdocs/README", False, (False, False, False)),
    ("name_exact_case", rules(E, ("name", "Thumbs.db")), "zqdocs/thumbs.DB", False, (True, True, True)),
    ("name_star_dir", rules(E, ("name", "onedrive*")), "OneDrive - Corp", True, (True, True, True)),
    ("name_question", rules(E, ("name", "a?c.txt")), "zqdocs/abc.txt", False, (True, True, True)),
    ("name_no_partial", rules(E, ("name", "abc")), "zqdocs/abc.txt", False, (False, False, False)),
    ("name_brackets_literal", rules(E, ("name", "[ab].txt")), "zqdocs/a.txt", False, (False, False, False)),
    ("name_brackets_exact", rules(E, ("name", "[ab].txt")), "zqdocs/[ab].txt", False, (True, True, True)),
    ("name_dir", rules(E, ("name", "zqsub")), "zqdocs/zqsub", True, (True, True, True)),
    ("path_contains", rules(E, ("path", "zqsub")), "zqdocs/zqsub/a.txt", False, (True, True, True)),
    ("path_backslash", rules(E, ("path", "zqdocs\\zqsub")), "zqdocs/zqsub/a.txt", False, (True, True, True)),
    ("path_forward_slash", rules(E, ("path", "zqdocs/zqsub")), "zqdocs/zqsub/a.txt", False, (True, True, True)),
    ("path_upper", rules(E, ("path", "ZQDOCS\\ZQSUB")), "zqdocs/zqsub/a.txt", False, (True, True, True)),
    ("path_wildcard", rules(E, ("path", "*/zqsub/*.txt")), "zqdocs/zqsub/a.txt", False, (True, True, True)),
    ("path_wildcard_no_match", rules(E, ("path", "*/zqsub/*.log")), "zqdocs/zqsub/a.txt", False, (False, False, False)),
    ("path_wildcard_no_leading_star", rules(E, ("path", "zqdocs/*.txt")), "zqdocs/a.txt", False, (False, False, False)),
    ("path_dir", rules(E, ("path", "zqdocs\\zqsub")), "zqdocs/zqsub", True, (True, True, True)),
    ("include_match", rules(I, ("ext", "txt")), "zqdocs/a.txt", False, (False, False, False)),
    ("include_no_match", rules(I, ("ext", "txt")), "zqdocs/b.log", False, (True, True, True)),
    ("include_not_for_dir", rules(I, ("ext", "txt")), "zqdocs", True, (False, False, False)),
    ("include_name_dir_not_applied", rules(I, ("name", "keep*")), "zqother", True, (False, False, False)),
    ("include_and_exclude", {"include": [{"type": "ext", "pattern": "txt"}],
                             "exclude": [{"type": "name", "pattern": "secret*"}]},
     "zqdocs/secret.txt", False, (True, True, True)),
    ("include_two_rules_second", {"include": [{"type": "ext", "pattern": "doc"},
                                              {"type": "ext", "pattern": "txt"}],
                                  "exclude": []}, "zqdocs/a.txt", False, (False, False, False)),
    ("empty_pattern_exclude", rules(E, ("ext", "   ")), "zqdocs/a.txt", False, (False, False, False)),
    ("empty_pattern_include_only", rules(I, ("ext", "")), "zqdocs/a.txt", False, (True, True, True)),
    ("none_pattern", {"include": [], "exclude": [{"type": "name", "pattern": None}]},
     "zqdocs/a.txt", False, (False, False, False)),
    ("unknown_type", rules(E, ("glob", "*.txt")), "zqdocs/a.txt", False, (False, False, False)),
    ("type_uppercase", rules(E, ("EXT", "txt")), "zqdocs/a.txt", False, (True, True, True)),
    ("missing_type_key", {"include": [], "exclude": [{"pattern": "a.txt"}]},
     "zqdocs/a.txt", False, (False, False, False)),
    ("missing_lists", {}, "zqdocs/a.txt", False, (False, False, False)),
    ("defaults_onedrive_dir", DEFAULTS, "OneDrive", True, (True, True, True)),
    ("defaults_desktop_ini", DEFAULTS, "zqdocs/Desktop.ini", False, (True, True, True)),
    ("defaults_recycle_dir", DEFAULTS, "$Recycle.Bin", True, (True, True, True)),
    ("defaults_normal", DEFAULTS, "zqdocs/normal.txt", False, (False, False, False)),
    # filters=None: GUI falls back to empty rules, Headless falls back to DEFAULT_FILTERS
    ("none_filters_desktop_ini", None, "zqdocs/desktop.ini", False, (False, False, True)),
    ("none_filters_normal", None, "zqdocs/normal.txt", False, (False, False, False)),
]


class _Var:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


def _make_entry(root, rel, is_dir):
    path = os.path.join(str(root), *rel.split("/"))
    if is_dir:
        os.makedirs(path, exist_ok=True)
    else:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write("x")
    return path


def _resolve(mod, filters):
    return mod.copy_default_filters() if filters == DEFAULTS else filters


def run_all(mod, filters, path, is_dir):
    """Return (gui, safe, headless) for one entry. Attribute policy is 'include' for both."""
    fake_app = SimpleNamespace(filters=filters,
                               hidden_mode_var=_Var("include"),
                               system_mode_var=_Var("include"))
    gui = mod.App.is_hidden(fake_app, path)
    safe = mod.App.is_hidden_safe(fake_app, path)
    # core_run_backup replaces filters=None with copy_default_filters() before matches_filter
    headless_filters = filters if filters is not None else mod.copy_default_filters()
    headless = matches_filter(path, headless_filters, is_dir)
    return (gui, safe, headless)


@pytest.mark.parametrize("case_id,filters,rel,is_dir,expected", CASES, ids=[c[0] for c in CASES])
def test_filter_golden(ezbak_mod, isolated_env, tmp_path,
                       case_id, filters, rel, is_dir, expected):
    path = _make_entry(tmp_path / "src", rel, is_dir)
    filters = _resolve(ezbak_mod, filters)
    assert run_all(ezbak_mod, filters, path, is_dir) == expected


def test_headless_matches_filter_ignores_existence(ezbak_mod, tmp_path):
    """GUI variants return False for nonexistent paths; matches_filter does not check existence."""
    path = str(tmp_path / "missing" / "a.txt")
    filters = rules(E, ("ext", "txt"))
    fake_app = SimpleNamespace(filters=filters, hidden_mode_var=_Var("include"),
                               system_mode_var=_Var("include"))
    assert (ezbak_mod.App.is_hidden(fake_app, path),
            ezbak_mod.App.is_hidden_safe(fake_app, path),
            matches_filter(path, filters, False)) == (False, False, True)


def test_default_filters_value(ezbak_mod):
    assert ezbak_mod.DEFAULT_FILTERS == {
        "include": [],
        "exclude": [
            {"type": "name", "pattern": "onedrive*"},
            {"type": "name", "pattern": "thumbs.db"},
            {"type": "name", "pattern": "desktop.ini"},
            {"type": "name", "pattern": "$recycle.bin"},
        ],
    }
    copied = ezbak_mod.copy_default_filters()
    copied["exclude"].clear()
    assert len(ezbak_mod.DEFAULT_FILTERS["exclude"]) == 4
