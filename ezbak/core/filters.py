"""Filter matching and hidden / system / reparse-point skip judgment (only implementation).

Replaces the three copies that lived in ezBAK.py (`App.is_hidden`,
`App.is_hidden_safe`, `core_run_backup.is_hidden_headless` / `matches_filter`).
GUI and Headless differences are selected only by `BackupPolicy` flags.

Known quirks are kept on purpose (Q1): `ext` compares only the last suffix,
`[ab]` is not a character class, an include list whose rules match nothing
(empty pattern, unknown type) skips every file, include rules never apply to
directories.
"""
from __future__ import annotations

import fnmatch
import os
from typing import Any, Optional

from ezbak.winsys import fileattrs

from .model import DEFAULT_FILTERS, BackupPolicy, FilterSet

# is_skipped() reasons. Both are truthy; "not skipped" is None.
SKIP_HIDDEN = "hidden"      # built-in names, AppData names, hidden/system/reparse attributes
SKIP_FILTERED = "filtered"  # user include/exclude rules

_BUILTIN_NAMES = ("thumbs.db", "desktop.ini", "$recycle.bin")
_APPDATA_NAMES = ("appdata", "application data", "local settings")
_NO_RULES = {"include": [], "exclude": []}


def _rule_matches(rule, is_dir, name_l, ext_l, filepath_lower):
    try:
        rtype = (rule.get('type') or '').lower()
        patt = (rule.get('pattern') or '').strip()
        if not patt:
            return False
        patt_lower = patt.lower()
        if rtype == 'ext':
            if is_dir:
                return False
            return ext_l == patt_lower.lstrip('*.')
        if rtype == 'name':
            if '*' in patt_lower or '?' in patt_lower:
                return fnmatch.fnmatch(name_l, patt_lower)
            return name_l == patt_lower
        if rtype == 'path':
            if patt_lower in filepath_lower:
                return True
            path_normalized = filepath_lower.replace('\\', '/')
            patt_normalized = patt_lower.replace('\\', '/')
            if patt_normalized in path_normalized:
                return True
            if '*' in patt_normalized or '?' in patt_normalized:
                return fnmatch.fnmatch(path_normalized, patt_normalized)
        return False
    except Exception:
        # Malformed rule (non-dict, non-str pattern): treated as "no match", as before.
        return False


def matches_filter(path: str, filters: Any, is_dir: bool = False) -> bool:
    """True when `path` is excluded by `filters` (JSON dict shape or FilterSet).

    Excludes win; include rules apply to files only and, when present, every
    file that matches none of them is excluded. None means no rules. The path
    is not checked for existence here.
    """
    if isinstance(filters, FilterSet):
        filters = filters.to_dict()
    filters = filters or _NO_RULES
    try:
        name_l = os.path.basename(path).lower()
        ext_l = os.path.splitext(name_l)[1].lstrip('.')
        filepath_lower = path.lower()
        for rule in filters.get('exclude', []) or []:
            if _rule_matches(rule, is_dir, name_l, ext_l, filepath_lower):
                return True
        inc_rules = filters.get('include', []) or []
        if inc_rules and not is_dir:
            for rule in inc_rules:
                if _rule_matches(rule, is_dir, name_l, ext_l, filepath_lower):
                    return False
            return True
    except Exception:
        # Malformed filter container: not excluded by any rule, as before.
        pass
    return False


def _read_attributes(path, policy):
    """Raw attribute mask, or None when it is not read (no pywin32 and no fallback) or fails."""
    if not fileattrs.HAS_PYWIN32 and not policy.attr_ctypes_fallback:
        return None
    try:
        return fileattrs.get_attributes(path)
    except Exception:
        # Unreadable attributes are recovered as "no attribute skip" (all three old copies did this).
        return None


def _hidden_reason(path, include_hidden, include_system, policy):
    base = os.path.basename(path).lower()
    if policy.skip_builtin_names and (base.startswith('.') or base in _BUILTIN_NAMES):
        return SKIP_HIDDEN
    if policy.skip_appdata_names and base in _APPDATA_NAMES:
        if not (include_hidden and include_system):
            return SKIP_HIDDEN
    attrs = _read_attributes(path, policy)
    if attrs is None:
        return None
    if attrs & fileattrs.FILE_ATTRIBUTE_REPARSE_POINT:
        return SKIP_HIDDEN
    if attrs & fileattrs.FILE_ATTRIBUTE_HIDDEN and not include_hidden:
        return SKIP_HIDDEN
    if attrs & fileattrs.FILE_ATTRIBUTE_SYSTEM and not include_system:
        return SKIP_HIDDEN
    return None


def is_skipped(path: str, job: Any, policy: BackupPolicy,
               is_dir: Optional[bool] = None, check_filters: bool = True) -> Optional[str]:
    """Why `path` is left out of a backup: SKIP_HIDDEN, SKIP_FILTERED or None (kept).

    `job` is a BackupJob or any object with `include_hidden`, `include_system`
    and (when `check_filters`) `filters` (JSON dict shape or FilterSet; None
    follows `policy.none_filters_use_default`). `is_dir=None` asks the file
    system. Name/attribute checks come first, so a path that is both hidden and
    filtered reports SKIP_HIDDEN. Use `is not None` (or bool()) for a plain answer.
    """
    exists = os.path.exists(path)
    if exists:
        reason = _hidden_reason(path, bool(job.include_hidden), bool(job.include_system), policy)
        if reason is not None:
            return reason
    if not check_filters or (policy.filter_requires_existing_path and not exists):
        return None
    filters = getattr(job, 'filters', None)
    if filters is None and policy.none_filters_use_default:
        filters = DEFAULT_FILTERS
    if is_dir is None:
        try:
            is_dir = os.path.isdir(path)
        except Exception:
            # Same fallback as App.is_hidden: treat as a file.
            is_dir = False
    return SKIP_FILTERED if matches_filter(path, filters, is_dir) else None
