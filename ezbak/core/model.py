"""Domain types shared by the GUI, Headless and scheduler paths (ADR 0002).

`DEFAULT_FILTERS` keeps the dict/list shape stored in settings and schedule
JSON (C3). `FilterSet.from_dict` / `to_dict` convert between that shape and
the typed form.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, ClassVar, Mapping, Optional

FILTER_TYPES = ("ext", "name", "path")

# The only definition. Shape is part of the settings/schedule JSON format (C3).
DEFAULT_FILTERS = {
    'include': [],
    'exclude': [
        {'type': 'name', 'pattern': 'onedrive*'},
        {'type': 'name', 'pattern': 'thumbs.db'},
        {'type': 'name', 'pattern': 'desktop.ini'},
        {'type': 'name', 'pattern': '$recycle.bin'},
    ],
}


@dataclass(frozen=True)
class FilterRule:
    """One include/exclude rule. `type` is normalized to lower case."""

    type: str
    pattern: str

    def __post_init__(self) -> None:
        normalized = (self.type or "").strip().lower()
        if normalized not in FILTER_TYPES:
            raise ValueError(f"unknown filter type: {self.type!r}")
        object.__setattr__(self, "type", normalized)
        object.__setattr__(self, "pattern", self.pattern if self.pattern is not None else "")

    @classmethod
    def from_dict(cls, data: Any) -> Optional["FilterRule"]:
        """Build a rule from a JSON dict; returns None for anything not ext/name/path."""
        if not isinstance(data, Mapping):
            return None
        rtype = data.get("type")
        if not isinstance(rtype, str) or rtype.strip().lower() not in FILTER_TYPES:
            return None
        pattern = data.get("pattern")
        return cls(rtype, pattern if isinstance(pattern, str) else "")

    def to_dict(self) -> dict:
        return {"type": self.type, "pattern": self.pattern}


@dataclass(frozen=True)
class FilterSet:
    include: tuple = ()
    exclude: tuple = ()

    @classmethod
    def from_dict(cls, data: Any) -> "FilterSet":
        """Convert the JSON shape; invalid rules are dropped, None/empty gives no rules."""
        data = data or {}

        def _rules(key: str) -> tuple:
            built = (FilterRule.from_dict(item) for item in (data.get(key) or []))
            return tuple(rule for rule in built if rule is not None)

        return cls(include=_rules("include"), exclude=_rules("exclude"))

    def to_dict(self) -> dict:
        return {
            "include": [rule.to_dict() for rule in self.include],
            "exclude": [rule.to_dict() for rule in self.exclude],
        }


@dataclass(frozen=True)
class BackupJob:
    user: str
    dest: str
    include_hidden: bool = False
    include_system: bool = False
    retention_count: int = 2
    log_retention_days: int = 30
    # JSON dict shape ({'include': [...], 'exclude': [...]}) as stored in settings /
    # schedule JSON. Not a FilterSet: from_dict drops invalid rules (core CLAUDE.md).
    filters: Optional[dict] = None
    # Accepted for C3 (CLI/schedule argument) but not used, as today.
    log_folder: Optional[str] = None


@dataclass(frozen=True)
class BackupPolicy:
    """GUI/Headless behaviour differences kept as explicit flags (ADR 0002, Q1).

    `GUI` reproduces `App.is_hidden_safe` + `App.run_backup`; `HEADLESS`
    reproduces `core_run_backup`. `App.is_hidden` (still used by restore,
    copy and space calculation) matches neither; its callers are to use
    `dataclasses.replace(BackupPolicy.GUI, skip_appdata_names=True,
    attr_ctypes_fallback=True)` rather than a third named constant.
    """

    # Skip names starting with '.' and thumbs.db / desktop.ini / $recycle.bin
    # regardless of attributes or filters.
    skip_builtin_names: bool
    # Skip AppData / Application Data / Local Settings unless both hidden and
    # system are included.
    skip_appdata_names: bool
    # Without pywin32, read attributes through ctypes; False skips the
    # hidden/system/reparse check entirely.
    attr_ctypes_fallback: bool
    # filters=None means DEFAULT_FILTERS; False means no rules.
    none_filters_use_default: bool
    # Filter rules apply only to existing paths (nonexistent -> not skipped).
    filter_requires_existing_path: bool
    # Compare total size with destination free space before copying.
    check_free_space: bool
    # Attribute skips and filter skips are checked and logged separately
    # ("Skipping hidden ..." vs "Skipping filtered directory").
    report_filtered_separately: bool
    # Count every attempted file copy (failed or empty ones too) and total the
    # bytes from progress; False counts only copies that returned > 0 bytes.
    count_attempted_copies: bool
    # Each browser's bookmark copy has its own try (one failure does not stop
    # the other); False wraps all bookmark copies in a single try.
    isolate_bookmark_errors: bool
    # Old-log cleanup runs a second scan/delete pass after the first one.
    log_cleanup_retry_pass: bool
    # Old-backup cleanup catches a failed folder delete per folder and goes on
    # ("Could not delete"); False lets it end the whole cleanup.
    isolate_backup_delete_errors: bool
    # An exception from the skip check of a file means "include it"; False
    # lets it fail the file ("ERROR: Failed to copy file ...").
    skip_errors_include: bool
    # Retention errors are logged and only that cleanup is skipped; the backup goes on.
    # A failed read of retention_count is logged (bk_error) and only the old-backup
    # cleanup is skipped (log cleanup still runs); any other cleanup error is logged
    # ("Cleanup failed") and ends the retention step. False lets both propagate.
    isolate_retention_errors: bool
    # A CRITICAL error of the tree walk ends the run and is returned in
    # BackupResult.critical_error (the wrapper reports it); False re-raises it.
    return_critical_errors: bool

    GUI: ClassVar["BackupPolicy"]
    HEADLESS: ClassVar["BackupPolicy"]


BackupPolicy.GUI = BackupPolicy(
    skip_builtin_names=False,
    skip_appdata_names=False,
    attr_ctypes_fallback=False,
    none_filters_use_default=False,
    filter_requires_existing_path=True,
    check_free_space=True,
    report_filtered_separately=False,
    count_attempted_copies=True,
    isolate_bookmark_errors=True,
    log_cleanup_retry_pass=True,
    isolate_backup_delete_errors=False,
    skip_errors_include=True,
    isolate_retention_errors=True,
    return_critical_errors=True,
)

BackupPolicy.HEADLESS = BackupPolicy(
    skip_builtin_names=True,
    skip_appdata_names=True,
    attr_ctypes_fallback=True,
    none_filters_use_default=True,
    filter_requires_existing_path=False,
    check_free_space=False,
    report_filtered_separately=True,
    count_attempted_copies=False,
    isolate_bookmark_errors=False,
    log_cleanup_retry_pass=False,
    isolate_backup_delete_errors=True,
    skip_errors_include=False,
    isolate_retention_errors=False,
    return_critical_errors=False,
)


@dataclass(frozen=True)
class BackupResult:
    backup_path: str
    log_path: str
    files: int = 0
    folders: int = 0
    bytes: int = 0
    # Set only with policy.return_critical_errors: the CRITICAL error that ended the run.
    critical_error: Optional[BaseException] = None
