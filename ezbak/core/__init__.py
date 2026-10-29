"""Backup/restore engine shared by GUI, Headless and scheduler. Use only these names.

The submodules diag (diagnostic log set-up and run_id, P6) and history (last backup
in state.json, U7-2) are public too.
"""
from . import diag, history
from .model import (
    DEFAULT_FILTERS,
    FILTER_TYPES,
    BackupJob,
    BackupPolicy,
    BackupResult,
    FilterRule,
    FilterSet,
)
from .filters import SKIP_FILTERED, SKIP_HIDDEN, is_skipped, matches_filter
from .fileops import (
    COPY_TIMEOUT_SECONDS,
    compute_total_size,
    copy_file_with_progress,
    copy_sources,
    handle_readonly_error,
    rmtree_readonly,
)
from .ports import BackupStore, BackupTarget, Reporter, TargetContext, TargetResult
from .joblog import JobLine, JobLog, JobLogFormat, LogItem, format_bytes, header_lines
from .retention import cleanup_old_backups, cleanup_old_logs
from .engine import InsufficientSpaceError, run_backup
from .restore import (KIND_DIRECT, KIND_PARENT, RestorePlan, RestoreResult,
                      RestoreSelectionError, locate_backup, run_restore, validate_selection)

__all__ = [
    "diag",
    "history",
    "KIND_DIRECT",
    "KIND_PARENT",
    "RestorePlan",
    "RestoreResult",
    "RestoreSelectionError",
    "locate_backup",
    "run_restore",
    "validate_selection",
    "InsufficientSpaceError",
    "run_backup",
    "cleanup_old_backups",
    "cleanup_old_logs",
    "BackupStore",
    "BackupTarget",
    "Reporter",
    "TargetContext",
    "TargetResult",
    "JobLine",
    "JobLog",
    "JobLogFormat",
    "LogItem",
    "format_bytes",
    "header_lines",
    "COPY_TIMEOUT_SECONDS",
    "compute_total_size",
    "copy_file_with_progress",
    "copy_sources",
    "handle_readonly_error",
    "rmtree_readonly",
    "SKIP_FILTERED",
    "SKIP_HIDDEN",
    "is_skipped",
    "matches_filter",
    "DEFAULT_FILTERS",
    "FILTER_TYPES",
    "BackupJob",
    "BackupPolicy",
    "BackupResult",
    "FilterRule",
    "FilterSet",
]
