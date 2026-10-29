"""Job log writer and GUI/Headless wording (ADR 0003 section 2).

The job log (`<user>_backup_<ts>.log`, restore logs) is a user artifact. Lines
are `[YYYY-MM-DD HH:MM:SS] <message>`; wording, order and line breaks follow
the ADR 0001 baseline. This module never goes through `logging` and never
reads the wall clock itself: time comes only from the `clock` argument.

The caller opens the file (mode, encoding and error handling stay with the
entry wrapper, ADR 0002 D2) and hands the stream to `JobLog`.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, ClassVar, Iterable, Iterator, Mapping, NamedTuple, Optional

TIMESTAMP_FORMAT = "%Y-%m-%d %H:%M:%S"


@dataclass(frozen=True)
class LogItem:
    """One wording entry. `text` is a str.format template; `lead`/`end` are
    written before the `[ts]` prefix and after the message."""

    text: str
    lead: str = ""
    end: str = "\n"

    def line(self, **values: Any) -> "JobLine":
        return JobLine(self.text.format(**values), self.lead, self.end)


class JobLine(NamedTuple):
    message: str
    lead: str = ""
    end: str = "\n"


@dataclass(frozen=True)
class JobLogFormat:
    """GUI / Headless wording. None means the line is not written.

    Header and retention items live here; the remaining wording moves in
    with each step (P2-5..P2-7).
    Template fields: start(user, source, backup_path),
    policy(hidden, system, backup_keep, log_keep), filters(include, exclude).
    Retention (core.retention): bk_* (count, keep, path, error), lg_* (folder,
    days, now, cutoff, count, name, date, age, flag, error, tb).
    Behaviour (per-folder try, retry pass) comes from BackupPolicy, not from
    these items. lg_header_first / lg_header: the old-log header before /
    after the disabled and missing-folder checks.
    Bookmarks (targets.browser): bm_* (browser, error). bm_saved_last /
    bm_missing_last are used for the last browser of the table.
    Engine (core.engine): size_* (size, error, path), space_short (required,
    free), tree copy items (path, rel, src, dst, error), completed (files,
    folders, size). rp_* are Reporter messages, not job log lines.
    """

    start: Optional[LogItem]
    policy: Optional[LogItem]
    filters: Optional[LogItem]
    separator: Optional[LogItem]
    # Old backup cleanup
    bk_header: Optional[LogItem]
    bk_disabled: Optional[LogItem]
    bk_found_keep: Optional[LogItem]
    bk_deleting: Optional[LogItem]
    bk_delete_failed: Optional[LogItem]
    bk_finished: Optional[LogItem]
    bk_nothing: Optional[LogItem]
    bk_error: Optional[LogItem]
    # Old log cleanup
    lg_header_first: Optional[LogItem]
    lg_disabled: Optional[LogItem]
    lg_folder_missing: Optional[LogItem]
    lg_header: Optional[LogItem]
    lg_scanning: Optional[LogItem]
    lg_current_time: Optional[LogItem]
    lg_cutoff: Optional[LogItem]
    lg_total_files: Optional[LogItem]
    lg_files_found: Optional[LogItem]
    lg_no_files: Optional[LogItem]
    lg_file: Optional[LogItem]
    lg_modified: Optional[LogItem]
    lg_age: Optional[LogItem]
    lg_should_delete: Optional[LogItem]
    lg_deleted: Optional[LogItem]
    lg_delete_failed: Optional[LogItem]
    lg_check_error: Optional[LogItem]
    lg_list_error: Optional[LogItem]
    lg_retry_deleted: Optional[LogItem]
    lg_retry_failed: Optional[LogItem]
    lg_done: Optional[LogItem]
    lg_nothing: Optional[LogItem]
    lg_error: Optional[LogItem]
    lg_traceback: Optional[LogItem]
    # Browser bookmarks
    bm_header: Optional[LogItem]
    bm_saved: Optional[LogItem]
    bm_missing: Optional[LogItem]
    bm_saved_last: Optional[LogItem]
    bm_missing_last: Optional[LogItem]
    bm_browser_error: Optional[LogItem]
    bm_error: Optional[LogItem]
    # Engine steps (core.engine): size, profile tree copy, completion
    size_total: Optional[LogItem]
    size_failed: Optional[LogItem]
    size_total_after_error: Optional[LogItem]
    size_walk_error: Optional[LogItem]
    size_file_error: Optional[LogItem]
    space_short: Optional[LogItem]
    processing_dir: Optional[LogItem]
    skip_hidden_dir: Optional[LogItem]
    skip_filtered_dir: Optional[LogItem]
    dir_check_error: Optional[LogItem]
    folder_created: Optional[LogItem]
    folder_failed: Optional[LogItem]
    skip_hidden_file: Optional[LogItem]
    skip_filtered_file: Optional[LogItem]
    copying: Optional[LogItem]
    copy_error: Optional[LogItem]
    copy_timeout: Optional[LogItem]
    file_failed: Optional[LogItem]
    dir_failed: Optional[LogItem]
    critical: Optional[LogItem]
    completed: Optional[LogItem]
    cleanup_failed: Optional[LogItem]
    finished: Optional[LogItem]
    # GUI entry wrapper lines (App.run_backup; P2-7)
    existing_deleted: Optional[LogItem]
    existing_delete_failed: Optional[LogItem]
    folder_created_new: Optional[LogItem]
    backup_error: Optional[LogItem]
    backup_traceback: Optional[LogItem]
    # Reporter messages (not job log lines; message text only, lead/end unused)
    rp_user: Optional[LogItem]
    rp_size_start: Optional[LogItem]
    rp_size_progress: Optional[LogItem]
    rp_space_check: Optional[LogItem]
    rp_space_short: Optional[LogItem]
    rp_starting: Optional[LogItem]
    rp_folder: Optional[LogItem]
    rp_files_log: Optional[LogItem]
    rp_files_status: Optional[LogItem]
    rp_copy_error: Optional[LogItem]
    rp_critical: Optional[LogItem]
    rp_done_log: Optional[LogItem]
    rp_done_status: Optional[LogItem]
    rp_error: Optional[LogItem]

    GUI: ClassVar["JobLogFormat"]
    HEADLESS: ClassVar["JobLogFormat"]


_POLICY = LogItem("Policy: Hidden={hidden}, System={system}, BackupKeep={backup_keep}, LogKeep={log_keep}")
_FILTERS = LogItem("Filters: Include=[{include}], Exclude=[{exclude}]")
_SEPARATOR = LogItem("-" * 50)

# Retention wording shared by both paths (identical text, lead and end).
_RETENTION = dict(
    bk_disabled=LogItem("Backup cleanup is disabled (retention count is 0 or less)."),
    bk_found_keep=LogItem("Found {count} backups. Keeping the {keep} most recent."),
    bk_deleting=LogItem("  Deleting old backup: {path}"),
    bk_finished=LogItem("Cleanup finished."),
    lg_disabled=LogItem("Log cleanup disabled (retention days <= 0)"),
    lg_folder_missing=LogItem("Log folder does not exist: {folder}"),
    lg_scanning=LogItem("Scanning '{folder}' for logs older than {days} days..."),
    lg_current_time=LogItem("Current time: {now}"),
    lg_cutoff=LogItem("Cutoff time: {cutoff}"),
    lg_files_found=LogItem("Log files found: {count}"),
    lg_file=LogItem("  File: {name}"),
    lg_modified=LogItem("    Modified: {date}"),
    lg_age=LogItem("    Age: {age:.1f} days"),
    lg_should_delete=LogItem("    Should delete: {flag}"),
    lg_deleted=LogItem("    ✓ Deleted: {name}"),
    lg_delete_failed=LogItem("    ✗ Delete failed: {error}"),
    lg_check_error=LogItem("  Error checking {name}: {error}"),
    lg_list_error=LogItem("Error listing files: {error}"),
    lg_retry_deleted=LogItem("  ✓ Deleted: {name}"),
    lg_retry_failed=LogItem("  ✗ Error deleting {name}: {error}"),
    lg_done=LogItem("Log cleanup completed. Deleted {count} files."),
    lg_nothing=LogItem("No old log files found to delete."),
    lg_error=LogItem("Error during log cleanup: {error}"),
)

# Engine wording shared by both paths (identical text, lead and end).
_SIZE_TOTAL = LogItem("Calculating total size... {size}")
_ENGINE = dict(
    size_total=_SIZE_TOTAL,
    space_short=LogItem("Insufficient free space for backup. Required={required} Free={free}"),
    skip_hidden_dir=LogItem("Skipping hidden directory: {path}"),
    skip_filtered_dir=LogItem("Skipping filtered directory: {path}"),
    dir_check_error=LogItem("Error checking directory {path}: {error}"),
    folder_created=LogItem("Created folder: {path}"),
    folder_failed=LogItem("ERROR: Failed to create folder {path}: {error}"),
    skip_hidden_file=LogItem("Skipping hidden file: {path}"),
    skip_filtered_file=LogItem("Skipping filtered file: {path}"),
    copying=LogItem("Copying file: {src} -> {dst}"),
    file_failed=LogItem("ERROR: Failed to copy file {path}: {error}"),
    dir_failed=LogItem("ERROR: Failed to process directory {path}: {error}"),
    critical=LogItem("CRITICAL ERROR in backup loop: {error}"),
    completed=LogItem("Backup completed. Files backed up: {files}, Folders: {folders}, Total size: {size}"),
)

# GUI-only lines and Reporter messages (App.run_backup / compute_total_bytes_safe /
# copy_file_with_progress_safe). Headless writes none of these.
_ENGINE_GUI = dict(
    size_failed=LogItem("Total bytes calculation error: {error}"),
    size_total_after_error=_SIZE_TOTAL,
    size_walk_error=LogItem("Walk error: {error}"),
    size_file_error=LogItem("Size calculation error for {path}: {error}"),
    processing_dir=LogItem("Processing directory: {path} (relative: {rel})"),
    copy_error=LogItem("ERROR: copying file {src} -> {dst}: {error}"),
    copy_timeout=LogItem("TIMEOUT: File copy timeout for {path}"),
    cleanup_failed=LogItem("Cleanup failed: {error}"),
    finished=LogItem("Backup successfully completed."),
    existing_deleted=LogItem("Deleted existing backup folder '{path}'."),
    existing_delete_failed=LogItem("Failed to delete existing backup folder: {error}"),
    folder_created_new=LogItem("Created new backup folder '{path}'."),
    backup_error=LogItem("Backup error: {error}"),
    backup_traceback=LogItem("Backup traceback: {tb}"),
    rp_user=LogItem("Backing up user: {user}"),
    rp_size_start=LogItem("Calculating total size..."),
    rp_size_progress=LogItem("Calculating total size... {count} files, {size}"),
    rp_space_check=LogItem("Checking available space..."),
    rp_space_short=LogItem("There is not enough space in the destination.\nRequired: {required}\nAvailable: {free}"),
    rp_starting=LogItem("Starting backup..."),
    rp_folder=LogItem("Processing folder {count}: {name}"),
    rp_files_log=LogItem("Copied {count} files..."),
    rp_files_status=LogItem("Copied {count} files"),
    rp_copy_error=None,
    rp_critical=LogItem("A critical error occurred during backup: {error}"),
    rp_done_log=LogItem("Backup complete (detailed log saved)."),
    rp_done_status=LogItem("Backup complete!"),
    rp_error=LogItem("An error occurred during backup: {error}"),
)

# Headless: "Note: ..." when the size total fails (no size line after it),
# copy errors go to stdout through the Reporter (rp_copy_error), not the job log.
_ENGINE_HEADLESS = {**{k: None for k in _ENGINE_GUI}, **dict(
    size_failed=LogItem("Note: Could not pre-calculate total size: {error}"),
    rp_copy_error=LogItem("ERROR copying file {src} -> {dst}: {error}"),
)}

_BM_HEADER = LogItem("Backing up Chrome and Edge browser bookmarks...")
_BM_SAVED = LogItem("- {browser} bookmarks backed up.")
_BM_MISSING = LogItem("- {browser} bookmark file not found.")

# GUI writes its start line from run_backup, after the header of open_log_file.
# GUI retention: "[ts] \n[Old ...]" headers, no per-folder try, log header
# after the disabled/missing checks, "Total files" line, traceback on error.
JobLogFormat.GUI = JobLogFormat(
    start=LogItem("Backup start: {source} -> {backup_path}"),
    policy=_POLICY,
    filters=_FILTERS,
    separator=_SEPARATOR,
    bk_header=LogItem("\n[Old backup cleanup]"),
    bk_delete_failed=None,
    bk_nothing=LogItem("Found {count} backups. No old backups to clean up (retention policy: keep {keep})."),
    bk_error=LogItem("ERROR: An error occurred during backup cleanup: {error}"),
    lg_header_first=None,
    lg_header=LogItem("\n[Old log cleanup]"),
    lg_total_files=LogItem("Total files in folder: {count}"),
    lg_no_files=LogItem("No .log files found in the folder."),
    lg_traceback=LogItem("Traceback: {tb}"),
    bm_header=_BM_HEADER,
    bm_saved=_BM_SAVED,
    bm_missing=_BM_MISSING,
    bm_saved_last=_BM_SAVED,
    bm_missing_last=_BM_MISSING,
    bm_browser_error=LogItem("- {browser} bookmark backup error: {error}"),
    bm_error=LogItem("Browser bookmark backup failed: {error}"),
    **_RETENTION,
    **_ENGINE,
    **_ENGINE_GUI,
)

# Headless writes its start line first, before Policy.
# Headless retention: "[Old ...]" headers without a blank line, per-folder
# "Could not delete", log header first. Every Headless item is one plain
# "[ts] msg\n" line with no lead: the baseline's missing line breaks (end="")
# and the "\n" leads that covered them were removed by P2-fix #3 (Q1 exception,
# ADR 0003 section 6). bm_*_last now equal bm_saved / bm_missing.
JobLogFormat.HEADLESS = JobLogFormat(
    start=LogItem("Backup start for '{user}' → {backup_path}"),
    policy=_POLICY,
    filters=_FILTERS,
    separator=_SEPARATOR,
    bk_header=LogItem("[Old backup cleanup]"),
    bk_delete_failed=LogItem("  ERROR: Could not delete {path}: {error}"),
    bk_nothing=LogItem("Found {count} backups. No old backups to clean up (retention policy: keep {keep})."),
    bk_error=LogItem("ERROR: An error occurred during backup cleanup: {error}"),
    lg_header_first=LogItem("[Old log cleanup]"),
    lg_header=None,
    lg_total_files=None,
    lg_no_files=None,
    lg_traceback=None,
    bm_header=_BM_HEADER,
    bm_saved=_BM_SAVED,
    bm_missing=_BM_MISSING,
    bm_saved_last=_BM_SAVED,
    bm_missing_last=_BM_MISSING,
    bm_browser_error=None,
    bm_error=LogItem("Browser bookmark backup error: {error}"),
    **_RETENTION,
    **_ENGINE,
    **_ENGINE_HEADLESS,
)


def _rules_text(rules: Any) -> str:
    rules = rules or []
    if not rules:
        return "None"
    return ", ".join(f"{r.get('type', '').lower()}:{r.get('pattern', '')}" for r in rules)


def header_lines(fmt: JobLogFormat, *, hidden: Any, system: Any, backup_keep: Any,
                 log_keep: Any, filters: Optional[Mapping] = None,
                 start: Optional[Mapping[str, Any]] = None) -> Iterator[JobLine]:
    """Yield the header: [start], Policy, [Filters], separator.

    Values are written as given (GUI passes its var strings). `filters` is the
    JSON dict shape; None leaves the Filters line out (restore logs). `start`
    holds the start-line template values; None leaves the start line out.
    Lines are produced lazily, so a line already yielded is written even if a
    later one fails to format (as the inline writes did).
    """
    if start is not None and fmt.start is not None:
        yield fmt.start.line(**{"user": "", "source": "", "backup_path": "", **start})
    if fmt.policy is not None:
        yield fmt.policy.line(hidden=hidden, system=system, backup_keep=backup_keep, log_keep=log_keep)
    if filters is not None and fmt.filters is not None:
        include = _rules_text(filters.get("include", []))
        exclude = _rules_text(filters.get("exclude", []))
        yield fmt.filters.line(include=include, exclude=exclude)
    if fmt.separator is not None:
        yield fmt.separator.line()


class JobLog:
    """Writes job log lines to an already open text stream."""

    def __init__(self, stream: Any, clock: Callable[[], Any]) -> None:
        self._stream = stream
        self._clock = clock

    def _timestamp(self) -> str:
        return self._clock().strftime(TIMESTAMP_FORMAT)

    def write(self, message: str, lead: str = "", end: str = "\n", timestamp: Optional[str] = None) -> None:
        ts = timestamp if timestamp is not None else self._timestamp()
        self._stream.write(f"{lead}[{ts}] {message}{end}")

    def emit(self, item: Optional[LogItem], **values: Any) -> None:
        """Format and write one wording item; None writes nothing."""
        if item is None:
            return
        line = item.line(**values)
        self.write(line.message, line.lead, line.end)

    def write_lines(self, lines: Iterable[JobLine], single_timestamp: bool = False) -> None:
        """Write lines in order. `single_timestamp` stamps all of them with one
        clock reading (GUI header); otherwise the clock is read per line."""
        ts = self._timestamp() if single_timestamp else None
        for line in lines:
            self.write(line.message, line.lead, line.end, timestamp=ts)

    def flush(self) -> None:
        self._stream.flush()


def format_bytes(size: Any) -> str:
    """Human-readable size (B..PB, two decimals); "N/A" when not a number."""
    try:
        size = float(size)
    except (ValueError, TypeError):
        return "N/A"
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024.0:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} PB"
