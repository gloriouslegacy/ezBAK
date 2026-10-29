"""Restore: selection check, backup location (P3-2) and the restore run (P3-3).

The flow is validate_selection -> locate_backup -> (wrapper opens the job log)
-> run_restore (P3-3). User answers come in as callbacks (ADR 0002 D1/D2);
this module knows nothing about queues, tkinter or polling.
Error texts are the current GUI show_error messages, unchanged (C3).
"""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import Any, Callable, List, Optional, Sequence

from .fileops import compute_total_size, copy_file_with_progress
from .joblog import format_bytes

_log = logging.getLogger(__name__)

KIND_DIRECT = "direct"
KIND_PARENT = "parent"

# Folders whose presence marks a directly selected backup folder.
USER_DATA_FOLDERS = ("Desktop", "Documents", "Downloads", "Pictures", "Videos", "Music")
BACKUP_MARK = "_backup_"


class RestoreSelectionError(Exception):
    """The selected path cannot be restored from. str(e) is the show_error text."""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


@dataclass(frozen=True)
class RestorePlan:
    """Where to restore from.

    kind: KIND_DIRECT (a backup folder was picked) or KIND_PARENT (a folder holding backups).
    backup_root: folder holding the backup folder (job log and old-log cleanup go here).
    folder_name: the backup folder name inside backup_root.
    backup_user: user name taken from folder_name ('' when it has no '_backup_').
    """

    kind: str
    backup_root: str
    folder_name: str
    backup_user: str

    @property
    def source_dir(self) -> str:
        return os.path.join(self.backup_root, self.folder_name)


def validate_selection(path: str) -> None:
    """Raise RestoreSelectionError when path is a log file, a file, or missing."""
    if path.lower().endswith(".log"):
        raise RestoreSelectionError(
            f"The selected path is a log file.\n\n"
            f"'{path}'\n\n"
            f"You must select a backup folder")
    if os.path.isfile(path):
        raise RestoreSelectionError(
            f"The selected path is a file. Please select a folder:\n{path}")
    if not os.path.isdir(path):
        raise RestoreSelectionError(f"Backup folder not found:\n{path}")


def _backup_user_of(folder_name: str) -> str:
    return folder_name.split(BACKUP_MARK)[0] if BACKUP_MARK in folder_name else ""


def _locate_direct(path: str, folder_name: str) -> RestorePlan:
    has_user_folders = any(
        os.path.exists(os.path.join(path, folder)) for folder in USER_DATA_FOLDERS)
    if not has_user_folders:
        raise RestoreSelectionError(
            f"The selected folder appears to be named like a backup but does not contain user data:\n{path}\n\n"
            f"Please select either:\n"
            f"1. A parent folder containing backup folders, or\n"
            f"2. A valid backup folder with user data")
    return RestorePlan(KIND_DIRECT, os.path.dirname(path), folder_name,
                       _backup_user_of(folder_name))


def _list_backup_folders(path: str) -> List[str]:
    """Names in path that are folders containing '_backup_', in listdir order."""
    try:
        return [item for item in os.listdir(path)
                if os.path.isdir(os.path.join(path, item)) and BACKUP_MARK in item]
    except Exception as e:  # translate: listing failure becomes the current show_error text
        _log.info("restore_locate listdir_failed path=%s error=%s", path, type(e).__name__)
        raise RestoreSelectionError(f"Error reading backup folder:\n{e}") from e


def order_candidates(folders: Sequence[str], user: str) -> List[str]:
    """Matching user's backups first, then the rest; newest first (case-insensitive) within each."""
    prefix = f"{user.lower()}{BACKUP_MARK}"
    matching = [d for d in folders if d.lower().startswith(prefix)]
    others = [d for d in folders if not d.lower().startswith(prefix)]
    return (sorted(matching, key=str.lower, reverse=True)
            + sorted(others, key=str.lower, reverse=True))


def _locate_parent(path: str, user: str,
                   choose_backup: Callable[[List[str]], Optional[str]]) -> Optional[RestorePlan]:
    all_folders = _list_backup_folders(path)
    if not all_folders:
        raise RestoreSelectionError(
            f"No backup folders found in:\n{path}\n\n"
            f"Please select either:\n"
            f"1. A parent folder containing backup folders (e.g., E:\\), or\n"
            f"2. A backup folder directly (e.g., test_backup_2025-12-04_091400)")
    backup_folders = order_candidates(all_folders, user)
    if len(backup_folders) == 1:
        selected = backup_folders[0]
    else:
        selected = choose_backup(backup_folders)
        if not selected:
            return None
    return RestorePlan(KIND_PARENT, path, selected, _backup_user_of(selected))


def locate_backup(path: str, user: str,
                  choose_backup: Callable[[List[str]], Optional[str]]) -> Optional[RestorePlan]:
    """Find the backup to restore from a validated folder selection.

    choose_backup(candidates) is called only when several backups are found and
    returns the chosen name, or a falsy value to cancel (then None is returned).
    Raises RestoreSelectionError with the current show_error text otherwise.
    """
    folder_name = os.path.basename(path)
    if BACKUP_MARK in folder_name and os.path.isdir(path):
        plan = _locate_direct(path, folder_name)
    else:
        plan = _locate_parent(path, user, choose_backup)
    _log.info("restore_locate kind=%s cancelled=%s",
              plan.kind if plan else "-", plan is None)
    return plan


# --- run_restore (P3-3): job log lines and messages, current text (C3) -------------

MSG_CROSS_USER = "Cross-user restore detected: {backup_user} → {user}"
MSG_CROSS_CANCELLED = "Cross-user restore cancelled by user"
MSG_CROSS_CONFIRMED = "User confirmed cross-user restore"
MSG_RESTORING = "Restoring from {source} to {dest}"
MSG_SIZE_ERROR = "Total size calculation error: {error}"
MSG_TOTAL = "Total bytes to restore: {total}"
MSG_NO_SPACE_LOG = "Insufficient free space. Required={required}, Free={free}"
MSG_NO_SPACE = "Not enough free space to restore.\nRequired: {required}\nAvailable: {free}"
MSG_CREATED = "Created folder: {path}"
MSG_SKIP_HIDDEN = "Skipping hidden file: {path}"
MSG_FILE = "Restoring file: {src} -> {dst}"
MSG_COPY_ERROR = "Error copying file {src} -> {dst}: {error}"
MSG_PROGRESS = "Restored {count} files..."
MSG_PROGRESS_LOG = "{count} files restored so far; bytes_copied={copied}"
MSG_DONE = "Restore completed. Files restored: {count}, Total size: {size}"
MSG_DONE_OK = "Restore successfully completed."
SIZE_FALLBACK = 1024 * 1024 * 100  # used when the size walk fails (current: 100MB)
PROGRESS_EVERY = 50

Out = Callable[[str], None]
Skip = Callable[[str, bool], bool]


@dataclass
class RestoreResult:
    """Outcome of run_restore.

    cancelled: the cross-user restore was declined. error: show_error text when the
    restore stopped before copying (not enough space). files / bytes: what was copied.
    """

    cancelled: bool = False
    error: Optional[str] = None
    files: int = 0
    bytes: int = 0
    total: int = 0


def _confirm_cross_user(plan: RestorePlan, user: str, out: Out,
                        confirm: Callable[[str, str, str], bool]) -> bool:
    if plan.backup_user.lower() == user.lower():
        return True
    out(MSG_CROSS_USER.format(backup_user=plan.backup_user, user=user))
    if not confirm(plan.backup_user, user, plan.folder_name):
        out(MSG_CROSS_CANCELLED)
        return False
    out(MSG_CROSS_CONFIRMED)
    return True


def _total_size(source: str, skip: Skip, out: Out) -> int:
    failed: List[Exception] = []
    total, _ = compute_total_size(source, skip, strict=True, on_error=failed.append)
    if failed:
        out(MSG_SIZE_ERROR.format(error=failed[0]))
        return SIZE_FALLBACK
    return total


def _copy_tree(source: str, dest_root: str, skip: Skip, out: Out, reporter: Any,
               result: RestoreResult) -> None:
    def on_chunk(delta: int) -> None:
        result.bytes += delta
        reporter.progress(result.bytes)

    for dirpath, dirnames, filenames in os.walk(source, topdown=True, onerror=lambda e: None):
        dirnames[:] = [d for d in dirnames if not skip(os.path.join(dirpath, d), True)]
        rel_dir = os.path.relpath(dirpath, source)
        dest_dir = os.path.join(dest_root, rel_dir) if rel_dir != "." else dest_root
        os.makedirs(dest_dir, exist_ok=True)
        out(MSG_CREATED.format(path=dest_dir))
        for f in filenames:
            src = os.path.join(dirpath, f)
            if skip(src, False):
                out(MSG_SKIP_HIDDEN.format(path=src))
                continue
            dst = os.path.join(dest_dir, f)
            out(MSG_FILE.format(src=src, dst=dst))
            copy_file_with_progress(src, dst, on_chunk, on_error=lambda e, s=src, d=dst: out(
                MSG_COPY_ERROR.format(src=s, dst=d, error=e)))
            result.files += 1
            if result.files % PROGRESS_EVERY == 0:
                reporter.log(MSG_PROGRESS.format(count=result.files))
                out(MSG_PROGRESS_LOG.format(count=result.files, copied=result.bytes))


def run_restore(plan: RestorePlan, user: str, destination_dir: str, out: Out, reporter: Any, *,
                skip: Skip, free_space: Callable[[str], int],
                confirm_cross_user: Callable[[str, str, str], bool]) -> RestoreResult:
    """Restore plan.source_dir into destination_dir (the job log is already open).

    out(text) writes one job log line. confirm_cross_user(backup_user, user, folder_name)
    is asked only when the backup belongs to another user. free_space(path) errors
    propagate, as do skip errors in the copy walk. The final progress value is always
    reported (reporter.progress with a value that did not grow).
    """
    result = RestoreResult()
    if not _confirm_cross_user(plan, user, out, confirm_cross_user):
        result.cancelled = True
        _log.info("restore_run cancelled=True reason=cross_user")
        return result
    source = plan.source_dir
    out(MSG_RESTORING.format(source=source, dest=destination_dir))
    result.total = _total_size(source, skip, out)
    reporter.progress_max(result.total)
    out(MSG_TOTAL.format(total=result.total))
    free = free_space(destination_dir)
    if result.total > free:
        out(MSG_NO_SPACE_LOG.format(required=result.total, free=free))
        result.error = MSG_NO_SPACE.format(required=format_bytes(result.total),
                                           free=format_bytes(free))
        _log.info("restore_run stopped=no_space required=%s free=%s", result.total, free)
        return result
    _copy_tree(source, destination_dir, skip, out, reporter, result)
    reporter.progress(result.bytes)
    out(MSG_DONE.format(count=result.files, size=format_bytes(result.bytes)))
    out(MSG_DONE_OK)
    _log.info("restore_run files=%s bytes=%s", result.files, result.bytes)
    return result
