"""Old backup / old log cleanup shared by GUI and Headless (P2-4).

Wording comes from JobLogFormat (bk_* / lg_* items), behaviour differences from
BackupPolicy (isolate_backup_delete_errors, log_cleanup_retry_pass). Every
line goes to `out(JobLine)`; the caller decides where it is written (GUI
detailed log, Headless JobLog). Backups are listed and removed only through
the BackupStore port; old .log files are removed from the folder directly.

Errors are handled here as the job log end of a work unit (they become log
lines, as in the baseline) and are not raised.
"""
from __future__ import annotations

import logging
import os
import traceback
from datetime import datetime
from typing import Any, Callable, List, Optional

from .joblog import JobLine, JobLogFormat, LogItem
from .model import BackupPolicy
from .ports import BackupStore

_log = logging.getLogger(__name__)

Out = Callable[[JobLine], None]
DAY_SECONDS = 86400
_DISPLAY_FORMAT = "%Y-%m-%d %H:%M:%S"


def _emitter(fmt_out: Out) -> Callable[..., None]:
    def emit(item: Optional[LogItem], **values: Any) -> None:
        if item is not None:
            fmt_out(item.line(**values))
    return emit


def cleanup_old_backups(store: BackupStore, user_name: str, keep: Any,
                        fmt: JobLogFormat, policy: BackupPolicy, out: Out) -> None:
    """Keep the `keep` newest backups of user_name and delete the rest.

    `keep` is used as given (Headless passes its argument unconverted; a bad
    value ends in the bk_error line). Per-entry delete messages of the store go
    to `out`. policy.isolate_backup_delete_errors False: an error deleting one
    folder ends the whole cleanup (GUI); True: it is logged (bk_delete_failed)
    and the next folder follows (Headless).
    """
    emit = _emitter(out)
    try:
        emit(fmt.bk_header)
        if keep <= 0:
            emit(fmt.bk_disabled)
            return
        names = store.list_backups(user_name)
        if len(names) <= keep:
            emit(fmt.bk_nothing, count=len(names), keep=keep)
            return
        emit(fmt.bk_found_keep, count=len(names), keep=keep)
        on_log = lambda message: out(JobLine(message))  # noqa: E731
        for name in names[:-keep]:
            path = store.path_of(name)
            emit(fmt.bk_deleting, path=path)
            _log.info("event=retention_delete_backup user=%s path=%s", user_name, path)
            if not policy.isolate_backup_delete_errors:
                store.delete_backup(name, on_log)
                continue
            try:
                store.delete_backup(name, on_log)
            except Exception as e:  # logged per folder, cleanup goes on (Headless)
                emit(fmt.bk_delete_failed, path=path, error=e)
        emit(fmt.bk_finished)
    except Exception as e:  # work-unit end: becomes the job log error line
        _log.warning("event=retention_backup_failed user=%s error=%s", user_name, e)
        emit(fmt.bk_error, error=e)


def _scan_and_delete(folder: str, log_files: List[str], now: float, cutoff: float,
                     fmt: JobLogFormat, emit: Callable[..., None]) -> int:
    deleted = 0
    for filename in log_files:
        filepath = os.path.join(folder, filename)
        try:
            if not os.path.isfile(filepath):
                continue
            file_mtime = os.path.getmtime(filepath)
            file_date = datetime.fromtimestamp(file_mtime).strftime(_DISPLAY_FORMAT)
            age_days = (now - file_mtime) / DAY_SECONDS
            emit(fmt.lg_file, name=filename)
            emit(fmt.lg_modified, date=file_date)
            emit(fmt.lg_age, age=age_days)
            emit(fmt.lg_should_delete, flag=file_mtime < cutoff)
            if file_mtime < cutoff:
                try:
                    os.remove(filepath)
                    emit(fmt.lg_deleted, name=filename)
                    deleted += 1
                except Exception as e:  # per-file job log line
                    emit(fmt.lg_delete_failed, error=e)
        except Exception as e:  # per-file job log line
            emit(fmt.lg_check_error, name=filename, error=e)
    return deleted


def _retry_pass(folder: str, log_files: List[str], cutoff: float,
                fmt: JobLogFormat, emit: Callable[..., None]) -> int:
    deleted = 0
    for filename in log_files:
        filepath = os.path.join(folder, filename)
        try:
            if os.path.isfile(filepath) and os.path.getmtime(filepath) < cutoff:
                os.remove(filepath)
                emit(fmt.lg_retry_deleted, name=filename)
                deleted += 1
        except Exception as e:  # per-file job log line
            emit(fmt.lg_retry_failed, name=filename, error=e)
    return deleted


def _scan_with_retry(folder: str, now: float, cutoff: float, fmt: JobLogFormat,
                     emit: Callable[..., None]) -> Optional[int]:
    """GUI shape: list/scan, on error re-scan from the first listing, then a
    second delete pass. None means the cleanup stopped early (no .log files)."""
    log_files: List[str] = []
    try:
        all_files = os.listdir(folder)
        emit(fmt.lg_total_files, count=len(all_files))
        log_files = [f for f in all_files if f.endswith('.log')]
        emit(fmt.lg_files_found, count=len(log_files))
        if not log_files:
            emit(fmt.lg_no_files)
            return None
        deleted = _scan_and_delete(folder, log_files, now, cutoff, fmt, emit)
    except Exception:  # baseline: re-scan; all_files unbound re-raises to the caller
        log_files = [f for f in all_files if f.endswith('.log')]
        emit(fmt.lg_files_found, count=len(log_files))
        deleted = _scan_and_delete(folder, log_files, now, cutoff, fmt, emit)
    return deleted + _retry_pass(folder, log_files, cutoff, fmt, emit)


def cleanup_old_logs(folder: str, days: Any, fmt: JobLogFormat, policy: BackupPolicy,
                     out: Out, *, clock: Callable[[], datetime],
                     now_ts: Callable[[], float]) -> None:
    """Delete `.log` files in `folder` older than `days` days.

    `now_ts` gives the age reference (epoch seconds, time.time in both paths);
    `clock` gives the displayed "Current time". policy.log_cleanup_retry_pass
    selects the GUI shape (second pass, re-scan on error) over the Headless one.
    """
    emit = _emitter(out)
    try:
        emit(fmt.lg_header_first)
        if days <= 0:
            emit(fmt.lg_disabled)
            return
        if not os.path.isdir(folder):
            emit(fmt.lg_folder_missing, folder=folder)
            return
        now = now_ts()
        cutoff = now - (days * DAY_SECONDS)
        emit(fmt.lg_header)
        emit(fmt.lg_scanning, folder=folder, days=days)
        emit(fmt.lg_current_time, now=clock().strftime(_DISPLAY_FORMAT))
        emit(fmt.lg_cutoff, cutoff=datetime.fromtimestamp(cutoff).strftime(_DISPLAY_FORMAT))
        _log.info("event=retention_logs folder=%s days=%s", folder, days)
        if policy.log_cleanup_retry_pass:
            deleted = _scan_with_retry(folder, now, cutoff, fmt, emit)
            if deleted is not None:
                emit(fmt.lg_done if deleted > 0 else fmt.lg_nothing, count=deleted)
            return
        try:
            all_files = os.listdir(folder)
            emit(fmt.lg_total_files, count=len(all_files))
            log_files = [f for f in all_files if f.endswith('.log')]
            emit(fmt.lg_files_found, count=len(log_files))
            deleted = _scan_and_delete(folder, log_files, now, cutoff, fmt, emit)
            emit(fmt.lg_done if deleted > 0 else fmt.lg_nothing, count=deleted)
        except Exception as e:  # job log line, cleanup ends (Headless)
            emit(fmt.lg_list_error, error=e)
    except Exception as e:  # work-unit end: becomes the job log error line(s)
        _log.warning("event=retention_logs_failed folder=%s error=%s", folder, e)
        emit(fmt.lg_error, error=e)
        emit(fmt.lg_traceback, tb=traceback.format_exc())
