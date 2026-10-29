"""Backup engine: size total -> profile tree copy -> targets -> retention (ADR 0002).

Opening the job log, preparing the backup folder (name, same-name delete, mkdir)
and checking that the user exists stay with the entry wrapper (ADR 0002 D2):
`run_backup` gets an open JobLog, the JobLogFormat and the backup folder path.
GUI / Headless differences come only from BackupPolicy flags (behaviour) and
JobLogFormat items (wording, line present or not, lead/end).
"""
from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass
from typing import Any, Callable, Iterable, List, Optional

from .fileops import COPY_TIMEOUT_SECONDS, compute_total_size, copy_file_with_progress
from .filters import SKIP_FILTERED, is_skipped
from .joblog import JobLog, JobLogFormat, LogItem, format_bytes
from .model import BackupPolicy, BackupResult
from .ports import BackupStore, BackupTarget, Reporter, TargetContext
from .retention import cleanup_old_backups, cleanup_old_logs

_log = logging.getLogger(__name__)

STATUS_EVERY_FOLDERS = 5   # GUI "Processing folder N" status
STATUS_EVERY_FILES = 10    # GUI "Copied N files" log + status


class InsufficientSpaceError(Exception):
    """Total size is larger than the store's free space (policy.check_free_space)."""

    def __init__(self, required: int, free: int) -> None:
        super().__init__(f"Required={required} Free={free}")
        self.required = required
        self.free = free


@dataclass
class _Run:
    """State of one engine run, passed between the step functions."""

    job: Any
    source: str
    backup_path: str
    log: JobLog
    fmt: JobLogFormat
    reporter: Reporter
    policy: BackupPolicy
    clock: Callable[[], Any]
    files: int = 0
    folders: int = 0
    ok_bytes: int = 0
    progress_bytes: int = 0

    @property
    def total_bytes(self) -> int:
        return self.progress_bytes if self.policy.count_attempted_copies else self.ok_bytes


def _say(send: Callable[[str], None], item: Optional[LogItem], **values: Any) -> None:
    """Send one Reporter message; a None item sends nothing."""
    if item is not None:
        send(item.line(**values).message)


def _skip_reason(run: _Run, path: str, is_dir: bool) -> Optional[str]:
    """is_skipped for this run. A failing check means "include" only with
    policy.skip_errors_include; otherwise the error propagates to the caller."""
    try:
        return is_skipped(path, run.job, run.policy, is_dir=is_dir)
    except Exception:
        if run.policy.skip_errors_include:
            return None
        raise


def _measure(run: _Run, store: BackupStore) -> Optional[int]:
    """Total size step; checks free space when policy.check_free_space (also
    with the partial total after a size error, as the GUI did).
    Returns the total, or None when it could not be computed."""
    log, fmt = run.log, run.fmt
    _say(run.reporter.status, fmt.rp_size_start)
    errors: List[Exception] = []
    total, _count = compute_total_size(
        run.source, lambda p, is_dir: _skip_reason(run, p, is_dir) is not None,
        on_progress=lambda n, t: _say(run.reporter.status, fmt.rp_size_progress,
                                      count=n, size=format_bytes(t)),
        on_walk_error=lambda e: log.emit(fmt.size_walk_error, error=e),
        on_file_error=lambda p, e: log.emit(fmt.size_file_error, path=p, error=e),
        on_error=errors.append)
    if errors:
        log.emit(fmt.size_failed, error=errors[0])
        _log.debug("size total failed error=%s", type(errors[0]).__name__)
        run.reporter.progress_max(total)
        log.emit(fmt.size_total_after_error, size=format_bytes(total))
    else:
        run.reporter.progress_max(total)
        log.emit(fmt.size_total, size=format_bytes(total))
    if run.policy.check_free_space:
        _check_space(run, store, total)
    return None if errors else total


def _check_space(run: _Run, store: BackupStore, total: int) -> None:
    _say(run.reporter.status, run.fmt.rp_space_check)
    free = store.free_space()
    if total > free:
        run.log.emit(run.fmt.space_short, required=total, free=free)
        raise InsufficientSpaceError(total, free)


def _filter_dirs(run: _Run, dirpath: str, dirnames: List[str]) -> None:
    """Prune dirnames in place, logging each skipped directory."""
    fmt = run.fmt
    original = list(dirnames)
    dirnames.clear()
    for d in original:
        full = os.path.join(dirpath, d)
        try:
            reason = _skip_reason(run, full, True)
        except Exception as e:
            # Work-unit boundary: the directory is kept when its check fails.
            run.log.emit(fmt.dir_check_error, path=full, error=e)
            dirnames.append(d)
            continue
        if reason is None:
            dirnames.append(d)
        elif reason == SKIP_FILTERED and run.policy.report_filtered_separately:
            run.log.emit(fmt.skip_filtered_dir, path=full)
        else:
            run.log.emit(fmt.skip_hidden_dir, path=full)


def _make_dest_dir(run: _Run, dirpath: str) -> Optional[str]:
    """Create the destination folder of dirpath; None when that fails (logged)."""
    rel_dir = os.path.relpath(dirpath, run.source)
    dest = os.path.join(run.backup_path, rel_dir) if rel_dir != "." else run.backup_path
    try:
        os.makedirs(dest, exist_ok=True)
        run.log.emit(run.fmt.folder_created, path=dest)
        run.folders += 1
        if run.folders % STATUS_EVERY_FOLDERS == 0:
            _say(run.reporter.status, run.fmt.rp_folder, count=run.folders,
                 name=os.path.basename(dirpath))
    except Exception as e:
        # Work-unit boundary: this folder's files are skipped, the walk goes on.
        run.log.emit(run.fmt.folder_failed, path=dest, error=e)
        return None
    return dest


def _progress(run: _Run, delta: int) -> None:
    run.progress_bytes += delta
    run.reporter.progress(run.progress_bytes)


def _copy_one(run: _Run, src: str, dst: str) -> None:
    """Skip check, copy and count of one file. Errors propagate (file_failed)."""
    log, fmt = run.log, run.fmt
    reason = _skip_reason(run, src, False)
    if reason is not None:
        separate = reason == SKIP_FILTERED and run.policy.report_filtered_separately
        log.emit(fmt.skip_filtered_file if separate else fmt.skip_hidden_file, path=src)
        return
    log.emit(fmt.copying, src=src, dst=dst)

    def on_error(e: Exception) -> None:
        log.emit(fmt.copy_error, src=src, dst=dst, error=e)
        _say(run.reporter.log, fmt.rp_copy_error, src=src, dst=dst, error=e)

    copied = copy_file_with_progress(
        src, dst, lambda delta: _progress(run, delta), timeout_seconds=COPY_TIMEOUT_SECONDS,
        clock=run.clock, on_timeout=lambda s: log.emit(fmt.copy_timeout, path=s), on_error=on_error)
    if run.policy.count_attempted_copies:
        run.files += 1
    elif copied > 0:
        run.files += 1
        run.ok_bytes += copied
    else:
        return
    if run.files % STATUS_EVERY_FILES == 0:
        _say(run.reporter.log, fmt.rp_files_log, count=run.files)
        _say(run.reporter.status, fmt.rp_files_status, count=run.files)


def _copy_dir(run: _Run, dirpath: str, dirnames: List[str], filenames: List[str]) -> None:
    run.log.emit(run.fmt.processing_dir, path=dirpath, rel=os.path.relpath(dirpath, run.source))
    _filter_dirs(run, dirpath, dirnames)
    dest = _make_dest_dir(run, dirpath)
    if dest is None:
        return
    for f in filenames:
        src = os.path.join(dirpath, f)
        try:
            _copy_one(run, src, os.path.join(dest, f))
        except Exception as e:
            # Work-unit boundary: one file fails, the backup goes on.
            run.log.emit(run.fmt.file_failed, path=src, error=e)


def _copy_profile(run: _Run) -> Optional[Exception]:
    """Walk the profile tree and copy it. An error outside the per-directory
    boundary is logged as CRITICAL, then returned (policy.return_critical_errors)
    or raised. Returns None when the walk completed."""
    _say(run.reporter.status, run.fmt.rp_starting)
    try:
        for dirpath, dirnames, filenames in os.walk(run.source, topdown=True, onerror=None):
            try:
                _copy_dir(run, dirpath, dirnames, filenames)
            except Exception as e:
                # Work-unit boundary: this directory fails, the walk goes on.
                run.log.emit(run.fmt.dir_failed, path=dirpath, error=e)
        run.reporter.progress(run.progress_bytes)
        run.log.emit(run.fmt.completed, files=run.files, folders=run.folders,
                     size=format_bytes(run.total_bytes))
    except Exception as e:
        run.log.emit(run.fmt.critical, error=e)
        if run.policy.return_critical_errors:
            return e
        raise
    return None


def _run_targets(run: _Run, targets: Iterable[BackupTarget]) -> None:
    ctx = TargetContext(backup_path=run.backup_path, dest=run.job.dest, fmt=run.fmt,
                        policy=run.policy, out=lambda line: run.log.write(*line),
                        profile=run.source)
    for target in targets:
        target.backup(ctx)


_NO_COUNT = object()  # retention count could not be read (skip old-backup cleanup)


def _retention_count(run: _Run) -> Any:
    """job.retention_count, read when the cleanup starts. With
    policy.isolate_retention_errors a failed read is logged (bk_error) and old-backup
    cleanup is skipped, log cleanup still runs (old GUI read its setting here)."""
    if not run.policy.isolate_retention_errors:
        return run.job.retention_count
    try:
        return run.job.retention_count
    except Exception as e:
        # Work-unit boundary (policy): no count, no backup deletion.
        run.log.emit(run.fmt.bk_error, error=e)
        return _NO_COUNT


def _cleanup(run: _Run, store: BackupStore, now_ts: Callable[[], float]) -> None:
    """Old backups, then old logs. job.log_retention_days is read after the backup
    cleanup; a failed read ends the step (Cleanup failed with isolate_retention_errors)."""
    job = run.job
    out = lambda line: run.log.write(*line)  # noqa: E731 (JobLine -> job log)
    keep = _retention_count(run)
    if keep is not _NO_COUNT:
        cleanup_old_backups(store, job.user, keep, run.fmt, run.policy, out)
    cleanup_old_logs(job.dest, job.log_retention_days, run.fmt, run.policy, out,
                     clock=run.clock, now_ts=now_ts)


def _run_retention(run: _Run, store: BackupStore, now_ts: Callable[[], float]) -> None:
    if not run.policy.isolate_retention_errors:
        _cleanup(run, store, now_ts)
        return
    try:
        _cleanup(run, store, now_ts)
    except Exception as e:
        # Work-unit boundary (policy): a failed cleanup does not fail the backup.
        run.log.emit(run.fmt.cleanup_failed, error=e)
        _log.debug("retention failed error=%s", type(e).__name__)


def run_backup(job: Any, backup_path: str, log: JobLog, fmt: JobLogFormat, store: BackupStore,
               targets: Iterable[BackupTarget], reporter: Reporter, policy: BackupPolicy,
               clock: Callable[[], Any], *, users_root: str, log_path: str = "",
               now_ts: Callable[[], float] = time.time) -> BackupResult:
    """Back up `<users_root>/<job.user>` into the prepared folder `backup_path`.

    `job` is a BackupJob (filters in JSON dict shape, see core CLAUDE.md) or any
    object with user, dest, include_hidden, include_system, filters,
    retention_count and log_retention_days. The source folder must exist
    (checked by the wrapper). Raises InsufficientSpaceError (check_free_space).
    A CRITICAL error of the tree walk is re-raised, or with
    policy.return_critical_errors ends the run and comes back in
    result.critical_error (no targets, no cleanup). `log_path` only goes into
    the result. `now_ts` is the log cleanup reference time (time.time today).
    """
    run = _Run(job=job, source=os.path.join(users_root, job.user), backup_path=backup_path,
               log=log, fmt=fmt, reporter=reporter, policy=policy, clock=clock)
    _log.debug("backup engine start")
    _measure(run, store)
    critical = _copy_profile(run)
    if critical is not None:
        _log.debug("backup engine critical error=%s", type(critical).__name__)
        return BackupResult(backup_path=backup_path, log_path=log_path, files=run.files,
                            folders=run.folders, bytes=run.total_bytes, critical_error=critical)
    _run_targets(run, targets)
    _run_retention(run, store, now_ts)
    log.emit(fmt.finished)
    _log.debug("backup engine done files=%d folders=%d bytes=%d",
               run.files, run.folders, run.total_bytes)
    return BackupResult(backup_path=backup_path, log_path=log_path, files=run.files,
                        folders=run.folders, bytes=run.total_bytes)
