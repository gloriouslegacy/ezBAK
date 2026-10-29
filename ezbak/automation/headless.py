"""Headless backup entry (scheduled task / CLI): run flow, folder prep, job log, engine call.

Folder preparation, log opening and the user check belong here, not in the engine
(ADR 0002 D2). The engine is referenced through this module's global
`engine_run_backup`, so the runpy harness stub (tests/fixtures/cli_harness.py) and
unit tests can replace it. Time comes only from `clock` (called at run time) and
`users_root` is passed by the caller, so tests can patch both in ezBAK.py.
"""
import copy
import json
import logging
import os
import time
import traceback

from ezbak.core import BackupJob, BackupPolicy, DEFAULT_FILTERS, diag, history
from ezbak.core import JobLog, JobLogFormat, header_lines, rmtree_readonly
from ezbak.core import run_backup as engine_run_backup
from ezbak.storage import LocalFolderStore
from ezbak.targets import default_targets

from . import schedule_config

_log = logging.getLogger(__name__)


class _HeadlessReporter:
    """ezbak.core.Reporter for Headless runs: messages (copy errors) go to stdout;
    status and progress are dropped.

    The class keeps its pre-split name because the CLI golden records the reporter
    as "<_HeadlessReporter>"; HeadlessReporter is the public name.
    """

    def log(self, message):
        print(message)

    def status(self, message):
        pass  # Headless has no status line

    def progress(self, value):
        pass  # Headless has no progress bar

    def progress_max(self, value):
        pass  # Headless has no progress bar


HeadlessReporter = _HeadlessReporter


def _prepare_backup_folder(user_name, dest_dir, users_root, clock):
    """Name the backup folder and log, recreate the folder, check the user exists.

    Returns (backup_path, log_path). Raises FileNotFoundError (no log written)
    when users_root/user_name is missing; the folder is created before the check.
    """
    timestamp = clock().strftime("%Y-%m-%d_%H%M%S")
    backup_path = os.path.join(dest_dir, f"{user_name}_backup_{timestamp}")
    log_timestamp = clock().strftime("%Y-%m-%d_%H%M%S")
    log_path = os.path.join(dest_dir, f"{user_name}_backup_{log_timestamp}.log")

    # Delete existing backup folder if exists (failure ignored, as before)
    if os.path.exists(backup_path):
        try:
            rmtree_readonly(backup_path)
        except Exception:  # current behaviour: makedirs below copes with leftovers
            pass

    os.makedirs(backup_path, exist_ok=True)
    user_home = os.path.join(users_root, user_name)
    if not os.path.exists(user_home):
        raise FileNotFoundError(f"User folder '{user_home}' not found.")
    return backup_path, log_path


def run_backup(user_name, dest_dir, include_hidden=False, include_system=False,
               log_folder=None, retention_count=2, log_retention_days=30, filters=None,
               *, users_root, clock):
    """Headless backup: prepares the folder and the job log, then runs the core engine.

    filters None means the default filters (a deep copy of DEFAULT_FILTERS).
    Returns (backup_path, log_path).
    """
    backup_path, log_path = _prepare_backup_folder(user_name, dest_dir, users_root, clock)
    _log.info("event=job_start kind=backup mode=headless job_log=%s", log_path)
    job_filters = filters if filters is not None else copy.deepcopy(DEFAULT_FILTERS)

    job = BackupJob(user=user_name, dest=dest_dir, include_hidden=include_hidden,
                    include_system=include_system, retention_count=retention_count,
                    log_retention_days=log_retention_days, filters=job_filters,
                    log_folder=log_folder)
    with open(log_path, "w", encoding="utf-8") as log:
        # Header: start, Policy, Filters, separator (ezbak.core.joblog)
        job_log = JobLog(log, clock=clock)
        job_log.write_lines(header_lines(
            JobLogFormat.HEADLESS,
            start={'user': user_name, 'backup_path': backup_path},
            hidden='include' if include_hidden else 'exclude',
            system='include' if include_system else 'exclude',
            backup_keep=retention_count, log_keep=log_retention_days,
            filters=job_filters))
        # Size, profile tree copy, targets and cleanup (ezbak.core.engine)
        engine_run_backup(job, backup_path, job_log, JobLogFormat.HEADLESS,
                          LocalFolderStore(dest_dir, clock=clock), default_targets(),
                          HeadlessReporter(), BackupPolicy.HEADLESS, clock,
                          users_root=users_root, log_path=log_path,
                          now_ts=lambda: time.time())
    history.record_backup(user=user_name, dest=dest_dir, folder=backup_path, when=clock(),
                          mode='headless')  # home card of the GUI (U7-2)

    return backup_path, log_path


def _run_config(args, users_root, clock):
    """--schedule-config run: 0 on success, 1 on a missing file or any error."""
    timestamp = clock().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] Scheduled backup start using config: {args.schedule_config}")
    try:
        if not os.path.exists(args.schedule_config):
            print(f"[{timestamp}] ERROR: Configuration file not found: {args.schedule_config}")
            return 1
        config = schedule_config.load(args.schedule_config)

        print(f"[{timestamp}] Configuration loaded successfully")
        print(f"[{timestamp}] Task: {config.get('task_name')}")
        print(f"[{timestamp}] User: {config.get('user')}")
        print(f"[{timestamp}] Destination: {config.get('dest')}")
        print(f"[{timestamp}] Retention: {config.get('retention_count')} backups, "
              f"{config.get('log_retention_days')} days logs")

        job = schedule_config.job_from_config(config, log_folder=args.log_folder)
        backup_path, log_path = _run_job(job, users_root, clock)
        print(f"[{timestamp}] Backup finished → {backup_path}")
        print(f"[{timestamp}] Log saved at {log_path}")
    except Exception as e:  # boundary: report on stdout + traceback, exit code 1
        _log.error("event=job_failed kind=backup mode=config", exc_info=True)
        print(f"[{timestamp}] ERROR during scheduled backup: {e}")
        traceback.print_exc()
        return 1
    return 0


def _run_direct(args, users_root, clock):
    """--user/--dest run: always 0 (errors are printed without a traceback, as before)."""
    timestamp = clock().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] Headless backup start for user '{args.user}'")
    print(f"[{timestamp}] Command line args: retention={args.retention}, "
          f"log_retention={args.log_retention}")

    filters = None
    if args.filters:
        try:
            filters = json.loads(args.filters)
            print(f"[{timestamp}] Filters loaded: {filters}")
        except Exception as e:  # recovery: run with the default filters
            print(f"[{timestamp}] WARNING: Failed to parse filters JSON: {e}")
            filters = None

    job = BackupJob(user=args.user, dest=args.dest, include_hidden=args.include_hidden,
                    include_system=args.include_system, retention_count=args.retention,
                    log_retention_days=args.log_retention, filters=filters,
                    log_folder=args.log_folder)
    try:
        backup_path, log_path = _run_job(job, users_root, clock)
        print(f"[{timestamp}] Backup finished → {backup_path}")
        print(f"[{timestamp}] Log saved at {log_path}")
    except Exception as e:  # terminal: current behaviour prints and still exits 0
        _log.error("event=job_failed kind=backup mode=direct", exc_info=True)
        print(f"[{timestamp}] ERROR during headless backup: {e}")
    return 0


def _run_job(job, users_root, clock):
    """Call this module's run_backup (looked up at call time so tests can replace it)."""
    return run_backup(job.user, job.dest, include_hidden=job.include_hidden,
                      include_system=job.include_system, log_folder=job.log_folder,
                      retention_count=job.retention_count,
                      log_retention_days=job.log_retention_days, filters=job.filters,
                      users_root=users_root, clock=clock)


def run(args, *, users_root, clock):
    """Run the Headless flow for parsed CLI args (ezbak.automation.cli.parse).

    Returns the exit code (config run: 0/1, direct --user/--dest run: always 0), or
    None when no Headless run was requested (e.g. --user alone): the caller starts the GUI.
    `clock` gives the stdout [timestamp], read once at the start of a run.
    """
    if not (args.schedule_config or (args.user and args.dest)):
        return None
    diag.setup()  # diagnostic log (ADR 0003 §2); a Headless run is one job, one run_id
    diag.bind(diag.new_run_id(clock))
    if args.schedule_config:
        return _run_config(args, users_root, clock)
    return _run_direct(args, users_root, clock)
