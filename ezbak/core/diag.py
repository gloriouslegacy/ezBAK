"""Diagnostic log set-up and the run_id of a job (ADR 0003 §2, P6).

Diagnostic lines (`_log = logging.getLogger(__name__)` in every module) go to one
rotating file, %APPDATA%\\ezBAK\\logs\\ezbak.log (1 MB x 3), with the format
`%(asctime)s %(levelname)-5s %(name)s [%(run_id)s] %(message)s`. The job logs in the
destination folder do not use logging (core/joblog.py) and never carry the run_id.

setup() is called once by the entry point. A job binds a new run_id when it starts and
binds it again at the top of its work thread (threads do not inherit contextvars).
Only the `ezbak` logger is configured; the root logger and other libraries are left alone.
"""
import contextvars
import logging
import logging.handlers
import os
import random
import sys

from ezbak.winsys import paths

LOGGER_NAME = "ezbak"
LOG_FILE = "ezbak.log"
MAX_BYTES = 1024 * 1024
BACKUP_COUNT = 3
FORMAT = "%(asctime)s %(levelname)-5s %(name)s [%(run_id)s] %(message)s"
LEVEL_ENV = "EZBAK_LOG_LEVEL"
DEFAULT_LEVEL = logging.INFO
NO_RUN_ID = "-"

_run_id = contextvars.ContextVar("ezbak_run_id", default=NO_RUN_ID)


def default_log_dir():
    """%APPDATA%\\ezBAK\\logs, or '' when APPDATA is not set."""
    appdata = paths.appdata_dir()
    return os.path.join(appdata, "ezBAK", "logs") if appdata else ""


def new_run_id(clock, rand=None):
    """`YYYYMMDD-HHMMSS-xxxx` from clock() (a datetime) and 4 random hex digits."""
    rand = rand or random.Random()
    return f"{clock():%Y%m%d-%H%M%S}-{rand.getrandbits(16):04x}"


def bind(run_id):
    """Make run_id the current one in this thread/context; returns it."""
    _run_id.set(run_id)
    return run_id


def current():
    """The bound run_id, or '-' when no job is running."""
    return _run_id.get()


class RunIdFilter(logging.Filter):
    """Puts the current run_id on every record the handler writes."""

    def filter(self, record):
        record.run_id = current()
        return True


def level_from_env(environ=None):
    """EZBAK_LOG_LEVEL (DEBUG, INFO, WARNING, ...) or INFO when unset or unknown."""
    name = (environ if environ is not None else os.environ).get(LEVEL_ENV, "").strip().upper()
    level = logging.getLevelName(name) if name else DEFAULT_LEVEL
    return level if isinstance(level, int) else DEFAULT_LEVEL


def _existing_handler(logger):
    for handler in logger.handlers:
        if getattr(handler, "_ezbak_diag", False):
            return handler
    return None


def setup(*, log_dir=None, level=None, environ=None):
    """Attach the rotating file handler to the `ezbak` logger (once).

    Returns the handler, or None when the folder cannot be made or opened: diagnostics
    must never stop a backup, so the failure is one line on stderr (when there is one).
    """
    logger = logging.getLogger(LOGGER_NAME)
    handler = _existing_handler(logger)
    if handler is not None:
        return handler
    log_dir = default_log_dir() if log_dir is None else log_dir
    if not log_dir:
        return None
    try:
        os.makedirs(log_dir, exist_ok=True)
        handler = logging.handlers.RotatingFileHandler(
            os.path.join(log_dir, LOG_FILE), maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT,
            encoding="utf-8")
    except OSError as e:
        if sys.stderr is not None:  # --windowed exe has no stderr
            sys.stderr.write(f"ezBAK: diagnostic log disabled ({type(e).__name__}: {e})\n")
        return None
    handler._ezbak_diag = True
    handler._ezbak_prev_level = logger.level  # teardown puts it back
    handler.setFormatter(logging.Formatter(FORMAT))
    handler.addFilter(RunIdFilter())
    logger.addHandler(handler)
    logger.setLevel(level_from_env(environ) if level is None else level)
    return handler


def teardown():
    """Remove and close the handler setup() attached and restore the logger level."""
    logger = logging.getLogger(LOGGER_NAME)
    handler = _existing_handler(logger)
    if handler is not None:
        logger.removeHandler(handler)
        logger.setLevel(handler._ezbak_prev_level)
        handler.close()
