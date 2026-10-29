"""External command execution: the only module that imports subprocess (ADR 0002).

Thin pass-through. shell, encoding, creationflags, capture_output, timeout and so on
are chosen by the caller; nothing is added here, so argv stays exactly as built.
"""
import logging
import subprocess

log = logging.getLogger(__name__)

# Re-exported so callers outside proc need not import subprocess (ADR 0002).
TimeoutExpired = subprocess.TimeoutExpired
list2cmdline = subprocess.list2cmdline
PIPE = subprocess.PIPE


def _program(argv):
    """Program name only, for diagnostics (argv may hold credentials, e.g. net use)."""
    if isinstance(argv, (list, tuple)):
        return str(argv[0]) if argv else ""
    return str(argv).split(" ", 1)[0]


def run(argv, **kwargs):
    """subprocess.run(argv, **kwargs) unchanged; returns its CompletedProcess."""
    log.debug("proc.run program=%s", _program(argv))
    return subprocess.run(argv, **kwargs)


def stream(argv, **kwargs):
    """subprocess.Popen(argv, **kwargs) unchanged; returns the Popen object."""
    log.debug("proc.stream program=%s", _program(argv))
    return subprocess.Popen(argv, **kwargs)
