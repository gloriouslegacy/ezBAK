"""Ports the engine talks through (ADR 0002 interface table).

Only Protocols and the minimal data types they need live here. This module imports
nothing from winsys or any other ezbak module: implementations live in storage/,
targets/ and the GUI / Headless entry wrappers.
"""
from dataclasses import dataclass, field
from typing import Any, Callable, List, Optional, Protocol, runtime_checkable


@runtime_checkable
class Reporter(Protocol):
    """Progress sink for the running job (P2 decision D3).

    GUI adapts these to its message queue; Headless ignores them.
    """

    def log(self, message: str) -> None: ...

    def status(self, message: str) -> None: ...

    def progress(self, value: int) -> None: ...

    def progress_max(self, value: int) -> None: ...


@runtime_checkable
class BackupStore(Protocol):
    """Where backups live (V2). One implementation for now: storage.LocalFolderStore.

    The store only names, lists and removes backups. Opening the job log, removing a
    same-name folder before a run and asking the user stay with the entry wrapper
    (P2 decision D2); retention log lines belong to core retention.
    """

    def new_backup_path(self, user_name: str) -> str:
        """Path for a new backup folder of user_name (not created here)."""
        ...

    def new_log_path(self, user_name: str) -> str:
        """Path for the job log file that goes with a new backup of user_name."""
        ...

    def list_backups(self, user_name: str) -> List[str]:
        """Names of existing backups of user_name, oldest first."""
        ...

    def path_of(self, name: str) -> str:
        """Full path of backup `name` as written in job log lines (display only;
        listing and removal still go through list_backups / delete_backup)."""
        ...

    def delete_backup(self, name: str, on_log: Optional[Callable[[str], None]] = None) -> None:
        """Remove one backup by name. Per-entry delete failures go to on_log as job log
        lines (not raised); other errors propagate to the caller."""
        ...

    def free_space(self) -> int:
        """Free bytes at the store location."""
        ...


@dataclass
class TargetContext:
    """What a BackupTarget gets for one run.

    backup_path: the new backup folder of this run. dest: the destination root the
    backup folders live in (bookmark copies go here, C3). fmt: JobLogFormat and
    policy: BackupPolicy of the calling path (GUI / HEADLESS). out: callback that
    takes one JobLine (as in core.retention); None drops the lines. profile: the
    backed-up user's profile folder (<users_root>/<job.user>); targets read their
    sources from it (bookmarks: <profile>/AppData/Local/..., P2-fix #1).
    """

    backup_path: str = ""
    log: Any = None
    dest: str = ""
    fmt: Any = None
    policy: Any = None
    out: Optional[Callable[[Any], None]] = None
    profile: str = ""


@dataclass
class TargetResult:
    """Outcome of one BackupTarget run. errors holds the error texts logged (ok is
    False when there is any)."""

    name: str
    ok: bool = True
    errors: List[str] = field(default_factory=list)


@runtime_checkable
class BackupTarget(Protocol):
    """One extra thing backed up next to the profile tree (V1), e.g. browser bookmarks."""

    name: str

    def backup(self, ctx: TargetContext) -> TargetResult: ...
