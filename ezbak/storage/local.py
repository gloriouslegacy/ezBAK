"""Local folder backup store (local disk, external drive or UNC path).

Names and listing rules are C3-fixed and match both current paths of ezBAK.py
(GUI App._cleanup_old_backups and Headless core_run_backup), which are identical:
- folder name `<user>_backup_<YYYY-MM-DD_HHMMSS>`, log `<same>.log` in dest
- list: entries of dest that are directories and whose lower-cased name starts with
  `<user lower>_backup_`, sorted by plain name order
"""
import logging
import os
from datetime import datetime
from typing import Callable, List, Optional

from ezbak.core import rmtree_readonly
from ezbak import winsys

_log = logging.getLogger(__name__)

NAME_TIMESTAMP_FORMAT = "%Y-%m-%d_%H%M%S"


class LocalFolderStore:
    """core.ports.BackupStore over a plain folder `dest`."""

    def __init__(self, dest: str, clock: Callable[[], datetime] = datetime.now) -> None:
        self.dest = dest
        self._clock = clock

    def _new_name(self, user_name: str) -> str:
        return f"{user_name}_backup_{self._clock().strftime(NAME_TIMESTAMP_FORMAT)}"

    def new_backup_path(self, user_name: str) -> str:
        """Path for a new backup folder. Reads the clock once; does not create it."""
        return os.path.join(self.dest, self._new_name(user_name))

    def new_log_path(self, user_name: str) -> str:
        """Path for the job log. Reads the clock on its own, as both current paths do."""
        return os.path.join(self.dest, self._new_name(user_name) + ".log")

    def list_backups(self, user_name: str) -> List[str]:
        """Backup folder names of user_name, oldest first (name order)."""
        prefix = f"{user_name.lower()}_backup_"
        names = [d for d in os.listdir(self.dest)
                 if os.path.isdir(os.path.join(self.dest, d)) and d.lower().startswith(prefix)]
        names.sort()
        return names

    def path_of(self, name: str) -> str:
        return os.path.join(self.dest, name)

    def delete_backup(self, name: str, on_log: Optional[Callable[[str], None]] = None) -> None:
        """Remove backup `name` including read-only entries.

        Per-entry failures are handled by core.rmtree_readonly as today: messages go to
        `on_log` (job log wording unchanged) and are not raised. Anything else propagates.
        """
        path = self.path_of(name)
        _log.info("event=delete_backup path=%s", path)
        rmtree_readonly(path, on_log)

    def free_space(self) -> int:
        return winsys.free_space(self.dest)
