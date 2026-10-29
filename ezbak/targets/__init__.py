"""Backup targets next to the profile tree (V1). Use only these names."""
from . import apps, drivers
from .browser import (BROWSERS, PROFILE_SOURCES, Bookmarks, Browser, ProfileSource, backup_profiles,
                      restore_bookmarks)


def default_targets():
    """Targets attached to every regular backup run (GUI and Headless)."""
    return [Bookmarks()]


__all__ = ["BROWSERS", "PROFILE_SOURCES", "Bookmarks", "Browser", "ProfileSource", "apps", "backup_profiles",
           "default_targets", "drivers", "restore_bookmarks"]
