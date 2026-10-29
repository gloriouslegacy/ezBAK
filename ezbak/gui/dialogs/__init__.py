"""Dialogs (moved from ezBAK.py in P5-3~15). Use only these names."""
from .message import Win11Dialog
# filters after message: filters -> shortcuts -> `.dialogs.Win11Dialog` must already be bound.
from .filters import FilterManagerDialog
from .sources import SelectSourcesDialog
from .schedule import ScheduleBackupDialog

__all__ = ["FilterManagerDialog", "ScheduleBackupDialog", "SelectSourcesDialog", "Win11Dialog"]
