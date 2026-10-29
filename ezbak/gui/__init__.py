"""GUI package. Use only these names.

Entry point: `run(entry_script)` (app.py, P5-16) creates App(entry_script=...) and runs
its mainloop; ezBAK.py calls it. App itself (app.py) is built from the mixins of the
gui modules, which stay internal.

Also public (tests and the test facade read them): the core-port adapters of
App.run_backup (bridge.py: GuiReporter, GuiLogStream, GuiFolderStore, GuiJob, int_or) and
begin_job (run_id and start line of a work thread, P6-3),
Translator (i18n.py), Win11Theme (theme.py), copy_default_filters (settings.py), the
dialogs Win11Dialog, FilterManagerDialog, SelectSourcesDialog, ScheduleBackupDialog
(dialogs/) and the CustomTkinter widgets Card, StatTile, Chip, SidebarItem, RingButton
(components.py, U3-3; Sparkline U7-4), and the submodules dpi (DPI factor helpers for the tk parts, U8) and home_data (home
data cards, U7-5).
"""
from . import dpi
from .app import App, run
from .bridge import GuiFolderStore, GuiJob, GuiLogStream, GuiReporter, begin_job, int_or
from .components import Card, Chip, RingButton, SidebarItem, Sparkline, StatTile
from .dialogs import FilterManagerDialog, ScheduleBackupDialog, SelectSourcesDialog, Win11Dialog
from .i18n import Translator
from .pages import home_data
from .settings import copy_default_filters
from .theme import Win11Theme

__all__ = ["App", "Card", "Chip", "FilterManagerDialog", "GuiFolderStore", "GuiJob", "GuiLogStream", "GuiReporter",
           "RingButton", "ScheduleBackupDialog", "SelectSourcesDialog", "SidebarItem", "Sparkline", "StatTile",
           "Translator", "Win11Dialog", "Win11Theme",
           "begin_job", "copy_default_filters", "dpi", "home_data", "int_or", "run"]
