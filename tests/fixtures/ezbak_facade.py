"""Test facade over the loaded ezBAK module (P5 decision D2=A).

ezBAK.py is a thin entry point (P5-16b); tests keep using one object `mod`:

- read  `mod.X`: root ezBAK module -> public names of `ezbak.gui` -> alias table.
- write `mod.X = v` (also pytest monkeypatch and its undo): set on the root module
  (when it has X, or when no gui module has it) and on every loaded `ezbak.gui*`
  module that already has X. An alias name the root lacks goes to its real target
  and also to every loaded `ezbak.gui*` module holding that name.
- `del mod.X`: removed wherever X is defined (root and gui modules), alias target too.

Unknown names raise AttributeError, so `getattr(mod, X, default)` and
monkeypatch's `raising=` checks behave as on a module.
"""
import importlib
import sys

GUI_PACKAGE = "ezbak.gui"

# test-only name -> (module name, attribute) it stands for; attribute None = the module
# itself (imported by name when it is not an ezbak module). Names ezBAK.py used to import.
# USERS_ROOT (P5 D3): ezBAK.py and ezbak.gui read ezbak.winsys.paths.USERS_ROOT at call
# time, so user_root.inject(mod) patches that single definition through this alias.
ALIASES = {
    "_have_pywin32": ("ezbak.winsys.fileattrs", "HAS_PYWIN32"),
    "USERS_ROOT": ("ezbak.winsys.paths", "USERS_ROOT"),
    "DEFAULT_FILTERS": ("ezbak.core", "DEFAULT_FILTERS"),
    "BackupPolicy": ("ezbak.core", "BackupPolicy"),
    "core_is_skipped": ("ezbak.core", "is_skipped"),
    "core_engine_run_backup": ("ezbak.core", "run_backup"),
    "winsys_free_space": ("ezbak.winsys", "free_space"),
    "winsys_sysinfo": ("ezbak.winsys", "sysinfo"),
    "get_system_info": ("ezbak.winsys.sysinfo", "system_name"),
    "resource_path": ("ezbak.winsys.paths", "resource_path"),
    "tk": ("tkinter", None),
    "ttk": ("tkinter.ttk", None),
    "messagebox": ("tkinter.messagebox", None),
    "filedialog": ("tkinter.filedialog", None),
    "simpledialog": ("tkinter.simpledialog", None),
    "os": ("os", None),
    "time": ("time", None),
    "shutil": ("shutil", None),
    "threading": ("threading", None),
}

_MISSING = object()


def _is_gui_name(name):
    return name == GUI_PACKAGE or name.startswith(GUI_PACKAGE + ".")


def capture_ezbak_modules(modules=None):
    """Snapshot {name: module} of the ezbak package currently in `modules` (sys.modules)."""
    modules = sys.modules if modules is None else modules
    return {name: module for name, module in list(modules.items())
            if module is not None and (name == "ezbak" or name.startswith("ezbak."))}


class EzbakFacade:
    """Module-like view of ezBAK; see the module docstring for the lookup rules.

    `packages`: {name: module} of the ezbak.* modules this copy of ezBAK uses.
    None means the live sys.modules (the normal import); the nopywin32 copy passes
    the snapshot taken before its fresh ezbak modules were removed from sys.modules.
    """

    __slots__ = ("_facade_root", "_facade_packages")

    def __init__(self, root, packages=None):
        object.__setattr__(self, "_facade_root", root)
        object.__setattr__(self, "_facade_packages", packages)

    # --- module lookup -------------------------------------------------------

    def _facade_modules(self):
        packages = object.__getattribute__(self, "_facade_packages")
        return capture_ezbak_modules() if packages is None else packages

    def _facade_gui_modules(self):
        mods = self._facade_modules()
        return [mods[name] for name in sorted(mods) if _is_gui_name(name)]

    def _facade_alias_target(self, name):
        module_name, attr = ALIASES[name]
        if module_name == "ezbak" or module_name.startswith("ezbak."):
            module = self._facade_modules().get(module_name)
            if module is None:
                raise AttributeError(f"ezBAK facade alias {name!r}: {module_name} is not loaded")
        else:
            module = importlib.import_module(module_name)
        return module, attr

    def _facade_alias_write(self, name, apply):
        """Alias the root lacks: apply(module, attr) on the target (unless it is a whole
        module) and on every gui module holding `name`; no place at all -> AttributeError."""
        module, attr = self._facade_alias_target(name)
        holders = [m for m in self._facade_gui_modules() if name in vars(m)]
        if attr is None and not holders:
            raise AttributeError(f"ezBAK facade alias {name!r}: no gui module holds it")
        if attr is not None:
            apply(module, attr)
        for holder in holders:
            if holder is not module or attr != name:
                apply(holder, name)

    # --- attribute protocol --------------------------------------------------

    def __getattr__(self, name):
        root = object.__getattribute__(self, "_facade_root")
        value = getattr(root, name, _MISSING)
        if value is not _MISSING:
            return value
        gui = self._facade_modules().get(GUI_PACKAGE)
        if gui is not None and name in getattr(gui, "__all__", ()):
            return getattr(gui, name)
        if name in ALIASES:
            module, attr = self._facade_alias_target(name)
            return module if attr is None else getattr(module, attr)
        raise AttributeError(f"ezBAK facade has no attribute {name!r}")

    def __setattr__(self, name, value):
        root = object.__getattribute__(self, "_facade_root")
        if name in ALIASES and not hasattr(root, name):
            self._facade_alias_write(name, lambda module, attr: setattr(module, attr, value))
            return
        holders = [m for m in self._facade_gui_modules() if name in vars(m)]
        if name in vars(root) or not holders:
            setattr(root, name, value)
        for module in holders:
            setattr(module, name, value)

    def __delattr__(self, name):
        root = object.__getattribute__(self, "_facade_root")
        if name in ALIASES and not hasattr(root, name):
            self._facade_alias_write(name, delattr)
            return
        holders = [m for m in [root, *self._facade_gui_modules()] if name in vars(m)]
        if not holders:
            raise AttributeError(f"ezBAK facade has no attribute {name!r}")
        for module in holders:
            delattr(module, name)

    def __dir__(self):
        root = object.__getattribute__(self, "_facade_root")
        gui = self._facade_modules().get(GUI_PACKAGE)
        names = set(dir(root)) | set(ALIASES)
        names |= set(getattr(gui, "__all__", ()))
        return sorted(names)

    def __repr__(self):
        root = object.__getattribute__(self, "_facade_root")
        return f"<EzbakFacade of {root!r}>"
