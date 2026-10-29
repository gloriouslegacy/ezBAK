"""Architecture checks a, b, c of ADR 0003 §3 and the ADR 0002 name-collision checks.

Only the standard library `ast` is used (no extra dev dependency, ADR 0003 §3 g).

- a: tkinter and customtkinter (and their submodules) are imported only under
     ezbak/gui/ (customtkinter: ADR 0004).
- b: imports between the six modules follow the ADR 0002 dependency table;
     core -> winsys only fileattrs / paths; another module is used only through
     its public names (`ezbak.<mod>` / `from ezbak.<mod> import <name>`), never
     `ezbak.<mod>.<file>`. A submodule listed in `ezbak.<mod>.__all__` is a public
     name: `from ezbak.<mod> import <sub>` and `import ezbak.<mod>.<sub>` are allowed
     (ADR 0002, P5-17 decision A); `from ezbak.<mod>.<sub> import x` is not.
     `__all__` is read with ast from the package __init__.py (no import).
     winsys is the exception: all its submodules are public entry points
     (`from ezbak.winsys import fileattrs`, `from ezbak.winsys.proc import run`).
- c: subprocess is imported only by ezbak/winsys/proc.py.
- names: no module or package inside ezbak/ shadows a standard library name, and
  `import ezbak` / `import ezBAK` resolve to the package folder / the root file.

Lazy imports inside functions are checked too (the whole AST is walked).

The root ezBAK.py is checked too: c since P4-7 (decision D7), a and b since P5-17
with the edges ezBAK.py -> automation, gui, winsys (ADR 0002).
- d (P6-6, decision D2 B): inside ezbak/ no `print` (a call or the name used as a
  value, e.g. `debug=print`) except automation/headless.py, whose stdout lines are a
  C3 output, and no bare `except:`. Diagnostics use logging (ADR 0003 §2).
"""
import ast
import importlib
import os
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
PACKAGE = "ezbak"
ROOT_FILE = "ezBAK.py"

MODULES = ("core", "targets", "storage", "automation", "winsys", "gui")

# ADR 0002 dependency direction: source module -> modules it may import.
ALLOWED_EDGES = {
    "gui": {"automation", "targets", "storage", "core", "winsys"},
    "automation": {"targets", "storage", "core", "winsys"},
    "targets": {"core", "winsys"},
    "storage": {"core", "winsys"},
    "core": {"winsys"},
    "winsys": set(),
    ROOT_FILE: {"automation", "gui", "winsys"},
}
# core -> winsys is limited to these winsys submodules.
CORE_WINSYS_ALLOWED = {"fileattrs", "paths"}
# Modules whose public entry points are their submodules.
SUBMODULE_PUBLIC = {"winsys"}

TKINTER_NAMES = {"tkinter", "_tkinter", "customtkinter"}
SUBPROCESS_ALLOWED = ("winsys", "proc.py")
PRINT_ALLOWED = {("automation", "headless.py")}  # Headless stdout is C3 (ADR 0003 §2)


# --------------------------------------------------------------------------- scanner

def _iter_py_files(pkg_dir):
    for path in sorted(pkg_dir.rglob("*.py")):
        if "__pycache__" not in path.parts:
            yield path


def _package_parts(pkg_dir, path):
    """Dotted package of `path` as a list, e.g. ['ezbak', 'core'] for ezbak/core/x.py."""
    rel = path.relative_to(pkg_dir.parent)
    return list(rel.parts[:-1])


def _imported_targets(node, pkg_parts):
    """Yield (dotted target, display name) for every name an import node brings in.

    `from X import a, b` yields X.a and X.b so `from ezbak import core` and
    `from ezbak.winsys import proc` are seen as the module they name.
    """
    if isinstance(node, ast.Import):
        for alias in node.names:
            yield alias.name, alias.name
        return
    if node.level:
        base = pkg_parts[: len(pkg_parts) - (node.level - 1)] if node.level > 1 else list(pkg_parts)
        prefix = ".".join(base + ([node.module] if node.module else []))
        shown = "." * node.level + (node.module or "")
    else:
        prefix = node.module or ""
        shown = prefix
    for alias in node.names:
        target = f"{prefix}.{alias.name}" if alias.name != "*" else prefix
        yield target, f"{shown} import {alias.name}"


def _source_module(pkg_dir, path):
    rel = path.relative_to(pkg_dir).parts
    return rel[0] if len(rel) > 1 and rel[0] in MODULES else None


def _is_internal_file(pkg_dir, mod, name):
    """True when `name` is a file or subpackage inside ezbak/<mod>/ (not a public name).

    Compared with the exact directory entries: Windows paths ignore case, so a
    public `JobLog` must not match the file `joblog.py`.
    """
    mod_dir = pkg_dir / mod
    if not mod_dir.is_dir():
        return False
    entries = os.listdir(mod_dir)
    if f"{name}.py" in entries and (mod_dir / f"{name}.py").is_file():
        return True
    return name in entries and (mod_dir / name).is_dir()


def _package_all(pkg_dir, mod):
    """Names in the module-level `__all__` of ezbak/<mod>/__init__.py, read with ast (no import).

    Handles `__all__ = [...]` / `(...)`, `__all__: list = [...]` and `__all__ += [...]`;
    entries that are not string literals are ignored.
    """
    init = pkg_dir / mod / "__init__.py"
    if not init.is_file():
        return set()
    names = set()
    for node in ast.parse(init.read_text(encoding="utf-8-sig"), filename=str(init)).body:
        if isinstance(node, ast.Assign):
            targets, value = node.targets, node.value
        elif isinstance(node, (ast.AnnAssign, ast.AugAssign)):
            targets, value = [node.target], node.value
        else:
            continue
        if not any(isinstance(t, ast.Name) and t.id == "__all__" for t in targets):
            continue
        if isinstance(value, (ast.List, ast.Tuple)):
            names.update(e.value for e in value.elts
                         if isinstance(e, ast.Constant) and isinstance(e.value, str))
    return names


def _check_edge(pkg_dir, src, target_parts, is_from_names):
    """Return the rule violated by an ezbak.* import from module `src`, or None."""
    if len(target_parts) < 2:
        return None  # bare `import ezbak`
    dst = target_parts[1]
    if dst not in MODULES:
        return None  # names in ezbak/__init__.py; nothing to check yet
    if src is None or dst == src:
        return None
    if dst not in ALLOWED_EDGES[src]:
        return f"b: 의존 방향 {src} -> {dst} 금지"
    depth = len(target_parts)
    if dst in SUBMODULE_PUBLIC:
        sub = target_parts[2] if depth > 2 else None
        if depth > 4 or (depth == 4 and not is_from_names):
            return f"b: {dst} 내부 파일 직접 import"
        if src == "core" and sub not in CORE_WINSYS_ALLOWED:
            return "b: core -> winsys 는 fileattrs·paths 만"
        return None
    # `from ezbak.<mod> import <name>` reaches depth 3 through a public name;
    # a submodule is public only when listed in ezbak.<mod>.__all__ (decision A).
    if depth > 3:
        return f"b: {dst} 내부 파일 직접 import ({dst}/__init__.py 공개 이름만)"
    if depth < 3:
        return None
    name = target_parts[2]
    internal = _is_internal_file(pkg_dir, dst, name)
    if internal and name in _package_all(pkg_dir, dst):
        return None
    if not is_from_names:
        return f"b: {dst} 내부 파일 직접 import ({dst}/__init__.py 공개 이름만)"
    if internal:
        return f"b: {dst} 내부 파일 import ({name} 은 __all__ 에 없는 파일)"
    return None


def collect_violations(pkg_dir):
    """Scan every .py under `pkg_dir` (an `ezbak` folder) and list rule violations."""
    pkg_dir = Path(pkg_dir)
    violations = []
    for path in _iter_py_files(pkg_dir):
        rel = path.relative_to(pkg_dir)
        shown_path = (Path(PACKAGE) / rel).as_posix()
        src = _source_module(pkg_dir, path)
        pkg_parts = _package_parts(pkg_dir, path)
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, (ast.Import, ast.ImportFrom)):
                continue
            # `from pkg import name`: name may be a module (depth +1) or a public name.
            is_from = isinstance(node, ast.ImportFrom)
            for target, shown in _imported_targets(node, pkg_parts):
                parts = target.split(".")
                rule = None
                if parts[0] in TKINTER_NAMES and src != "gui":
                    rule = "a: tkinter 는 gui/ 만"
                elif parts[0] == "subprocess" and rel.parts != SUBPROCESS_ALLOWED:
                    rule = "c: subprocess 는 winsys/proc.py 만"
                elif parts[0] == PACKAGE:
                    rule = _check_edge(pkg_dir, src, parts, is_from)
                if rule:
                    violations.append(f"{shown_path}:{node.lineno} import {shown} ({rule})")
    return violations + _collect_root_violations(pkg_dir.parent / ROOT_FILE)


def _collect_root_violations(path):
    """Checks a, b, c for the root ezBAK.py next to the package (edges in ALLOWED_EDGES[ROOT_FILE])."""
    if not path.is_file():
        return []
    pkg_dir = path.parent / PACKAGE
    violations = []
    tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    for node in ast.walk(tree):
        if not isinstance(node, (ast.Import, ast.ImportFrom)) or getattr(node, "level", 0):
            continue
        is_from = isinstance(node, ast.ImportFrom)
        for target, shown in _imported_targets(node, []):
            parts = target.split(".")
            rule = None
            if parts[0] in TKINTER_NAMES:
                rule = "a: tkinter 는 gui/ 만"
            elif parts[0] == "subprocess":
                rule = "c: subprocess 는 winsys/proc.py 만"
            elif parts[0] == PACKAGE:
                rule = _check_edge(pkg_dir, ROOT_FILE, parts, is_from)
            if rule:
                violations.append(f"{ROOT_FILE}:{node.lineno} import {shown} ({rule})")
    return violations


def collect_stdlib_name_clashes(pkg_dir):
    """List modules / packages inside `pkg_dir` named like a standard library module."""
    clashes = []
    for path in sorted(Path(pkg_dir).rglob("*")):
        if "__pycache__" in path.parts:
            continue
        if path.is_dir():
            name = path.name
        elif path.suffix == ".py":
            name = path.stem
        else:
            continue
        if name in sys.stdlib_module_names:
            rel = (Path(PACKAGE) / path.relative_to(pkg_dir)).as_posix()
            clashes.append(f"{rel} (표준 라이브러리 이름 {name})")
    return clashes


def collect_d_violations(pkg_dir):
    """Check d: `print` outside PRINT_ALLOWED and bare `except:` anywhere in `pkg_dir`."""
    pkg_dir = Path(pkg_dir)
    violations = []
    for path in _iter_py_files(pkg_dir):
        rel = path.relative_to(pkg_dir)
        shown_path = (Path(PACKAGE) / rel).as_posix()
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Name) and node.id == "print"
                    and isinstance(node.ctx, ast.Load) and rel.parts not in PRINT_ALLOWED):
                violations.append(f"{shown_path}:{node.lineno} print (d: 진단은 logging)")
            elif isinstance(node, ast.ExceptHandler) and node.type is None:
                violations.append(f"{shown_path}:{node.lineno} except: (d: 맨 except 금지)")
    return violations


def _assert_none(found):
    assert not found, "위반 {} 건:\n".format(len(found)) + "\n".join(found)


# --------------------------------------------------------------------------- real tree

def test_ezbak_imports_follow_rules():
    _assert_none(collect_violations(REPO_ROOT / PACKAGE))


def test_ezbak_has_no_print_or_bare_except():
    _assert_none(collect_d_violations(REPO_ROOT / PACKAGE))


def test_ezbak_has_no_stdlib_name_clash():
    _assert_none(collect_stdlib_name_clashes(REPO_ROOT / PACKAGE))


def test_import_ezbak_is_package_folder():
    mod = importlib.import_module(PACKAGE)
    assert Path(mod.__file__).resolve() == (REPO_ROOT / PACKAGE / "__init__.py").resolve()


def test_import_ezBAK_is_root_file(ezbak_mod):
    assert Path(ezbak_mod.__file__).resolve() == (REPO_ROOT / "ezBAK.py").resolve()
    assert ezbak_mod.__name__ == "ezBAK"


# --------------------------------------------------------------------------- self tests

def _make_tree(root, files):
    pkg = root / PACKAGE
    for rel, text in files.items():
        path = pkg / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
    for mod in MODULES:
        init = pkg / mod / "__init__.py"
        init.parent.mkdir(parents=True, exist_ok=True)
        if not init.exists():
            init.write_text("", encoding="utf-8")
    (pkg / "__init__.py").touch()
    return pkg


ALLOWED_SOURCES = {
    "gui/app.py": "import tkinter\nfrom tkinter import ttk\nimport customtkinter as ctk\nfrom ezbak.core import run_backup\n"
                  "from ezbak import automation\nimport ezbak.storage\nfrom . import bridge\n",
    "gui/bridge.py": "from .app import x\n",
    "core/engine.py": "from ezbak.winsys import fileattrs, paths\nfrom .model import BackupJob\n"
                      "from . import joblog\nimport ezbak.winsys.paths\n",
    "automation/cli.py": "from ezbak.winsys import proc\nfrom ezbak.winsys.proc import run\n"
                         "from ezbak.targets import Bookmarks\nfrom ..core import BackupJob\n",
    "core/restore.py": "from ezbak.core import engine\n",
    "storage/local.py": "",
    "automation/schedule_config.py": "from ezbak.storage import LocalFolderStore\n",
    "winsys/proc.py": "import subprocess\nfrom subprocess import PIPE\n",
    "winsys/fileattrs.py": "import ctypes\nfrom . import paths\n",
}


def test_checker_accepts_allowed_imports(tmp_path):
    pkg = _make_tree(tmp_path, ALLOWED_SOURCES)
    assert collect_violations(pkg) == []
    assert collect_stdlib_name_clashes(pkg) == []


@pytest.mark.parametrize("rel, source, rule", [
    ("core/engine.py", "import tkinter\n", "a: tkinter"),
    ("storage/local.py", "from tkinter import messagebox\n", "a: tkinter"),
    ("targets/browser.py", "def f():\n    import tkinter.ttk\n", "a: tkinter"),
    ("core/engine.py", "import customtkinter\n", "a: tkinter"),
    ("automation/cli.py", "from customtkinter import CTkFrame\n", "a: tkinter"),
    ("core/engine.py", "import subprocess\n", "c: subprocess"),
    ("winsys/elevation.py", "from subprocess import run\n", "c: subprocess"),
    ("automation/cli.py", "def f():\n    import subprocess\n", "c: subprocess"),
    ("core/engine.py", "from ezbak.gui import App\n", "b: 의존 방향 core -> gui"),
    ("core/engine.py", "from ezbak import storage\n", "b: 의존 방향 core -> storage"),
    ("targets/browser.py", "from ezbak.storage import LocalFolderStore\n", "b: 의존 방향 targets -> storage"),
    ("winsys/paths.py", "from ezbak.core import model\n", "b: 의존 방향 winsys -> core"),
    ("core/restore.py", "from ..gui import App\n", "b: 의존 방향 core -> gui"),
    ("storage/local.py", "def f():\n    from ezbak.automation import cli\n", "b: 의존 방향 storage -> automation"),
    ("core/engine.py", "from ezbak.winsys import proc\n", "b: core -> winsys 는 fileattrs·paths 만"),
    ("core/engine.py", "import ezbak.winsys.sysinfo\n", "b: core -> winsys 는 fileattrs·paths 만"),
    ("core/engine.py", "import ezbak.winsys\n", "b: core -> winsys 는 fileattrs·paths 만"),
    ("gui/app.py", "from ezbak.core.engine import run\n", "b: core 내부 파일"),
    ("gui/app.py", "import ezbak.core.model\n", "b: core 내부 파일"),
    ("automation/cli.py", "from ezbak.targets.browser import Bookmarks\n", "b: targets 내부 파일"),
    ("gui/app.py", "from ezbak.winsys.proc.helpers import x\n", "b: winsys 내부 파일"),
    ("gui/app.py", "from ezbak.core import engine\n", "b: core 내부 파일 import (engine"),
    ("gui/app.py", "from ezbak.automation import scheduler\n", "b: automation 내부 파일 import (scheduler"),
    ("gui/app.py", "import ezbak.automation.scheduler\n", "b: automation 내부 파일 직접 import"),
    ("gui/app.py", "from ezbak.automation.cli import parse\n", "b: automation 내부 파일 직접 import"),
    ("storage/local.py", "from ezbak.core import model\n", "b: core 내부 파일 import (model"),
    ("automation/cli.py", "from ezbak.storage import backends\n", "b: storage 내부 파일 import (backends"),
])
def test_checker_reports_violation(tmp_path, rel, source, rule):
    extra = {"core/engine.py": "", "storage/backends/__init__.py": "", "core/model.py": "",
             "core/__init__.py": '__all__ = ["run_backup", MODEL]\nMODEL = "model"\n',
             "automation/__init__.py": '__all__ = ["cli"]\n', "automation/cli.py": "",
             "automation/scheduler.py": ""}
    pkg = _make_tree(tmp_path, {**extra, rel: source})
    found = collect_violations(pkg)
    assert len(found) == 1, found
    assert found[0].startswith(f"ezbak/{rel}:")
    assert f"({rule}" in found[0]


def test_checker_accepts_public_name_differing_from_file_only_in_case(tmp_path):
    pkg = _make_tree(tmp_path, {
        "core/joblog.py": "",
        "gui/app.py": "from ezbak.core import JobLog\n",
    })
    assert collect_violations(pkg) == []


@pytest.mark.parametrize("source, rule", [
    ("from ezbak.core.joblog import JobLog\n", "b: core 내부 파일 직접 import"),
    ("from ezbak.core import joblog\n", "b: core 내부 파일 import (joblog"),
])
def test_checker_reports_internal_file_next_to_case_twin(tmp_path, source, rule):
    pkg = _make_tree(tmp_path, {"core/joblog.py": "", "gui/app.py": source})
    found = collect_violations(pkg)
    assert len(found) == 1, found
    assert f"({rule}" in found[0]


def test_checker_reports_all_violations_with_line(tmp_path):
    pkg = _make_tree(tmp_path, {
        "core/engine.py": "import os\nimport tkinter\n\ndef f():\n    import subprocess\n",
    })
    found = collect_violations(pkg)
    assert found == [
        "ezbak/core/engine.py:2 import tkinter (a: tkinter 는 gui/ 만)",
        "ezbak/core/engine.py:5 import subprocess (c: subprocess 는 winsys/proc.py 만)",
    ]


SUBMODULE_ALL_SOURCES = {
    "automation/__init__.py": '__all__ = ["cli", "headless", "Spec"]\n__all__ += ("scheduler",)\n',
    "automation/cli.py": "",
    "automation/headless.py": "def f():\n    import ezbak.storage.nas\n",
    "automation/scheduler.py": "",
    "storage/__init__.py": '__all__: list = ["nas"]\n',
    "storage/nas.py": "",
    "gui/app.py": "from ezbak.automation import cli, headless, Spec\nimport ezbak.automation.scheduler\n"
                  "from ezbak.storage import nas\n",
}


def test_checker_accepts_submodule_listed_in_all(tmp_path):
    pkg = _make_tree(tmp_path, SUBMODULE_ALL_SOURCES)
    assert collect_violations(pkg) == []


def test_package_all_is_read_without_import(tmp_path):
    pkg = _make_tree(tmp_path, {**SUBMODULE_ALL_SOURCES,
                                "core/__init__.py": "raise RuntimeError('imported')\n__all__ = ['x', NAME]\n"})
    assert _package_all(pkg, "automation") == {"cli", "headless", "Spec", "scheduler"}
    assert _package_all(pkg, "storage") == {"nas"}
    assert _package_all(pkg, "core") == {"x"}
    assert _package_all(pkg, "targets") == set()


def test_checker_accepts_root_allowed_edges(tmp_path):
    pkg = _make_tree(tmp_path, {**SUBMODULE_ALL_SOURCES, "winsys/elevation.py": "", "winsys/paths.py": ""})
    (tmp_path / ROOT_FILE).write_text(
        "import os\nfrom ezbak import gui\nfrom ezbak.winsys import elevation\n"
        "from ezbak.winsys.paths import USERS_ROOT\nfrom ezbak.automation import headless, cli\n",
        encoding="utf-8")
    assert collect_violations(pkg) == []


def test_checker_reports_root_a_b_c(tmp_path):
    """Root ezBAK.py: tkinter (a), a forbidden edge and internal files (b), subprocess (c)."""
    pkg = _make_tree(tmp_path, {**SUBMODULE_ALL_SOURCES, "core/engine.py": ""})
    (tmp_path / ROOT_FILE).write_text(
        "import tkinter\nfrom ezbak.core import run_backup\nfrom ezbak.gui.app import App\n"
        "from ezbak.automation.cli import parse\nimport ezbak.storage\n"
        "\ndef f():\n    from subprocess import run\n    from tkinter import ttk\n",
        encoding="utf-8")
    internal = "내부 파일 직접 import ({0}/__init__.py 공개 이름만)"
    assert collect_violations(pkg) == [
        "ezBAK.py:1 import tkinter (a: tkinter 는 gui/ 만)",
        "ezBAK.py:2 import ezbak.core import run_backup (b: 의존 방향 ezBAK.py -> core 금지)",
        "ezBAK.py:3 import ezbak.gui.app import App (b: gui " + internal.format("gui") + ")",
        "ezBAK.py:4 import ezbak.automation.cli import parse (b: automation " + internal.format("automation") + ")",
        "ezBAK.py:5 import ezbak.storage (b: 의존 방향 ezBAK.py -> storage 금지)",
        "ezBAK.py:8 import subprocess import run (c: subprocess 는 winsys/proc.py 만)",
        "ezBAK.py:9 import tkinter import ttk (a: tkinter 는 gui/ 만)",
    ]


@pytest.mark.parametrize("rel", ["core/logging.py", "platform/__init__.py", "gui/queue.py", "core/copy.py"])
def test_checker_reports_stdlib_name_clash(tmp_path, rel):
    pkg = _make_tree(tmp_path, {rel: ""})
    found = collect_stdlib_name_clashes(pkg)
    assert len(found) == 1, found
    assert Path(rel).stem in found[0] or Path(rel).parent.name in found[0]


@pytest.mark.parametrize("rel, source, expected", [
    ("gui/app.py", "print('DEBUG: x')\n", ["ezbak/gui/app.py:1 print (d: 진단은 logging)"]),
    ("gui/app.py", "def f(debug=print):\n    debug('x')\n", ["ezbak/gui/app.py:1 print (d: 진단은 logging)"]),
    ("core/engine.py", "try:\n    pass\nexcept:\n    pass\n",
     ["ezbak/core/engine.py:3 except: (d: 맨 except 금지)"]),
    ("automation/headless.py", "try:\n    print('ok')\nexcept:\n    pass\n",
     ["ezbak/automation/headless.py:3 except: (d: 맨 except 금지)"]),
])
def test_checker_reports_d(tmp_path, rel, source, expected):
    assert collect_d_violations(_make_tree(tmp_path, {rel: source})) == expected


def test_checker_accepts_d(tmp_path):
    pkg = _make_tree(tmp_path, {
        "automation/headless.py": "print('[ts] Backup finished')\n",
        "gui/app.py": "import logging\n_log = logging.getLogger(__name__)\n"
                      "try:\n    pass\nexcept Exception:\n    _log.debug('x', exc_info=True)\n"
                      "printer = 1\nobj.print = 2\n",
    })
    assert collect_d_violations(pkg) == []
