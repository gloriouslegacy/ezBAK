"""Characterization (golden) of the tool features' cases (P4-0a), on the untouched ezBAK.py.

Covered: get_system_info (registry -> wmic x3 -> "Unknown", clean_name, DEBUG prints),
App.get_used_drive_letters / find_free_drive_letter, driver backup and restore
(start_* -> run_* with the worker inline), and open_device_manager when every way fails.

argv_golden.json already fixes the argv and kwargs of these calls; here the rest is fixed:
the return value, the DEBUG prints, the GUI queue in order, the job logs (<TS> for
[YYYY-mm-dd HH:MM:SS] stamps, <TMP> for the temp root, traceback frames -> <FRAMES>),
the files and folders made, the dialogs (name, title, message) and the command texts.

Nothing real runs: winreg OpenKey/QueryValueEx/CloseKey are fakes (tool_harness.FakeWinreg),
every external command is answered by record_subprocess, kernel32 and os.path.exists are
fakes for the drive letter cases, the clock is fixed_clock and the preferred encoding is
fixed to "<PREFERRED>". Driver folders live under <TMP>/drivers.

To rebuild the golden from the current code (Q1-approved changes only):
    set EZBAK_REGEN_GOLDEN=1 && python -m pytest tests/characterization/test_tool_cases.py
"""
import importlib.util
import locale
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.windows

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
FIXTURES = Path(__file__).resolve().parent.parent / "fixtures"
GOLDEN_PATH = FIXTURES / "tool_cases_golden.json"

_spec = importlib.util.spec_from_file_location("ezbak_test_tool_harness", FIXTURES / "tool_harness.py")
harness = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(harness)

BOARD = "TESTBOARD"
STAMP = "2024-01-02_030405"  # fixed_clock


def _commands(rec):
    return [[c["api"], c["args"][0] if isinstance(c["args"][0], str) else " ".join(c["args"][0])]
            for c in rec.calls]


@pytest.fixture
def env(ezbak_mod, tmp_path, monkeypatch, isolated_env, fixed_clock, record_subprocess):
    # ezbak_mod is requested before fixed_clock so the session import is not made under the patch.
    fixed_clock.inject(ezbak_mod)
    monkeypatch.setattr(locale, "getpreferredencoding", lambda do_setlocale=True: "<PREFERRED>")
    return SimpleNamespace(mod=ezbak_mod, rec=record_subprocess, mp=monkeypatch, tmp=tmp_path,
                           norm=harness.normalizer(tmp_path, REPO_ROOT))


# ---------------------------------------------------------------- get_system_info
# registry: None -> OpenKey fails; dict -> the key's values. wmic: {match: respond kwargs}.

SYSINFO = {
    "sysinfo_registry_hit": {"registry": {"BaseBoardProduct": "PRIME Z490-P"}, "wmic": {}},
    "sysinfo_registry_value_missing": {"registry": {}, "wmic": {"baseboard": {"stdout": "Product=B450M\r\n"}}},
    "sysinfo_oem_skipped": {
        "registry": {"BaseBoardProduct": "To be filled by O.E.M."},
        "wmic": {"baseboard": {"stdout": "Product=To Be Filled By O.E.M.\r\n"},
                 "model": {"stdout": "\r\nModel=System Product Name\r\n\r\n"},
                 "manufacturer": {"stdout": "Manufacturer=ACME Corp\r\n"}}},
    "sysinfo_clean_chars": {
        "registry": None,
        "wmic": {"baseboard": {"stdout": 'Product=  My<Board>:"X"/Y\\Z|?*  Pro \t Max  \r\n'}}},
    "sysinfo_whitespace_only": {
        "registry": {"BaseBoardProduct": "   "},
        "wmic": {"baseboard": {"stdout": "Product=<>:?*\r\n"},
                 "model": {"stdout": "Model=   \r\n"},
                 "manufacturer": {"stdout": "Manufacturer=Dell Inc.\r\n"}}},
    "sysinfo_rc_nonzero": {
        "registry": None,
        "wmic": {"baseboard": {"returncode": 1, "stdout": "Product=IGNORED\r\n"},
                 "model": {"stdout": "Model=OptiPlex 7090\r\n"}}},
    "sysinfo_wmic_raises": {
        "registry": None,
        "wmic": {"baseboard": {"raises": OSError("wmic missing")},
                 "model": {"raises": FileNotFoundError(2, "No such file")},
                 "manufacturer": {"stdout": "Manufacturer=HP\r\n"}}},
    "sysinfo_all_fail_unknown": {
        "registry": None,
        "wmic": {"baseboard": {"returncode": 1},
                 "model": {"raises": OSError("wmic missing")},
                 "manufacturer": {"stdout": "Manufacturer=\r\n"}}},
}


@pytest.mark.parametrize("case", list(SYSINFO))
def test_sysinfo_case(case, env, capsys):
    spec = SYSINFO[case]
    if spec["registry"] is None:
        reg = harness.blocked_registry(env.mp)
    else:
        reg = harness.FakeWinreg(env.mp, values=spec["registry"])
    for match, kwargs in spec["wmic"].items():
        env.rec.respond(match, **kwargs)
    capsys.readouterr()
    result = env.mod.get_system_info()
    out = capsys.readouterr().out.splitlines()
    harness.check_golden(GOLDEN_PATH, case, {
        "result": result, "stdout": out, "registry": reg.calls, "commands": _commands(env.rec)})


# ---------------------------------------------------------------- drive letters

def _bits(letters):
    return sum(1 << (ord(c) - 65) for c in letters)


def _used_letters(mod, monkeypatch, bitmask=None, existing=()):
    """App.get_used_drive_letters with kernel32 (bitmask; None -> raises) and exists faked."""
    probed = []

    def get_logical_drives():
        if bitmask is None:
            raise OSError("GetLogicalDrives failed")
        return bitmask

    def exists(path):
        probed.append(path)
        return path in existing

    kernel32 = SimpleNamespace(GetLogicalDrives=get_logical_drives)
    with pytest.MonkeyPatch.context() as mp:
        # sysinfo.used_drive_letters looks kernel32 up through its own ctypes (P5-15)
        mp.setattr(mod.winsys_sysinfo, "ctypes", SimpleNamespace(windll=SimpleNamespace(kernel32=kernel32)))
        mp.setattr(mod.os.path, "exists", exists)
        used = mod.App.get_used_drive_letters(None)
    return sorted(used), probed


def _free_letter(mod, used):
    return mod.App.find_free_drive_letter(SimpleNamespace(get_used_drive_letters=lambda: set(used)))


def test_drive_letters_bitmask(env):
    used, probed = _used_letters(env.mod, env.mp, bitmask=_bits("ACDZ"))
    harness.check_golden(GOLDEN_PATH, "drive_letters_bitmask", {
        "used": used, "probed": probed, "free": _free_letter(env.mod, used)})


def test_drive_letters_exists_probe(env):
    used, probed = _used_letters(env.mod, env.mp, bitmask=None, existing=("C:\\", "E:\\"))
    harness.check_golden(GOLDEN_PATH, "drive_letters_exists_probe", {
        "used": used, "probed": probed, "free": _free_letter(env.mod, used)})


def test_drive_letters_free_order(env):
    letters = [chr(c) for c in range(65, 91)]
    d_to_z = letters[3:]
    table = {
        "none_used": [],
        "z_used": ["Z"],
        "e_to_z_used": letters[4:],
        "d_to_z_used": d_to_z,
        "c_to_z_used": ["C"] + d_to_z,
        "a_c_to_z_used": ["A", "C"] + d_to_z,
        "all_used": letters,
    }
    harness.check_golden(GOLDEN_PATH, "drive_letters_free_order",
                         {name: _free_letter(env.mod, used) for name, used in table.items()})


# ---------------------------------------------------------------- GUI: drivers, devmgr

@pytest.fixture
def gui(env, gui_app, capsys):
    """env plus the session App (reset), inline worker threads and the dialog recorder."""
    app = gui_app(env.mod)
    harness.inline_threads(env.mod, env.mp)
    env.app, env.dialogs, env.capsys = app, [], capsys
    env.drivers = env.tmp / "drivers"
    env.drivers.mkdir()
    harness.drain(app)
    yield env
    app.close_log_file()
    app.reset_progress_bar()
    app.set_buttons_state("normal")
    harness.drain(app)


def _run_driver(env, start, folder, yes=()):
    harness.patch_dialogs(env.mod, env.mp, env.dialogs, yes=yes, folder=folder)
    env.capsys.readouterr()
    getattr(env.app, start)()
    out = env.capsys.readouterr().out.splitlines()
    return env.norm({
        "stdout": out,
        "dialogs": env.dialogs,
        "queue": harness.drain(env.app),
        "tree": harness.walk(env.drivers),
        "logs": harness.read_logs(env.drivers),
        "commands": _commands(env.rec),
    })


# backup: pnputil respond kwargs
DRIVER_BACKUP = {
    "driver_backup_rc0": {"stdout": "Exporting driver package: oem1.inf\nExported 1 driver package.\n"},
    "driver_backup_rc1": {"returncode": 1, "stdout": "Failed to export driver packages.\n"},
    "driver_backup_popen_raises": {"raises": OSError("pnputil missing")},
    "driver_backup_stderr": {"stdout": "Exported 1 driver package.\n",
                             "stderr": "Warning: one driver skipped\nsecond line\n"},
}


@pytest.mark.gui
@pytest.mark.parametrize("case", list(DRIVER_BACKUP))
def test_driver_backup_case(case, gui):
    harness.registry_value(gui.mp, BOARD)
    gui.rec.respond("pnputil", **DRIVER_BACKUP[case])
    result = _run_driver(gui, "start_driver_backup_thread", str(gui.drivers))
    harness.check_golden(GOLDEN_PATH, case, result)


# restore: folder name under <TMP>/drivers, askyesno answers, pnputil respond kwargs
DRIVER_RESTORE = {
    "driver_restore_not_backup_folder": {"folder": "random_folder", "yes": [], "pnputil": {}},
    "driver_restore_mismatch_yes": {"folder": f"OTHERBOARD_drivers_backup_{STAMP}", "yes": [True],
                                    "pnputil": {"stdout": "Added 1 driver package.\n"}},
    "driver_restore_mismatch_no": {"folder": f"OTHERBOARD_drivers_backup_{STAMP}", "yes": [False],
                                   "pnputil": {}},
    "driver_restore_uppercase_marker": {"folder": f"OTHERBOARD_DRIVERS_BACKUP_{STAMP}", "yes": [True],
                                        "pnputil": {"stdout": "Added 1 driver package.\n"}},
    "driver_restore_rc_nonzero": {"folder": f"{BOARD}_drivers_backup_{STAMP}", "yes": [],
                                  "pnputil": {"returncode": 259, "stdout": "Failed to add.\n",
                                              "stderr": "Access is denied.\n"}},
    "driver_restore_popen_raises": {"folder": f"{BOARD}_drivers_backup_{STAMP}", "yes": [],
                                    "pnputil": {"raises": OSError("pnputil missing")}},
}


@pytest.mark.gui
@pytest.mark.parametrize("case", list(DRIVER_RESTORE))
def test_driver_restore_case(case, gui):
    spec = DRIVER_RESTORE[case]
    harness.registry_value(gui.mp, BOARD)
    gui.rec.respond("pnputil", **spec["pnputil"])
    folder = gui.drivers / spec["folder"]
    folder.mkdir()
    (folder / "oem1.inf").write_text("[Version]\n", encoding="utf-8")
    result = _run_driver(gui, "start_driver_restore_thread", str(folder), yes=spec["yes"])
    harness.check_golden(GOLDEN_PATH, case, result)


@pytest.mark.gui
def test_devmgr_all_fail(gui):
    windir = gui.tmp / "Windows"
    (windir / "system32").mkdir(parents=True)  # no devmgmt.msc
    gui.mp.setenv("WINDIR", str(windir))
    gui.rec.respond("mmc", raises=OSError("mmc missing"))
    gui.rec.respond("start devmgmt", raises=OSError("start failed"))
    harness.patch_dialogs(gui.mod, gui.mp, gui.dialogs)
    logged = []
    original_log = gui.app.log

    def log(message):
        logged.append(message)
        original_log(message)

    gui.mp.setattr(gui.app, "log", log)
    gui.app.open_device_manager()
    harness.check_golden(GOLDEN_PATH, "devmgr_all_fail", gui.norm({
        "log": logged, "dialogs": gui.dialogs, "queue": harness.drain(gui.app),
        "commands": _commands(gui.rec)}))
