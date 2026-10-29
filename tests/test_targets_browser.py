"""ezbak.targets browser bookmarks (P2-5, source from the target profile: P2-fix #1) and full-profile backup (P4-6)."""
import datetime as _dt
import importlib.util
import os
import shutil
import sys
from pathlib import Path

import pytest

from ezbak.core import BackupPolicy, BackupTarget, InsufficientSpaceError, JobLogFormat, TargetContext
from ezbak.targets import (BROWSERS, PROFILE_SOURCES, Bookmarks, backup_profiles, default_targets,
                           restore_bookmarks)

GUI = (JobLogFormat.GUI, BackupPolicy.GUI)
HEADLESS = (JobLogFormat.HEADLESS, BackupPolicy.HEADLESS)


def render(lines):
    return "".join(f"{ln.lead}[TS] {ln.message}{ln.end}" for ln in lines)


def _profile(tmp_path):
    return tmp_path / "Users" / "alice"


def _write_bookmark(profile, browser, text):
    path = os.path.join(str(profile), "AppData", "Local", *browser.relative)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def _run(dest, path, profile=None):
    fmt, policy = path
    profile = dest.parent / "Users" / "alice" if profile is None else profile
    lines = []
    result = Bookmarks().backup(TargetContext(backup_path=str(dest), dest=str(dest), fmt=fmt,
                                              policy=policy, out=lines.append,
                                              profile=str(profile)))
    return result, render(lines)


def test_browsers_table():
    assert [(b.name, b.backup_name) for b in BROWSERS] == [
        ("Chrome", "Chrome_Bookmarks"), ("Edge", "Edge_Bookmarks")]
    assert BROWSERS[0].relative == ("Google", "Chrome", "User Data", "Default", "Bookmarks")
    assert BROWSERS[1].relative == ("Microsoft", "Edge", "User Data", "Default", "Bookmarks")


def test_source_is_under_profile_local(isolated_env, tmp_path):
    profile = str(_profile(tmp_path))
    assert BROWSERS[0].source(profile) == os.path.join(profile, "AppData", "Local",
                                                        *BROWSERS[0].relative)
    assert not BROWSERS[0].source(profile).startswith(str(isolated_env["LOCALAPPDATA"]))


@pytest.mark.parametrize("path", [GUI, HEADLESS])
def test_reads_target_profile_not_running_account(isolated_env, tmp_path, path):
    dest = tmp_path / "dest"
    dest.mkdir()
    profile = _profile(tmp_path)
    _write_bookmark(profile, BROWSERS[0], "target-chrome")
    # Decoys in the running account's %LOCALAPPDATA%: must not be read.
    for browser in BROWSERS:
        path_ = os.path.join(str(isolated_env["LOCALAPPDATA"]), *browser.relative)
        os.makedirs(os.path.dirname(path_), exist_ok=True)
        with open(path_, "w", encoding="utf-8") as f:
            f.write("decoy")
    result, text = _run(dest, path, profile)
    assert result.ok
    assert "Chrome bookmarks backed up." in text and "Edge bookmark file not found." in text
    assert (dest / "Chrome_Bookmarks").read_text(encoding="utf-8") == "target-chrome"
    assert not (dest / "Edge_Bookmarks").exists()


def test_protocol_and_default_targets():
    assert isinstance(Bookmarks(), BackupTarget)
    targets = default_targets()
    assert len(targets) == 1 and isinstance(targets[0], Bookmarks)


HEADER = "[TS] Backing up Chrome and Edge browser bookmarks...\n"


@pytest.mark.parametrize("path, tail", [(GUI, "\n"), (HEADLESS, "\n")])
def test_no_bookmarks(isolated_env, tmp_path, path, tail):
    dest = tmp_path / "dest"
    dest.mkdir()
    result, text = _run(dest, path)
    assert result.ok and result.errors == []
    assert text == (HEADER + "[TS] - Chrome bookmark file not found.\n"
                    "[TS] - Edge bookmark file not found." + tail)
    assert os.listdir(dest) == []


@pytest.mark.parametrize("path, tail", [(GUI, "\n"), (HEADLESS, "\n")])
def test_one_bookmark(isolated_env, tmp_path, path, tail):
    dest = tmp_path / "dest"
    dest.mkdir()
    _write_bookmark(_profile(tmp_path), BROWSERS[0], '{"chrome": 1}')
    result, text = _run(dest, path)
    assert result.ok
    assert text == (HEADER + "[TS] - Chrome bookmarks backed up.\n"
                    "[TS] - Edge bookmark file not found." + tail)
    assert (dest / "Chrome_Bookmarks").read_text(encoding="utf-8") == '{"chrome": 1}'
    assert not (dest / "Edge_Bookmarks").exists()


@pytest.mark.parametrize("path, tail", [(GUI, "\n"), (HEADLESS, "\n")])
def test_both_bookmarks(isolated_env, tmp_path, path, tail):
    dest = tmp_path / "dest"
    dest.mkdir()
    _write_bookmark(_profile(tmp_path), BROWSERS[0], "c")
    _write_bookmark(_profile(tmp_path), BROWSERS[1], "e")
    result, text = _run(dest, path)
    assert result.ok
    assert text == (HEADER + "[TS] - Chrome bookmarks backed up.\n"
                    "[TS] - Edge bookmarks backed up." + tail)
    assert sorted(os.listdir(dest)) == ["Chrome_Bookmarks", "Edge_Bookmarks"]


def _fail_chrome_copy(monkeypatch):
    real_copy = shutil.copy

    def copy(src, dst, *args, **kwargs):
        if os.path.basename(dst) == "Chrome_Bookmarks":
            raise OSError("boom")
        return real_copy(src, dst, *args, **kwargs)

    monkeypatch.setattr(shutil, "copy", copy)


def test_copy_error_gui_isolates_each_browser(isolated_env, tmp_path, monkeypatch):
    dest = tmp_path / "dest"
    dest.mkdir()
    _write_bookmark(_profile(tmp_path), BROWSERS[0], "c")
    _write_bookmark(_profile(tmp_path), BROWSERS[1], "e")
    _fail_chrome_copy(monkeypatch)
    result, text = _run(dest, GUI)
    assert not result.ok and result.errors == ["Chrome: boom"]
    assert text == (HEADER + "[TS] - Chrome bookmark backup error: boom\n"
                    "[TS] - Edge bookmarks backed up.\n")
    assert os.listdir(dest) == ["Edge_Bookmarks"]


def test_copy_error_headless_stops_at_first(isolated_env, tmp_path, monkeypatch):
    dest = tmp_path / "dest"
    dest.mkdir()
    _write_bookmark(_profile(tmp_path), BROWSERS[0], "c")
    _write_bookmark(_profile(tmp_path), BROWSERS[1], "e")
    _fail_chrome_copy(monkeypatch)
    result, text = _run(dest, HEADLESS)
    assert not result.ok and result.errors == ["boom"]
    assert text == HEADER + "[TS] Browser bookmark backup error: boom\n"
    assert os.listdir(dest) == []


def test_gui_outer_error(isolated_env, tmp_path, monkeypatch):
    def broken(*args, **kwargs):
        raise ValueError("bad path")

    monkeypatch.setattr(os.path, "join", broken)
    lines = []
    result = Bookmarks().backup(TargetContext(dest="d", fmt=JobLogFormat.GUI, policy=BackupPolicy.GUI,
                                              out=lines.append))
    monkeypatch.undo()
    assert not result.ok
    assert render(lines) == HEADER + "[TS] Browser bookmark backup failed: bad path\n"


# --- restore_bookmarks (P3-4, fix-d): back to the target user's profile ---

RESTORE_HEADER = "Restoring Chrome and Edge browser bookmarks..."


def _restore_setup(tmp_path, present):
    backup = tmp_path / "backup"
    profile = tmp_path / "Users" / "tester"
    backup.mkdir()
    for browser in BROWSERS:
        os.makedirs(os.path.dirname(browser.source(str(profile))), exist_ok=True)
        if browser.name in present:
            (backup / browser.backup_name).write_text(f"{browser.name}-data", encoding="utf-8")
    return backup, profile


def _restore(backup, profile):
    lines = []
    restore_bookmarks(str(backup), lines.append, str(profile))
    return lines


def test_restore_both_present_copies_under_profile(tmp_path):
    backup, profile = _restore_setup(tmp_path, {"Chrome", "Edge"})
    assert _restore(backup, profile) == [
        RESTORE_HEADER, "- Chrome bookmarks restored.", "- Edge bookmarks restored."]
    for browser in BROWSERS:
        target = os.path.join(str(profile), "AppData", "Local", *browser.relative)
        with open(target, encoding="utf-8") as f:
            assert f.read() == f"{browser.name}-data"


def test_restore_only_one_present(tmp_path):
    backup, profile = _restore_setup(tmp_path, {"Edge"})
    assert _restore(backup, profile) == [
        RESTORE_HEADER, "- Backed-up Chrome bookmark file not found.", "- Edge bookmarks restored."]
    assert not os.path.exists(BROWSERS[0].source(str(profile)))


def test_restore_none_present(tmp_path):
    backup, profile = _restore_setup(tmp_path, set())
    assert _restore(backup, profile) == [
        RESTORE_HEADER, "- Backed-up Chrome bookmark file not found.",
        "- Backed-up Edge bookmark file not found."]


def test_restore_copy_error_is_per_browser(tmp_path, monkeypatch):
    backup, profile = _restore_setup(tmp_path, {"Chrome", "Edge"})
    real_copy = shutil.copy
    targets = []

    def fake_copy(src, dst):
        targets.append(dst)
        if os.path.basename(src) == "Chrome_Bookmarks":
            raise OSError("denied")
        return real_copy(src, dst)

    monkeypatch.setattr(shutil, "copy", fake_copy)
    assert _restore(backup, profile) == [
        RESTORE_HEADER, "- Chrome bookmark restore error: denied", "- Edge bookmarks restored."]
    assert targets == [b.source(str(profile)) for b in BROWSERS]


def test_restore_does_not_write_running_account_local_appdata(tmp_path, isolated_env):
    backup, profile = _restore_setup(tmp_path, {"Chrome", "Edge"})
    local = str(isolated_env["LOCALAPPDATA"])
    for browser in BROWSERS:
        os.makedirs(os.path.join(local, *browser.relative[:-1]), exist_ok=True)
    _restore(backup, profile)
    for browser in BROWSERS:
        assert not os.path.exists(os.path.join(local, *browser.relative))
        assert os.path.exists(browser.source(str(profile)))

# ---------------------------------------------------------------- full profiles (P4-6)
_USER_TREE = "ezbak_test_user_tree"
if _USER_TREE not in sys.modules:
    _spec = importlib.util.spec_from_file_location(
        _USER_TREE, Path(__file__).resolve().parent / "fixtures" / "user_tree.py")
    sys.modules[_USER_TREE] = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(sys.modules[_USER_TREE])
user_tree = sys.modules[_USER_TREE]

_NOW = _dt.datetime(2024, 1, 2, 3, 4, 5)
_BIG = 10**12


class _Progress:
    def __init__(self):
        self.calls = []

    def progress_max(self, value):
        self.calls.append(("max", value))

    def progress(self, value):
        self.calls.append(("progress", value))


def _put(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _roots(tmp_path):
    local = tmp_path / "env" / "Local"
    roaming = tmp_path / "env" / "Roaming"
    local.mkdir(parents=True)
    roaming.mkdir(parents=True)
    return local, roaming


def _backup(tmp_path, local, roaming, free=_BIG):
    dest = tmp_path / "dest"
    dest.mkdir(exist_ok=True)
    lines, progress, asked = [], _Progress(), []

    def free_space(path):
        asked.append(path)
        return free
    base = backup_profiles(str(dest), lines.append, progress, local=str(local), roaming=str(roaming),
                           free_space=free_space, clock=lambda: _NOW)
    return base, lines, progress, asked


def _tree(root):
    found = []
    for dirpath, _dirs, files in os.walk(root):
        for f in files:
            found.append(os.path.relpath(os.path.join(dirpath, f), root).replace(os.sep, "/"))
    return sorted(found)


def test_profile_sources_table():
    assert [(s.name, s.root, s.relative, s.dest) for s in PROFILE_SOURCES] == [
        ("Chrome", "local", ("Google", "Chrome", "User Data"), ("Chrome", "User Data")),
        ("Edge", "local", ("Microsoft", "Edge", "User Data"), ("Edge", "User Data")),
        ("Firefox", "roaming", ("Mozilla", "Firefox", "Profiles"), ("Firefox", "Profiles")),
    ]
    assert PROFILE_SOURCES[0].path("L", "R") == os.path.join("L", "Google", "Chrome", "User Data")
    assert PROFILE_SOURCES[2].path("L", "R") == os.path.join("R", "Mozilla", "Firefox", "Profiles")
    # The bookmark table is a different table and is unchanged.
    assert [b.name for b in BROWSERS] == ["Chrome", "Edge"]


def test_profiles_copy_tree(tmp_path):
    local, roaming = _roots(tmp_path)
    _put(local / "Google" / "Chrome" / "User Data" / "Default" / "Login Data", "chrome")
    _put(local / "Microsoft" / "Edge" / "User Data" / "Local State", "edge")
    _put(roaming / "Mozilla" / "Firefox" / "Profiles" / "a.default" / "prefs.js", "ff")
    _put(roaming / "Mozilla" / "Firefox" / "profiles.ini", "[P]")
    base, lines, progress, asked = _backup(tmp_path, local, roaming)
    assert os.path.basename(base) == "BrowserProfiles_backup_2024-01-02_030405"
    assert asked == [str(tmp_path / "dest")]  # checked on dest_root, before the folder exists
    assert _tree(base) == ["Chrome/User Data/Default/Login Data", "Edge/User Data/Local State",
                           "Firefox/Profiles/a.default/prefs.js", "Firefox/profiles.ini"]
    total = len("chrome") + len("edge") + len("ff") + len("[P]")
    assert progress.calls[0] == ("max", total)
    assert progress.calls[-1] == ("progress", total)
    assert lines[0] == f"Browser profiles total bytes: {total}"
    assert lines[-1] == "Browser profiles backup completed."
    assert sum(ln.startswith("Copying directory: ") for ln in lines) == 3


def test_profiles_no_source(tmp_path):
    local, roaming = _roots(tmp_path)
    base, lines, progress, asked = _backup(tmp_path, local, roaming)
    # g-o: nothing to back up -> None, no folder, the total line is still written
    assert base is None
    assert os.listdir(tmp_path / "dest") == []
    assert lines == ["Browser profiles total bytes: 0"]
    assert progress.calls == [("max", 0)]
    assert asked == []


def test_profiles_only_profiles_ini_still_backs_up(tmp_path):
    local, roaming = _roots(tmp_path)
    _put(roaming / "Mozilla" / "Firefox" / "profiles.ini", "[P]")
    base, lines, _progress, _asked = _backup(tmp_path, local, roaming)
    assert _tree(base) == ["Firefox/profiles.ini"]
    assert lines[-1] == "Browser profiles backup completed."


def test_profiles_insufficient_space(tmp_path):
    local, roaming = _roots(tmp_path)
    _put(local / "Google" / "Chrome" / "User Data" / "Local State", "0123456789")
    dest = tmp_path / "dest"
    dest.mkdir()
    lines, asked = [], []

    def free_space(path):
        asked.append(path)
        return 3
    with pytest.raises(InsufficientSpaceError) as info:
        backup_profiles(str(dest), lines.append, _Progress(), local=str(local),
                        roaming=str(roaming), free_space=free_space, clock=lambda: _NOW)
    assert (info.value.required, info.value.free) == (10, 3)
    assert lines[-1] == "Insufficient free space. Required=10 Free=3"
    # g-n: space is checked on dest_root before the timestamped folder is made; none is left
    assert asked == [str(dest)]
    assert os.listdir(dest) == []


@pytest.mark.windows
def test_profiles_skip_junction(tmp_path):
    local, roaming = _roots(tmp_path)
    data = local / "Google" / "Chrome" / "User Data"
    _put(data / "Local State", "x")
    _put(tmp_path / "outside" / "secret.txt", "must not be copied")
    user_tree.make_junction(str(data / "Linked"), str(tmp_path / "outside"))
    base, lines, progress, _asked = _backup(tmp_path, local, roaming)
    assert _tree(base) == ["Chrome/User Data/Local State"]
    assert progress.calls[0] == ("max", 1)
    assert not any("secret.txt" in ln for ln in lines)
