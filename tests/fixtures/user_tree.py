"""User profile tree used by the backup golden tests (T0-2).

Each entry is (relative path, kind, attributes, content):
- kind "file" writes `content` (UTF-8), "dir" creates the folder,
  "junction" creates a directory junction pointing at the relative path in `content`.
- attributes are FILE_ATTRIBUTE_* bits applied with SetFileAttributesW after creation.

Entries are created in list order, so parents come first and attributes are set
after the folder's children exist.
"""
import ctypes
import os
import subprocess

# Captured at import time: test fixtures (block_subprocess, record_subprocess) patch
# subprocess.Popen (which subprocess.run uses), but junctions must still be created for real.
_REAL_POPEN = subprocess.Popen

USER_NAME = "tester"

H = 0x02  # FILE_ATTRIBUTE_HIDDEN
S = 0x04  # FILE_ATTRIBUTE_SYSTEM

CHROME_BOOKMARKS = os.path.join("AppData", "Local", "Google", "Chrome",
                                "User Data", "Default", "Bookmarks")

TREE = [
    # ordinary data
    ("Documents", "dir", 0, None),
    ("Documents/report.txt", "file", 0, "quarterly report\n"),
    ("Documents/notes.md", "file", 0, "# notes\n- one\n- two\n"),
    ("Documents/sub", "dir", 0, None),
    ("Documents/sub/deep.txt", "file", 0, "deep file\n"),
    ("Desktop", "dir", 0, None),
    ("Desktop/todo.txt", "file", 0, "todo\n"),
    ("Pictures", "dir", 0, None),
    ("Pictures/photo.jpg", "file", 0, "JPEGDATA" * 8),
    ("Empty", "dir", 0, None),
    # default filter targets (name rules: onedrive*, thumbs.db, desktop.ini, $recycle.bin)
    ("Desktop/desktop.ini", "file", 0, "[.ShellClassInfo]\n"),
    ("Pictures/Thumbs.db", "file", 0, "thumbs"),
    ("OneDrive", "dir", 0, None),
    ("OneDrive/cloud.txt", "file", 0, "cloud\n"),
    ("OneDrive - Company", "dir", 0, None),
    ("OneDrive - Company/work.txt", "file", 0, "work\n"),
    ("$Recycle.Bin", "dir", 0, None),
    ("$Recycle.Bin/trash.txt", "file", 0, "trash\n"),
    # hidden / system attributes
    ("hidden.txt", "file", H, "hidden\n"),
    ("system.txt", "file", S, "system\n"),
    ("NTUSER.DAT", "file", H | S, "registry hive\n"),
    ("Documents/hiddendir", "dir", 0, None),
    ("Documents/hiddendir/inner.txt", "file", 0, "inside hidden dir\n"),
    ("Documents/hiddendir", "attr", H, None),
    # dot files
    (".gitconfig", "file", 0, "[user]\n"),
    (".ssh", "dir", 0, None),
    (".ssh/known_hosts", "file", 0, "host key\n"),
    # AppData (hidden on a real profile) and the browser bookmark spot
    ("AppData", "dir", 0, None),
    ("AppData/Roaming", "dir", 0, None),
    ("AppData/Roaming/app.cfg", "file", 0, "setting=1\n"),
    ("AppData/Local", "dir", 0, None),
    ("AppData/Local/Google/Chrome/User Data/Default", "dir", 0, None),
    (CHROME_BOOKMARKS.replace(os.sep, "/"), "file", 0, '{"roots": {}}\n'),
    ("AppData", "attr", H, None),
    # legacy junction found in real profiles (non-cyclic here)
    ("Application Data", "junction", 0, "AppData/Roaming"),
]


def set_attributes(path, attrs):
    set_attrs = ctypes.windll.kernel32.SetFileAttributesW
    set_attrs.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
    set_attrs.restype = ctypes.c_int
    if not set_attrs(path, attrs):
        raise ctypes.WinError()


def make_junction(link, target):
    """Create a directory junction `link` -> `target` in a separate process.

    _winapi.CreateJunction in the test process made later Tk() creations fail now and
    then (tcl library files not found), so the junction is made by `cmd /c mklink /J`.
    """
    with _REAL_POPEN(["cmd", "/c", "mklink", "/J", link, target],
                     stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True) as proc:
        out, err = proc.communicate()
    if proc.returncode != 0 or not os.path.isdir(link):
        raise OSError(f"mklink /J failed rc={proc.returncode} link={link} "
                      f"stdout={out.strip()} stderr={err.strip()}")


def build_user_tree(home):
    """Create TREE under `home` (a str path) and return it."""
    os.makedirs(home, exist_ok=True)
    for rel, kind, attrs, content in TREE:
        path = os.path.join(home, *rel.split("/"))
        if kind == "file":
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(content)
        elif kind == "dir":
            os.makedirs(path, exist_ok=True)
        elif kind == "junction":
            make_junction(path, os.path.join(home, *content.split("/")))
        if attrs:
            set_attributes(path, attrs)
    return home
