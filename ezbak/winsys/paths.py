"""Well-known paths: C:\\Users, %APPDATA%, PyInstaller _MEIPASS.

Environment variables are read at call time (tests monkeypatch them).
Values match ezBAK.py as is; callers keep their own fallbacks
(e.g. settings use `appdata_dir() or os.path.expanduser('~')`).
"""
import os
import sys

# The single definition of C:\Users (P5 decision D3). Callers read it at call time
# (paths.USERS_ROOT or users_root()), so tests can patch this one name.
USERS_ROOT = os.path.join("C:", os.sep, "Users")


def users_root():
    """C:\\Users: the current value of USERS_ROOT."""
    return USERS_ROOT


def appdata_dir():
    """%APPDATA% as os.environ.get('APPDATA', ''): empty string when unset."""
    return os.environ.get('APPDATA', '')


def resource_path(relative_path):
    """Absolute path to a bundled resource: base is sys._MEIPASS, else the current directory."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def list_users(root=None):
    """User profile folder names under `root` (default users_root()), as App.load_users lists them.

    Directories only, skipping names starting with '.' and 'All Users' (any case).
    Any error while listing yields [].
    """
    users_path = users_root() if root is None else root
    try:
        users = [d for d in os.listdir(users_path) if os.path.isdir(os.path.join(users_path, d)) and not d.startswith('.') and d.lower() != "all users"]
    except Exception:
        users = []
    return users
