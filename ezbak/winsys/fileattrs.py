"""Raw Windows file attribute lookup (pywin32 -> ctypes fallback).

Moved from ezBAK.py module level (the `import win32api` block and the ctypes
`_get_file_attributes` wrapper). No decisions here (hidden / system / reparse
policy belongs to core.filters).

Failure behaviour is the same as ezBAK.py:
- pywin32 path: whatever win32api.GetFileAttributes raises (pywintypes.error).
- ctypes path: OSError on INVALID_FILE_ATTRIBUTES, OSError("GetFileAttributesW
  unavailable") when the ctypes wrapper could not be built.

HAS_PYWIN32 is read at call time, so tests can monkeypatch it to False to take
the ctypes path, or re-import this module with win32api hidden in sys.modules.
"""
import ctypes

try:
    import win32api
    import win32con  # noqa: F401  (kept to match the ezBAK.py availability check)
    HAS_PYWIN32 = True
except Exception:
    win32api = None
    win32con = None
    HAS_PYWIN32 = False

# Windows file attribute constants (from WinBase.h); same values as win32con.
FILE_ATTRIBUTE_READONLY = 0x01
FILE_ATTRIBUTE_HIDDEN = 0x02
FILE_ATTRIBUTE_SYSTEM = 0x04
FILE_ATTRIBUTE_DIRECTORY = 0x10
FILE_ATTRIBUTE_ARCHIVE = 0x20
FILE_ATTRIBUTE_DEVICE = 0x40
FILE_ATTRIBUTE_NORMAL = 0x80
FILE_ATTRIBUTE_REPARSE_POINT = 0x400

INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF

# ctypes wrapper for GetFileAttributesW
try:
    from ctypes import wintypes

    _GetFileAttributesW = ctypes.windll.kernel32.GetFileAttributesW
    _GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
    _GetFileAttributesW.restype = wintypes.DWORD

    def _ctypes_get_attributes(path):
        res = _GetFileAttributesW(path)
        if res == INVALID_FILE_ATTRIBUTES:
            raise OSError(f"GetFileAttributesW failed for {path}")
        return int(res)
except Exception:
    # ctypes fallback could not be built (non-Windows or broken kernel32)
    def _ctypes_get_attributes(path):
        raise OSError("GetFileAttributesW unavailable")


def get_attributes(path):
    """Return the raw attribute bitmask of `path` (int). Raises on failure (see module doc)."""
    if HAS_PYWIN32 and win32api is not None:
        return win32api.GetFileAttributes(path)
    return _ctypes_get_attributes(path)
