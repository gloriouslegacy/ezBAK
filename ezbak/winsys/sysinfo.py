"""System information: free disk space (P2), system name and used drive letters (P4).

free_space is moved from ezBAK.App.get_free_space_reliable without behaviour change:
1. shutil.disk_usage(path).free
2. ctypes GetDiskFreeSpaceExW(path) -> free bytes; a zero return value
   (API failure) moves on to step 3 instead of reporting 0 bytes (fix-b)
3. shutil.disk_usage(<drive>\\).free when path has a drive part
4. 999 TB fallback so the caller keeps going

system_name and used_drive_letters are moved from ezBAK.get_system_info and
App.get_used_drive_letters without behaviour change (P4-1, decision D2=A).
"""
import ctypes
import locale
import logging
import os
import re
import shutil

from . import proc

FALLBACK_FREE_BYTES = 999 * 1024**4  # 999 TB

_log = logging.getLogger(__name__)


def _debug(text):
    # Default reporter of system_name: the diagnostic log (P6-5d; was print to stdout)
    _log.debug(text.removeprefix("DEBUG: "))


def free_space(path):
    """Free bytes available at path, never raises (999 TB when every method fails)."""
    try:
        return shutil.disk_usage(path).free
    except Exception:
        free = _api_free_space(path)
        if free is not None:
            return free
        try:
            drive = os.path.splitdrive(path)[0]
            if drive:
                return shutil.disk_usage(drive + os.sep).free
        except Exception:
            # recovery: fall through to the "large enough" value below
            pass
    return FALLBACK_FREE_BYTES


def _api_free_space(path):
    """GetDiskFreeSpaceExW free bytes, or None when the call raises or returns 0 (failure)."""
    try:
        free_bytes = ctypes.c_ulonglong(0)
        ok = ctypes.windll.kernel32.GetDiskFreeSpaceExW(
            ctypes.c_wchar_p(path),
            ctypes.pointer(free_bytes),
            None,
            None
        )
    except Exception:
        # recovery: the caller moves on to the drive root step
        return None
    return free_bytes.value if ok else None


_PLACEHOLDER_NAMES = ('to be filled by o.e.m.', 'system product name', 'system version')

# (wmic argv, value prefix, DEBUG label for a hit, DEBUG label for a failure), in priority order
_WMIC_QUERIES = (
    (['wmic', 'baseboard', 'get', 'product', '/value'], 'Product=',
     "Using baseboard product", "Baseboard product query failed"),
    (['wmic', 'computersystem', 'get', 'model', '/value'], 'Model=',
     "Using computer system model", "Computer system model query failed"),
    (['wmic', 'computersystem', 'get', 'manufacturer', '/value'], 'Manufacturer=',
     "Using manufacturer", "Manufacturer query failed"),
)


def clean_name(name):
    """Name usable in a folder name, or None for empty / placeholder values (C3 folder names)."""
    if not name or name.lower() in _PLACEHOLDER_NAMES:
        return None
    cleaned = re.sub(r'[<>:"/\\|?*]', '', name)
    cleaned = re.sub(r'\s+', '_', cleaned.strip())
    return cleaned if cleaned else None


def system_name(*, debug=_debug):
    """Baseboard / model / manufacturer name, "Unknown" when every source fails.

    Order: registry BaseBoardProduct -> registry SystemProductName -> registry
    SystemManufacturer -> wmic baseboard product -> wmic computersystem model -> wmic
    computersystem manufacturer -> "Unknown". Every step reports through
    debug with the original "DEBUG: ..." text (default: the diagnostic log at DEBUG,
    without the prefix); failures never raise.
    """
    name = "Unknown"
    # Use system default encoding (e.g., cp949 on Korean Windows), read at call time
    system_encoding = locale.getpreferredencoding()

    found = _registry_name(debug)
    if found:
        return found

    for cmd, prefix, hit_label, fail_label in _WMIC_QUERIES:
        try:
            result = proc.run(cmd, capture_output=True, text=True, encoding=system_encoding, errors='replace')
            if result.returncode == 0:
                for line in result.stdout.splitlines():
                    if line.startswith(prefix):
                        cleaned = clean_name(line.split('=', 1)[1].strip())
                        if cleaned:
                            debug(f"DEBUG: {hit_label}: {cleaned}")
                            return cleaned
        except Exception as e:
            # recovery: report and move on to the next source
            debug(f"DEBUG: {fail_label}: {e}")

    debug(f"DEBUG: Using fallback name: {name}")
    return name


# Extra BIOS key values tried after BaseBoardProduct, in order (g-s: wmic is gone on 24H2)
_REGISTRY_FALLBACK_VALUES = ("SystemProductName", "SystemManufacturer")


def _registry_name(debug):
    """Cleaned HKLM BIOS BaseBoardProduct, then SystemProductName, then SystemManufacturer,
    or None (winreg looked up at call time for tests). A missing extra value prints nothing."""
    try:
        import winreg
        reg_path = r"HARDWARE\DESCRIPTION\System\BIOS"
        reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
        try:
            try:
                baseboard_product, _ = winreg.QueryValueEx(reg_key, "BaseBoardProduct")
                clean_product = clean_name(baseboard_product)
                if clean_product:
                    debug(f"DEBUG: Using registry BaseBoardProduct: {clean_product}")
                    return clean_product
            except FileNotFoundError:
                debug("DEBUG: BaseBoardProduct not found in registry")
            for value_name in _REGISTRY_FALLBACK_VALUES:
                try:
                    value, _ = winreg.QueryValueEx(reg_key, value_name)
                except FileNotFoundError:
                    # recovery: a missing value moves on silently (g-s)
                    continue
                cleaned = clean_name(value)
                if cleaned:
                    debug(f"DEBUG: Using registry {value_name}: {cleaned}")
                    return cleaned
        finally:
            winreg.CloseKey(reg_key)
    except Exception as e:
        # recovery: fall back to the wmic queries
        debug(f"DEBUG: Registry query failed: {e}")
    return None


def used_drive_letters(get_logical_drives=None):
    """Used drive letters ('A'..'Z', no colon).

    GetLogicalDrives bitmask first; when it raises, probe <letter>:\\ with os.path.exists.
    get_logical_drives replaces ctypes.windll.kernel32.GetLogicalDrives (both resolved at
    call time, inside the try, so a failed lookup takes the probe path too).
    """
    used = set()
    try:
        if get_logical_drives is None:
            bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        else:
            bitmask = get_logical_drives()
        for i in range(26):
            if bitmask & (1 << i):
                used.add(chr(65 + i))  # 'A'..'Z'
    except Exception:
        # recovery: fallback probe
        for i in range(65, 91):
            letter = chr(i)
            try:
                if os.path.exists(f"{letter}:\\"):
                    used.add(letter)
            except Exception:
                # recovery: skip a letter that cannot be probed
                continue
    return used
