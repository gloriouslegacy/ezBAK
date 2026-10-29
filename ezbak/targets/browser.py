"""Browser bookmark backup (V1 target).

BROWSERS lists each browser's bookmark file under a profile's AppData/Local and
the file name it gets in the destination root (Chrome_Bookmarks, Edge_Bookmarks:
C3). A new browser is one more row in BROWSERS. restore_bookmarks() copies the
files back under the restore target user's profile (fix-d).

Backup reads from and restore writes to the user's profile (<users_root>/<user>)
through the same Browser.source(), never the running account's %LOCALAPPDATA%
(P2-fix #1, fix-d).

backup_profiles() is the full-profile backup (Chrome / Edge User Data, Firefox
Profiles + profiles.ini). Its sources are a separate table, PROFILE_SOURCES, so
BROWSERS (bookmark targets and bm_* wording) stays as it is. The roots are the
caller's %LOCALAPPDATA% / %APPDATA% values (the running account, as before).
"""
import logging
import os
import shutil
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, List, Optional, Tuple

from ezbak.core import InsufficientSpaceError, TargetContext, TargetResult, copy_file_with_progress
from ezbak.winsys import fileattrs

_log = logging.getLogger(__name__)


@dataclass(frozen=True)
class Browser:
    name: str
    # Bookmark file path relative to <profile>/AppData/Local
    relative: Tuple[str, ...]
    # File name in the destination root (C3)
    backup_name: str

    def source(self, profile: str) -> str:
        """Bookmark file of this browser inside the user profile folder `profile`."""
        return os.path.join(profile, "AppData", "Local", *self.relative)


BROWSERS: Tuple[Browser, ...] = (
    Browser("Chrome", ("Google", "Chrome", "User Data", "Default", "Bookmarks"), "Chrome_Bookmarks"),
    Browser("Edge", ("Microsoft", "Edge", "User Data", "Default", "Bookmarks"), "Edge_Bookmarks"),
)


def _emit(out, item, **values) -> None:
    if out is not None and item is not None:
        out(item.line(**values))


class Bookmarks:
    """Copies each browser's bookmark file into ctx.dest (the destination root).

    ctx.policy.isolate_bookmark_errors: True gives each browser its own try
    (GUI); False wraps the whole step in one try (Headless). Wording, line
    breaks and which lines exist come from ctx.fmt (bm_* items).
    """

    name = "bookmarks"

    def __init__(self, browsers: Tuple[Browser, ...] = BROWSERS) -> None:
        self.browsers = tuple(browsers)

    def backup(self, ctx: TargetContext) -> TargetResult:
        fmt, out = ctx.fmt, ctx.out
        result = TargetResult(self.name)
        try:
            _emit(out, fmt.bm_header)
            sources = [(browser, browser.source(ctx.profile)) for browser in self.browsers]
            last = len(sources) - 1
            for index, (browser, src) in enumerate(sources):
                if ctx.policy.isolate_bookmark_errors:
                    try:
                        self._copy_one(ctx, browser, src, index == last)
                    except Exception as e:
                        _log.warning("bookmark_copy_failed browser=%s error=%s", browser.name, e)
                        _emit(out, fmt.bm_browser_error, browser=browser.name, error=e)
                        result.errors.append(f"{browser.name}: {e}")
                else:
                    self._copy_one(ctx, browser, src, index == last)
        except Exception as e:
            _log.warning("bookmark_backup_failed dest=%s error=%s", ctx.dest, e)
            _emit(out, fmt.bm_error, error=e)
            result.errors.append(str(e))
        result.ok = not result.errors
        return result

    @staticmethod
    def _copy_one(ctx: TargetContext, browser: Browser, src: str, is_last: bool) -> None:
        fmt = ctx.fmt
        if os.path.exists(src):
            shutil.copy(src, os.path.join(ctx.dest, browser.backup_name))
            _emit(ctx.out, fmt.bm_saved_last if is_last else fmt.bm_saved, browser=browser.name)
        else:
            _emit(ctx.out, fmt.bm_missing_last if is_last else fmt.bm_missing, browser=browser.name)


def restore_bookmarks(backup_root: str, out, profile: str,
                      browsers: Tuple[Browser, ...] = BROWSERS) -> None:
    """Copies <backup_root>/<backup_name> back to browser.source(profile) per browser.

    `profile` is the restore target user's profile folder (<users_root>/<user>),
    the same place the backup reads from. `out(text)` receives each log line.
    Each browser has its own try; a missing backup file is reported, not raised,
    and a missing target folder is reported as an error line, not created.
    """
    out("Restoring Chrome and Edge browser bookmarks...")
    for browser in browsers:
        target = browser.source(profile)
        try:
            src = os.path.join(backup_root, browser.backup_name)
            if os.path.exists(src):
                shutil.copy(src, target)
                out(f"- {browser.name} bookmarks restored.")
            else:
                out(f"- Backed-up {browser.name} bookmark file not found.")
        except Exception as e:
            _log.warning("bookmark_restore_failed browser=%s error=%s", browser.name, e)
            out(f"- {browser.name} bookmark restore error: {e}")


# ---------------------------------------------------------------- full profiles
LOCAL = "local"      # root: the caller's %LOCALAPPDATA% value
ROAMING = "roaming"  # root: the caller's %APPDATA% value


@dataclass(frozen=True)
class ProfileSource:
    name: str
    root: str                 # LOCAL or ROAMING
    relative: Tuple[str, ...]  # folder under the root
    dest: Tuple[str, ...]      # folder under BrowserProfiles_backup_<ts>

    def path(self, local: str, roaming: str) -> str:
        return os.path.join(local if self.root == LOCAL else roaming, *self.relative)


PROFILE_SOURCES: Tuple[ProfileSource, ...] = (
    ProfileSource("Chrome", LOCAL, ("Google", "Chrome", "User Data"), ("Chrome", "User Data")),
    ProfileSource("Edge", LOCAL, ("Microsoft", "Edge", "User Data"), ("Edge", "User Data")),
    ProfileSource("Firefox", ROAMING, ("Mozilla", "Firefox", "Profiles"), ("Firefox", "Profiles")),
)
# Firefox profiles.ini: under ROAMING, copied to <dest_base>/Firefox/profiles.ini after the folders
FIREFOX_PROFILES_INI = ("Mozilla", "Firefox", "profiles.ini")
FIREFOX_PROFILES_INI_DEST = ("Firefox", "profiles.ini")


def _is_reparse(path: str) -> bool:
    """True for a reparse point (junction, symlink); any lookup failure counts as False."""
    try:
        return bool(fileattrs.get_attributes(path) & fileattrs.FILE_ATTRIBUTE_REPARSE_POINT)
    except Exception:
        return False


def _walk(src: str):
    """os.walk(src) that never descends into reparse-point folders; errors are ignored."""
    for dirpath, dirnames, filenames in os.walk(src, topdown=True, onerror=lambda e: None):
        dirnames[:] = [d for d in dirnames if not _is_reparse(os.path.join(dirpath, d))]
        yield dirpath, filenames


def _total_size(sources: List[Tuple[str, str]], ini: str) -> int:
    total = 0
    for src, _rel in sources:
        for dirpath, filenames in _walk(src):
            for f in filenames:
                p = os.path.join(dirpath, f)
                try:
                    if not _is_reparse(p):
                        total += os.path.getsize(p)
                except Exception as e:
                    _log.debug("profile_size_skip path=%s error=%s", p, type(e).__name__)
    if os.path.exists(ini):
        try:
            total += os.path.getsize(ini)
        except Exception as e:
            _log.debug("profile_size_skip path=%s error=%s", ini, type(e).__name__)
    return total


def _copy_folder(src: str, dst_dir: str, out: Callable[[str], None], copy) -> None:
    out(f"Copying directory: {src} -> {dst_dir}")
    if os.path.exists(dst_dir):
        try:
            shutil.rmtree(dst_dir)
            out(f"  - Removed existing destination {dst_dir}")
        except Exception as e:
            out(f"  - Failed to remove existing destination: {e}")
    os.makedirs(dst_dir, exist_ok=True)
    for dirpath, filenames in _walk(src):
        relp = os.path.relpath(dirpath, src)
        target_dir = os.path.join(dst_dir, relp) if relp != '.' else dst_dir
        os.makedirs(target_dir, exist_ok=True)
        for f in filenames:
            sfile = os.path.join(dirpath, f)
            if _is_reparse(sfile):
                continue
            dfile = os.path.join(target_dir, f)
            out(f"Copying file: {sfile} -> {dfile}")
            copy(sfile, dfile)


def backup_profiles(dest_root: str, out: Callable[[str], None], progress, *, local: str,
                    roaming: str, free_space: Callable[[str], int],
                    clock: Callable[[], datetime] = datetime.now) -> Optional[str]:
    """Copy every existing PROFILE_SOURCES folder (and Firefox profiles.ini) into a new
    <dest_root>/BrowserProfiles_backup_<ts> folder; returns that folder.

    `out(text)` gets the work-log lines; `progress` is a core Reporter (progress_max,
    progress with the running byte count, the final count reported again at the end).
    Reparse points are skipped. Missing sources are skipped silently. When no source
    exists at all (profiles.ini included), returns None after the total line and creates
    no folder. When the total is larger than free_space(dest_root), InsufficientSpaceError
    is raised after the log line and before any folder is created. Other errors propagate.
    """
    sources = [(s.path(local, roaming), os.path.join(*s.dest)) for s in PROFILE_SOURCES]
    sources = [(src, rel) for src, rel in sources if os.path.exists(src)]
    ff_ini = os.path.join(roaming, *FIREFOX_PROFILES_INI)
    dest_base = os.path.join(dest_root, f"BrowserProfiles_backup_{clock().strftime('%Y-%m-%d_%H%M%S')}")
    _log.info("event=profiles_backup_start dest=%s sources=%d", dest_base, len(sources))

    total = _total_size(sources, ff_ini)
    progress.progress_max(total)
    out(f"Browser profiles total bytes: {total}")
    if not sources and not os.path.exists(ff_ini):
        _log.info("event=profiles_backup_no_source dest_root=%s", dest_root)
        return None
    free = free_space(dest_root)
    if total > free:
        out(f"Insufficient free space. Required={total} Free={free}")
        raise InsufficientSpaceError(total, free)
    os.makedirs(dest_base, exist_ok=True)

    copied = 0

    def copy(src: str, dst: str) -> None:
        def on_chunk(delta: int) -> None:
            nonlocal copied
            copied += delta
            progress.progress(copied)
        copy_file_with_progress(src, dst, on_chunk,
                                on_error=lambda e: out(f"Error copying file {src} -> {dst}: {e}"))

    for src, rel in sources:
        _copy_folder(src, os.path.join(dest_base, rel), out, copy)
    if os.path.exists(ff_ini):
        dst_ini = os.path.join(dest_base, *FIREFOX_PROFILES_INI_DEST)
        os.makedirs(os.path.dirname(dst_ini), exist_ok=True)
        try:
            out(f"Copying file: {ff_ini} -> {dst_ini}")
            copy(ff_ini, dst_ini)
        except Exception as e:
            out(f"Failed to copy profiles.ini: {e}")

    progress.progress(copied)
    out("Browser profiles backup completed.")
    _log.info("event=profiles_backup_done dest=%s bytes=%d", dest_base, copied)
    return dest_base
