"""File operations shared by GUI and Headless: chunked copy, size total, read-only rmtree, multi-source copy.

Moved from ezBAK.py (`App.copy_file_with_progress[_safe]`, `copy_file_safe_headless`,
`App.compute_total_bytes_safe`, `calculate_total_size`, the size loop of
`App.run_space_calculation`, the rmtree error handlers and the copy loop of
`App.run_copy_thread`). Nothing here writes log text or touches a UI: callers
pass callbacks and keep their own wording (Q1). Skip judgment is a callback
`skip(path, is_dir) -> bool` so each caller keeps its own policy (GUI
`is_hidden_safe`, the `App.is_hidden` variant, Headless names + filters).
Errors are swallowed exactly where the old copies swallowed them.
"""
from __future__ import annotations

import os
import shutil
import stat
from datetime import datetime
from functools import partial
from typing import Callable, Iterable, Optional, Tuple

BUFFER_SIZE = 64 * 1024
COPY_TIMEOUT_SECONDS = 30

Skip = Callable[[str, bool], bool]
Log = Callable[[str], None]


def copy_file_with_progress(src: str, dst: str, progress: Optional[Callable[[int], None]] = None, *,
                            buffer_size: int = BUFFER_SIZE, timeout_seconds: Optional[float] = None,
                            clock: Callable[[], datetime] = datetime.now,
                            on_timeout: Optional[Callable[[str], None]] = None,
                            on_error: Optional[Callable[[Exception], None]] = None) -> int:
    """Copy `src` to `dst` in chunks; returns bytes copied, or 0 on any error.

    `progress(delta)` is called after every chunk. With `timeout_seconds` the
    elapsed time (from `clock`) is checked before every read; when exceeded
    `on_timeout(src)` is called and the copy fails with TimeoutError. Any error
    is passed to `on_error(exc)` and swallowed. copystat failures are ignored.
    """
    try:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        copied = 0
        start = clock() if timeout_seconds is not None else None
        with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
            while True:
                if start is not None and (clock() - start).total_seconds() > timeout_seconds:
                    if on_timeout is not None:
                        on_timeout(src)
                    raise TimeoutError(f"File copy timeout: {src}")
                buf = fsrc.read(buffer_size)
                if not buf:
                    break
                fdst.write(buf)
                copied += len(buf)
                if progress is not None:
                    progress(len(buf))
        try:
            shutil.copystat(src, dst)
        except Exception:
            pass  # attribute copy failure is ignored (all old copies did this)
        return copied
    except Exception as e:
        # Work-unit boundary: one failed file does not stop the backup; caller logs it.
        if on_error is not None:
            on_error(e)
        return 0


def _keep_dir(skip: Skip, path: str, strict: bool) -> bool:
    if strict:
        return not skip(path, True)
    try:
        return not skip(path, True)
    except Exception:
        return True  # include the directory when the check fails (old behaviour)


def _file_size(skip: Skip, fp: str, strict: bool, on_file_error: Optional[Callable]) -> Optional[int]:
    """Size of a counted file, or None when skipped or unreadable."""
    if strict:
        if skip(fp, False):
            return None
        try:
            return os.path.getsize(fp)
        except OSError:
            return None
    try:
        if skip(fp, False):
            return None
        return os.path.getsize(fp)
    except Exception as e:
        if on_file_error is not None:
            on_file_error(fp, e)
        return None


def compute_total_size(root: str, skip: Skip, *, strict: bool = False,
                       on_progress: Optional[Callable[[int, int], None]] = None, progress_every: int = 100,
                       on_walk_error: Optional[Callable[[OSError], None]] = None,
                       on_file_error: Optional[Callable[[str, Exception], None]] = None,
                       on_error: Optional[Callable[[Exception], None]] = None) -> Tuple[int, int]:
    """(total bytes, file count) under `root`, leaving out what `skip` rejects.

    Not `strict`: a failing directory check keeps the directory, a failing file
    check or size read goes to `on_file_error`. `strict`: check errors propagate
    and only OSError from the size read is ignored (old space calculation).
    `on_progress(count, total)` runs every `progress_every` files. With
    `on_error`, an unexpected error ends the walk and the partial total is kept.
    """
    total = 0
    count = 0
    try:
        for dirpath, dirnames, filenames in os.walk(root, topdown=True, onerror=on_walk_error):
            dirnames[:] = [d for d in dirnames if _keep_dir(skip, os.path.join(dirpath, d), strict)]
            for f in filenames:
                size = _file_size(skip, os.path.join(dirpath, f), strict, on_file_error)
                if size is None:
                    continue
                total += size
                count += 1
                if on_progress is not None and count % progress_every == 0:
                    on_progress(count, total)
    except Exception as e:
        if on_error is None:
            raise
        on_error(e)
    return total, count


def handle_readonly_error(func, path, exc_info, on_log: Log) -> None:
    """shutil.rmtree onerror body: clear read-only and retry on PermissionError.

    Messages (unchanged wording) go to `on_log`; nothing is raised.
    """
    if isinstance(exc_info[1], PermissionError):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
            on_log(f"INFO: Removed read-only and successfully deleted {path}")
        except Exception as e:
            on_log(f"ERROR: Could not delete {path} even after removing read-only flag: {e}")
    else:
        on_log(f"ERROR: Deletion failed for {path}: {exc_info[1]}")


def rmtree_readonly(path: str, on_log: Optional[Log] = None) -> None:
    """shutil.rmtree that also removes read-only entries (see handle_readonly_error)."""
    shutil.rmtree(path, onerror=partial(handle_readonly_error, on_log=on_log or (lambda _m: None)))


def _copy_tree(src_root: str, dest_base: str, skip: Skip, copy_one, on_copy, on_count,
               count_every: int, done: int) -> int:
    for dirpath, dirnames, filenames in os.walk(src_root, topdown=True, onerror=None):
        dirnames[:] = [d for d in dirnames if not skip(os.path.join(dirpath, d), True)]
        rel_dir = os.path.relpath(dirpath, src_root)
        dest_dir = dest_base if rel_dir == "." else os.path.join(dest_base, rel_dir)
        os.makedirs(dest_dir, exist_ok=True)
        for f in filenames:
            src_file = os.path.join(dirpath, f)
            dest_file = os.path.join(dest_dir, f)
            if skip(src_file, False):
                continue
            if on_copy is not None:
                on_copy(src_file, dest_file)
            copy_one(src_file, dest_file)
            done += 1
            if on_count is not None and done % count_every == 0:
                on_count(done)
    return done


def copy_sources(sources: Iterable[str], dest_root: str, skip: Skip, *,
                 copy_one: Optional[Callable[[str, str], int]] = None,
                 on_copy: Optional[Callable[[str, str], None]] = None,
                 on_count: Optional[Callable[[int], None]] = None, count_every: int = 10) -> int:
    """Copy each source into `dest_root`; returns the number of files processed.

    A directory is copied as `dest_root/<its name>/...`, a file as
    `dest_root/<its name>`. `on_copy(src, dst)` runs before each copy and
    `on_count(n)` every `count_every` files copied from directories (single
    files are not counted toward it, as before). Errors propagate.
    """
    copy_one = copy_one or copy_file_with_progress
    done = 0
    for path in sources:
        if os.path.isdir(path):
            dest_base = os.path.join(dest_root, os.path.basename(path))
            done = _copy_tree(path, dest_base, skip, copy_one, on_copy, on_count, count_every, done)
        elif os.path.isfile(path) and not skip(path, False):
            dest_file = os.path.join(dest_root, os.path.basename(path))
            if on_copy is not None:
                on_copy(path, dest_file)
            copy_one(path, dest_file)
            done += 1
    return done
