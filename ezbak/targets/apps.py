"""Installed-apps list export with winget and its fallbacks (P4-4).

Command lists, the winget runs in their current order, output handling and the
job log lines. The GUI queue, dialogs and the "Possible solutions" error text
stay with the caller (decision D1=A). Commands, kwargs, encodings and log text
are kept character for character.

Order: --version (diagnostic only) -> export (streamed) -> basic export ->
list --accept-source-agreements -> list -> RuntimeError.
"""
import logging
import os
import time
from collections import namedtuple

from ezbak.winsys import proc

_log = logging.getLogger(__name__)

_STAMP = "%Y%m%d_%H%M%S"
_RUN_KW = dict(shell=False, capture_output=True, text=True, encoding='utf-8', errors='ignore')

# kind: "export" | "basic" | "list" | "simple"; path: file written; rc: primary export code.
AppsExport = namedtuple("AppsExport", "kind path rc")


def build_version_cmd():
    """winget command printing its version (diagnostics)."""
    return ['winget', '--version']


def build_export_cmd(output_path, basic=False):
    """winget export command; basic=True drops every option but --output."""
    if basic:
        return ['winget', 'export', '--output', output_path]
    return ['winget', 'export', '--output', output_path, '--include-versions',
            '--accept-source-agreements', '--disable-interactivity']


def build_list_cmd():
    """winget list accepting source agreements (first text fallback)."""
    return ['winget', 'list', '--accept-source-agreements']


def build_simple_list_cmd():
    """Plain winget list (last fallback)."""
    return ['winget', 'list']


def export_file_name(now):
    """Primary output file name: winget_export_<YYYYmmdd_HHMMSS>.json."""
    return f"winget_export_{now.strftime(_STAMP)}.json"


def _log_version(out):
    try:
        ver = proc.run(build_version_cmd(), **_RUN_KW)
        out(f"winget --version -> rc={ver.returncode}, out={ver.stdout.strip() or ver.stderr.strip()}")
    except Exception:  # diagnostics only: a missing winget shows up at the export run
        _log.debug("apps.version unavailable")


def _run_primary(output_path, out):
    cmd = build_export_cmd(output_path)
    out(f"Executing winget export command: {' '.join(cmd)}")
    process = proc.stream(cmd, shell=False, stdout=proc.PIPE, stderr=proc.PIPE,
                          universal_newlines=True, encoding='utf-8', errors='ignore')
    while True:
        line = process.stdout.readline()
        if not line:
            if process.poll() is not None:
                break
            time.sleep(0.05)
            continue
        out(line.strip())
    err = process.stderr.read()
    if err:
        out(f"winget export stderr: {err.strip()}")
    rc = process.returncode
    out(f"winget export completed with return code: {rc}")
    return rc


def _try_basic(output_path, out):
    try:
        basic_cmd = build_export_cmd(output_path, basic=True)
        out(f"Executing basic winget export: {' '.join(basic_cmd)}")
        basic_result = proc.run(basic_cmd, timeout=300, **_RUN_KW)
        if basic_result.returncode == 0:
            out("Basic winget export succeeded!")
            out(f"Winget export completed: {output_path}")
            return True
        out(f"Basic export also failed with code {basic_result.returncode}")
        if basic_result.stderr:
            out(f"Basic export stderr: {basic_result.stderr.strip()}")
    except Exception as basic_e:  # fallback chain: log and try the next command
        out(f"Basic export attempt failed: {basic_e}")
    return False


def _try_list(folder, out, clock):
    out("Attempting fallback: winget list.")
    fb_path = os.path.join(folder, f"winget_list_{clock().strftime(_STAMP)}.txt")
    try:
        list_cmd = build_list_cmd()
        out(f"Executing winget list: {' '.join(list_cmd)}")
        p2 = proc.run(list_cmd, timeout=300, **_RUN_KW)
        if p2.returncode == 0:
            try:
                with open(fb_path, 'w', encoding='utf-8', errors='ignore') as f:
                    f.write(p2.stdout)
                out(f"Saved fallback list to {fb_path}")
                return fb_path
            except Exception as fe:  # fallback chain: log and try the next command
                out(f"Failed writing fallback file: {fe}")
        else:
            out(f"winget list fallback failed with code {p2.returncode}")
            if p2.stderr:
                out(f"winget list stderr: {p2.stderr.strip()}")
    except proc.TimeoutExpired:
        out("winget list command timed out after 300 seconds")
    except Exception as e2:  # fallback chain: log and try the next command
        out(f"winget list fallback error: {e2}")
    return None


def _try_simple(folder, out, clock):
    # A non-zero return code adds no job log line (as before); it goes to the
    # diagnostic log (g-l, P6-4).
    out("Attempting most basic winget list.")
    try:
        p3 = proc.run(build_simple_list_cmd(), timeout=180, **_RUN_KW)
        if p3.returncode == 0:
            simple_fb_path = os.path.join(folder, f"winget_simple_list_{clock().strftime(_STAMP)}.txt")
            with open(simple_fb_path, 'w', encoding='utf-8', errors='ignore') as f:
                f.write(p3.stdout)
            out(f"Saved simple fallback list to {simple_fb_path}")
            return simple_fb_path
        _log.warning("event=winget_simple_list failed rc=%s", p3.returncode)
    except Exception as e3:  # last fallback: log, the caller raises RuntimeError
        out(f"Simple winget list also failed: {e3}")
    return None


def export(output_path, out, *, clock):
    """Export the installed-apps list to output_path, falling back to winget list text.

    out(line) gets every job log line; clock() gives the time for fallback file
    names, read at the same moments as before. Returns AppsExport(kind, path, rc).
    FileNotFoundError from starting the primary export propagates; when every
    attempt fails, RuntimeError("All winget export/list attempts failed. ...").
    """
    _log_version(out)
    rc = _run_primary(output_path, out)
    if rc == 0:
        out(f"Winget export completed successfully: {output_path}")
        return _done("export", output_path, rc)
    out(f"winget export failed with code {rc}. Trying basic export command.")
    if _try_basic(output_path, out):
        return _done("basic", output_path, rc)
    folder = os.path.dirname(output_path)
    fb_path = _try_list(folder, out, clock)
    if fb_path:
        return _done("list", fb_path, rc)
    simple_path = _try_simple(folder, out, clock)
    if simple_path:
        return _done("simple", simple_path, rc)
    _log.debug("apps.export all_failed rc=%s", rc)
    raise RuntimeError(f"All winget export/list attempts failed. Primary error code: {rc}")


def _done(kind, path, rc):
    _log.debug("apps.export kind=%s rc=%s", kind, rc)
    return AppsExport(kind, path, rc)
