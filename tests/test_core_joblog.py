"""Unit tests for ezbak.core job log writer and header wording (flag values: test_core_model)."""
import io
from dataclasses import fields
from datetime import datetime, timedelta

import pytest

from ezbak.core import (
    DEFAULT_FILTERS,
    JobLine,
    JobLog,
    JobLogFormat,
    LogItem,
    format_bytes,
    header_lines,
)

SEP = "-" * 50


class StepClock:
    """Each call returns the next second, starting at 2024-01-02 03:04:05."""

    def __init__(self):
        self.calls = 0

    def __call__(self):
        value = datetime(2024, 1, 2, 3, 4, 5) + timedelta(seconds=self.calls)
        self.calls += 1
        return value


def _log():
    stream = io.StringIO()
    clock = StepClock()
    return JobLog(stream, clock), stream, clock


# --- JobLog line format, lead / end ---------------------------------------

def test_write_line_format():
    log, stream, _ = _log()
    log.write("hello")
    assert stream.getvalue() == "[2024-01-02 03:04:05] hello\n"


def test_write_lead_and_end():
    log, stream, _ = _log()
    log.write("A", lead="\n")
    log.write("B", end="")
    log.write("C")
    assert stream.getvalue() == ("\n[2024-01-02 03:04:05] A\n"
                                 "[2024-01-02 03:04:06] B"
                                 "[2024-01-02 03:04:07] C\n")


def test_emit_item_and_none():
    log, stream, clock = _log()
    log.emit(None, x=1)
    assert stream.getvalue() == "" and clock.calls == 0
    log.emit(LogItem("x={x}", lead="\n", end=""), x=1)
    assert stream.getvalue() == "\n[2024-01-02 03:04:05] x=1"


def test_write_lines_clock_per_line_or_single():
    lines = [JobLine("a"), JobLine("b")]
    log, stream, clock = _log()
    log.write_lines(lines)
    assert stream.getvalue() == "[2024-01-02 03:04:05] a\n[2024-01-02 03:04:06] b\n"
    assert clock.calls == 2

    log, stream, clock = _log()
    log.write_lines(lines, single_timestamp=True)
    assert stream.getvalue() == "[2024-01-02 03:04:05] a\n[2024-01-02 03:04:05] b\n"
    assert clock.calls == 1


def test_write_lines_keeps_lines_before_a_failure():
    log, stream, _ = _log()
    bad = {'include': [{'type': None, 'pattern': 'x'}], 'exclude': []}
    with pytest.raises(AttributeError):
        log.write_lines(header_lines(JobLogFormat.GUI, hidden='exclude', system='exclude',
                                     backup_keep='2', log_keep='30', filters=bad))
    assert stream.getvalue().splitlines() == [
        "[2024-01-02 03:04:05] Policy: Hidden=exclude, System=exclude, BackupKeep=2, LogKeep=30"]


# --- header_lines --------------------------------------------------------

def test_header_gui_backup():
    filters = {'include': [{'type': 'EXT', 'pattern': '.txt'}],
               'exclude': [{'type': 'name', 'pattern': 'tmp'}, {'type': 'Path', 'pattern': 'a\\b'}]}
    lines = list(header_lines(JobLogFormat.GUI, hidden='include', system='exclude',
                              backup_keep='3', log_keep='7', filters=filters))
    assert [l.message for l in lines] == [
        "Policy: Hidden=include, System=exclude, BackupKeep=3, LogKeep=7",
        "Filters: Include=[ext:.txt], Exclude=[name:tmp, path:a\\b]",
        SEP,
    ]
    assert all(l.lead == "" and l.end == "\n" for l in lines)


def test_header_restore_has_no_filters_line():
    lines = list(header_lines(JobLogFormat.GUI, hidden='exclude', system='exclude',
                              backup_keep='2', log_keep='30', filters=None))
    assert [l.message for l in lines] == [
        "Policy: Hidden=exclude, System=exclude, BackupKeep=2, LogKeep=30", SEP]


def test_header_empty_filters_are_none():
    lines = list(header_lines(JobLogFormat.GUI, hidden='exclude', system='exclude',
                              backup_keep='2', log_keep='30',
                              filters={'include': [], 'exclude': None}))
    assert lines[1].message == "Filters: Include=[None], Exclude=[None]"


def test_header_headless_with_start():
    lines = list(header_lines(JobLogFormat.HEADLESS,
                              start={'user': 'bob', 'backup_path': 'D:\\bk\\bob_backup_x'},
                              hidden='exclude', system='include', backup_keep=2, log_keep=30,
                              filters=DEFAULT_FILTERS))
    assert [l.message for l in lines] == [
        "Backup start for 'bob' \u2192 D:\\bk\\bob_backup_x",
        "Policy: Hidden=exclude, System=include, BackupKeep=2, LogKeep=30",
        "Filters: Include=[None], Exclude=[name:onedrive*, name:thumbs.db, "
        "name:desktop.ini, name:$recycle.bin]",
        SEP,
    ]


def test_header_without_start_values_skips_start():
    lines = list(header_lines(JobLogFormat.HEADLESS, hidden='exclude', system='exclude',
                              backup_keep=2, log_keep=30))
    assert lines[0].message.startswith("Policy:")


def test_headless_items_are_plain_lines():
    """P2-fix #3 (Q1 exception): every Headless item is one `[ts] msg\\n` line."""
    fmt = JobLogFormat.HEADLESS
    items = {f.name: getattr(fmt, f.name) for f in fields(fmt)}
    bad = {name: (item.lead, item.end) for name, item in items.items()
           if item is not None and (item.lead != "" or item.end != "\n")}
    assert bad == {}
    assert fmt.bm_saved_last == fmt.bm_saved and fmt.bm_missing_last == fmt.bm_missing


def test_gui_start_line():
    line = JobLogFormat.GUI.start.line(source='C:\\Users\\bob', backup_path='D:\\bk')
    assert line.message == "Backup start: C:\\Users\\bob -> D:\\bk"


# --- format_bytes ----------------------------------------------------------

@pytest.mark.parametrize("size, text", [
    (0, "0.00 B"),
    (1023, "1023.00 B"),
    (1024, "1.00 KB"),
    (1536, "1.50 KB"),
    (1024 ** 3 * 5, "5.00 GB"),
    (1024 ** 5, "1.00 PB"),
    (1024 ** 6, "1024.00 PB"),
    ("2048", "2.00 KB"),
    (None, "N/A"),
    ("abc", "N/A"),
])
def test_format_bytes(size, text):
    assert format_bytes(size) == text
