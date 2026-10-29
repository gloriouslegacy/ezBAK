"""Unit tests for ezbak.targets.apps (P4-4). winget never runs: proc.run and proc.stream are faked."""
import logging
import os
from datetime import datetime
from types import SimpleNamespace

import pytest

from ezbak.targets import apps
from ezbak.winsys import proc

NOW = datetime(2025, 1, 15, 12, 0, 5)
RUN_KW = dict(shell=False, capture_output=True, text=True, encoding='utf-8', errors='ignore')


class FakeProcess:
    """Popen stand-in: stdout lines, then '' with poll() done; stderr read once."""

    def __init__(self, lines, stderr="", returncode=0):
        self._lines = list(lines)
        self.returncode = returncode
        self.stdout = SimpleNamespace(readline=lambda: self._lines.pop(0) if self._lines else "")
        self.stderr = SimpleNamespace(read=lambda: stderr)

    def poll(self):
        return self.returncode


@pytest.fixture
def winget(monkeypatch):
    """Answers by command: runs[key] is a CompletedProcess-like or an exception."""
    state = SimpleNamespace(calls=[], runs={}, primary=FakeProcess(["Done\n"]))

    def key(argv):
        return " ".join(argv[1:3]) if argv[1] != "--version" else "--version"

    def run(argv, **kwargs):
        state.calls.append(("run", argv, kwargs))
        answer = state.runs.get(key(argv), SimpleNamespace(returncode=0, stdout="", stderr=""))
        if isinstance(answer, BaseException):
            raise answer
        return answer

    def stream(argv, **kwargs):
        state.calls.append(("stream", argv, kwargs))
        if isinstance(state.primary, BaseException):
            raise state.primary
        return state.primary

    monkeypatch.setattr(proc, "run", run)
    monkeypatch.setattr(proc, "stream", stream)
    return state


def done(rc=0, stdout="", stderr=""):
    return SimpleNamespace(returncode=rc, stdout=stdout, stderr=stderr)


def clock_counter():
    reads = []

    def clock():
        reads.append(NOW)
        return NOW
    return clock, reads


def test_commands():
    assert apps.build_version_cmd() == ['winget', '--version']
    assert apps.build_export_cmd("o.json") == [
        'winget', 'export', '--output', 'o.json', '--include-versions',
        '--accept-source-agreements', '--disable-interactivity']
    assert apps.build_export_cmd("o.json", basic=True) == ['winget', 'export', '--output', 'o.json']
    assert apps.build_list_cmd() == ['winget', 'list', '--accept-source-agreements']
    assert apps.build_simple_list_cmd() == ['winget', 'list']


def test_export_file_name():
    assert apps.export_file_name(NOW) == "winget_export_20250115_120005.json"


def test_primary_ok(winget, tmp_path):
    out_path = str(tmp_path / "o.json")
    lines = []
    clock, reads = clock_counter()
    winget.runs["--version"] = done(stdout="v1.9.0\n")
    result = apps.export(out_path, lines.append, clock=clock)
    assert result == apps.AppsExport("export", out_path, 0)
    assert reads == []
    assert [c[0] for c in winget.calls] == ["run", "stream"]
    assert winget.calls[0][2] == RUN_KW
    assert winget.calls[1][2] == dict(shell=False, stdout=proc.PIPE, stderr=proc.PIPE,
                                      universal_newlines=True, encoding='utf-8', errors='ignore')
    assert lines == ["winget --version -> rc=0, out=v1.9.0",
                     f"Executing winget export command: {' '.join(apps.build_export_cmd(out_path))}",
                     "Done", "winget export completed with return code: 0",
                     f"Winget export completed successfully: {out_path}"]


def _fail_primary(winget):
    winget.primary = FakeProcess([], stderr="export error\n", returncode=1)


def test_basic_ok(winget, tmp_path):
    _fail_primary(winget)
    out_path = str(tmp_path / "o.json")
    lines = []
    result = apps.export(out_path, lines.append, clock=clock_counter()[0])
    assert result == apps.AppsExport("basic", out_path, 1)
    assert winget.calls[2][1] == apps.build_export_cmd(out_path, basic=True)
    assert winget.calls[2][2] == dict(timeout=300, **RUN_KW)
    assert lines[-2:] == ["Basic winget export succeeded!", f"Winget export completed: {out_path}"]


def test_list_ok_writes_file(winget, tmp_path):
    _fail_primary(winget)
    winget.runs["export --output"] = done(2, stderr="basic error\n")
    winget.runs["list --accept-source-agreements"] = done(stdout="Name Id\n")
    clock, reads = clock_counter()
    result = apps.export(str(tmp_path / "o.json"), [].append, clock=clock)
    path = os.path.join(str(tmp_path), "winget_list_20250115_120005.txt")
    assert result == apps.AppsExport("list", path, 1)
    assert len(reads) == 1
    assert winget.calls[3][2] == dict(timeout=300, **RUN_KW)
    with open(path, encoding='utf-8') as f:
        assert f.read() == "Name Id\n"


def test_fallback_order_and_simple_ok(winget, tmp_path):
    _fail_primary(winget)
    winget.runs["export --output"] = done(2)
    winget.runs["list --accept-source-agreements"] = proc.TimeoutExpired(["winget"], 300)
    winget.runs["list"] = done(stdout="Name Id\n")
    lines = []
    clock, reads = clock_counter()
    result = apps.export(str(tmp_path / "o.json"), lines.append, clock=clock)
    path = os.path.join(str(tmp_path), "winget_simple_list_20250115_120005.txt")
    assert result == apps.AppsExport("simple", path, 1)
    assert len(reads) == 2
    assert [c[1][1:] for c in winget.calls] == [
        ['--version'], apps.build_export_cmd(str(tmp_path / "o.json"))[1:],
        ['export', '--output', str(tmp_path / "o.json")], ['list', '--accept-source-agreements'], ['list']]
    assert winget.calls[4][2] == dict(timeout=180, **RUN_KW)
    assert "winget list command timed out after 300 seconds" in lines
    assert lines[-1] == f"Saved simple fallback list to {path}"


def test_all_fail_raises_current_text(winget, tmp_path, caplog):
    _fail_primary(winget)
    winget.runs["export --output"] = OSError("basic raised")
    winget.runs["list --accept-source-agreements"] = done(3)
    winget.runs["list"] = done(4, stderr="simple error\n")
    lines = []
    with pytest.raises(RuntimeError, match=r"^All winget export/list attempts failed\. Primary error code: 1$"):
        apps.export(str(tmp_path / "o.json"), lines.append, clock=clock_counter()[0])
    assert "Basic export attempt failed: basic raised" in lines
    assert lines[-1] == "Attempting most basic winget list."  # simple rc != 0: no job log line
    assert os.listdir(tmp_path) == []
    assert ("ezbak.targets.apps", logging.WARNING,
            "event=winget_simple_list failed rc=4") in caplog.record_tuples  # g-l, P6-4


def test_file_not_found_propagates(winget, tmp_path):
    winget.runs["--version"] = FileNotFoundError(2, "missing")
    winget.primary = FileNotFoundError(2, "missing")
    lines = []
    with pytest.raises(FileNotFoundError):
        apps.export(str(tmp_path / "o.json"), lines.append, clock=clock_counter()[0])
    assert [c[0] for c in winget.calls] == ["run", "stream"]
    assert lines == [f"Executing winget export command: "
                     f"{' '.join(apps.build_export_cmd(str(tmp_path / 'o.json')))}"]
