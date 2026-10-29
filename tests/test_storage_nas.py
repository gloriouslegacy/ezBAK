"""Unit tests for ezbak.storage.nas (P4-5). net use never runs: proc.run is faked.

The password is the dummy value "dummy-pw" only.
"""
import logging
from types import SimpleNamespace

import pytest

from ezbak.storage import nas
from ezbak.winsys import proc

UNC = "\\\\nas\\share"
PW = "dummy-pw"  # test value only, never a real secret


@pytest.fixture
def fake_run(monkeypatch):
    """Install a proc.run stand-in answering from a list (exception or result per call)."""
    calls = []

    def install(*answers):
        queue = list(answers)

        def run(argv, **kwargs):
            calls.append((argv, kwargs))
            answer = queue.pop(0)
            if isinstance(answer, BaseException):
                raise answer
            return answer
        monkeypatch.setattr(proc, "run", run)
        return calls
    return install


def _done(rc=0, stdout="", stderr=""):
    return SimpleNamespace(returncode=rc, stdout=stdout, stderr=stderr)


@pytest.mark.parametrize("drive, user, persistent, expected", [
    ("Z:", "nasuser", True, ["net", "use", "Z:", UNC, "/user:nasuser", PW, "/persistent:yes"]),
    ("Z:", "", False, ["net", "use", "Z:", UNC, "/persistent:no"]),
    (None, "nasuser", False, ["net", "use", UNC, "/user:nasuser", PW, "/persistent:no"]),
    (None, "", True, ["net", "use", UNC, "/persistent:yes"]),
])
def test_build_connect_cmd(drive, user, persistent, expected):
    spec = nas.NasSpec(UNC, drive, user, PW if user else None, persistent)
    assert nas.build_connect_cmd(spec) == expected


def test_nasspec_repr_hides_password():
    spec = nas.NasSpec(UNC, "Z:", "nasuser", PW, True)
    assert PW not in repr(spec)
    assert PW not in str(spec)


@pytest.mark.parametrize("text, expected", [
    (UNC, (["net", "use", UNC, "/delete", "/y"], UNC)),
    ("z", (["net", "use", "Z:", "/delete", "/y"], "Z:")),
    ("z:", (["net", "use", "Z:", "/delete", "/y"], "Z:")),
    ("Y::", (["net", "use", "Y:", "/delete", "/y"], "Y:")),
    ("zz", None),
    ("1", None),
    ("", None),
])
def test_disconnect_target(text, expected):
    assert nas.disconnect_target(text) == expected


def test_run_net_use_cp949_ok(fake_run):
    calls = fake_run(_done(0, " ok \r\n", ""))
    res = nas.run_net_use(["net", "use", UNC])
    assert calls == [(["net", "use", UNC], {"shell": False, "capture_output": True, "text": True,
                                             "encoding": "cp949", "errors": "ignore"})]
    assert res == nas.NetUseResult(0, "ok", "", "")


def test_run_net_use_retries_without_encoding(fake_run):
    calls = fake_run(OSError("boom"), _done(2, "", ""))
    res = nas.run_net_use(["net", "use", UNC])
    assert [kw for _, kw in calls] == [
        {"shell": False, "capture_output": True, "text": True, "encoding": "cp949", "errors": "ignore"},
        {"shell": False, "capture_output": True, "text": True}]
    assert res.message == "Failed with code 2"


def test_run_net_use_second_failure_propagates(fake_run):
    calls = fake_run(OSError("first"), OSError("second"))
    with pytest.raises(OSError, match="second"):
        nas.run_net_use(["net", "use", UNC])
    assert len(calls) == 2


@pytest.mark.parametrize("rc, out, err, message", [
    (2, "out text", "System error 53", "System error 53"),
    (1, "not found", "", "not found"),
    (2, "", "", "Failed with code 2"),
    (0, "done", "warn", ""),
])
def test_run_net_use_message(fake_run, rc, out, err, message):
    fake_run(_done(rc, out + "\r\n", err))
    assert nas.run_net_use(["net", "use", UNC]).message == message


def test_password_not_in_result_or_logs(fake_run, caplog):
    caplog.set_level(logging.DEBUG)
    argv = nas.build_connect_cmd(nas.NasSpec(UNC, "Z:", "nasuser", PW, True))
    fake_run(OSError("raised"), _done(2, "", ""))
    res = nas.run_net_use(argv)
    assert PW not in repr(res)
    assert PW not in res.message
    assert PW not in caplog.text
    assert PW not in "".join(str(r.args) for r in caplog.records)


@pytest.mark.parametrize("used, expected", [
    (set(), "Z"),
    ({"Z", "Y"}, "X"),
    ({chr(c) for c in range(68, 91)}, "C"),
    ({chr(c) for c in range(67, 91)}, "B"),
    ({chr(c) for c in range(66, 91)}, "Z"),
])
def test_free_drive_letter(used, expected):
    assert nas.free_drive_letter(used) == expected
