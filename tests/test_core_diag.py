"""P6-1: ezbak.core.diag — diagnostic log handler, level, run_id (ADR 0003 §2).

setup() writes only under tmp_path here (log_dir), never the real %APPDATA%; every
test removes the handler again (teardown) so the session keeps no file handler.
"""
import logging
import random
import threading
from datetime import datetime

import pytest

from ezbak.core import diag


@pytest.fixture(autouse=True)
def clean():
    diag.teardown()
    token_level = logging.getLogger("ezbak").level
    yield
    diag.teardown()
    logging.getLogger("ezbak").setLevel(token_level)
    diag.bind(diag.NO_RUN_ID)


def lines(path):
    return path.read_text(encoding="utf-8").splitlines()


def test_default_log_dir(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert diag.default_log_dir() == str(tmp_path / "ezBAK" / "logs")
    monkeypatch.delenv("APPDATA")
    assert diag.default_log_dir() == ""


def test_setup_without_appdata_attaches_nothing(monkeypatch):
    monkeypatch.delenv("APPDATA", raising=False)
    assert diag.setup() is None
    assert diag._existing_handler(logging.getLogger("ezbak")) is None


def test_setup_writes_the_adr_format(tmp_path):
    handler = diag.setup(log_dir=str(tmp_path / "logs"), level=logging.INFO)
    assert handler.maxBytes == 1024 * 1024 and handler.backupCount == 3
    logging.getLogger("ezbak.core.test").info("event=hello")
    logging.getLogger("ezbak.core.test").debug("event=hidden")
    handler.flush()
    (line,) = lines(tmp_path / "logs" / "ezbak.log")
    assert line.endswith(" INFO  ezbak.core.test [-] event=hello")


def test_setup_twice_keeps_one_handler(tmp_path):
    first = diag.setup(log_dir=str(tmp_path))
    assert diag.setup(log_dir=str(tmp_path / "other")) is first
    assert [h for h in logging.getLogger("ezbak").handlers if getattr(h, "_ezbak_diag", False)] == [first]


def test_other_loggers_are_not_written(tmp_path):
    handler = diag.setup(log_dir=str(tmp_path), level=logging.DEBUG)
    logging.getLogger("other.lib").warning("not ours")
    handler.flush()
    assert lines(tmp_path / "ezbak.log") == []


def test_folder_failure_returns_none(tmp_path, capsys):
    blocker = tmp_path / "file"
    blocker.write_text("x")
    assert diag.setup(log_dir=str(blocker / "logs")) is None
    assert "diagnostic log disabled" in capsys.readouterr().err


@pytest.mark.parametrize("value, expected", [
    (None, logging.INFO), ("", logging.INFO), ("debug", logging.DEBUG),
    (" WARNING ", logging.WARNING), ("nonsense", logging.INFO)])
def test_level_from_env(value, expected):
    environ = {} if value is None else {"EZBAK_LOG_LEVEL": value}
    assert diag.level_from_env(environ) == expected


def test_setup_reads_the_env_level(tmp_path):
    diag.setup(log_dir=str(tmp_path), environ={"EZBAK_LOG_LEVEL": "DEBUG"})
    assert logging.getLogger("ezbak").level == logging.DEBUG


def test_new_run_id_shape():
    clock = lambda: datetime(2026, 10, 27, 9, 5, 7)
    assert diag.new_run_id(clock, random.Random(0)) == "20261027-090507-" + f"{random.Random(0).getrandbits(16):04x}"
    assert len(diag.new_run_id(clock)) == len("20261027-090507-abcd")


def test_bound_run_id_is_written_and_not_inherited_by_threads(tmp_path):
    handler = diag.setup(log_dir=str(tmp_path), level=logging.INFO)
    log = logging.getLogger("ezbak.core.test")
    diag.bind("20261027-090507-00ff")
    log.info("event=main")
    seen = []

    def work():
        seen.append(diag.current())
        diag.bind("20261027-090507-00ff")
        log.info("event=thread")
    thread = threading.Thread(target=work)
    thread.start()
    thread.join()
    handler.flush()
    assert seen == ["-"]
    assert [line[line.index("["):] for line in lines(tmp_path / "ezbak.log")] == [
        "[20261027-090507-00ff] event=main", "[20261027-090507-00ff] event=thread"]
