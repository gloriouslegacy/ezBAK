"""U7-2: ezbak.core.history — last backup in %APPDATA%\\ezBAK\\state.json (decision D1 A)."""
import json
import logging
import os
from datetime import datetime

from ezbak.core import history

WHEN = datetime(2026, 10, 27, 9, 5, 7)


def test_state_path(isolated_env, monkeypatch):
    assert history.state_path() == os.path.join(str(isolated_env["APPDATA"]), "ezBAK", "state.json")
    monkeypatch.delenv("APPDATA")
    assert history.state_path() == ""


def test_record_and_read(isolated_env):
    assert history.last_backup() is None
    assert history.record_backup(user="kim", dest="D:\\B", folder="D:\\B\\kim_backup_x",
                                 when=WHEN, mode="gui") is True
    assert history.last_backup() == {"time": "2026-10-27 09:05:07", "user": "kim", "dest": "D:\\B",
                                     "folder": "D:\\B\\kim_backup_x", "mode": "gui"}
    assert not os.path.exists(history.state_path() + ".tmp")


def test_record_keeps_other_keys_and_replaces_last(isolated_env):
    path = history.state_path()
    os.makedirs(os.path.dirname(path))
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"other": 1, "last_backup": {"time": "old"}}, f)
    history.record_backup(user="lee", dest="E:\\", folder="E:\\lee_backup_y", when=WHEN, mode="headless")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    assert data["other"] == 1 and data["last_backup"]["user"] == "lee"


def test_broken_or_odd_file_reads_as_none(isolated_env):
    path = history.state_path()
    os.makedirs(os.path.dirname(path))
    for text in ("{broken", "[1, 2]", '{"last_backup": "x"}', '{"last_backup": {"time": 5}}'):
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        assert history.last_backup() is None, text
    assert history.record_backup(user="u", dest="d", folder="f", when=WHEN, mode="gui")  # rewrites it


def test_without_appdata_nothing_is_written(isolated_env, monkeypatch):
    monkeypatch.delenv("APPDATA")
    assert history.record_backup(user="u", dest="d", folder="f", when=WHEN, mode="gui") is False
    assert history.last_backup() is None


def test_write_failure_is_a_warning(isolated_env, caplog):
    blocker = os.path.join(str(isolated_env["APPDATA"]), "ezBAK")
    with open(blocker, "w") as f:
        f.write("not a folder")
    with caplog.at_level(logging.WARNING, logger="ezbak.core.history"):
        assert history.record_backup(user="u", dest="d", folder="f", when=WHEN, mode="gui") is False
    assert any("event=last_backup_write failed" in m for m in caplog.messages)
