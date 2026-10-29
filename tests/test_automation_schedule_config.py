"""Unit tests for automation.schedule_config (P3-5)."""
import json
import os
from datetime import datetime

import pytest

from ezbak.automation import ScheduleSpec, schedule_config

NOW = datetime(2026, 10, 12, 9, 30, 5)
KEYS = ['task_name', 'user', 'dest', 'schedule', 'time_str', 'monthly_day',
        'include_hidden', 'include_system', 'retention_count', 'log_retention_days',
        'filters', 'created_at']


def _spec(**kw):
    values = dict(task_name="ezBAK_Daily", user="홍길동", dest="D:\\백업", schedule="DAILY",
                  time_str="14:30", filters=[{"type": "Extension", "mode": "exclude", "pattern": ".TMP"}])
    values.update(kw)
    return ScheduleSpec(**values)


def test_config_dir_under_appdata(isolated_env):
    assert schedule_config.config_dir() == os.path.join(str(isolated_env["APPDATA"]), 'ezBAK', 'ezBAK_Schedule')


@pytest.mark.parametrize("value", [None, ""])
def test_config_dir_without_appdata_raises(monkeypatch, value):
    if value is None:
        monkeypatch.delenv("APPDATA", raising=False)
    else:
        monkeypatch.setenv("APPDATA", value)
    with pytest.raises(RuntimeError, match="^APPDATA environment variable not found$"):
        schedule_config.config_dir()
    with pytest.raises(RuntimeError):
        schedule_config.save(_spec(), now=NOW)


def test_config_path_replaces_invalid_chars(isolated_env):
    path = schedule_config.config_path('a<b>c:d"e/f\\g|h?i*j')
    assert os.path.basename(path) == "a_b_c_d_e_f_g_h_i_j.json"
    assert os.path.dirname(path) == schedule_config.config_dir()


def test_save_writes_fixed_key_order_indent_and_utf8(isolated_env):
    path = schedule_config.save(_spec(), now=NOW)
    assert path == schedule_config.config_path("ezBAK_Daily")
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    data = json.loads(text, object_pairs_hook=lambda pairs: pairs)
    assert [k for k, _ in data] == KEYS
    assert "홍길동" in text and "\\u" not in text  # ensure_ascii=False
    assert text.startswith('{\n  "task_name": "ezBAK_Daily",\n')  # indent=2
    expected = {**_spec().to_dict("2026-10-12 09:30:05")}
    assert text == json.dumps(expected, indent=2, ensure_ascii=False)


def test_save_keeps_defaults_and_filter_type_case(isolated_env):
    spec = _spec(filters=[{"type": "FOLDER", "mode": "Exclude", "pattern": "Temp"}])
    data = schedule_config.load(schedule_config.save(spec, now=NOW))
    assert data["filters"] == [{"type": "FOLDER", "mode": "Exclude", "pattern": "Temp"}]
    assert (data["monthly_day"], data["include_hidden"], data["include_system"],
            data["retention_count"], data["log_retention_days"]) == (1, False, False, 2, 30)
    assert data["created_at"] == "2026-10-12 09:30:05"


def test_save_load_roundtrip(isolated_env):
    spec = _spec(schedule="MONTHLY", monthly_day=15, include_hidden=True, retention_count=5)
    data = schedule_config.load(schedule_config.save(spec, now=NOW))
    assert data == spec.to_dict("2026-10-12 09:30:05")
    assert list(data) == KEYS


def test_load_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        schedule_config.load(str(tmp_path / "none.json"))


def test_delete_existing_and_missing(isolated_env):
    path = schedule_config.save(_spec(task_name="a/b"), now=NOW)
    assert schedule_config.delete("a/b") == path
    assert not os.path.exists(path)
    assert schedule_config.delete("a/b") is None


def test_job_from_config_defaults():
    job = schedule_config.job_from_config({})
    assert (job.user, job.dest, job.include_hidden, job.include_system, job.retention_count,
            job.log_retention_days, job.filters, job.log_folder) == (
        None, None, False, False, 2, 30, None, None)


def test_job_from_config_keeps_values_unconverted():
    filters = {"include": [], "exclude": [{"type": "Extension", "pattern": ".TMP"}]}
    config = {"user": "u", "dest": "D:\b", "include_hidden": "yes", "include_system": 0,
              "retention_count": "3", "log_retention_days": "7", "filters": filters,
              "log_folder": "ignored"}
    job = schedule_config.job_from_config(config, log_folder="logs")
    assert job.include_hidden == "yes" and job.include_system == 0
    assert job.retention_count == "3" and job.log_retention_days == "7"
    assert job.filters is filters and job.filters["exclude"][0]["type"] == "Extension"
    assert job.log_folder == "logs"


def test_job_from_config_non_dict_raises_attribute_error():
    with pytest.raises(AttributeError):
        schedule_config.job_from_config([1, 2])


def test_read_raw_none_without_file(isolated_env):
    assert schedule_config.read_raw("ezBAK_Daily") is None


def test_read_raw_returns_bytes_by_safe_name(isolated_env):
    path = schedule_config.save(_spec(task_name="a:b"), now=NOW)
    assert os.path.basename(path) == "a_b.json"
    with open(path, 'rb') as f:
        assert schedule_config.read_raw("a:b") == f.read()


def test_restore_raw_rewrites_previous_bytes(isolated_env):
    prev = b'{"task_name": "pre-existing"}\r\n'
    path = schedule_config.save(_spec(), now=NOW)
    schedule_config.restore_raw("ezBAK_Daily", prev)
    with open(path, 'rb') as f:
        assert f.read() == prev


def test_restore_raw_none_removes_file(isolated_env):
    path = schedule_config.save(_spec(task_name="x|y"), now=NOW)
    schedule_config.restore_raw("x|y", None)
    assert not os.path.exists(path)
    schedule_config.restore_raw("x|y", None)  # no file: nothing to do


def test_load_all_reads_every_config_and_skips_broken(isolated_env):
    folder = os.path.join(str(isolated_env["APPDATA"]), 'ezBAK', 'ezBAK_Schedule')
    os.makedirs(folder)
    for name, text in (("b.json", '{"task_name": "b"}'), ("a.json", '{"task_name": "a"}'),
                       ("broken.json", "{not json"), ("notes.txt", "x")):
        with open(os.path.join(folder, name), "w", encoding="utf-8") as f:
            f.write(text)
    assert schedule_config.load_all() == [{"task_name": "a"}, {"task_name": "b"}]


def test_load_all_without_folder_or_appdata(isolated_env, monkeypatch):
    assert schedule_config.load_all() == []
    monkeypatch.delenv("APPDATA")
    assert schedule_config.load_all() == []
