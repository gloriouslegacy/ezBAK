"""Unit tests for ezbak.automation.cli (argparse moved as-is, P3-9a)."""
import pytest

from ezbak.automation import cli

DEFAULTS = {
    "schedule_config": None,
    "user": None,
    "dest": None,
    "include_hidden": False,
    "include_system": False,
    "retention": 2,
    "log_retention": 30,
    "log_folder": "logs",
    "filters": None,
}


def test_defaults_table():
    args = cli.parse([])
    assert vars(args) == DEFAULTS


def test_dest_order_matches_current_parser():
    dests = [a.dest for a in cli.build_parser()._actions if a.dest != "help"]
    assert dests == list(DEFAULTS)


def test_unknown_arguments_are_ignored():
    args = cli.parse(["--auto-backup", "--user", "u", "extra", "--foo=1"])
    assert args.user == "u"
    assert not hasattr(args, "auto_backup")


def test_retention_not_int_exits_2(capsys):
    with pytest.raises(SystemExit) as exc:
        cli.parse(["--retention", "abc"])
    assert exc.value.code == 2
    err = capsys.readouterr().err
    assert "usage:" in err
    assert "argument --retention: invalid int value: 'abc'" in err


def test_schedule_config_path():
    path = r"C:\Users\x\AppData\Roaming\ezBAK\schedules\task.json"
    args = cli.parse(["--schedule-config", path])
    assert args.schedule_config == path
    assert args.user is None and args.dest is None


def test_direct_arguments_combination():
    args = cli.parse([
        "--user", "alice", "--dest", r"D:\bk", "--include-hidden", "--include-system",
        "--retention", "5", "--log-retention", "7", "--log-folder", "L",
        "--filters", '{"include": []}',
    ])
    assert vars(args) == {
        "schedule_config": None,
        "user": "alice",
        "dest": r"D:\bk",
        "include_hidden": True,
        "include_system": True,
        "retention": 5,
        "log_retention": 7,
        "log_folder": "L",
        "filters": '{"include": []}',
    }


def test_parse_defaults_to_sys_argv(monkeypatch):
    monkeypatch.setattr("sys.argv", ["ezBAK.py", "--user", "bob", "--auto-backup"])
    args = cli.parse()
    assert args.user == "bob"


def test_prog_follows_sys_argv_at_call_time(monkeypatch):
    monkeypatch.setattr("sys.argv", ["ezBAK.exe"])
    assert cli.build_parser().format_usage().startswith("usage: ezBAK.exe ")
