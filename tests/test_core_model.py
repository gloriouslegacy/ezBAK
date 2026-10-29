"""Unit tests for ezbak.core domain types (model)."""
import dataclasses

import pytest

from ezbak.core import (
    DEFAULT_FILTERS,
    BackupJob,
    BackupPolicy,
    BackupResult,
    FilterRule,
    FilterSet,
)

# Value of ezBAK.py DEFAULT_FILTERS before it moved to ezbak.core (Phase 0, 99e4d37).
LEGACY_DEFAULT_FILTERS = {
    'include': [],
    'exclude': [
        {'type': 'name', 'pattern': 'onedrive*'},
        {'type': 'name', 'pattern': 'thumbs.db'},
        {'type': 'name', 'pattern': 'desktop.ini'},
        {'type': 'name', 'pattern': '$recycle.bin'},
    ],
}


def test_default_filters_unchanged():
    assert DEFAULT_FILTERS == LEGACY_DEFAULT_FILTERS
    assert list(DEFAULT_FILTERS) == ['include', 'exclude']


def test_monolith_uses_core_default_filters(ezbak_mod):
    assert ezbak_mod.DEFAULT_FILTERS is DEFAULT_FILTERS
    assert ezbak_mod.copy_default_filters() == LEGACY_DEFAULT_FILTERS


@pytest.mark.parametrize("raw", ["NAME", "Name", " name "])
def test_filter_rule_type_is_lowercased(raw):
    assert FilterRule(raw, "x").type == "name"


@pytest.mark.parametrize("raw", ["", None, "glob", "extension"])
def test_filter_rule_rejects_unknown_type(raw):
    with pytest.raises(ValueError):
        FilterRule(raw, "x")


def test_filter_rule_from_dict_drops_unknown():
    assert FilterRule.from_dict({'type': 'EXT', 'pattern': '*.tmp'}) == FilterRule('ext', '*.tmp')
    assert FilterRule.from_dict({'type': 'regex', 'pattern': 'a'}) is None
    assert FilterRule.from_dict({'pattern': 'a'}) is None
    assert FilterRule.from_dict("name") is None


def test_filter_rule_is_frozen():
    with pytest.raises(dataclasses.FrozenInstanceError):
        FilterRule('ext', 'x').type = 'name'


def test_filter_set_round_trip_default():
    fs = FilterSet.from_dict(DEFAULT_FILTERS)
    assert len(fs.exclude) == 4 and fs.include == ()
    assert fs.to_dict() == LEGACY_DEFAULT_FILTERS


def test_filter_set_from_dict_normalizes_and_drops():
    fs = FilterSet.from_dict({
        'include': [{'type': 'EXT', 'pattern': 'docx'}, {'type': 'bad', 'pattern': 'x'}],
        'exclude': [{'type': 'Path', 'pattern': 'C:\\Temp'}],
    })
    assert fs.to_dict() == {
        'include': [{'type': 'ext', 'pattern': 'docx'}],
        'exclude': [{'type': 'path', 'pattern': 'C:\\Temp'}],
    }


@pytest.mark.parametrize("raw", [None, {}, {'include': None, 'exclude': None}])
def test_filter_set_empty_inputs(raw):
    assert FilterSet.from_dict(raw) == FilterSet()
    assert FilterSet.from_dict(raw).to_dict() == {'include': [], 'exclude': []}


def test_backup_job_defaults_match_core_run_backup():
    job = BackupJob(user="u", dest="d")
    assert (job.include_hidden, job.include_system) == (False, False)
    assert (job.retention_count, job.log_retention_days) == (2, 30)
    assert job.filters is None and job.log_folder is None


def test_backup_result_fields():
    r = BackupResult("b", "l", files=1, folders=2, bytes=3)
    assert (r.backup_path, r.log_path, r.files, r.folders, r.bytes) == ("b", "l", 1, 2, 3)


POLICY_FLAGS = {
    # flag: (GUI, HEADLESS)
    "skip_builtin_names": (False, True),
    "skip_appdata_names": (False, True),
    "attr_ctypes_fallback": (False, True),
    "none_filters_use_default": (False, True),
    "filter_requires_existing_path": (True, False),
    "check_free_space": (True, False),
    "report_filtered_separately": (False, True),
    "count_attempted_copies": (True, False),
    "isolate_bookmark_errors": (True, False),
    "log_cleanup_retry_pass": (True, False),
    "isolate_backup_delete_errors": (False, True),
    "skip_errors_include": (True, False),
    "isolate_retention_errors": (True, False),
    "return_critical_errors": (True, False),
}


def test_backup_policy_flag_set():
    assert {f.name for f in dataclasses.fields(BackupPolicy)} == set(POLICY_FLAGS)


@pytest.mark.parametrize("flag", sorted(POLICY_FLAGS))
def test_backup_policy_constants(flag):
    gui, headless = POLICY_FLAGS[flag]
    assert getattr(BackupPolicy.GUI, flag) is gui
    assert getattr(BackupPolicy.HEADLESS, flag) is headless


def test_backup_policy_variant_for_is_hidden():
    variant = dataclasses.replace(BackupPolicy.GUI, skip_appdata_names=True, attr_ctypes_fallback=True)
    assert variant not in (BackupPolicy.GUI, BackupPolicy.HEADLESS)
    assert BackupPolicy.GUI.skip_appdata_names is False
