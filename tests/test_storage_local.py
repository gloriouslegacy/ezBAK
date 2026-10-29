"""Unit tests for storage.LocalFolderStore (P2-3)."""
import os
import stat
from datetime import datetime

import pytest

import ezbak.winsys
from ezbak.core import BackupStore
from ezbak.storage import LocalFolderStore

FIXED = datetime(2026, 10, 12, 9, 5, 7)


def _store(tmp_path):
    return LocalFolderStore(str(tmp_path), clock=lambda: FIXED)


def test_new_backup_path_name(tmp_path):
    store = _store(tmp_path)
    assert store.new_backup_path("Alice") == os.path.join(str(tmp_path), "Alice_backup_2026-10-12_090507")
    assert not os.path.exists(store.new_backup_path("Alice"))


def test_new_log_path_name(tmp_path):
    store = _store(tmp_path)
    assert store.new_log_path("Alice") == os.path.join(str(tmp_path), "Alice_backup_2026-10-12_090507.log")


def test_list_backups_rules(tmp_path):
    for d in ("alice_backup_2026-01-02_000000", "ALICE_backup_2025-12-31_000000",
              "Alice_backup_2026-01-01_000000", "Bob_backup_2026-01-01_000000",
              "Alice_restore_2026-01-01_000000", "xAlice_backup_2026-01-01_000000"):
        (tmp_path / d).mkdir()
    (tmp_path / "Alice_backup_2026-01-03_000000.log").write_text("log")
    (tmp_path / "Alice_backup_2026-01-04_000000").write_text("file, not dir")

    names = _store(tmp_path).list_backups("Alice")

    # plain name sort: upper-case 'A' sorts before lower-case 'a'
    assert names == ["ALICE_backup_2025-12-31_000000", "Alice_backup_2026-01-01_000000",
                     "alice_backup_2026-01-02_000000"]


def test_list_backups_empty(tmp_path):
    assert _store(tmp_path).list_backups("Alice") == []


def test_delete_backup_with_readonly_file(tmp_path):
    folder = tmp_path / "Alice_backup_2026-01-01_000000"
    (folder / "sub").mkdir(parents=True)
    ro = folder / "sub" / "ro.txt"
    ro.write_text("x")
    os.chmod(ro, stat.S_IREAD)

    _store(tmp_path).delete_backup(folder.name)

    assert not folder.exists()


def test_delete_backup_missing_logs_not_raises(tmp_path):
    lines = []
    _store(tmp_path).delete_backup("Alice_backup_missing", on_log=lines.append)
    assert lines and all(m.startswith("ERROR: Deletion failed for ") for m in lines)


def test_delete_backup_other_error_propagates(tmp_path, monkeypatch):
    def boom(path, on_log=None):
        raise RuntimeError("boom")
    monkeypatch.setattr("ezbak.storage.local.rmtree_readonly", boom)
    with pytest.raises(RuntimeError):
        _store(tmp_path).delete_backup("x")


def test_free_space_delegates(tmp_path, monkeypatch):
    seen = []
    monkeypatch.setattr(ezbak.winsys, "free_space", lambda p: seen.append(p) or 12345)
    assert _store(tmp_path).free_space() == 12345
    assert seen == [str(tmp_path)]


def test_satisfies_backup_store_protocol(tmp_path):
    assert isinstance(_store(tmp_path), BackupStore)
