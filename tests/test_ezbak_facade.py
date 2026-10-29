"""Self-test of the ezBAK test facade (tests/fixtures/ezbak_facade.py, P5 D2) on
synthetic modules, so it does not depend on which gui modules exist yet."""
import importlib.util
import types
from pathlib import Path

import pytest

_PATH = Path(__file__).resolve().parent / "fixtures" / "ezbak_facade.py"
_spec = importlib.util.spec_from_file_location("ezbak_test_facade_selftest", _PATH)
facade_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(facade_mod)


def _module(name, **attrs):
    module = types.ModuleType(name)
    for key, value in attrs.items():
        setattr(module, key, value)
    return module


@pytest.fixture
def synthetic():
    """root + ezbak.gui (public name) + two gui submodules + fileattrs alias target."""
    shared = object()
    root = _module("ezBAK_synthetic", threading="root-threading", only_root=1, shared=shared)
    gui = _module("ezbak.gui", GuiThing="gui-thing", hidden="not public",
                  __all__=["GuiThing"])
    gui_a = _module("ezbak.gui.a", threading="a-threading", only_a="a")
    gui_b = _module("ezbak.gui.b", threading="b-threading", shared=shared)
    fileattrs = _module("ezbak.winsys.fileattrs", HAS_PYWIN32=True)
    packages = {"ezbak.gui": gui, "ezbak.gui.a": gui_a, "ezbak.gui.b": gui_b,
                "ezbak.winsys.fileattrs": fileattrs}
    facade = facade_mod.EzbakFacade(root, packages)
    return types.SimpleNamespace(facade=facade, root=root, gui=gui, a=gui_a, b=gui_b,
                                 fileattrs=fileattrs)


def test_read_order_root_then_gui_public_then_alias(synthetic):
    f = synthetic.facade
    assert f.threading == "root-threading"
    assert f.GuiThing == "gui-thing"
    assert f._have_pywin32 is True
    assert f.__name__ == "ezBAK_synthetic"


def test_unknown_and_non_public_names_raise_attribute_error(synthetic):
    f = synthetic.facade
    for name in ("no_such_name", "hidden", "only_a"):
        with pytest.raises(AttributeError):
            getattr(f, name)
    assert getattr(f, "no_such_name", "default") == "default"
    assert not hasattr(f, "hidden")


def test_setattr_spreads_to_root_and_gui_modules_that_have_the_name(synthetic):
    s = synthetic
    s.facade.threading = "patched"
    assert (s.root.threading, s.a.threading, s.b.threading) == ("patched",) * 3
    assert not hasattr(s.gui, "threading")


def test_new_name_lands_on_root_only(synthetic):
    s = synthetic
    s.facade.brand_new = 5
    assert s.root.brand_new == 5
    assert not any(hasattr(m, "brand_new") for m in (s.gui, s.a, s.b))


def test_monkeypatch_undo_restores_every_module(synthetic):
    s = synthetic
    original_shared = s.root.shared
    mp = pytest.MonkeyPatch()
    mp.setattr(s.facade, "threading", "patched")
    mp.setattr(s.facade, "shared", "patched-shared")
    mp.setattr(s.facade, "added", 1, raising=False)
    assert s.b.shared == "patched-shared"
    mp.undo()
    assert s.root.threading == "root-threading"
    # undo writes the value read through the facade (the root's) back everywhere
    assert s.a.threading == "root-threading" and s.b.threading == "root-threading"
    assert s.root.shared is original_shared and s.b.shared is original_shared
    assert not hasattr(s.root, "added")


def test_alias_forwards_set_and_undo_to_target(synthetic):
    s = synthetic
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(s.facade, "_have_pywin32", False)
        assert s.fileattrs.HAS_PYWIN32 is False
        assert s.facade._have_pywin32 is False
        assert not hasattr(s.root, "_have_pywin32")
    assert s.fileattrs.HAS_PYWIN32 is True


def test_delattr_removes_from_every_holder_and_unknown_raises(synthetic):
    s = synthetic
    del s.facade.threading
    assert not any(hasattr(m, "threading") for m in (s.root, s.a, s.b))
    with pytest.raises(AttributeError):
        del s.facade.no_such_name


def test_live_sys_modules_when_no_snapshot(monkeypatch):
    root = _module("ezBAK_live", x=1)
    gui_c = _module("ezbak.gui.c", x=1)
    monkeypatch.setitem(__import__("sys").modules, "ezbak.gui.c", gui_c)
    f = facade_mod.EzbakFacade(root)
    monkeypatch.setattr(f, "x", 9)
    assert (root.x, gui_c.x) == (9, 9)
    monkeypatch.undo()
    assert (root.x, gui_c.x) == (1, 1)


def test_module_alias_reads_the_module_and_spreads_to_gui_holders(synthetic, monkeypatch):
    """P5-16b: a name only gui modules still import (threading, filedialog, ...) is read
    through its alias (imported by name) and patched in every gui module holding it."""
    import string
    s = synthetic
    monkeypatch.setitem(facade_mod.ALIASES, "strmod", ("string", None))
    s.a.strmod = string
    assert s.facade.strmod is string
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(s.facade, "strmod", "patched")
        assert s.a.strmod == "patched"
        assert not hasattr(s.root, "strmod") and not hasattr(s.b, "strmod")
    assert s.a.strmod is string
    del s.facade.strmod
    assert not hasattr(s.a, "strmod")
    with pytest.raises(AttributeError):
        s.facade.strmod = "nowhere"


def test_attribute_alias_sets_target_and_gui_holders(synthetic, monkeypatch):
    s = synthetic
    monkeypatch.setitem(facade_mod.ALIASES, "attrs_free", ("ezbak.winsys.fileattrs", "FREE"))
    s.fileattrs.FREE = "real"
    s.b.attrs_free = "real"
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(s.facade, "attrs_free", "fake")
        assert (s.fileattrs.FREE, s.b.attrs_free) == ("fake", "fake")
        assert s.facade.attrs_free == "fake"
        assert not hasattr(s.root, "attrs_free") and not hasattr(s.a, "attrs_free")
    assert (s.fileattrs.FREE, s.b.attrs_free) == ("real", "real")
