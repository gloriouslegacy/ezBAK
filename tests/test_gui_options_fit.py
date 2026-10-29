"""F1 (docs/spec/ui-fixes.md): Backup Options fit the default window (1000x700).

The System Exclude radio was cut off at the default width. The session gui_app is
withdrawn and never runs mainloop, so real geometry needs a child Python (same as
test_gui_window_shown.py) that builds the App, shows the backup page and measures.
"""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = [pytest.mark.gui, pytest.mark.windows,
              pytest.mark.skipif(not sys.platform.startswith("win"),
                                 reason="real window geometry is checked on Windows only")]

pytest.importorskip("customtkinter")

REPO_ROOT = Path(__file__).resolve().parent.parent

CHILD = """
import json, sys
from ezbak.gui import App, Win11Dialog
for name in ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel"):
    setattr(Win11Dialog, name, lambda *a, **k: None)
app = App(entry_script=sys.argv[1])
seen = {}
def walk(widget):
    yield widget
    for child in widget.winfo_children():
        yield from walk(child)
def record():
    app.show_page('backup')
    app.update()
    body = app.page_bodies['options']
    radios = [w for w in walk(body) if w.winfo_class() == 'Radiobutton'
              and str(w.cget('variable')) == str(app.system_mode_var)
              and w.cget('value') == 'exclude']
    # CTk.geometry("1000x700") scales by the window scaling (int(value * scaling))
    scale = app._get_window_scaling()
    seen.update(lang=app.translator.current_lang, size=[app.winfo_width(), app.winfo_height()],
                want_size=[int(1000 * scale), int(700 * scale)],
                radios=len(radios), card_right=body.winfo_rootx() + body.winfo_width(),
                radio_right=[r.winfo_rootx() + r.winfo_reqwidth() for r in radios],
                radio_width=[[r.winfo_width(), r.winfo_reqwidth()] for r in radios])
    app.destroy()
app.after(1500, record)
app.mainloop()
print("PROBE " + json.dumps(seen))
"""


@pytest.mark.parametrize("lang", ["en", "ko"])
def test_system_exclude_fits_options_card(tmp_path, lang):
    env = dict(os.environ)
    for var in ("APPDATA", "LOCALAPPDATA"):
        (tmp_path / var).mkdir()
        env[var] = str(tmp_path / var)
    cfg = tmp_path / "APPDATA" / "ezbak"
    cfg.mkdir()
    (cfg / "settings.json").write_text(json.dumps({"language": lang}), encoding="utf-8")
    proc = subprocess.run([sys.executable, "-c", CHILD, str(REPO_ROOT / "ezBAK.py")],
                          cwd=str(REPO_ROOT), env=env, capture_output=True, text=True,
                          timeout=120)
    assert proc.returncode == 0, proc.stderr
    probe = [line for line in proc.stdout.splitlines() if line.startswith("PROBE ")]
    assert len(probe) == 1, proc.stdout + proc.stderr
    seen = json.loads(probe[0][len("PROBE "):])
    assert seen["lang"] == lang
    assert seen["size"] == seen["want_size"], seen  # not shrunk by a small screen
    assert seen["radios"] == 1, seen
    width, reqwidth = seen["radio_width"][0]
    assert width >= reqwidth, seen  # pack did not cut the radio
    assert seen["radio_right"][0] <= seen["card_right"], seen
