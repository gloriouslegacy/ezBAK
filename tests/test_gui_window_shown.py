"""v0.10.0 fix: the App window is shown once mainloop runs on Windows (black screen bug).

CTk's withdraw() before the window exists sets a flag that makes CTk.mainloop() on
Windows hide the window for the titlebar colour and then skip deiconify(), so App must
hide itself with Tk's own withdraw. The session gui_app is withdrawn and never runs
mainloop (one Tk root per process), so a child Python builds the App and runs it.
"""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = [pytest.mark.gui, pytest.mark.windows,
              pytest.mark.skipif(not sys.platform.startswith("win"),
                                 reason="CTk's mainloop titlebar path runs on Windows only")]

pytest.importorskip("customtkinter")

REPO_ROOT = Path(__file__).resolve().parent.parent

CHILD = """
import json, sys
from ezbak.gui import App, Win11Dialog
for name in ("showwarning", "showerror", "showinfo", "askyesno", "askokcancel"):
    setattr(Win11Dialog, name, lambda *a, **k: None)
app = App(entry_script=sys.argv[1])
seen = {}
def record():
    seen.update(state=app.state(), mapped=app.winfo_ismapped(), viewable=app.winfo_viewable())
    app.destroy()
app.after(1500, record)
app.mainloop()
print("PROBE " + json.dumps(seen))
"""


def test_window_is_shown_after_mainloop(tmp_path):
    env = dict(os.environ)
    for var in ("APPDATA", "LOCALAPPDATA"):
        (tmp_path / var).mkdir()
        env[var] = str(tmp_path / var)
    proc = subprocess.run([sys.executable, "-c", CHILD, str(REPO_ROOT / "ezBAK.py")],
                          cwd=str(REPO_ROOT), env=env, capture_output=True, text=True,
                          timeout=120)
    assert proc.returncode == 0, proc.stderr
    probe = [line for line in proc.stdout.splitlines() if line.startswith("PROBE ")]
    assert len(probe) == 1, proc.stdout + proc.stderr
    seen = json.loads(probe[0][len("PROBE "):])
    assert seen == {"state": "normal", "mapped": 1, "viewable": 1}
