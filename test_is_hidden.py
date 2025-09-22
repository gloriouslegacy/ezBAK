import importlib.util
import os
import subprocess
import tempfile
import sys

# prepare temp file
tmp_dir = tempfile.gettempdir()
path = os.path.join(tmp_dir, 'ezbak_test_hidden.txt')
with open(path, 'w') as f:
    f.write('test')

# set hidden and system attributes using Windows attrib
subprocess.run(f'attrib +h +s "{path}"', shell=True)

try:
    spec = importlib.util.spec_from_file_location('ezbakmod', r'c:/Users/asdf/Downloads/ezbak/dataDriverBAK/ezBAK.py')
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    app = mod.App()

    print('Temp file:', path)
    print('Default modes: hidden=%s, system=%s' % (app.hidden_mode_var.get(), app.system_mode_var.get()))
    print('is_hidden default:', app.is_hidden(path))

    app.hidden_mode_var.set('include')
    print('After hidden include -> is_hidden:', app.is_hidden(path))

    app.system_mode_var.set('include')
    print('After system include -> is_hidden:', app.is_hidden(path))

finally:
    # cleanup: remove attributes and file
    subprocess.run(f'attrib -h -s "{path}"', shell=True)
    try:
        os.remove(path)
    except Exception:
        pass

print('Done')
