"""NAS connect / disconnect dialogs of App (moved from ezBAK.py in P5-15).

NasDialogsMixin is mixed into App before ctk.CTk; the methods use the same `self`
attributes as before. Module-level names (os, tk, resource_path, Win11Dialog,
storage_nas, winsys_sysinfo) match ezBAK.py so facade patches reach this module too.
net use command lines, target parsing and result messages are built in
ezbak.storage.nas; the password goes only into its argv, never into the queue, logs
or dialogs. The long methods are split into module functions that run the same
statements in the same order (P5 decision D5); blocks the two dialogs share
character for character (help label + form, buttons + centering + wait) are one
helper each. Module functions, not App methods, so gui_harness.recording never stubs
them. get_used_drive_letters uses winsys.sysinfo.used_drive_letters' own ctypes
lookup (the P4 lambda only existed so a test patch on the root ctypes reached it).
Colours from the design tokens for the theme's mode through _c (U6-5, ADR 0004).
"""
import logging
import os
import tkinter as tk
from types import SimpleNamespace

from ezbak.storage import nas as storage_nas
from ezbak.winsys import sysinfo as winsys_sysinfo
from ezbak.winsys.paths import resource_path

from .. import dpi
from ..theme import token_colors
from .message import Win11Dialog

# Diagnostic lines carry only the exception type (never exc_info): an exception from
# the connect path may hold the net use argv, which contains the password.
_log = logging.getLogger(__name__)


def _c(app, name):
    """Design token colour for the App's theme mode (U6-5). A module function, not an
    App method: the tests' recording replaces App methods."""
    return token_colors(app.theme)[name]


def _log_net_use_output(out, err, password):
    # g-j (P6-4): net use output reached only the job log, which is usually closed here.
    # The password is masked in case a message ever echoes it (it is not in the argv log).
    for stream, text in (("stdout", out), ("stderr", err)):
        if text:
            if password:
                text = text.replace(password, "***")
            _log.info("event=nas_connect_output stream=%s text=%s", stream, text.strip())


def _help_and_form(self, dlg, help_text, wrap):
    # Shared by both dialogs (identical lines): the command help label and the form frame
    tk.Label(dlg, text=help_text, bg=_c(self, 'card'), fg=_c(self, 'fg'), justify='left', anchor='w', font=("Consolas", 10), wraplength=dpi.px(dlg, wrap)).pack(fill='x', padx=12, pady=(12, 8))

    form = tk.Frame(dlg, bg=_c(self, 'card'))
    form.pack(fill='both', expand=True, padx=16)
    return form


def _finish_dialog(self, dlg, btns, on_ok, on_cancel):
    # Shared by both dialogs (identical lines): OK/Cancel, keys, centering, modal wait
    # Windows 11 style buttons
    ok_btn = Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_ok'), on_ok, self.theme, is_primary=True)
    ok_btn.pack(side='right', padx=8)
    cancel_btn = Win11Dialog._create_dialog_button(btns, self.translator.get('dialog_cancel'), on_cancel, self.theme, is_primary=False)
    cancel_btn.pack(side='right', padx=8)

    dlg.bind('<Return>', lambda e: on_ok())
    dlg.bind('<Escape>', lambda e: on_cancel())
    dlg.update_idletasks()
    try:
        # Center on parent
        px = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        py = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(px,0)}+{max(py,0)}")
    except Exception:
        pass  # centering is cosmetic; the dialog stays where Tk put it (as before)
    self.wait_window(dlg)


def _connect_fields(self, form):
    # Layout: the connect form rows (UNC, drive letter, credentials, persistent)
    row = 0
    def add_label(text):
        nonlocal row
        tk.Label(form, text=text, bg=_c(self, 'card'), fg=_c(self, 'fg')).grid(row=row, column=0, sticky='w', pady=6)
    def add_entry(var, width=42, show=None):
        nonlocal row
        e = tk.Entry(form, textvariable=var, width=width, show=show, bg=_c(self, 'bg'), fg=_c(self, 'fg'), insertbackground=_c(self, 'fg'))
        e.grid(row=row, column=1, sticky='w', pady=6)
        row += 1
        return e

    # UNC Path
    add_label(self.translator.get('unc_path_label'))
    unc_var = tk.StringVar()
    add_entry(unc_var)

    # Map drive option
    map_var = tk.BooleanVar(value=True)
    def_drive = self.find_free_drive_letter() or 'Z'
    drive_var = tk.StringVar(value=def_drive)

    map_frame = tk.Frame(form, bg=_c(self, 'card'))
    tk.Checkbutton(map_frame, text=self.translator.get('map_as_drive_letter'), variable=map_var, bg=_c(self, 'card'), fg=_c(self, 'fg'), selectcolor=_c(self, 'card'), activebackground=_c(self, 'card'), activeforeground=_c(self, 'fg')).pack(side='left')
    tk.Label(map_frame, text=self.translator.get('letter'), bg=_c(self, 'card'), fg=_c(self, 'fg')).pack(side='left', padx=(12,4))
    drive_entry = tk.Entry(map_frame, textvariable=drive_var, width=6, bg=_c(self, 'bg'), fg=_c(self, 'fg'), insertbackground=_c(self, 'fg'))
    drive_entry.pack(side='left')
    map_frame.grid(row=row, column=0, columnspan=2, sticky='w', pady=6)
    row += 1

    def toggle_drive():
        try:
            state = 'normal' if map_var.get() else 'disabled'
            drive_entry.config(state=state)
        except Exception:
            pass  # entry state is cosmetic; a destroyed entry is left alone (as before)
    toggle_drive()
    map_var.trace_add('write', lambda *_: toggle_drive())  # follow the checkbox (P5-fix h-j)
    try:
        map_frame.children
        # Bind after creation
        map_frame.bind('<Map>', lambda e: toggle_drive())
    except Exception:
        pass  # the <Map> re-check is cosmetic (as before)
    tk.Checkbutton(map_frame, variable=map_var).destroy()  # no-op cleanup (ensures lint calm)

    # Credentials
    add_label(self.translator.get('username_optional'))
    user_var = tk.StringVar()
    add_entry(user_var)

    add_label(self.translator.get('password_required_if_username'))
    pwd_var = tk.StringVar()
    add_entry(pwd_var, show='*')

    # Persistent
    persist_var = tk.BooleanVar(value=True)
    persist_chk = tk.Checkbutton(form, text=self.translator.get('reconnect_at_logon_persistent'), variable=persist_var, bg=_c(self, 'card'), fg=_c(self, 'fg'), selectcolor=_c(self, 'card'), activebackground=_c(self, 'card'), activeforeground=_c(self, 'fg'))
    persist_chk.grid(row=row, column=0, columnspan=2, sticky='w', pady=8)
    row += 1
    return SimpleNamespace(unc_var=unc_var, map_var=map_var, drive_var=drive_var,
                           user_var=user_var, pwd_var=pwd_var, persist_var=persist_var)


def _connect_on_ok(self, dlg, v, result):
    # Body of the connect dialog's on_ok: input checks, then the answer into result
    unc = (v.unc_var.get() or '').strip()
    if not unc or not unc.startswith('\\\\'):
        Win11Dialog.showerror(self.translator.get('error'), self.translator.get('unc_path_required'), parent=dlg, theme=self.theme, translator=self.translator)
        return
    user = (v.user_var.get() or '').strip()
    pwd = (v.pwd_var.get() or '')
    if user and not pwd:
        Win11Dialog.showerror(self.translator.get('error'), self.translator.get('password_required'), parent=dlg, theme=self.theme, translator=self.translator)
        return
    drv = None
    if v.map_var.get():
        d = (v.drive_var.get() or '').strip().rstrip(':').upper()
        if not (len(d) == 1 and 'A' <= d <= 'Z'):
            Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_drive_letter'), parent=dlg, theme=self.theme, translator=self.translator)
            return
        drv = d + ':'
    result.update(ok=True, unc=unc, drive=drv, user=user, pwd=pwd, persistent=v.persist_var.get())
    dlg.destroy()


def _run_connect(self, result):
    # Execute connection (after the dialog closed with OK)
    unc = result['unc']
    drive = result['drive']
    username = result['user']
    password = result['pwd']
    persistent = result['persistent']

    try:
        cmd = storage_nas.build_connect_cmd(storage_nas.NasSpec(unc, drive, username, password, persistent))

        self.message_queue.put(('log', f"Connecting to {unc}..."))
        res = storage_nas.run_net_use(cmd)

        out, err = res.out, res.err
        _log_net_use_output(out, err, password)
        if out:
            try:
                self.write_detailed_log(out)
            except Exception as exc:
                # the detailed log copy is best effort (as before)
                _log.debug("event=nas_connect_stdout_log failed error_type=%s", type(exc).__name__)
        if err:
            try:
                self.write_detailed_log(err)
            except Exception as exc:
                # the detailed log copy is best effort (as before)
                _log.debug("event=nas_connect_stderr_log failed error_type=%s", type(exc).__name__)

        if res.rc == 0:
            self.message_queue.put(('log', f"NAS connected: {drive or unc}"))
            try:
                if Win11Dialog.askyesno(self.translator.get('connected'), self.translator.get('open_share_now'), parent=self, theme=self.theme, translator=self.translator):
                    os.startfile(drive if drive else unc)
            except Exception as exc:
                # opening the share is optional; the connection is already logged (as before)
                _log.debug("event=nas_open_share failed error_type=%s", type(exc).__name__)
        else:
            msg = res.message
            self.message_queue.put(('log', f"NAS connect failed: {msg}"))
            try:
                Win11Dialog.showerror(self.translator.get('nas_connect'), self.translator.get('failed_to_connect').format(msg), parent=self, theme=self.theme, translator=self.translator)
            except Exception as exc:
                # the failure is already in the queue log (as before)
                _log.debug("event=nas_connect_failed_dialog failed error_type=%s", type(exc).__name__)
    except Exception as e:
        self.message_queue.put(('log', f"NAS connect error: {e}"))
        try:
            Win11Dialog.showerror(self.translator.get('error'), self.translator.get('nas_connect_error').format(e), parent=self, theme=self.theme, translator=self.translator)
        except Exception as exc:
            # the error is already in the queue log (as before)
            _log.debug("event=nas_connect_error_dialog failed error_type=%s", type(exc).__name__)
        return


def _run_disconnect(self, target):
    # Execute disconnection (after the dialog closed with OK)
    try:
        parsed = storage_nas.disconnect_target(target)
        if parsed is None:
            Win11Dialog.showerror(self.translator.get('error'), self.translator.get('invalid_drive_letter'), parent=self, theme=self.theme, translator=self.translator)
            return
        args, label = parsed

        self.message_queue.put(('log', f"Disconnecting {label}..."))
        res = storage_nas.run_net_use(args)
        if res.rc == 0:
            self.message_queue.put(('log', f"NAS disconnected: {label}"))
            try:
                Win11Dialog.showinfo(self.translator.get('nas_disconnect'), self.translator.get('disconnected').format(label), parent=self, theme=self.theme, translator=self.translator)
            except Exception as exc:
                # the result is already in the queue log (as before)
                _log.debug("event=nas_disconnected_dialog failed error_type=%s", type(exc).__name__)
        else:
            msg = res.message
            self.message_queue.put(('log', f"NAS disconnect failed: {msg}"))
            try:
                Win11Dialog.showerror(self.translator.get('nas_disconnect'), self.translator.get('failed_to_disconnect').format(msg), parent=self, theme=self.theme, translator=self.translator)
            except Exception as exc:
                # the failure is already in the queue log (as before)
                _log.debug("event=nas_disconnect_failed_dialog failed error_type=%s", type(exc).__name__)
    except Exception as e:
        self.message_queue.put(('log', f"NAS disconnect error: {e}"))
        try:
            Win11Dialog.showerror(self.translator.get('error'), self.translator.get('nas_disconnect_error').format(e), parent=self, theme=self.theme, translator=self.translator)
        except Exception as exc:
            # the error is already in the queue log (as before)
            _log.debug("event=nas_disconnect_error_dialog failed error_type=%s", type(exc).__name__)
        return


class NasDialogsMixin:
    """get_used_drive_letters, find_free_drive_letter and the NAS dialogs of App."""

    def get_used_drive_letters(self):
        """Returns a set of used drive letters (uppercase, without colon).

        Delegates to ezbak.winsys.sysinfo.used_drive_letters, which looks kernel32 up
        through its own ctypes at call time; self is unused (golden calls it with None)."""
        return winsys_sysinfo.used_drive_letters()

    def find_free_drive_letter(self):
        """Finds a free drive letter, preferring Z backward to D (avoid A,B, system)."""
        return storage_nas.free_drive_letter(self.get_used_drive_letters())

    def open_nas_connect_dialog(self):
        """Open a large dialog for NAS connect with full command examples, then run 'net use'."""
        # Build dialog UI
        dlg = tk.Toplevel(self)
        dlg.title(self.translator.get('network_share_connect'))
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # the window icon is cosmetic (as before)
        dlg.configure(bg=_c(self, 'card'))
        dlg.geometry(dpi.geometry(dlg, 680, 301))  # px at 100 % (U8)
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass  # modal grab is best effort (as before)

        wrap = 640
        help_text = (
            "Connect: net use [<drive_letter>:] \\Server\\share [/user:id pwd] /persistent:{yes|no}\n\n"
            "Disconnect: net use [<drive_letter>:|\\Server\\share] /delete /y"
        )
        form = _help_and_form(self, dlg, help_text, wrap)
        fields = _connect_fields(self, form)

        # Buttons
        btns = tk.Frame(dlg, bg=_c(self, 'card'))
        btns.pack(fill='x', pady=(6,12))

        result = {'ok': False}
        def on_ok():
            _connect_on_ok(self, dlg, fields, result)
        def on_cancel():
            dlg.destroy()

        _finish_dialog(self, dlg, btns, on_ok, on_cancel)

        if not result.get('ok'):
            self.message_queue.put(('log', "NAS connect cancelled."))
            return

        _run_connect(self, result)

    def open_nas_disconnect_dialog(self):
        """Open a large dialog for NAS disconnect with full command examples, then run 'net use ... /delete'."""
        # Build dialog UI
        dlg = tk.Toplevel(self)
        dlg.title(self.translator.get('network_share_disconnect'))
        try:
            dlg.iconbitmap(resource_path('./icon/ezbak.ico'))
        except Exception:
            pass  # the window icon is cosmetic (as before)
        dlg.configure(bg=_c(self, 'card'))
        dlg.geometry(dpi.geometry(dlg, 680, 178))  # px at 100 % (U8)
        dlg.transient(self)
        try:
            dlg.grab_set()
        except Exception:
            pass  # modal grab is best effort (as before)

        wrap = 640
        help_text = (
            "Connect: net use [<drive_letter>:] \\\\server\\share [/user:username password] /persistent:{yes|no}\n\n"
            "Disconnect: net use [drive_letter:|\\\\server\\share] /delete /y"
        )
        form = _help_and_form(self, dlg, help_text, wrap)

        tk.Label(form, text=self.translator.get('target_drive_or_unc'), bg=_c(self, 'card'), fg=_c(self, 'fg')).grid(row=0, column=0, sticky='w', pady=6)
        target_var = tk.StringVar()
        tk.Entry(form, textvariable=target_var, width=42, bg=_c(self, 'bg'), fg=_c(self, 'fg'), insertbackground=_c(self, 'fg')).grid(row=0, column=1, sticky='w', pady=6)

        btns = tk.Frame(dlg, bg=_c(self, 'card'))
        btns.pack(fill='x', pady=(6,12))

        result = {'ok': False}
        def on_ok():
            val = (target_var.get() or '').strip()
            if not val:
                Win11Dialog.showerror(self.translator.get('error'), self.translator.get('input_target'), parent=dlg, theme=self.theme, translator=self.translator)
                return
            result.update(ok=True, target=val)
            dlg.destroy()
        def on_cancel():
            dlg.destroy()

        _finish_dialog(self, dlg, btns, on_ok, on_cancel)

        if not result.get('ok'):
            self.message_queue.put(('log', "NAS disconnect cancelled."))
            return

        _run_disconnect(self, result['target'])
