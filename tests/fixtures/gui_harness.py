"""Widget snapshot helpers for the GUI goldens (P5-0b App, P5-0c dialogs).

A snapshot is plain JSON data, the same on every run and machine:
- no pixel geometry (winfo_width/x/...), no DPI or screen values;
- widget path names (!frame2) are not recorded: children are listed in creation order,
  so the order is the identity. Widgets held by an owner attribute get "attr";
- only options that differ from the widget default are kept (what the code sets);
- Tcl command names lose their id prefix ("2471...start_backup_thread" -> name),
  Tk variable names become the owner attribute name ("PY_VAR3" -> "hidden_mode_var"),
  [YYYY-mm-dd HH:MM:SS] clock texts become <TS>.

CustomTkinter widgets (U2-2, ADR 0004) are snapshotted through cget of CTK_OPTIONS
(their configure() returns nothing); the canvas, labels and frames a CTk widget keeps
in its own private attributes are drawing parts, not recorded (they depend on scaling and on
the drawing engine). Children the code adds to a CTk frame are recorded as usual.

Generic on purpose: snapshot_tree / widget_node / menu_tree / recording /
invoke_binding take any widget or object, so dialog goldens reuse them.

Loaded with importlib (tests/fixtures is not a package), like tool_harness.py.
"""
import contextlib
import copy
import importlib.util
import queue
import re
import tkinter as tk
from pathlib import Path


def _load_sibling(name, file):
    spec = importlib.util.spec_from_file_location(name, Path(__file__).with_name(file))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_ctk = _load_sibling("ezbak_test_gui_ctk", "gui_ctk.py")
is_ctk, ctk_parts, widget_text, is_button = _ctk.is_ctk, _ctk.ctk_parts, _ctk.widget_text, _ctk.is_button
is_ctk_class = _ctk.is_ctk_class

TS_RE = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")
CMD_ID_RE = re.compile(r"^\d+")
COMMAND_OPTIONS = ("command", "yscrollcommand", "xscrollcommand", "postcommand",
                   "validatecommand", "invalidcommand")
VARIABLE_OPTIONS = ("textvariable", "variable", "listvariable")
SKIP_CLASSES = ("Menu", "Toplevel")
# Widget state a running App changes (progress, busy buttons, user list): masked when
# comparing a live App with a fresh build.
LIVE_OPTIONS = ("state", "value", "maximum", "mode", "values")
# Options Tk fills from the Windows system settings when the code leaves them unset
# (the menu font name is in the system code page): not recorded.
PLATFORM_OPTIONS = {"Menu": ("font", "activeborderwidth")}


def plain(value):
    """Tcl objects, tuples and widgets -> JSON values (str, int, float, bool, None, list)."""
    if isinstance(value, (list, tuple)):
        return [plain(v) for v in value]
    if isinstance(value, dict):
        return {str(k): plain(v) for k, v in value.items()}
    if value is None or isinstance(value, (bool, int, float)):
        return value
    return TS_RE.sub("<TS>", str(value))


def command_name(value):
    """Tcl command name -> the Python function name tkinter registered it for."""
    return CMD_ID_RE.sub("", str(value))


class Namer:
    """Names for one owner object (App, dialog, host): its widget and variable attributes."""

    def __init__(self, *owners):
        self.widgets, self.variables = {}, {}
        for owner in owners:
            for name, value in vars(owner).items():
                if isinstance(value, tk.Variable):
                    self.variables.setdefault(str(value), name)
                elif isinstance(value, tk.Misc):
                    self.widgets.setdefault(str(value), name)

    def attr(self, widget):
        return self.widgets.get(str(widget))

    def variable(self, tcl_name):
        return self.variables.get(str(tcl_name), "<var>")


def changed_options(config, namer, ignore=()):
    """{option: value} for options whose value differs from the default.

    config: widget.configure() / menu.entryconfigure(i) / text.tag_configure(tag) output.
    ignore: option names left out (PLATFORM_OPTIONS of the widget class).
    """
    out = {}
    for name, spec in config.items():
        if not isinstance(spec, tuple) or len(spec) < 5 or name in ignore:
            continue  # alias such as bd -> borderwidth, or an ignored option
        default, value = spec[3], spec[4]
        if str(default) == str(value):
            continue
        if name in COMMAND_OPTIONS:
            out[name] = command_name(value)
        elif name in VARIABLE_OPTIONS:
            out[name] = namer.variable(value)
        else:
            out[name] = plain(value)
    return out


def layout(widget):
    """Geometry manager and its options ('in' dropped: the tree already says the parent)."""
    manager = widget.winfo_manager()
    info = {}
    if manager == "pack":
        info = widget.pack_info()
    elif manager == "grid":
        info = widget.grid_info()
    elif manager == "place":
        info = widget.place_info()
    info = {k: plain(v) for k, v in info.items() if k != "in"}
    return manager, info


def container(widget):
    """pack/grid propagate flags and grid column/row settings of a parent widget
    (a CTk widget's drawing parts do not count as slaves)."""
    out = {}
    parts = ctk_parts(widget)
    if [w for w in widget.pack_slaves() if str(w) not in parts]:
        out["pack_propagate"] = bool(widget.pack_propagate())
    if [w for w in widget.grid_slaves() if str(w) not in parts]:
        out["grid_propagate"] = bool(widget.grid_propagate())
        cols, rows = widget.grid_size()
        out["columns"] = [plain(widget.columnconfigure(i)) for i in range(cols)]
        out["rows"] = [plain(widget.rowconfigure(i)) for i in range(rows)]
    return out


def bindings(widget):
    """{sequence: command name} of the widget's own bindings (not class or all)."""
    out = {}
    bind = tk.Misc.bind  # CTk widgets override bind (it binds their canvas)
    for seq in bind(widget):
        match = re.search(r"\[(\S+) ", bind(widget, seq))
        out[seq] = command_name(match.group(1)) if match else plain(bind(widget, seq))
    return out


def canvas_items(canvas):
    items = []
    for item in canvas.find_all():
        items.append({"type": canvas.type(item), "coords": plain(canvas.coords(item)),
                      "fill": plain(canvas.itemcget(item, "fill")),
                      "outline": plain(canvas.itemcget(item, "outline"))})
    return items


def text_tags(text, namer):
    return {tag: changed_options(text.tag_configure(tag), namer)
            for tag in text.tag_names() if tag != "sel"}


def widget_node(widget, namer, skip=SKIP_CLASSES):
    node = {"class": widget.winfo_class(), "py": type(widget).__name__}
    attr = namer.attr(widget)
    if attr:
        node["attr"] = attr
    if is_ctk(widget):
        node["options"] = _ctk.ctk_options(widget, namer, plain)
    else:
        # the class's configure: a CTk child replaces configure on its tk master instance
        node["options"] = changed_options(type(widget).configure(widget), namer,
                                          PLATFORM_OPTIONS.get(node["class"], ()))
    manager, info = layout(widget)
    node["manager"] = manager
    if info:
        node["layout"] = info
    node.update(container(widget))
    binds = bindings(widget)
    if binds:
        node["bindings"] = binds
    if isinstance(widget, tk.Canvas) and not is_ctk(widget):
        node["items"] = canvas_items(widget)
    if isinstance(widget, tk.Text):
        node["tags"] = text_tags(widget, namer)
    children = snapshot_tree(widget, namer, skip)
    if children:
        node["children"] = children
    return node


def snapshot_tree(parent, namer, skip=SKIP_CLASSES):
    """Child nodes of parent in creation order; Menu and Toplevel children are skipped
    (menus are read through menu_tree, dialogs have their own snapshot), and so are the
    drawing parts of a CTk parent."""
    parts = ctk_parts(parent)
    return [widget_node(child, namer, skip) for child in parent.winfo_children()
            if child.winfo_class() not in skip and str(child) not in parts]


def menu_tree(menu, namer):
    """Options and entries of a menu; cascades are followed."""
    entries = []
    end = menu.index("end")
    for i in range(0 if end is None else end + 1):
        entry = {"type": menu.type(i)}
        if entry["type"] not in ("separator", "tearoff"):
            opts = changed_options(menu.entryconfigure(i), namer)
            sub = opts.pop("menu", None)
            entry.update(opts)
            if sub:
                entry["menu"] = menu_tree(menu.nametowidget(sub), namer)
        entries.append(entry)
    ignore = PLATFORM_OPTIONS["Menu"]
    return {"options": changed_options(menu.configure(), namer, ignore), "entries": entries}


def menubar(window, namer):
    name = str(window.cget("menu"))
    return menu_tree(window.nametowidget(name), namer) if name else None


def mask_live(nodes, texts_of=("status_label",)):
    """Copy of a tree without LIVE_OPTIONS and with the text of `texts_of` attrs masked."""
    out = []
    for node in copy.deepcopy(nodes):
        for key in LIVE_OPTIONS:
            node["options"].pop(key, None)
        if node.get("attr") in texts_of:
            node["options"]["text"] = "<LIVE>"
        if "children" in node:
            node["children"] = mask_live(node["children"], texts_of)
        out.append(node)
    return out


# --- calls made by bindings and menu entries ----------------------------------------

def own_members(cls):
    """{name: attribute} defined by cls and its bases outside tkinter and customtkinter
    (mixins included), dunder names left out; later classes in the MRO are overridden
    by earlier ones."""
    members = {}
    for klass in reversed(cls.__mro__):
        if klass is object or klass.__module__.split(".")[0] in ("tkinter", "customtkinter"):
            continue
        members.update({k: v for k, v in vars(klass).items()
                        if not (k.startswith("__") and k.endswith("__"))})
    return members


@contextlib.contextmanager
def recording(obj, prefix, calls, keep=()):
    """Within the block, obj's own methods (own_members of its class) are recorders on
    the instance appending [prefix.name, *args] to calls.

    Afterwards the instance dict is exactly as before. Not MonkeyPatch: its undo would
    store the class's bound methods on the instance, and the App's _underscore ones
    would then outlive gui_app's reset and shadow a later __class__ switch.
    """
    saved = dict(vars(obj))
    names = [name for name, value in own_members(type(obj)).items()
             if callable(value) and not isinstance(value, type) and name not in keep]
    try:
        for name in names:
            def recorder(*args, _name=f"{prefix}.{name}", **kwargs):
                calls.append([_name] + [plain(a) for a in args])
            setattr(obj, name, recorder)
        yield calls
    finally:
        for name in names:
            if name in saved:
                setattr(obj, name, saved[name])
            else:
                vars(obj).pop(name, None)


def own_binding_sequences(widget):
    """Sequences bound on widget itself, less those a CTk root binds to its own methods."""
    ctk_names = {name for klass in type(widget).__mro__ if is_ctk_class(klass)
                 for name in vars(klass)} - set(own_members(type(widget)))
    out = []
    for seq in tk.Misc.bind(widget):
        match = re.search(r"\[(\S+) ", tk.Misc.bind(widget, seq))
        if not (match and command_name(match.group(1)) in ctk_names):
            out.append(seq)
    return out


def invoke_binding(widget, seq):
    """Run the Tcl command bound to seq on widget as Tk would, with a dummy event."""
    match = re.search(r"\[(\S+) ", widget.bind(seq))
    fields = ["0"] * len(widget._subst_format)
    fields[widget._subst_format.index("%W")] = str(widget)
    return widget.tk.call(match.group(1), *fields)


def menu_commands(menu, path=()):
    """[(label path, menu, index)] of every command entry, cascades followed."""
    out = []
    end = menu.index("end")
    for i in range(0 if end is None else end + 1):
        kind = menu.type(i)
        if kind == "command":
            out.append(("/".join(path + (menu.entrycget(i, "label"),)), menu, i))
        elif kind == "cascade":
            sub = menu.nametowidget(menu.entrycget(i, "menu"))
            out += menu_commands(sub, path + (menu.entrycget(i, "label"),))
    return out


# --- a fresh build of the App window ------------------------------------------------

class _NoAfter:
    def after(self, ms, func=None, *args):
        """No timers on a host: the header clock would re-arm every second."""
        return "after#host"


def make_host(mod, app, theme_mode, lang):
    """Toplevel child of the session App carrying the App's own methods (all mixins), with
    a fresh theme, translator and copies of the App's Tk variables, filters and style.

    App.__init__ is not run (it would need a second Tk); the caller runs the steps it
    wants (create_widgets, update_ui_texts). Destroy the host afterwards.
    """
    members = own_members(mod.App)
    host_cls = type("SnapshotHost", (_NoAfter, tk.Toplevel), members)
    host = host_cls(app)
    host.withdraw()
    host.theme = mod.Win11Theme(initial_mode=theme_mode)
    host.translator = mod.Translator(initial_lang=lang)
    for name, value in list(vars(app).items()):
        if isinstance(value, tk.Variable):
            setattr(host, name, type(value)(master=host, value=value.get()))
    host.filters = copy.deepcopy(app.filters)
    host.style = app.style
    host.message_queue = queue.Queue()
    return host


# --- dialogs (P5-0c) ------------------------------------------------------------------

class DialogNamer(Namer):
    """Namer for a dialog: Tk variables no owner attribute holds (locals of the code that
    built the dialog) are named var1, var2, ... in the order the snapshot meets them.
    values: {variable name: value} of every variable met, read from Tcl."""

    def __init__(self, window, *owners):
        super().__init__(window, *owners)
        self.window = window
        self.values = {}
        self._unnamed = 0

    def variable(self, tcl_name):
        name = self.variables.get(str(tcl_name))
        if name is None:
            self._unnamed += 1
            name = self.variables[str(tcl_name)] = f"var{self._unnamed}"
        if name not in self.values:
            self.values[name] = plain(self.window.getvar(str(tcl_name)))
        return name


def dialog_snapshot(window, *owners):
    """Title, wm flags, the Toplevel node (options, bindings, children) and the values of
    the Tk variables its widgets use. No geometry; child Toplevels are left out."""
    namer = DialogNamer(window, *owners)
    node = widget_node(window, namer)
    return {"title": window.title(), "resizable": plain(window.resizable()),
            "transient": bool(str(window.transient())),
            "protocol": {p: command_name(window.protocol(p)) for p in window.protocol()},
            "node": node, "variables": namer.values}


def descendants(widget, cls):
    """Widgets of type cls under widget, depth first in creation order."""
    out = []
    for child in widget.winfo_children():
        if isinstance(child, cls):
            out.append(child)
        out += descendants(child, cls)
    return out


def toplevels(widget):
    return [w for w in widget.winfo_children() if isinstance(w, tk.Toplevel)]


def button(window, text):
    """The tk.Button or CTkButton under window whose text is `text` (exactly one)."""
    buttons = [w for w in descendants(window, tk.Misc) if is_button(w)]
    found = [b for b in buttons if b.cget("text") == text]
    assert len(found) == 1, (text, [b.cget("text") for b in buttons])
    return found[0]


def type_into(entry, text):
    entry.delete(0, "end")
    entry.insert(0, text)


@contextlib.contextmanager
def modal_hook(handler):
    """Within the block, tk.Misc.wait_window runs handler(window) instead of waiting, then
    destroys the window if it is still open (the user closing it); grab_set and
    grab_release do nothing (the dialogs' parent App is withdrawn, so a real grab fails).
    Yields the event list: ["wait", title], ["grab_set", title], ["grab_release", title].
    The class attributes are put back as they were (plain setattr, no MonkeyPatch)."""
    events = []
    names = ("wait_window", "grab_set", "grab_release")
    saved = {name: vars(tk.Misc)[name] for name in names}

    def wait_window(self, window=None):
        window = self if window is None else window
        events.append(["wait", window.title()])
        handler(window)
        if window.winfo_exists():
            window.destroy()

    def grab(name):
        def call(self):
            events.append([name, self.title() if isinstance(self, tk.Wm) else str(self)])
        return call

    try:
        tk.Misc.wait_window = wait_window
        tk.Misc.grab_set = grab("grab_set")
        tk.Misc.grab_release = grab("grab_release")
        yield events
    finally:
        for name, value in saved.items():
            setattr(tk.Misc, name, value)


@contextlib.contextmanager
def override(obj, **values):
    """Instance attributes set for the block; obj's instance dict is exactly as before
    afterwards (same rule as recording)."""
    saved = dict(vars(obj))
    try:
        for name, value in values.items():
            setattr(obj, name, value)
        yield obj
    finally:
        for name in values:
            if name in saved:
                setattr(obj, name, saved[name])
            else:
                vars(obj).pop(name, None)


def window_bindings(window):
    """{sequence: [calls]} of the window's own bindings, run with its methods (and the
    Tk destroy) recorded."""
    calls, out = [], {}
    with recording(window, "dlg", calls),             override(window, destroy=lambda: calls.append(["dlg.destroy"])):
        for seq in sorted(window.bind()):
            invoke_binding(window, seq)
            out[seq], calls[:] = list(calls), []
    return out


def path_normalizer(tmp_path):
    """value -> value with tmp_path (backslash or slash form) as <TMP>, lists/dicts walked."""
    subs = [(str(tmp_path), "<TMP>"), (str(tmp_path).replace("\\", "/"), "<TMP>")]

    def norm(value):
        if isinstance(value, str):
            for old, new in subs:
                value = value.replace(old, new)
            return value
        if isinstance(value, (list, tuple)):
            return [norm(v) for v in value]
        if isinstance(value, dict):
            return {k: norm(v) for k, v in value.items()}
        return value

    return norm


@contextlib.contextmanager
def app_unchanged(app):
    """Afterwards: Toplevels the block left are destroyed, and none may be left; the App's
    instance dict and filters equal what they were before."""
    before, filters = dict(vars(app)), copy.deepcopy(app.filters)
    children = list(app.winfo_children())
    try:
        yield
    finally:
        leftovers = [w for w in app.winfo_children() if w not in children]
        for widget in leftovers:
            widget.destroy()
    assert leftovers == []
    assert dict(vars(app)) == before
    assert app.filters == filters
