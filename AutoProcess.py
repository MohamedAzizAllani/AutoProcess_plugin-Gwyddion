"""
AutoProcess – Gwyddion Python 2.7 plugin

Purpose
-------
Batch toolbox for SPM data in Gwyddion:
- Color range control (fixed/full, invert, zero-to-min).
- Palette application from known gradients.
- Cropping (by selection or coordinates), optional new image, offset handling.
- Channel renaming.
- Macro replay from parsed processing logs.
- Batch save of selected channels to .gwy while ensuring logs/ranges exist.
- File/channel selection helpers and live browser sync.

"""
# ---------- Plugin metadata required by Gwyddion ----------

plugin_menu = "/AutoProcess"       # Where the plugin appears in Gwyddion's menu
plugin_desc = ("Replays processing tools, sets color palettes, applies fixed "
               "color ranges, inverts mapping, sets zero to minimum, "
               "deletes/removes files, and automates cropping.")
plugin_type = "PROCESS"            # Gwyddion plugin type

# -----------------------------
# Standard & Gwyddion Imports
# -----------------------------
import os
import re
import sys
import time
import gtk               # GTK for GUI
import gtk.gdk           # GTK GDK (pixbuf, colors)
import pango             # Font attributes (bold, weights)
import gobject           # GObject & signals
import logging           # Logging and debug
import tempfile          # Temp directory for logs
from datetime import datetime

import gwy               # Gwyddion main Python module (data browser, processing)

# -----------------------------
# Logging Setup
# -----------------------------
log_dir = tempfile.gettempdir()
log_file = os.path.join(log_dir, "SPM_autoprocess.log")

logger = logging.getLogger('SPM_autoprocess')
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter("%(asctime)s,%(msecs)03d: %(message)s",
                              datefmt='%Y-%m-%d %H:%M:%S')

try:
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    file_handler = logging.FileHandler(log_file, mode='w')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.debug("Logger initialized with file handler: %s", log_file)
except Exception as e:
    # Fallback to console if file handler cannot be created
    logger.debug("Failed to initialize file handler for %s: %s", log_file, str(e))
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.debug("Using console handler due to file handler failure")


class StderrToLogger(object):
    """Redirect sys.stderr to the plugin logger as warnings.
    Keeps Gwyddion Python console tidy and captures tracebacks in the log.
    """
    def __init__(self, _logger):
        self.logger = _logger

    def write(self, message):
        if message.strip():
            self.logger.warning(message.strip())

    def flush(self):
        # File-like API compliance
        pass


# Redirect stderr to our logger
sys.stderr = StderrToLogger(logger)

# -----------------------------
# Frequently Used Container Keys
# -----------------------------
DATA_KEY          = "/%d/data"
BASE_MIN_KEY      = "/%d/base/min"
BASE_MAX_KEY      = "/%d/base/max"
RANGE_TYPE_KEY    = "/%d/base/range-type"
VISIBLE_KEY       = "/%d/data/visible"
SELECTION_KEYS    = ["/%d/select/rectangle", "/%d/data/selection"]
FILENAME_KEY      = "/filename"
TITLE_KEY         = "/%d/data/title"
ORIGINAL_MIN_KEY  = "/%d/base/original_min"
ORIGINAL_MAX_KEY  = "/%d/base/original_max"

# -----------------------------
# State Holder to Avoid Globals
# -----------------------------
class PluginState(object):
    """Hold all volatile UI references and runtime state.

    Keeping everything together helps avoid accidental globals and makes
    teardown/cleanup simpler.
    """
    def __init__(self):
        # Macro (parsed entries from a log file)
        self.macro = []

        # GTK data models
        self.liststore = gtk.ListStore(int, str, str)
        self.channel_liststore = gtk.ListStore(bool, str, bool, object, int, str,
                                               gtk.gdk.Pixbuf, gtk.gdk.Pixbuf)

        # Top-level window and major widgets
        self.window = None
        self.palette_combobox = None
        self.min_entry = None
        self.max_entry = None
        self.x_entry = None
        self.y_entry = None
        self.width_entry = None
        self.height_entry = None
        self.create_new_check = None
        self.keep_offsets_check = None
        self.rename_entry = None
        self.select_all_check = None
        self.select_dropdown = None
        self.select_store = None

        # Runtime bookkeeping
        self.selection_connections = []
        self.timeout_id = None
        self.data_browser_timeout_id = None
        self.current_container = None
        self.current_data_id = None
        self.last_containers = None


# Keep a single open GUI instance (avoid duplicates)
_plugin_gui_instance = None

# --------------------------------
# Log Parsing Utilities
# --------------------------------
def parse_log_entry(entry):
    """Parse a single 'proc::func(params)@timestamp' line into a dict.

    Returns:
        dict or None
        Example: {'function': 'func', 'parameters': {...}, 'param_string': '...',
                  'timestamp': '...', 'order': int (added by file parser)}
    """
    try:
        match = re.match(r"proc::(\w+)\((.*?)\)@(.+?)(?:Z|$)", entry)
        if not match:
            logger.debug("Skipping non-proc log entry")
            return None

        function, params, tstamp = match.groups()
        param_string = params.strip()
        param_dict = {}

        if param_string:
            # Split on commas not inside quotes
            parts = re.split(r",\s*(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", param_string)
            for param in parts:
                if '=' in param:
                    key, value = param.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    try:
                        if value.lower() == 'true':
                            param_dict[key] = True
                        elif value.lower() == 'false':
                            param_dict[key] = False
                        elif value.replace('.', '', 1).isdigit():
                            param_dict[key] = (float(value) if '.' in value else int(value))
                        else:
                            param_dict[key] = value.strip('\"')
                    except Exception:
                        param_dict[key] = value

        return {
            "function": function,
            "parameters": param_dict,
            "param_string": param_string,
            "timestamp": tstamp.strip()
        }
    except Exception:
        return None


def parse_log_file(file_path):
    """Parse a log file and return a list of parsed entries (with 'order')."""
    log_entries = []
    try:
        with open(file_path, "r") as f:
            for i, line in enumerate(f):
                parsed = parse_log_entry(line.strip())
                if parsed:
                    parsed["order"] = i + 1
                    log_entries.append(parsed)
        logger.info("Parsed %d proc entries from %s", len(log_entries), file_path)
    except IOError:
        logger.error("Error reading log file %s", file_path)
    return log_entries


def update_macro_view(liststore, macro):
    """Refresh macro table (order, function, parameter string)."""
    liststore.clear()
    for i, entry in enumerate(macro):
        liststore.append([i + 1, entry["function"], entry["param_string"]])


def load_log_file(button, entry, liststore, macro):
    """Open a file chooser, parse selected log, update macro table/model."""
    dialog = gtk.FileChooserDialog("Select Log File", None,
                                   gtk.FILE_CHOOSER_ACTION_OPEN,
                                   (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                    gtk.STOCK_OPEN, gtk.RESPONSE_OK))
    dialog.set_default_response(gtk.RESPONSE_OK)
    response = dialog.run()
    if response == gtk.RESPONSE_OK:
        file_path = dialog.get_filename()
        entry.set_text(file_path)
        if file_path and os.path.exists(file_path):
            macro[:] = parse_log_file(file_path)
            update_macro_view(liststore, macro)
            if not macro:
                logger.warning("No valid proc entries in %s", file_path)
                show_message_dialog(gtk.MESSAGE_WARNING,
                                    "No valid processing tools found in the log file.")
        else:
            logger.error("Log file does not exist: %s", file_path)
            show_message_dialog(gtk.MESSAGE_ERROR,
                                "Log file does not exist: %s" % file_path)
    dialog.destroy()


# --------------------------------
# Generic GTK Dialog Helpers
# --------------------------------
def show_message_dialog(msg_type, message, parent=None):
    """Simple OK dialog for info/warning/error messages."""
    dialog = gtk.MessageDialog(parent=parent, flags=0, type=msg_type,
                               buttons=gtk.BUTTONS_OK, message_format=message)
    dialog.run()
    dialog.destroy()


def show_rename_confirmation_dialog(new_names, parent):
    """Confirm the list of (old → new) channel renames before applying.

    Args:
        new_names: list of tuples (old_title, new_title, container, data_id, filename)
    """
    message = "The following channels will be renamed:\n\n"
    for old_name, new_name, _, _, _ in new_names:
        message += "%s -> %s\n" % (old_name, new_name)

    dialog = gtk.MessageDialog(parent=parent, flags=gtk.DIALOG_MODAL,
                               type=gtk.MESSAGE_QUESTION,
                               buttons=gtk.BUTTONS_OK_CANCEL,
                               message_format=message)
    dialog.set_title("Confirm Rename")
    response = dialog.run()
    dialog.destroy()

    ok = (response == gtk.RESPONSE_OK)
    logger.info("User %s rename operation", "confirmed" if ok else "cancelled")
    return ok


# --------------------------------
# Rename Selected Channels
# --------------------------------
def apply_rename(button, channel_liststore, state):
    """Rename selected channels to the same base name (exact replacement)."""
    base_name = state.rename_entry.get_text().strip()
    if not base_name:
        logger.error("No base name provided for renaming")
        show_message_dialog(gtk.MESSAGE_ERROR, "Please enter a valid base name.")
        return

    selected = []
    for row in channel_liststore:
        checked, title, is_channel, container, data_id, filename, _, _ = row
        if checked and container and is_channel and data_id != -1:
            selected.append((container, data_id, title, filename))

    if not selected:
        logger.error("No channels selected for renaming")
        show_message_dialog(gtk.MESSAGE_ERROR, "No channels selected for renaming")
        return

    # Prepare new names
    new_names = []
    for container, data_id, title, filename in selected:
        new_name = base_name
        new_names.append((title, new_name, container, data_id, filename))

    if not show_rename_confirmation_dialog(new_names, state.window):
        logger.info("Rename operation cancelled by user")
        return

    def operation(container, data_id, title, filename):
        new_name = next(n for t, n, c, d, f in new_names
                        if c == container and d == data_id)
        container.set_string_by_name(TITLE_KEY % data_id, new_name)
        logger.info("Renamed data_id %d from %s to %s in %s",
                    data_id, title, new_name, filename)

    process_selected_channels(channel_liststore, operation,
                              "No valid channels to rename",
                              "Renamed %d channels", state)
    populate_data_channels(channel_liststore, state)


# --------------------------------
# One-Instance Window Signal
# --------------------------------
_gui_close_signal = None
if not gobject.signal_lookup("close-gui", gtk.Window):
    _gui_close_signal = gobject.signal_new("close-gui", gtk.Window,
                                           gobject.SIGNAL_RUN_FIRST,
                                           gobject.TYPE_NONE, ())

# --------------------------------
# GUI Construction
# --------------------------------
def create_gui(state):
    """Instantiate the main plugin window and build the complete UI."""
    # If an instance already exists (defensive), present and return.
    existing = _find_autoprocess_window()
    if existing is not None:
        try:
            existing.present()
        except Exception:
            pass
        return

    # Window
    state.window = gtk.Window()
    state.window.set_title("AutoProcess")
    state.window.set_resizable(True)
    state.window.set_size_request(680, 600)
    state.window.connect("delete-event",
                         lambda w, e: on_window_delete_event(w, e, state))
    logger.debug("Created main window")
    state.window.set_data("autoprocess_singleton", True)

    # Root layout
    main_hbox = gtk.HBox(spacing=5)
    state.window.add(main_hbox)

    left_vbox = gtk.VBox(spacing=5)
    right_vbox = gtk.VBox(spacing=5)
    right_vbox.set_size_request(305, -1)  # Fixed right pane width
    main_hbox.pack_start(left_vbox, True, True, 2)
    main_hbox.pack_start(right_vbox, False, False, 2)

    # ---------------- Color Range Block ----------------
    color_range_label = gtk.Label("<b>Fixed Color Range</b>")
    color_range_label.set_use_markup(True)
    color_range_label.set_alignment(0, 0.5)
    left_vbox.pack_start(color_range_label, False, False, 2)

    color_range_vbox = gtk.VBox(spacing=5)
    left_vbox.pack_start(color_range_vbox, False, False, 0)

    # Start (min) row
    hbox_min = gtk.HBox(spacing=5)
    color_range_vbox.pack_start(hbox_min, True, True, 0)

    label_min = gtk.Label("Start:")
    label_min.set_alignment(0, 0.5)
    hbox_min.pack_start(label_min, False, False, 0)

    state.min_entry = gtk.Entry()
    hbox_min.pack_start(state.min_entry, True, True, 0)

    # End (max) row
    hbox_max = gtk.HBox(spacing=5)
    color_range_vbox.pack_start(hbox_max, True, True, 0)

    label_max = gtk.Label("End: ")
    label_max.set_alignment(0, 0.5)
    hbox_max.pack_start(label_max, False, False, 0)

    state.max_entry = gtk.Entry()
    hbox_max.pack_start(state.max_entry, True, True, 0)

    # Buttons row 1
    hbox_min_buttons = gtk.HBox(spacing=5)
    color_range_vbox.pack_start(hbox_min_buttons, True, True, 0)

    apply_range_button = gtk.Button("Apply Fixed Range")
    hbox_min_buttons.pack_start(apply_range_button, True, True, 0)
    apply_range_button.connect("clicked",
                               lambda b: apply_fixed_color_range(b, state.channel_liststore, state))

    invert_button = gtk.Button("Invert Mapping")
    hbox_min_buttons.pack_start(invert_button, True, True, 0)
    invert_button.connect("clicked",
                          lambda b: invert_mapping(b, state.channel_liststore, state))

    # Buttons row 2
    hbox_max_buttons = gtk.HBox(spacing=5)
    color_range_vbox.pack_start(hbox_max_buttons, True, True, 0)

    full_range_button = gtk.Button("Set Full Range")
    hbox_max_buttons.pack_start(full_range_button, True, True, 0)
    full_range_button.connect("clicked",
                              lambda b: set_to_full_range(b, state.channel_liststore, state))

    zero_min_button = gtk.Button("Zero to Min")
    hbox_max_buttons.pack_start(zero_min_button, True, True, 0)
    zero_min_button.connect("clicked",
                            lambda b: set_zero_to_minimum(b, state.channel_liststore, state))

    separator1 = gtk.HSeparator()
    left_vbox.pack_start(separator1, False, False, 5)

    # ---------------- Color + Rename Block ----------------
    hbox_color_rename = gtk.HBox(spacing=7)
    left_vbox.pack_start(hbox_color_rename, False, False, 2)

    # Color palette sub-block
    vbox_color = gtk.VBox(spacing=5)
    hbox_color_rename.pack_start(vbox_color, True, True, 5)

    change_color_label = gtk.Label("<b>Change Color</b>")
    change_color_label.set_use_markup(True)
    change_color_label.set_alignment(0, 0.5)
    vbox_color.pack_start(change_color_label, False, False, 2)

    hbox_palette = gtk.HBox(spacing=5)
    vbox_color.pack_start(hbox_palette, False, False, 0)

    palette_store = gtk.ListStore(str, gtk.gdk.Pixbuf)
    for name, pixbuf in get_gradient_names():
        palette_store.append([name, pixbuf])

    state.palette_combobox = gtk.ComboBox(palette_store)
    state.palette_combobox.set_size_request(-1, -1)
    renderer_text = gtk.CellRendererText()
    state.palette_combobox.pack_start(renderer_text, True)
    state.palette_combobox.add_attribute(renderer_text, "text", 0)
    renderer_pixbuf = gtk.CellRendererPixbuf()
    state.palette_combobox.pack_start(renderer_pixbuf, False)
    state.palette_combobox.add_attribute(renderer_pixbuf, "pixbuf", 1)
    state.palette_combobox.set_active(0)
    hbox_palette.pack_start(state.palette_combobox, True, True, 0)

    hbox_apply_palette = gtk.HBox(spacing=5)
    vbox_color.pack_start(hbox_apply_palette, False, False, 0)

    apply_palette_button = gtk.Button("Apply Color Gradient")
    apply_palette_button.set_size_request(-1, -1)
    apply_palette_button.connect("clicked",
                                 lambda b: apply_palette(b, state.channel_liststore, state))
    hbox_apply_palette.pack_start(apply_palette_button, True, True, 0)

    # Vertical separator
    vertical_separator = gtk.VSeparator()
    hbox_color_rename.pack_start(vertical_separator, False, False, 5)

    # Rename sub-block
    vbox_rename = gtk.VBox(spacing=5)
    hbox_color_rename.pack_start(vbox_rename, True, True, 5)

    rename_files_label = gtk.Label("<b>Rename Files</b>")
    rename_files_label.set_use_markup(True)
    rename_files_label.set_alignment(0, 0.5)
    vbox_rename.pack_start(rename_files_label, False, False, 2)

    hbox_rename_entry = gtk.HBox(spacing=5)
    vbox_rename.pack_start(hbox_rename_entry, False, False, 0)

    state.rename_entry = gtk.Entry()
    rename_placeholder_text = "Insert New name..."

    # Placeholder behavior
    state.rename_entry.set_text(rename_placeholder_text)
    state.rename_entry.modify_text(gtk.STATE_NORMAL, gtk.gdk.color_parse("gray"))

    def on_rename_entry_focus_in(entry, event):
        if entry.get_text() == rename_placeholder_text:
            entry.set_text("")
            entry.modify_text(gtk.STATE_NORMAL, None)

    def on_rename_entry_focus_out(entry, event):
        if entry.get_text().strip() == "":
            entry.set_text(rename_placeholder_text)
            entry.modify_text(gtk.STATE_NORMAL, gtk.gdk.color_parse("gray"))

    state.rename_entry.connect("focus-in-event", on_rename_entry_focus_in)
    state.rename_entry.connect("focus-out-event", on_rename_entry_focus_out)
    hbox_rename_entry.pack_start(state.rename_entry, True, True, 0)

    hbox_apply_rename = gtk.HBox(spacing=5)
    vbox_rename.pack_start(hbox_apply_rename, False, False, 0)

    apply_rename_button = gtk.Button("Apply")
    apply_rename_button.set_size_request(-1, -1)
    apply_rename_button.connect("clicked",
                                lambda b: apply_rename(b, state.channel_liststore, state))
    hbox_apply_rename.pack_start(apply_rename_button, True, True, 0)

    separator2 = gtk.HSeparator()
    left_vbox.pack_start(separator2, False, False, 2)

    # ---------------- Crop Data Block ----------------
    crop_data_expander = gtk.Expander()
    crop_data_expander.set_use_markup(True)
    crop_data_expander.set_label("<b>Crop Data</b>")
    crop_data_expander.set_expanded(True)
    left_vbox.pack_start(crop_data_expander, False, False, 0)

    crop_data_vbox = gtk.VBox(spacing=2)
    crop_data_expander.add(crop_data_vbox)

    # Row 1: origin
    hbox_crop1 = gtk.HBox(spacing=5)
    crop_data_vbox.pack_start(hbox_crop1, False, False, 0)

    label_x = gtk.Label("Origin X (px):")
    label_x.set_alignment(0, 2)
    label_x.set_size_request(100, -1)
    hbox_crop1.pack_start(label_x, False, False, 0)

    state.x_entry = gtk.Entry()
    state.x_entry.set_text("0")
    hbox_crop1.pack_start(state.x_entry, True, True, 5)

    label_y = gtk.Label("Origin Y (px):")
    label_y.set_alignment(0, 0.5)
    label_y.set_size_request(100, -1)
    hbox_crop1.pack_start(label_y, False, False, 5)

    state.y_entry = gtk.Entry()
    state.y_entry.set_text("0")
    hbox_crop1.pack_start(state.y_entry, True, True, 5)

    # Row 2: size
    hbox_crop2 = gtk.HBox(spacing=5)
    crop_data_vbox.pack_start(hbox_crop2, False, False, 0)

    label_width = gtk.Label("Width (px):")
    label_width.set_alignment(0, 0.5)
    label_width.set_size_request(100, -1)
    hbox_crop2.pack_start(label_width, False, False, 0)

    state.width_entry = gtk.Entry()
    state.width_entry.set_text("100")
    hbox_crop2.pack_start(state.width_entry, True, True, 5)

    label_height = gtk.Label("Height (px):")
    label_height.set_alignment(0, 0.5)
    label_height.set_size_request(100, -1)
    hbox_crop2.pack_start(label_height, False, False, 5)

    state.height_entry = gtk.Entry()
    state.height_entry.set_text("100")
    hbox_crop2.pack_start(state.height_entry, True, True, 5)

    # Row 3: options + apply
    hbox_crop3 = gtk.HBox(spacing=5)
    crop_data_vbox.pack_start(hbox_crop3, False, False, 5)

    state.create_new_check = gtk.CheckButton("Create new image")
    state.create_new_check.set_active(False)
    hbox_crop3.pack_start(state.create_new_check, False, False, 5)

    state.keep_offsets_check = gtk.CheckButton("Keep lateral offsets")
    state.keep_offsets_check.set_active(False)
    hbox_crop3.pack_start(state.keep_offsets_check, False, False, 5)

    apply_crop_button = gtk.Button("    Apply crop    ")
    apply_crop_button.connect("clicked",
                              lambda b: apply_crop(b, state.channel_liststore, state))
    hbox_crop3.pack_start(apply_crop_button, False, False, 5)

    separator3 = gtk.HSeparator()
    left_vbox.pack_start(separator3, False, False, 5)

    # ---------------- Data Process Block ----------------
    data_process_expander = gtk.Expander()
    data_process_expander.set_use_markup(True)
    data_process_expander.set_label("<b>Data Process Functionalities</b>")
    data_process_expander.set_expanded(True)
    left_vbox.pack_start(data_process_expander, True, True, 0)

    data_process_vbox = gtk.VBox(spacing=5)
    data_process_expander.add(data_process_vbox)

    spacer = gtk.Label("")
    spacer.set_size_request(-1, 5)
    data_process_vbox.pack_start(spacer, False, False, 0)

    # Log file chooser row
    hbox_log = gtk.HBox(spacing=5)
    log_entry = gtk.Entry()
    placeholder_text = "Insert the Log file path..."
    log_entry.set_text(placeholder_text)
    log_entry.modify_text(gtk.STATE_NORMAL, gtk.gdk.color_parse("gray"))

    def on_log_entry_focus_in(entry, event):
        if entry.get_text() == placeholder_text:
            entry.set_text("")
            entry.modify_text(gtk.STATE_NORMAL, None)

    def on_log_entry_focus_out(entry, event):
        if entry.get_text().strip() == "":
            entry.set_text(placeholder_text)
            entry.modify_text(gtk.STATE_NORMAL, gtk.gdk.color_parse("gray"))

    log_entry.connect("focus-in-event", on_log_entry_focus_in)
    log_entry.connect("focus-out-event", on_log_entry_focus_out)
    hbox_log.pack_start(log_entry, True, True, 5)

    load_button = gtk.Button("Load Log File")
    load_button.connect("clicked",
                        lambda b: load_log_file(b, log_entry, state.liststore, state.macro))
    hbox_log.pack_start(load_button, False, False, 1)
    data_process_vbox.pack_start(hbox_log, False, False, 0)

    # Macro table
    scrolled_macro = gtk.ScrolledWindow()
    scrolled_macro.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)

    treeview_macro = gtk.TreeView(state.liststore)
    renderer_text = gtk.CellRendererText()
    treeview_macro.append_column(gtk.TreeViewColumn("#", renderer_text, text=0))
    treeview_macro.append_column(gtk.TreeViewColumn("Function", renderer_text, text=1))
    treeview_macro.append_column(gtk.TreeViewColumn("Parameters", renderer_text, text=2))
    scrolled_macro.add(treeview_macro)

    data_process_vbox.pack_start(scrolled_macro, True, True, 0)

    # Replay macro button
    replay_button = gtk.Button("Replay Selected Channels")
    replay_button.connect("clicked",
                          lambda b: replay_selected_channels(b, state.channel_liststore, state))
    data_process_vbox.pack_start(replay_button, False, False, 0)

    # ---------------- Right Pane: Files & Save ----------------
    hbox_spm = gtk.HBox(spacing=7)
    right_vbox.pack_start(hbox_spm, False, False, 2)

    vbox_spm = gtk.VBox(spacing=5)
    hbox_spm.pack_start(vbox_spm, True, True, 2)

    OpenSPM_Files_label = gtk.Label()
    OpenSPM_Files_label.set_markup("<b>List of Open SPM Files</b>")
    OpenSPM_Files_label.set_alignment(0, 0.5)
    vbox_spm.pack_start(OpenSPM_Files_label, False, False, 2)

    # Row: select-all + per-index selection + save .gwy
    hbox_select = gtk.HBox(spacing=5)
    vbox_spm.pack_start(hbox_select, False, False, 0)

    state.select_all_check = gtk.CheckButton("Select All")
    state.select_all_check.set_active(False)
    state.select_all_check.connect("toggled", sync_select_all_check,
                                   state.channel_liststore, state)
    hbox_select.pack_start(state.select_all_check, False, False, 5)

    temp_store = gtk.ListStore(str, bool, str)
    state.select_dropdown = gtk.ComboBox(temp_store)
    state.select_dropdown.set_size_request(350, 25)
    renderer_text = gtk.CellRendererText()
    state.select_dropdown.pack_start(renderer_text, True)
    state.select_dropdown.add_attribute(renderer_text, "text", 0)
    state.select_dropdown.set_active(0)
    state.select_dropdown.connect("changed", select_dropdown_changed,
                                  state.channel_liststore, state)
    hbox_select.pack_start(state.select_dropdown, True, True, 5)

    save_gwy_button = gtk.Button("Save As .GWY")
    save_gwy_button.set_size_request(-1, 25)
    save_gwy_button.connect("clicked",
                            lambda b: save_as_gwy(b, state.channel_liststore, state))
    hbox_select.pack_start(save_gwy_button, False, False, 0)

    # File + channel table
    scrolled_channels = gtk.ScrolledWindow()
    scrolled_channels.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)

    treeview_channels = gtk.TreeView(state.channel_liststore)
    renderer_toggle = gtk.CellRendererToggle()
    renderer_toggle.set_property("activatable", True)
    renderer_toggle.connect("toggled", toggle_channel_selection, state.channel_liststore)

    renderer_text_select = gtk.CellRendererText()
    column_toggle = gtk.TreeViewColumn("Select")
    column_toggle.pack_start(renderer_toggle, False)
    column_toggle.pack_start(renderer_text_select, False)
    column_toggle.set_cell_data_func(renderer_toggle, render_channel_column, treeview_channels)
    column_toggle.set_cell_data_func(renderer_text_select, render_channel_column, treeview_channels)
    treeview_channels.append_column(column_toggle)

    renderer_text = gtk.CellRendererText()
    renderer_text.set_property("markup", True)
    column_files = gtk.TreeViewColumn("SPM File / Channel", renderer_text, markup=1)
    column_files.set_alignment(0.5)
    treeview_channels.append_column(column_files)

    renderer_delete = gtk.CellRendererText()
    renderer_delete.set_property("xalign", 0.0)
    column_delete = gtk.TreeViewColumn("Close", renderer_delete)
    column_delete.set_cell_data_func(renderer_delete, render_delete_column, treeview_channels)
    treeview_channels.append_column(column_delete)

    treeview_channels.add_events(gtk.gdk.POINTER_MOTION_MASK | gtk.gdk.LEAVE_NOTIFY_MASK)
    treeview_channels.connect("button-press-event",
                              lambda t, e: on_treeview_button_press(t, e, state.channel_liststore, state))
    treeview_channels.connect("motion-notify-event",
                              lambda t, e: on_treeview_motion(t, e, state.channel_liststore))
    treeview_channels.connect("leave-notify-event", on_treeview_leave)

    scrolled_channels.add(treeview_channels)
    scrolled_channels.set_size_request(300, -1)
    right_vbox.pack_start(scrolled_channels, True, True, 2)

    # Populate and start background checks
    populate_data_channels(state.channel_liststore, state)
    check_current_selection(state)

    state.timeout_id = gtk.timeout_add(500, check_current_selection, state)
    state.last_containers = set(id(c) for c in gwy.gwy_app_data_browser_get_containers())
    state.data_browser_timeout_id = gtk.timeout_add(1000, check_data_browser_changes,
                                                    state.channel_liststore, state)
    logger.debug("Started periodic data browser check")

    state.window.set_default_size(700, 600)
    state.window.show_all()


# --------------------------------
# Remember Last Save Directory
# --------------------------------
LAST_SAVE_DIR = os.path.expanduser("~/Desktop")


def save_last_dir(save_dir):
    """Persist last chosen directory across sessions (in user's home)."""
    try:
        with open(os.path.expanduser("~/.gwyddion_last_dir"), "w") as f:
            f.write(save_dir)
        logger.info("Saved last directory: %s", save_dir)
    except Exception as e:
        logger.warning("Failed to save last directory: %s", str(e))


def load_last_dir():
    """Retrieve last chosen directory if valid, else Desktop."""
    try:
        with open(os.path.expanduser("~/.gwyddion_last_dir"), "r") as f:
            last_dir = f.read().strip()
        if os.path.isdir(last_dir) and os.access(last_dir, os.W_OK):
            logger.info("Loaded last directory: %s", last_dir)
            return last_dir
        else:
            logger.warning("Last directory %s is invalid or non-writable", last_dir)
    except Exception:
        logger.info("No last directory found, using Desktop")
    return os.path.expanduser("~/Desktop")


def show_save_confirmation_dialog(save_files, parent):
    """Confirm that N files will be saved into a chosen directory.

    Args:
        save_files: list of tuples (basename, [titles], save_path)
    """
    num_files = len(save_files)
    save_dir = os.path.dirname(save_files[0][2]) if save_files else ""
    message = "%d files will be saved under '%s' " % (num_files, save_dir)

    dialog = gtk.MessageDialog(parent=parent, flags=gtk.DIALOG_MODAL,
                               type=gtk.MESSAGE_QUESTION, buttons=gtk.BUTTONS_OK_CANCEL,
                               message_format=message)
    dialog.set_title("Confirm Save as .gwy")
    response = dialog.run()
    dialog.destroy()

    ok = (response == gtk.RESPONSE_OK)
    logger.info("User %s save as .gwy operation", "confirmed" if ok else "cancelled")
    return ok


def get_save_dir(parent, channel_liststore):
    """Prompt once for a directory to save all .gwy outputs into."""
    global LAST_SAVE_DIR

    dialog = gtk.FileChooserDialog(title="Select Save Directory for All SPM Files",
                                   parent=parent,
                                   action=gtk.FILE_CHOOSER_ACTION_SELECT_FOLDER,
                                   buttons=(gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                            gtk.STOCK_OK, gtk.RESPONSE_OK))

    # Try the directory of any selected SPM file first
    initial_dir = None
    for row in channel_liststore:
        if row[5]:
            file_dir = os.path.dirname(row[5])
            if os.path.isdir(file_dir) and os.access(file_dir, os.W_OK):
                initial_dir = file_dir
                logger.info("Using SPM file directory: %s", initial_dir)
                break
    if not initial_dir:
        initial_dir = load_last_dir()
        logger.info("No valid SPM file directory, using last directory: %s", initial_dir)

    dialog.set_current_folder(initial_dir)
    response = dialog.run()
    if response == gtk.RESPONSE_OK:
        save_dir = dialog.get_filename()
        logger.info("User selected save directory: %s", save_dir)
        if not os.access(save_dir, os.W_OK):
            logger.warning("No write access to %s, falling back to Desktop", save_dir)
            save_dir = os.path.expanduser("~/Desktop")
        LAST_SAVE_DIR = save_dir
        save_last_dir(save_dir)
    else:
        save_dir = None
        logger.info("User cancelled directory selection")
    dialog.destroy()
    return save_dir


# --------------------------------
# Ensure Log & Color Range Consistency
# --------------------------------
def ensure_processing_log(container, data_id, filename,
                          log_file="c:\\users\\allani\\appdata\\local\\temp\\SPM_autoprocess.log"):
    """Populate '/%d/log' with synthetic proc lines when possible.

    Tries to reconstruct a minimal log from the plugin log file for the given
    (container, data_id, file) to make saved .gwy self-descriptive.
    """
    try:
        with open(log_file, "r") as f:
            lines = f.readlines()

        log_entries = []
        search_str = "data_id %d in %s" % (data_id, filename)

        for line in lines:
            if search_str in line:
                for op in ["Ran ", "Cropped "]:
                    if op in line:
                        timestamp = line.split(" ")[0]
                        operation = line.split(" ")[-1].strip()
                        if op == "Ran ":
                            operation = line.split("Ran ")[1].split(" on ")[0].strip()
                            log_entries.append("proc::%s@%s" % (operation, timestamp))
                        elif op == "Cropped ":
                            crop_line = next((l for l in lines
                                              if "Cropped in place data_id %d" % data_id in l), None)
                            if crop_line:
                                crop_params = next((l for l in lines
                                                    if "tool::GwyToolCrop" in l and
                                                       "data_id %d" % data_id in l), None)
                                if crop_params:
                                    log_entries.append(crop_params.strip())

        log_value = "\n".join(log_entries) if log_entries else None
        if log_value:
            container.set_string_by_name("/%d/log" % data_id, log_value)
            logger.info("Set processing log for data_id %d in %s", data_id, filename)
        else:
            logger.warning("No processing log constructed for data_id %d in %s",
                           data_id, filename)
    except Exception as e:
        logger.warning("Failed to set log for data_id %d in %s: %s",
                       data_id, filename, str(e))


def ensure_color_range(container, data_id, filename):
    """If no color range metadata exists, set defaults from actual data min/max."""
    try:
        data_field = container.get_object_by_name("/%d/data" % data_id)
        if not container.contains_by_name("/%d/base/range" % data_id):
            min_val, max_val = gwy.gwy_data_field_get_min_max(data_field)
            container.set_value_by_name("/%d/base/range" % data_id, (min_val, max_val))
            logger.info("Set fallback color range for data_id %d in %s: min=%f, max=%f",
                        data_id, filename, min_val, max_val)
        if not container.contains_by_name("/%d/base/range-type" % data_id):
            container.set_int32_by_name("/%d/base/range-type" % data_id, 1)  # GWY_LAYER_RANGE_FIXED
            logger.info("Set fixed color range type for data_id %d in %s", data_id, filename)
    except Exception as e:
        logger.warning("Failed to set color range for data_id %d in %s: %s",
                       data_id, filename, str(e))


# --------------------------------
# Batch Save to .gwy
# --------------------------------
def save_as_gwy(button, channel_liststore, state):
    """Save each SPM file's selected channels into a single .gwy file (report files saved)."""
    DATA_KEY_L = "/%d/data"
    TITLE_KEY_L = "/%d/data/title"
    SHOW_KEY_L = "/%d/data/visible"
    PALETTE_KEY_L = "/%d/base/palette"
    LOG_KEY_L = "/%d/log"
    RANGE_KEY_L = "/%d/base/range"
    RANGE_TYPE_KEY_L = "/%d/base/range-type"

    # Gather selected unique (filename, data_id)
    selected = []
    seen = set()
    for row in channel_liststore:
        checked, title, is_channel, container, data_id, filename, _, _ = row
        if checked and container and is_channel and data_id != -1:
            key = (filename, data_id)
            if key not in seen:
                logger.info("Processing channel: title=%s, data_id=%d, filename=%s",
                            title, data_id, filename)
                selected.append((container, data_id, title, filename))
                seen.add(key)

    if not selected:
        logger.error("No channels selected for saving")
        show_message_dialog(gtk.MESSAGE_ERROR, "No channels selected for saving")
        return

    # Group by filename (SPM file)
    groups = {}
    for container, data_id, title, filename in selected:
        groups.setdefault(filename, []).append((container, data_id, title))

    if not groups:
        logger.error("No valid SPM files found for selected channels")
        show_message_dialog(gtk.MESSAGE_ERROR, "No valid SPM files found for saving")
        return

    # Directory selection (single root for all outputs)
    save_dir = get_save_dir(state.window, channel_liststore)
    if save_dir is None:
        logger.info("Save as .gwy operation cancelled by user in file chooser")
        return

    # Prepare output names (avoid overwrite by suffixing _processed_N)
    save_files = []
    for filename, channels in groups.items():
        base = os.path.splitext(os.path.basename(filename))[0]
        out_path = os.path.join(save_dir, "%s.gwy" % base)
        counter = 1
        while os.path.exists(out_path):
            out_path = os.path.join(save_dir, "%s_processed_%d.gwy" % (base, counter))
            counter += 1
        save_files.append((base, [t for _, _, t in channels], out_path))

    # Confirm and write
    if not show_save_confirmation_dialog(save_files, state.window):
        logger.info("Save as .gwy operation cancelled by user")
        return

    def _save_group(filename, channels, save_path):
        logger.info("Attempting to save %d channels to %s", len(channels), save_path)

        container = channels[0][0]  # All channels are from same container
        success = True

        # Ensure logs/ranges exist for each channel prior to save
        for _, data_id, title in channels:
            try:
                if not container.contains_by_name(DATA_KEY_L % data_id):
                    logger.error("No data field for data_id %d (%s) in %s",
                                 data_id, title, filename)
                    success = False
                    continue
                ensure_processing_log(container, data_id, filename)
                ensure_color_range(container, data_id, filename)
                logger.info("Prepared data_id %d (%s) for %s", data_id, title, save_path)
            except Exception as e:
                logger.error("Failed to prepare data_id %d (%s) for %s: %s",
                             data_id, title, save_path, str(e))
                success = False

        try:
            # First, try generic save
            op = gwy.gwy_file_save(container, save_path, gwy.RUN_NONINTERACTIVE)
            if op == 0:
                logger.warning("gwy_file_save failed, fallback to gwy_file_func_run_save")
                success = gwy.gwy_file_func_run_save("gwyddion", container,
                                                     save_path, gwy.RUN_NONINTERACTIVE)
                if not success:
                    logger.error("Failed to save %s using gwy_file_func_run_save", save_path)
                    show_message_dialog(gtk.MESSAGE_ERROR, "Failed to save %s" % save_path)
                    return False

            if not os.path.exists(save_path):
                logger.error("File %s was not created", save_path)
                show_message_dialog(gtk.MESSAGE_ERROR, "File %s was not created" % save_path)
                return False

            logger.info("Saved %s", save_path)
            return True
        except Exception as e:
            logger.error("Failed to save %s: %s", save_path, str(e))
            show_message_dialog(gtk.MESSAGE_ERROR, "Failed to save %s: %s" % (save_path, str(e)))
            return False

    # Execute per file group and count files saved (not channels)
    files_saved = 0
    for filename, channels in groups.items():
        out = next(path for fname, _, path in save_files
                   if fname == os.path.splitext(os.path.basename(filename))[0])
        if _save_group(filename, channels, out):
            files_saved += 1

    if files_saved == 0:
        logger.error("No items successfully processed")
        show_message_dialog(gtk.MESSAGE_ERROR, "No items successfully processed")
    else:
        logger.info("Saved %d .gwy files", files_saved)
        show_message_dialog(gtk.MESSAGE_INFO, "Saved %d files as .gwy" % files_saved)

    populate_data_channels(channel_liststore, state)



# --------------------------------
# Window Close / Cleanup
# --------------------------------
def on_window_delete_event(widget, event, state):
    """Cleanup timers, signals, and release the global singleton when closing."""
    global _plugin_gui_instance

    # Stop periodic timers
    if getattr(state, 'timeout_id', None) is not None:
        try:
            gobject.source_remove(state.timeout_id)
        except Exception:
            pass
        state.timeout_id = None

    if getattr(state, 'data_browser_timeout_id', None) is not None:
        try:
            gobject.source_remove(state.data_browser_timeout_id)
        except Exception:
            pass
        state.data_browser_timeout_id = None

    # Disconnect selection signals
    for conn_id, container, data_id in list(getattr(state, 'selection_connections', [])):
        try:
            for key in [SELECTION_KEYS[0] % data_id, SELECTION_KEYS[1] % data_id]:
                if container and container.contains_by_name(key):
                    selection = container.get_object_by_name(key)
                    selection.disconnect(conn_id)
        except Exception:
            pass
    state.selection_connections = []
    # Clear current references
    state.current_container = None
    state.current_data_id = None

    # Optional: logger handler tidy-up (matches your existing behavior)
    try:
        if state.window is not None:
            try:
                state.window.set_data("autoprocess_singleton", None)
            except Exception:
                pass
            state.window.destroy()
    except Exception:
        pass
    state.window = None

    # Optional: also clear your module-global, if you keep one
    try:
        global _plugin_gui_instance
        _plugin_gui_instance = None
    except Exception:
        pass

    return True  



# --------------------------------
# Channel & File Table Utilities
# --------------------------------
def toggle_channel_selection(cell, path, channel_liststore):
    """Toggle the checkbox for a channel row (not file header rows)."""
    if channel_liststore[path][2]:
        channel_liststore[path][0] = not channel_liststore[path][0]
        logger.debug("Toggled %s to %s", channel_liststore[path][1], channel_liststore[path][0])


def select_all_channels(button, channel_liststore, select=True):
    """Check/uncheck all selectable channel rows (helper; unused by UI directly)."""
    def set_selection(model, path, _iter, _select):
        if model.get_value(_iter, 2):
            model.set_value(_iter, 0, _select)
    channel_liststore.foreach(set_selection, select)


def delete_file(cell, path, channel_liststore, state):
    """Close an SPM file (container) from the data browser via 'Close' column."""
    container = channel_liststore[path][3]
    if container and channel_liststore[path][4] == -1 and not channel_liststore[path][2]:
        filename = channel_liststore[path][5]
        logger.info("Attempting to delete SPM file: %s", filename)
        try:
            gwy.gwy_app_data_browser_remove(container)
            populate_data_channels(channel_liststore, state)
        except Exception:
            logger.error("Failed to delete SPM file %s", filename)


def create_pixbuf(stock_id, fallback_color):
    """Small helper to create an icon pixbuf with a GTK stock as primary source."""
    try:
        image = gtk.Image()
        image.set_from_stock(stock_id, gtk.ICON_SIZE_BUTTON)
        pixbuf = image.get_pixbuf()
        if pixbuf:
            return pixbuf
    except Exception:
        pass
    pixbuf = gtk.gdk.Pixbuf(gtk.gdk.COLORSPACE_RGB, True, 8, 16, 16)
    pixbuf.fill(fallback_color)
    return pixbuf


def populate_data_channels(channel_liststore, state):
    """List all open SPM files and their channels into the right pane table.
    Also populates the per-index selection dropdown with dynamic options.
    """
    # Preserve current checkbox states per (container, data_id)
    checkbox_states = {}
    for row in channel_liststore:
        container, data_id, filename = row[3], row[4], row[5]
        if container and data_id != -1:
            key = (id(container), data_id)
            checkbox_states[key] = row[0]
        elif container and data_id == -1 and row[1] != "──────────────────":
            key = (id(container), -1)
            checkbox_states[key] = row[0]

    # Disconnect old selection signals
    for conn_id, container, data_id in state.selection_connections:
        try:
            for key in [SELECTION_KEYS[0] % data_id, SELECTION_KEYS[1] % data_id]:
                if container.contains_by_name(key):
                    selection = container.get_object_by_name(key)
                    selection.disconnect(conn_id)
                    logger.debug("Disconnected selection signal for data_id %d, key %s", data_id, key)
        except:
            logger.debug("Error disconnecting selection signal for data_id %d", data_id)
    state.selection_connections = []

    # Determine max number of channels across all files; gather names by index
    containers = gwy.gwy_app_data_browser_get_containers()
    max_channels = 0
    channel_names_by_index = {}
    for container in containers:
        data_ids = gwy.gwy_app_data_browser_get_data_ids(container)
        max_channels = max(max_channels, len(data_ids))
        for i, data_id in enumerate(data_ids):
            title = container.get_string_by_name(TITLE_KEY % data_id) or "Data %d" % data_id
            if i not in channel_names_by_index:
                channel_names_by_index[i] = set()
            channel_names_by_index[i].add(title)

    # Prepare the dropdown model (with a placeholder first row)
    if state.select_store is None:
        state.select_store = gtk.ListStore(str, bool, str)
    state.select_store.clear()
    state.select_store.append(["Select Options...", False, "Placeholder to guide user"])

    # Fill entries "First/Second/Third..." up to max_channels
    for i in range(max_channels):
        if i == 0:
            option_label = "First Datachannels"
        elif i == 1:
            option_label = "Second Datachannels"
        elif i == 2:
            option_label = "Third Datachannels"
        elif i == 3:
            option_label = "Fourth Datachannels"
        elif i == 4:
            option_label = "Fifth Datachannels"
        elif i == 5:
            option_label = "Sixth Datachannels"
        elif i == 6:
            option_label = "Seventh Datachannels"
        elif i == 7:
            option_label = "Eighth Datachannels"
        else:
            suffix = "th"
            if (i + 1) % 10 == 1 and (i + 1) != 11:
                suffix = "st"
            elif (i + 1) % 10 == 2 and (i + 1) != 12:
                suffix = "nd"
            elif (i + 1) % 10 == 3 and (i + 1) != 13:
                suffix = "rd"
            option_label = "%d%s Datachannels" % (i + 1, suffix)

        names = ", ".join(sorted(channel_names_by_index.get(i, {"Unknown"})))
        tooltip_text = "Select channels: %s" % names
        state.select_store.append([option_label, False, tooltip_text])

    state.select_dropdown.set_model(state.select_store)
    state.select_dropdown.set_active(0)

    # Tooltip for active item
    state.select_dropdown.set_has_tooltip(True)

    def query_tooltip(combo, x, y, keyboard_mode, tooltip):
        if combo.get_active_iter():
            tooltip_text = combo.get_model()[combo.get_active()][2]
            tooltip.set_text(tooltip_text)
            return True
        return False

    state.select_dropdown.connect("query-tooltip", query_tooltip)

    # Fill the table
    channel_liststore.clear()
    delete_pixbuf = create_pixbuf(gtk.STOCK_CLOSE, 0xff0000ff)
    remove_pixbuf = create_pixbuf(gtk.STOCK_REMOVE, 0xffa500ff)

    for idx, container in enumerate(containers, 1):
        filename = container.get_string_by_name(FILENAME_KEY) or "Container %d" % id(container)
        filename = os.path.basename(filename) if filename else "Unknown SPM File"

        file_key = (id(container), -1)
        file_checked = checkbox_states.get(file_key, False)

        channel_liststore.append([file_checked, "<b>File%d: %s</b>" % (idx, filename),
                                  False, container, -1, filename, delete_pixbuf, remove_pixbuf])

        for data_id in gwy.gwy_app_data_browser_get_data_ids(container):
            title = container.get_string_by_name(TITLE_KEY % data_id) or "Data %d" % data_id
            channel_key = (id(container), data_id)
            channel_checked = checkbox_states.get(channel_key, False)
            channel_liststore.append([channel_checked, "  %s" % title, True,
                                      container, data_id, filename, None, None])

            for selection_key in [SELECTION_KEYS[0] % data_id, SELECTION_KEYS[1] % data_id]:
                if container.contains_by_name(selection_key):
                    selection = container.get_object_by_name(selection_key)
                    try:
                        conn_id = selection.connect("changed", selection_changed,
                                                    container, data_id, state)
                        state.selection_connections.append((conn_id, container, data_id))
                        logger.debug("Connected selection signal for data_id %d", data_id)
                    except Exception as e:
                        logger.error("Failed to connect selection signal for data_id %d: %s",
                                     data_id, str(e))

        channel_liststore.append([False, "──────────────────", False, None, -1, "", None, None])

    logger.info("Populated %d data channels from %d SPM files, max channels: %d",
                sum(len(gwy.gwy_app_data_browser_get_data_ids(c)) for c in containers),
                len(containers), max_channels)


# --------------------------------
# Selection Helpers
# --------------------------------
def get_selection_params(container, data_id):
    """Return integer crop rectangle (x, y, w, h) from current rectangle selection.

    Coordinates are converted from real units to pixel indices.
    """
    try:
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            logger.error("No data field for data_id %d", data_id)
            return None, None, None, None

        dx, dy = data_field.get_dx(), data_field.get_dy()
        selection_key = SELECTION_KEYS[0] % data_id

        if container.contains_by_name(selection_key):
            selection = container.get_object_by_name(selection_key)
            try:
                coords = selection.get_data()[:4] if hasattr(selection, 'get_data') else None
                if coords and len(coords) == 4:
                    logger.debug("Raw selection coords for data_id %d: %s", data_id, coords)
                    x1 = int(coords[0] / dx)
                    y1 = int(coords[1] / dy)
                    x2 = int(coords[2] / dx)
                    y2 = int(coords[3] / dy)

                    width = x2 - x1
                    height = y2 - y1
                    x = x1 if width >= 0 else x2
                    y = y1 if height >= 0 else y2
                    width = abs(width)
                    height = abs(height)

                    logger.debug("Normalized selection for data_id %d: x=%d, y=%d, width=%d, height=%d",
                                 data_id, x, y, width, height)
                    return x, y, width, height
            except Exception as e:
                logger.error("Failed to process selection for data_id %d: %s", data_id, str(e))
        else:
            logger.debug("No selection found for data_id %d at %s", data_id, selection_key)
        return None, None, None, None
    except Exception as e:
        logger.error("Failed to get selection for data_id %d: %s", data_id, str(e))
        return None, None, None, None


def selection_changed(selection, index, container, data_id, state, *args):
    """GTK signal: update crop fields when selection rectangle changes."""
    try:
        if state.window is None:
            logger.debug("Skipping selection update for data_id %d: GUI is closed", data_id)
            return
        x, y, width, height = get_selection_params(container, data_id)
        if x is not None:
            state.x_entry.set_text(str(x))
            state.y_entry.set_text(str(y))
            state.width_entry.set_text(str(width))
            state.height_entry.set_text(str(height))
            logger.debug("Dynamic selection update for data_id %d: x=%d, y=%d, width=%d, height=%d",
                         data_id, x, y, width, height)
        else:
            state.x_entry.set_text("")
            state.y_entry.set_text("")
            state.width_entry.set_text("")
            state.height_entry.set_text("")
            logger.debug("Cleared selection fields for data_id %d due to no valid selection",
                         data_id)
    except Exception as e:
        logger.error("Error in selection_changed for data_id %d: %s", data_id, str(e))
        if state.window is not None:
            state.x_entry.set_text("")
            state.y_entry.set_text("")
            state.width_entry.set_text("")
            state.height_entry.set_text("")


def check_current_selection(state):
    """Periodic task: track active container/data_id; attach a rectangle layer.

    Also synchronizes crop entry fields with the current selection.
    """
    if not gwy.gwy_app_data_browser_get_containers():
        return True

    current_container = gwy.gwy_app_data_browser_get_current(gwy.APP_CONTAINER)
    current_data_id = (gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_ID)
                       if current_container else None)

    if (current_container, current_data_id) != (state.current_container, state.current_data_id):
        # Disconnect old signals
        for conn_id, container, data_id in state.selection_connections:
            try:
                for key in [SELECTION_KEYS[0] % data_id, SELECTION_KEYS[1] % data_id]:
                    if container.contains_by_name(key):
                        selection = container.get_object_by_name(key)
                        selection.disconnect(conn_id)
            except:
                pass
        state.selection_connections = []

        state.current_container, state.current_data_id = current_container, current_data_id
        if current_container and current_data_id is not None:
            data_view = gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_VIEW)
            if not data_view or not isinstance(data_view, gobject.GObject):
                return True

            layer = gobject.new(gobject.type_from_name('GwyLayerRectangle'))
            selection_key = SELECTION_KEYS[0] % current_data_id
            layer.set_selection_key(selection_key)
            layer.set_property("is-crop", True)
            data_view.set_top_layer(layer)

            for key in ["/%d/select/pointer" % current_data_id,
                        "/%d/select/line" % current_data_id]:
                if current_container.contains_by_name(key):
                    current_container.remove_by_name(key)

            if not current_container.contains_by_name(selection_key):
                selection = gobject.new(gobject.type_from_name('GwySelectionRectangle'))
                selection.set_max_objects(1)
                current_container.set_object_by_name(selection_key, selection)

            data_field = current_container.get_object_by_name(DATA_KEY % current_data_id)
            dx, dy = data_field.get_dx(), data_field.get_dy()
            xres, yres = data_field.get_xres(), data_field.get_yres()
            default_width = min(0, xres // 2)
            default_height = min(0, yres // 2)
            default_coords = [0.0, 0.0, default_width * dx, default_height * dy]
            selection = current_container.get_object_by_name(selection_key)
            selection.set_object(0, default_coords)
            selection.crop(0.0, 0.0, xres * dx, yres * dy)

            try:
                conn_id = selection.connect("changed", selection_changed,
                                            current_container, current_data_id, state)
                state.selection_connections.append((conn_id, current_container, current_data_id))
            except Exception as e:
                pass

            x, y, width, height = get_selection_params(current_container, current_data_id)
            if all(v is not None for v in (x, y, width, height)):
                state.x_entry.set_text(str(x))
                state.y_entry.set_text(str(y))
                state.width_entry.set_text(str(width))
                state.height_entry.set_text(str(height))
            else:
                state.x_entry.set_text("")
                state.y_entry.set_text("")
                state.width_entry.set_text("")
                state.height_entry.set_text("")
    else:
        if current_container and current_data_id is not None:
            x, y, width, height = get_selection_params(current_container, current_data_id)
            if all(v is not None for v in (x, y, width, height)):
                if (state.x_entry.get_text().strip() != str(x) or
                    state.y_entry.get_text().strip() != str(y) or
                    state.width_entry.get_text().strip() != str(width) or
                    state.height_entry.get_text().strip() != str(height)):
                    state.x_entry.set_text(str(x))
                    state.y_entry.set_text(str(y))
                    state.width_entry.set_text(str(width))
                    state.height_entry.set_text(str(height))
            else:
                state.x_entry.set_text("")
                state.y_entry.set_text("")
                state.width_entry.set_text("")
                state.height_entry.set_text("")
    return True


def data_browser_changed(obj, arg, channel_liststore, state):
    """Callback: data browser changed (unused externally); refresh listing."""
    logger.debug("Data browser changed, updating channel list")
    populate_data_channels(channel_liststore, state)


def check_data_browser_changes(channel_liststore, state):
    """Periodic task: detect addition/removal of containers; auto-close GUI
    if Gwyddion data browser empties (likely app shutdown).
    """
    current_containers = gwy.gwy_app_data_browser_get_containers()
    if not current_containers and state.window is not None:
        logger.debug("No containers in data browser, Gwyddion likely closed; shutting down GUI")
        on_window_delete_event(state.window, None, state)
        gtk.main_quit()
        return False

    current_container_ids = set(id(c) for c in current_containers)
    if (not hasattr(state, 'last_containers') or
        state.last_containers != current_container_ids or
        len(current_containers) != len(state.last_containers)):
        logger.debug("Data browser containers changed or count mismatch, updating channel list")
        populate_data_channels(channel_liststore, state)
        state.last_containers = current_container_ids
    return True


# --------------------------------
# Processing Helpers
# --------------------------------
def get_min_max(container, data_id):
    """Return min/max over the channel or over all channels of a file row."""
    try:
        if data_id == -1:
            data_ids = gwy.gwy_app_data_browser_get_data_ids(container)
            if not data_ids:
                return None, None
            global_min, global_max = float('inf'), float('-inf')
            for did in data_ids:
                data_field = container.get_object_by_name(DATA_KEY % did)
                if data_field:
                    global_min = min(global_min, data_field.get_min())
                    global_max = max(global_max, data_field.get_max())
            return global_min, global_max
        else:
            data_field = container.get_object_by_name(DATA_KEY % data_id)
            return (data_field.get_min(), data_field.get_max()) if data_field else (None, None)
    except Exception:
        return None, None


def validate_crop_params(data_field, x, y, width, height, filename, spm_filename):
    """Ensure crop rectangle is positive and within image bounds."""
    xres, yres = data_field.get_xres(), data_field.get_yres()
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        return False, "Invalid crop parameters for %s in %s" % (filename, spm_filename)
    if x + width > xres or y + height > yres:
        return False, ("Crop area out of bounds for %s in %s: x=%d, y=%d, width=%d, height=%d"
                       % (filename, spm_filename, x, y, width, yres))
    return True, None


def process_selected_channels(channel_liststore, operation, no_selection_msg, success_msg, state):
    """Generic batch runner for per-channel operations.

    Args:
        operation(container, data_id, title, filename): function applied to each
    """
    selected = []
    for row in channel_liststore:
        checked, title, is_channel, container, data_id, filename, _, _ = row
        if checked and container and (is_channel or data_id == -1):
            selected.append((container, data_id, title, filename))

    if not selected:
        logger.error(no_selection_msg)
        show_message_dialog(gtk.MESSAGE_ERROR, no_selection_msg)
        return

    success_count = 0
    for container, data_id, title, filename in selected:
        try:
            operation(container, data_id, title, filename)
            success_count += 1
        except Exception as e:
            logger.error("Failed to process %s, data_id %d: %s", filename, data_id, str(e))

    if success_count > 0:
        logger.info(success_msg % success_count)
        show_message_dialog(gtk.MESSAGE_INFO, success_msg % success_count)
    else:
        logger.error("No items successfully processed")
        show_message_dialog(gtk.MESSAGE_ERROR, "No items successfully processed")


def apply_palette(button, channel_liststore, state):
    """Assign the selected gradient palette to checked channels."""
    active_iter = state.palette_combobox.get_active_iter()
    if not active_iter:
        logger.error("No palette selected")
        show_message_dialog(gtk.MESSAGE_ERROR, "No palette selected. Please choose a palette.")
        return
    palette_name = state.palette_combobox.get_model().get_value(active_iter, 0)

    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            raise ValueError("No data field")
        container.set_string_by_name("/%d/base/palette" % data_id, palette_name)
        data_field.data_changed()
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        logger.info("Set palette %s on data_id %d (%s) in %s",
                    palette_name, data_id, title, filename)

    process_selected_channels(channel_liststore, operation,
                              "No channels selected for palette change",
                              "Palette %s applied to %%d channels" % palette_name, state)


def apply_fixed_color_range(button, channel_liststore, state):
    """Set a fixed display range from 'Start' and 'End' entries (no swap)."""
    try:
        start_val = float(state.min_entry.get_text().strip())
        end_val = float(state.max_entry.get_text().strip())
    except ValueError:
        show_message_dialog(gtk.MESSAGE_ERROR,
                            "Invalid input: Please enter valid numeric values for Start and End.")
        return

    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FIXED)
        container.set_double_by_name(BASE_MIN_KEY % data_id, start_val)
        container.set_double_by_name(BASE_MAX_KEY % data_id, end_val)
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        logger.info("Applied fixed color range: Start=%f, End=%f on data_id=%d in %s",
                    start_val, end_val, data_id, filename)

    process_selected_channels(channel_liststore, operation,
                              "No channels selected for color range",
                              "Fixed color range applied to %d channels", state)


def set_to_full_range(button, channel_liststore, state):
    """Restore full range display. If zero-to-min was used earlier, undo offset."""
    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            raise ValueError("No data field")
        if (container.contains_by_name(ORIGINAL_MIN_KEY % data_id) and
            container.contains_by_name(ORIGINAL_MAX_KEY % data_id)):
            original_min = container.get_double_by_name(ORIGINAL_MIN_KEY % data_id)
            current_min = data_field.get_min()
            if original_min != current_min:
                data_field.add(original_min - current_min)
                data_field.data_changed()
            container.remove_by_name(ORIGINAL_MIN_KEY % data_id)
            container.remove_by_name(ORIGINAL_MAX_KEY % data_id)
            logger.info("Restored original min=%g for data_id %d in %s",
                        original_min, data_id, filename)

        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FULL)
        if container.contains_by_name(BASE_MIN_KEY % data_id):
            container.remove_by_name(BASE_MIN_KEY % data_id)
        if container.contains_by_name(BASE_MAX_KEY % data_id):
            container.remove_by_name(BASE_MAX_KEY % data_id)

        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        current_data_id = (gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_ID)
                           if container == gwy.gwy_app_data_browser_get_current(gwy.APP_CONTAINER)
                           else None)
        if current_data_id == data_id:
            min_val, max_val = data_field.get_min(), data_field.get_max()
            state.min_entry.set_text("%.6g" % min_val if min_val is not None else "")
            state.max_entry.set_text("%.6g" % max_val if max_val is not None else "")
        logger.info("Set full range for data_id %d in %s", data_id, filename)

    process_selected_channels(channel_liststore, operation,
                              "No channels selected for full range",
                              "Full range applied to %d channels", state)


def invert_mapping(button, channel_liststore, state):
    """Swap current min/max (either from explicit base or live data)."""
    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            raise ValueError("No data field")

        current_min = (container.get_double_by_name(BASE_MIN_KEY % data_id)
                       if container.contains_by_name(BASE_MIN_KEY % data_id) else data_field.get_min())
        current_max = (container.get_double_by_name(BASE_MAX_KEY % data_id)
                       if container.contains_by_name(BASE_MAX_KEY % data_id) else data_field.get_max())

        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FIXED)
        container.set_double_by_name(BASE_MIN_KEY % data_id, current_max)
        container.set_double_by_name(BASE_MAX_KEY % data_id, current_min)
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        logger.info("Inverted color range for data_id %d in %s", data_id, filename)

    process_selected_channels(channel_liststore, operation,
                              "No channels selected for invert mapping",
                              "Color range inverted for %d channels", state)


def set_zero_to_minimum(button, channel_liststore, state):
    """Shift data so min becomes 0; cache original min/max to allow restore."""
    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            raise ValueError("No data field")

        current_min, current_max = data_field.get_min(), data_field.get_max()
        if not container.contains_by_name(ORIGINAL_MIN_KEY % data_id):
            container.set_double_by_name(ORIGINAL_MIN_KEY % data_id, current_min)
        if not container.contains_by_name(ORIGINAL_MAX_KEY % data_id):
            container.set_double_by_name(ORIGINAL_MAX_KEY % data_id, current_max)

        data_field.add(-current_min)
        data_field.data_changed()

        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FIXED)
        container.set_double_by_name(BASE_MIN_KEY % data_id, 0.0)
        container.set_double_by_name(BASE_MAX_KEY % data_id, current_max - current_min)

        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        current_data_id = (gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_ID)
                           if container == gwy.gwy_app_data_browser_get_current(gwy.APP_CONTAINER)
                           else None)
        if current_data_id == data_id:
            state.min_entry.set_text("0")
            state.max_entry.set_text("%.6g" % (current_max - current_min))

        logger.info("Set zero to minimum for data_id %d in %s, stored original min=%g, max=%g",
                    data_id, filename, current_min, current_max)

    process_selected_channels(channel_liststore, operation,
                              "No channels selected for set zero to minimum",
                              "Zero to minimum applied to %d channels", state)


def apply_crop(button, channel_liststore, state):
    """Crop selected channels using X/Y/W/H (px). Supports in-place or new channel."""
    try:
        x = int(state.x_entry.get_text().strip())
        y = int(state.y_entry.get_text().strip())
        width = int(state.width_entry.get_text().strip())
        height = int(state.height_entry.get_text().strip())
        create_new = state.create_new_check.get_active()
        keep_offsets = state.keep_offsets_check.get_active()
    except ValueError:
        logger.error("Invalid crop parameters")
        show_message_dialog(gtk.MESSAGE_ERROR,
                            "Invalid crop parameters. Please enter valid integer values.")
        return

    # Build selection list across files
    selected = []
    valid_channels = []
    invalid_channels = []
    for row in channel_liststore:
        checked, title, is_channel, container, data_id, filename, _, _ = row
        if checked and container and (is_channel or data_id == -1):
            selected.append((container, data_id, title, filename))

    if not selected:
        logger.error("No files or channels selected for cropping")
        show_message_dialog(gtk.MESSAGE_ERROR, "No files or channels selected for cropping")
        return

    for container, data_id, title, filename in selected:
        data_ids = (gwy.gwy_app_data_browser_get_data_ids(container) if data_id == -1 else [data_id])
        for did in data_ids:
            data_field = container.get_object_by_name(DATA_KEY % did)
            if not data_field:
                invalid_channels.append((container, did, title, filename, "No data field"))
                continue
            valid, error_msg = validate_crop_params(data_field, x, y, width, height, title, filename)
            if valid:
                valid_channels.append((container, did, title, filename))
            else:
                invalid_channels.append((container, did, title, filename, error_msg))

    if invalid_channels:
        response = show_crop_conflict_dialog(invalid_channels, valid_channels, channel_liststore,
                                             state, x, y, width, height, create_new, keep_offsets)
        if response in ["cancel", "cancel_list"]:
            logger.info("Crop operation cancelled by user")
            return
        selected = valid_channels
    else:
        selected = valid_channels

    def operation(container, data_id, title, filename):
        crop_channel(container, data_id, title, filename, x, y, width, height, create_new, keep_offsets)

    if selected:
        process_selected_channels(channel_liststore, operation, "No valid channels to crop",
                                 "Cropping applied to %d items", state)
        populate_data_channels(channel_liststore, state)
    else:
        logger.error("No valid channels to crop after validation")
        show_message_dialog(gtk.MESSAGE_ERROR, "No valid channels to crop after validation")


def crop_channel(container, data_id, title, filename, x, y, width, height, create_new, keep_offsets):
    """Perform the actual crop, either creating a new data field or in-place resize.

    Also appends a synthetic 'tool::GwyToolCrop(...)' line to '/%d/log'.
    """
    data_field = container.get_object_by_name(DATA_KEY % data_id)
    if not data_field:
        raise ValueError("No data field for data_id %d" % data_id)

    valid, error_msg = validate_crop_params(data_field, x, y, width, height, title, filename)
    if not valid:
        raise ValueError(error_msg)

    log_entry = ("tool::GwyToolCrop(all=%s, hold_selection=4, keep_offsets=%s, new_channel=%s, "
                 "x=%d, y=%d, width=%d, height=%d)@%s" %
                 (str(data_id == -1), str(keep_offsets), str(create_new),
                  x, y, width, height, datetime.now().isoformat()))
    logger.info(log_entry)

    log_key = "/%d/log" % data_id
    current_log = container.get_string_by_name(log_key) or ""
    container.set_string_by_name(log_key, current_log + log_entry + "\n")
    logger.debug("Manually added log entry to %s for data_id %d", log_key, data_id)

    new_id = None
    if create_new:
        # Create a new channel from the selected area
        new_data_field = data_field.area_extract(x, y, width, height)
        new_id = gwy.gwy_app_data_browser_add_data_field(new_data_field, container, True)

        old_title = container.get_string_by_name(TITLE_KEY % data_id) or "Data %d" % data_id
        container.set_string_by_name(TITLE_KEY % new_id, old_title + " (Cropped)")

        if container.contains_by_name("/%d/base" % data_id):
            new_data_field.copy(container.get_object_by_name(DATA_KEY % data_id), True)

        dx, dy = data_field.get_dx(), data_field.get_dy()
        new_data_field.set_xreal(width * dx)
        new_data_field.set_yreal(height * dy)

        if keep_offsets:
            new_data_field.set_xoffset(data_field.get_xoffset() + x * dx)
            new_data_field.set_yoffset(data_field.get_yoffset() + y * dy)
        else:
            new_data_field.set_xoffset(0.0)
            new_data_field.set_yoffset(0.0)

        new_data_field.data_changed()
        if container:
            gwy.gwy_app_data_browser_select_data_field(container, new_id)
        logger.info("Cropped to new data_id %d in %s", new_id, filename)
    else:
        # In-place crop
        data_field.resize(x, y, x + width, y + height)
        data_field.data_changed()
        if container:
            gwy.gwy_app_data_browser_select_data_field(container, data_id)
        logger.info("Cropped in place data_id %d in %s", data_id, filename)


def show_crop_conflict_dialog(invalid_channels, valid_channels, channel_liststore, state,
                              x, y, width, height, create_new, keep_offsets):
    """Inform the user that some channels are invalid for cropping; offer options."""
    total_channels = len(invalid_channels) + len(valid_channels)
    message = ("%d out of %d selected DataChannels cannot be processed. "
               "Proceed without them?" % (len(invalid_channels), total_channels))

    dialog = gtk.MessageDialog(parent=state.window, flags=gtk.DIALOG_MODAL,
                               type=gtk.MESSAGE_WARNING, buttons=gtk.BUTTONS_NONE,
                               message_format=message)
    dialog.add_button("Cancel", gtk.RESPONSE_CANCEL)
    dialog.add_button("Proceed", gtk.RESPONSE_OK)
    dialog.add_button("Cancel and list conflicts", gtk.RESPONSE_REJECT)
    dialog.add_button("Proceed and list conflicts", gtk.RESPONSE_APPLY)
    dialog.set_default_response(gtk.RESPONSE_CANCEL)
    response = dialog.run()
    dialog.destroy()

    response_map = {gtk.RESPONSE_CANCEL: "cancel",
                    gtk.RESPONSE_OK: "proceed",
                    gtk.RESPONSE_REJECT: "cancel_list",
                    gtk.RESPONSE_APPLY: "proceed_list"}
    response_str = response_map.get(response, "cancel")

    if response_str in ["cancel_list", "proceed_list"]:
        show_conflict_list_dialog(invalid_channels, state.window)

    logger.info("User selected %s for crop conflict dialog", response_str)
    return response_str


def show_conflict_list_dialog(invalid_channels, parent):
    """Display a scrollable list of invalid channels and their error reasons."""
    dialog = gtk.Dialog(title="Crop Conflicts", parent=parent, flags=gtk.DIALOG_MODAL,
                        buttons=(gtk.STOCK_OK, gtk.RESPONSE_OK))
    dialog.set_default_size(600, 300)

    scrolled = gtk.ScrolledWindow()
    scrolled.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)

    liststore = gtk.ListStore(str, str, str)
    for container, data_id, title, filename, error_msg in invalid_channels:
        liststore.append([title, filename, error_msg])

    treeview = gtk.TreeView(liststore)
    treeview.append_column(gtk.TreeViewColumn("Channel", gtk.CellRendererText(), text=0))
    treeview.append_column(gtk.TreeViewColumn("File", gtk.CellRendererText(), text=1))
    treeview.append_column(gtk.TreeViewColumn("Error", gtk.CellRendererText(), text=2))

    scrolled.add(treeview)
    dialog.vbox.pack_start(scrolled, True, True, 5)
    dialog.show_all()
    dialog.run()
    dialog.destroy()
    logger.info("Displayed conflict list dialog with %d invalid channels", len(invalid_channels))


def replay_selected_channels(button, channel_liststore, state):
    """Execute macro entries (from parsed log) on checked channels, in order."""
    if not state.macro:
        logger.error("No tools in macro to replay")
        show_message_dialog(gtk.MESSAGE_ERROR,
                            "No tools in macro to replay. Please load a log file.")
        return

    settings = gwy.gwy_app_settings_get()

    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        for entry in state.macro:
            function, params = entry["function"], entry["parameters"]
            for key, value in params.items():
                settings_key = "/module/%s/%s" % (function, key)
                try:
                    settings[settings_key] = value
                except ValueError:
                    logger.error("Invalid setting %s=%s for %s", settings_key, value, function)
                    raise ValueError("Invalid setting %s=%s for %s" % (key, value, function))
            gwy.gwy_app_undo_checkpoint(container, DATA_KEY % data_id)
            gwy.gwy_process_func_run(function, container, gwy.RUN_IMMEDIATE)
            logger.info("Ran %s on data_id %d in %s", function, data_id, filename)

    process_selected_channels(channel_liststore, operation,
                              "No channels selected for replay",
                              "Macro replay completed on %d channels", state)


# --------------------------------
# Gradients Inventory
# --------------------------------
def get_gradient_names():
    """Return list of available gradients with pre-rendered pixbufs if possible.

    Falls back to a small subset of names when sampling is not available.
    """
    known_gradients = [
        'Blend1', 'Blend2', 'Blue', 'Blue-Cyan', 'Blue-Violet', 'Blue-Yellow', 'Body', 'BW1', 'BW2',
        'Caribbean', 'Clusters', 'Code-V', 'Cold', 'DFit', 'Digitalis', 'Gold', 'Gray-inverted',
        'Green', 'Green-Cyan', 'Green-Stripes-4', 'Green-Violet', 'Green-Yellow', 'Gwyddion.net',
        'Halcyon', 'Lines', 'Maple', 'MetroPro', 'Neon', 'NT-MDT', 'Olive', 'Painbow', 'Pink',
        'Plum', 'Pm3d', 'Rainbow1', 'Rainbow2', 'Red', 'Red-Cyan', 'Red-Stripes-5', 'Red-Violet',
        'Red-Yellow', 'RGB-Blue', 'RGB-Green', 'RGB-Red', 'Rust', 'Saw1', 'Shame', 'Sky', 'Sm2',
        'Spectral', 'Spectral-white', 'Spring', 'Viridis', 'Warm', 'Warpp-mono', 'Warpp-spectral',
        'Wyko', 'Yellow', 'Zones'
    ]
    try:
        gradient_inventory = gwy.gwy_gradients()
        palettes = []
        for name in known_gradients:
            try:
                gradient = gwy.gwy_gradients_get_gradient(name)
                pixbuf = gtk.gdk.Pixbuf(gtk.gdk.COLORSPACE_RGB, True, 8, 100, 20)
                gradient.sample_to_pixbuf(pixbuf)
                palettes.append((name, pixbuf))
            except Exception:
                pass
        if palettes:
            palettes.sort(key=lambda x: x[0])
            logger.info("Loaded %d gradient names", len(palettes))
            return palettes
    except Exception:
        pass
    return [('Gwyddion.net', None), ('Green', None), ('Blue', None)]


def render_channel_column(column, cell, model, iter, treeview):
    """Cell data func: manage 'Select' column visibility/checkbox per row type."""
    is_selectable = model.get_value(iter, 2)
    is_file_row = (not is_selectable and model.get_value(iter, 4) == -1 and
                   model.get_value(iter, 1) != "──────────────────")
    path = model.get_path(iter)
    select_hover_path = treeview.get_data("select_hover_path")

    if is_selectable:
        if isinstance(cell, gtk.CellRendererToggle):
            cell.set_property("visible", True)
            cell.set_property("active", model.get_value(iter, 0))
            cell.set_property("activatable", True)
        elif isinstance(cell, gtk.CellRendererText):
            cell.set_property("visible", False)
    else:
        if isinstance(cell, gtk.CellRendererToggle):
            cell.set_property("visible", False)
            cell.set_property("activatable", False)
        elif isinstance(cell, gtk.CellRendererText):
            cell.set_property("visible", False)


def render_delete_column(column, cell, model, iter, treeview):
    """Cell data func: show red 'X' to close SPM files on header rows only."""
    is_file_row = (not model.get_value(iter, 2) and model.get_value(iter, 4) == -1 and
                   model.get_value(iter, 1) != "──────────────────")
    path = model.get_path(iter)
    close_hover_path = treeview.get_data("close_hover_path")

    if is_file_row:
        cell.set_property("visible", True)
        cell.set_property("text", "X")
        cell.set_property("weight", pango.WEIGHT_BOLD)
        cell.set_property("foreground", "red" if close_hover_path == path else "black")
    else:
        cell.set_property("visible", False)


def on_treeview_button_press(treeview, event, channel_liststore, state):
    """Mouse click handler: toggle checkboxes; select data; close files."""
    if event.button == 1:
        pos = treeview.get_path_at_pos(int(event.x), int(event.y))
        if pos:
            path, column, cell_x, cell_y = pos
            # Column 0 = Select
            if column == treeview.get_column(0):
                if channel_liststore[path][2]:  # Channel row
                    toggle_channel_selection(None, path, channel_liststore)
                    return True
                return False

            # Channel title column: focus/select data
            elif channel_liststore[path][2]:
                container, data_id = channel_liststore[path][3], channel_liststore[path][4]
                if data_id != -1:
                    gwy.gwy_app_data_browser_select_data_field(container, data_id)
                    min_val = (container.get_double_by_name(BASE_MIN_KEY % data_id)
                               if container.contains_by_name(BASE_MIN_KEY % data_id) else None)
                    max_val = (container.get_double_by_name(BASE_MAX_KEY % data_id)
                               if container.contains_by_name(BASE_MAX_KEY % data_id) else None)
                    if min_val is None or max_val is None:
                        min_val, max_val = get_min_max(container, data_id)
                    state.min_entry.set_text("%.6g" % min_val if min_val is not None else "")
                    state.max_entry.set_text("%.6g" % max_val if max_val is not None else "")
                    x, y, width, height = get_selection_params(container, data_id)
                    if all(v is not None for v in (x, y, width, height)):
                        state.x_entry.set_text(str(x))
                        state.y_entry.set_text(str(y))
                        state.width_entry.set_text(str(width))
                        state.height_entry.set_text(str(height))

            # Column 2 = Close file on header rows
            if column == treeview.get_column(2) and channel_liststore[path][4] == -1 and not channel_liststore[path][2]:
                delete_file(None, path, channel_liststore, state)
                return True
    return False


def on_treeview_motion(treeview, event, channel_liststore):
    """Track hover to color the red 'X' and (optionally) select column."""
    pos = treeview.get_path_at_pos(int(event.x), int(event.y))
    old_select_hover_path = treeview.get_data("select_hover_path")
    old_close_hover_path = treeview.get_data("close_hover_path")
    new_select_hover_path = None
    new_close_hover_path = None

    if pos:
        path, column, cell_x, cell_y = pos
        if column == treeview.get_column(0):
            new_select_hover_path = path
        elif column == treeview.get_column(2):
            new_close_hover_path = path

    if old_select_hover_path != new_select_hover_path or old_close_hover_path != new_close_hover_path:
        treeview.set_data("select_hover_path", new_select_hover_path)
        treeview.set_data("close_hover_path", new_close_hover_path)
        treeview.queue_draw()
    return True

def on_treeview_leave(treeview, event):
    """Clear hover visuals when cursor leaves the treeview area."""
    old_select_hover_path = treeview.get_data("select_hover_path")
    old_close_hover_path = treeview.get_data("close_hover_path")
    if old_select_hover_path or old_close_hover_path:
        treeview.set_data("select_hover_path", None)
        treeview.set_data("close_hover_path", None)
        treeview.queue_draw()
    return True

def select_dropdown_changed(combo, channel_liststore, state):
    """Handle dropdown selection for dynamic channel indices."""
    active = combo.get_active()
    if active == 0:  # "Select Options" selected
        return
    row_index = active - 1  # Map to 0-based channel index
    new_state = not state.select_store[active][1]
    state.select_store[active][1] = new_state
    current_file_container = None
    current_row_index = -1
    for row in channel_liststore:
        container, data_id = row[3], row[4]
        if data_id == -1 and row[1] != "──────────────────":
            current_file_container = container
            current_row_index = -1
        elif current_file_container == container and row[2]:
            current_row_index += 1
            if current_row_index == row_index and row[2]:
                row[0] = new_state
                logger.debug("%s Channel %d for file %s", "Selected" if new_state else "Deselected", row_index + 1, row[5])
    combo.set_active(0)

def sync_select_all_check(checkbutton, channel_liststore, state):
    """Toggle all channel checkboxes based on Select All state."""
    active = checkbutton.get_active()
    for row in channel_liststore:
        if row[2]:  # Only update actual channels, not file headers or separators
            row[0] = active
            logger.debug("%s channel %s for file %s", "Selected" if active else "Deselected", row[1], row[5])
    logger.debug("Select All %s", "enabled" if active else "disabled")

def _find_autoprocess_window():
    """Return the existing AutoProcess window if it's already open, else None."""
    try:
        # PyGTK 2.x: list all toplevel GTK windows
        for w in gtk.window_list_toplevels():
            try:
                if isinstance(w, gtk.Window) and w.get_data("autoprocess_singleton") is True:
                    return w
            except Exception:
                continue
    except Exception:
        pass
    return None


# ---------- Entry point required by Gwyddion ----------
def run(data, mode):
    """
    Gwyddion entry point. Enforce a strict single AutoProcess window:
    if one exists, present it; otherwise create a fresh one.
    """
    # 1) If a window already exists, bring it to front and bail.
    existing = _find_autoprocess_window()
    if existing is not None:
        try:
            if hasattr(existing, "present"):
                existing.present()
        except Exception:
            pass
        dlg = gtk.MessageDialog(
            parent=None,
            flags=gtk.DIALOG_MODAL,
            type=gtk.MESSAGE_INFO,
            buttons=gtk.BUTTONS_OK,
            message_format="AutoProcess GUI is already open."
        )
        dlg.run()
        dlg.destroy()
        return

    # 2) Otherwise create a new GUI
    state = PluginState()
    try:
        key = gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_KEY)
        gwy.gwy_app_undo_qcheckpoint(data, [key])
    except Exception:
        pass

    create_gui(state)

