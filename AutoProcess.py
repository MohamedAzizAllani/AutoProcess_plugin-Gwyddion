# Plugin metadata for Gwyddion integration
plugin_menu = "/AutoProcess"  # Menu path in Gwyddion
plugin_desc = "Replays processing tools, sets color palettes, applies fixed color ranges, inverts mapping, sets zero to minimum, deletes/removes files, and automates cropping."  # Plugin description
plugin_type = "PROCESS"  # Plugin type for Gwyddion

# Import required libraries
import gwy  # Gwyddion library for SPM data processing
import gtk  # GTK for GUI creation
import os  # OS operations for file handling
import re  # Regular expressions for log parsing
import logging  # Logging for debugging and error tracking
import tempfile  # Temporary file handling for logs
import gtk.gdk  # GTK graphics for pixbufs
import pango  # Pango for text formatting in GUI
import gobject  # GObject for Gwyddion object management
from datetime import datetime  # Date/time for log timestamps
import sys  # System operations for stderr redirection
import time




# Logging setup
log_dir = tempfile.gettempdir()  # Get temporary directory for log file
log_file = os.path.join(log_dir, "SPM_autoprocess.log")  # Define log file path
logger = logging.getLogger('SPM_autoprocess')  # Create logger instance
logger.setLevel(logging.DEBUG)  # Set logging level to DEBUG
formatter = logging.Formatter("%(asctime)s,%(msecs)03d: %(message)s", datefmt='%Y-%m-%d %H:%M:%S')  # Define log format
try:
    if not os.path.exists(log_dir):  # Create log directory if it doesn't exist
        os.makedirs(log_dir)
    file_handler = logging.FileHandler(log_file, mode='w')  # Create file handler for logging
    file_handler.setFormatter(formatter)  # Apply format to file handler
    logger.addHandler(file_handler)  # Add file handler to logger
    logger.debug("Logger initialized with file handler: %s", log_file)  # Log successful initialization
except Exception as e:
    logger.debug("Failed to initialize file handler for %s: %s", log_file, str(e))  # Log file handler failure
    console_handler = logging.StreamHandler()  # Fallback to console handler
    console_handler.setFormatter(formatter)  # Apply format to console handler
    logger.addHandler(console_handler)  # Add console handler to logger
    logger.debug("Using console handler due to file handler failure")  # Log fallback

# Redirect stderr to logger
class StderrToLogger:
    def __init__(self, logger):  # Initialize with logger instance
        self.logger = logger
    def write(self, message):  # Redirect stderr messages to logger
        if message.strip():  # Log non-empty messages
            self.logger.warning(message.strip())
    def flush(self):  # Required for file-like object compatibility
        pass

sys.stderr = StderrToLogger(logger)  # Replace stderr with logger

# Constants for Gwyddion data keys
DATA_KEY = "/%d/data"  # Key for data field
BASE_MIN_KEY = "/%d/base/min"  # Key for minimum color range
BASE_MAX_KEY = "/%d/base/max"  # Key for maximum color range
RANGE_TYPE_KEY = "/%d/base/range-type"  # Key for range type
VISIBLE_KEY = "/%d/data/visible"  # Key for visibility
SELECTION_KEYS = ["/%d/select/rectangle", "/%d/data/selection"]  # Keys for selection types
FILENAME_KEY = "/filename"  # Key for filename
TITLE_KEY = "/%d/data/title"  # Key for data title
ORIGINAL_MIN_KEY = "/%d/base/original_min"  # Key for original minimum
ORIGINAL_MAX_KEY = "/%d/base/original_max"  # Key for original maximum

# State management class to avoid global variables
class PluginState:
    def __init__(self):
        self.macro = []
        self.liststore = gtk.ListStore(int, str, str)
        self.channel_liststore = gtk.ListStore(bool, str, bool, object, int, str, gtk.gdk.Pixbuf, gtk.gdk.Pixbuf)
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
        self.selection_connections = []
        self.timeout_id = None
        self.current_container = None
        self.current_data_id = None
        self.select_all_check = None
        self.select_dropdown = None
        self.last_containers = None
        self.data_browser_timeout_id = None
        self.rename_entry = None  # New field for rename text entry

# Global variable to track GUI instance
_plugin_gui_instance = None
_gui_close_signal = None

# Log parsing functions
def parse_log_entry(entry):  # Parse a single log entry
    try:
        match = re.match(r"proc::(\w+)\((.*?)\)@(.+?)(?:Z|$)", entry)  # Match proc entries
        if not match:
            logger.debug("Skipping non-proc log entry")  # Log invalid entry
            return None
        function, params, time = match.groups()  # Extract function, parameters, and timestamp
        param_string = params.strip()  # Clean parameter string
        param_dict = {}  # Dictionary for parsed parameters
        if param_string:
            for param in re.split(r",\s*(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", param_string):  # Split parameters
                if '=' in param:
                    key, value = param.split('=', 1)  # Split key-value pair
                    key = key.strip()
                    value = value.strip()
                    try:
                        if value.lower() == 'true':
                            param_dict[key] = True  # Parse boolean true
                        elif value.lower() == 'false':
                            param_dict[key] = False  # Parse boolean false
                        elif value.replace('.', '', 1).isdigit():
                            param_dict[key] = float(value) if '.' in value else int(value)  # Parse numbers
                        else:
                            param_dict[key] = value.strip('"')  # Parse strings
                    except Exception:
                        param_dict[key] = value  # Fallback to raw value
        return {"function": function, "parameters": param_dict, "param_string": param_string, "timestamp": time.strip()}
    except Exception:
        return None  # Return None on parsing error

def parse_log_file(file_path):  # Parse entire log file
    log_entries = []  # List to store parsed entries
    try:
        with open(file_path, "r") as f:  # Open log file
            for i, line in enumerate(f):
                parsed = parse_log_entry(line.strip())  # Parse each line
                if parsed:
                    parsed["order"] = i + 1  # Assign order
                    log_entries.append(parsed)
        logger.info("Parsed %d proc entries from %s", len(log_entries), file_path)  # Log success
    except IOError:
        logger.error("Error reading log file %s", file_path)  # Log file error
    return log_entries

def update_macro_view(liststore, macro):  # Update macro table in GUI
    liststore.clear()  # Clear existing entries
    for i, entry in enumerate(macro):
        liststore.append([i + 1, entry["function"], entry["param_string"]])  # Add macro entries

def load_log_file(button, entry, liststore, macro):  # Load log file via GUI
    dialog = gtk.FileChooserDialog("Select Log File", None, gtk.FILE_CHOOSER_ACTION_OPEN,
                                  (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL, gtk.STOCK_OPEN, gtk.RESPONSE_OK))  # Create file chooser
    dialog.set_default_response(gtk.RESPONSE_OK)
    response = dialog.run()
    if response == gtk.RESPONSE_OK:
        file_path = dialog.get_filename()  # Get selected file
        entry.set_text(file_path)
        if file_path and os.path.exists(file_path):
            macro[:] = parse_log_file(file_path)  # Parse and store log entries
            update_macro_view(liststore, macro)  # Update GUI
            if not macro:
                logger.warning("No valid proc entries in %s", file_path)  # Warn if no valid entries
                show_message_dialog(gtk.MESSAGE_WARNING, "No valid processing tools found in the log file.")
        else:
            logger.error("Log file does not exist: %s", file_path)  # Log missing file
            show_message_dialog(gtk.MESSAGE_ERROR, "Log file does not exist: %s" % file_path)
    dialog.destroy()  # Close dialog

def show_message_dialog(msg_type, message, parent=None):  # Display message dialog
    dialog = gtk.MessageDialog(parent=parent, flags=0, type=msg_type, buttons=gtk.BUTTONS_OK, message_format=message)
    dialog.run()
    dialog.destroy()

def show_rename_confirmation_dialog(new_names, parent):
    """Show a dialog to confirm renaming selected channels."""
    message = "The following channels will be renamed:\n\n"
    for old_name, new_name, _, _, _ in new_names:
        message += "%s -> %s\n" % (old_name, new_name)
    dialog = gtk.MessageDialog(
        parent=parent,
        flags=gtk.DIALOG_MODAL,
        type=gtk.MESSAGE_QUESTION,
        buttons=gtk.BUTTONS_OK_CANCEL,
        message_format=message
    )
    dialog.set_title("Confirm Rename")
    response = dialog.run()
    dialog.destroy()
    logger.info("User %s rename operation", "confirmed" if response == gtk.RESPONSE_OK else "cancelled")
    return response == gtk.RESPONSE_OK

def apply_rename(button, channel_liststore, state):
    """Rename selected DataChannels with the exact base name."""
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

    # Prepare new names (same base name for all channels)
    new_names = []
    for container, data_id, title, filename in selected:
        new_name = base_name
        new_names.append((title, new_name, container, data_id, filename))

    # Show confirmation dialog
    if not show_rename_confirmation_dialog(new_names, state.window):
        logger.info("Rename operation cancelled by user")
        return

    # Apply renaming
    def operation(container, data_id, title, filename):
        new_name = next(n for t, n, c, d, f in new_names if c == container and d == data_id)
        container.set_string_by_name(TITLE_KEY % data_id, new_name)
        logger.info("Renamed data_id %d from %s to %s in %s", data_id, title, new_name, filename)

    process_selected_channels(
        channel_liststore,
        operation,
        "No valid channels to rename",
        "Renamed %d channels",
        state
    )
    populate_data_channels(channel_liststore, state)

# GUI creation
def create_gui(state):  # Create main plugin GUI
    global _plugin_gui_instance
    if _plugin_gui_instance is not None:
        logger.debug("GUI already open, closing existing window")
        on_window_delete_event(_plugin_gui_instance, None, state)
        return

    state.window = gtk.Window()  # Initialize main window
    state.window.set_title("AutoProcess")  # Set title
    state.window.set_resizable(True)
    state.window.set_size_request(600, 500)  # Connect close event
    state.window.connect("delete-event", lambda w, e: on_window_delete_event(w, e, state))
    logger.debug("Created main window")
    _plugin_gui_instance = state.window

    vbox = gtk.VBox(spacing=5)  # Main vertical box
    state.window.add(vbox)

    # Fixed Color Range section
    color_range_label = gtk.Label()
    color_range_label.set_markup("<b>Fixed Color Range</b>")
    color_range_label.set_alignment(0, 0.5)
    vbox.pack_start(color_range_label, False, False, 2)
    
    hbox_color_range = gtk.HBox(spacing=5)  # Horizontal box for color range controls
    label_min = gtk.Label("Min:")
    hbox_color_range.pack_start(label_min, False, False, 5)
    state.min_entry = gtk.Entry()  # Entry for min value
    state.min_entry.set_width_chars(10)
    hbox_color_range.pack_start(state.min_entry, False, False, 5)
    label_max = gtk.Label("Max:")
    hbox_color_range.pack_start(label_max, False, False, 5)
    state.max_entry = gtk.Entry()  # Entry for max value
    state.max_entry.set_width_chars(10)
    hbox_color_range.pack_start(state.max_entry, False, False, 5)
    apply_range_button = gtk.Button("Apply Fixed Range")  # Button to apply fixed range
    apply_range_button.connect("clicked", lambda b: apply_fixed_color_range(b, state.channel_liststore, state))
    hbox_color_range.pack_start(apply_range_button, False, False, 1)
    full_range_button = gtk.Button("Full Range")  # Button to set full range
    full_range_button.connect("clicked", lambda b: set_to_full_range(b, state.channel_liststore, state))
    hbox_color_range.pack_start(full_range_button, False, False, 1)
    invert_button = gtk.Button("Invert Mapping")  # Button to invert mapping
    invert_button.connect("clicked", lambda b: invert_mapping(b, state.channel_liststore, state))
    hbox_color_range.pack_start(invert_button, False, False, 1)
    zero_min_button = gtk.Button("Zero to Min")  # Button to set zero to minimum
    zero_min_button.connect("clicked", lambda b: set_zero_to_minimum(b, state.channel_liststore, state))
    hbox_color_range.pack_start(zero_min_button, False, False, 1)
    vbox.pack_start(hbox_color_range, False, False, 2)

    separator1 = gtk.HSeparator()  # Separator
    vbox.pack_start(separator1, False, False, 5)

    # Crop Data and Change Color section
    hbox_section2 = gtk.HBox(spacing=2)
    vbox.pack_start(hbox_section2, False, False, 2)

    # Crop Data subsection
    vbox_crop = gtk.VBox(spacing=5)
    hbox_section2.pack_start(vbox_crop, False, False, 2)
    
    crop_data_label = gtk.Label()
    crop_data_label.set_markup("<b>Crop Data</b>")
    crop_data_label.set_alignment(0, 0.5)
    vbox_crop.pack_start(crop_data_label, False, False, 2)
    
    hbox_crop1 = gtk.HBox(spacing=5)  # Crop coordinates
    label_x = gtk.Label("Origin X (px):")
    label_x.set_size_request(80, -1)
    hbox_crop1.pack_start(label_x, False, False, 5)
    state.x_entry = gtk.Entry()
    state.x_entry.set_width_chars(8)
    state.x_entry.set_size_request(60, -1)
    state.x_entry.set_text("0")
    hbox_crop1.pack_start(state.x_entry, False, False, 5)
    label_y = gtk.Label("Origin Y (px):")
    label_y.set_size_request(80, -1)
    hbox_crop1.pack_start(label_y, False, False, 5)
    state.y_entry = gtk.Entry()
    state.y_entry.set_width_chars(8)
    state.y_entry.set_size_request(60, -1)
    state.y_entry.set_text("0")
    hbox_crop1.pack_start(state.y_entry, False, False, 5)
    vbox_crop.pack_start(hbox_crop1, False, False, 2)
    
    hbox_crop2 = gtk.HBox(spacing=5)  # Crop dimensions
    label_width = gtk.Label("Width (px):")
    label_width.set_size_request(80, -1)
    hbox_crop2.pack_start(label_width, False, False, 5)
    state.width_entry = gtk.Entry()
    state.width_entry.set_width_chars(8)
    state.width_entry.set_size_request(60, -1)
    state.width_entry.set_text("100")
    hbox_crop2.pack_start(state.width_entry, False, False, 5)
    label_height = gtk.Label("Height (px):")
    label_height.set_size_request(80, -1)
    hbox_crop2.pack_start(label_height, False, False, 5)
    state.height_entry = gtk.Entry()
    state.height_entry.set_width_chars(8)
    state.height_entry.set_size_request(60, -1)
    state.height_entry.set_text("100")
    hbox_crop2.pack_start(state.height_entry, False, False, 5)
    vbox_crop.pack_start(hbox_crop2, False, False, 2)
    
    hbox_crop3 = gtk.HBox(spacing=5)  # Crop options
    state.create_new_check = gtk.CheckButton("Create new image")
    state.create_new_check.set_active(False)
    hbox_crop3.pack_start(state.create_new_check, False, False, 5)
    state.keep_offsets_check = gtk.CheckButton("Keep lateral offsets")
    state.keep_offsets_check.set_active(False)
    hbox_crop3.pack_start(state.keep_offsets_check, False, False, 5)
    apply_crop_button = gtk.Button("Apply Crop")
    apply_crop_button.connect("clicked", lambda b: apply_crop(b, state.channel_liststore, state))
    hbox_crop3.pack_start(apply_crop_button, False, False, 1)
    vbox_crop.pack_start(hbox_crop3, False, False, 2)
    
    separator_vertical = gtk.VSeparator()  # Vertical separator
    hbox_section2.pack_start(separator_vertical, False, False, 2)
    
    # Change Color subsection
    vbox_color = gtk.VBox(spacing=5)
    hbox_section2.pack_start(vbox_color, False, False, 2)
    
    change_color_label = gtk.Label()
    change_color_label.set_markup("<b>Change Color</b>")
    change_color_label.set_alignment(0, 0.5)
    vbox_color.pack_start(change_color_label, False, False, 2)
    
    hbox_color1 = gtk.HBox(spacing=5)
    palette_store = gtk.ListStore(str, gtk.gdk.Pixbuf)  # Store for palette names and previews
    for name, pixbuf in get_gradient_names():
        palette_store.append([name, pixbuf])
    state.palette_combobox = gtk.ComboBox(palette_store)
    renderer_text = gtk.CellRendererText()
    state.palette_combobox.pack_start(renderer_text, True)
    state.palette_combobox.add_attribute(renderer_text, "text", 0)
    renderer_pixbuf = gtk.CellRendererPixbuf()
    state.palette_combobox.pack_start(renderer_pixbuf, False)
    state.palette_combobox.add_attribute(renderer_pixbuf, "pixbuf", 1)
    state.palette_combobox.set_active(0)
    hbox_color1.pack_start(state.palette_combobox, False, False, 5)
    vbox_color.pack_start(hbox_color1, False, False, 2)
    
    hbox_color2 = gtk.HBox(spacing=5)
    apply_palette_button = gtk.Button("Apply Palette")
    apply_palette_button.connect("clicked", lambda b: apply_palette(b, state.channel_liststore, state))
    alignment = gtk.Alignment(xalign=0.5, yalign=0.5)
    alignment.add(apply_palette_button)
    hbox_color2.pack_start(alignment, False, False, 5)
    vbox_color.pack_start(hbox_color2, False, False, 2)

    separator2 = gtk.HSeparator()
    vbox.pack_start(separator2, False, False, 5)

    # Data Process Functionalities section
    data_process_label = gtk.Label()
    data_process_label.set_markup("<b>Data Process Functionalities</b>")
    data_process_label.set_alignment(0, 0.5)
    vbox.pack_start(data_process_label, False, False, 2)
    
    vbox_data_process = gtk.VBox(spacing=5)
    hbox_log = gtk.HBox(spacing=5)
    log_entry = gtk.Entry()
    log_entry.set_text("Insert the Log file path")
    log_entry.connect("focus-in-event", lambda w, e: w.set_text("") if w.get_text() == "Insert the Log file path" else None)
    log_entry.connect("focus-out-event", lambda w, e: w.set_text("Insert the Log file path") if not w.get_text() else None)
    hbox_log.pack_start(log_entry, True, True, 5)
    load_button = gtk.Button("Load Log File")
    load_button.connect("clicked", lambda b: load_log_file(b, log_entry, state.liststore, state.macro))
    hbox_log.pack_start(load_button, False, False, 1)
    vbox_data_process.pack_start(hbox_log, False, False, 2)
    vbox.pack_start(vbox_data_process, False, False, 2)
    
    # Macro table
    scrolled_macro = gtk.ScrolledWindow()
    scrolled_macro.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
    treeview_macro = gtk.TreeView(state.liststore)
    renderer_text = gtk.CellRendererText()
    treeview_macro.append_column(gtk.TreeViewColumn("#", renderer_text, text=0))
    treeview_macro.append_column(gtk.TreeViewColumn("Function", renderer_text, text=1))
    treeview_macro.append_column(gtk.TreeViewColumn("Parameters", renderer_text, text=2))
    scrolled_macro.add(treeview_macro)
    scrolled_macro.set_size_request(-1, 100)
    vbox.pack_start(scrolled_macro, True, True, 2)
    
    # Replay button
    replay_button = gtk.Button("Replay Selected Channels")
    replay_button.connect("clicked", lambda b: replay_selected_channels(b, state.channel_liststore, state))
    vbox.pack_start(replay_button, False, False, 1)
    
    separator4 = gtk.HSeparator()
    vbox.pack_start(separator4, False, False, 5)
    
    # Open SPM Files and Rename Files section
    hbox_spm_rename = gtk.HBox(spacing=2)
    vbox.pack_start(hbox_spm_rename, False, False, 2)

    # Open SPM Files subsection
    vbox_spm = gtk.VBox(spacing=5)
    hbox_spm_rename.pack_start(vbox_spm, False, False, 2)
    
    OpenSPM_Files_label = gtk.Label()
    OpenSPM_Files_label.set_markup("<b>Open SPM Files</b>")
    OpenSPM_Files_label.set_alignment(0, 0.5)
    vbox_spm.pack_start(OpenSPM_Files_label, False, False, 2)

    hbox_select = gtk.HBox(spacing=5)
    state.select_all_check = gtk.CheckButton("Select All")
    state.select_all_check.set_active(False)
    
    select_store = gtk.ListStore(str, bool)
    select_store.append(["Select Options", False])
    select_store.append(["First Rows", False])
    select_store.append(["Second Rows", False])
    select_store.append(["Third Rows", False])
    select_store.append(["Fourth Rows", False])
    state.select_dropdown = gtk.ComboBox(select_store)
    renderer_text = gtk.CellRendererText()
    state.select_dropdown.pack_start(renderer_text, True)
    state.select_dropdown.add_attribute(renderer_text, "text", 0)
    state.select_dropdown.set_active(0)
    state.select_dropdown.connect("changed", select_dropdown_changed, state.channel_liststore, select_store)
    state.select_all_check.connect("toggled", sync_select_all_check, state.channel_liststore, select_store)
    
    hbox_select.pack_start(state.select_all_check, False, False, 5)
    hbox_select.pack_start(state.select_dropdown, False, False, 5)
    vbox_spm.pack_start(hbox_select, False, False, 2)

    separator5 = gtk.VSeparator()  # Vertical separator
    hbox_spm_rename.pack_start(separator5, False, False, 2)
    
    # Rename Files subsection
    vbox_rename = gtk.VBox(spacing=5)
    hbox_spm_rename.pack_start(vbox_rename, False, False, 2)
    
    rename_files_label = gtk.Label()
    rename_files_label.set_markup("<b>Rename Files</b>")
    rename_files_label.set_alignment(0, 0.5)
    vbox_rename.pack_start(rename_files_label, False, False, 2)
    
    hbox_rename = gtk.HBox(spacing=5)
    label_base_name = gtk.Label("Base Name:")
    label_base_name.set_size_request(80, -1)
    hbox_rename.pack_start(label_base_name, False, False, 5)
    state.rename_entry = gtk.Entry()
    state.rename_entry.set_width_chars(15)
    state.rename_entry.set_size_request(100, -1)
    state.rename_entry.set_text("")  # Empty by default
    hbox_rename.pack_start(state.rename_entry, False, False, 5)
    apply_rename_button = gtk.Button("Apply Rename")
    apply_rename_button.connect("clicked", lambda b: apply_rename(b, state.channel_liststore, state))
    hbox_rename.pack_start(apply_rename_button, False, False, 1)
    vbox_rename.pack_start(hbox_rename, False, False, 2)

    separator6 = gtk.VSeparator()  # Vertical separator for Save as .gwy
    hbox_spm_rename.pack_start(separator6, False, False, 2)

    # Save as .gwy button
    save_gwy_button = gtk.Button("Save as .gwy")
    save_gwy_button.connect("clicked", lambda b: save_as_gwy(b, state.channel_liststore, state))
    hbox_spm_rename.pack_start(save_gwy_button, False, False, 2)
    
    # SPM File and Channel table
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
    column_delete = gtk.TreeViewColumn("Close File", renderer_delete)
    column_delete.set_cell_data_func(renderer_delete, render_delete_column, treeview_channels)
    treeview_channels.append_column(column_delete)
    treeview_channels.add_events(gtk.gdk.POINTER_MOTION_MASK | gtk.gdk.LEAVE_NOTIFY_MASK)
    treeview_channels.connect("button-press-event", lambda t, e: on_treeview_button_press(t, e, state.channel_liststore, state))
    treeview_channels.connect("motion-notify-event", lambda t, e: on_treeview_motion(t, e, state.channel_liststore))
    treeview_channels.connect("leave-notify-event", on_treeview_leave)
    scrolled_channels.add(treeview_channels)
    scrolled_channels.set_size_request(-1, 200)
    vbox.pack_start(scrolled_channels, True, True, 2)

    populate_data_channels(state.channel_liststore, state)  # Populate channel list
    check_current_selection(state)  # Initialize selection
    state.timeout_id = gtk.timeout_add(500, check_current_selection, state)  # Periodic selection check
    
    state.last_containers = set(id(c) for c in gwy.gwy_app_data_browser_get_containers())  # Track containers
    state.data_browser_timeout_id = gtk.timeout_add(1000, check_data_browser_changes, state.channel_liststore, state)  # Periodic data browser check
    logger.debug("Started periodic data browser check")

    state.window.set_default_size(600, 600)
    state.window.show_all()

import os
import gwy
import gtk
from datetime import datetime

# Store last selected directory
LAST_SAVE_DIR = os.path.expanduser("~/Desktop")

def save_last_dir(save_dir):
    """Save the last selected directory to a file."""
    try:
        with open(os.path.expanduser("~/.gwyddion_last_dir"), "w") as f:
            f.write(save_dir)
        logger.info("Saved last directory: %s", save_dir)
    except Exception as e:
        logger.warning("Failed to save last directory: %s", str(e))

def load_last_dir():
    """Load the last selected directory from a file."""
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
    """Show a dialog to confirm saving .gwy files."""
    message = "The following files will be saved:\n\n"
    for filename, channels, path in save_files:
        message += "%s: %s -> %s\n" % (filename, ", ".join(channels), path)
    dialog = gtk.MessageDialog(
        parent=parent,
        flags=gtk.DIALOG_MODAL,
        type=gtk.MESSAGE_QUESTION,
        buttons=gtk.BUTTONS_OK_CANCEL,
        message_format=message
    )
    dialog.set_title("Confirm Save as .gwy")
    response = dialog.run()
    dialog.destroy()
    logger.info("User %s save as .gwy operation", "confirmed" if response == gtk.RESPONSE_OK else "cancelled")
    return response == gtk.RESPONSE_OK

def get_save_dir(parent):
    """Prompt user to select a single save directory for all SPM files."""
    global LAST_SAVE_DIR
    dialog = gtk.FileChooserDialog(
        title="Select Save Directory for All SPM Files",
        parent=parent,
        action=gtk.FILE_CHOOSER_ACTION_SELECT_FOLDER,
        buttons=(gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL, gtk.STOCK_OK, gtk.RESPONSE_OK)
    )
    dialog.set_current_folder(load_last_dir())
    response = dialog.run()
    if response == gtk.RESPONSE_OK:
        save_dir = dialog.get_filename()
        logger.info("User selected save directory: %s", save_dir)
    else:
        save_dir = os.path.expanduser("~/Desktop")
        logger.info("User cancelled directory selection, using Desktop")
    dialog.destroy()
    if not os.access(save_dir, os.W_OK):
        logger.warning("No write access to %s, falling back to Desktop", save_dir)
        save_dir = os.path.expanduser("~/Desktop")
    LAST_SAVE_DIR = save_dir
    save_last_dir(save_dir)
    return save_dir

def ensure_processing_log(container, data_id, filename, log_file="c:\\users\\allani\\appdata\\local\\temp\\SPM_autoprocess.log"):
    """Ensure processing log is set in the container for a data_id."""
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
                            crop_line = next((l for l in lines if "Cropped in place data_id %d" % data_id in l), None)
                            if crop_line:
                                crop_params = next((l for l in lines if "tool::GwyToolCrop" in l and "data_id %d" % data_id in l), None)
                                if crop_params:
                                    log_entries.append(crop_params.strip())
        log_value = "\n".join(log_entries) if log_entries else None
        if log_value:
            container.set_string_by_name("/%d/log" % data_id, log_value)
            logger.info("Set processing log for data_id %d in %s", data_id, filename)
        else:
            logger.warning("No processing log constructed for data_id %d in %s", data_id, filename)
    except Exception as e:
        logger.warning("Failed to set log for data_id %d in %s: %s", data_id, filename, str(e))

def ensure_color_range(container, data_id, filename):
    """Ensure color range is set in the container for a data_id."""
    try:
        data_field = container.get_object_by_name("/%d/data" % data_id)
        if not container.contains_by_name("/%d/base/range" % data_id):
            min_val, max_val = gwy.gwy_data_field_get_min_max(data_field)
            container.set_value_by_name("/%d/base/range" % data_id, (min_val, max_val))
            logger.info("Set fallback color range for data_id %d in %s: min=%f, max=%f", data_id, filename, min_val, max_val)
        if not container.contains_by_name("/%d/base/range-type" % data_id):
            container.set_int32_by_name("/%d/base/range-type" % data_id, 1)  # GWY_LAYER_RANGE_FIXED
            logger.info("Set fixed color range type for data_id %d in %s", data_id, filename)
    except Exception as e:
        logger.warning("Failed to set color range for data_id %d in %s: %s", data_id, filename, str(e))

def save_as_gwy(button, channel_liststore, state):
    """Save selected DataChannels as .gwy files named after their SPM files."""
    DATA_KEY = "/%d/data"
    TITLE_KEY = "/%d/data/title"
    SHOW_KEY = "/%d/data/visible"
    PALETTE_KEY = "/%d/base/palette"
    LOG_KEY = "/%d/log"
    RANGE_KEY = "/%d/base/range"
    RANGE_TYPE_KEY = "/%d/base/range-type"

    selected = []
    seen_ids = set()
    for row in channel_liststore:
        checked, title, is_channel, container, data_id, filename, _, _ = row
        if checked and container and is_channel and data_id != -1:
            unique_key = (filename, data_id)
            if unique_key not in seen_ids:
                logger.info("Processing channel: title=%s, data_id=%d, filename=%s", title, data_id, filename)
                selected.append((container, data_id, title, filename))
                seen_ids.add(unique_key)

    if not selected:
        logger.error("No channels selected for saving")
        show_message_dialog(gtk.MESSAGE_ERROR, "No channels selected for saving")
        return

    # Group channels by SPM file (filename)
    file_groups = {}
    for container, data_id, title, filename in selected:
        if filename not in file_groups:
            file_groups[filename] = []
        file_groups[filename].append((container, data_id, title))

    if not file_groups:
        logger.error("No valid SPM files found for selected channels")
        show_message_dialog(gtk.MESSAGE_ERROR, "No valid SPM files found for saving")
        return

    # Get single save directory for all files
    save_dir = get_save_dir(state.window)
    logger.info("Using save directory for all files: %s", save_dir)

    # Prepare save paths
    save_files = []
    for filename, channels in file_groups.items():
        base_name = os.path.splitext(os.path.basename(filename))[0]
        save_path = os.path.join(save_dir, "%s.gwy" % base_name)
        counter = 1
        while os.path.exists(save_path):
            save_path = os.path.join(save_dir, "%s_processed_%d.gwy" % (base_name, counter))
            counter += 1
        save_files.append((base_name, [title for _, _, title in channels], save_path))

    # Show confirmation dialog
    if not show_save_confirmation_dialog(save_files, state.window):
        logger.info("Save as .gwy operation cancelled by user")
        return

    # Save each group as a .gwy file
    def save_group(filename, channels, save_path):
        logger.info("Attempting to save %d channels to %s", len(channels), save_path)
        # Use the first container as the base, as it contains all channels
        container = channels[0][0]
        success = True
        # Ensure logs and color ranges are set for each channel
        for _, data_id, title in channels:
            try:
                if not container.contains_by_name(DATA_KEY % data_id):
                    logger.error("No data field found for data_id %d (%s) in %s", data_id, title, filename)
                    success = False
                    continue
                # Ensure processing log
                ensure_processing_log(container, data_id, filename)
                # Ensure color range
                ensure_color_range(container, data_id, filename)
                logger.info("Prepared data_id %d (%s) for %s", data_id, title, save_path)
            except Exception as e:
                logger.error("Failed to prepare data_id %d (%s) for %s: %s", data_id, title, save_path, str(e))
                success = False
        try:
            # Try gwy_file_save first
            operation = gwy.gwy_file_save(container, save_path, gwy.RUN_NONINTERACTIVE)
            if operation == 0:
                logger.warning("gwy_file_save failed, falling back to gwy_file_func_run_save")
                # Fallback to explicit .gwy save
                success = gwy.gwy_file_func_run_save("gwyddion", container, save_path, gwy.RUN_NONINTERACTIVE)
                if not success:
                    logger.error("Failed to save %s using gwy_file_func_run_save", save_path)
                    show_message_dialog(gtk.MESSAGE_ERROR, "Failed to save %s" % save_path)
                    return False
            if not os.path.exists(save_path):
                logger.error("File %s was not created", save_path)
                show_message_dialog(gtk.MESSAGE_ERROR, "File %s was not created" % save_path)
                return False
            logger.info("Saved %d channels to %s", len(channels), save_path)
            return success
        except Exception as e:
            logger.error("Failed to save %s: %s", save_path, str(e))
            show_message_dialog(gtk.MESSAGE_ERROR, "Failed to save %s: %s" % (save_path, str(e)))
            return False

    # Process each file group and count successes
    success_count = 0
    for filename, channels in file_groups.items():
        save_path = next(path for fname, _, path in save_files if fname == os.path.splitext(os.path.basename(filename))[0])
        if save_group(filename, channels, save_path):
            success_count += len(channels)

    if success_count == 0:
        logger.error("No items successfully processed")
        show_message_dialog(gtk.MESSAGE_ERROR, "No items successfully processed")
    else:
        logger.info("Saved %d channels as .gwy files", success_count)
        show_message_dialog(gtk.MESSAGE_INFO, "Saved %d channels as .gwy files" % success_count)

    populate_data_channels(channel_liststore, state)

# window close handler
def on_window_delete_event(widget, event, state):  # Handle window close
    # Clear GUI open flag
    container = gwy.gwy_app_data_browser_get_containers()[0]
    gui_key = "/module/autoprocess/gui_open"
    if container.contains_by_name(gui_key):
        container.set_boolean_by_name(gui_key, False)
        logger.debug("Cleared GUI open flag in container")
    
    if state.timeout_id is not None:
        gobject.source_remove(state.timeout_id)  # Remove selection timeout
        state.timeout_id = None
        logger.debug("Removed selection timeout handler")
    
    if hasattr(state, 'data_browser_timeout_id') and state.data_browser_timeout_id is not None:
        gobject.source_remove(state.data_browser_timeout_id)  # Remove data browser timeout
        state.data_browser_timeout_id = None
        logger.debug("Removed data browser timeout handler")
    
    for conn_id, container, data_id in state.selection_connections:
        try:
            for key in [SELECTION_KEYS[0] % data_id, SELECTION_KEYS[1] % data_id]:
                if container and container.contains_by_name(key):
                    selection = container.get_object_by_name(key)
                    selection.disconnect(conn_id)  # Disconnect selection signals
        except:
            pass
    
    state.selection_connections = []
    state.last_crop_operation = []  # Clear crop history
    state.current_container = None
    state.current_data_id = None
    
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter("%(asctime)s,%(msecs)03d: %(message)s", datefmt='%Y-%m-%d %H:%M:%S'))
    logger.addHandler(console_handler)
    for handler in logger.handlers[:]:
        if handler != console_handler:
            logger.removeHandler(handler)  # Remove logger handlers
    logger.removeHandler(console_handler)
    
    state.window.destroy()  # Destroy the window
    state.window = None
    logger.debug("GUI instance closed and cleared")
    return True  # Prevent default GTK behavior

# Channel and file management
def toggle_channel_selection(cell, path, channel_liststore):  # Toggle channel checkbox
    if channel_liststore[path][2]:
        channel_liststore[path][0] = not channel_liststore[path][0]
        logger.debug("Toggled %s to %s", channel_liststore[path][1], channel_liststore[path][0])

def select_all_channels(button, channel_liststore, select=True):  # Select/deselect all channels
    def set_selection(model, path, iter, select):
        if model.get_value(iter, 2):
            model.set_value(iter, 0, select)
    channel_liststore.foreach(set_selection, select)

def delete_file(cell, path, channel_liststore, state):  # Delete SPM file
    container = channel_liststore[path][3]
    if container and channel_liststore[path][4] == -1 and not channel_liststore[path][2]:
        filename = channel_liststore[path][5]
        logger.info("Attempting to delete SPM file: %s", filename)
        try:
            gwy.gwy_app_data_browser_remove(container)  # Remove from Gwyddion
            populate_data_channels(channel_liststore, state)  # Refresh list
        except Exception:
            logger.error("Failed to delete SPM file %s", filename)

def remove_file_from_list(cell, path, channel_liststore):  # Remove file from list
    container = channel_liststore[path][3]
    filename = channel_liststore[path][5]
    if container and channel_liststore[path][4] == -1 and not channel_liststore[path][2] and channel_liststore[path][1] != "──────────────────":
        logger.info("Attempting to remove SPM file from list: %s", filename)
        try:
            iter_to_remove = channel_liststore.get_iter(path)
            while iter_to_remove is not None:
                if (channel_liststore[iter_to_remove][3] == container or 
                    channel_liststore[iter_to_remove][1] == "──────────────────"):
                    next_iter = channel_liststore.iter_next(iter_to_remove)
                    channel_liststore.remove(iter_to_remove)
                    iter_to_remove = next_iter
                else:
                    break
            logger.info("Removed SPM file %s from list", filename)
        except Exception as e:
            logger.error("Failed to remove SPM file %s from list: %s", filename, str(e))

def create_pixbuf(stock_id, fallback_color):  # Create pixbuf for icons
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

def populate_data_channels(channel_liststore, state):  # Populate channel list
    checkbox_states = {}  # Store checkbox states
    for row in channel_liststore:
        container, data_id, filename = row[3], row[4], row[5]
        if container and data_id != -1:
            key = (id(container), data_id)
            checkbox_states[key] = row[0]
        elif container and data_id == -1 and row[1] != "──────────────────":
            key = (id(container), -1)
            checkbox_states[key] = row[0]

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

    channel_liststore.clear()
    delete_pixbuf = create_pixbuf(gtk.STOCK_CLOSE, 0xff0000ff)  # Red icon for delete
    remove_pixbuf = create_pixbuf(gtk.STOCK_REMOVE, 0xffa500ff)  # Orange icon for remove
    containers = gwy.gwy_app_data_browser_get_containers()
    for idx, container in enumerate(containers, 1):
        filename = container.get_string_by_name(FILENAME_KEY) or "Container %d" % id(container)
        filename = os.path.basename(filename) if filename else "Unknown SPM File"
        file_key = (id(container), -1)
        file_checked = checkbox_states.get(file_key, False)
        channel_liststore.append([file_checked, "<b>File%d: %s</b>" % (idx, filename), False, container, -1, filename, delete_pixbuf, remove_pixbuf])
        for data_id in gwy.gwy_app_data_browser_get_data_ids(container):
            title = container.get_string_by_name(TITLE_KEY % data_id) or "Data %d" % data_id
            channel_key = (id(container), data_id)
            channel_checked = checkbox_states.get(channel_key, False)
            channel_liststore.append([channel_checked, "  %s" % title, True, container, data_id, filename, None, None])
            for selection_key in [SELECTION_KEYS[0] % data_id, SELECTION_KEYS[1] % data_id]:
                if container.contains_by_name(selection_key):
                    selection = container.get_object_by_name(selection_key)
                    try:
                        conn_id = selection.connect("changed", selection_changed, container, data_id, state)
                        state.selection_connections.append((conn_id, container, data_id))
                        logger.debug("Connected selection signal for data_id %d", data_id)
                    except Exception as e:
                        logger.error("Failed to connect selection signal for data_id %d: %s", data_id, str(e))
        channel_liststore.append([False, "──────────────────", False, None, -1, "", None, None])
    logger.info("Populated %d data channels from %d SPM files", sum(len(gwy.gwy_app_data_browser_get_data_ids(c)) for c in containers), len(containers))

# Selection handling
def get_selection_params(container, data_id):  # Get crop parameters from selection
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
                    logger.debug("Normalized selection for data_id %d: x=%d, y=%d, width=%d, height=%d", data_id, x, y, width, height)
                    return x, y, width, height
            except Exception as e:
                logger.error("Failed to process selection for data_id %d: %s", data_id, str(e))
        else:
            logger.debug("No selection found for data_id %d at %s", data_id, selection_key)
        return None, None, None, None
    except Exception as e:
        logger.error("Failed to get selection for data_id %d: %s", data_id, str(e))
        return None, None, None, None

def selection_changed(selection, index, container, data_id, state, *args):  # Handle selection change
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
            logger.debug("Dynamic selection update for data_id %d: x=%d, y=%d, width=%d, height=%d", data_id, x, y, width, height)
        else:
            state.x_entry.set_text("")
            state.y_entry.set_text("")
            state.width_entry.set_text("")
            state.height_entry.set_text("")
            logger.debug("Cleared selection fields for data_id %d due to no valid selection", data_id)
    except Exception as e:
        logger.error("Error in selection_changed for data_id %d: %s", data_id, str(e))
        if state.window is not None:
            state.x_entry.set_text("")
            state.y_entry.set_text("")
            state.width_entry.set_text("")
            state.height_entry.set_text("")

def check_current_selection(state):  # Check and update current selection
    if not gwy.gwy_app_data_browser_get_containers():
        return True

    current_container = gwy.gwy_app_data_browser_get_current(gwy.APP_CONTAINER)
    current_data_id = gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_ID) if current_container else None

    if (current_container, current_data_id) != (state.current_container, state.current_data_id):
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

            for key in ["/%d/select/pointer" % current_data_id, "/%d/select/line" % current_data_id]:
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
                conn_id = selection.connect("changed", selection_changed, current_container, current_data_id, state)
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
                if (state.x_entry.get_text().strip() != str(x) or state.y_entry.get_text().strip() != str(y) or
                    state.width_entry.get_text().strip() != str(width) or state.height_entry.get_text().strip() != str(height)):
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

def data_browser_changed(obj, arg, channel_liststore, state):  # Handle data browser changes
    logger.debug("Data browser changed, updating channel list")
    populate_data_channels(channel_liststore, state)

def check_data_browser_changes(channel_liststore, state):  # Periodically check data browser
    current_containers = gwy.gwy_app_data_browser_get_containers()
    if not current_containers and state.window is not None:
        logger.debug("No containers in data browser, Gwyddion likely closed; shutting down GUI")
        on_window_delete_event(state.window, None, state)
        gtk.main_quit()
        return False
    
    current_container_ids = set(id(c) for c in current_containers)
    if not hasattr(state, 'last_containers') or state.last_containers != current_container_ids or \
       len(current_containers) != len(state.last_containers):
        logger.debug("Data browser containers changed or count mismatch, updating channel list")
        populate_data_channels(channel_liststore, state)
        state.last_containers = current_container_ids
    return True

# Data processing
def get_min_max(container, data_id):  # Get min/max values for channel or file
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
            return data_field.get_min(), data_field.get_max() if data_field else (None, None)
    except Exception:
        return None, None

def validate_crop_params(data_field, x, y, width, height, filename, spm_filename):  # Validate crop parameters
    xres, yres = data_field.get_xres(), data_field.get_yres()
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        return False, "Invalid crop parameters for %s in %s" % (filename, spm_filename)
    if x + width > xres or y + height > yres:
        return False, "Crop area out of bounds for %s in %s: x=%d, y=%d, width=%d, height=%d" % (filename, spm_filename, x, y, width, yres)
    return True, None

def process_selected_channels(channel_liststore, operation, no_selection_msg, success_msg, state):  # Process selected channels
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
            #show_message_dialog(gtk.MESSAGE_ERROR, "Failed to process %s, data_id %d: %s" % (filename, data_id, str(e)))
    
    if success_count > 0:
        logger.info(success_msg % success_count)
        show_message_dialog(gtk.MESSAGE_INFO, success_msg % success_count)
    else:
        logger.error("No items successfully processed")
        show_message_dialog(gtk.MESSAGE_ERROR, "No items successfully processed")

def apply_palette(button, channel_liststore, state):  # Apply selected palette
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
        logger.info("Set palette %s on data_id %d (%s) in %s", palette_name, data_id, title, filename)
    
    process_selected_channels(channel_liststore, operation, "No channels selected for palette change",
                             "Palette %s applied to %%d channels" % palette_name, state)

def apply_fixed_color_range(button, channel_liststore, state):  # Apply fixed color range
    try:
        min_val = float(state.min_entry.get_text().strip())
        max_val = float(state.max_entry.get_text().strip())
        if min_val >= max_val:
            show_message_dialog(gtk.MESSAGE_ERROR, "Invalid range: Minimum value must be less than maximum value.You Can Invert Mapping instead")
            return
        user_provided = True
    except ValueError:
        user_provided = False
    
    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        min_val_local, max_val_local = (min_val, max_val) if user_provided else get_min_max(container, data_id)
        if min_val_local is None or max_val_local is None:
            raise ValueError("No valid min/max")
        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FIXED)
        container.set_double_by_name(BASE_MIN_KEY % data_id, min_val_local)
        container.set_double_by_name(BASE_MAX_KEY % data_id, max_val_local)
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
    
    process_selected_channels(channel_liststore, operation, "No channels selected for color range",
                             "Fixed color range applied to %d channels", state)

def set_to_full_range(button, channel_liststore, state):  # Set to full range
    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            raise ValueError("No data field")
        if container.contains_by_name(ORIGINAL_MIN_KEY % data_id) and container.contains_by_name(ORIGINAL_MAX_KEY % data_id):
            original_min = container.get_double_by_name(ORIGINAL_MIN_KEY % data_id)
            current_min = data_field.get_min()
            if original_min != current_min:  
                data_field.add(original_min - current_min)  
                data_field.data_changed()
            container.remove_by_name(ORIGINAL_MIN_KEY % data_id)
            container.remove_by_name(ORIGINAL_MAX_KEY % data_id)
            logger.info("Restored original min=%g for data_id %d in %s", original_min, data_id, filename)
        # Reset to full range
        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FULL)
        if container.contains_by_name(BASE_MIN_KEY % data_id):
            container.remove_by_name(BASE_MIN_KEY % data_id)
        if container.contains_by_name(BASE_MAX_KEY % data_id):
            container.remove_by_name(BASE_MAX_KEY % data_id)
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        current_data_id = gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_ID) if container == gwy.gwy_app_data_browser_get_current(gwy.APP_CONTAINER) else None
        if current_data_id == data_id:
            min_val, max_val = data_field.get_min(), data_field.get_max()
            state.min_entry.set_text("%.6g" % min_val if min_val is not None else "")
            state.max_entry.set_text("%.6g" % max_val if max_val is not None else "")
        logger.info("Set full range for data_id %d in %s", data_id, filename)
    
    process_selected_channels(channel_liststore, operation, "No channels selected for full range",
                             "Full range applied to %d channels", state)

def invert_mapping(button, channel_liststore, state):  # Invert color mapping
    def operation(container, data_id, title, filename):
        if data_id == -1:
            raise ValueError("Invalid channel")
        data_field = container.get_object_by_name(DATA_KEY % data_id)
        if not data_field:
            raise ValueError("No data field")
        current_min = container.get_double_by_name(BASE_MIN_KEY % data_id) if container.contains_by_name(BASE_MIN_KEY % data_id) else data_field.get_min()
        current_max = container.get_double_by_name(BASE_MAX_KEY % data_id) if container.contains_by_name(BASE_MAX_KEY % data_id) else data_field.get_max()
        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FIXED)
        container.set_double_by_name(BASE_MIN_KEY % data_id, current_max)
        container.set_double_by_name(BASE_MAX_KEY % data_id, current_min)
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        logger.info("Inverted color range for data_id %d in %s", data_id, filename)
    
    process_selected_channels(channel_liststore, operation, "No channels selected for invert mapping",
                             "Color range inverted for %d channels", state)

def set_zero_to_minimum(button, channel_liststore, state):  # Set minimum to zero
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
        data_field.add(-current_min)  # Shift data to set minimum to zero
        data_field.data_changed()
        container.set_int32_by_name(RANGE_TYPE_KEY % data_id, gwy.LAYER_BASIC_RANGE_FIXED)
        container.set_double_by_name(BASE_MIN_KEY % data_id, 0.0)
        container.set_double_by_name(BASE_MAX_KEY % data_id, current_max - current_min)
        gwy.gwy_app_data_browser_select_data_field(container, data_id)
        current_data_id = gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_ID) if container == gwy.gwy_app_data_browser_get_current(gwy.APP_CONTAINER) else None
        if current_data_id == data_id:
            state.min_entry.set_text("0")
            state.max_entry.set_text("%.6g" % (current_max - current_min))
        logger.info("Set zero to minimum for data_id %d in %s, stored original min=%g, max=%g", data_id, filename, current_min, current_max)
    
    process_selected_channels(channel_liststore, operation, "No channels selected for set zero to minimum",
                             "Zero to minimum applied to %d channels", state)

def apply_crop(button, channel_liststore, state):
    try:
        x = int(state.x_entry.get_text().strip())
        y = int(state.y_entry.get_text().strip())
        width = int(state.width_entry.get_text().strip())
        height = int(state.height_entry.get_text().strip())
        create_new = state.create_new_check.get_active()
        keep_offsets = state.keep_offsets_check.get_active()
    except ValueError:
        logger.error("Invalid crop parameters")
        show_message_dialog(gtk.MESSAGE_ERROR, "Invalid crop parameters. Please enter valid integer values.")
        return

    # Scan selected channels for dimension compatibility
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
        if data_id == -1:
            data_ids = gwy.gwy_app_data_browser_get_data_ids(container)
        else:
            data_ids = [data_id]
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
        response = show_crop_conflict_dialog(invalid_channels, valid_channels, channel_liststore, state, x, y, width, height, create_new, keep_offsets)
        if response in ["cancel", "cancel_list"]:
            logger.info("Crop operation cancelled by user")
            return
        # Proceed with valid channels only
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
    data_field = container.get_object_by_name(DATA_KEY % data_id)
    if not data_field:
        raise ValueError("No data field for data_id %d" % data_id)
    valid, error_msg = validate_crop_params(data_field, x, y, width, height, title, filename)
    if not valid:
        raise ValueError(error_msg)
    log_entry = "tool::GwyToolCrop(all=%s, hold_selection=4, keep_offsets=%s, new_channel=%s, x=%d, y=%d, width=%d, height=%d)@%s" % (
        str(data_id == -1), str(keep_offsets), str(create_new), x, y, width, height, datetime.now().isoformat())
    logger.info(log_entry)
    log_key = "/%d/log" % data_id
    current_log = container.get_string_by_name(log_key) or ""
    container.set_string_by_name(log_key, current_log + log_entry + "\n")
    logger.debug("Manually added log entry to %s for data_id %d", log_key, data_id)
    new_id = None
    if create_new:
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
        data_field.resize(x, y, x + width, y + height)
        data_field.data_changed()
        if container:
            gwy.gwy_app_data_browser_select_data_field(container, data_id)
        logger.info("Cropped in place data_id %d in %s", data_id, filename)

def show_crop_conflict_dialog(invalid_channels, valid_channels, channel_liststore, state, x, y, width, height, create_new, keep_offsets):
    total_channels = len(invalid_channels) + len(valid_channels)
    message = "%d out of %d selected DataChannels cannot be processed. Proceed without them?" % (len(invalid_channels), total_channels)
    dialog = gtk.MessageDialog(
        parent=state.window,
        flags=gtk.DIALOG_MODAL,
        type=gtk.MESSAGE_WARNING,
        buttons=gtk.BUTTONS_NONE,
        message_format=message
    )
    dialog.add_button("Cancel", gtk.RESPONSE_CANCEL)
    dialog.add_button("Proceed", gtk.RESPONSE_OK)
    dialog.add_button("Cancel and list conflicts", gtk.RESPONSE_REJECT)
    dialog.add_button("Proceed and list conflicts", gtk.RESPONSE_APPLY)
    dialog.set_default_response(gtk.RESPONSE_CANCEL)
    response = dialog.run()
    dialog.destroy()

    response_map = {
        gtk.RESPONSE_CANCEL: "cancel",
        gtk.RESPONSE_OK: "proceed",
        gtk.RESPONSE_REJECT: "cancel_list",
        gtk.RESPONSE_APPLY: "proceed_list"
    }
    response_str = response_map.get(response, "cancel")

    if response_str in ["cancel_list", "proceed_list"]:
        show_conflict_list_dialog(invalid_channels, state.window)

    logger.info("User selected %s for crop conflict dialog", response_str)
    return response_str

def show_conflict_list_dialog(invalid_channels, parent):
    dialog = gtk.Dialog(
        title="Crop Conflicts",
        parent=parent,
        flags=gtk.DIALOG_MODAL,
        buttons=(gtk.STOCK_OK, gtk.RESPONSE_OK)
    )
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

def replay_selected_channels(button, channel_liststore, state):  # Replay macro on channels
    if not state.macro:
        logger.error("No tools in macro to replay")
        show_message_dialog(gtk.MESSAGE_ERROR, "No tools in macro to replay. Please load a log file.")
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
    
    process_selected_channels(channel_liststore, operation, "No channels selected for replay",
                             "Macro replay completed on %d channels", state)

# Utility
def get_gradient_names():  # Retrieve Gwyddion gradients
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

def render_channel_column(column, cell, model, iter, treeview):  # Render Select column
    is_selectable = model.get_value(iter, 2)
    is_file_row = not is_selectable and model.get_value(iter, 4) == -1 and model.get_value(iter, 1) != "──────────────────"
    path = model.get_path(iter)
    select_hover_path = treeview.get_data("select_hover_path")

    if is_selectable:
        if isinstance(cell, gtk.CellRendererToggle):
            cell.set_property("visible", True)
            cell.set_property("active", model.get_value(iter, 0))
        elif isinstance(cell, gtk.CellRendererText):
            cell.set_property("visible", False)
    elif is_file_row:
        if isinstance(cell, gtk.CellRendererToggle):
            cell.set_property("visible", False)
        elif isinstance(cell, gtk.CellRendererText):
            cell.set_property("visible", True)
            cell.set_property("text", "–")
            cell.set_property("foreground", "red" if select_hover_path == path else "black")
            cell.set_property("underline", pango.UNDERLINE_SINGLE if select_hover_path == path else pango.UNDERLINE_NONE)
    else:
        if isinstance(cell, gtk.CellRendererToggle):
            cell.set_property("visible", False)
        elif isinstance(cell, gtk.CellRendererText):
            cell.set_property("visible", False)

def render_delete_column(column, cell, model, iter, treeview):  # Render Close File column
    is_file_row = not model.get_value(iter, 2) and model.get_value(iter, 4) == -1 and model.get_value(iter, 1) != "──────────────────"
    path = model.get_path(iter)
    close_hover_path = treeview.get_data("close_hover_path")

    if is_file_row:
        cell.set_property("visible", True)
        cell.set_property("text", "X")
        cell.set_property("weight", pango.WEIGHT_BOLD)
        cell.set_property("foreground", "blue" if close_hover_path == path else "black")
    else:
        cell.set_property("visible", False)

def render_remove_column(column, cell, model, iter):  # Render remove button
    is_file_row = not model.get_value(iter, 2) and model.get_value(iter, 4) == -1 and model.get_value(iter, 1) != "──────────────────"
    cell.set_property("visible", is_file_row)
    if is_file_row:
        pixbuf = model.get_value(iter, 7)
        if pixbuf and isinstance(pixbuf, gtk.gdk.Pixbuf):
            cell.set_property("pixbuf", pixbuf)

def on_treeview_button_press(treeview, event, channel_liststore, state):  # Handle TreeView clicks
    if event.button == 1:
        pos = treeview.get_path_at_pos(int(event.x), int(event.y))
        if pos:
            path, column, cell_x, cell_y = pos
            if column == treeview.get_column(0):
                if channel_liststore[path][2]:
                    toggle_channel_selection(None, path, channel_liststore)
                    return True
                elif not channel_liststore[path][2] and channel_liststore[path][4] == -1 and channel_liststore[path][1] != "──────────────────":
                    remove_file_from_list(None, path, channel_liststore)
                    return True
            elif channel_liststore[path][2]:
                container, data_id = channel_liststore[path][3], channel_liststore[path][4]
                if data_id != -1:
                    gwy.gwy_app_data_browser_select_data_field(container, data_id)
                    min_val = container.get_double_by_name(BASE_MIN_KEY % data_id) if container.contains_by_name(BASE_MIN_KEY % data_id) else None
                    max_val = container.get_double_by_name(BASE_MAX_KEY % data_id) if container.contains_by_name(BASE_MAX_KEY % data_id) else None
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
            if column == treeview.get_column(2) and channel_liststore[path][4] == -1 and not channel_liststore[path][2]:
                delete_file(None, path, channel_liststore, state)
                return True
    return False

def on_treeview_motion(treeview, event, channel_liststore):  # Handle mouse motion for hover effects
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

def on_treeview_leave(treeview, event):  # Clear hover states on mouse leave
    old_select_hover_path = treeview.get_data("select_hover_path")
    old_close_hover_path = treeview.get_data("close_hover_path")
    if old_select_hover_path or old_close_hover_path:
        treeview.set_data("select_hover_path", None)
        treeview.set_data("close_hover_path", None)
        treeview.queue_draw()
    return True

def select_dropdown_changed(combo, channel_liststore, select_store):  # Handle dropdown selection
    active = combo.get_active()
    if active == 0:
        return
    
    row_index = active - 1
    new_state = not select_store[active][1]
    select_store[active][1] = new_state
    
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
                logger.debug("%s row %d for file %s", "Selected" if new_state else "Deselected", row_index + 1, row[5])
    
    combo.set_active(0)

def sync_select_all_check(checkbutton, channel_liststore, select_store):  # Sync Select All checkbox
    select_all_state = checkbutton.get_active()
    
    for i, row in enumerate(select_store):
        if i == 0:
            continue
        row[1] = select_all_state
    
    for row in channel_liststore:
        if row[2]:
            row[0] = select_all_state
    
    logger.debug("%s all channels and dropdown states", "Selected" if select_all_state else "Deselected")

#plugin entry point
def run(data, mode):  # Main plugin entry point
    global _gui_close_signal
    # Check for existing GUI in Gwyddion's data browser
    container = gwy.gwy_app_data_browser_get_containers()[0]
    gui_key = "/module/autoprocess/gui_open"
    if container.contains_by_name(gui_key) and container.get_boolean_by_name(gui_key):
        logger.debug("GUI already open, showing message dialog and emitting close signal")
        # Show message dialog
        dialog = gtk.MessageDialog(
            parent=None,
            flags=gtk.DIALOG_MODAL,
            type=gtk.MESSAGE_INFO,
            buttons=gtk.BUTTONS_OK,
            message_format="AutoProcess GUI is already open."
        )
        dialog.run()
        dialog.destroy()
        if _gui_close_signal is not None:
            gobject.signal_emit(_gui_close_signal, "close-gui")
        return  # Exit without creating new GUI

    state = PluginState()
    key = gwy.gwy_app_data_browser_get_current(gwy.APP_DATA_FIELD_KEY)
    gwy.gwy_app_undo_qcheckpoint(data, [key])  # Create undo checkpoint
    create_gui(state)  # Initialize GUI
    # Set GUI open flag
    container.set_boolean_by_name(gui_key, True)
    logger.debug("Set GUI open flag in container")
    # Register close signal
    if _gui_close_signal is None:
        _gui_close_signal = gobject.signal_new("close-gui", gtk.Window, gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ())
    state.window.connect("close-gui", lambda w: on_window_delete_event(w, None, state))
