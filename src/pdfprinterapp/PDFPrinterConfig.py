import sys
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import win32print
import logging
from PIL import Image  # Import PIL to handle icon loading for the tray

# Setup logging
logging.basicConfig(level=logging.INFO, filename='config_manager.log', format='%(asctime)s %(levelname)s:%(message)s')

CONFIG_FILE = "config/config.json"

def load_config():
    """
    Load or create the config file. If the config file is missing, default settings are used.
    """
    default_config = {
        "pdf_folder": os.path.expanduser("~"),  # Default to user home folder
        "default_printer": win32print.GetDefaultPrinter()
    }

    # Ensure the config directory exists
    os.makedirs("config", exist_ok=True)

    try:
        # Check if config.json exists
        if not os.path.exists(CONFIG_FILE):
            # Create config.json with default settings
            with open(CONFIG_FILE, "w") as f:
                json.dump(default_config, f, indent=4)
            logging.info("Created new config file with default settings.")
            return default_config

        # Load existing config from file
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
        logging.info("Config file loaded successfully.")
        return config

    except Exception as e:
        logging.error(f"Error loading config file: {str(e)}")
        messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")
        return default_config  # Fallback to default settings in case of failure


def save_config(updated_config):
    """
    Save the updated configuration to the config file.
    """
    try:
        # Write config to a temporary file and rename it for atomic writes
        temp_file = CONFIG_FILE + ".tmp"
        with open(temp_file, "w") as f:
            json.dump(updated_config, f, indent=4)

        # Safely replace the original file
        os.replace(temp_file, CONFIG_FILE)
        logging.info("Configuration saved successfully.")
        messagebox.showinfo("Success", "Settings have been saved successfully!")

    except Exception as e:
        logging.error(f"Failed to save config file: {str(e)}")
        messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")


def select_pdf_folder():
    """
    Open a folder dialog for the user to select a folder for saving PDFs.
    """
    folder = filedialog.askdirectory()
    if folder:
        folder_label.config(text=folder)
        config = load_config()
        config['pdf_folder'] = folder
        save_config(config)


def select_default_printer():
    """
    Open a dialog to let the user select a default printer from the available printers.
    """
    try:
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        printer_names = [printer[2] for printer in printers]

        if not printer_names:
            raise ValueError("No printers available")

        # Popup window to select the printer
        printer_window = tk.Toplevel(root)
        printer_window.title("Select Default Printer")
        printer_var = tk.StringVar(printer_window)

        # Create a dropdown to select a printer
        tk.Label(printer_window, text="Select Printer:").pack(pady=10)
        printer_menu = tk.OptionMenu(printer_window, printer_var, *printer_names)
        printer_menu.pack(pady=10)

        def set_printer():
            selected_printer = printer_var.get()
            if selected_printer:
                config = load_config()
                config['default_printer'] = selected_printer
                save_config(config)
                printer_window.quit()  # Close the window after setting the printer

        tk.Button(printer_window, text="Set Default Printer", command=set_printer).pack(pady=10)

    except Exception as e:
        logging.error(f"Failed to list printers: {str(e)}")
        messagebox.showerror("Error", f"Failed to list printers: {str(e)}")


# Function to load custom icon from file (supports PyInstaller)
def load_icon(icon_name):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS  # PyInstaller extracts files to this temp directory
    else:
        base_path = os.path.abspath(".")  # Running from script directly

    icon_path = os.path.join(base_path, icon_name)
    return Image.open(icon_path)

# Create the GUI window using Tkinter
root = tk.Tk()
root.title("PDF and Printer Configuration")

# Load the existing configuration
config = load_config()

# Folder selection label and button
tk.Label(root, text="Select Folder for Saving PDFs:").pack(pady=10)
folder_label = tk.Label(root, text=config['pdf_folder'], bg="lightgrey", width=50)
folder_label.pack(pady=10)
tk.Button(root, text="Choose Folder", command=select_pdf_folder).pack(pady=10)

# Printer selection button
tk.Button(root, text="Choose Default Printer", command=select_default_printer).pack(pady=20)

# Set up the system tray icon if needed (if you want a system tray)
# icon = load_icon("365PrintAppIcon.ico")  # Example of loading the icon
# You can use this icon wherever needed, such as for a tray app or other purposes

# Start the Tkinter GUI event loop
root.mainloop()
