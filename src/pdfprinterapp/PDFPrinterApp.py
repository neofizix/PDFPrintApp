import sys
import os
import json
import base64
import tempfile
import threading
import logging
import win32print
import win32api
import subprocess
from flask import Flask, request, jsonify
from flask_cors import CORS
from pystray import Icon, MenuItem, Menu
from PIL import Image

app = Flask(__name__)
CORS(app)

# Setup logging
logging.basicConfig(level=logging.INFO, filename='service.log', format='%(asctime)s %(levelname)s:%(message)s')

CONFIG_FILE = "config/config.json"

def load_config():
    """ Load or create the config file. """
    default_config = {
        "pdf_folder": tempfile.gettempdir(),
        "default_printer": win32print.GetDefaultPrinter()
    }

    os.makedirs("config", exist_ok=True)
    
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w") as f:
            json.dump(default_config, f, indent=4)
        return default_config
    
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def get_pdf_save_folder():
    config = load_config()
    return config.get("pdf_folder", tempfile.gettempdir())

def get_default_printer():
    config = load_config()
    return config.get("default_printer", win32print.GetDefaultPrinter())

def save_pdf(payload_base64, doc_name):
    try:
        pdf_data = base64.b64decode(payload_base64)
        pdf_save_folder = get_pdf_save_folder()
        pdf_path = os.path.join(pdf_save_folder, doc_name)

        with open(pdf_path, "wb") as pdf_file:
            pdf_file.write(pdf_data)

        logging.info(f"PDF saved at {pdf_path}")
        return pdf_path, None

    except Exception as e:
        logging.error(f"Failed to save PDF: {str(e)}")
        return None, f"Failed to save PDF: {str(e)}"

def print_pdf_silently(pdf_path):
    try:
        if not os.path.exists(pdf_path):
            return False, f"File {pdf_path} does not exist"

        printer_name = get_default_printer()

        win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
        logging.info(f"Print job sent for {pdf_path} to {printer_name}")
        return True, "Print job sent successfully"

    except Exception as e:
        logging.error(f"Failed to print PDF: {str(e)}")
        return False, f"Failed to print: {str(e)}"

@app.route('/api/printer/printraw', methods=['POST'])
def print_raw():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data provided"}), 400

        payload_base64 = data.get("PayloadBase64")
        doc_name = data.get("DocName", "Document.pdf")
        if not payload_base64:
            return jsonify({"error": "No payload provided"}), 400

        pdf_path, error = save_pdf(payload_base64, doc_name)
        if error:
            return jsonify({"error": error}), 500

        success, message = print_pdf_silently(pdf_path)
        if not success:
            return jsonify({"error": message}), 500

        return jsonify({"message": message}), 200

    except Exception as e:
        logging.error(f"Error in print_raw route: {str(e)}")
        return jsonify({"error": str(e)}), 500

# Function to run Flask in a background thread
def run_flask():
    logging.info("Starting Flask app...")
    try:
        # Bind to localhost to ensure it's only available on the machine it's running on
        app.run(host="127.0.0.1", port=8072)
    except Exception as e:
        logging.error(f"Error running Flask app: {str(e)}")

# Function to load custom icon from file
def load_icon(icon_name):
    # Check if running as a PyInstaller bundle
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS  # PyInstaller extracts files to this temp directory
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))  # Running directly from source

    icon_path = os.path.join(base_path, icon_name)
    
    # Check if the icon exists before loading it
    if not os.path.exists(icon_path):
        raise FileNotFoundError(f"Icon file not found: {icon_path}")
    
    return Image.open(icon_path)

# Function to quit the app from system tray
def quit_app(icon, item):
    logging.info("Quitting app...")
    icon.stop()

# Function to launch the PDFPrinterConfig script
def open_settings(icon, item):
    config_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'PDFPrinterConfig.py')
    logging.info(f"Attempting to open settings: {config_script}")
    try:
        # Launch the config script using subprocess
        subprocess.Popen([sys.executable, config_script])
        logging.info("Opened PDF Printer Config successfully")
    except Exception as e:
        logging.error(f"Failed to open config script: {str(e)}")

# Setup system tray icon and menu
def setup_tray(icon_path):
    # Creating a menu with two options: Settings and Quit
    try:
        menu = Menu(
            MenuItem('Settings', open_settings),  # Open the PDFPrinterConfig script
            MenuItem('Quit', quit_app)            # Quit the application
        )
        logging.info("System tray menu created with Settings and Quit")
        
        # Create the system tray icon
        tray_icon = Icon("PDF Printer Service", load_icon(icon_path), "PDF Printer Service", menu)
        tray_icon.run()
    except Exception as e:
        logging.error(f"Failed to create system tray: {str(e)}")

# Main entry point
if __name__ == '__main__':
    # Start Flask in a background thread
    flask_thread = threading.Thread(target=run_flask)
    flask_thread.daemon = True
    flask_thread.start()

    # Icon file is expected to be bundled in the PyInstaller executable
    icon_path = "365PrintAppIcon.ico"

    # Setup system tray with custom icon
    setup_tray(icon_path)
