import sys
import os
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import Qt
from doc_handler_app import DocHandlerApp
from background_processor import integrate_background_processor, TaskType
from pathlib import Path
import logging
from config import CONFIG

# Define constants
TEST_FILE_NAME = "test_file.docx"

def setup_logging():
    """
    Set up logging for the application.
    """
    try:
        log_dir = Path(CONFIG['LOG_FILE_PATH']).parent

        # Ensure the log directory exists
        if not log_dir.exists():
            log_dir.mkdir(parents=True, exist_ok=True)

        # Configure logging
        logging.basicConfig(
            filename=CONFIG['LOG_FILE_PATH'],
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        logging.info("Logging initialized successfully.")
    except Exception as e:
        print(f"Failed to initialize logging: {e}")
        raise RuntimeError(f"Logging setup failed: {e}")

def ensure_directories():
    """
    Ensure required directories exist for the application.
    """
    try:
        os.makedirs(os.path.dirname(CONFIG['FILE_NAME_PORTIONS_PATH']), exist_ok=True)
        os.makedirs(os.path.dirname(CONFIG['COMPANY_NAMES_PATH']), exist_ok=True)
        os.makedirs(CONFIG['DEFAULT_SAVE_DIR'], exist_ok=True)
        os.makedirs(os.path.dirname(CONFIG['LOG_FILE_PATH']), exist_ok=True)
        logging.info("All required directories verified or created.")
    except Exception as e:
        logging.error(f"Error ensuring directories: {e}")
        raise RuntimeError(f"Failed to ensure directories: {e}")


def ensure_test_file_exists(default_save_dir):
    """
    Ensure that the test file exists in the save directory.

    :param default_save_dir: Path to the default save directory.
    """
    test_file_path = os.path.join(default_save_dir, TEST_FILE_NAME)
    try:
        # Ensure the directory exists
        if not os.path.exists(default_save_dir):
            logging.warning(f"Default save directory does not exist. Creating: {default_save_dir}")
            os.makedirs(default_save_dir, exist_ok=True)

        # Create the test file if it doesn't exist
        if not os.path.exists(test_file_path):
            with open(test_file_path, "w") as test_file:
                test_file.write("This is a placeholder test file for validation.")
            logging.info(f"Placeholder test file created at: {test_file_path}")
    except Exception as e:
        logging.error(f"Failed to create test file at {test_file_path}: {e}")
        raise FileNotFoundError(
            f"Failed to create test file at {test_file_path}. Please check permissions or disk space."
        )

def setup_application_style():
    """
    Set up modern application-wide styling.
    """
    return """
        QToolTip {
            background-color: #1F2937;
            color: white;
            padding: 8px;
            border-radius: 4px;
            border: none;
        }
        QMessageBox {
            background-color: white;
        }
        QMessageBox QPushButton {
            padding: 8px 16px;
            border-radius: 4px;
            background-color: #3B82F6;
            color: white;
            font-weight: bold;
            min-width: 80px;
        }
        QMessageBox QPushButton:hover {
            background-color: #2563EB;
        }
    """

def get_icon_path():
    """
    Ensures the application can locate the main application icon correctly.
    """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS  # ✅ PyInstaller extracts bundled files here
    else:
        base_path = os.path.abspath(".")  # ✅ Development mode

    icon_path = os.path.join(base_path, "resources/icons/main_application_icon.ico")

    if not os.path.exists(icon_path):
        logging.warning(f"WARNING: Icon file not found at {icon_path}")

    return icon_path

def main():
    """
    Entry point for the DocHandler application.
    """
    try:
        # ✅ Ensure directories and logging are set up
        ensure_directories()
        setup_logging()
        logging.info("Starting DocHandler application")

        # ✅ Initialize the application
        app = QApplication(sys.argv)

        # ✅ Set application style
        app.setStyleSheet(setup_application_style())

        # ✅ Set main application icon using the correct path
        app.setWindowIcon(QIcon(get_icon_path()))

        # ✅ Create the main window
        window = DocHandlerApp()
        window.hide()  # Keep window hidden during initialization

        # ✅ Ensure test file exists
        try:
            ensure_test_file_exists(window.default_save_dir)
        except FileNotFoundError as e:
            logging.error(f"Test file creation failed: {e}")
            QMessageBox.critical(None, "Critical Error", str(e))
            sys.exit(1)

        # ✅ Initialize background processor with integration
        try:
            integrate_background_processor(window)
            logging.info("BackgroundProcessor integrated successfully.")
        except Exception as e:
            logging.error(f"Failed to integrate BackgroundProcessor: {e}")
            QMessageBox.warning(None, "Background Processor Error", f"Background processor failed to start: {e}")

        # ✅ Handle first run setup
        if window.is_first_run():
            try:
                logging.info("First run detected. Prompting user to set default save location.")
                QMessageBox.information(window, "Welcome to DocHandler",
                                        "Welcome to DocHandler! Please set your default save location to get started.")
                window.set_default_save_location()
            except Exception as e:
                logging.error(f"First run setup failed: {e}")
                QMessageBox.critical(None, "Setup Error", f"First run setup failed: {e}")
                sys.exit(1)

        # ✅ Show the main window
        window.show()
        logging.info("DocHandler window displayed")

        # ✅ Execute the application
        exit_code = app.exec()
        logging.info(f"Application exited with code: {exit_code}")
        sys.exit(exit_code)

    except Exception as e:
        # ✅ Log and handle critical startup errors
        logging.critical(f"Application failed to start: {e}", exc_info=True)
        QMessageBox.critical(None, "Critical Error", f"Unexpected error during startup: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
