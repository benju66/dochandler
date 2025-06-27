import os
import sys
import logging
import shutil
import subprocess
import tempfile
import json
from pathlib import Path
from contextlib import contextmanager

from PyQt6.QtGui import QCursor, QGuiApplication, QIcon
from PyQt6.QtWidgets import (
    QWidget,
    QMessageBox,
    QFileDialog,
    QInputDialog,
    QApplication,
    QLineEdit,
    QDialog,
    QVBoxLayout,
    QLabel,
    QScrollArea,
    QFrame,
)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal

# Local modules
from ui_components import UIComponents
from background_processor import TaskType
from file_operations import FileOperations
from pdf_operations import PDFOperations
from outlook_handler import OutlookHandler
from outlook_handler import OutlookWorker
from edit_list_dialog import EditListDialog
from resource_manager import ResourceManager
from config import CONFIG

from workers import WordToPDFWorker 

class RetryableError(Exception):
    """Custom exception for handling retryable COM server errors."""
    pass

class DocHandlerApp(QWidget):
    def __init__(self):
        super().__init__()
        try:
            self._initialize_state()
            self._initialize_save_directories()
            self._initialize_components()
            self._initialize_ui()

            # ✅ Set main window icon with correct path handling
            icon_path = self._get_icon_path()
            self.setWindowIcon(QIcon(icon_path))

            self.setup_logging()
            self._finalize_setup()
            logging.info("DocHandlerApp initialized successfully.")
        except Exception as e:
            logging.critical(f"Error during initialization: {e}", exc_info=True)
            QMessageBox.critical(self, "Initialization Error", 
                                f"Failed to initialize the application: {e}. Please check the logs.")
            sys.exit(1)

    def get_data_path():
        """
        Ensures the application can locate the 'data/' directory whether running
        from source or as a packaged application.
        """
        if getattr(sys, 'frozen', False):
            # ✅ Running as a packaged app, use PyInstaller's temp directory
            base_path = Path(sys._MEIPASS)
        else:
            # ✅ Running in development mode
            base_path = Path(os.path.abspath("."))

        return base_path / "data"

    # ✅ Define paths to specific files
    data_folder = get_data_path()
    company_names_file = data_folder / "company_names.txt"
    file_name_portions_file = data_folder / "file_name_portions.txt"
    recent_filename_portions_file = data_folder / "recent_filename_portions.txt"
    recent_save_locations_file = data_folder / "recent_save_locations.txt"
    theme_config_file = data_folder / "theme_config.txt"

    # ✅ Ensure the path exists (for debugging)
    if not data_folder.exists():
        print(f"WARNING: Data folder not found at {data_folder}")

    def _get_icon_path(self):
        """
        Determines the correct path for the application icon.
        Ensures it works in both development mode and when packaged with PyInstaller.
        """
        if getattr(sys, 'frozen', False):
            # ✅ Running as a packaged application (PyInstaller)
            base_path = sys._MEIPASS
        else:
            # ✅ Running in development mode
            base_path = os.path.abspath(".")

        # ✅ Construct full icon path inside "resources/icons/"
        icon_path = os.path.join(base_path, "resources/icons/main_application_icon.ico")

        if not os.path.exists(icon_path):
            logging.warning(f"WARNING: Icon file not found at {icon_path}")

        return icon_path


    def _initialize_state(self):
        """Set initial application state and attributes."""
        try:
            self.pending_files = []
            self.current_file = None
            self.auto_convert_enabled = False
            self.dark_mode = False
            self.location_dialog_open = False
            self.processed_files = set()  # Move here from __init__
            
            config_dir = Path.home() / '.dochandler'
            config_dir.mkdir(parents=True, exist_ok=True)
            self.config_file = config_dir / 'config.json'
            
            self.session_save_dir = None
            
            logging.info(f"Application state initialized with config file: {self.config_file}")
        except Exception as e:
            logging.error(f"Error initializing application state: {str(e)}")
            raise

    def _initialize_save_directories(self):
        """Load and validate save directories."""
        try:
            save_dir = self.load_save_location()
            if not save_dir or not os.path.isdir(save_dir):  # Validate save_dir
                raise ValueError(f"Invalid save directory: {save_dir}")
            self.default_save_dir = Path(save_dir)  # Convert to Path object
        except Exception as e:
            logging.error(f"Error loading save location: {e}")
            self.default_save_dir = Path(CONFIG.get('DEFAULT_SAVE_DIR', Path.home() / 'Downloads' / 'DocHandler'))
            logging.warning("Using fallback default save directory.")

        if not self.default_save_dir.exists():
            self.default_save_dir.mkdir(parents=True, exist_ok=True)


    def _initialize_components(self):
        """Initialize core components of the application."""
        try:
            # Initialize resource manager
            self.resource_manager = ResourceManager()

            # Initialize file operations and handlers
            self.file_ops = FileOperations()
            self.pdf_ops = PDFOperations(self.file_ops, self.resource_manager)
            self.outlook_handler = OutlookHandler(self.file_ops, self.pdf_ops, self.resource_manager)

        except Exception as e:
            logging.error(f"Error initializing components: {e}", exc_info=True)
            raise

    def _initialize_ui(self):
        """Set up the user interface."""
        try:
            self.ui_components = UIComponents(self)
            self.init_ui()
            self.setup_logging()
            self.ui_components.save_location_label.update_path(self.default_save_dir)

            # Configure UI initial states
            self.ui_components.toggle_theme(enabled=False)
            self.ui_components.toggle_recent_portions(False)

            # Load and update filename portions
            self._initialize_filename_portions()

        except Exception as e:
            logging.error(f"Error initializing UI: {e}", exc_info=True)
            raise

    def _initialize_filename_portions(self):
        """Load and update filename portions and related components."""
        try:
            self.filename_portions_enabled = False
            self.filename_portion = ""
            self.portions_list = self.file_ops.load_file_name_portions()
            self.company_names = self.file_ops.load_company_names()
            self.recent_portions = self.file_ops.load_recent_filename_portions()

            self.ui_components.update_filename_portions_widget(self.portions_list)
            self.ui_components.update_recent_portions_list_widget(self.recent_portions)

        except Exception as e:
            logging.error(f"Error initializing filename portions: {e}")
            self.portions_list = []
            self.recent_portions = []

    def _finalize_setup(self):
        """Perform final setup steps."""
        self.update_drag_drop_text("Drag and drop an Outlook email attachments, PDFs, or Word document here.")
        self.setAcceptDrops(True)
        self.hide()  # Keep the window hidden until fully initialized


    def init_ui(self):
        """Initialize the UI components and layout."""
        try:
            # Hide the window during initialization
            self.hide()

            # Set window properties before UI setup
            self.setWindowTitle("DocHandler - v1.1.0")

            # Center the window on the screen
            self.center_window()

            # Initialize UI components with the window hidden
            self.ui_components.setup_ui()

            # Connect signals
            self.connect_signals()

            logging.info("UI initialized successfully.")
        except Exception as e:
            logging.error(f"Error initializing UI: {e}", exc_info=True)
            raise


    def center_window(self):
        """Center the window on the screen."""
        screen = QGuiApplication.primaryScreen().availableGeometry()
        window_width = 800  # Default width
        window_height = 600  # Default height

        # Set the window size
        self.resize(window_width, window_height)

        # Calculate the position to center the window
        x = (screen.width() - window_width) // 2
        y = (screen.height() - window_height) // 2

        # Move the window to the centered position
        self.move(x, y)


    def closeEvent(self, event):
        """Handle application close event."""
        try:
            # Save application state
            self.save_config(self.default_save_dir)

            # Clean up all managed resources
            self.resource_manager.cleanup_all()

            logging.info("Application shutting down cleanly")

        except Exception as e:
            logging.error(f"Error during application shutdown: {str(e)}", exc_info=True)
        finally:
            event.accept()


    def resizeEvent(self, event):
        """Handle window resize events."""
        super().resizeEvent(event)

    @contextmanager
    def busy_cursor(self):
        """Context manager to show both pointer and wait cursor."""
        try:
            app = QApplication.instance()
            if app:
                # First set the arrow cursor
                app.setOverrideCursor(QCursor(Qt.CursorShape.ArrowCursor))
                # Then add the wait cursor
                app.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
                app.processEvents()
            yield
        finally:
            app = QApplication.instance()
            if app:
                # Restore cursors (needs to be called twice since we set two cursors)
                app.restoreOverrideCursor()
                app.restoreOverrideCursor()
                app.processEvents()

    def setup_logging(self, log_level=logging.DEBUG):
        """
        Configure logging to output to both a file and the console.
        """

         # Check if handlers already exist to avoid duplicate initialization
        logger = logging.getLogger("doc_handler_logger")
        if logger.handlers:  # If handlers are already present, do nothing
            return
        
        try:
            
            log_file = CONFIG['LOG_FILE_PATH']
            log_dir = Path(log_file).parent
            log_dir.mkdir(parents=True, exist_ok=True)

            # Define logging format
            log_format = '%(asctime)s - %(levelname)s - %(message)s'

            # Create a logger specifically for DocHandler
            logger = logging.getLogger("doc_handler_logger")
            logger.setLevel(log_level)

            # Avoid adding duplicate handlers
            if not logger.handlers:
                # File Handler (persistent logging)
                file_handler = logging.FileHandler(log_file)
                file_handler.setLevel(log_level)  # Log level for file
                file_handler.setFormatter(logging.Formatter(log_format))

                # Console Handler (real-time monitoring)
                console_handler = logging.StreamHandler()
                console_handler.setLevel(logging.INFO)  # Log level for console
                console_handler.setFormatter(logging.Formatter(log_format))

                # Add handlers to the logger
                logger.addHandler(file_handler)
                logger.addHandler(console_handler)

            logger.info("Logging initialized: logs to file and console.")
        except Exception as e:
            print(f"Failed to initialize logging: {e}")


    def connect_signals(self):
        try:
            # File Menu signals
            self.ui_components.save_location_action.triggered.connect(self.set_save_location)
            self.ui_components.default_save_location_action.triggered.connect(self.set_default_save_location)
            self.ui_components.recent_locations_menu.aboutToShow.connect(self.ui_components.show_recent_locations_menu)
            self.ui_components.toggle_auto_convert_action.triggered.connect(self.toggle_auto_convert)
            self.ui_components.exit_action.triggered.connect(self.close)

            # **New: Check for Updates Button**
            self.ui_components.check_update_action.triggered.connect(self.check_for_updates)

            # Processing Menu signals
            self.ui_components.enable_filename_portions_action.triggered.connect(self.toggle_save_quotes)
            self.ui_components.edit_company_names_action.triggered.connect(self.edit_company_names)
            self.ui_components.edit_file_name_portions_action.triggered.connect(self.edit_file_name_portions)

            # View Menu signals
            self.ui_components.toggle_dark_mode_action.triggered.connect(self.toggle_theme)
            self.ui_components.toggle_recent_files_action.triggered.connect(self.ui_components.toggle_recent_files)
            self.ui_components.toggle_recent_portions_action.triggered.connect(self.toggle_recent_portions)
            self.ui_components.debug_action.triggered.connect(self.toggle_debug_mode)

            # Help Menu signals
            self.ui_components.help_action.triggered.connect(self.show_help)
            self.ui_components.about_action.triggered.connect(self.show_about)

            # Other UI component signals
            self.ui_components.search_bar.textChanged.connect(self.filter_portions_list)
            self.ui_components.portions_list_widget.itemClicked.connect(self.select_filename_portion)
            self.ui_components.recent_files_list.itemClicked.connect(self.open_recent_file)
            self.ui_components.recent_portions_list_widget.itemClicked.connect(self.select_filename_portion)
            self.ui_components.clear_recent_portions_button.clicked.connect(self.clear_recent_portions)
            self.ui_components.save_button.clicked.connect(self.save_current_file)
            self.ui_components.progress_bar.valueChanged.connect(self.on_progress_changed)

            logging.info("All signals connected successfully.")
        except Exception as e:
            logging.error(f"Error connecting signals: {e}", exc_info=True)


    def load_save_location(self):
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    save_dir = config.get('default_save_dir')
                    if save_dir and os.path.isdir(save_dir):
                        return save_dir
                    else:
                        raise ValueError(f"Configured save directory is invalid: {save_dir}")
            else:
                self.save_config(CONFIG['DEFAULT_SAVE_DIR'])
                return CONFIG['DEFAULT_SAVE_DIR']
        except Exception as e:
            logging.error(f"Error loading save location: {e}")
            raise


        
    def save_config(self, save_dir):
        """Save application configuration to file with proper error handling."""
        try:
            # Ensure the directory exists
            config_dir = self.config_file.parent
            config_dir.mkdir(parents=True, exist_ok=True)

            # Prepare configuration data
            config = {}
            if self.config_file.exists():
                try:
                    with open(self.config_file, 'r') as f:
                        config = json.load(f)
                except json.JSONDecodeError:
                    logging.warning("Invalid JSON in config file. Creating new configuration.")
                    config = {}

            # Update configuration
            config['default_save_dir'] = str(save_dir)  # Convert Path to string

            # Write configuration with proper permissions
            temp_config_file = self.config_file.with_suffix('.tmp')
            
            # First write to a temporary file
            with open(temp_config_file, 'w') as f:
                json.dump(config, f, indent=4)

            # Then replace the original file
            if os.name == 'nt':  # Windows
                # Remove original file if it exists (Windows often needs this)
                if self.config_file.exists():
                    self.config_file.unlink()
                temp_config_file.replace(self.config_file)
            else:  # Unix-like systems
                os.replace(temp_config_file, self.config_file)

            # Set appropriate permissions
            if os.name != 'nt':  # Skip on Windows
                os.chmod(self.config_file, 0o644)

            logging.info(f"Configuration saved successfully to {self.config_file}")
        except PermissionError as pe:
            logging.error(f"Permission denied when saving configuration: {str(pe)}")
            QMessageBox.warning(
                self, 
                "Save Error",
                "Permission denied when saving configuration. Please check folder permissions."
            )
        except OSError as oe:
            logging.error(f"OS error when saving configuration: {str(oe)}")
            QMessageBox.warning(
                self, 
                "Save Error",
                f"Error saving the configuration: {str(oe)}"
            )
        except Exception as e:
            logging.error(f"Unexpected error saving configuration: {str(e)}", exc_info=True)
            QMessageBox.warning(
                self, 
                "Save Error",
                "Unexpected error saving the configuration. Check the application logs."
            )

    def is_first_run(self):
        return not self.config_file.exists()

    def handle_first_run(self):
        logging.info("First run detected. Prompting for default save location.")
        QMessageBox.information(self, "Welcome", 
                              "Welcome to DocHandler! Please set your default save location.")
        self.set_default_save_location()

    def set_save_location(self):
        if self.location_dialog_open:
            return
        self.location_dialog_open = True
        try:
            # Convert Path to string
            initial_dir = str(self.session_save_dir or self.default_save_dir)
            save_dir = QFileDialog.getExistingDirectory(
                self, "Select Save Folder", initial_dir
            )
            if save_dir:
                self.session_save_dir = save_dir
                recent_locations = self.file_ops.load_recent_save_locations()
                if save_dir not in recent_locations:
                    recent_locations.insert(0, save_dir)
                    self.file_ops.save_recent_save_locations(recent_locations)
                self.ui_components.update_recent_locations_menu(recent_locations)
                self.ui_components.save_location_label.update_path(save_dir)
                QMessageBox.information(self, "Save Location", f"Session save location set to: {save_dir}")
        finally:
            self.location_dialog_open = False


    
    def set_save_location_from_recent(self, location):
        if location:
            self.session_save_dir = location
            
            # Update recency order
            recent_locations = self.file_ops.load_recent_save_locations()
            if location in recent_locations:
                recent_locations.remove(location)
            recent_locations.insert(0, location)
            self.file_ops.save_recent_save_locations(recent_locations)
            
            # Update UI
            self.ui_components.update_recent_locations_menu(recent_locations)
            self.ui_components.save_location_label.update_path(location)
            QMessageBox.information(self, "Save Location", 
                                f"Session save location set to: {location}")



    def set_default_save_location(self):
        if self.location_dialog_open:
            logging.warning("Default save location dialog is already open.")
            return
        self.location_dialog_open = True
        self.ui_components.default_save_location_action.setEnabled(False)
        try:
            save_dir = QFileDialog.getExistingDirectory(
                self, "Select Default Save Folder", self.default_save_dir
            )
            if save_dir and save_dir != self.default_save_dir:
                self.default_save_dir = save_dir
                self.session_save_dir = save_dir
                self.save_config(save_dir)
                
                # Update the save location label
                self.ui_components.save_location_label.update_path(save_dir)
                
                message = f"Default save location set to: {self.default_save_dir}"
                logging.info(message)
                self.ui_components.set_label_text(f"Drag and drop here. {message}")
                QMessageBox.information(self, "Default Save Location", message)
            elif not save_dir:
                logging.info("User canceled the default save location selection.")
                self.ui_components.set_label_text(
                    "Drag and drop an Outlook email, attachment, or Word document here."
                )
        finally:
            self.location_dialog_open = False
            self.ui_components.default_save_location_action.setEnabled(True)

    def toggle_maximize(self):
        """Toggle between maximized and normal window state."""
        if self.isMaximized():
            self.showNormal()
            self.ui_components.maximize_action.setText("Maximize")
        else:
            self.showMaximized()
            self.ui_components.maximize_action.setText("Restore")

    def show_help(self):
        """Show the help dialog."""
        help_content = (
            "DocHandler Help:\n\n"
            "Drag and drop files to process them\n"
             "1. Supported File Types\n"
            "   • Email Attachments: Outlook email attachments\n"
            "   • Word Documents: .doc, .docx\n"
            "   • PDFs: Standard and OCR-supported image PDFs\n\n"
            "2. File Management\n"
            "   • Drag and drop files to process them\n"
            "   • Default save location: ~/Downloads/DocHandler\n"
            "   • Change default location: Settings → Set Default Save Location\n"
            "   • Temporary save location: Settings → Set Save Location\n\n"
            "3. Save Quotes\n"
            "   • Navigate to Settings > Save Quotes\n"
            "   • Set the non-default save location.\n"
            "   • Search for or select a Filename portion from the list (e.g. 02-4100 - Demolition)\n"
            "   • Drop quote into drop area\n"
            "   • The user will be prompted to confirm the company name if it's in the existing company database or the user will be prompted to enter the company name in text box\n"
            "   • The user can select a filename portion from the 'Recently Used Portions' and/or clear this window by clicking the 'Clear Recent Portions' button\n\n"
            "4. Recent Files\n"
            "   • View and open recent files in default applications\n"
            "   • Clear list: Click 'Clear Recent Portions'\n\n"
            "5. Preview Panel\n"
            "   • Toggle: Settings → Show Preview Panel\n"
            "   • Navigation: Next/Previous buttons or arrow keys\n"
            "   • Zoom controls available in toolbar\n\n"
            "6. Advanced Features\n"
            "   • Outlook integration for email attachments\n"
            "   • Edit company names and filename portions\n"
            "   • Dark mode toggle: Settings → Dark Mode\n"
            "   • Debug mode: Settings → Debug Mode\n\n"
            "7. Keyboard Shortcuts\n"
            "   • Zoom in: Ctrl+shift++\n"
            "   • Zoom out: Ctrl+-\n"
            "   • Reset zoom: Ctrl+0\n"
        )
        dialog = ScrollableDialog("Help", help_content, self)
        dialog.exec()

    def show_about(self):
        """Show the about dialog."""
        about_content = (
            "DocHandler v1.1.0\n\n"
            "A tool for handling documents and email attachments.\n\n"
            "New in v1.1.0:\n"
            "• Auto-focus company name field\n"
            "• Enhanced filename handling\n"
            "• Enhanced filename company name saving\n"
            "• Access recently saved file locations\n\n"
            "Features:\n"
            "• Document conversion and processing\n"
            "• Convert .doc and .docx files to PDFs\n"
            "• Automatically save PDFs and converted files to default or selected folder\n"
            "• Save files by selected scope and add company name\n"
            "• Filename customization\n"
            "• Dark/Light themes\n\n"
            "Known Issues\n"
            "   • Preview Window Tool Bar Page numbering and navigation not updating correctly\n\n"
            "Future Updates:\n"
            "   • Support more document types\n"
            "   • Open Save folder\n"
            "   • Better rendering\n"
            "   • Improved company name recognition\n"
            "   • Merge Files\n"
            "   • Unlock PDFs\n"
            "   • Support more document types\n\n"
        )
        dialog = ScrollableDialog("About DocHandler", about_content, self)
        dialog.exec()

    def on_progress_changed(self, value):
        """Handle progress bar value changes."""
        if value == 100:
            # Auto-hide progress bar after 1 second when complete
            QTimer.singleShot(1000, self.ui_components.hide_progress)

    def toggle_auto_convert(self, enabled):
        """Toggle auto-convert functionality."""
        self.auto_convert_enabled = enabled
        
        # Clear pending files when enabling auto-convert
        if enabled:
            self.pending_files.clear()
            self.ui_components.update_pending_files_list([])
        
        # Show save button when auto-convert is OFF
        self.ui_components.save_button.setVisible(not enabled)
        
        # Update UI text based on state
        if enabled:
            if self.current_file:
                self.ui_components.set_label_text(f"{self.current_file}")
                # Process the current file
                self.process_dropped_file(self.current_file)
            else:
                self.ui_components.set_label_text("Drag and drop an Outlook email attachment, PDFs, or Word document here.")
        else:
            if len(self.pending_files) > 0:
                self.ui_components.set_label_text(f"Ready to save: {len(self.pending_files)} files")
            else:
                self.ui_components.set_label_text("Drag and drop an Outlook email attachment, PDFs, or Word document here.")
        
        logging.info(f"Auto convert {'enabled' if enabled else 'disabled'}")

    def toggle_recent_portions(self, enabled):
        """Toggle visibility of recent portions area."""
        self.recent_portions_visible = enabled
        self.ui_components.recent_portions_group.setVisible(enabled)
        logging.info(f"Recent portions visibility {'enabled' if enabled else 'disabled'}")

    def save_current_file(self):
        """Handle save button click with support for multiple files and Save Quotes mode."""
        if not self.pending_files and not self.current_file:
            self.ui_components.show_warning_message("No Files", "No files available to save.")
            self.update_drag_drop_text("Drag and drop an Outlook email attachment, PDFs, or Word documents here.")
            return

        try:
            self.update_drag_drop_text("Processing files... Please wait.")
            active_save_dir = self.session_save_dir or self.default_save_dir

            # Check if we have multiple files to merge
            if len(self.pending_files) > 1:
                # Convert all files to PDF first if needed
                pdf_files = []
                for file_path in self.pending_files:
                    try:
                        if file_path.lower().endswith('.pdf'):
                            pdf_files.append(file_path)
                        else:
                            # Convert non-PDF files to PDF first
                            if self.filename_portions_enabled:
                                company_name = self.scan_document_for_company_names(file_path)
                                if not company_name:
                                    company_name = self.prompt_for_company_name()
                                    if not company_name:  # User canceled
                                        continue
                                temp_name = f"{self.filename_portion} - {company_name}.pdf"
                            else:
                                temp_name = os.path.splitext(os.path.basename(file_path))[0] + '.pdf'
                            
                            temp_dir = self.file_ops.create_temp_folder()
                            temp_pdf = self.pdf_ops.convert_to_pdf(file_path, temp_dir, temp_name)
                            pdf_files.append(temp_pdf)

                    except Exception as e:
                        logging.error(f"Error processing file {file_path}: {str(e)}")
                        self.ui_components.show_error_message(
                            "Processing Error", 
                            f"Could not process {os.path.basename(file_path)}: {str(e)}"
                        )

                if pdf_files:
                    try:
                        # Generate merged filename
                        if self.filename_portions_enabled:
                            company_name = self.prompt_for_company_name()
                            if company_name:
                                merged_name = f"{self.filename_portion} - {company_name}_merged.pdf"
                            else:
                                merged_name = "merged_document.pdf"
                        else:
                            merged_name = "merged_document.pdf"

                        # Get unique filename
                        output_path = self.file_ops.get_unique_filename(active_save_dir, merged_name)

                        # Merge PDFs
                        merged_path = self.pdf_ops.merge_pdfs(pdf_files, output_path)

                        # Update UI
                        self.ui_components.update_recent_files(merged_path)
                        self.ui_components.set_label_text(f"Merged file saved: {merged_path}")

                        logging.info(f"Successfully merged files to: {merged_path}")

                    except Exception as e:
                        logging.error(f"Error merging PDFs: {str(e)}")
                        self.ui_components.show_error_message("Merge Error", f"Error merging PDFs: {str(e)}")

                # Clean up temporary files
                for pdf_file in pdf_files:
                    if pdf_file not in self.pending_files:  # Only delete temporary conversions
                        try:
                            os.remove(pdf_file)
                        except Exception as e:
                            logging.warning(f"Failed to clean up temporary file {pdf_file}: {str(e)}")

            else:
                # Single file processing (existing code for single file)
                file_path = self.pending_files[0]
                try:
                    if self.filename_portions_enabled:
                        company_name = self.scan_document_for_company_names(file_path)
                        if not company_name:
                            company_name = self.prompt_for_company_name()
                            if not company_name:  # User canceled
                                return
                        
                        new_file_name = f"{self.filename_portion} - {company_name}.pdf"
                    else:
                        new_file_name = os.path.basename(file_path)
                        if not new_file_name.lower().endswith('.pdf'):
                            new_file_name = os.path.splitext(new_file_name)[0] + '.pdf'

                    saved_path = self.pdf_ops.convert_to_pdf(file_path, active_save_dir, new_file_name)
                    self.ui_components.update_recent_files(saved_path)
                    self.ui_components.set_label_text(f"File saved: {saved_path}")
                    self.ui_components.drag_drop_area.set_status("Done! Let's save another file.")

                except Exception as e:
                    logging.error(f"Error processing file: {str(e)}")
                    self.ui_components.show_error_message(
                        "Processing Error",
                        f"Could not process {os.path.basename(file_path)}: {str(e)}"
                    )

            # Clear pending files after processing
            self.pending_files.clear()
            self.ui_components.update_pending_files_list([])

        except Exception as e:
            logging.error(f"Error saving files: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Save Error", f"Error saving files: {str(e)}")
        finally:
            self.ui_components.hide_progress()

    def _handle_merged_files(self, active_save_dir):
        """Handle merging multiple files into a single PDF."""
        try:
            # Convert all files to PDF first
            pdf_files = []
            temp_files = []  # Track temporary files for cleanup

            # Show progress
            self.ui_components.show_progress(0)
            total_files = len(self.pending_files)

            for index, file_path in enumerate(self.pending_files):
                try:
                    # Update progress
                    progress = int((index / total_files) * 50)  # First 50% for conversion
                    self.ui_components.show_progress(progress)
                    
                    if file_path.lower().endswith('.pdf'):
                        pdf_files.append(file_path)
                    else:
                        # Create temporary file for conversion
                        temp_dir = self.file_ops.create_temp_folder()
                        temp_name = f"temp_{index}_{os.path.basename(file_path)}"
                        temp_pdf = self.pdf_ops.convert_to_pdf(
                            file_path,
                            temp_dir,
                            temp_name
                        )
                        pdf_files.append(temp_pdf)
                        temp_files.append(temp_pdf)

                except Exception as e:
                    logging.error(f"Error converting file {file_path}: {str(e)}")
                    self.ui_components.show_warning_message(
                        "Conversion Warning",
                        f"Could not convert {os.path.basename(file_path)}. Skipping file."
                    )
                    continue

            if not pdf_files:
                raise ValueError("No valid PDF files to merge")

            # Generate output filename
            if self.filename_portions_enabled:
                company_name = self.prompt_for_company_name()
                if company_name:
                    merged_filename = f"{self.filename_portion} - {company_name}_merged.pdf"
                else:
                    merged_filename = "merged_document.pdf"
            else:
                merged_filename = "merged_document.pdf"

            # Get unique filename for merged file
            output_path = self.file_ops.get_unique_filename(active_save_dir, merged_filename)

            # Merge the PDFs
            self.ui_components.show_progress(75)  # Show merge progress
            merged_path = self.pdf_ops.merge_pdfs(pdf_files, output_path)

            # Update UI and recent files
            self.ui_components.update_recent_files(merged_path)
            self.ui_components.set_label_text(f"Merged file saved: {merged_path}")
            
            # Clear pending files
            self.pending_files.clear()
            self.ui_components.update_pending_files_list([])
            self.ui_components.show_progress(100)

            logging.info(f"Successfully merged files to: {merged_path}")

        except Exception as e:
            logging.error(f"Error merging files: {str(e)}", exc_info=True)
            raise
        finally:
            # Clean up temporary files
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except Exception as e:
                    logging.warning(f"Failed to clean up temporary file {temp_file}: {str(e)}")

    def _handle_individual_files(self, active_save_dir):
        """Handle saving files individually."""
        total_files = len(self.pending_files)
        
        for index, file_path in enumerate(self.pending_files):
            try:
                # Update progress
                progress = int((index / total_files) * 100)
                self.ui_components.show_progress(progress)

                if self.filename_portions_enabled:
                    # Save Quotes mode - scan for company name
                    company_name = self.scan_document_for_company_names(file_path)
                    if not company_name:
                        company_name = self.prompt_for_company_name()
                        if not company_name:  # User canceled
                            continue
                    
                    new_file_name = f"{self.filename_portion} - {company_name}.pdf"
                else:
                    new_file_name = os.path.basename(file_path)
                    if not new_file_name.lower().endswith('.pdf'):
                        new_file_name = os.path.splitext(new_file_name)[0] + '.pdf'

                # Convert and save the file
                saved_path = self.pdf_ops.convert_to_pdf(file_path, active_save_dir, new_file_name)

                # Update recent files and UI
                self.ui_components.update_recent_files(saved_path)
                self.ui_components.set_label_text(f"File saved: {saved_path}")

                logging.info(f"Processed file saved to {saved_path}")

            except Exception as e:
                logging.error(f"Error processing file {file_path}: {str(e)}", exc_info=True)
                self.ui_components.show_error_message(
                    "Processing Error",
                    f"Could not process {os.path.basename(file_path)}: {str(e)}"
                )

        # Clear pending files after processing
        self.pending_files.clear()
        self.ui_components.update_pending_files_list([])
        self.ui_components.show_progress(100)

    def toggle_debug_mode(self):
        """Toggle debug mode for development purposes."""
        debug_enabled = not getattr(self, 'debug_mode', False)
        self.debug_mode = debug_enabled
        logging.getLogger().setLevel(logging.DEBUG if debug_enabled else logging.INFO)
        self.ui_components.show_info_message(
            "Debug Mode", 
            f"Debug mode {'enabled' if debug_enabled else 'disabled'}"
        )
    
    def toggle_save_quotes(self, enabled):
        """Enable or disable Save Quotes functionality."""
        self.filename_portions_enabled = enabled
        self.ui_components.filename_portions_group.setVisible(enabled)

        if enabled:
            self.ui_components.set_label_text("Select a filename portion and add files.")
        else:
            self.pending_files.clear()
            self.ui_components.update_pending_files_list([])

        logging.info(f"Save Quotes {'enabled' if enabled else 'disabled'}.")
    
    def toggle_theme(self, enabled):
        """Toggle the UI theme between light and dark."""
        self.ui_components.toggle_theme(enabled)
    
    def filter_portions_list(self, search_text):
        """Filter and display filename portions based on the search text."""
        filtered_portions = [
            portion for portion in self.portions_list
            if search_text.lower() in portion.lower()
        ]
        self.ui_components.update_filename_portions_widget(filtered_portions)
        logging.debug(f"Filtered filename portions with text '{search_text}': {len(filtered_portions)} results.")


    def select_filename_portion(self, item):
        portion_text = item.text()
        self.filename_portion = portion_text
        self.ui_components.search_bar.setText(self.filename_portion)
        logging.info(f"Filename portion selected: {self.filename_portion}")
        self.update_recent_portions(portion_text)
    
    def update_recent_portions(self, selected_portion):
        if selected_portion in self.recent_portions:
            self.recent_portions.remove(selected_portion)
        self.recent_portions.insert(0, selected_portion)
        self.recent_portions = self.recent_portions[:10]
        self.file_ops.save_recent_filename_portions(self.recent_portions)
        self.ui_components.update_recent_portions_list_widget(self.recent_portions)

    def clear_recent_portions(self):
        """Clear recent portions list after user confirmation."""
        # Ask for confirmation using UIComponents
        if self.ui_components.get_confirmation(
            'Clear Recent Portions', 'Are you sure you want to clear all recent portions?'
        ):
            # Clear the recent portions data
            self.recent_portions.clear()
            self.file_ops.save_recent_filename_portions([])  # Persist empty list

            # Update the UI to reflect the cleared list
            self.ui_components.update_recent_portions_list_widget([])

            # Log the action and show an information message to the user
            logging.info("Recent filename portions cleared")
            self.ui_components.show_info_message("Cleared", "Recent portions have been cleared.")

    def edit_company_names(self):
        dialog = EditListDialog(CONFIG['COMPANY_NAMES_PATH'], "Edit Company Names", self)
        dialog.exec()  # Display the dialog

    def edit_file_name_portions(self):
        dialog = EditListDialog(CONFIG['FILE_NAME_PORTIONS_PATH'], "Edit Filename Portions", self)
        dialog.exec()  # Display the dialog


    def edit_text_file(self, file_path, title):
        try:
            with open(file_path, 'r+', encoding='utf-8') as file:
                current_text = file.read()
            
            text, ok = self.ui_components.get_text_input(title, f"Current entries:\n{current_text}")
            
            if ok:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(text)
                logging.info(f"{title} updated successfully.")
                self.ui_components.show_info_message("Success", f"{title} has been updated.")
        except Exception as e:
            logging.error(f"Error editing {title}: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Edit Error", f"Could not edit {title}.")



    def open_recent_file(self, item):
        try:
            file_name = item.text()
            logging.info(f"Attempting to open recent file: {file_name}")
            if not self.default_save_dir:
                raise Exception("Save location not set. Please set a save location in the settings first.")
            active_save_dir = self.session_save_dir or self.default_save_dir
            full_path = os.path.join(active_save_dir, file_name)
            self.file_ops.open_file(full_path)
            logging.info(f"File opened successfully: {full_path}")
        except Exception as e:
            error_msg = f"Error opening recent file: {str(e)}"
            logging.error(error_msg, exc_info=True)
            QMessageBox.warning(self, "File Not Found", error_msg)

    def dragEnterEvent(self, event):
        """Handle drag enter events and validate drop data."""
        if self.file_ops.is_valid_drop(event.mimeData()):
            event.acceptProposedAction()
            logging.debug("Drag event accepted")
        else:
            event.ignore()
            logging.debug("Drag event ignored")

    def update_drag_drop_text(self, text):
        """Update the drag-and-drop area text dynamically."""
        if hasattr(self.ui_components, 'drag_drop_area'):
            self.ui_components.drag_drop_area.set_status(text)



    def generate_filename_with_company(self, file_path, filename_portion):
        """Generate filename with company name if available."""
        try:
            if not filename_portion:
                return os.path.basename(file_path)

            company_name = self.scan_document_for_company_names(file_path)
            if company_name:
                base, ext = os.path.splitext(os.path.basename(file_path))
                return f"{filename_portion} - {company_name}{ext}"
            else:
                return self.file_ops.generate_new_filename(file_path, filename_portion)
        except Exception as e:
            logging.error(f"Error generating filename with company: {str(e)}", exc_info=True)
            return self.file_ops.generate_new_filename(file_path, filename_portion)

    def dropEvent(self, event):
        """Handle files dropped into the application."""
        try:
            mime_data = event.mimeData()
            if not self.file_ops.is_valid_drop(mime_data):
                raise ValueError("Invalid drop data")

            if self.file_ops.is_outlook_item(mime_data):
                temp_file_path = self.outlook_handler.get_outlook_item(mime_data)
                if temp_file_path and os.path.exists(temp_file_path):
                    self.process_dropped_file(temp_file_path)
            else:
                # Handle regular file drops
                for url in mime_data.urls():
                    file_path = url.toLocalFile()
                    if file_path:
                        file_ext = os.path.splitext(file_path)[1].lower()
                        if file_ext in ['.pdf', '.doc', '.docx']:
                            self.process_dropped_file(file_path)
                        else:
                            logging.warning(f"Skipping unsupported file type: {file_path}")

            event.acceptProposedAction()

        except Exception as e:
            logging.error(f"Error processing dropped file: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Drop Error", str(e))


    def handle_file_drop(self, mime_data, temp_folder):
        try:
            if self.file_ops.is_outlook_item(mime_data):
                temp_file_path = self.outlook_handler.get_outlook_item(mime_data)
                if temp_file_path and os.path.exists(temp_file_path):
                    self.process_dropped_file(temp_file_path)  # Changed from process_single_file
                    try:
                        os.remove(temp_file_path)
                    except Exception as e:
                        logging.warning(f"Error removing temp file: {str(e)}")
                return

            for url in mime_data.urls():
                file_path = url.toLocalFile()
                if file_path:
                    file_ext = os.path.splitext(file_path)[1].lower()
                    if file_ext in ['.pdf', '.doc', '.docx']:
                        self.process_dropped_file(file_path)  # Changed from process_single_file
                    else:
                        logging.warning(f"Skipping unsupported file type: {file_path}")

        except Exception as e:
            logging.error(f"Error handling dropped files: {str(e)}", exc_info=True)
            self.ui_components.show_error_message(
                "Drop Error",
                f"Error processing dropped files: {str(e)}"
            )
            self.update_drag_drop_text("An error occurred. Please try again.")


    def _handle_com_error(self, e: Exception) -> None:
        if isinstance(e, RetryableError):
            self.ui_components.show_warning_message(
                "COM Server Error",
                "Connection to Microsoft Office failed. Retrying..."
            )
            QApplication.processEvents()
            return True
        return False

    def process_dropped_file(self, file_path):
        """Process a dropped file, ensuring filename conventions apply correctly in auto-convert mode."""
        if file_path in self.processed_files:
            logging.info(f"File already processed: {file_path}")
            return

        try:
            with self.busy_cursor():
                logging.info(f"Processing dropped file: {file_path}")
                self.ui_components.show_progress(0)
                self.ui_components.set_label_text(f"Processing: {file_path}")

                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"File not found: {file_path}")

                active_save_dir = self.session_save_dir or self.default_save_dir
                file_ext = os.path.splitext(file_path)[1].lower()

                # Handle non-auto-convert mode
                if not self.auto_convert_enabled:
                    if file_ext in ['.pdf', '.doc', '.docx']:
                        if file_path not in self.pending_files:
                            self.pending_files.append(file_path)
                        self.ui_components.update_pending_files_list(self.pending_files)
                        self.ui_components.set_label_text(f"Ready to save: {len(self.pending_files)} files")
                        return
                    else:
                        raise ValueError(f"Unsupported file type: {file_ext}")

                # --- Auto-Convert Mode ---
                if file_ext in ['.doc', '.docx', '.pdf']:
                    # Apply filename portion if Save Quotes is enabled
                    new_filename = os.path.basename(file_path)
                    if self.filename_portions_enabled:
                        if not self.filename_portion:
                            raise ValueError("Filename portion is required in Save Quotes mode.")

                        # Extract or prompt for company name
                        company_name = self.scan_document_for_company_names(file_path)
                        if not company_name:
                            company_name = self.prompt_for_company_name()
                            if not company_name:
                                return  # User canceled

                        # Format filename correctly
                        new_filename = f"{self.filename_portion} - {company_name}.pdf"

                    # Convert Word files to PDF
                    if file_ext in ['.doc', '.docx']:
                        converted_path = self.pdf_ops.convert_to_pdf(file_path, active_save_dir, new_filename)
                        self._on_file_processed(converted_path)
                    else:  # PDF files
                        saved_path = self.file_ops.save_file(file_path, active_save_dir, new_filename)
                        self._on_file_processed(saved_path)

                else:
                    raise ValueError(f"Unsupported file type: {file_ext}")

        except Exception as e:
            logging.error(f"Error processing dropped file: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Processing Error", str(e))
            self.update_drag_drop_text("An error occurred. Please try again.")
        finally:
            self.ui_components.hide_progress()

    def process_single_file(self, file_path):
        """Process a single file with correct naming conventions in auto-convert mode."""
        try:
            logging.info(f"Starting to process file: {file_path}")
            self.ui_components.show_progress(0)
            self.update_drag_drop_text(f"Processing: {file_path}")

            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")

            active_save_dir = self.session_save_dir or self.default_save_dir
            file_ext = os.path.splitext(file_path)[1].lower()

            # Handle non-auto-convert mode
            if not self.auto_convert_enabled:
                if file_ext in ['.pdf', '.doc', '.docx']:
                    if self.filename_portions_enabled and not self.filename_portion:
                        self.ui_components.show_warning_message(
                            "No Portion Selected",
                            "Please select a filename portion before processing files."
                        )
                        return

                    if file_path not in self.pending_files:
                        self.pending_files.append(file_path)
                    self.ui_components.update_pending_files_list(self.pending_files)
                    self.ui_components.set_label_text(f"Ready to save: {len(self.pending_files)} files")
                    return
                else:
                    raise ValueError(f"Unsupported file type: {file_ext}")

            # --- Auto-Convert Mode ---
            if file_ext in ['.doc', '.docx', '.pdf']:
                # Apply filename portion if Save Quotes is enabled
                new_filename = os.path.basename(file_path)
                if self.filename_portions_enabled:
                    if not self.filename_portion:
                        raise ValueError("Filename portion is required in Save Quotes mode.")

                    # Extract or prompt for company name
                    company_name = self.scan_document_for_company_names(file_path)
                    if not company_name:
                        company_name = self.prompt_for_company_name()
                        if not company_name:
                            return  # User canceled

                    # Format filename correctly
                    new_filename = f"{self.filename_portion} - {company_name}.pdf"

                # Convert Word files to PDF
                if file_ext in ['.doc', '.docx']:
                    converted_path = self.pdf_ops.convert_to_pdf(file_path, active_save_dir, new_filename)
                    self._on_file_processed(converted_path)
                else:  # PDF files
                    saved_path = self.file_ops.save_file(file_path, active_save_dir, new_filename)
                    self._on_file_processed(saved_path)

            else:
                raise ValueError(f"Unsupported file type: {file_ext}")

            self.ui_components.show_progress(20)

        except Exception as e:
            logging.error(f"Error processing file: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Processing Error", str(e))
            self.update_drag_drop_text("An error occurred. Please try again.")
        finally:
            self.ui_components.hide_progress()

    def _on_file_processed(self, new_path):
        """Handle updates after a file is processed."""
        try:
            if not isinstance(new_path, str) or not os.path.exists(new_path):
                raise ValueError(f"Invalid file path: {new_path}")

            self.current_file = new_path
            self.ui_components.update_recent_files(new_path)

            self.ui_components.set_label_text(f"File processed: {new_path}")
            logging.info(f"File successfully processed: {new_path}")
            
            # Add to processed files set
            self.processed_files.add(new_path)
            
        except Exception as e:
            logging.error(f"Error handling processed file: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Processing Error", str(e))
    
    def copy_file_to_save_directory(self, source_path, target_file_name):
        """
        Copy a file to the default save directory.

        :param source_path: Path of the source file to copy.
        :param target_file_name: The name of the file in the save directory.
        :return: The full path of the copied file.
        """
        try:
            if not os.path.exists(source_path):
                raise FileNotFoundError(f"Source file does not exist: {source_path}")

            target_path = os.path.join(self.default_save_dir, target_file_name)
            shutil.copy(source_path, target_path)
            logging.info(f"File copied to save directory: {target_path}")
            return target_path
        except Exception as e:
            logging.error(f"Error copying file: {str(e)}", exc_info=True)
            raise

    def replace_file(self, source_path, target_path):
        """
        Replace an existing file at the target location with a new file.

        :param source_path: Path of the source file.
        :param target_path: Path where the file will be replaced.
        :return: None
        """
        try:
            if not os.path.exists(source_path):
                raise FileNotFoundError(f"Source file does not exist: {source_path}")

            if os.path.exists(target_path):
                logging.info(f"Replacing file: {target_path}")
                os.remove(target_path)

            shutil.copy(source_path, target_path)
            logging.info(f"File replaced successfully at: {target_path}")
        except Exception as e:
            logging.error(f"Error replacing file: {str(e)}", exc_info=True)
            raise

    def handle_outlook_drop(self, mime_data):
        """Handle the drop event for Outlook items asynchronously."""
        try:
            # Show busy cursor and update UI to indicate processing
            self.ui_components.set_label_text("Processing Outlook item...")
            self.setCursor(Qt.CursorShape.BusyCursor)

            # Ensure the worker is cleaned up after use
            if hasattr(self, "outlook_worker") and self.outlook_worker.isRunning():
                self.outlook_worker.terminate()
                self.outlook_worker.wait()

            # Initialize the worker thread
            self.outlook_worker = OutlookWorker(
                mime_data=mime_data,
                file_ops=self.file_ops,
                pdf_ops=self.pdf_ops,
                resource_manager=self.resource_manager,
            )

            # Connect signals to handle results
            self.outlook_worker.finished.connect(self.on_outlook_processed)
            self.outlook_worker.error.connect(self.on_outlook_error)

            # Start the worker thread
            self.outlook_worker.start()
            logging.info("OutlookWorker thread started for processing drop event.")

        except Exception as e:
            logging.error(f"Error initializing Outlook drop processing: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Error", f"Unable to process Outlook item: {str(e)}")
        finally:
            # Restore cursor to default
            self.setCursor(Qt.CursorShape.ArrowCursor)


    def on_outlook_processed(self, temp_file_path):
        """Handle successful processing of Outlook items."""
        try:
            if not temp_file_path or not os.path.exists(temp_file_path):
                raise ValueError("Temporary file does not exist or could not be processed")

            # Process the temporary file
            self.process_dropped_file(temp_file_path)

            # Clean up temporary file
            try:
                os.remove(temp_file_path)
                logging.info(f"Temporary file removed: {temp_file_path}")
            except Exception as e:
                logging.warning(f"Error cleaning up temporary file: {str(e)}")

            # Update the UI to show success
            self.ui_components.set_label_text(f"File saved: {temp_file_path}")

        except Exception as e:
            logging.error(f"Error during post-processing: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Error", f"Failed to process the file: {str(e)}")
        finally:
            self.setCursor(Qt.CursorShape.ArrowCursor)

    def on_outlook_error(self, error_message):
        """Handle errors that occur during Outlook processing."""
        logging.error(f"Outlook processing error: {error_message}")
        self.ui_components.show_error_message("Outlook Error", error_message)
        self.ui_components.set_label_text("Failed to process Outlook item.")
        self.setCursor(Qt.CursorShape.ArrowCursor)


    def process_outlook_email(self, email):
        """Process an Outlook email and its attachments."""
        try:
            logging.debug(f"Processing email: {email.Subject}")
            self.ui_components.set_label_text(f"Processing email: {email.Subject}")
            
            attachments = self.outlook_handler.get_email_attachments(email)
            if not attachments:
                message = "No attachments found in the email."
                self.ui_components.set_label_text(message)
                logging.info(message)
                return

            temp_folder = self.file_ops.create_temp_folder()
            try:
                last_processed_path = None
                for attachment in attachments:
                    temp_path = self.outlook_handler.save_attachment(
                        attachment, 
                        temp_folder
                    )
                    if temp_path:
                        self.process_dropped_file(temp_path)
                        last_processed_path = self.current_file  # Store the path of the last processed file
                        try:
                            os.remove(temp_path)
                        except Exception as e:
                            logging.warning(f"Error removing temporary file: {str(e)}")
            finally:
                self.file_ops.cleanup_temp_folder(temp_folder)

        except Exception as e:
            error_msg = f"Error processing email: {str(e)}"
            logging.error(error_msg, exc_info=True)
            raise Exception(error_msg)

    def scan_document_for_company_names(self, file_path):
        """Extract company name from the document."""
        try:
            # Only scan if in Save Quotes mode
            if not self.filename_portions_enabled:
                return None

            text_content = self.file_ops.extract_text_from_file(file_path)
            company_names = self.file_ops.load_company_names()

            # Match extracted text against known company names
            for name in company_names:
                if name.lower() in text_content.lower():
                    # Found a match - ask for confirmation
                    confirmed = self.ui_components.get_confirmation(
                        "Confirm Company Name",
                        f"Found company name: {name}\nIs this correct?"
                    )
                    if confirmed:
                        logging.info(f"Company name confirmed: {name}")
                        return name
                    break  # Break if user rejected the found name

            # If no match found or match was rejected, prompt for new name
            logging.info("Company name not found or rejected. Prompting user.")
            return None
        except Exception as e:
            logging.error(f"Error extracting company name: {e}", exc_info=True)
            return None

        
    def detect_new_company_names(self, text):
        """Detect potential new company names based on text patterns."""
        new_names = set()
        for line in text.splitlines():
            # Example pattern for detecting company names (adjust as needed)
            if "Company:" in line:
                name = line.split("Company:", 1)[1].strip()
                new_names.add(name)
        return new_names

    def prompt_for_company_name(self):
        """Prompt the user to enter a company name with proper activation."""
        logging.debug("Prompting user for company name")

        # Create and configure the dialog
        dialog = QInputDialog(self)
        dialog.setWindowTitle('Company Name')
        dialog.setLabelText('Enter the company name:')
        dialog.setTextValue("")  # Default empty text

        # Ensure the dialog is fully activated
        dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        dialog.activateWindow()
        dialog.raise_()

        # Access the input field and enforce focus
        input_field = dialog.findChild(QLineEdit)
        if input_field:
            def enforce_focus():
                dialog.activateWindow()  # Ensure the dialog is active
                input_field.setFocus()
                input_field.selectAll()  # Optional: Highlight any default text
            QTimer.singleShot(0, enforce_focus)

        # Execute the dialog and block until input is complete
        if dialog.exec():
            company_name = dialog.textValue().strip()
            if company_name:
                logging.debug(f"User entered company name: {company_name}")
                self.file_ops.add_company_name(company_name)
                self.company_names = self.file_ops.load_company_names()
                return company_name
        else:
            logging.debug("User canceled the company name input")
            return ""
        
    def validate_file_path(file_path):
        """
        Validate if a file exists at the specified path.

        :param file_path: Path to the file to validate.
        :return: True if the file exists, False otherwise.
        """
        try:
            if not os.path.exists(file_path):
                logging.error(f"File not found: {file_path}")
                return False
            return True
        except Exception as e:
            logging.error(f"Error validating file path: {file_path}, {str(e)}")
            return False

    
    def handle_task_completed(self, task_type, result):
        """
        Handle task completion from the BackgroundProcessor.
        """
        logging.debug(f"Task completed. Type: {task_type}, Result: {result} ({type(result)})")
        try:
            # Validate the result
            if isinstance(result, str) and os.path.exists(result):
                # Update recent files in the UI
                self.ui_components.update_recent_files(result)

                logging.info(f"Task completed successfully: {result}")
                self.ui_components.set_label_text(f"Task completed: {os.path.basename(result)}")

            else:
                # Handle invalid task results
                error_message = f"Task completed with an invalid result: {result}. Expected a valid file path."
                logging.error(error_message)
                self.ui_components.show_error_message("Invalid Task Result", error_message)

        except Exception as e:
            # Log and display unexpected errors
            logging.error(f"Error handling completed task: {str(e)}", exc_info=True)
            self.ui_components.show_error_message("Task Completion Error", f"An unexpected error occurred: {str(e)}")
        finally:
            # Ensure UI updates and cleanup are consistent
            self.ui_components.hide_progress()


    def handle_task_failed(self, task_type, error_message):
        """Handle task failures from the BackgroundProcessor."""
        logging.error(f"Task failed: {task_type} with error: {error_message}")
        self.ui_components.show_error_message("Task Failed", error_message)

    def get_version_file_path(self):
        """Dynamically determine the user's OneDrive path for the version file."""
        user_home = Path.home()
        return str(user_home / "OneDrive - Fendler Patterson Construction, Inc" /
                "Documents - FP Shared" / "Estimates" / "Misc Downloads" / "version.txt")

    def get_update_exe_path(self):
        """Dynamically determine the user's OneDrive path for the update file."""
        user_home = Path.home()
        return str(user_home / "OneDrive - Fendler Patterson Construction, Inc" /
                "Documents - FP Shared" / "Estimates" / "Misc Downloads" / "DocHandler.exe")

    def get_current_version(self):
        """Retrieve the current version of the app."""
        return "2.0.0"  # Update this each time a new version is built

    def check_for_updates(self):
        """Check if a new version is available."""
        version_file_path = self.get_version_file_path()

        try:
            if not os.path.exists(version_file_path):
                QMessageBox.warning(self, "Update Error", "Version file not found. Check OneDrive settings.")
                return

            # Read the latest version from the file
            with open(version_file_path, "r", encoding="utf-8") as f:
                latest_version = f.read().strip()

            current_version = self.get_current_version()

            if latest_version > current_version:
                reply = QMessageBox.question(self, "Update Available",
                                            f"A new version {latest_version} is available.\n"
                                            "Do you want to update now?",
                                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

                if reply == QMessageBox.StandardButton.Yes:
                    self.download_update()

            else:
                QMessageBox.information(self, "Up to Date", "You already have the latest version.")

        except Exception as e:
            QMessageBox.warning(self, "Update Error", f"Could not check for updates: {e}")

    def download_update(self):
        """Download and replace the old executable."""
        update_exe_path = self.get_update_exe_path()

        try:
            if not os.path.exists(update_exe_path):
                QMessageBox.warning(self, "Update Error", "Update file not found in OneDrive.")
                return

            new_exe_path = os.path.join(os.path.dirname(sys.executable), "DocHandler_new.exe")

            # Copy new file from OneDrive
            shutil.copy2(update_exe_path, new_exe_path)

            # Replace the running executable with the new version
            update_script = os.path.join(os.path.dirname(sys.executable), "update_script.bat")

            with open(update_script, "w") as f:
                f.write(f"""
    @echo off
    timeout /t 2
    del "{sys.executable}" /Q
    move "{new_exe_path}" "{sys.executable}"
    start "" "{sys.executable}"
    del "{update_script}" /Q
    """)

            # Run the update script and close the app
            subprocess.Popen(["cmd.exe", "/c", update_script], shell=True)
            sys.exit(0)

        except Exception as e:
            QMessageBox.warning(self, "Update Failed", f"Could not update: {e}")


class AutoFocusInputDialog(QInputDialog):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.line_edit = self.findChild(QLineEdit)

    def showEvent(self, event):
        super().showEvent(event)
        if self.line_edit:
            # Ensure focus is set after the dialog is fully displayed
            QTimer.singleShot(0, self.line_edit.setFocus)

    def focusInEvent(self, event):
        super().focusInEvent(event)
        if self.line_edit:
            # Ensure focus remains on the input box
            self.line_edit.setFocus()

class ScrollableDialog(QDialog):
    def __init__(self, title, content, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(600, 400)  # Set the desired width and height

        # Main layout
        layout = QVBoxLayout(self)

        # Create a scroll area
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)

        # Create a content widget for the scroll area
        content_widget = QFrame()
        content_layout = QVBoxLayout(content_widget)

        # Add the content as a QLabel
        label = QLabel(content)
        label.setWordWrap(True)  # Enable text wrapping
        content_layout.addWidget(label)

        # Set the content widget in the scroll area
        scroll_area.setWidget(content_widget)
        layout.addWidget(scroll_area)

        self.setLayout(layout)
    

if __name__ == '__main__':
    import sys
    from PyQt6.QtWidgets import QApplication

    # Create the application
    app = QApplication(sys.argv)
    
    try:
        # Create and show the main window
        window = DocHandlerApp()
        window.show()
        
        # Start the application event loop
        sys.exit(app.exec())
    except Exception as e:
        logging.critical(f"Application failed to start: {str(e)}", exc_info=True)
        QMessageBox.critical(None, "Fatal Error", 
                           f"Application failed to start: {str(e)}\n\nPlease check the log file for details.")
        sys.exit(1)
