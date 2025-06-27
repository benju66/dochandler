import os  # For file operations
import subprocess
import logging
from functools import partial

from PyQt6.QtWidgets import (
    QVBoxLayout,
    QLabel,
    QListWidget,
    QLineEdit,
    QFrame,
    QProgressBar,
    QMenu,
    QMenuBar,
    QGroupBox,
    QPushButton,
    QWidget,
    QHBoxLayout,
    QMessageBox,
    QInputDialog,
    QListWidgetItem,
)
from PyQt6.QtCore import (
    Qt,
    pyqtSignal,
    QPropertyAnimation,
    QRect,
    QSize,
    QTimer,
)
from PyQt6.QtGui import QAction, QCursor, QGuiApplication

# Worker thread for Word-to-PDF conversion
from workers import WordToPDFWorker  # Import from workers.py to avoid circular imports


# Theme Definitions (add it here)
THEMES = {
    "light": {
        "background": "#f0f2f5",
        "text": "black",
        "button_bg": "#ffffff",
        "button_hover": "#e0e0e0",
    },
    "dark": {
        "background": "#1e1e1e",
        "text": "#d4d4d4",
        "button_bg": "#3a3d41",
        "button_hover": "#505355",
    },
}

class DragDropArea(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

        # Initialize attributes
        self.pending_updates = []  # Queue for status updates
        self.update_timer = QTimer()  # Timer for batching updates
        self.update_timer.setInterval(500)  # Update every 500ms
        self.update_timer.timeout.connect(self.flush_updates)

        # Initialize the label
        self.label = QLabel(
            "Drag and drop files here:\n"
            "- Outlook email attachments\n"
            "- PDFs\n"
            "- Word documents (.doc/.docx)",
            self
        )
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setWordWrap(True)
        self.label.setStyleSheet("font-size: 14px; color: #555; padding: 10px;")

        # Set up the layout
        layout = QVBoxLayout(self)
        layout.addWidget(self.label)

        # Apply light mode styling initially
        self.apply_light_mode()


    def set_status(self, text):
        """Queue status updates and flush them periodically."""
        self.pending_updates.append(text)
        if not self.update_timer.isActive():
            self.update_timer.start()

    def flush_updates(self):
        """Apply the latest status update and clear the queue."""
        if self.pending_updates:
            self.label.setText(self.pending_updates[-1])
            self.pending_updates.clear()
        self.update_timer.stop()

    def apply_light_mode(self):
        """Set the DragDropArea style for light mode."""
        self.setStyleSheet("""
            DragDropArea {
                background-color: #f8f8f8;
                min-height: 150px;
                border: 2px dashed #aaa;
                border-radius: 10px;
                margin: 10px;
            }
            DragDropArea:hover {
                background-color: #e8e8e8;
            }
        """)
        self.label.setStyleSheet("color: #333;")

    def apply_dark_mode(self):
        """Set the DragDropArea style for dark mode."""
        self.setStyleSheet("""
            DragDropArea {
                background-color: #2b2b2b;
                min-height: 150px;
                border: 2px dashed #555;
                border-radius: 10px;
                margin: 10px;
            }
            DragDropArea:hover {
                background-color: #3a3a3a;
            }
        """)
        self.label.setStyleSheet("color: #ddd;")

    def show_context_menu(self, position):
        menu = QMenu()
        clear_action = menu.addAction("Clear Drop Area")
        action = menu.exec(self.mapToGlobal(position))
        
        if action == clear_action:
            self.label.setText("Drag and drop files here:\n- Outlook email attachments\n- PDFs\n- Word documents (.doc/.docx)")


class UIComponents:
    def __init__(self, parent):
        self.parent = parent
        self.layout = QVBoxLayout()

        # Initialize attributes that might be accessed later
        self.base_width = 400
        self.resize_animation = None

        # Create all UI elements first
        self.create_ui_elements()
        

        # Default states
        self.auto_convert_enabled = False
        self.recent_portions_visible = False
        self.recent_files_visible = False

        # Initialize visibility after creating components
        self.setup_visibility()

        # Set up button styles
        self.setup_button_styles()

        #Lazy Load
        self.setup_lazy_loading()


    def create_ui_elements(self):
        # Create menu bar and menus
        self.menu_bar = QMenuBar(self.parent)
        self.file_menu = QMenu("File", self.parent)
        self.processing_menu = QMenu("Processing", self.parent)
        self.view_menu = QMenu("View", self.parent)
        self.help_menu = QMenu("Help", self.parent)
        self.edit_lists_menu = QMenu("Edit Lists", self.parent)
        self.recent_locations_menu = QMenu("Recent Locations", self.parent)

        # File Menu
        self.save_location_action = QAction("Set Save Location", self.parent)
        self.default_save_location_action = QAction("Set Default Location", self.parent)
        self.toggle_auto_convert_action = QAction("Auto Convert", self.parent, checkable=True)
        self.exit_action = QAction("Exit", self.parent)
        self.check_update_action = QAction("Check for Updates", self.parent)
        self.file_menu.addAction(self.check_update_action)


        # Processing Menu
        self.enable_filename_portions_action = QAction("Save Quotes Mode", self.parent, checkable=True)
        self.edit_company_names_action = QAction("Edit Company Names", self.parent)
        self.edit_file_name_portions_action = QAction("Edit Scope of Work", self.parent)

        # View Menu
        self.toggle_dark_mode_action = QAction("Dark Mode", self.parent, checkable=True)
        self.toggle_pending_files_action = QAction("Pending Files Panel", self.parent, checkable=True)
        self.toggle_recent_files_action = QAction("Recent Files Panel", self.parent, checkable=True)
        self.toggle_recent_portions_action = QAction("Recent Scopes of Work List", self.parent, checkable=True)
        self.debug_action = QAction("Debug Mode", self.parent, checkable=True)

        # Help Menu
        self.help_action = QAction("Quick Guide", self.parent)
        self.about_action = QAction("About", self.parent)

        # Create UI components
        self.drag_drop_area = DragDropArea(self.parent)
        self.search_bar = QLineEdit(self.parent)
        self.search_bar.setPlaceholderText("Search filename portions...")
        self.search_bar.setClearButtonEnabled(True)

        self.portions_list_widget = QListWidget(self.parent)
        self.progress_bar = QProgressBar(self.parent)
        self.recent_files_list = QListWidget(self.parent)
        self.recent_files_label = QLabel("Recent Files", self.parent)
        self.recent_portions_list_widget = QListWidget(self.parent)
        self.clear_recent_portions_button = QPushButton("Clear Recent Portions", self.parent)
        self.save_button = QPushButton("Save", self.parent)
        self.save_location_label = ClickablePathLabel(self.parent)

        # Create pending files list with context menu
        self.add_pending_files_components()
        self.pending_files_list = QListWidget(self.parent)
        self.pending_files_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.pending_files_list.customContextMenuRequested.connect(self.show_context_menu)
        self.pending_files_list.setMaximumHeight(150)

        # Initialize groups
        self.filename_portions_group = QGroupBox("Scope of Work", self.parent)
        self.recent_portions_group = QGroupBox("Recently Used Portions", self.parent)
        self.recent_files_group = QGroupBox("Recent Files", self.parent)

        # Set up group layouts
        filename_portions_layout = QVBoxLayout()
        filename_portions_layout.addWidget(self.search_bar)
        filename_portions_layout.addWidget(self.portions_list_widget)
        self.filename_portions_group.setLayout(filename_portions_layout)

        recent_portions_layout = QVBoxLayout()
        recent_portions_layout.addWidget(self.recent_portions_list_widget)
        recent_portions_layout.addWidget(self.clear_recent_portions_button)
        self.recent_portions_group.setLayout(recent_portions_layout)

        recent_files_layout = QVBoxLayout()
        recent_files_layout.addWidget(self.recent_files_label)
        recent_files_layout.addWidget(self.recent_files_list)
        self.recent_files_group.setLayout(recent_files_layout)

    def setup_button_styles(self):
        """Set up styles for buttons."""
        self.clear_recent_portions_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: black;
                padding: 5px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)

    def setup_ui(self):
        """Initialize the user interface layout and components."""
        self.setup_menu()  # Set up the menu bar and menu actions

        # Configure layout properties
        self.layout.setSpacing(8)
        self.layout.setContentsMargins(5, 5, 5, 5)

        # Add menu bar to the layout
        self.menu_bar.setStyleSheet("QMenuBar {padding-top: 0px; padding-bottom: 0px;}")
        self.layout.addWidget(self.menu_bar)

        # Add save location label and drag-and-drop area
        self.layout.addWidget(self.save_location_label)
        self.layout.addSpacing(5)
        self.layout.addWidget(self.drag_drop_area)

        self.pending_files_group = QGroupBox("Files", self.parent)
        pending_layout = QVBoxLayout()
        pending_layout.addWidget(self.pending_files_list)
        self.pending_files_group.setLayout(pending_layout)
        self.layout.addWidget(self.pending_files_group)
    
        # Configure and add "Filename Portions" group (Scope of Work)
        self.filename_portions_group = QGroupBox("Scope of Work")
        filename_portions_layout = QVBoxLayout()
        filename_portions_layout.addWidget(self.search_bar)
        filename_portions_layout.addWidget(self.portions_list_widget)
        self.filename_portions_group.setLayout(filename_portions_layout)
        self.layout.addWidget(self.filename_portions_group)

        # Configure and add "Recently Used Portions" group
        self.recent_portions_group = QGroupBox("Recently Used Portions")
        recent_portions_layout = QVBoxLayout()
        recent_portions_layout.addWidget(self.recent_portions_list_widget)
        recent_portions_layout.addWidget(self.clear_recent_portions_button)
        self.recent_portions_group.setLayout(recent_portions_layout)
        self.layout.addWidget(self.recent_portions_group)

        # Configure and add "Recent Files" group
        self.recent_files_group = QGroupBox()
        recent_files_layout = QVBoxLayout()
        recent_files_layout.addWidget(self.recent_files_label)
        recent_files_layout.addWidget(self.recent_files_list)
        self.recent_files_group.setLayout(recent_files_layout)
        self.layout.addWidget(self.recent_files_group)

        # Add the save button and progress bar
        self.layout.addWidget(self.save_button)
        self.layout.addWidget(self.progress_bar)

        # Ensure extra space is added at the bottom of the layout
        self.layout.addStretch(1)

        # Set the layout for the parent widget
        self.parent.setLayout(self.layout)
        self.parent.setWindowTitle("DocHandler - Outlook Email/Attachment & Word Document Handler")

        # Handle visibility and styles in one place
        self.setup_visibility()  # Initial visibility states
        self.setup_tooltips()    # Add tooltips to components
        self.apply_styles()      # Apply consistent styles to the UI

        # Ensure the window size is updated based on content
        QTimer.singleShot(100, self.update_window_size)

        # Set save button visibility last, after all other setup is done
        self.save_button.setVisible(not self.auto_convert_enabled)


    def setup_menu(self):
        self.menu_bar.clear()
        
        # Apply menu styling
        menu_style = """
            QMenu {
                background-color: #1e1e1e;
                color: #d4d4d4;
                border: 1px solid #3c3c3c;
            }
            QMenu::item {
                background-color: #1e1e1e;
                color: #d4d4d4;
            }
            QMenu::item:selected {
                background-color: #3a3d41;
            }
        """
        
        for menu in [self.file_menu, self.processing_menu, self.view_menu, 
                    self.help_menu, self.recent_locations_menu, self.edit_lists_menu]:
            menu.setStyleSheet(menu_style)

        # File Menu
        self.file_menu.addAction(self.default_save_location_action)
        self.file_menu.addAction(self.save_location_action)
        self.file_menu.addMenu(self.recent_locations_menu)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.toggle_auto_convert_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.exit_action)

        # Processing Menu
        self.processing_menu.addAction(self.enable_filename_portions_action)
        self.processing_menu.addSeparator()
        self.edit_lists_menu.addAction(self.edit_company_names_action)
        self.edit_lists_menu.addAction(self.edit_file_name_portions_action)
        self.processing_menu.addMenu(self.edit_lists_menu)

        # View Menu
        self.view_menu.addAction(self.toggle_dark_mode_action)
        self.view_menu.addAction(self.toggle_pending_files_action)
        self.toggle_pending_files_action.triggered.connect(self.toggle_pending_files_group)

        self.view_menu.addAction(self.toggle_recent_files_action)
        self.view_menu.addAction(self.toggle_recent_portions_action)
        self.view_menu.addAction(self.debug_action)

        # Help Menu
        self.help_menu.addAction(self.help_action)
        self.help_menu.addAction(self.about_action)

        # Add menus to menu bar
        self.menu_bar.addMenu(self.file_menu)
        self.menu_bar.addMenu(self.processing_menu)
        self.menu_bar.addMenu(self.view_menu)
        self.menu_bar.addMenu(self.help_menu)

    def set_label_text(self, text):
        """Update the label text."""
        self.drag_drop_area.label.setText(f"{text}")
        logging.debug(f"Label text updated to: {text}")

    def setup_lazy_loading(self):
        """Enable lazy loading for recent portions and recent files when the user scrolls."""
        self.recent_portions_list_widget.verticalScrollBar().valueChanged.connect(self.lazy_load_more_portions)
        self.recent_files_list.verticalScrollBar().valueChanged.connect(self.lazy_load_more_files)


    def show_recent_locations_menu(self):
        """Update and display the recent locations submenu."""
        recent_locations = self.parent.file_ops.load_recent_save_locations()
        self.update_recent_locations_menu(recent_locations)

    def update_recent_locations_menu(self, locations):
        """Update the recent save locations submenu."""
        self.recent_locations_menu.clear()
        for location in locations:
            # Create shortened display path
            parts = os.path.normpath(location).split(os.sep)
            display_parts = parts[-4:] if len(parts) >= 5 else parts
            display_path = os.path.join(*display_parts)
            
            action = QAction(display_path, self.parent)
            action.setData(location)  # Store full path as data
            action.triggered.connect(lambda checked, loc=location: self.parent.set_save_location_from_recent(loc))
            self.recent_locations_menu.addAction(action)

    def update_recent_portions_list_widget(self, portions, load_more=False):
        """Update the recent portions list widget with lazy loading."""
        if not load_more:
            self.recent_portions_list_widget.clear()

        for portion in portions:
            self.recent_portions_list_widget.addItem(portion)

        if len(portions) >= 20:  # Show "Load More" button if more items exist
            load_more_button = QListWidgetItem("Load More...")
            load_more_button.setFlags(Qt.ItemFlag.ItemIsEnabled)  # Make it non-selectable
            self.recent_portions_list_widget.addItem(load_more_button)


    def update_filename_portions_widget(self, portions):
        """Update the filename portions list widget with the provided portions."""
        self.portions_list_widget.clear()
        for portion in portions:
            self.portions_list_widget.addItem(portion)
        logging.debug(f"Updated filename portions widget with {len(portions)} items.")
        # Add size update
        QTimer.singleShot(100, self.update_window_size)

    def setup_lazy_loading(self):
        """Enable lazy loading when scrolling in the recent portions list."""
        self.recent_portions_list_widget.verticalScrollBar().valueChanged.connect(self.lazy_load_more_portions)

    def lazy_load_more_portions(self):
        """Load more recent portions when the user scrolls down."""
        scrollbar = self.recent_portions_list_widget.verticalScrollBar()
        if scrollbar.value() == scrollbar.maximum():  # User scrolled to the bottom
            current_count = self.recent_portions_list_widget.count() - 1  # Exclude "Load More"
            more_portions = self.parent.file_ops.load_recent_filename_portions(limit=10)  # Load next 10 items
            self.update_recent_portions_list_widget(more_portions, load_more=True)

    def lazy_load_more_files(self):
        """Load more recent files when the user scrolls down."""
        scrollbar = self.recent_files_list.verticalScrollBar()
        if scrollbar.value() == scrollbar.maximum():  # User reached the bottom
            current_count = self.recent_files_list.count()
            more_files = self.parent.file_ops.load_recent_save_locations(limit=10)  # Load next 10 items
            self.update_recent_files_list(more_files, load_more=True)

    def update_recent_files_list(self, files, load_more=False):
        """Update the recent files list widget with lazy loading."""
        if not load_more:
            self.recent_files_list.clear()

        for file in files:
            self.recent_files_list.addItem(file)

        if len(files) >= 20:  # Show "Load More" button if more items exist
            load_more_button = QListWidgetItem("Load More...")
            load_more_button.setFlags(Qt.ItemFlag.ItemIsEnabled)  # Make it non-selectable
            self.recent_files_list.addItem(load_more_button)


    def clear_filename_portions_widget(self):
        """Clear all items from the filename portions list widget."""
        self.portions_list_widget.clear()
        logging.debug("Cleared filename portions widget.")

    def setup_visibility(self):
        """Set initial visibility of UI components."""
        if hasattr(self, 'pending_files_group'):
            self.pending_files_group.setVisible(False)
            self.toggle_pending_files_action.setChecked(self.pending_files_group.isVisible())
        else:
            logging.error("Pending Files Group is not initialized")
        self.filename_portions_group.setVisible(False)
        self.recent_portions_group.setVisible(self.recent_portions_visible)
        self.progress_bar.setVisible(False)
        self.recent_files_group.setVisible(self.recent_files_visible)
        self.save_button.setVisible(True)

        QTimer.singleShot(100, self.update_window_size)

    def setup_tooltips(self):
        self.save_location_action.setToolTip("Choose where to save processed files")
        self.enable_filename_portions_action.setToolTip("Enable to use predefined file name portions")
        self.search_bar.setToolTip("Search for specific file name portions")

    def add_pending_files_components(self):
        """Add UI components for the pending files list and initialize visibility."""
        self.pending_files_group = QGroupBox("Pending Files", self.parent)
        pending_layout = QVBoxLayout()
        self.pending_files_list = QListWidget(self.parent)
        pending_layout.addWidget(self.pending_files_list)

        self.pending_files_group.setLayout(pending_layout)
        self.pending_files_group.setVisible(False)

        # Insert below the drag-and-drop area
        drag_drop_index = self.layout.indexOf(self.drag_drop_area)
        if drag_drop_index >= 0:
            self.layout.insertWidget(drag_drop_index + 1, self.pending_files_group)


    def toggle_pending_files_group(self):
        """Toggle visibility of the pending files group."""
        if hasattr(self, 'pending_files_group'):
            is_visible = not self.pending_files_group.isVisible()
            self.pending_files_group.setVisible(is_visible)
            self.toggle_pending_files_action.setChecked(is_visible)
            QTimer.singleShot(100, self.update_window_size)
            return is_visible
        return False


    
    def show_context_menu(self, position):
        menu = QMenu()
        remove_action = menu.addAction("Remove File")
        clear_action = menu.addAction("Clear All")
        
        action = menu.exec(self.pending_files_list.mapToGlobal(position))
        if action == remove_action:
            self.remove_selected_file()
        elif action == clear_action:
            self.clear_pending_files()

    def remove_selected_file(self):
        current_item = self.pending_files_list.currentItem()
        if current_item:
            row = self.pending_files_list.row(current_item)
            self.pending_files_list.takeItem(row)
            self.parent.pending_files.pop(row)
            
            # Update save button text
            if self.pending_files_list.count() > 1:
                self.save_button.setText(f"Merge and Save {self.pending_files_list.count()} Files")
            elif self.pending_files_list.count() == 1:
                self.save_button.setText("Save File")
            else:
                self.save_button.setText("Save")

    def clear_pending_files(self):
        self.pending_files_list.clear()
        self.parent.pending_files.clear()
        self.save_button.setText("Save")

    def apply_styles(self):
        self.parent.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                font-family: 'Open Sans', sans-serif;
                color: black;
            }
            QLabel {
                font-size: 15px;
                color: black;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 6px;
                color: black;
                background-color: white;
            }
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 6px;
                margin: 5px 0;
                color: black;
                background-color: white;
            }
            QListWidget::item {
                color: black;
            }
            QPushButton {
                background-color: #45A049;
                color: black;
                padding: 8px 15px;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #45A049;
            }
            QProgressBar {
                text-align: center;
                border: 1px solid #ddd;
                border-radius: 6px;
                color: black;
            }
            QTabWidget::pane {
                border-top: 2px solid #4CAF50;
            }
            QTabBar::tab {
                background: #e0e0e0;
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 6px;
                margin: 2px;
                color: black;
            }
            QTabBar::tab:selected {
                background: #4CAF50;
                color: black;
            }
            QMenuBar {
                color: black;
                background-color: #f0f2f5;
            }
            QMenuBar::item {
                color: black;
                background-color: #f0f2f5;
            }
            QMenu {
                color: black;
                background-color: white;
            }
            QMenu::item {
                color: black;
                background-color: white;
            }
            QGroupBox {
                color: black;
            }
            QGroupBox::title {
                color: black;
            }
        """)

    def calculate_content_width(self):
            """Calculate the required width based on list content."""
            max_width = self.base_width
            
            # Get the font metrics for accurate text width calculation
            font_metrics = self.portions_list_widget.fontMetrics()
            
            # Check width needed for filename portions
            for i in range(self.portions_list_widget.count()):
                item_width = font_metrics.horizontalAdvance(self.portions_list_widget.item(i).text()) + 40  # Add padding
                max_width = max(max_width, item_width)
                
            # Check width needed for recent portions
            for i in range(self.recent_portions_list_widget.count()):
                item_width = font_metrics.horizontalAdvance(self.recent_portions_list_widget.item(i).text()) + 40
                max_width = max(max_width, item_width)
                
            # Add margins and padding
            max_width += 40  # Account for scrollbar and margins
            
            return max_width

    def animate_resize(self, new_width=None, new_height=None, keep_position=True):
        """
        Animate the window resize, supporting both width and height adjustments.
        Optionally keep the current window position unchanged.
        """
        if self.resize_animation:
            self.resize_animation.stop()

        # Get the current geometry
        current_geometry = self.parent.geometry()

        # Calculate new geometry
        x_pos = current_geometry.x() if keep_position else current_geometry.x()
        y_pos = current_geometry.y() if keep_position else current_geometry.y()
        width = new_width if new_width else current_geometry.width()
        height = new_height if new_height else current_geometry.height()

        new_geometry = QRect(x_pos, y_pos, width, height)

        # Create and start the animation
        self.resize_animation = QPropertyAnimation(self.parent, b"geometry")
        self.resize_animation.setDuration(200)  # Animation duration in milliseconds
        self.resize_animation.setStartValue(current_geometry)
        self.resize_animation.setEndValue(new_geometry)
        self.resize_animation.start()


    def calculate_content_height(self):
        """Calculate the required height based on the layout content, excluding hidden widgets."""
        total_height = 0
        for i in range(self.layout.count()):
            item = self.layout.itemAt(i)
            widget = item.widget()
            if widget and widget.isVisible():
                total_height += widget.sizeHint().height()
        return total_height + 40  # Add padding or margins as needed


    def update_window_size(self):
        """Debounced window size update to prevent excessive UI lag."""
        if not hasattr(self, "_resize_timer"):
            self._resize_timer = QTimer(self.parent)
            self._resize_timer.setSingleShot(True)
            self._resize_timer.timeout.connect(self._apply_window_update)

        # Restart the timer on each call, ensuring the resize only happens after 100ms
        self._resize_timer.start(100)

    def _apply_window_update(self):
        """Apply the actual window resizing after debounce delay."""
        new_width = self.calculate_content_width()
        new_height = self.calculate_content_height()

        # Get available screen geometry
        screen_geometry = QGuiApplication.primaryScreen().availableGeometry()
        max_width = screen_geometry.width()
        max_height = screen_geometry.height()

        # Constrain the dimensions to the screen size
        constrained_width = min(new_width, max_width)
        constrained_height = min(new_height, max_height)

        # Animate the resize with constrained dimensions
        self.animate_resize(constrained_width, constrained_height)

    def update_pending_files_list(self, files):
        """Update the pending files list in the UI."""
        valid_files = [f for f in files if os.path.splitext(f)[1].lower() in ['.pdf', '.doc', '.docx']]
        self.pending_files_list.clear()

        for file in valid_files:
            self.pending_files_list.addItem(os.path.basename(file))

        if len(valid_files) > 1:
            self.save_button.setText(f"Merge and Save {len(valid_files)} Files")
        elif valid_files:
            self.save_button.setText("Save File")
        else:
            self.save_button.setText("Save")

        self.pending_files = valid_files
        QTimer.singleShot(100, self.update_window_size)
        logging.debug(f"Pending files updated: {len(valid_files)} valid files.")

    def update_recent_files(self, file_path):
        """Update the recent files list in the UI."""
        if not hasattr(self, "recent_files_list"):
            logging.error("Recent files list is not initialized.")
            return

        if os.path.exists(file_path):
            # Avoid duplicates by removing the existing entry
            items = [self.recent_files_list.item(i).text() for i in range(self.recent_files_list.count())]
            if file_path in items:
                index = items.index(file_path)
                self.recent_files_list.takeItem(index)

            # Add the new file path at the top
            self.recent_files_list.insertItem(0, file_path)
            logging.info(f"Added file to recent files: {file_path}")
        else:
            logging.warning(f"File path does not exist: {file_path}")

        # Only update visibility if the toggle is enabled
        if self.recent_files_group.isVisible():
            self.recent_files_group.setVisible(self.recent_files_list.count() > 0)

        QTimer.singleShot(100, self.update_window_size)




    def reset_progress(self):
        """Reset the progress bar value and visibility."""
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
    
    def show_progress(self, value=0):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(value)

    def hide_progress(self):
        self.progress_bar.setVisible(False)

    def update_portions_list_widget(self, portions):
        self.portions_list_widget.clear()
        for portion in portions:
            self.portions_list_widget.addItem(portion)

    def toggle_recent_files(self):
        self.recent_files_visible = not self.recent_files_visible
        self.recent_files_group.setVisible(self.recent_files_visible)
        toggle_text = "Hide Recent Files" if self.recent_files_visible else "Show Recent Files"
        self.toggle_recent_files_action.setText(toggle_text)
        QTimer.singleShot(100, self.update_window_size)


    def toggle_filename_portions(self, enabled):
        """Toggle visibility of the filename portions and manage related components."""
        self.filename_portions_group.setVisible(enabled)
        
        # Ensure recent portions respect filename portions state
        if enabled:
            self.recent_portions_group.setVisible(self.recent_portions_visible)
        else:
            self.recent_portions_group.setVisible(False)
        
        if not enabled:
            self.search_bar.clear()
        QTimer.singleShot(100, self.update_window_size)

        # Debugging visibility states
        logging.debug(f"filename_portions_group visible: {self.filename_portions_group.isVisible()}")
        logging.debug(f"recent_portions_group visible: {self.recent_portions_group.isVisible()}")

    
    def toggle_recent_portions(self, enabled):
        """Toggle visibility of the 'Recently Used Portions' group."""
        self.recent_portions_visible = enabled

        # Only show if filename portions are enabled and recent portions are toggled on
        if self.filename_portions_group.isVisible():
            self.recent_portions_group.setVisible(enabled)
        else:
            self.recent_portions_group.setVisible(False)

        # Sync action state
        self.toggle_recent_portions_action.setChecked(enabled)
        QTimer.singleShot(100, self.update_window_size)

        # Debugging visibility states
        logging.info(f"'Recently Used Portions' visibility {'enabled' if enabled else 'disabled'}.")
        logging.debug(f"recent_portions_group visible: {self.recent_portions_group.isVisible()}")


    def show_progress(self, value):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(value)

    def hide_progress(self):
        self.progress_bar.setVisible(False)

    def show_error_message(self, title, message):
        QMessageBox.critical(self.parent, title, message)

    def show_info_message(self, title, message):
        QMessageBox.information(self.parent, title, message)

    def show_warning_message(self, title, message):
        QMessageBox.warning(self.parent, title, message)

    def get_confirmation(self, title, message):
        """Display a confirmation dialog and return True if Yes is selected."""
        return QMessageBox.question(
            self.parent, title, message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        ) == QMessageBox.StandardButton.Yes

    def update_recent_portions_list_widget(self, portions):
        """Update or clear the recent portions list widget based on the provided portions list."""
        self.recent_portions_list_widget.clear()
        for portion in portions:
            self.recent_portions_list_widget.addItem(portion)

        # Debugging state and count
        logging.debug(f"Updated recent portions list widget with {len(portions)} items.")
        logging.debug(f"recent_portions_group visible: {self.recent_portions_group.isVisible()}")
        QTimer.singleShot(100, self.update_window_size)


    def get_text_input(self, title, message):
        text, ok = QInputDialog.getText(self.parent, title, message)
        return (text, ok)

    def toggle_theme(self, enabled):
        """Toggle between light and dark mode."""
        # Select the appropriate theme
        theme = THEMES["dark"] if enabled else THEMES["light"]

        # Apply global widget styles
        self.parent.setStyleSheet(f"""
            QWidget {{
                background-color: {theme["background"]};
                color: {theme["text"]};
                font-family: 'Segoe UI', sans-serif;
            }}
            QPushButton {{
                background-color: {theme["button_bg"]};
                color: {theme["text"]};
                border: 1px solid {theme["button_bg"]};
                padding: 8px 15px;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {theme["button_hover"]};
            }}
            QLineEdit {{
                background-color: {theme["background"]};
                color: {theme["text"]};
                border: 1px solid {theme["button_bg"]};
                border-radius: 6px;
                padding: 8px;
            }}
            QListWidget {{
                background-color: {theme["background"]};
                color: {theme["text"]};
                border: 1px solid {theme["button_bg"]};
                border-radius: 6px;
            }}
            QProgressBar {{
                text-align: center;
                background-color: {theme["background"]};
                color: {theme["text"]};
                border: 1px solid {theme["button_bg"]};
                border-radius: 6px;
            }}
            QLabel {{
                color: {theme["text"]};
            }}
            QMenu {{
                background-color: {theme["background"]};
                color: {theme["text"]};
                border: 1px solid {theme["button_bg"]};
            }}
            QMenu::item {{
                background-color: {theme["background"]};
                color: {theme["text"]};
            }}
            QMenu::item:selected {{
                background-color: {theme["button_hover"]};
                color: {theme["text"]};
            }}
        """)

        # Apply mode-specific styles for DragDropArea
        if enabled:
            self.drag_drop_area.apply_dark_mode()
        else:
            self.drag_drop_area.apply_light_mode()

        # Log the change
        logging.info(f"Theme changed to {'Dark' if enabled else 'Light'} mode")

    def convert_word_to_pdf_async(self, doc_path, save_dir, base_name):
        self.worker = WordToPDFWorker(self.file_ops.pdf_ops, doc_path, save_dir, base_name)
        self.worker.progress.connect(self.update_progress)  # Optional: Update a progress bar
        self.worker.finished.connect(self.on_conversion_complete)
        self.worker.error.connect(self.on_conversion_error)
        self.worker.start()

    def update_progress(self, value):
        # Update a progress bar or label
        self.ui_components.set_label_text(f"Conversion Progress: {value}%")

    def on_conversion_complete(self, pdf_path):
        self.ui_components.set_label_text(f"Word converted to PDF: {pdf_path}")

    def on_conversion_error(self, error_message):
        self.ui_components.set_label_text(f"Error: {error_message}")



class ClickablePathLabel(QLabel):
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)  # Add this line
        self.setStyleSheet("""
            QLabel {
                color: #0066cc;
                text-decoration: underline;
                padding: 5px;
                background: transparent;
            }
            QLabel:hover {
                color: #003366;
            }
        """)
        self.full_path = ""
        
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
            if self.full_path and os.path.exists(self.full_path):
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(self.full_path)
                    else:  # macOS and Linux
                        opener = 'open' if os.name == 'darwin' else 'xdg-open'
                        subprocess.run([opener, self.full_path])
                except Exception as e:
                    print(f"Error opening directory: {e}")

    def apply_dark_mode(self):
        self.setStyleSheet("""
            QLabel {
                color: #61afef;
                text-decoration: underline;
                padding: 5px;
                background: transparent;
            }
            QLabel:hover {
                color: #88c0d0;
            }
        """)

    def apply_light_mode(self):
        self.setStyleSheet("""
            QLabel {
                color: #0066cc;
                text-decoration: underline;
                padding: 5px;
                background: transparent;
            }
            QLabel:hover {
                color: #003366;
            }
        """)

    def update_path(self, full_path):
        self.full_path = full_path
        if full_path:
            # Split the path and get the last three parts
            parts = os.path.normpath(full_path).split(os.sep)
            display_parts = parts[-3:] if len(parts) >= 3 else parts
            display_path = os.path.join(*display_parts)
            self.setText(f"Save Location: {display_path}")
        else:
            self.setText("No save location set")