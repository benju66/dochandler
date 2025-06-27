from PyQt6.QtWidgets import QDialog, QVBoxLayout, QListWidget, QPushButton, QHBoxLayout, QInputDialog, QMessageBox, QLineEdit
from PyQt6.QtCore import Qt, QTimer

class EditListDialog(QDialog):
    def __init__(self, file_path, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.file_path = file_path
        self.parent = parent
        self.file_ops = parent.file_ops

        # Main layout
        layout = QVBoxLayout(self)

        # List widget for scrollable display
        self.list_widget = QListWidget()
        self.load_items()
        layout.addWidget(self.list_widget)

        # Buttons
        button_layout = QHBoxLayout()
        add_button = QPushButton("Add")
        edit_button = QPushButton("Edit")
        delete_button = QPushButton("Delete")
        add_button.clicked.connect(self.add_item)
        edit_button.clicked.connect(self.edit_item)
        delete_button.clicked.connect(self.delete_item)
        button_layout.addWidget(add_button)
        button_layout.addWidget(edit_button)
        button_layout.addWidget(delete_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def load_items(self):
        """Load items from the file into the list widget in alphabetical order."""
        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                items = sorted(line.strip() for line in file if line.strip())
                self.list_widget.addItems(items)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not load items: {str(e)}")

    def save_items(self):
        """Save current list items back to the file in alphabetical order."""
        try:
            items = []
            for i in range(self.list_widget.count()):
                items.append(self.list_widget.item(i).text())
            
            items.sort()
            with open(self.file_path, 'w', encoding='utf-8') as file:
                for item in items:
                    file.write(f"{item}\n")
            
            # Reload lists in main application
            if "company_names.txt" in self.file_path:
                self.parent.company_names = self.parent.file_ops.load_company_names()
            elif "file_name_portions.txt" in self.file_path:
                self.parent.portions_list = self.parent.file_ops.load_file_name_portions()
                self.parent.ui_components.update_filename_portions_widget(self.parent.portions_list)
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not save items: {str(e)}")

    def add_item(self):
        """Add a new item to the list."""
        dialog = QInputDialog(self)
        dialog.setWindowTitle("Add Item")
        dialog.setLabelText("Enter new item:")
        dialog.setTextValue("")

        # Access the QLineEdit and set focus after the dialog is fully displayed
        input_field = dialog.findChild(QLineEdit)
        if input_field:
            QTimer.singleShot(0, lambda: input_field.setFocus())

        # Execute the dialog
        ok = dialog.exec()
        text = dialog.textValue()
        if ok and text:
            self.list_widget.addItem(text)
            self.save_items()
            self.refresh_list()


    def edit_item(self):
        """Edit the selected item in the list."""
        current_item = self.list_widget.currentItem()
        if current_item:
            text, ok = QInputDialog.getText(self, "Edit Item", "Edit selected item:", text=current_item.text())
            if ok and text:
                current_item.setText(text)
                self.save_items()
                self.refresh_list()

    def delete_item(self):
        """Delete the selected item from the list."""
        current_item = self.list_widget.currentItem()
        if current_item:
            confirm = QMessageBox.question(self, "Delete Item", "Are you sure you want to delete this item?",
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if confirm == QMessageBox.StandardButton.Yes:
                self.list_widget.takeItem(self.list_widget.row(current_item))
                self.save_items()
                self.refresh_list()

    def refresh_list(self):
        """Reloads and sorts items in the list widget after modification."""
        self.list_widget.clear()
        self.load_items()