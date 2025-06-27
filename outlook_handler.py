import win32com.client
import pythoncom
import tempfile
from io import BytesIO
import logging
import os
from contextlib import contextmanager
from config import CONFIG
from PyQt6.QtCore import Qt, QByteArray, QThread, pyqtSignal
import time

class OutlookWorker(QThread):
    finished = pyqtSignal(str)  # Signal for successfully processed file
    error = pyqtSignal(str)     # Signal for any errors

    def __init__(self, mime_data, file_ops, pdf_ops, resource_manager=None):
        super().__init__()
        self.mime_data = mime_data
        self.file_ops = file_ops
        self.pdf_ops = pdf_ops
        self.resource_manager = resource_manager

    def run(self):
        """Run Outlook processing in a separate thread."""
        try:
            handler = OutlookHandler(self.file_ops, self.pdf_ops, self.resource_manager)
            temp_file_path = handler.get_outlook_item(self.mime_data)
            if temp_file_path:
                self.finished.emit(temp_file_path)
            else:
                self.error.emit("Failed to process Outlook item")
        except Exception as e:
            self.error.emit(str(e))


class OutlookHandler:
    def __init__(self, file_ops, pdf_ops, resource_manager=None):
        self.file_ops = file_ops
        self.pdf_ops = pdf_ops
        self.resource_manager = resource_manager

    @contextmanager
    def get_outlook_instance(self):
        """Create and manage an Outlook instance."""
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.DispatchEx('Outlook.Application')
            mapi = outlook.GetNamespace("MAPI")
            # Test connection to ensure Outlook is responsive
            mapi.GetDefaultFolder(6)  # 6 = olFolderInbox
            yield mapi
        except pythoncom.com_error as e:
            logging.error(f"COM Error creating Outlook instance: {e}")
            raise ValueError("Unable to connect to Outlook - COM error")
        except Exception as e:
            logging.error(f"Error creating Outlook instance: {e}")
            raise ValueError("Unable to connect to Outlook")
        finally:
            try:
                outlook = None
                pythoncom.CoUninitialize()
            except:
                pass

    def get_outlook_item(self, mime_data):
        if not self.resource_manager:
            raise RuntimeError("ResourceManager not initialized")

        try:
            # Create temp directory if it doesn't exist
            temp_dir = tempfile.gettempdir()
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)

            if mime_data.hasFormat("application/x-qt-windows-mime;value=\"Ole Object\""):
                return self._handle_ole_object(mime_data)
            elif mime_data.hasFormat("FileGroupDescriptor") and mime_data.hasFormat("FileContents"):
                return self._handle_file_content(mime_data)
            else:
                raise ValueError("Unsupported Outlook item format")
        except Exception as e:
            logging.error(f"Error processing Outlook item: {str(e)}", exc_info=True)
            raise ValueError(f"Failed to process Outlook attachment: {str(e)}")



    def _handle_ole_object(self, mime_data):
        """Handle item dragged directly from Outlook."""
        with self.get_outlook_instance() as outlook:
            ole_data = mime_data.data("application/x-qt-windows-mime;value=\"Ole Object\"")
            if not ole_data:
                raise ValueError("Empty OLE object data")

            # Handle QByteArray properly
            if isinstance(ole_data, QByteArray):
                ole_bytes = ole_data.data()
            else:
                ole_bytes = ole_data

            try:
                # Try decoding with null termination handling
                entry_id = ole_bytes.decode('utf-16le').split('\x00')[0]
                item = outlook.GetItemFromID(entry_id)
                
                if not item or not hasattr(item, 'Attachments') or item.Attachments.Count == 0:
                    raise ValueError("No valid attachments found")

                attachment = item.Attachments.Item(1)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".tmp") as temp_file:
                    temp_file.write(attachment.Content)
                    return temp_file.name

            except Exception as e:
                logging.error(f"Error handling OLE object: {str(e)}")
                raise ValueError(f"Failed to process Outlook attachment: {str(e)}")



    def _handle_file_content(self, mime_data):
        """Handle file content data."""
        contents = mime_data.data("FileContents")
        descriptor = mime_data.data("FileGroupDescriptor")
        
        if not contents or not descriptor:
            raise ValueError("Missing file data")

        try:
            filename = self._extract_filename_from_descriptor(descriptor)
            content_data = contents.data()
            
            # Create a non-temporary file path that won't be deleted
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, filename or f"document_{int(time.time())}.pdf")
            
            with open(temp_path, 'wb') as f:
                f.write(content_data)
            
            return temp_path

        except Exception as e:
            logging.error(f"Error saving file content: {str(e)}")
            raise

    def _extract_filename_from_descriptor(self, descriptor):
        """Extract filename from descriptor data."""
        desc_data = descriptor.data()
        
        for encoding in ['utf-16le', 'utf-8', 'ascii']:
            try:
                text = desc_data.decode(encoding)
                parts = text.split('\x00')
                for part in parts:
                    if part.lower().endswith(('.pdf', '.doc', '.docx')):
                        return self.file_ops.sanitize_filename(part)
            except:
                continue
        
        return None


    def is_email(self, outlook_item):
        """Check if item is an email."""
        return getattr(outlook_item, 'Class', None) == 43

    def is_attachment(self, outlook_item):
        """Check if item is an attachment."""
        return hasattr(outlook_item, 'Parent') and getattr(outlook_item.Parent, 'Name', '') == "Attachments"

    def save_attachment(self, attachment, save_dir, filename_portion=None):
        """Save an attachment to specified directory."""
        try:
            if not hasattr(attachment, 'FileName') or not hasattr(attachment, 'SaveAsFile'):
                raise ValueError("Invalid attachment object")

            # Get filename
            file_name = self.file_ops.sanitize_filename(attachment.FileName)
            if filename_portion:
                file_name = f"{filename_portion} - {file_name}"

            # Save file
            save_path = self.file_ops.get_unique_filename(save_dir, file_name)
            attachment.SaveAsFile(save_path)
            logging.info(f"Saved attachment: {save_path}")

            # Convert Word to PDF if needed
            if save_path.lower().endswith(('.doc', '.docx')):
                pdf_path = self.pdf_ops.convert_word_to_pdf(save_path, save_dir, file_name)
                os.remove(save_path)
                return pdf_path

            return save_path

        except Exception as e:
            logging.error(f"Error saving attachment: {str(e)}")
            raise