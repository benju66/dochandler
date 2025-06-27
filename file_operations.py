import os
import time
import shutil
import logging
from docx import Document
import PyPDF2
import fitz  # PyMuPDF
from config import CONFIG
from pathlib import Path
import sys
import subprocess
from urllib.parse import unquote
import tempfile
from io import BytesIO
from threading import Lock
filename_lock = Lock()
from pdf_operations import PDFOperations


class FileOperations:
    def __init__(self):
        self.company_names_path = CONFIG['COMPANY_NAMES_PATH']
        self.file_name_portions_path = CONFIG['FILE_NAME_PORTIONS_PATH']
        # Ensure data directory exists
        os.makedirs(os.path.dirname(self.company_names_path), exist_ok=True)
        os.makedirs(os.path.dirname(self.file_name_portions_path), exist_ok=True)
        # Initialize pdf_ops
        self.pdf_ops = PDFOperations(self)
    
    def load_file_name_portions(self):
        """Load filename portions from the data file."""
        try:
            if not os.path.exists(self.file_name_portions_path):
                logging.error(f"File not found: {self.file_name_portions_path}")
                return []
                
            with open(self.file_name_portions_path, 'r', encoding='utf-8') as file:
                portions = [line.strip() for line in file if line.strip()]
            logging.info(f"Loaded {len(portions)} filename portions from {self.file_name_portions_path}")
            return portions
        except Exception as e:
            logging.error(f"Error loading filename portions: {str(e)}", exc_info=True)
            return []

    def load_company_names(self):
        """Load company names from the data file, ensuring alphabetical order."""
        try:
            if not os.path.exists(self.company_names_path):
                logging.warning(f"Company names file not found: {self.company_names_path}")
                return []

            with open(self.company_names_path, 'r', encoding='utf-8') as file:
                names = sorted(line.strip() for line in file if line.strip())

            logging.info(f"Loaded {len(names)} company names.")
            return names
        except Exception as e:
            logging.error(f"Error loading company names: {e}", exc_info=True)
            return []

    def add_company_name(self, company_name):
        """Add a new company name, ensuring no duplicates."""
        try:
            company_name = company_name.strip()
            if not company_name:
                raise ValueError("Company name cannot be empty.")

            current_names = self.load_company_names()
            if company_name.lower() not in (name.lower() for name in current_names):
                current_names.append(company_name)
                current_names.sort()

                with open(self.company_names_path, 'w', encoding='utf-8') as file:
                    for name in current_names:
                        file.write(f"{name}\n")

                logging.info(f"Added new company name: {company_name}")
                return True
            return False
        except Exception as e:
            logging.error(f"Error adding company name: {e}", exc_info=True)
            raise

    
    def load_recent_filename_portions(self, limit=20):
        """Load recent filename portions with a default limit."""
        recent_portions_path = os.path.join(os.path.dirname(self.file_name_portions_path), 'recent_filename_portions.txt')
        try:
            if not os.path.exists(recent_portions_path):
                return []
            
            with open(recent_portions_path, 'r', encoding='utf-8') as file:
                portions = [line.strip() for line in file if line.strip()]
            
            logging.info(f"Loaded {len(portions)} recent portions (showing {limit})")
            return portions[:limit]  # Load only the first `limit` items
        except Exception as e:
            logging.error(f"Error loading recent portions: {str(e)}", exc_info=True)
            return []


    def save_recent_filename_portions(self, portions):
        """Save recent filename portions."""
        recent_portions_path = os.path.join(os.path.dirname(self.file_name_portions_path), 'recent_filename_portions.txt')
        try:
            with open(recent_portions_path, 'w', encoding='utf-8') as file:
                for portion in portions:
                    file.write(f"{portion}\n")
            logging.info(f"Saved {len(portions)} recent portions")
        except Exception as e:
            logging.error(f"Error saving recent portions: {str(e)}", exc_info=True)
            raise

    def load_recent_save_locations(self, limit=20):
        """Load recent save locations with a default limit."""
        try:
            recent_locations_path = CONFIG['RECENT_SAVE_LOCATIONS']
            if not os.path.exists(recent_locations_path):
                return []
            
            with open(recent_locations_path, 'r', encoding='utf-8') as file:
                locations = [line.strip() for line in file if line.strip()]
            
            logging.info(f"Loaded {len(locations)} recent save locations (showing {limit})")
            return locations[:limit]  # Load only the first `limit` items
        except Exception as e:
            logging.error(f"Error loading recent save locations: {str(e)}")
            return []



    def save_recent_save_locations(self, locations):
        """Save recent save locations."""
        try:
            recent_locations_path = CONFIG['RECENT_SAVE_LOCATIONS']
            os.makedirs(os.path.dirname(recent_locations_path), exist_ok=True)
            
            # Limit to 10 most recent locations
            locations = locations[:10]
            
            with open(recent_locations_path, 'w', encoding='utf-8') as file:
                file.writelines([f"{location}\n" for location in locations])
            logging.info(f"Saved {len(locations)} recent locations")
        except Exception as e:
            logging.error(f"Error saving recent save locations: {str(e)}")


    def get_unique_filename(self, directory, filename):
        with filename_lock:
            base_name, ext = os.path.splitext(filename)
            counter = 1
            unique_filename = filename
            file_path = os.path.join(directory, unique_filename)

            while os.path.exists(file_path):
                unique_filename = f"{base_name}_{counter}{ext}"
                file_path = os.path.join(directory, unique_filename)
                counter += 1

            return file_path


    def extract_text_from_file(self, file_path):
        """Extract text content from a file."""
        try:
            logging.debug(f"Attempting to extract text from: {file_path}")
            file_ext = os.path.splitext(file_path)[1].lower()

            if file_ext == '.docx':
                return self.extract_text_from_word(file_path)
            elif file_ext == '.pdf':
                return self.extract_text_from_pdf(file_path)
            elif file_ext == '.doc':
                with tempfile.TemporaryDirectory() as temp_dir:
                    docx_path = self.pdf_ops._convert_doc_to_docx(file_path)
                    if docx_path:
                        return self.extract_text_from_word(docx_path)
                    else:
                        raise RuntimeError("Failed to convert .doc to .docx.")
            else:
                raise ValueError(f"Unsupported file type: {file_ext}")
        except Exception as e:
            logging.error(f"Error extracting text from file: {str(e)}", exc_info=True)
            return f"Error extracting text: {str(e)}"



    def extract_text_from_pdf(self, file_path):
        """Extract text from a PDF using PyMuPDF."""
        try:
            logging.debug(f"Opening PDF document: {file_path}")
            text_content = ""

            pdf_document = fitz.open(file_path)
            for page in pdf_document:
                text_content += page.get_text() + "\n"

            logging.debug(f"Extracted {len(text_content)} characters from PDF")

            # If no text content was extracted, fallback to OCR
            if not text_content.strip():
                logging.info("No text content extracted from PDF, attempting OCR.")
                text_content = self.pdf_ops.extract_text_from_image_pdf(file_path)
                if not text_content.strip():
                    raise RuntimeError("Failed to extract text from PDF using both standard and OCR methods.")

            return text_content
        except Exception as e:
            logging.error(f"Error extracting text from PDF: {e}", exc_info=True)
            return f"Error extracting text: {e}"



    def extract_text_from_word(self, file_path):
        """Extract text from a Word document."""
        try:
            logging.debug(f"Opening Word document: {file_path}")
            with open(file_path, 'rb') as f:
                docx_buffer = BytesIO(f.read())
            doc = Document(docx_buffer)
            paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
            return '\n'.join(paragraphs)
        except Exception as e:
            logging.error(f"Error extracting text from Word document: {e}")
            return f"Error extracting text: {e}"



    def save_file(self, source_path, save_dir, new_file_name):
        """
        Save a file to the specified directory with a unique name.
        Handles path normalization and timestamp updates.
        """
        try:
            # Normalize paths
            source_path = os.path.normpath(source_path)
            save_dir = os.path.normpath(save_dir)
            new_file_name = os.path.basename(new_file_name)  # Ensure no path in filename
            
            # Validate source file existence
            if not os.path.exists(source_path):
                raise FileNotFoundError(f"Source file does not exist: {source_path}")

            # Generate a unique file path
            save_path = self.get_unique_filename(save_dir, new_file_name)
            
            # Copy the file
            shutil.copy(source_path, save_path)
            
            # Update modification time to current time
            current_time = time.time()
            os.utime(save_path, (current_time, current_time))
            
            logging.info(f"File saved successfully: {save_path}")
            return save_path
        except Exception as e:
            logging.error(f"Error saving file: {str(e)}")
            raise



    def generate_new_filename(self, original_filename, filename_portion=None):
        """Generate a new filename with an optional portion."""
        base_name = os.path.basename(unquote(original_filename))
        if filename_portion:
            return f"{filename_portion} - {base_name}"
        return base_name


    def is_valid_drop(self, mime_data):
        return mime_data.hasUrls() or \
               mime_data.hasFormat("FileGroupDescriptor") or \
               mime_data.hasFormat("FileContents") or \
               mime_data.hasFormat("application/x-qt-windows-mime;value=\"Ole Object\"")

    def handle_mime_type(self, mime_data):
        """Process different MIME types based on their content."""
        try:
            if self.is_outlook_item(mime_data):
                logging.info("Processing Outlook item.")
                # Call appropriate Outlook logic
            elif mime_data.hasUrls():
                logging.info("Processing standard file URLs.")
                for url in mime_data.urls():
                    file_path = url.toLocalFile()
                    self.process_dropped_file(file_path)
            else:
                logging.warning("Unsupported MIME type detected.")
                raise ValueError("Unsupported MIME type.")
        except Exception as e:
            logging.error(f"Error handling MIME type: {str(e)}", exc_info=True)
            raise

    def is_outlook_item(self, mime_data):
        """Check if the dropped item is from Outlook."""
        return (
            mime_data.hasFormat("application/x-qt-windows-mime;value=\"Ole Object\"") or
            mime_data.hasFormat("application/x-qt-windows-mime;value=\"Outlook Message Format\"") or
            (
                mime_data.hasFormat("FileGroupDescriptor") and
                mime_data.hasFormat("FileContents")
            )
        )
    
    def load_recent_filename_portions(self):
        recent_portions_path = os.path.join(os.path.dirname(self.file_name_portions_path), 'recent_filename_portions.txt')
        try:
            with open(recent_portions_path, 'r') as file:
                return [line.strip() for line in file if line.strip()]
        except FileNotFoundError:
            return []

    def save_recent_filename_portions(self, portions):
        """Save recent filename portions."""
        recent_portions_path = os.path.join(os.path.dirname(self.file_name_portions_path), 'recent_filename_portions.txt')
        try:
            with open(recent_portions_path, 'w', encoding='utf-8') as file:
                for portion in portions:
                    file.write(f"{portion}\n")
            logging.info(f"Saved {len(portions)} recent portions")
        except Exception as e:
            logging.error(f"Error saving recent portions: {str(e)}", exc_info=True)
            raise
    def open_file(self, file_path):
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"The file '{file_path}' could not be found.")
            
            if sys.platform == "win32":
                os.startfile(file_path)
            else:
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, file_path])
            logging.info(f"File opened successfully: {file_path}")
        except Exception as e:
            logging.error(f"Error opening file: {str(e)}", exc_info=True)
            raise

    def determine_file_type(self, file_path):
        try:
            with open(file_path, 'rb') as f:
                header = f.read(8)
            
            if header.startswith(b'%PDF'):
                return 'pdf'
            elif header.startswith(b'\x50\x4B\x03\x04'):  # ZIP archive (could be DOCX)
                return 'docx'
            else:
                logging.warning(f"Unable to determine file type for {file_path}")
                return None
        except Exception as e:
            logging.error(f"Error determining file type: {str(e)}", exc_info=True)
            raise

    def sanitize_filename(self, filename):
        # Remove invalid characters, but keep spaces
        valid_chars = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        sanitized = ''.join(c for c in filename if c in valid_chars)
        # Replace multiple spaces with a single space
        sanitized = ' '.join(sanitized.split())
        # Trim to maximum length (e.g., 255 characters)
        sanitized = sanitized[:255]
        # Ensure the filename is not empty
        if not sanitized:
            sanitized = "unnamed_file"
        return sanitized.strip()  # Remove leading/trailing spaces

    def create_directory_if_not_exists(self, directory):
        if not os.path.exists(directory):
            os.makedirs(directory)
            logging.info(f"Created directory: {directory}")

    def move_file(self, source_path, destination_path):
        try:
            shutil.move(source_path, destination_path)
            logging.info(f"Moved file from {source_path} to {destination_path}")
        except Exception as e:
            logging.error(f"Error moving file: {str(e)}", exc_info=True)
            raise

    def delete_file(self, file_path):
        try:
            os.remove(file_path)
            logging.info(f"Deleted file: {file_path}")
        except Exception as e:
            logging.error(f"Error deleting file: {str(e)}", exc_info=True)
            raise

    def get_file_size(self, file_path):
        try:
            return os.path.getsize(file_path)
        except Exception as e:
            logging.error(f"Error getting file size: {str(e)}", exc_info=True)
            raise

    def get_file_creation_time(self, file_path):
        try:
            return os.path.getctime(file_path)
        except Exception as e:
            logging.error(f"Error getting file creation time: {str(e)}", exc_info=True)
            raise

    def get_file_modification_time(self, file_path):
        try:
            return os.path.getmtime(file_path)
        except Exception as e:
            logging.error(f"Error getting file modification time: {str(e)}", exc_info=True)
            raise

# Add any additional file operation methods as needed