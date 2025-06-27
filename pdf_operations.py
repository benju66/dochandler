from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from docx import Document
import os
import time
import pikepdf
import PyPDF2
import logging
import subprocess
from PIL import Image
import fitz  # PyMuPDF
import pytesseract
from urllib.parse import unquote
import win32com.client
import pythoncom
import tempfile
from io import BytesIO
import time
from PyQt6.QtWidgets import QApplication, QLineEdit
from PyQt6.QtGui import QCursor
from PyQt6.QtCore import Qt
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
import shutil

class PDFOperations:
    def __init__(self, file_ops, resource_manager=None):
        self.file_ops = file_ops
        self.resource_manager = resource_manager
        pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Burness\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

    @contextmanager
    def get_word_instance(self):
        """Create an isolated Word instance for conversion."""
        pythoncom.CoInitialize()
        word_app = None
        try:
            word_app = win32com.client.DispatchEx('Word.Application')
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            yield word_app
        except Exception as e:
            logging.error(f"Error creating Word instance: {str(e)}", exc_info=True)
            raise Exception("Microsoft Word is required for document conversion. Please ensure Word is installed.")
        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def extract_text_from_image_pdf(self, pdf_path):
        """Extract text from image-based PDFs using OCR."""
        if not self.resource_manager:
            raise RuntimeError("ResourceManager not initialized")

        with self.resource_manager.busy_cursor():
            text_content = ""
            try:
                pdf_document = fitz.open(pdf_path)
                for page_num in range(pdf_document.page_count):
                    page = pdf_document.load_page(page_num)
                    pix = page.get_pixmap()
                    
                    # Use in-memory buffer
                    img_buffer = BytesIO()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    
                    # Perform OCR on in-memory image
                    page_text = pytesseract.image_to_string(Image.open(img_buffer))
                    text_content += f"Page {page_num + 1}\n{page_text}\n\n"
                return text_content
            except Exception as e:
                logging.error(f"Error performing OCR on PDF: {str(e)}", exc_info=True)
                return ""

    def _convert_image_to_pdf(self, image_path, save_dir, new_file_name):
        """Convert image to PDF using PIL and in-memory operations."""
        if not self.resource_manager:
            raise RuntimeError("ResourceManager not initialized")

        with self.resource_manager.busy_cursor():
            try:
                pdf_name = os.path.splitext(new_file_name)[0] + '.pdf'
                pdf_path = self.file_ops.get_unique_filename(save_dir, pdf_name)
                
                # Use in-memory buffer
                img = Image.open(image_path)
                if img.mode in ['RGBA', 'LA', 'P']:
                    img = img.convert('RGB')
                
                pdf_buffer = BytesIO()
                img.save(pdf_buffer, 'PDF', resolution=300.0)
                pdf_buffer.seek(0)
                
                # Save to final path
                with open(pdf_path, 'wb') as f:
                    f.write(pdf_buffer.read())
                logging.info(f"Successfully converted image to PDF: {pdf_path}")
                return pdf_path
            except Exception as e:
                logging.error(f"Error converting image to PDF: {str(e)}", exc_info=True)
                raise



    def _convert_doc_to_docx(self, doc_path):
        """Convert .doc file to .docx format."""
        try:
            with self.get_word_instance() as word:
                # Get unique docx path
                docx_path = os.path.splitext(doc_path)[0] + '.docx'
                doc = word.Documents.Open(doc_path)
                try:
                    doc.SaveAs2(docx_path, FileFormat=16)  # wdFormatDocumentDefault = 16
                    return docx_path
                finally:
                    doc.Close()
        except Exception as e:
            logging.error(f"Error converting .doc to .docx: {str(e)}")
            return None
    
    def convert_word_to_pdf(self, doc_path, save_dir, base_name):
        """
        Convert a Word document to PDF and save it to the specified directory.
        Uses context manager for Word instance and handles unique filenames.
        """
        try:
            # Normalize paths
            doc_path = os.path.normpath(doc_path)
            save_dir = os.path.normpath(save_dir)
            
            if not os.path.exists(doc_path):
                raise FileNotFoundError(f"Word file not found: {doc_path}")

            # Ensure base_name doesn't include .pdf extension
            base_name = os.path.splitext(os.path.basename(base_name))[0]
            
            # Generate unique filename through FileOperations
            pdf_path = self.file_ops.get_unique_filename(save_dir, f"{base_name}.pdf")

            with self.get_word_instance() as word:
                try:
                    # Convert paths to absolute to avoid any path resolution issues
                    abs_doc_path = os.path.abspath(doc_path)
                    abs_pdf_path = os.path.abspath(pdf_path)
                    
                    doc = word.Documents.Open(abs_doc_path)
                    try:
                        doc.SaveAs(abs_pdf_path, FileFormat=17)  # 17 = wdFormatPDF
                        # Update modification time
                        current_time = time.time()
                        os.utime(abs_pdf_path, (current_time, current_time))
                        logging.info(f"Converted Word to PDF: {pdf_path}")
                        return pdf_path
                    finally:
                        doc.Close(SaveChanges=False)
                except Exception as word_error:
                    logging.error(f"Error in Word conversion process: {str(word_error)}")
                    raise RuntimeError(f"Word conversion failed: {str(word_error)}")
        except Exception as e:
            logging.error(f"Error converting Word to PDF: {str(e)}")
            raise RuntimeError(f"Failed to convert {doc_path} to PDF: {str(e)}")

    def convert_to_pdf(self, file_path, save_dir, new_file_name):
        try:
            # Validate input file
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Input file not found: {file_path}")
            if os.path.getsize(file_path) == 0:
                raise ValueError(f"Input file is empty: {file_path}")

            # Ensure new_file_name has .pdf extension
            if not new_file_name.lower().endswith('.pdf'):
                new_file_name = f"{os.path.splitext(new_file_name)[0]}.pdf"

            # Get unique filename using the existing method
            pdf_path = self.file_ops.get_unique_filename(save_dir, new_file_name)

            # Handle conversion based on file type
            file_ext = os.path.splitext(file_path)[1].lower()
            if file_ext in ['.doc', '.docx']:
                return self.convert_word_to_pdf(file_path, save_dir, os.path.splitext(new_file_name)[0])
            elif file_ext == '.pdf':
                # Copy PDF and update modification time
                shutil.copy(file_path, pdf_path)
                current_time = time.time()
                os.utime(pdf_path, (current_time, current_time))
                logging.info(f"PDF file copied to: {pdf_path} with updated timestamp")
                return pdf_path
            else:
                raise ValueError(f"Unsupported file type: {file_ext}")

        except Exception as e:
            logging.error(f"Error converting file to PDF: {e}", exc_info=True)
            raise RuntimeError(f"Failed to convert {file_path} to PDF: {str(e)}")




    @contextmanager
    def get_excel_instance(self):
        """Create an isolated Excel instance for conversion."""
        pythoncom.CoInitialize()
        excel_app = None
        try:
            excel_app = win32com.client.DispatchEx('Excel.Application')
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            yield excel_app
        except Exception as e:
            logging.error(f"Error creating Excel instance: {str(e)}", exc_info=True)
            raise Exception("Microsoft Excel is required for conversion. Please ensure Excel is installed.")
        finally:
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def _convert_excel_to_pdf(self, excel_path, save_dir, new_file_name):
        """Convert Excel to PDF using COM automation."""
        if not self.resource_manager:
            raise RuntimeError("ResourceManager not initialized")
                
        try:
            pdf_name = os.path.splitext(new_file_name)[0] + '.pdf'
            pdf_path = self.file_ops.get_unique_filename(save_dir, pdf_name)
            
            # Use get_excel_instance for proper COM initialization
            with self.get_excel_instance() as excel:
                try:
                    # Create absolute paths
                    abs_excel_path = os.path.abspath(excel_path)
                    abs_pdf_path = os.path.abspath(pdf_path)
                    
                    # Open workbook and export to PDF
                    wb = excel.Workbooks.Open(abs_excel_path)
                    try:
                        wb.ExportAsFixedFormat(0, abs_pdf_path)
                        logging.info(f"Successfully converted Excel to PDF: {pdf_path}")
                        return pdf_path
                    finally:
                        wb.Close(False)
                except Exception as e:
                    raise Exception(f"Failed to save PDF: {str(e)}")
                    
        except Exception as e:
            logging.error(f"Error converting Excel to PDF: {str(e)}", exc_info=True)
            raise


    def merge_pdfs(self, file_paths, output_path):
        """Merge multiple PDFs into a single PDF."""
        try:
            merger = PdfMerger()

            for file_path in file_paths:
                if os.path.exists(file_path):
                    merger.append(file_path)
                else:
                    logging.warning(f"File not found and skipped: {file_path}")

            # Save the merged PDF
            merger.write(output_path)
            merger.close()

            logging.info(f"Merged PDF saved to: {output_path}")
            return output_path
        except Exception as e:
            logging.error(f"Error merging PDFs: {str(e)}", exc_info=True)
            raise


    def extract_pages(self, input_pdf, output_pdf, pages):
        """Extract specific pages from a PDF."""
        try:
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            for page_num in pages:
                if 0 < page_num <= len(reader.pages):
                    writer.add_page(reader.pages[page_num - 1])

            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
            logging.info(f"Extracted pages saved to: {output_pdf}")
            return output_pdf
        except Exception as e:
            logging.error(f"Error extracting PDF pages: {str(e)}", exc_info=True)
            raise

    def split_pdf(self, input_pdf, output_directory):
        """Split a PDF into individual pages."""
        try:
            reader = PdfReader(input_pdf)
            self.file_ops.create_directory_if_not_exists(output_directory)

            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                output_filename = f'page_{i + 1}.pdf'
                output_path = os.path.join(output_directory, output_filename)
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)

            logging.info(f"PDF split into individual pages in directory: {output_directory}")
            return output_directory
        except Exception as e:
            logging.error(f"Error splitting PDF: {str(e)}", exc_info=True)
            raise

    def rotate_pdf(self, input_pdf, output_pdf, rotation):
        """Rotate all pages of a PDF by a specified angle."""
        try:
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            for page in reader.pages:
                page.rotate(rotation)
                writer.add_page(page)

            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
            logging.info(f"Rotated PDF saved to: {output_pdf}")
            return output_pdf
        except Exception as e:
            logging.error(f"Error rotating PDF: {str(e)}", exc_info=True)
            raise

    def compress_pdf(self, input_pdf, output_pdf, power=0):
        """Compress a PDF using Ghostscript."""
        try:
            if not self._is_ghostscript_installed():
                raise EnvironmentError("Ghostscript is not installed or not in PATH")

            quality = {
                0: '/default',
                1: '/prepress',
                2: '/printer',
                3: '/ebook',
                4: '/screen'
            }

            subprocess.call([
                'gs', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                f'-dPDFSETTINGS={quality[power]}', '-dNOPAUSE', '-dQUIET',
                '-dBATCH', f'-sOutputFile={output_pdf}', input_pdf
            ])

            logging.info(f"Compressed PDF saved to: {output_pdf}")
            return output_pdf  # Fixed: returning output_pdf instead of output_path
        except Exception as e:
            logging.error(f"Error compressing PDF: {str(e)}", exc_info=True)
            raise

    def _is_ghostscript_installed(self):
        """Check if Ghostscript is installed."""
        try:
            subprocess.call(['gs', '--version'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except FileNotFoundError:
            return False

    def add_watermark(self, input_pdf, watermark_pdf, output_pdf):
        """Add a watermark to a PDF."""
        try:
            watermark = PdfReader(watermark_pdf).pages[0]
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            for page in reader.pages:
                page.merge_page(watermark)
                writer.add_page(page)

            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)

            logging.info(f"PDF with watermark saved to: {output_pdf}")
            return output_pdf
        except Exception as e:
            logging.error(f"Error adding watermark to PDF: {str(e)}", exc_info=True)
            raise

    def extract_text_from_pdf(self, pdf_path):
        """Extract text from a PDF using multi-threading."""
        try:
            pdf_document = fitz.open(pdf_path)
            text_content = [""] * pdf_document.page_count  # Pre-allocate for thread-safe storage

            def extract_page_text(page_num):
                try:
                    page = pdf_document.load_page(page_num)
                    return page.get_text()
                except Exception as e:
                    logging.error(f"Error extracting text from page {page_num}: {e}")
                    return ""

            with ThreadPoolExecutor(max_workers=4) as executor:  # Adjust max_workers as needed
                results = executor.map(extract_page_text, range(pdf_document.page_count))

            # Gather results
            for idx, text in enumerate(results):
                text_content[idx] = text

            pdf_document.close()
            return "\n".join(text_content)  # Combine all page texts
        except Exception as e:
            logging.error(f"Error extracting text from PDF: {e}", exc_info=True)
            return f"Error extracting text: {e}"

    def decrypt_pdf(self, input_path, output_path=None):
        """Remove password protection from PDF without requiring password."""
        try:
            output_path = output_path or input_path
            pdf = pikepdf.open(input_path)
            pdf.save(output_path)
            pdf.close()
            logging.info(f"Successfully decrypted PDF: {output_path}")
            return output_path
        except Exception as e:
            logging.error(f"Failed to decrypt PDF: {e}")
            return False

    def process_pdf(self, file_path, save_dir, new_filename):
        """Process PDF file, removing all encryption."""
        try:
            output_path = os.path.join(save_dir, new_filename)
            
            # Try common passwords
            passwords = ['', 'password', 'admin', '1234', '12345', 'test']
            
            for password in passwords:
                try:
                    with pikepdf.open(file_path, password=password) as pdf:
                        # Force remove all encryption and restrictions
                        pdf.save(output_path, encryption=False)
                        return output_path
                except pikepdf.PasswordError:
                    continue
                
            # If we get here, try one last time with no security
            with pikepdf.Pdf.new() as new_pdf:
                with pikepdf.open(file_path) as old_pdf:
                    new_pdf.pages.extend(old_pdf.pages)
                new_pdf.save(output_path, encryption=False)
                
            return output_path
                
        except Exception as e:
            logging.error(f"Error processing PDF: {str(e)}")
            raise ValueError("Could not decrypt PDF - encryption too strong")
