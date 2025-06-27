from PyQt6.QtCore import QThread, pyqtSignal

class WordToPDFWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, pdf_operations, doc_path, save_dir, base_name):
        super().__init__()
        self.pdf_operations = pdf_operations
        self.doc_path = doc_path
        self.save_dir = save_dir
        self.base_name = base_name

    def run(self):
        """Run the Word-to-PDF conversion in a background thread."""
        try:
            self.progress.emit(0)  # Notify the UI (optional)
            pdf_path = self.pdf_operations.convert_word_to_pdf(self.doc_path, self.save_dir, self.base_name)
            self.progress.emit(100)
            self.finished.emit(pdf_path)
        except Exception as e:
            self.error.emit(str(e))
