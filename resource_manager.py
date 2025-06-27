from contextlib import contextmanager
from typing import Optional, Any, Generator, TypeVar, Iterator
import win32com.client
import pythoncom
import tempfile
import os
import logging
import shutil
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QCursor
from PyQt6.QtCore import Qt, QTimer
import gc
import time
import weakref

T = TypeVar('T')

class RetryableError(Exception):
    pass

class ResourceManager:
    def __init__(self):
        self._temp_directories: set[str] = set()
        self._temp_files: set[str] = set()
        self._com_objects = weakref.WeakSet()
        self._cleanup_scheduled = False
        self._max_resource_age = 3600  # 1 hour
        self._resource_timestamps = {}
        self._max_retries = 3
        self._retry_delay = 1  # seconds

    @contextmanager
    def manage_com_object(self, com_class: str, visible: bool = False):
        obj = None
        with self._temp_files_lock:
            for attempt in range(self._max_retries):
                try:
                    pythoncom.CoInitialize()
                    obj = win32com.client.DispatchEx(com_class)
                    if hasattr(obj, 'Visible'):
                        obj.Visible = visible
                    if hasattr(obj, 'DisplayAlerts'):
                        obj.DisplayAlerts = False
                    self._com_objects.add(obj)
                    self._resource_timestamps[id(obj)] = time.time()
                    yield obj
                    break
                except Exception as e:
                    logging.warning(f"Attempt {attempt + 1} failed for COM class {com_class}: {e}")
                    if attempt == self._max_retries - 1:
                        raise
                    time.sleep(self._retry_delay * (2 ** attempt))
                finally:
                    if obj:
                        try:
                            if hasattr(obj, 'Quit'):
                                obj.Quit()
                        except Exception as quit_error:
                            logging.warning(f"Failed to quit COM object {com_class}: {quit_error}")
                        self._com_objects.discard(obj)
                    pythoncom.CoUninitialize()


    @contextmanager
    def temp_directory(self) -> Iterator[str]:
        """Create and manage a temporary directory."""
        temp_dir = tempfile.mkdtemp()
        self._temp_directories.add(temp_dir)
        try:
            yield temp_dir
        finally:
            self.cleanup_directory(temp_dir)

    @contextmanager
    def temp_file(self, suffix: Optional[str] = None) -> Iterator[str]:
        """Create and manage a temporary file."""
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        temp_file.close()
        self._temp_files.add(temp_file.name)
        try:
            yield temp_file.name
        finally:
            self.cleanup_file(temp_file.name)

    @contextmanager
    def busy_cursor(self) -> Iterator[None]:
        """Show busy cursor while operation is in progress."""
        app = QApplication.instance()
        if app:
            try:
                # Set arrow cursor first, then wait cursor
                app.setOverrideCursor(QCursor(Qt.CursorShape.ArrowCursor))
                app.setOverrideCursor(QCursor(Qt.CursorShape.WaitCursor))
                app.processEvents()
                yield
            finally:
                # Restore cursors (twice because we set two)
                app.restoreOverrideCursor()
                app.restoreOverrideCursor()
                app.processEvents()
        else:
            yield

    def cleanup_directory(self, directory: str) -> None:
        """Safely remove a temporary directory and its contents."""
        if directory in self._temp_directories:
            try:
                shutil.rmtree(directory, ignore_errors=True)
            except Exception as e:
                logging.warning(f"Failed to remove temporary directory {directory}: {e}")
            finally:
                self._temp_directories.discard(directory)

    def cleanup_file(self, file_path: str) -> None:
        """Safely remove a temporary file."""
        if file_path in self._temp_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception as e:
                logging.warning(f"Failed to remove temporary file {file_path}: {e}")
            finally:
                self._temp_files.discard(file_path)

    def schedule_cleanup(self):
        if not self._cleanup_scheduled:
            QTimer.singleShot(10000, self.cleanup_all)
            self._cleanup_scheduled = True

    def cleanup_all(self):
        current_time = time.time()
        for directory in list(self._temp_directories):
            if os.path.exists(directory) and current_time - os.path.getmtime(directory) > self._max_resource_age:
                self.cleanup_directory(directory)
        for file_path in list(self._temp_files):
            if os.path.exists(file_path) and current_time - os.path.getmtime(file_path) > self._max_resource_age:
                self.cleanup_file(file_path)
        self._cleanup_scheduled = False

    def _cleanup_old_resources(self, current_time: float) -> None:
        for resource_id, timestamp in list(self._resource_timestamps.items()):
            if current_time - timestamp > self._max_resource_age:
                for obj in self._com_objects:
                    if id(obj) == resource_id:
                        try:
                            if hasattr(obj, 'Quit'):
                                obj.Quit()
                        except:
                            pass
                        self._com_objects.discard(obj)
                        break
                self._resource_timestamps.pop(resource_id)