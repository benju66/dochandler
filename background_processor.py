# background_processor.py

from PyQt6.QtCore import QThread, pyqtSignal
import queue
import logging
import os
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Callable
from enum import Enum
from queue import PriorityQueue

class TaskType(Enum):
    FILE_CONVERSION = "file_conversion"
    TEXT_EXTRACTION = "text_extraction"
    FILE_ORGANIZATION = "file_organization"

@dataclass
class Task:
    type: TaskType
    func: Callable
    args: tuple
    kwargs: dict
    callback: Callable = None

class BackgroundProcessor(QThread):
    task_completed = pyqtSignal(TaskType, object)
    task_failed = pyqtSignal(TaskType, str)
    progress_updated = pyqtSignal(int)

    def __init__(self, max_workers=None):
        super().__init__()
        self.task_queue = PriorityQueue()
        self.executor = ThreadPoolExecutor(max_workers=max_workers or os.cpu_count())
        self.running = True
        self._processed_files = set()

    def update_progress(self, progress):
        self.progress_updated.emit(progress)

    def add_task(self, priority, task_type, func, *args, callback=None, **kwargs):
        if args and isinstance(args[0], str) and os.path.isfile(args[0]):
            file_path = args[0]
            if file_path in self._processed_files or self.is_processing_file(file_path):
                logging.info(f"File already processed or in queue: {file_path}")
                return
            self._processed_files.add(file_path)
            
        task = Task(task_type, func, args, kwargs, callback)
        self.task_queue.put((priority, task))
        logging.debug(f"Added task to queue with priority {priority}: {task_type.value}")

    def is_processing_file(self, file_path):
        with self.task_queue.mutex:
            return any(
                isinstance(item[1], Task) and 
                item[1].args and 
                isinstance(item[1].args[0], str) and 
                item[1].args[0] == file_path 
                for item in self.task_queue.queue
            )

    def run(self):
        while self.running:
            try:
                priority, task = self.task_queue.get(timeout=1)
                logging.info(f"Processing task: {task.type.value}")

                future = self.executor.submit(task.func, *task.args, **task.kwargs)
                try:
                    result = future.result()
                    if task.callback:
                        task.callback(result)
                    self.task_completed.emit(task.type, result)
                except Exception as task_error:
                    logging.error(f"Task execution failed: {str(task_error)}", exc_info=True)
                    self.task_failed.emit(task.type, str(task_error))
                    if task.args and isinstance(task.args[0], str):
                        self._processed_files.discard(task.args[0])

            except queue.Empty:
                continue
            except Exception as e:
                logging.error(f"Task processing failed: {str(e)}", exc_info=True)
                if 'task' in locals():
                    self.task_failed.emit(task.type, str(e))

    def stop(self):
        self.running = False
        self.executor.shutdown(wait=True)
        while not self.task_queue.empty():
            try:
                _, task = self.task_queue.get_nowait()
                logging.warning(f"Unprocessed task dropped: {task.type.value}")
            except:
                pass
        self._processed_files.clear()
        self.wait()

def integrate_background_processor(doc_handler_app):
    """
    Add background processing to the existing DocHandlerApp.

    :param doc_handler_app: Instance of the DocHandlerApp.
    """
    try:
        def handle_task_completed(task_type, result):
            logging.info(f"Task completed: {task_type}, Result: {result}")
            if task_type == TaskType.FILE_CONVERSION:
                doc_handler_app.ui_components.set_label_text(f"File converted: {result}")
                doc_handler_app.ui_components.update_recent_files(result)
                if hasattr(doc_handler_app, 'preview_window') and doc_handler_app.preview_window.isVisible():
                    doc_handler_app.preview_window.set_preview(result)
            elif task_type == TaskType.FILE_ORGANIZATION:
                doc_handler_app.ui_components.set_label_text(f"File processed: {result}")
                doc_handler_app.ui_components.update_recent_files(result)
                if hasattr(doc_handler_app, 'preview_window') and doc_handler_app.preview_window.isVisible():
                    doc_handler_app.preview_window.set_preview(result)

        def handle_task_failed(task_type, error_message):
            logging.error(f"Task failed: {task_type}, Error: {error_message}")
            doc_handler_app.ui_components.show_error_message("Task Failed", error_message)

        # Add methods to the class
        setattr(doc_handler_app.__class__, 'handle_task_completed', handle_task_completed)
        setattr(doc_handler_app.__class__, 'handle_task_failed', handle_task_failed)

        # Add and start the background processor
        doc_handler_app.background_processor = BackgroundProcessor()
        doc_handler_app.background_processor.start()

        # Connect signals
        doc_handler_app.background_processor.task_completed.connect(doc_handler_app.handle_task_completed)
        doc_handler_app.background_processor.task_failed.connect(doc_handler_app.handle_task_failed)
        doc_handler_app.background_processor.progress_updated.connect(doc_handler_app.ui_components.show_progress)
        logging.info("BackgroundProcessor integrated successfully.")
    except Exception as e:
        logging.error(f"Failed to integrate BackgroundProcessor: {e}")
        raise RuntimeError(f"Background processor integration failed: {e}")
