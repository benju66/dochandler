from pathlib import Path
import os

# Base directory for the application
BASE_DIR = Path(__file__).resolve().parent

# Ensure essential directories exist
def ensure_directory(path):
    Path(path).mkdir(parents=True, exist_ok=True)

# Configuration
CONFIG = {
    'FILE_NAME_PORTIONS_PATH': BASE_DIR / 'data' / 'file_name_portions.txt',
    'COMPANY_NAMES_PATH': BASE_DIR / 'data' / 'company_names.txt',
    'DEFAULT_SAVE_DIR': Path.home() / 'Downloads' / 'DocHandler',
    'LOG_FILE_PATH': Path.home() / 'DocHandlerLogs' / 'dochandler.log',
    'RECENT_SAVE_LOCATIONS': Path.home() / 'DocHandlerLogs' / 'recent_save_locations.txt',
}

# Ensure required directories exist
ensure_directory(CONFIG['DEFAULT_SAVE_DIR'].parent)
ensure_directory(CONFIG['LOG_FILE_PATH'].parent)
