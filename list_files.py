import os

# Define folders and files to exclude
EXCLUDED_FOLDERS = {
    "venv", "__pycache__", ".git", ".vscode", "extlibs", "libxml", "libxslt", "libexslt", "isoschematron",
    "lxml-5.3.0.dist-info", "pip", "Scripts", "_vendor", "distutils", "site-packages", "main", "hooks", "build", "dist"
}
EXCLUDED_FILES = {
    ".DS_Store", "Thumbs.db", "pyvenv.cfg", "activate", "activate.bat", "Activate.ps1", "deactivate.bat", "pip.exe",
    "pip3.11.exe", "pip3.exe", "python.exe", "pythonw.exe", "INSTALLER", "LICENSE.txt", "METADATA", "RECORD", "WHEEL", "main.spec"
}
EXCLUDED_EXTENSIONS = {".pxd", ".h", ".pxi", ".pyc", ".pyd", ".dist-info", ".svg", ".zip", ".spev"}

def print_directory_structure(path, indent=""):
    for item in os.listdir(path):
        item_path = os.path.join(path, item)
        # Skip if the item is in excluded lists or has an excluded extension
        if (
            item in EXCLUDED_FOLDERS or
            item in EXCLUDED_FILES or
            any(item.endswith(ext) for ext in EXCLUDED_EXTENSIONS)
        ):
            continue
        print(f"{indent}- {item}")
        if os.path.isdir(item_path):
            print_directory_structure(item_path, indent + "  ")

# Set your actual root folder path
root_path = 'C:\document_handler'
print_directory_structure(root_path)
