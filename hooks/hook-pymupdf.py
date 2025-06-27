from PyInstaller.utils.hooks import collect_data_files

# Include additional data files required by pymupdf
datas = collect_data_files('pymupdf')
