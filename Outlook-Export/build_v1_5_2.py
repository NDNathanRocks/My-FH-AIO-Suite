"""Build dist\\InboxEmailExporter.exe with PyInstaller (run on Windows: py build_v1_5_2.py)."""
from pathlib import Path

import PyInstaller.__main__

p = Path(__file__).resolve().parent
PyInstaller.__main__.run([
    str(p / "inbox_email_exporter_v1_5_2.py"), "--name=InboxEmailExporter", "--onefile", "--windowed", "--clean", "--noconfirm",
    f"--paths={p}",
    "--hidden-import=win32com.client", "--hidden-import=pythoncom", "--hidden-import=pywintypes", "--hidden-import=openpyxl",
    "--hidden-import=exporter_core.outlook_com",
])
print(r"Built: dist\InboxEmailExporter.exe")
