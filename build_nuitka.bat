@echo off
cd /d D:\JEB
.venv\Scripts\python -m nuitka --standalone --windows-disable-console --enable-plugin=pyside6 --windows-icon-from-ico=icon.ico --include-data-files=icon.ico=icon.ico --output-dir=nuitka_dist --output-filename=JsonEditor main.py
pause
