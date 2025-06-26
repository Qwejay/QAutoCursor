@echo off
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --icon=icon.ico --add-data "targets;targets" main.py
pause