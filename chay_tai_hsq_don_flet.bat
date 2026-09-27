@echo off
cd /d "%~dp0"
venv\Scripts\python.exe tai_hsq_don_flet.py
if errorlevel 1 pause
