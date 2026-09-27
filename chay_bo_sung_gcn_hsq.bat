@echo off
cd /d "%~dp0"
venv\Scripts\python.exe bo_sung_gcn_hsq_flet.py
if errorlevel 1 pause
