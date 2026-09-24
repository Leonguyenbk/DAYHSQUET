@echo off
cd /d "%~dp0"
python cap_nhat_don_hsq_flet.py
if errorlevel 1 pause
