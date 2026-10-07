@echo off
cd /d "%~dp0"
python bao_cao_thong_tin_thieu_mplis.py
if errorlevel 1 pause
