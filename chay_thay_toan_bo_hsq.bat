@echo off
chcp 65001 >nul
title Thay The Toan Bo Ho So Quet Theo Mo Ta
echo Dang khoi dong tool...
"%~dp0.venv\Scripts\python.exe" "%~dp0thay_toan_bo_hsq_theo_mota.py"
pause
