@echo off
chcp 65001 >nul
title 库存自动化控制面板

cd /d "%~dp0"
call venv\Scripts\activate.bat
python control_panel.py
pause
