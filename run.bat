@echo off
title Word to PDF Converter
echo ========================================
echo    Word to PDF Converter - Running
echo ========================================
echo.

:: فعالسازی محیط مجاز و اجرای برنامه
call venv\Scripts\activate.bat
python run.py

:: اگر برنامه بسته شد
echo.
echo Program closed.
pause