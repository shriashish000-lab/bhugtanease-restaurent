@echo off
setlocal
title BhugtanEase - EXE Builder
echo ============================================
echo   BhugtanEase EXE Builder
echo ============================================
echo.

REM ---------------------------------------------------------------
REM  IMPORTANT (Windows 7 compatibility):
REM  Python 3.8 is the LAST official Python version that runs on
REM  Windows 7. If you need the EXE to work on Windows 7, you MUST
REM  run this script using a Python 3.8 installation.
REM  Download (if needed): https://www.python.org/downloads/release/python-3810/
REM  -> choose "Windows installer (64-bit)"
REM  If you only need Windows 8/10/11, any modern Python (3.9-3.12) is fine.
REM ---------------------------------------------------------------

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python nahi mila. Pehle Python install karo:
    echo         https://www.python.org/downloads/
    echo         Install karte waqt "Add Python to PATH" tick karna zaroori hai.
    pause
    exit /b 1
)

echo Using this Python:
python --version
echo.

echo [1/5] Virtual environment ban rahi hai ...
python -m venv venv
call venv\Scripts\activate.bat

echo [2/5] pip upgrade ho raha hai ...
python -m pip install --upgrade pip >nul

echo [3/5] Zaroori packages install ho rahe hain (openpyxl, reportlab, pywin32, pyinstaller) ...
pip install -r requirements.txt

echo [4/5] pywin32 post-install (printer support ke liye) ...
python venv\Scripts\pywin32_postinstall.py -install >nul 2>nul

echo [5/5] EXE bana rahe hain (logo ke saath, single file, no console) ...
pyinstaller --noconfirm --onefile --windowed ^
    --name "BhugtanEase" ^
    --icon "BhugtanEase.ico" ^
    --hidden-import win32print ^
    --hidden-import win32api ^
    --hidden-import openpyxl ^
    --hidden-import reportlab ^
    --collect-submodules openpyxl ^
    --collect-submodules reportlab ^
    BhugtanEase_Final.pyw

if exist "dist\BhugtanEase.exe" (
    copy /Y "dist\BhugtanEase.exe" "BhugtanEase.exe" >nul
    echo.
    echo ============================================
    echo   SUCCESS! BhugtanEase.exe taiyaar hai.
    echo   Ye file isi folder mein bhi copy ho gayi hai.
    echo   Customer ko SIRF "BhugtanEase.exe" file dena hai.
    echo ============================================
) else (
    echo.
    echo [ERROR] EXE nahi ban paayi. Upar ke error messages dekho.
    pause
    exit /b 1
)

echo.
echo Ab installer ^(Setup.exe^) banane ki koshish ho rahi hai ...
echo.
call build_installer.bat

endlocal
