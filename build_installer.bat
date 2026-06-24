@echo off
setlocal
echo ============================================
echo   BhugtanEase Installer Builder (Inno Setup)
echo ============================================
echo.

if not exist "BhugtanEase.exe" (
    echo [ERROR] "BhugtanEase.exe" nahi mili.
    echo         Pehle "build.bat" chalao jisse EXE bane.
    pause
    exit /b 1
)

REM ---- ISCC.exe (Inno Setup Compiler) dhoondo ----
set "ISCC="
where ISCC.exe >nul 2>nul
if not errorlevel 1 (
    set "ISCC=ISCC.exe"
) else (
    if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
    if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
)

if "%ISCC%"=="" (
    echo [ERROR] Inno Setup nahi mila.
    echo.
    echo Pehle Inno Setup install karo ^(free hai^):
    echo     https://jrsoftware.org/isdl.php
    echo Installer chalakar default settings se install kar do, phir
    echo yeh script ^(build_installer.bat^) dobara chalao.
    pause
    exit /b 1
)

echo Inno Setup mil gaya: %ISCC%
echo.
echo Installer ban raha hai ...
"%ISCC%" "installer.iss"

if exist "installer_output\BhugtanEase_Setup.exe" (
    echo.
    echo ============================================
    echo   SUCCESS! "BhugtanEase_Setup.exe" taiyaar hai.
    echo   Folder: installer_output\BhugtanEase_Setup.exe
    echo   Customer ko SIRF yeh ek file dena hai.
    echo   Double-click karega -> install wizard chalega ->
    echo   Desktop/Start Menu shortcut ban jayega.
    echo ============================================
) else (
    echo.
    echo [ERROR] Installer nahi ban paaya. Upar ke messages dekho.
)

echo.
pause
endlocal
