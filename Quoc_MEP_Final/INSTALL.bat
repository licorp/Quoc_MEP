@echo off
echo ================================================
echo   Quoc MEP Universal Installer
echo ================================================
echo.

set "ADDIN_PATH=%APPDATA%\Autodesk\Revit\Addins"

echo Chon phien ban Revit cua ban:
echo 1. Revit 2020
echo 2. Revit 2021
echo 3. Revit 2022
echo 4. Revit 2023
echo 5. Revit 2024
echo 6. Revit 2025
echo 7. Revit 2026
echo.

set /p choice="Nhap so (1-7): "

if "%choice%"=="1" set "VERSION=2020"
if "%choice%"=="2" set "VERSION=2021"
if "%choice%"=="3" set "VERSION=2022"
if "%choice%"=="4" set "VERSION=2023"
if "%choice%"=="5" set "VERSION=2024"
if "%choice%"=="6" set "VERSION=2025"
if "%choice%"=="7" set "VERSION=2026"

if not defined VERSION (
    echo Lua chon khong hop le!
    pause
    exit /b 1
)

set "DEST=%ADDIN_PATH%\%VERSION%"

echo.
echo Dang cai dat vao: %DEST%
echo.

if not exist "%DEST%" mkdir "%DEST%"

xcopy /Y /E /I "%~dp0*" "%DEST%\"

echo.
echo ================================================
echo   Cai dat thanh cong!
echo ================================================
echo.
echo Khoi dong lai Revit %VERSION% de su dung.
echo.
pause
