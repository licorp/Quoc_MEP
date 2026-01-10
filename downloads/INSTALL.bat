@echo off
setlocal enabledelayedexpansion
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
echo 8. Tat ca cac phien ban (Full)
echo.

set /p choice="Nhap so (1-8): "

if "%choice%"=="1" set "VERSION=2020"
if "%choice%"=="2" set "VERSION=2021"
if "%choice%"=="3" set "VERSION=2022"
if "%choice%"=="4" set "VERSION=2023"
if "%choice%"=="5" set "VERSION=2024"
if "%choice%"=="6" set "VERSION=2025"
if "%choice%"=="7" set "VERSION=2026"
if "%choice%"=="8" set "VERSION=FULL"

if not defined VERSION (
    echo Lua chon khong hop le!
    pause
    exit /b 1
)

if "%VERSION%"=="FULL" (
    echo.
    echo Dang cai dat cho tat ca cac phien ban...
    echo.
    
    REM Cai dat cho tung version
    for %%V in (2020 2021 2022 2023 2024 2025 2026) do (
        set "DEST=%ADDIN_PATH%\%%V"
        echo Cai dat vao: !DEST!
        
        if not exist "!DEST!" mkdir "!DEST!"
        
        REM Copy file .addin va Loader DLL
        copy /Y "%~dp0Quoc_MEP_Universal.addin" "!DEST!\"
        copy /Y "%~dp0Quoc_MEP_Loader.dll" "!DEST!\"
        
        REM Copy thu muc version tuong ung
        if exist "%~dp0Revit%%V" (
            xcopy /Y /E /I "%~dp0Revit%%V" "!DEST!\Revit%%V\"
        )
    )
    
    echo.
    echo ================================================
    echo   Cai dat thanh cong cho tat ca cac phien ban!
    echo ================================================
) else (
    set "DEST=%ADDIN_PATH%\%VERSION%"
    
    echo.
    echo Dang cai dat vao: !DEST!
    echo.
    
    if not exist "!DEST!" mkdir "!DEST!"
    
    REM Copy file .addin va Loader DLL
    echo Copy Quoc_MEP_Universal.addin...
    copy /Y "%~dp0Quoc_MEP_Universal.addin" "!DEST!\" || echo ERROR: Copy addin failed
    
    echo Copy Quoc_MEP_Loader.dll...
    copy /Y "%~dp0Quoc_MEP_Loader.dll" "!DEST!\" || echo ERROR: Copy loader failed
    
    REM Copy chi thu muc version duoc chon
    if exist "%~dp0Revit%VERSION%" (
        echo Copy Revit%VERSION% folder...
        xcopy /Y /E /I "%~dp0Revit%VERSION%" "!DEST!\Revit%VERSION%\" || echo ERROR: xcopy failed
        echo.
        echo Da copy thanh cong Revit%VERSION% folder
    ) else (
        echo.
        echo CANH BAO: Khong tim thay thu muc Revit%VERSION%
        echo Thu muc hien tai: %~dp0
        dir "%~dp0" | findstr /I "Revit"
    )
    
    echo.
    echo ================================================
    echo   Cai dat thanh cong!
    echo ================================================
    echo.
    echo Khoi dong lai Revit %VERSION% de su dung.
    echo.
    echo Kiem tra thu muc: !DEST!
)

echo.
pause
