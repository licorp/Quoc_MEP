@echo off
echo ================================================
echo   Unblock DLL Files - Giai phong cac file DLL
echo ================================================
echo.
echo Dang unblock cac file DLL da download...
echo.

REM Unblock file addin va loader
powershell -Command "Get-ChildItem -Path '%~dp0' -Filter '*.dll' -Recurse | Unblock-File"
powershell -Command "Get-ChildItem -Path '%~dp0' -Filter '*.addin' -Recurse | Unblock-File"

echo.
echo ================================================
echo   Da unblock tat ca cac file!
echo ================================================
echo.
echo Bay gio ban co the chay INSTALL.bat de cai dat.
echo.
pause
