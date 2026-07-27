@echo off
title Build SMW ROM Hack Tracker Community Edition
cd /d "%~dp0"
echo.
echo Building SMW_ROM_Hack_Tracker_Community.xlsm from the uploaded workbook template...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build_Community_Tracker.ps1"
echo.
pause
