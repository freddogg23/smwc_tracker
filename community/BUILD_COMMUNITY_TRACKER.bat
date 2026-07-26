@echo off
setlocal
title Build SMW ROM Hack Tracker - Community Edition
cd /d "%~dp0"
echo.
echo Building the macro-enabled Community Edition...
echo Close Excel before continuing.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build_Community_Tracker.ps1"
echo.
pause
