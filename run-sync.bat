@echo off
title Network Asset Register - Watcher Sync
echo -----------------------------------------
echo   Initializing Network Asset Sync...
echo -----------------------------------------
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync.ps1"

echo.
echo -----------------------------------------
echo   Sync complete. You may close this window.
echo -----------------------------------------
pause