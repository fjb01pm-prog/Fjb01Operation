@echo off
setlocal enabledelayedexpansion
title FJB Operations Control - Auto Deploy System
color 0B

echo ============================================================
echo   FJB OPERATIONS CONTROL SYSTEM - AUTO DEPLOY AUTOMATION
echo ============================================================
echo.

node "%~dp0scripts\auto-deploy.js"

echo.
pause
exit /b %errorlevel%
