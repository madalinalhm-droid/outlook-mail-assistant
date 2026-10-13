@echo off
title Outlook Mail Assistant
cd /d "%~dp0"

if not exist "OutlookMailAssistant.ps1" (
    echo OutlookMailAssistant.ps1 is missing from:
    echo %cd%
    echo.
    echo Put both files in the same folder and start again.
    echo.
    pause
    exit /b 1
)

:menu
cls
echo ==========================================================
echo   OUTLOOK MAIL ASSISTANT - all mailboxes in this profile
echo ==========================================================
echo.
echo   1. Report of one day's e-mails (read only)
echo   2. Print attachments (asks for each one)
echo   3. Reply to e-mails (draft or send)
echo   4. Open the reports folder
echo   5. Demo report (fictitious e-mails, no Outlook needed)
echo.
echo   0. Exit
echo.
set "option="
set /p "option=Choose an option and press Enter: "

if "%option%"=="1" goto report
if "%option%"=="2" goto print
if "%option%"=="3" goto reply
if "%option%"=="4" goto folder
if "%option%"=="5" goto demo
if "%option%"=="0" exit /b 0
goto menu

:report
cls
echo Day: Enter = today, y = yesterday, or a date like 05.10.2026
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0OutlookMailAssistant.ps1" report
echo.
pause
goto menu

:print
cls
echo For each attachment:
echo   P = print   N = don't print   V = open it to view
echo   S = skip the whole e-mail     Q = stop printing
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0OutlookMailAssistant.ps1" print
echo.
pause
goto menu

:reply
cls
echo Pick the e-mail number, then the reply. At the end:
echo   D = save as draft   S = send now   C = cancel
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0OutlookMailAssistant.ps1" reply
echo.
pause
goto menu

:folder
if not exist "Reports" mkdir "Reports"
start "" "Reports"
goto menu

:demo
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0OutlookMailAssistant.ps1" demo
echo.
pause
goto menu
