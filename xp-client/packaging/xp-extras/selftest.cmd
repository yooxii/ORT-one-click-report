@echo off
rem ---------------------------------------------------------------------------
rem ORT XP lite client - self test (double-click me on the XP machine).
rem
rem ASCII-ONLY on purpose: cmd.exe reads batch files with the OEM code page, and
rem Chinese text here would show up garbled on a cp950 machine. No PowerShell
rem either - Windows XP does not ship it; WSH/.vbs is used for the shortcut instead.
rem
rem Runs the packaged client's own checks, then opens the log in Notepad. The log
rem is written as UTF-8 with a BOM, so Notepad shows the Chinese text correctly
rem even when the console code page cannot.
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

echo === ORT XP client self test ===
echo.
echo Note: on a console whose code page cannot show Simplified Chinese the text
echo below may look escaped (for example \u901a\u8fc7). The log file opened at the
echo end is always readable.
echo.

echo [1/4] version
"%~dp0ORT-XP.exe" --version
echo.

echo [2/4] environment self test (interpreter, sqlite, TLS, DPAPI, database)
"%~dp0ORT-XP.exe" --selftest
echo.

echo [3/4] UI assembly self test (builds every window, checks it becomes visible)
"%~dp0ORT-XP.exe" --ui-smoke
echo.

echo [4/4] plan deadline reminder (dry run - does NOT send mail)
"%~dp0ORT-XP.exe" --check-plans-email
echo.

echo === done ===
if exist "%~dp0Logs\fatal_*.log" (
    echo WARNING: a fatal error log exists - opening it.
    start notepad "%~dp0Logs"
) else (
    echo Opening the run log: %~dp0Logs\ort_xp.log
    start notepad "%~dp0Logs\ort_xp.log"
)
echo Press any key to close this window...
pause >nul
