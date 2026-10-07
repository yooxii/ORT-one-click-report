@echo off
rem ---------------------------------------------------------------------------
rem ORT XP lite client - run the plan deadline reminder ONCE and really send mail.
rem
rem Meant to be started by the Windows Task Scheduler (see docs/04 - deployment
rem and hand-over). Nothing is scheduled by this file itself: the client never
rem sends reminders on its own, so nothing happens until YOU create the task.
rem
rem ASCII-ONLY on purpose (cmd.exe reads batch files with the OEM code page).
rem
rem Coordination with the main program (important, see docs/04 section 7):
rem   both ends share mail_logs, so with "dedupe days" >= 1 the second end skips
rem   the same plan (kind=Warning, ref=Plan/<JobNo>). Keep that setting >= 1 if
rem   the main program may also send reminders.
rem
rem Exit codes: 0 = nothing failed, 1 = at least one mail failed (look at
rem Logs\ort_xp.log). Details of every run are appended to Logs\reminder.log.
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

if not exist "%~dp0Logs" mkdir "%~dp0Logs"
set "REMINDERLOG=%~dp0Logs\reminder.log"

echo [%DATE% %TIME%] reminder start >> "%REMINDERLOG%"
"%~dp0ORT-XP.exe" --check-plans-email --send --no-dialog
set "RC=%ERRORLEVEL%"
echo [%DATE% %TIME%] reminder done, exit=%RC% >> "%REMINDERLOG%"

exit /b %RC%
