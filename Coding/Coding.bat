@Echo Preparing the Coding Daily Snapshot Emails

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%coding_main.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed.
pause