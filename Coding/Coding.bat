@Echo Preparing the Coding Daily Snapshot Emails

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%CM1235.py
python -u "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%CM1249.py
python -u "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%CM6146.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed.
pause