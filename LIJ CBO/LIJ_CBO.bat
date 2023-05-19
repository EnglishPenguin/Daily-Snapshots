@Echo Preparing the LIJ CBO Daily Snapshot Email

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%Itemized_Statement.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed.
pause