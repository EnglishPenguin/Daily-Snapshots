@Echo Preparing the Front End Daily Snapshot Emails

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%MCD_MCO.py
python -u "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%MCRNotPrimary.py
python -u "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%MCRPartBInactive.py
python -u "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%TPLAlert.py
python -u "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%MCR_Advantage.py
python -u "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed.
pause