@echo off
setlocal
set "documentsDir=%userprofile%\Documents"

set "targetDir=%documentsDir%\UnrealEngine\Python"

if not exist "%targetDir%" (
    mkdir "%targetDir%"
)

copy "credentials.json" "%targetDir%"
copy "VMGoogleSheetShotTracker.py" "%targetDir%"


cd /d ##UnrealPath##\UnrealEngine\5.2.1\Windows\Engine\Binaries\ThirdParty\Python3\Win64
python.exe -m pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib


echo VMShotTracker installed to %DESTINATION_FOLDER%
pause
