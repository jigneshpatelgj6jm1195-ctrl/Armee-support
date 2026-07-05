@echo off
rem Double-click this file, type the ssgujarat.org portal password, press Enter.
rem It is stored ONLY in the local .env file next to this script.
echo.
set /p PW="Enter the ssgujarat.org portal password (user 1001): "
powershell -NoProfile -Command "(Get-Content '%~dp0.env' -Raw) -replace 'PORTAL_PASS=.*', 'PORTAL_PASS=%PW%' | Set-Content '%~dp0.env' -Encoding utf8 -NoNewline"
echo.
echo Saved to .env - you can close this window now.
pause
