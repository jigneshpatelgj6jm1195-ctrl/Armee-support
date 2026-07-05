@echo off
rem Runs one scrape cycle and appends everything to logs\scraper.log
cd /d "%~dp0"
if not exist logs mkdir logs
echo ================ %date% %time% ================>> logs\scraper.log
"C:\Program Files\nodejs\node.exe" run.js >> logs\scraper.log 2>&1
