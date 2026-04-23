@echo off
cd /d "%~dp0"
echo.
echo  Starting SF Case Manager...
echo  Open: http://localhost:3737
echo.
"C:\Program Files\sf\client\bin\node.exe" sf_report_server.js
pause
