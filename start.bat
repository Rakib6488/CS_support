@echo off
setlocal
cd /d "%~dp0"

netstat -ano -p tcp 2>nul | findstr /R /C:"127.0.0.1:3001 .*LISTENING" >nul
if errorlevel 1 (
    start "Alimail" /D "%~dp0alimail" cmd /k npm run dev
)

netstat -ano -p tcp 2>nul | findstr /R /C:"127.0.0.1:3002 .*LISTENING" >nul
if errorlevel 1 (
    start "WhatsApp Bridge" /D "%~dp0whatsapp-bridge" cmd /k npm start
)

echo.
echo Alimail:         http://localhost:3001
echo WhatsApp Bridge: http://localhost:3002
echo.
