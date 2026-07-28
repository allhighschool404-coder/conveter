@echo off
title File Converter Server Launcher
setlocal enabledelayedexpansion

echo ==================================================
echo [1/3] Checking Required Python Libraries...
echo ==================================================
pip install flask pandas openpyxl python-docx PyMuPDF

echo.
echo ==================================================
echo [2/3] Starting Flask Server in Background...
echo ==================================================
start /b python App.py

ping -n 4 127.0.0.1 >nul

echo.
echo ==================================================
echo [3/3] Opening Web App in Browser...
echo ==================================================
start http://127.0.0.1:5000/

echo.
echo ====================================================================
echo                 SERVER STARTED SUCCESSFULLY!
echo ====================================================================
echo.
echo [1] THIS COMPUTER LINK (Localhost):
echo     http://127.0.0.1:5000/
echo.
echo [2] OTHER COMPUTER / MOBILE / LAN NETWORK LINK:
echo     (Connect the other device to the same Wi-Fi / Local Network)
echo     Open the link below in the other computer's browser:
echo.
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set "ip=%%a"
    set "ip=!ip: =!"
    echo     URL -^> http://!ip!:5000/
)
echo.
echo ====================================================================
echo  NOTICE: Do NOT close this window. Closing it will stop the server.
echo ====================================================================
pause