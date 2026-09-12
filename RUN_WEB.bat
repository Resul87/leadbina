@echo off
chcp 65001 > nul
title BinaLeadPro Cloud Web Server
color 0A

echo ======================================================================
echo           🚀 BİNALAUNCH PRO CLOUD — WEB SAAS PLATFORMASI
echo ======================================================================
echo.
echo [1/2] Server hazirlanir...
start "" "http://127.0.0.1:8000"

echo [2/2] Uvicorn FastAPI Server basladilir...
echo.
echo ----------------------------------------------------------------------
echo Brauzer avtomatik acilacaq: http://127.0.0.1:8000
echo Serveri dayandirmaq ucun bu pencerede CTRL + C basin.
echo ----------------------------------------------------------------------
echo.

python -m uvicorn web_server:app --host 0.0.0.0 --port 8000

pause
