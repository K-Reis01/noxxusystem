@echo off
cd /d "%~dp0"
start "Noxxus Fichas Diarias" cmd /k "node local-fichas-service.mjs"
start https://noxxusystem.vercel.app/fichas-diarias.html
