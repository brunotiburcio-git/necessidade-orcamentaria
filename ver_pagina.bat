@echo off
REM Abre a pagina web (docs\index.html) no navegador, em http://localhost:8000, para conferir antes de publicar.
REM Para desligar: feche a janela ou aperte Ctrl+C.
cd /d "%~dp0docs"
start "" /b cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:8000/"
py -m http.server 8000 --bind 127.0.0.1
pause
