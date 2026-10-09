@echo off
REM Inicia o sistema NecOrc. Deixe esta janela aberta enquanto as pessoas estiverem usando.
REM Para desligar: feche a janela ou aperte Ctrl+C.
cd /d "%~dp0"
echo.
echo Endereco para as outras pessoas (mesma rede): http://%COMPUTERNAME%:8501
echo.
.venv\Scripts\python -m streamlit run app.py --server.port 8501
pause
