@echo off
REM ============================================================
REM Script batch per avviare l'analisi mensile dei flussi di cassa
REM Doppio click su questo file per eseguire l'analisi completa
REM ============================================================

REM Imposta la directory corrente alla posizione dello script
cd /d "%~dp0"

echo.
echo ============================================================
echo   ANALISI FLUSSI DI CASSA - Avvio automatico
echo ============================================================
echo.

echo Attivazione ambiente Python...
call ".venv\Scripts\activate.bat"

echo.
echo Avvio analisi...
echo.

".venv\Scripts\python.exe" analisi_mensile.py

echo.
echo ============================================================
echo   Premi un tasto per chiudere...
echo ============================================================
pause > nul
