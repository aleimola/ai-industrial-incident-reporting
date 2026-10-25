@echo off
echo ========================================================
echo   Avvio SicuraFlow - Server Locale
echo ========================================================
echo.
echo Controllo e installazione delle dipendenze necessarie in corso...
pip install -r requirements.txt

echo.
echo Avvio del server in corso...
echo Una volta avviato, tieni aperta questa finestra nera e
echo apri il browser all'indirizzo http://localhost:8000
echo.

python server.py
pause
