@echo off
echo ========================================
echo   ROBOT FIRMA JUSTIFICACIONES
echo   Iniciando servidor...
echo ========================================
echo.

cd /d "%~dp0just-signer"

REM Iniciar el servidor Python en segundo plano
start /B python app.py

REM Esperar 3 segundos a que el servidor inicie
echo Esperando a que el servidor inicie...
timeout /t 3 /nobreak >nul

REM Abrir Chrome con la interfaz del robot
echo Abriendo Chrome...
start chrome "http://localhost:8771"

echo.
echo ========================================
echo   Robot iniciado en http://localhost:8771
echo   Presiona cualquier tecla para detener
echo ========================================
pause >nul

REM Al presionar una tecla, matar el proceso Python
taskkill /F /IM python.exe /FI "WINDOWTITLE eq *app.py*" 2>nul
