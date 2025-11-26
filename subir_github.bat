@echo off
chcp 65001 > nul
echo ========================================
echo    SUBIR CAMBIOS A GITHUB
echo ========================================
echo.

REM Pedir mensaje de commit
set /p mensaje="Introduce el mensaje del commit: "

REM Validar que no esté vacío
if "%mensaje%"=="" (
    echo [ERROR] El mensaje no puede estar vacío
    pause
    exit /b 1
)

echo.
echo [1/3] Agregando archivos...
git add .

echo [2/3] Creando commit...
git commit -m "%mensaje%"

echo [3/3] Subiendo a GitHub...
git push origin main

echo.
echo ========================================
echo    ✓ CAMBIOS SUBIDOS EXITOSAMENTE
echo ========================================
pause
