@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo.
echo ============================================================
echo   Regenerando Dashboard Siegfried
echo ============================================================
echo.

py build_dashboard.py
set ERR=%ERRORLEVEL%

echo.
if %ERR% NEQ 0 (
    echo ============================================================
    echo   ERROR: no se pudo generar el dashboard.
    echo   - Asegurate de cerrar los archivos Excel
    echo     (MAESTRO.xlsx, QLICK_VTA_INTERNA.xlsx,
    echo      X_MOLECULA_2.xlsx, X_ATC.xlsx^)
    echo     si los tenes abiertos.
    echo   - Si el problema persiste, revisa el error de arriba.
    echo ============================================================
    pause
    exit /b %ERR%
)

echo ============================================================
echo   Dashboard generado correctamente.
echo   Abriendo dashboard.html en el navegador...
echo ============================================================
echo.

start "" "dashboard.html"
timeout /t 2 >nul
endlocal
