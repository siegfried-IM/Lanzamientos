@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

echo.
echo ============================================================
echo   Subiendo dashboard a GitHub (rama main)
echo ============================================================
echo.

REM Stage solo dashboard.html (la version regenerada)
git add dashboard.html
if errorlevel 1 (
    echo ERROR: git add fallo. ^Esta inicializado el repo?
    pause & exit /b 1
)

REM Verificar si hay cambios para commitear
git diff --cached --quiet
if not errorlevel 1 (
    echo No hay cambios en dashboard.html desde el ultimo push.
    echo.
    pause
    exit /b 0
)

for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set FECHA=%%a/%%b/%%c
for /f "tokens=1-2 delims=:. " %%a in ('time /t') do set HORA=%%a:%%b

git commit -m "Actualizacion dashboard %FECHA% %HORA%"
if errorlevel 1 (
    echo ERROR: git commit fallo.
    pause & exit /b 1
)

git push origin main
if errorlevel 1 (
    echo.
    echo ERROR: git push fallo. Revisa la conexion / credenciales.
    pause & exit /b 1
)

echo.
echo ============================================================
echo   OK. Dashboard subido a:
echo   https://github.com/siegfried-IM/Lanzamientos
echo.
echo   URL publica (una vez activado GitHub Pages):
echo   https://siegfried-im.github.io/Lanzamientos/
echo ============================================================
echo.
pause
endlocal
