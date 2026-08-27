@echo off
REM Lanzador de doble clic para la migracion Access -> Supabase.
REM Vive junto a Migrar_A_Supabase.ps1, asi que funciona sin importar en que
REM carpeta este clonado el repositorio (%~dp0 = carpeta de este archivo).

echo ============================================
echo   Migracion Base_Produccion.accdb -^> Supabase
echo ============================================
echo.
set /p DESDE="Migrar produccion desde que fecha? (aaaa-mm-dd, o ENTER para migrar TODO): "

if "%DESDE%"=="" (
    echo.
    echo Migrando el historico COMPLETO. Puede tardar y superar el limite
    echo de 500 MB del plan gratis de Supabase.
    powershell -ExecutionPolicy Bypass -File "%~dp0Migrar_A_Supabase.ps1"
) else (
    echo.
    echo Migrando produccion desde %DESDE%.
    powershell -ExecutionPolicy Bypass -File "%~dp0Migrar_A_Supabase.ps1" -DesdeFecha "%DESDE%"
)

echo.
echo Proceso terminado. Revisa los mensajes de arriba.
pause
