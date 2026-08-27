@echo off
REM Lanzador de doble clic para la migracion Access -> Supabase.
REM Vive junto a Migrar_A_Supabase.ps1, asi que funciona sin importar en que
REM carpeta este clonado el repositorio (%~dp0 = carpeta de este archivo).

setlocal
set RECOMENDADA=2025-03-01

echo ==============================================
echo   Migracion Base_Produccion.accdb -^> Supabase
echo ==============================================
echo.
echo El plan gratis de Supabase son 0.5 GB.
echo   18 meses de produccion = 0.34 GB  (cabe)
echo   Historico completo     = 0.99 GB  (NO cabe)
echo.
echo Opciones:
echo   [ENTER]        migrar desde %RECOMENDADA% (18 meses, recomendado)
echo   aaaa-mm-dd     migrar desde otra fecha
echo   TODO           migrar el historico completo (requiere plan de pago)
echo.
set /p DESDE="Tu eleccion: "

if "%DESDE%"=="" set DESDE=%RECOMENDADA%

if /i "%DESDE%"=="TODO" (
    echo.
    echo Migrando el historico COMPLETO. Va a superar el plan gratis.
    powershell -ExecutionPolicy Bypass -File "%~dp0Migrar_A_Supabase.ps1"
) else (
    echo.
    echo Migrando produccion desde %DESDE%.
    echo Puede tardar varios minutos. Es seguro cortar y volver a correrlo:
    echo no duplica lo ya cargado.
    powershell -ExecutionPolicy Bypass -File "%~dp0Migrar_A_Supabase.ps1" -DesdeFecha "%DESDE%"
)

echo.
echo Proceso terminado. Revisa los mensajes de arriba.
pause
endlocal
