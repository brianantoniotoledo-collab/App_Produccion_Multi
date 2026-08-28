@echo off
REM Sube a la nube solo los JSON de configuracion (zonas, cerdos por dia,
REM peso de vara, tipos de cambio). NO abre el Access, asi que sirve igual
REM desde el notebook, que no tiene la base de produccion.
REM
REM Lee los JSON de la carpeta app_original del propio repositorio.

echo ==========================================
echo   Subir configuraciones a Supabase
echo ==========================================
echo.
echo Sube: zonas del cerdo, cerdos por dia, peso de vara y tipos de cambio.
echo No necesita el Access: lee los JSON del repositorio.
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0Migrar_A_Supabase.ps1" ^
    -SoloConfiguraciones -CarpetaJson "%~dp0..\app_original"

echo.
echo Proceso terminado. Revisa los mensajes de arriba.
pause
