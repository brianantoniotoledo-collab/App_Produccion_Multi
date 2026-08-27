' Lanzador sin ventana de consola para la sincronizacion Access -> Supabase.
' Pensado para el Programador de tareas: corre cada 20 minutos detras de
' Importador_Produccion.ps1 y sube a la nube solo lo que llego nuevo.
'
' Mismo patron que KPI_Produccion.vbs: resuelve su propia carpeta, asi funciona
' sin importar donde este clonado el repositorio.

Dim carpeta, shell, comando
carpeta = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set shell = CreateObject("WScript.Shell")

comando = "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
          carpeta & "\Migrar_A_Supabase.ps1"" -SoloNuevos"

' 0 = ventana oculta, False = no esperar a que termine
shell.Run comando, 0, False
