<#
Lista todas las tablas y columnas reales de Base_Produccion.accdb (nombre,
tipo de dato, tamano). Sirve para construir el esquema de la base en la nube
a partir de la estructura real y actual, no de una version documentada que
puede haber quedado desactualizada por cambios hechos despues.
#>

param(
    [string]$RutaAccess = 'C:\Produccion\Base_Produccion.accdb',
    [string]$ArchivoSalida = 'C:\Produccion\Esquema_Base_Produccion.txt'
)

$ErrorActionPreference = 'Stop'

$cadena = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$RutaAccess;"
$conexion = New-Object System.Data.OleDb.OleDbConnection($cadena)
$conexion.Open()

$lineas = New-Object System.Collections.Generic.List[string]
$tablas = $conexion.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)

foreach ($fila in $tablas.Rows) {
    if ($fila['TABLE_TYPE'] -ne 'TABLE') { continue }
    $nombreTabla = [string]$fila['TABLE_NAME']
    if ($nombreTabla -like 'MSys*') { continue }

    $lineas.Add("TABLA: $nombreTabla")

    $restricciones = [object[]]@($null, $null, $nombreTabla, $null)
    $columnas = $conexion.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Columns, $restricciones)
    $columnasOrdenadas = $columnas.Rows | Sort-Object { [int]$_['ORDINAL_POSITION'] }
    foreach ($col in $columnasOrdenadas) {
        $lineas.Add("    $($col['COLUMN_NAME'])  -  tipo $($col['DATA_TYPE'])  -  tamano $($col['CHARACTER_MAXIMUM_LENGTH'])")
    }
    $lineas.Add('')
}

$conexion.Close()
$lineas | Out-File -FilePath $ArchivoSalida -Encoding UTF8
Write-Host "Esquema guardado en $ArchivoSalida"
