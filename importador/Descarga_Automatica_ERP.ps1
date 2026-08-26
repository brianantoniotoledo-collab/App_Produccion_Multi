<#
Automatiza la descarga de "Detalle Cajas Etiquetadas" desde el Sistema Productivo
Comafri (SPC.exe) y deja el Excel resultante en Descargas_ERP para que
Importador_Produccion.ps1 lo procese como siempre.

Requiere que SPC.exe ya este abierto y logeado (Brian lo deja siempre abierto).
Pensado para correr por Programador de tareas, unos minutos antes de cada
corrida de Importador_Produccion.ps1.

PRIMERA VERSION: los nombres de ventana/boton se tomaron del video del proceso,
no de la app corriendo en vivo. Es esperable ajustar tiempos de espera o algun
nombre exacto la primera vez que se pruebe contra el SPC.exe real.
#>

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

$ErrorActionPreference = 'Stop'

$RutaBaseAccess    = 'C:\Produccion\Base_Produccion.accdb'
$CarpetaDescargas  = 'C:\Produccion\Descargas_ERP'
$ArchivoLog        = 'C:\Produccion\Descarga_Automatica_ERP_log.txt'
$TituloVentanaApp  = 'SISTEMA PRODUCTIVO COMAFRI'
$NombreProcesoApp  = 'ProjectDesposte_Insumos'
$TimeoutBusquedaSeg = 60

function Escribir-Log {
    param([string]$Mensaje)
    $linea = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Mensaje"
    Add-Content -Path $ArchivoLog -Value $linea
    Write-Host $linea
}

function Obtener-FechaDesde {
    $cadena = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$RutaBaseAccess;"
    $conexion = New-Object System.Data.OleDb.OleDbConnection($cadena)
    try {
        $conexion.Open()
        $comando = $conexion.CreateCommand()
        $comando.CommandText = 'SELECT MAX(FechaDesposte) AS Ultima FROM Cajas'
        $resultado = $comando.ExecuteScalar()
        if ($resultado -and $resultado -ne [DBNull]::Value) {
            # Se repite el ultimo dia ya importado como "desde": el indice unico
            # de NumeroCaja en Cajas evita duplicados si ese dia se vuelve a traer.
            return [datetime]$resultado
        }
    }
    finally {
        $conexion.Close()
    }
    return (Get-Date).AddDays(-7)
}

function Buscar-Elemento {
    param(
        [System.Windows.Automation.AutomationElement]$Origen = [System.Windows.Automation.AutomationElement]::RootElement,
        [string]$Nombre,
        [System.Windows.Automation.ControlType]$TipoControl,
        [int]$TimeoutSeg = $TimeoutBusquedaSeg
    )
    # Comparacion tolerante (sin distinguir mayusculas ni espacios sobrantes):
    # UI Automation puede exponer el texto del control distinto a como se ve en pantalla.
    if ($TipoControl) {
        $condicion = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $TipoControl)
    }
    else {
        $condicion = [System.Windows.Automation.Condition]::TrueCondition
    }

    $cronometro = [Diagnostics.Stopwatch]::StartNew()
    while ($cronometro.Elapsed.TotalSeconds -lt $TimeoutSeg) {
        $candidatos = $Origen.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condicion)
        foreach ($candidato in $candidatos) {
            if ($candidato.Current.Name.Trim() -ieq $Nombre.Trim()) { return $candidato }
        }
        Start-Sleep -Milliseconds 500
    }
    throw "No se encontro el elemento '$Nombre' en $TimeoutSeg segundos."
}

function Volcar-NombresDescendientes {
    param([System.Windows.Automation.AutomationElement]$Origen, [string]$Etiqueta)
    try {
        $todos = $Origen.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        $nombres = $todos | ForEach-Object { $_.Current.Name } | Where-Object { $_ } | Select-Object -Unique
        Escribir-Log "Diagnostico ($Etiqueta) - controles visibles: $($nombres -join ' | ')"
    }
    catch {
        Escribir-Log "Diagnostico ($Etiqueta) - no se pudo enumerar: $($_.Exception.Message)"
    }
}

function Buscar-ElementosPorClase {
    param(
        [System.Windows.Automation.AutomationElement]$Origen,
        [string]$NombreClase,
        [int]$TimeoutSeg = $TimeoutBusquedaSeg
    )
    $condicion = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ClassNameProperty, $NombreClase)
    $cronometro = [Diagnostics.Stopwatch]::StartNew()
    while ($cronometro.Elapsed.TotalSeconds -lt $TimeoutSeg) {
        $encontrados = $Origen.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condicion)
        if ($encontrados.Count -gt 0) { return $encontrados }
        Start-Sleep -Milliseconds 500
    }
    throw "No se encontraron controles de clase '$NombreClase' en $TimeoutSeg segundos."
}

function Invocar {
    param([System.Windows.Automation.AutomationElement]$Elemento)
    $patron = $null
    if ($Elemento.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$patron)) {
        $patron.Invoke()
        return
    }
    if ($Elemento.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$patron)) {
        $patron.Expand()
        return
    }
    throw "El elemento '$($Elemento.Current.Name)' no soporta Invoke ni ExpandCollapse."
}

function Escribir-Fecha {
    param(
        [System.Windows.Automation.AutomationElement]$Control,
        [datetime]$Fecha
    )
    $Control.SetFocus()
    Start-Sleep -Milliseconds 300
    # El DateTimePicker de Windows edita el segmento resaltado al tipear
    # digitos y avanza solo de dia -> mes -> anio (formato dd-MM-yyyy).
    $texto = $Fecha.ToString('ddMMyyyy')
    [System.Windows.Forms.SendKeys]::SendWait($texto)
    Start-Sleep -Milliseconds 300
}

function Esperar-Ventana {
    param([string]$TituloParcial, [int]$TimeoutSeg = $TimeoutBusquedaSeg)
    $cronometro = [Diagnostics.Stopwatch]::StartNew()
    while ($cronometro.Elapsed.TotalSeconds -lt $TimeoutSeg) {
        $condicion = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
        $ventanas = [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $condicion)
        foreach ($v in $ventanas) {
            if ($v.Current.Name -like "*$TituloParcial*") { return $v }
        }
        Start-Sleep -Milliseconds 500
    }
    throw "No aparecio la ventana '$TituloParcial' en $TimeoutSeg segundos."
}

try {
    Escribir-Log 'Inicio de descarga automatica.'

    if (-not (Get-Process -Name $NombreProcesoApp -ErrorAction SilentlyContinue)) {
        throw "SPC.exe no esta abierto. Debe quedar logeado en el PC para que esto funcione."
    }

    $fechaDesde = Obtener-FechaDesde
    $fechaHasta = Get-Date
    Escribir-Log "Rango a descargar: $($fechaDesde.ToString('dd-MM-yyyy')) a $($fechaHasta.ToString('dd-MM-yyyy'))."

    $ventanaApp = Esperar-Ventana -TituloParcial $TituloVentanaApp -TimeoutSeg 10
    $ventanaApp.SetFocus()

    $menuConsultas = Buscar-Elemento -Origen $ventanaApp -Nombre 'Consultas_Produccion'
    Invocar -Elemento $menuConsultas
    Start-Sleep -Milliseconds 500

    $itemDetalle = Buscar-Elemento -Nombre 'Detalle Cajas Etiquetadas'
    Invocar -Elemento $itemDetalle

    $ventanaBusqueda = Esperar-Ventana -TituloParcial 'Consulta Detalle Cajas Etiquetadas'

    $fechas = Buscar-ElementosPorClase -Origen $ventanaBusqueda -NombreClase 'SysDateTimePick32'
    if ($fechas.Count -lt 2) {
        throw "Se esperaban 2 campos de fecha y se encontraron $($fechas.Count)."
    }
    Escribir-Fecha -Control $fechas[0] -Fecha $fechaDesde
    Escribir-Fecha -Control $fechas[1] -Fecha $fechaHasta

    $botonBuscar = Buscar-Elemento -Origen $ventanaBusqueda -Nombre 'Buscar' -TipoControl ([System.Windows.Automation.ControlType]::Button)
    Invocar -Elemento $botonBuscar

    $botonExcel = Buscar-Elemento -Nombre 'Excel' -TipoControl ([System.Windows.Automation.ControlType]::Button) -TimeoutSeg 90
    Escribir-Log 'Grilla cargada, exportando a Excel.'

    $procesosExcelAntes = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    Invocar -Elemento $botonExcel

    $cronometro = [Diagnostics.Stopwatch]::StartNew()
    $procesoExcelNuevo = $null
    while ($cronometro.Elapsed.TotalSeconds -lt 30) {
        $procesoExcelNuevo = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Where-Object { $procesosExcelAntes -notcontains $_.Id }
        if ($procesoExcelNuevo) { break }
        Start-Sleep -Milliseconds 500
    }
    if (-not $procesoExcelNuevo) {
        throw 'Excel no se abrio despues de pulsar el boton Excel.'
    }
    Start-Sleep -Seconds 2

    if (-not (Test-Path $CarpetaDescargas)) {
        New-Item -ItemType Directory -Path $CarpetaDescargas | Out-Null
    }
    $nombreArchivo = "Produccion_ERP_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    $rutaCompleta = Join-Path $CarpetaDescargas $nombreArchivo

    [System.Windows.Forms.SendKeys]::SendWait('^s')
    Start-Sleep -Seconds 1
    $dialogoGuardar = Esperar-Ventana -TituloParcial 'Guardar como' -TimeoutSeg 15
    $dialogoGuardar.SetFocus()
    Start-Sleep -Milliseconds 300
    [System.Windows.Forms.SendKeys]::SendWait($rutaCompleta)
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Seconds 1
    # Excel puede preguntar si se mantiene el formato .xlsx; Enter confirma la opcion por defecto.
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -Seconds 2

    if (-not (Test-Path $rutaCompleta)) {
        throw "No se genero el archivo esperado en $rutaCompleta."
    }
    Escribir-Log "Archivo guardado: $rutaCompleta"

    Stop-Process -Id $procesoExcelNuevo.Id -Force -ErrorAction SilentlyContinue

    Escribir-Log 'Descarga automatica finalizada con exito.'
}
catch {
    Escribir-Log "ERROR: $($_.Exception.Message)"
    if ($ventanaApp) {
        Volcar-NombresDescendientes -Origen $ventanaApp -Etiqueta 'ventana principal al fallar'
    }
    throw
}
