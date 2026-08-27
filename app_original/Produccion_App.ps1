# ============================================================
#  GIVEAWAY PRODUCCION - COMAFRI   (app standalone, v3)
#  Primera pieza de la app "solo Produccion".
#  Se actualiza sola a medida que el importador carga datos.
#
#  Patron reciclado de KPI_Produccion_App.ps1:
#   - Ventana WPF que queda abierta
#   - Conexion a Access abierta y cerrada POR OPERACION
#   - Reglas: excluye COM_DSK_SUBPROD, solo AASA y Agrosuper
#
#  Autor: Brian + Claude | Comafri
# ============================================================

# ------------------------------------------------------------
#  CONFIGURACION
# ------------------------------------------------------------
$RutaBase          = "C:\Produccion\Base_Produccion.accdb"
$SegundosRefresco  = 90       # cada cuanto se relee la base sola
$PasoBin           = 0.05     # ancho del intervalo del histograma, en kg
$UsuarioExcluido   = "COM_DSK_SUBPROD"
$Cliente           = "AASA PORK LIMITADA"   # regla global: solo AASA
$UmbralGiveawayAlto= 3.0      # sobre este % la barra se pinta naranja

# Productos de PESO FIJO. Son los unicos donde aplica giveaway.
# Giveaway de cada caja = Peso Neto - Peso Neto Etiqueta
#
# OJO: se filtra por CODIGO, no por nombre. Patas y Manos tienen
# una version a granel con el MISMO nombre pero peso variable
# (130191 y 130193), que NO debe entrar al calculo.
$ProductosGiveaway = @(
    [pscustomobject]@{ Codigo = 130154; Nombre = "FILETE CERDO C/CABEZA CONGELADO VP UE AF" }
    [pscustomobject]@{ Codigo = 130073; Nombre = "PATAS CERDO CONGELADO" }
    [pscustomobject]@{ Codigo = 130145; Nombre = "MANOS CERDO CONGELADO" }
    [pscustomobject]@{ Codigo = 130374; Nombre = "TIRA DE LOMO DE CERDO CONGELADO" }
    [pscustomobject]@{ Codigo = 130051; Nombre = "PLATEADA DE LOMO CONGELADO VP" }
)


# Precio de venta por KILO de cada producto de peso fijo, tal como
# viene en las facturas de exportacion (cada uno en su moneda).
# Con esto el giveaway en kilos se traduce a plata real.
$PreciosGiveaway = @{
    130154 = @{ Precio = 4.30;   Moneda = 'EUR' }   # FILETE CERDO C/CABEZA VP UE AF
    130073 = @{ Precio = 3.25;   Moneda = 'USD' }   # PATAS CERDO CONGELADO
    130145 = @{ Precio = 3.30;   Moneda = 'USD' }   # MANOS CERDO CONGELADO
    130374 = @{ Precio = 450.00; Moneda = 'JPY' }   # TIRA DE LOMO DE CERDO CONGELADO
    130051 = @{ Precio = 6.50;   Moneda = 'USD' }   # PLATEADA DE LOMO CONGELADO VP
}

# Cuanto vale 1 unidad de cada moneda en pesos chilenos.
# Se editan desde el boton "Tipos de cambio" y quedan guardados
# en Tipos_Cambio.json; estos son solo el valor de partida.
$RutaTiposCambio  = "C:\Produccion\Tipos_Cambio.json"
$TiposCambioBase  = @{ USD = 913.83; EUR = 1055.00; JPY = 5.76 }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
#  PALETA
# ------------------------------------------------------------
$C = @{
    Fondo     = "#070B18"
    Panel     = "#0E1530"
    Campo     = "#0A1128"
    Linea     = "#22305C"
    Texto     = "#E8EDFA"
    Texto2    = "#8C9BC4"
    Azul      = "#1B3F94"
    Naranja   = "#F7941D"
    NeonAzul  = "#3D8BFF"
    NeonCyan  = "#00E5FF"
    NeonNaranja="#FFA829"
    NeonRojo  = "#FF3B6B"
    NeonVerde = "#00FFA3"
    NeonVioleta="#A96BFF"
}

function Pincel { param([string]$Hex) return (New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($Hex))) }

# ------------------------------------------------------------
#  CAPA DE DATOS
#  'return ,$tabla' -> la coma evita que PowerShell desenrolle
#  el DataTable en un arreglo de DataRow.
# ------------------------------------------------------------
function Invoke-ConsultaAccess {
    param([string]$Sql)

    $cadena = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$RutaBase;Persist Security Info=False;"
    $conexion = New-Object System.Data.OleDb.OleDbConnection($cadena)
    $tabla = New-Object System.Data.DataTable
    try {
        $conexion.Open()
        $comando = New-Object System.Data.OleDb.OleDbCommand($Sql, $conexion)
        $comando.CommandTimeout = 600
        $adaptador = New-Object System.Data.OleDb.OleDbDataAdapter($comando)
        [void]$adaptador.Fill($tabla)
    }
    finally {
        if ($conexion.State -ne [System.Data.ConnectionState]::Closed) { $conexion.Close() }
        $conexion.Dispose()
    }
    return ,$tabla
}

function Format-FechaAccess {
    param([datetime]$Fecha)
    return "#" + $Fecha.ToString("MM/dd/yyyy") + "#"
}

# ------------------------------------------------------------
#  EXTRACCION AGREGADA
#  Giveaway por caja = PesoNeto - PesoNetoEtiqueta
#  Se agrupa en SQL para no traer cientos de miles de filas.
# ------------------------------------------------------------
function Get-DatosGiveaway {
    param([datetime]$Desde, [datetime]$Hasta)

    $listaCod = ($ProductosGiveaway | ForEach-Object { [string]$_.Codigo }) -join ","
    $fDesde = Format-FechaAccess $Desde.Date
    $fHasta = Format-FechaAccess $Hasta.Date.AddDays(1)
    $factor = [int][math]::Round(1.0 / $PasoBin, 0)

    $sql = @"
SELECT CodigoProducto,
       ProductoEspanol,
       Format(FechaPesaje, 'yyyy-mm-dd') AS Dia,
       Int((PesoNeto - PesoNetoEtiqueta) * $factor) AS Bin,
       IIF(PesoNeto < PesoNetoEtiqueta, 1, 0) AS EsBajo,
       COUNT(*)               AS Cajas,
       SUM(PesoNeto)          AS KgReal,
       SUM(PesoNetoEtiqueta)  AS KgEtiqueta
FROM Cajas
WHERE PesoNeto > 0
  AND PesoNetoEtiqueta > 0
  AND FechaPesaje >= $fDesde
  AND FechaPesaje <  $fHasta
  AND Usuario <> '$UsuarioExcluido'
  AND NombreCliente = '$Cliente'
  AND CodigoProducto IN ($listaCod)
GROUP BY CodigoProducto, ProductoEspanol, Format(FechaPesaje, 'yyyy-mm-dd'),
         Int((PesoNeto - PesoNetoEtiqueta) * $factor),
         IIF(PesoNeto < PesoNetoEtiqueta, 1, 0)
"@

    $tabla = Invoke-ConsultaAccess -Sql $sql

    $hechos = New-Object System.Collections.Generic.List[object]
    foreach ($fila in $tabla.Rows) {
        $dia = [string]$fila["Dia"]
        $hechos.Add([pscustomobject]@{
            Codigo     = [int]$fila["CodigoProducto"]
            Producto   = ([string]$fila["ProductoEspanol"]).Trim()
            Dia        = $dia
            Mes        = $dia.Substring(0, 7)
            Bin        = [int]$fila["Bin"]
            Cajas      = [int]$fila["Cajas"]
            KgReal     = [double]$fila["KgReal"]
            KgEtiqueta = [double]$fila["KgEtiqueta"]
            Bajo       = ([int]$fila["EsBajo"] -eq 1)
        })
    }
    return $hechos
}

# ------------------------------------------------------------
#  CUADRE DE PIEZAS - ZONAS Y CONFIGURACION
#  Las asignaciones material->zona y los cerdos por dia viven en
#  archivos JSON aparte (NO en la base Access), para no escribir
#  en la base que llena el importador.
# ------------------------------------------------------------
$RutaConfigZonas  = "C:\Produccion\Config_Zonas.json"
$RutaCerdosDia    = "C:\Produccion\Cerdos_Por_Dia.json"
$CerdosPorDefecto = 700

# Silueta del cerdo (calcada de la referencia britanica, 2 patas, cola enroscada)
$CerdoCuerpo = 'F0 M726.9,90.2 L722.6,81.4 L715.4,75.6 L708.8,73.7 L698.9,74.6 L689.2,81.6 L692.1,90.2 L691.4,96.3 L682.2,103.6 L676.9,104.5 L670.3,102.1 L649.4,88.8 L630.7,80.6 L601.0,72.1 L568.0,67.1 L510.8,66.0 L454.7,70.4 L414.0,70.4 L382.1,67.1 L359.0,67.1 L326.0,71.3 L273.2,85.0 L240.2,98.1 L232.5,99.0 L211.6,90.8 L180.8,86.7 L162.6,76.2 L162.6,81.4 L172.6,94.6 L182.1,115.5 L190.9,125.4 L191.5,133.1 L182.2,151.8 L178.9,169.4 L171.4,180.4 L159.9,192.5 L148.9,200.5 L142.3,202.4 L129.6,202.4 L128.6,210.1 L134.3,224.4 L142.3,235.1 L147.8,237.4 L159.9,236.7 L172.0,244.7 L188.5,242.0 L225.9,244.2 L246.8,239.0 L261.1,238.7 L275.4,241.8 L293.0,251.1 L313.4,258.3 L316.1,266.2 L322.7,326.7 L313.3,357.5 L297.4,376.2 L297.9,379.0 L300.7,380.9 L330.4,381.9 L338.1,390.0 L357.9,390.5 L367.5,387.7 L370.9,375.6 L381.3,371.6 L382.9,365.2 L377.7,342.1 L377.7,325.6 L389.3,270.6 L393.3,264.5 L481.1,275.0 L498.7,275.0 L524.0,271.2 L557.0,260.7 L583.4,249.5 L593.3,248.6 L597.9,251.9 L614.0,293.7 L616.4,303.6 L616.4,316.8 L610.9,335.5 L601.3,356.4 L595.2,365.2 L580.4,377.3 L579.6,382.5 L590.0,386.1 L607.6,385.9 L611.2,383.7 L615.5,374.9 L626.9,371.6 L629.1,367.4 L631.2,348.7 L637.0,332.2 L653.8,302.5 L657.3,299.7 L662.6,299.2 L668.6,302.5 L674.7,314.6 L678.0,343.2 L676.7,359.7 L665.1,378.4 L665.0,381.5 L669.2,383.9 L689.0,383.1 L695.8,380.1 L697.2,368.7 L705.7,364.7 L707.1,356.4 L700.6,337.7 L698.9,324.5 L702.8,278.3 L691.5,266.2 L687.9,259.6 L685.7,249.7 L685.7,231.0 L698.9,201.3 L703.6,183.7 L705.5,166.1 L700.2,132.0 L702.4,129.2 L712.1,125.9 L717.0,121.0 L726.4,102.3 Z M690.8,106.7 L694.5,104.5 L706.6,105.8 L710.8,110.0 L710.3,117.5 L704.4,121.0 L694.5,119.7 L690.3,115.5 Z M694.1,86.9 L700.0,81.8 L713.2,82.5 L718.3,88.0 L718.5,96.8 L716.5,98.8 L711.0,99.0 L695.6,96.6 L693.4,93.5 Z'

# Paneles de la derecha. No son zonas del cerdo: el cuero cubre todo el
# animal y el recorte sale del prolijado de varias piezas, asi que no
# tienen posicion anatomica. Se miden en KG POR CERDO, no en piezas.
$PanelesCuadre = @(
    [pscustomobject]@{ Id='cuero';    Nombre='CUERO Y GRASAS';        Nota='cubre todo el animal, no una zona';   Color='#FFA829' }
    [pscustomobject]@{ Id='recortes'; Nombre='TRIMMINGS Y RECORTES';  Nota='prolijado de pulpas, filete, despuntes'; Color='#00E5FF' }
    [pscustomobject]@{ Id='huesos';   Nombre='HUESOS';                Nota='hueso de todas las piezas';            Color='#5B6EA8' }
    [pscustomobject]@{ Id='despojos'; Nombre='DESPOJOS Y DECOMISOS';  Nota='NO se venden - nombre trae DESPOJO/DECOMISO'; Color='#A96BFF' }
    [pscustomobject]@{ Id='subprod';  Nombre='SUBPRODUCTOS';          Nota='SI se venden - estacion COM_DSK_SUBPROD';     Color='#F06BC8' }
)
$IdsPaneles = @('cuero','recortes','huesos','despojos','subprod')

# Peso promedio de la vara fria. Rango real de planta: 95 a 102 kg.
$RutaPesoVara  = "C:\Produccion\Peso_Vara.json"
$PesoVaraBase  = 98.5

$EstriasBaby = @(
    'M419.2,91.0 L419.2,121.0',
    'M431.4,91.0 L431.4,121.0',
    'M443.5,91.0 L443.5,121.0',
    'M455.7,91.0 L455.7,121.0',
    'M467.8,91.0 L467.8,121.0',
    'M480.0,91.0 L480.0,121.0',
    'M492.2,91.0 L492.2,121.0',
    'M504.3,91.0 L504.3,121.0',
    'M516.5,91.0 L516.5,121.0',
    'M528.6,91.0 L528.6,121.0',
    'M540.8,91.0 L540.8,121.0'
)

# 14 zonas comerciales. Mult = piezas por cerdo.
# Modo de etiqueta: In = dentro del cuerpo, T = arriba con linea guia,
# L = columna izquierda, R = columna derecha.
# El filete es un corte interno (bajo el lomo, region lumbar hacia el jamon):
# se dibuja como lente punteada sobre el limite lomo/panceta.
$ZonasCuadre = @(
    [pscustomobject]@{ Id='cabeza'; Nombre='Cabeza'; Mult=1; Modo='In'; LX=205; LY=116; AX=0; AY=0; D='M81.8,-28.6 L294.1,-28.6 L297.4,42.9 L300.7,108.9 L300.7,183.7 L197.3,243.1 L81.8,243.1 Z' }
    [pscustomobject]@{ Id='papada'; Nombre='Papada'; Mult=2; Modo='In'; LX=258; LY=212; AX=0; AY=0; D='M300.7,183.7 L197.3,243.1 L188.5,290.4 L301.8,290.4 L285.3,240.9 L296.3,207.9 Z' }
    [pscustomobject]@{ Id='collar'; Nombre='Collar / Vetado'; Mult=2; Modo='In'; LX=352; LY=96; AX=0; AY=0; D='M294.1,20.9 L386.5,20.9 L381.0,85.8 L375.5,108.9 L367.8,132.0 L352.4,155.1 L308.4,178.2 L301.8,157.3 L298.5,116.6 L294.1,75.9 Z' }
    [pscustomobject]@{ Id='paleta'; Nombre='Paleta'; Mult=2; Modo='In'; LX=342; LY=232; AX=0; AY=0; D='M308.4,178.2 L352.4,155.1 L367.8,132.0 L375.5,108.9 L381.0,85.8 L386.5,20.9 L395.3,20.9 L395.3,299.2 L272.1,299.2 L283.1,240.9 L294.1,205.7 Z' }
    [pscustomobject]@{ Id='pernilmano'; Nombre='Pernil mano'; Mult=2; Modo='L'; LX=272; LY=300; AX=330; AY=322; D='M276.5,299.2 L404.1,299.2 L404.1,346.5 L276.5,346.5 Z' }
    [pscustomobject]@{ Id='mano'; Nombre='Mano'; Mult=2; Modo='L'; LX=272; LY=350; AX=330; AY=370; D='M276.5,346.5 L404.1,346.5 L404.1,403.7 L276.5,403.7 Z' }
    [pscustomobject]@{ Id='centro'; Nombre='Chuleta / Lomo centro'; Mult=2; Modo='T'; LX=470; LY=22; AX=470; AY=80; D='M395.3,20.9 L598.8,20.9 L584.5,64.9 L569.1,108.9 L563.6,155.1 L568.0,185.9 L571.3,203.5 L395.3,203.5 Z' }
    [pscustomobject]@{ Id='baby'; Nombre='Baby back'; Mult=2; Modo='In'; LX=440; LY=134; AX=0; AY=0; D='M556.0,100.5 L552.4,89.8 L546.6,86.0 L541.2,86.1 L537.4,88.9 L532.8,88.8 L528.2,87.4 L505.3,88.0 L500.2,86.1 L496.9,86.0 L493.0,89.5 L438.8,89.8 L435.4,90.5 L427.3,89.4 L421.4,91.7 L413.5,91.2 L409.4,93.0 L404.9,93.4 L404.0,116.7 L405.8,122.4 L408.4,124.0 L420.4,124.0 L424.7,122.8 L429.8,124.0 L434.2,123.1 L442.8,123.8 L448.2,121.8 L454.0,122.8 L463.0,121.7 L467.3,122.7 L472.4,122.0 L490.2,124.0 L494.6,125.7 L503.2,124.6 L511.1,126.0 L520.3,124.8 L527.2,126.0 L534.6,124.4 L542.2,125.7 L545.0,124.6 L550.1,125.6 L554.4,121.9 L554.5,112.7 L556.0,108.6 Z' }
    [pscustomobject]@{ Id='panceta'; Nombre='Pecho / Panceta'; Mult=2; Modo='In'; LX=452; LY=210; AX=0; AY=0; D='M395.3,203.5 L571.3,203.5 L582.3,221.1 L595.5,244.2 L595.5,309.1 L395.3,309.1 Z' }
    [pscustomobject]@{ Id='filete'; Nombre='Filete'; Mult=2; Modo='In'; LX=526; LY=166; AX=0; AY=0; D='M472.3,185.9 L486.6,170.5 L513.0,161.7 L547.1,160.6 L580.1,174.9 L551.5,196.9 L513.0,203.5 L485.5,199.1 Z' }
    [pscustomobject]@{ Id='pierna'; Nombre='Pierna'; Mult=2; Modo='In'; LX=652; LY=168; AX=0; AY=0; D='M598.8,20.9 L780.3,20.9 L780.3,298.1 L595.5,298.1 L595.5,244.2 L582.3,221.1 L571.3,203.5 L568.0,185.9 L563.6,155.1 L569.1,108.9 L584.5,64.9 Z' }
    [pscustomobject]@{ Id='pernilpierna'; Nombre='Pernil pierna'; Mult=2; Modo='R'; LX=742; LY=300; AX=660; AY=322; D='M566.9,298.1 L734.1,298.1 L734.1,346.5 L566.9,346.5 Z' }
    [pscustomobject]@{ Id='pata'; Nombre='Pata'; Mult=2; Modo='R'; LX=742; LY=350; AX=660; AY=370; D='M566.9,346.5 L734.1,346.5 L734.1,403.7 L566.9,403.7 Z' }
    [pscustomobject]@{ Id='cola'; Nombre='Cola'; Mult=1; Modo='R'; LX=748; LY=84; AX=724; AY=102; D='ELIPSE:677,74,51,56' }
)
$SumaMultiplicadores = 0
foreach ($z in $ZonasCuadre) { $SumaMultiplicadores = $SumaMultiplicadores + $z.Mult }

# ---- JSON: mapas simples clave -> valor ----
function Read-MapaJson {
    param([string]$Ruta)
    $mapa = @{}
    if (Test-Path $Ruta) {
        try {
            $obj = Get-Content -Path $Ruta -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $obj.PSObject.Properties) { $mapa[[string]$p.Name] = $p.Value }
        } catch { }
    }
    return $mapa
}

function Write-MapaJson {
    param([string]$Ruta, [hashtable]$Mapa)
    $obj = New-Object PSObject
    foreach ($k in ($Mapa.Keys | Sort-Object)) {
        $obj | Add-Member -MemberType NoteProperty -Name ([string]$k) -Value $Mapa[$k]
    }
    ($obj | ConvertTo-Json -Depth 3) | Set-Content -Path $Ruta -Encoding UTF8
}

# ---- Asignacion automatica por nombre (solo la primera vez) ----
function Get-ZonaPorNombre {
    param([string]$Nombre)
    $n = ([string]$Nombre).ToUpperInvariant()

    # Perniles primero (contienen MANO / PIERNA)
    if ($n.Contains('PERNIL MANO'))   { return 'pernilmano' }
    if ($n.Contains('PERNIL PIERNA')) { return 'pernilpierna' }
    # Pierna entera antes de la exclusion por CUERO
    if ($n.Contains('PIERNA ENTERA')) { return 'pierna' }

    # Excluidos: no son piezas anatomicas contables.
    # Antes de revisar, se limpian calificativos que contienen palabras
    # excluidas pero NO convierten al producto en subproducto
    # (DESGRASADA contiene GRASA, C/CUERO contiene CUERO, C/HUESO contiene HUESO).
    $nl = $n
    foreach ($q in @('DESGRASAD','DESHUESAD','C/CUERO','S/CUERO','CON CUERO','SIN CUERO','C/HUESO','S/HUESO','CON HUESO','SIN HUESO','C/GRASA','S/GRASA','CON GRASA','SIN GRASA')) {
        $nl = $nl.Replace($q, ' ')
    }
    # Estos ya no se excluyen: alimentan los paneles de la derecha.
    # El orden importa: HUESO antes que nada porque "HUESO DE PIERNA"
    # debe ir a huesos y no a la zona pierna.
    $rinon = 'RI' + [char]209 + 'ON'
    if ($nl.Contains('HUESO') -or $nl.Contains('COSTILLAR HUESO')) { return 'huesos' }
    foreach ($pal in @('TRIMMING','TRIMING','RECORTE','DESPUNTE','RETAZO','PROLIJAD','FALDA')) {
        if ($nl.Contains($pal)) { return 'recortes' }
    }
    foreach ($pal in @('CUERO','GRASA','GORDURA','MANTECA','TOCINO','LARDO','EMPELLA')) {
        if ($nl.Contains($pal)) { return 'cuero' }
    }
    # Solo por nombre explicito: lo demas llega marcado por la estacion.
    foreach ($pal in @('DECOMISO','DESPOJO')) {
        if ($nl.Contains($pal)) { return 'despojos' }
    }

    # Filete trae "C/CABEZA" en el nombre: se evalua antes que CABEZA
    # para que no caiga en la zona equivocada.
    if ($n.Contains('FILETE')) { return 'filete' }

    if ($n.Contains('PATAS'))  { return 'pata' }
    if ($n.Contains('MANOS'))  { return 'mano' }
    if ($n.Contains('COLA'))   { return 'cola' }
    if ($n.Contains('CABEZA')) { return 'cabeza' }
    if ($n.Contains('PAPADA')) { return 'papada' }
    if ($n.Contains('COLLAR') -or $n.Contains('CHULETA VETADA') -or $n.Contains('LOMO VETADO')) { return 'collar' }
    # BABY BACK va antes que LOMO: son las costillas dorsales, zona propia.
    # Antes caia en 'centro' e inflaba el porcentaje del lomo.
    if ($n.Contains('BABY BACK') -or $n.Contains('BABYBACK')) { return 'baby' }
    if ($n.Contains('CHULETA CENTRO') -or $n.Contains('LOMO CENTRO') -or $n.Contains('LOMO MM')) { return 'centro' }
    if ($n.Contains('PANCETA') -or $n.Contains('PECHITO') -or $n.Contains('PECHO')) { return 'panceta' }
    if ($n.Contains('PULPA PIERNA')) { return 'pierna' }
    if ($n.Contains('PALETA')) { return 'paleta' }
    return $null
}

function Initialize-ConfigZonas {
    $script:CfgZonas = Read-MapaJson $RutaConfigZonas
    # 130136 despunte y 130084 falda van a recortes; 130223 a decomisos.
    # Se corrige una sola vez la clasificacion vieja que los dejaba excluidos.
    foreach ($par in @(@('130136','recortes'), @('130084','recortes'), @('130223','despojos'))) {
        if ($script:CfgZonas[$par[0]] -eq 'excluido') { $script:CfgZonas[$par[0]] = $par[1] }
    }
    # Se siembran por nombre los codigos que aun no tienen decision guardada.
    # Los que el usuario ya asigno (incluido 'sinasignar') no se tocan.
    try {
        $fDesde = Format-FechaAccess (Get-Date).Date.AddDays(-60)
        $sql = @"
SELECT CodigoProducto, ProductoEspanol
FROM Cajas
WHERE Piezas > 0
  AND FechaPesaje >= $fDesde
  AND Usuario <> '$UsuarioExcluido'
  AND NombreCliente = '$Cliente'
GROUP BY CodigoProducto, ProductoEspanol
"@
        $t = Invoke-ConsultaAccess -Sql $sql
        $nuevos = 0
        foreach ($f in $t.Rows) {
            $cod = [string][int]$f["CodigoProducto"]
            if ($script:CfgZonas.ContainsKey($cod)) { continue }
            $zid = Get-ZonaPorNombre ([string]$f["ProductoEspanol"])
            if ($zid) { $script:CfgZonas[$cod] = $zid; $nuevos = $nuevos + 1 }
        }
        if ($nuevos -gt 0) { Write-MapaJson $RutaConfigZonas $script:CfgZonas }
    } catch { }
}

function Get-PesoVara {
    if (Test-Path $RutaPesoVara) {
        try {
            $j = Get-Content $RutaPesoVara -Raw -Encoding UTF8 | ConvertFrom-Json
            $v = 0.0
            if ([double]::TryParse([string]$j.PesoVara, [ref]$v) -and $v -ge 40 -and $v -le 200) { return $v }
        } catch { }
    }
    return $PesoVaraBase
}

function Set-PesoVara {
    param([double]$Valor)
    $script:PesoVara = $Valor
    try {
        $o = New-Object psobject
        Add-Member -InputObject $o -MemberType NoteProperty -Name 'PesoVara' -Value $Valor
        $o | ConvertTo-Json | Set-Content -Path $RutaPesoVara -Encoding UTF8
    } catch { }
}

function Get-CerdosDeFecha {
    param([datetime]$Fecha)
    $k = $Fecha.ToString('yyyy-MM-dd')
    if ($script:CerdosPorDia.ContainsKey($k)) { return [int]$script:CerdosPorDia[$k] }
    return $CerdosPorDefecto
}

# ---- Consultas del cuadre ----
function Get-MaterialesDia {
    param([datetime]$Fecha)
    $fDesde = Format-FechaAccess $Fecha.Date
    $fHasta = Format-FechaAccess $Fecha.Date.AddDays(1)
    $sql = @"
SELECT CodigoProducto, ProductoEspanol,
       IIF(Usuario = '$UsuarioExcluido', 1, 0) AS EsSub,
       SUM(Piezas)   AS Piezas,
       SUM(PesoNeto) AS Kg,
       COUNT(*)      AS Cajas
FROM Cajas
WHERE FechaPesaje >= $fDesde
  AND FechaPesaje <  $fHasta
  AND NombreCliente = '$Cliente'
GROUP BY CodigoProducto, ProductoEspanol, IIF(Usuario = '$UsuarioExcluido', 1, 0)
"@
    $t = Invoke-ConsultaAccess -Sql $sql
    $lista = New-Object System.Collections.Generic.List[object]
    foreach ($f in $t.Rows) {
        $lista.Add([pscustomobject]@{
            Codigo   = [int]$f["CodigoProducto"]
            Producto = ([string]$f["ProductoEspanol"]).Trim()
            Piezas   = $(if ($f["Piezas"] -is [System.DBNull]) { 0 } else { [int][math]::Round([double]$f["Piezas"], 0) })
            Kg       = $(if ($f["Kg"] -is [System.DBNull]) { 0.0 } else { [double]$f["Kg"] })
            EsSub    = $(if ($f["EsSub"] -is [System.DBNull]) { 0 } else { [int]$f["EsSub"] })
            Cajas    = [int]$f["Cajas"]
        })
    }
    return $lista.ToArray()
}

function Get-TendenciaCuadre {
    param([datetime]$Fecha)
    $iniMes = Get-Date -Year $Fecha.Year -Month $Fecha.Month -Day 1
    $ini14  = $Fecha.Date.AddDays(-13)
    $ini    = $iniMes.Date
    if ($ini14 -lt $ini) { $ini = $ini14 }
    $fDesde = Format-FechaAccess $ini
    $fHasta = Format-FechaAccess $Fecha.Date.AddDays(1)
    $sql = @"
SELECT Format(FechaPesaje, 'yyyy-mm-dd') AS Dia,
       CodigoProducto,
       SUM(Piezas) AS Piezas
FROM Cajas
WHERE Piezas > 0
  AND FechaPesaje >= $fDesde
  AND FechaPesaje <  $fHasta
  AND Usuario <> '$UsuarioExcluido'
  AND NombreCliente = '$Cliente'
GROUP BY Format(FechaPesaje, 'yyyy-mm-dd'), CodigoProducto
"@
    $t = Invoke-ConsultaAccess -Sql $sql
    $raw = @{}
    foreach ($f in $t.Rows) {
        $d = [string]$f["Dia"]
        if (-not $raw.ContainsKey($d)) { $raw[$d] = @{} }
        $raw[$d][[string][int]$f["CodigoProducto"]] = [int][math]::Round([double]$f["Piezas"], 0)
    }
    return $raw
}

# ---- Calculos puros (testeables sin interfaz) ----

function Get-HistorialAprovechamiento {
    # Reparto del kilaje por dia para los ultimos N dias. Sirve para ver
    # si una categoria se dispara (tipico: decomisos).
    param([datetime]$Hasta, [int]$Dias = 30, [hashtable]$Cfg)

    $fDesde = Format-FechaAccess $Hasta.Date.AddDays(-($Dias - 1))
    $fHasta = Format-FechaAccess $Hasta.Date.AddDays(1)
    $sql = @"
SELECT Format(FechaPesaje,'yyyy-mm-dd') AS Dia,
       CodigoProducto, ProductoEspanol,
       IIF(Usuario = '$UsuarioExcluido', 1, 0) AS EsSub,
       SUM(PesoNeto) AS Kg
FROM Cajas
WHERE FechaPesaje >= $fDesde
  AND FechaPesaje <  $fHasta
  AND NombreCliente = '$Cliente'
GROUP BY Format(FechaPesaje,'yyyy-mm-dd'), CodigoProducto, ProductoEspanol,
         IIF(Usuario = '$UsuarioExcluido', 1, 0)
ORDER BY Format(FechaPesaje,'yyyy-mm-dd')
"@
    try { $t = Invoke-ConsultaAccess -Sql $sql } catch { return ,@() }

    $porDia = [ordered]@{}
    foreach ($f in $t.Rows) {
        $dia = [string]$f["Dia"]
        if (-not $porDia.Contains($dia)) {
            $o = [ordered]@{ Dia = $dia; Total = 0.0; cortes = 0.0 }
            foreach ($p in $PanelesCuadre) { $o[$p.Id] = 0.0 }
            $porDia[$dia] = $o
        }
        $kg = 0.0
        if ($f["Kg"] -isnot [System.DBNull]) { $kg = [double]$f["Kg"] }
        if ($kg -le 0) { continue }
        $porDia[$dia].Total = $porDia[$dia].Total + $kg

        $destino = 'cortes'
        if ([int]$f["EsSub"] -eq 1) {
            $destino = 'subprod'
        } else {
            $cod = [string][int]$f["CodigoProducto"]
            $zid = $Cfg[$cod]
            if ($zid -and ($IdsPaneles -contains $zid)) { $destino = $zid }
            elseif ($zid -eq 'excluido' -or $zid -eq 'sinasignar') { $destino = 'cortes' }
        }
        $porDia[$dia][$destino] = $porDia[$dia][$destino] + $kg
    }

    $lista = New-Object System.Collections.Generic.List[object]
    foreach ($k in $porDia.Keys) {
        $d = $porDia[$k]
        if ($d.Total -le 0) { continue }
        $o = [pscustomobject]@{ Dia = $d.Dia; Total = $d.Total; Cortes = $d.cortes }
        foreach ($p in $PanelesCuadre) {
            Add-Member -InputObject $o -MemberType NoteProperty -Name $p.Id -Value $d[$p.Id]
        }
        [void]$lista.Add($o)
    }
    return ,$lista.ToArray()
}

function Get-AggZonas {
    param([array]$Materiales, [hashtable]$Cfg)
    $agg = @{}
    foreach ($z in $ZonasCuadre) {
        $agg[$z.Id] = [pscustomobject]@{
            Zona = $z; Piezas = 0
            Mats = New-Object System.Collections.Generic.List[object]
        }
    }
    # Los paneles acumulan KILOS, no piezas: son productos a granel
    $pan = @{}
    foreach ($p in $PanelesCuadre) {
        $pan[$p.Id] = [pscustomobject]@{
            Panel = $p; Kg = 0.0
            Mats = New-Object System.Collections.Generic.List[object]
        }
    }
    $sin  = New-Object System.Collections.Generic.List[object]
    $excl = 0
    $kgTotal = 0.0
    $kgZonas = 0.0
    foreach ($m in $Materiales) {
        $kgTotal = $kgTotal + $m.Kg
        # Todo lo etiquetado en la estacion de subproductos es subproducto,
        # sin importar el nombre ni lo que diga la configuracion.
        if ($m.EsSub -eq 1) {
            $pan['subprod'].Kg = $pan['subprod'].Kg + $m.Kg
            $pan['subprod'].Mats.Add($m)
            continue
        }
        $zid = $Cfg[[string]$m.Codigo]
        if ($zid -eq 'excluido') { $excl = $excl + $m.Piezas; continue }
        if ($zid -and $pan.ContainsKey($zid)) {
            $pan[$zid].Kg = $pan[$zid].Kg + $m.Kg
            $pan[$zid].Mats.Add($m)
            continue
        }
        if ($zid -and $agg.ContainsKey($zid)) {
            $kgZonas = $kgZonas + $m.Kg
            $agg[$zid].Piezas = $agg[$zid].Piezas + $m.Piezas
            $agg[$zid].Mats.Add($m)
        } else {
            $sin.Add($m)
        }
    }
    $tot = 0
    foreach ($z in $ZonasCuadre) { $tot = $tot + $agg[$z.Id].Piezas }
    $sinArr = $sin.ToArray()
    return [pscustomobject]@{ Agg = $agg; Sin = $sinArr; PiezasExcluidas = $excl; TotalAsignadas = $tot
                              Pan = $pan; KgTotal = $kgTotal; KgZonas = $kgZonas }
}

function Compute-Tendencia {
    param([hashtable]$Raw, [hashtable]$CfgZ, [hashtable]$CerdosDia, [datetime]$Fecha, [string]$ZonaSel = '')

    $multSel = $SumaMultiplicadores
    if ($ZonaSel) {
        foreach ($z in $ZonasCuadre) { if ($z.Id -eq $ZonaSel) { $multSel = $z.Mult } }
    }
    $lista = @()
    foreach ($d in @($Raw.Keys | Sort-Object)) {
        $codigos = $Raw[$d]
        $p = 0
        foreach ($k in $codigos.Keys) {
            $zid = $CfgZ[[string]$k]
            if (-not $zid -or $zid -eq 'excluido') { continue }
            if (-not $ZonaSel -or $zid -eq $ZonaSel) { $p = $p + $codigos[$k] }
        }
        if ($p -le 0) { continue }
        $cerd = $CerdosPorDefecto
        if ($CerdosDia.ContainsKey($d)) { $cerd = [int]$CerdosDia[$d] }
        $esp = $cerd * $multSel
        $pct = 0.0
        if ($esp -gt 0) { $pct = $p / $esp * 100 }
        $lista += [pscustomobject]@{ Dia = $d; Pct = $pct }
    }
    $mesKey = $Fecha.ToString('yyyy-MM')
    $delMes = @($lista | Where-Object { $_.Dia.StartsWith($mesKey) })
    $ult    = @($lista | Select-Object -Last 14)
    $prom14 = 0.0
    if ($ult.Count -gt 0)    { $prom14  = ($ult    | Measure-Object -Property Pct -Average).Average }
    $promMes = 0.0
    if ($delMes.Count -gt 0) { $promMes = ($delMes | Measure-Object -Property Pct -Average).Average }
    return [pscustomobject]@{ Barras = $ult; Prom14 = $prom14; PromMes = $promMes; MesTexto = $Fecha.ToString('MM-yyyy') }
}

function Get-ColorPct {
    param([double]$P)
    if ($P -gt 105) { return $C.NeonVioleta }
    if ($P -ge 98)  { return $C.NeonVerde }
    if ($P -ge 85)  { return $C.NeonNaranja }
    return $C.NeonRojo
}

function Get-EstadoPct {
    param([double]$P)
    if ($P -gt 105) { return 'Sobre' }
    if ($P -ge 98)  { return 'OK' }
    if ($P -ge 85)  { return 'Parcial' }
    return 'Bajo'
}

function Build-CuadreCsv {
    param([array]$Materiales, [hashtable]$Cfg, [int]$Cerdos, [datetime]$Fecha)
    $r = Get-AggZonas -Materiales $Materiales -Cfg $Cfg
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Fecha;Cerdos;Zona;Codigo;Material;Piezas;Cajas;EsperadoZona;PctZona")
    $fk = $Fecha.ToString('yyyy-MM-dd')
    $cult = [System.Globalization.CultureInfo]::GetCultureInfo("es-CL")
    foreach ($z in $ZonasCuadre) {
        $dz  = $r.Agg[$z.Id]
        $esp = $Cerdos * $z.Mult
        $pct = 0.0
        if ($esp -gt 0) { $pct = $dz.Piezas / $esp * 100 }
        $pctTxt = $pct.ToString("N1", $cult)
        foreach ($m in $dz.Mats) {
            [void]$sb.AppendLine("$fk;$Cerdos;$($z.Nombre);$($m.Codigo);$($m.Producto);$($m.Piezas);$($m.Cajas);$esp;$pctTxt")
        }
        if ($dz.Mats.Count -eq 0) {
            [void]$sb.AppendLine("$fk;$Cerdos;$($z.Nombre);;;0;0;$esp;$pctTxt")
        }
    }
    foreach ($m in $r.Sin) {
        [void]$sb.AppendLine("$fk;$Cerdos;SIN ASIGNAR;$($m.Codigo);$($m.Producto);$($m.Piezas);$($m.Cajas);;")
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Zona;Piezas;Esperado;Pct")
    foreach ($z in $ZonasCuadre) {
        $dz  = $r.Agg[$z.Id]
        $esp = $Cerdos * $z.Mult
        $pct = 0.0
        if ($esp -gt 0) { $pct = $dz.Piezas / $esp * 100 }
        [void]$sb.AppendLine("$($z.Nombre);$($dz.Piezas);$esp;$($pct.ToString('N1', $cult))")
    }
    return $sb.ToString()
}


# ------------------------------------------------------------
#  BUSCAR PRODUCTO
#  Consulta libre por rango de fechas y multiples materiales.
#
#  La tabla Cajas viene del export del ERP y no todas las
#  instalaciones traen las mismas columnas, asi que en vez de
#  fijar los nombres a ciegas se lee el esquema real una vez y
#  se resuelve cada campo logico contra lo que exista.
# ------------------------------------------------------------
$TopeDetalle    = 5000    # cajas maximas a mostrar en la grilla de detalle
$script:ColsBase = $null  # cache del esquema de Cajas
$script:MapaCols = $null  # campo logico -> nombre real de columna
$script:BusDet   = @()    # ultimo detalle consultado (para exportar)
$script:BusSel   = $null  # producto cuyas estadisticas se muestran en las tarjetas
$script:BusRes   = @()    # ultimo resumen consultado (para exportar)

function Get-ColumnasCajas {
    if ($null -ne $script:ColsBase) { return $script:ColsBase }
    $cols = New-Object System.Collections.Generic.List[string]
    try {
        # 1=0 no trae filas pero si la definicion de columnas
        $t = Invoke-ConsultaAccess -Sql "SELECT * FROM Cajas WHERE 1=0"
        foreach ($c in $t.Columns) { [void]$cols.Add([string]$c.ColumnName) }
    } catch { }
    $script:ColsBase = $cols
    return $cols
}

function Resolve-Columna {
    # Devuelve el nombre real de la primera variante que exista, o $null.
    param([string[]]$Variantes)
    $cols = Get-ColumnasCajas
    foreach ($v in $Variantes) {
        foreach ($c in $cols) {
            if ($c -eq $v) { return $c }
        }
    }
    # segundo intento: sin distinguir mayusculas ni espacios
    foreach ($v in $Variantes) {
        $vn = ($v -replace '\s', '').ToLower()
        foreach ($c in $cols) {
            if (($c -replace '\s', '').ToLower() -eq $vn) { return $c }
        }
    }
    return $null
}

function Initialize-MapaColumnas {
    if ($null -ne $script:MapaCols) { return $script:MapaCols }
    $m = @{}
    $m['NumeroCaja'] = Resolve-Columna @('NumeroCaja', 'Numero caja', 'NumeroDeCaja', 'NroCaja')
    # El lote correcto es "Lote Unico". El Numero SAG NO sirve como lote:
    # es el correlativo de la etiqueta, distinto en cada caja.
    $m['Lote']       = Resolve-Columna @('LoteUnico', 'Lote Unico', 'Lote_Unico', 'Lote', 'NumeroLote', 'Nro Lote')
    $m['FDesposte']  = Resolve-Columna @('FechaDesposte', 'Fecha Desposte', 'Desposte', 'FechaDeDesposte')
    $m['Piezas']     = Resolve-Columna @('Piezas', 'NumeroPiezas', 'CantidadPiezas')
    $script:MapaCols = $m
    return $m
}

function Get-CatalogoProductos {
    # Catalogo de materiales vistos en el rango. Se muestra
    # "codigo - nombre" para poder buscar por cualquiera de los dos.
    param([datetime]$Desde, [datetime]$Hasta)

    $fDesde = Format-FechaAccess $Desde.Date
    $fHasta = Format-FechaAccess $Hasta.Date.AddDays(1)
    $sql = @"
SELECT CodigoProducto, ProductoEspanol, COUNT(*) AS Cajas
FROM Cajas
WHERE FechaPesaje >= $fDesde AND FechaPesaje < $fHasta
  AND Usuario <> '$UsuarioExcluido'
  AND NombreCliente = '$Cliente'
GROUP BY CodigoProducto, ProductoEspanol
ORDER BY ProductoEspanol
"@
    $t = Invoke-ConsultaAccess -Sql $sql
    $lista = New-Object System.Collections.Generic.List[object]
    foreach ($f in $t.Rows) {
        $cod = 0
        if ($f["CodigoProducto"] -isnot [System.DBNull]) { $cod = [int]$f["CodigoProducto"] }
        $nom = [string]$f["ProductoEspanol"]
        if ([string]::IsNullOrWhiteSpace($nom)) { continue }
        [void]$lista.Add([pscustomobject]@{
            Codigo   = $cod
            Producto = $nom.Trim()
            Cajas    = [int]$f["Cajas"]
            Texto    = ("{0,-7} {1}  ({2} cajas)" -f $cod, $nom.Trim(), [int]$f["Cajas"])
        })
    }
    return ,$lista.ToArray()
}

function Get-DetalleCajas {
    # Detalle caja por caja de los materiales elegidos.
    param([datetime]$Desde, [datetime]$Hasta, [object[]]$Productos)

    if (-not $Productos -or $Productos.Count -eq 0) { return ,@() }
    $m = Initialize-MapaColumnas
    $fDesde = Format-FechaAccess $Desde.Date
    $fHasta = Format-FechaAccess $Hasta.Date.AddDays(1)

    # Se filtra por codigo cuando existe; los materiales sin codigo
    # (CodigoProducto 0 o vacio) se piden por nombre exacto.
    $cods = @(); $noms = @()
    foreach ($p in $Productos) {
        if ($p.Codigo -gt 0) { $cods += [string]$p.Codigo }
        else { $noms += "'" + ($p.Producto -replace "'", "''") + "'" }
    }
    $filtros = @()
    if ($cods.Count -gt 0) { $filtros += "CodigoProducto IN (" + ($cods -join ",") + ")" }
    if ($noms.Count -gt 0) { $filtros += "ProductoEspanol IN (" + ($noms -join ",") + ")" }
    $filtroProd = "(" + ($filtros -join " OR ") + ")"

    # Solo se piden las columnas que existen de verdad en la base
    $sel = @("CodigoProducto", "ProductoEspanol", "PesoNeto", "FechaPesaje")
    foreach ($k in @('NumeroCaja', 'Piezas', 'Lote', 'FDesposte')) {
        if ($m[$k]) { $sel += "[" + $m[$k] + "] AS " + $k }
    }
    $orden = if ($m['NumeroCaja']) { "FechaPesaje, [" + $m['NumeroCaja'] + "]" } else { "FechaPesaje" }

    $sql = @"
SELECT TOP $TopeDetalle $($sel -join ", ")
FROM Cajas
WHERE FechaPesaje >= $fDesde AND FechaPesaje < $fHasta
  AND Usuario <> '$UsuarioExcluido'
  AND NombreCliente = '$Cliente'
  AND $filtroProd
ORDER BY $orden
"@
    $t = Invoke-ConsultaAccess -Sql $sql

    $filas = New-Object System.Collections.Generic.List[object]
    foreach ($f in $t.Rows) {
        $peso = 0.0
        if ($f["PesoNeto"] -isnot [System.DBNull]) { $peso = [double]$f["PesoNeto"] }
        $cod = 0
        if ($f["CodigoProducto"] -isnot [System.DBNull]) { $cod = [int]$f["CodigoProducto"] }

        $pz = 0
        if ($t.Columns.Contains('Piezas') -and $f["Piezas"] -isnot [System.DBNull]) { $pz = [int]$f["Piezas"] }
        $ncaja = ""
        if ($t.Columns.Contains('NumeroCaja') -and $f["NumeroCaja"] -isnot [System.DBNull]) { $ncaja = [string]$f["NumeroCaja"] }
        $lote = ""
        if ($t.Columns.Contains('Lote') -and $f["Lote"] -isnot [System.DBNull]) { $lote = [string]$f["Lote"] }
        $fdesp = ""
        if ($t.Columns.Contains('FDesposte') -and $f["FDesposte"] -isnot [System.DBNull]) {
            try { $fdesp = ([datetime]$f["FDesposte"]).ToString("dd-MM-yyyy") } catch { $fdesp = [string]$f["FDesposte"] }
        }
        $fpes = ""
        if ($f["FechaPesaje"] -isnot [System.DBNull]) { $fpes = ([datetime]$f["FechaPesaje"]).ToString("dd-MM-yyyy") }

        [void]$filas.Add([pscustomobject]@{
            NumeroCaja = $ncaja
            Codigo     = $cod
            Producto   = ([string]$f["ProductoEspanol"]).Trim()
            PesoNetoN  = $peso
            PesoNeto   = Fmt $peso 2
            Piezas     = $pz
            Lote       = $lote
            FDesposte  = $fdesp
            FPesaje    = $fpes
        })
    }
    return ,$filas.ToArray()
}

function Get-Estadisticas {
    # Promedio, min, max y desviacion estandar muestral.
    param([double[]]$Valores)
    $n = $Valores.Count
    if ($n -eq 0) {
        return [pscustomobject]@{ N=0; Suma=0.0; Prom=0.0; Min=0.0; Max=0.0; Desv=0.0 }
    }
    $suma = 0.0
    foreach ($v in $Valores) { $suma += $v }
    $prom = $suma / $n
    $mn = $Valores[0]; $mx = $Valores[0]
    foreach ($v in $Valores) {
        if ($v -lt $mn) { $mn = $v }
        if ($v -gt $mx) { $mx = $v }
    }
    $desv = 0.0
    if ($n -gt 1) {
        $acum = 0.0
        foreach ($v in $Valores) { $acum += [math]::Pow($v - $prom, 2) }
        $desv = [math]::Sqrt($acum / ($n - 1))
    }
    return [pscustomobject]@{ N=$n; Suma=$suma; Prom=$prom; Min=$mn; Max=$mx; Desv=$desv }
}

function Get-ResumenBusqueda {
    # Agrupa el detalle por material y calcula estadistica de cada uno.
    param([object[]]$Detalle)
    if (-not $Detalle -or $Detalle.Count -eq 0) { return ,@() }

    $porClave = @{}
    foreach ($d in $Detalle) {
        $clave = "{0}|{1}" -f $d.Codigo, $d.Producto
        if (-not $porClave.ContainsKey($clave)) {
            $porClave[$clave] = [pscustomobject]@{
                Codigo   = $d.Codigo
                Producto = $d.Producto
                Pesos    = (New-Object System.Collections.Generic.List[double])
                Piezas   = 0
                KgConPz  = 0.0   # kilos de las cajas que si traen piezas
                PzTot    = 0     # piezas de esas mismas cajas
            }
        }
        [void]$porClave[$clave].Pesos.Add($d.PesoNetoN)
        $porClave[$clave].Piezas = $porClave[$clave].Piezas + $d.Piezas
        # El peso por pieza se promedia ponderado: kilos totales / piezas
        # totales. Las cajas sin piezas declaradas se dejan fuera de ambos
        # acumuladores para no ensuciar el promedio ni dividir por cero.
        if ($d.Piezas -gt 0) {
            $porClave[$clave].KgConPz = $porClave[$clave].KgConPz + $d.PesoNetoN
            $porClave[$clave].PzTot   = $porClave[$clave].PzTot + $d.Piezas
        }
    }

    $filas = New-Object System.Collections.Generic.List[object]
    foreach ($k in ($porClave.Keys | Sort-Object)) {
        $g = $porClave[$k]
        $e = Get-Estadisticas $g.Pesos.ToArray()
        $kgPz = 0.0
        if ($g.PzTot -gt 0) { $kgPz = $g.KgConPz / $g.PzTot }
        [void]$filas.Add([pscustomobject]@{
            Codigo   = $g.Codigo
            Producto = $g.Producto
            Cajas    = $e.N
            KgTotalN = $e.Suma
            KgTotal  = Fmt $e.Suma 1
            Promedio = Fmt $e.Prom 3
            Minimo   = Fmt $e.Min 2
            Maximo   = Fmt $e.Max 2
            Desv     = Fmt $e.Desv 3
            Piezas   = $g.Piezas
            KgPorPzN = $kgPz
            KgPorPz  = $(if ($g.PzTot -gt 0) { Fmt $kgPz 3 } else { "-" })
        })
    }
    # @(...) es obligatorio: con UN solo material Sort-Object devuelve el
    # objeto suelto, "return ,$x" lo desenrolla y el DataGrid recibe un
    # PSCustomObject en vez de una lista -> excepcion al asignar ItemsSource.
    $orden = @($filas.ToArray() | Sort-Object -Property KgTotalN -Descending)
    return ,$orden
}

function Build-BusquedaCsv {
    # Dos secciones en un mismo archivo: resumen y detalle.
    param([object[]]$Resumen, [object[]]$Detalle, [datetime]$Desde, [datetime]$Hasta)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("BUSCAR PRODUCTO - COMAFRI")
    [void]$sb.AppendLine("Cliente;$Cliente")
    [void]$sb.AppendLine("Desde;" + $Desde.ToString("dd-MM-yyyy") + ";Hasta;" + $Hasta.ToString("dd-MM-yyyy"))
    [void]$sb.AppendLine("Generado;" + (Get-Date).ToString("dd-MM-yyyy HH:mm"))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("RESUMEN POR PRODUCTO")
    [void]$sb.AppendLine("Material;Producto;Cajas;Kg total;Promedio;Minimo;Maximo;Desviacion;Piezas;Kg por pieza")
    foreach ($r in $Resumen) {
        [void]$sb.AppendLine(($r.Codigo, $r.Producto, $r.Cajas, $r.KgTotal, $r.Promedio,
                              $r.Minimo, $r.Maximo, $r.Desv, $r.Piezas, $r.KgPorPz) -join ";")
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("DETALLE DE CAJAS")
    [void]$sb.AppendLine("Numero caja;Material;Producto;Peso neto;Piezas;Lote;Fecha desposte;Fecha pesaje")
    foreach ($d in $Detalle) {
        [void]$sb.AppendLine(($d.NumeroCaja, $d.Codigo, $d.Producto, $d.PesoNeto, $d.Piezas,
                              $d.Lote, $d.FDesposte, $d.FPesaje) -join ";")
    }
    return $sb.ToString()
}


# ------------------------------------------------------------
#  GIVEAWAY EN DINERO
# ------------------------------------------------------------
function Read-TiposCambio {
    $tc = @{}
    foreach ($k in $TiposCambioBase.Keys) { $tc[$k] = [double]$TiposCambioBase[$k] }
    if (Test-Path $RutaTiposCambio) {
        try {
            $j = Get-Content $RutaTiposCambio -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) {
                $v = 0.0
                if ([double]::TryParse([string]$p.Value, [ref]$v) -and $v -gt 0) { $tc[$p.Name] = $v }
            }
        } catch { }
    }
    return $tc
}

function Write-TiposCambio {
    param([hashtable]$Tipos)
    $o = New-Object psobject
    foreach ($k in ($Tipos.Keys | Sort-Object)) {
        Add-Member -InputObject $o -MemberType NoteProperty -Name $k -Value ([double]$Tipos[$k])
    }
    $o | ConvertTo-Json | Set-Content -Path $RutaTiposCambio -Encoding UTF8
}

function Set-TiposCambio {
    # Puente para actualizar el valor desde un closure: dentro de
    # GetNewClosure "$script:" no llega al script principal.
    param([hashtable]$Tipos)
    $script:TiposCambio = $Tipos
}

function Get-PesosRegalados {
    # Kilos regalados de un producto -> pesos chilenos.
    # Devuelve 0 si el producto no tiene precio cargado.
    param([int]$Codigo, [double]$KgRegalados, [hashtable]$Tipos)
    if (-not $PreciosGiveaway.ContainsKey($Codigo)) { return 0.0 }
    $p = $PreciosGiveaway[$Codigo]
    $tasa = 0.0
    if ($Tipos.ContainsKey($p.Moneda)) { $tasa = [double]$Tipos[$p.Moneda] }
    if ($tasa -le 0) { return 0.0 }
    return $KgRegalados * [double]$p.Precio * $tasa
}

function Get-DetallePrecio {
    # Texto corto para mostrar de donde sale el valor (ej. "4,30 EUR/kg").
    param([int]$Codigo)
    if (-not $PreciosGiveaway.ContainsKey($Codigo)) { return "sin precio" }
    $p = $PreciosGiveaway[$Codigo]
    return (Fmt ([double]$p.Precio) 2) + " " + $p.Moneda + "/kg"
}

# ------------------------------------------------------------
#  INTERFAZ
# ------------------------------------------------------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Produccion - Comafri" Height="900" Width="1480"
        WindowStartupLocation="CenterScreen" Background="#070B18"
        FontFamily="Segoe UI" WindowState="Maximized">

  <Window.Resources>
    <SolidColorBrush x:Key="Panel"  Color="#0E1530"/>
    <SolidColorBrush x:Key="Campo"  Color="#0A1128"/>
    <SolidColorBrush x:Key="Linea"  Color="#22305C"/>
    <SolidColorBrush x:Key="Txt"    Color="#E8EDFA"/>
    <SolidColorBrush x:Key="Txt2"   Color="#8C9BC4"/>

    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Txt}"/>
    </Style>

    <Style x:Key="Etiqueta" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Txt2}"/>
      <Setter Property="FontSize" Value="10.5"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,7,0"/>
    </Style>

    <Style x:Key="Tarjeta" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource Panel}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Linea}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="15,13"/>
      <Setter Property="Margin" Value="0,0,11,0"/>
    </Style>

    <Style x:Key="Caja" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource Panel}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Linea}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="16,14"/>
    </Style>

    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="#DDE5F5"/>
      <Setter Property="Foreground" Value="#0A1128"/>
      <Setter Property="BorderBrush" Value="{StaticResource Linea}"/>
      <Setter Property="Height" Value="27"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="DatePicker">
      <Setter Property="Height" Value="27"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Width" Value="118"/>
    </Style>

    <Style x:Key="BotonRango" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource Campo}"/>
      <Setter Property="Foreground" Value="{StaticResource Txt2}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Linea}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,4"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="14" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#3D8BFF"/>
                <Setter Property="Foreground" Value="#E8EDFA"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="TituloCaja" TargetType="TextBlock">
      <Setter Property="FontSize" Value="13.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="NotaCaja" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Foreground" Value="{StaticResource Txt2}"/>
      <Setter Property="Margin" Value="0,2,0,10"/>
    </Style>

    <Style x:Key="GridHeader" TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#0A1128"/>
      <Setter Property="Foreground" Value="#8C9BC4"/>
      <Setter Property="FontSize" Value="10.5"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="BorderBrush" Value="#22305C"/>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
    </Style>

    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="#8C9BC4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" BorderBrush="Transparent"
                    BorderThickness="0,0,0,2" Padding="16,9,16,7" Margin="0,0,4,0" Cursor="Hand">
              <TextBlock Text="{TemplateBinding Header}" FontSize="13.5" FontWeight="SemiBold"
                         Foreground="{TemplateBinding Foreground}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#F7941D"/>
                <Setter Property="Foreground" Value="#E8EDFA"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#E8EDFA"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- ENCABEZADO -->
    <Border Grid.Row="0" Margin="18,16,18,0" CornerRadius="0,10,10,0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
          <GradientStop Color="#1B3F94" Offset="0"/>
          <GradientStop Color="#00070B18" Offset="0.9"/>
        </LinearGradientBrush>
      </Border.Background>
      <Grid>
        <Border Width="4" HorizontalAlignment="Left" Background="#F7941D"/>
        <Grid Margin="20,14">
          <StackPanel HorizontalAlignment="Left">
            <TextBlock Text="PRODUCCION" FontSize="23" FontWeight="SemiBold"/>
            <TextBlock x:Name="txtSubtitulo" Text="Fase 1 Produccion - Cliente AASA" FontSize="12" Foreground="#C6D4F5" Margin="0,3,0,0"/>
          </StackPanel>
          <StackPanel HorizontalAlignment="Right" VerticalAlignment="Center">
            <TextBlock Text="COMAFRI - AASA" FontSize="11" Foreground="#9FB2DE" HorizontalAlignment="Right"/>
            <TextBlock x:Name="txtActualizado" Text="" FontSize="11" Foreground="#9FB2DE" HorizontalAlignment="Right" Margin="0,3,0,0"/>
          </StackPanel>
        </Grid>
      </Grid>
    </Border>

    <!-- PESTANAS -->
    <TabControl Grid.Row="1" Margin="18,8,18,0">
      <TabItem Header="Giveaway">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
    <!-- FILTROS -->
    <Border Grid.Row="0" Style="{StaticResource Caja}" Margin="0,13,0,0" Padding="16,11">
      <WrapPanel>
        <TextBlock Text="DESDE" Style="{StaticResource Etiqueta}"/>
        <DatePicker x:Name="dpDesde" Margin="0,0,12,0"/>
        <TextBlock Text="HASTA" Style="{StaticResource Etiqueta}"/>
        <DatePicker x:Name="dpHasta" Margin="0,0,16,0"/>
        <Button x:Name="btnHoy"  Content="Hoy"     Style="{StaticResource BotonRango}"/>
        <Button x:Name="btn7"    Content="7 dias"  Style="{StaticResource BotonRango}"/>
        <Button x:Name="btn30"   Content="30 dias" Style="{StaticResource BotonRango}"/>
        <Button x:Name="btn90"   Content="90 dias" Style="{StaticResource BotonRango}"/>
        <TextBlock Text="CLIENTE" Style="{StaticResource Etiqueta}" Margin="14,0,7,0"/>
        <Border Background="#0A1128" BorderBrush="#22305C" BorderThickness="1" CornerRadius="13"
                Padding="12,4" Margin="0,0,14,0" VerticalAlignment="Center">
          <TextBlock Text="AASA PORK LIMITADA" FontSize="12" Foreground="#3D8BFF"/>
        </Border>
        <TextBlock Text="PRODUCTO" Style="{StaticResource Etiqueta}"/>
        <ComboBox x:Name="cbProducto" Width="320" Margin="0,0,14,0"/>
        <CheckBox x:Name="chkAuto" Content="Actualizacion automatica" IsChecked="True"
                  Foreground="#8C9BC4" FontSize="12" VerticalAlignment="Center" Margin="0,0,12,0"/>
        <Button x:Name="btnTiposCambio" Content="Tipos de cambio" Style="{StaticResource BotonRango}"/>
        <Button x:Name="btnRefrescar" Content="Actualizar ahora" Style="{StaticResource BotonRango}"/>
      </WrapPanel>
    </Border>

    <!-- CONTENIDO -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,13,0,0">
      <StackPanel>

        <Grid x:Name="gridTarjetas" Margin="0,0,0,13">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <Border Grid.Column="0" Style="{StaticResource Tarjeta}">
            <Grid>
              <Border Width="3" HorizontalAlignment="Left" Margin="-15,-13,0,-13" Background="#3D8BFF"/>
              <StackPanel Margin="6,0,0,0">
                <TextBlock Text="GIVEAWAY PROMEDIO" FontSize="10.5" Foreground="#8C9BC4"/>
                <TextBlock x:Name="vGw" Text="-" FontSize="29" FontWeight="SemiBold" Foreground="#00E5FF" Margin="0,7,0,0"/>
                <TextBlock x:Name="sGw" Text="peso real vs etiqueta" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Grid.Column="1" Style="{StaticResource Tarjeta}">
            <Grid>
              <Border Width="3" HorizontalAlignment="Left" Margin="-15,-13,0,-13" Background="#FFA829"/>
              <StackPanel Margin="6,0,0,0">
                <TextBlock Text="KILOS REGALADOS" FontSize="10.5" Foreground="#8C9BC4"/>
                <TextBlock x:Name="vKg" Text="-" FontSize="29" FontWeight="SemiBold" Foreground="#FFA829" Margin="0,7,0,0"/>
                <TextBlock x:Name="sKg" Text="" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Grid.Column="2" Style="{StaticResource Tarjeta}">
            <Grid>
              <Border Width="3" HorizontalAlignment="Left" Margin="-15,-13,0,-13" Background="#00FFA3"/>
              <StackPanel Margin="6,0,0,0">
                <TextBlock Text="CAJAS EVALUADAS" FontSize="10.5" Foreground="#8C9BC4"/>
                <TextBlock x:Name="vCajas" Text="-" FontSize="29" FontWeight="SemiBold" Foreground="#00FFA3" Margin="0,7,0,0"/>
                <TextBlock x:Name="sCajas" Text="solo productos de peso fijo" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Grid.Column="3" Style="{StaticResource Tarjeta}">
            <Grid>
              <Border Width="3" HorizontalAlignment="Left" Margin="-15,-13,0,-13" Background="#A96BFF"/>
              <StackPanel Margin="6,0,0,0">
                <TextBlock Text="GIVEAWAY POR CAJA" FontSize="10.5" Foreground="#8C9BC4"/>
                <TextBlock x:Name="vProm" Text="-" FontSize="29" FontWeight="SemiBold" Foreground="#A96BFF" Margin="0,7,0,0"/>
                <TextBlock x:Name="sProm" Text="kilos regalados en cada caja" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Grid.Column="4" Style="{StaticResource Tarjeta}">
            <Grid>
              <Border Width="3" HorizontalAlignment="Left" Margin="-15,-13,0,-13" Background="#FF3B6B"/>
              <StackPanel Margin="6,0,0,0">
                <TextBlock Text="DINERO REGALADO" FontSize="10.5" Foreground="#8C9BC4"/>
                <TextBlock x:Name="vPlata" Text="-" FontSize="29" FontWeight="SemiBold" Foreground="#FF3B6B" Margin="0,7,0,0"/>
                <TextBlock x:Name="sPlata" Text="valor de venta de lo regalado" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Grid.Column="5" Style="{StaticResource Tarjeta}" Margin="0">
            <Grid>
              <Border x:Name="barraBajo" Width="3" HorizontalAlignment="Left" Margin="-15,-13,0,-13" Background="#00FFA3"/>
              <StackPanel Margin="6,0,0,0">
                <TextBlock Text="CAJAS BAJO PESO" FontSize="10.5" Foreground="#8C9BC4"/>
                <TextBlock x:Name="vBajo" Text="-" FontSize="29" FontWeight="SemiBold" Foreground="#00FFA3" Margin="0,7,0,0"/>
                <TextBlock x:Name="sBajo" Text="riesgo normativo" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
              </StackPanel>
            </Grid>
          </Border>
        </Grid>

        <Border x:Name="panelAviso" Background="#22FF3B6B" BorderBrush="#59FF3B6B" BorderThickness="1"
                CornerRadius="9" Padding="14,11" Margin="0,0,0,13" Visibility="Collapsed">
          <TextBlock x:Name="txtAviso" Foreground="#FFC2D2" FontSize="12.5" TextWrapping="Wrap"/>
        </Border>

        <Grid Margin="0,0,0,13">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Border Grid.Column="0" Style="{StaticResource Caja}" Margin="0,0,6,0">
            <StackPanel>
              <TextBlock Text="Giveaway por producto" Style="{StaticResource TituloCaja}"/>
              <TextBlock Text="Peso real entregado por sobre el peso declarado en la etiqueta" Style="{StaticResource NotaCaja}"/>
              <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Left">
                <Canvas x:Name="cvBarras" Width="720" Height="260" Background="Transparent"/>
              </Viewbox>
            </StackPanel>
          </Border>
          <Border Grid.Column="1" Style="{StaticResource Caja}" Margin="6,0,0,0">
            <StackPanel>
              <TextBlock Text="Kilos regalados por mes" Style="{StaticResource TituloCaja}"/>
              <TextBlock Text="Evolucion del sobrepeso entregado" Style="{StaticResource NotaCaja}"/>
              <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Left">
                <Canvas x:Name="cvLinea" Width="720" Height="260" Background="Transparent"/>
              </Viewbox>
            </StackPanel>
          </Border>
        </Grid>

        <Border Style="{StaticResource Caja}" Margin="0,0,0,13">
          <StackPanel>
            <TextBlock Text="Distribucion del giveaway por caja" Style="{StaticResource TituloCaja}"/>
            <TextBlock x:Name="txtNotaHist" Text="" Style="{StaticResource NotaCaja}"/>
            <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Left">
              <Canvas x:Name="cvHist" Width="1400" Height="320" Background="Transparent"/>
            </Viewbox>
          </StackPanel>
        </Border>

        <Grid Margin="0,0,0,16">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Border Grid.Column="0" Style="{StaticResource Caja}" Margin="0,0,6,0">
            <StackPanel>
              <TextBlock Text="Mapa de calor: giveaway por producto y mes" Style="{StaticResource TituloCaja}"/>
              <TextBlock Text="Mas intenso = mas sobrepeso en proporcion" Style="{StaticResource NotaCaja}"/>
              <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Left">
                <Canvas x:Name="cvMapa" Width="720" Height="260" Background="Transparent"/>
              </Viewbox>
            </StackPanel>
          </Border>
          <Border Grid.Column="1" Style="{StaticResource Caja}" Margin="6,0,0,0">
            <StackPanel>
              <TextBlock Text="Detalle por producto" Style="{StaticResource TituloCaja}"/>
              <TextBlock Text="Ordenado por kilos regalados" Style="{StaticResource NotaCaja}"/>
              <DataGrid x:Name="grid" AutoGenerateColumns="False" IsReadOnly="True" Height="250"
                        Background="Transparent" BorderThickness="0" RowBackground="Transparent"
                        AlternatingRowBackground="#12193A" Foreground="#E8EDFA" GridLinesVisibility="None"
                        HeadersVisibility="Column" FontSize="12" RowHeight="27" CanUserAddRows="False">
                <DataGrid.ColumnHeaderStyle>
                  <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#0A1128"/>
                    <Setter Property="Foreground" Value="#8C9BC4"/>
                    <Setter Property="FontSize" Value="10.5"/>
                    <Setter Property="Padding" Value="8,6"/>
                    <Setter Property="BorderBrush" Value="#22305C"/>
                    <Setter Property="BorderThickness" Value="0,0,0,1"/>
                  </Style>
                </DataGrid.ColumnHeaderStyle>
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Producto"   Binding="{Binding Producto}" Width="*"/>
                  <DataGridTextColumn Header="Etiqueta"   Binding="{Binding Etiqueta}" Width="70"/>
                  <DataGridTextColumn Header="Cajas"      Binding="{Binding Cajas}"    Width="62"/>
                  <DataGridTextColumn Header="Peso real"  Binding="{Binding Promedio}" Width="76"/>
                  <DataGridTextColumn Header="Kg / caja"  Binding="{Binding PorCaja}"  Width="76"/>
                  <DataGridTextColumn Header="Giveaway"   Binding="{Binding Giveaway}" Width="78"/>
                  <DataGridTextColumn Header="Kg reg."    Binding="{Binding KgReg}"    Width="74"/>
                  <DataGridTextColumn Header="Precio"     Binding="{Binding Precio}"   Width="96"/>
                  <DataGridTextColumn Header="$ regalado" Binding="{Binding Plata}"    Width="104"/>
                  <DataGridTextColumn Header="Bajo peso"  Binding="{Binding Bajo}"     Width="76"/>
                </DataGrid.Columns>
              </DataGrid>
            </StackPanel>
          </Border>
        </Grid>

      </StackPanel>
    </ScrollViewer>

        </Grid>
      </TabItem>
      <TabItem Header="Cuadre de Piezas">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Style="{StaticResource Caja}" Margin="0,13,0,0" Padding="16,11">
            <WrapPanel>
              <TextBlock Text="FECHA" Style="{StaticResource Etiqueta}"/>
              <DatePicker x:Name="dpFechaCuadre" Margin="0,0,14,0"/>
              <TextBlock Text="CERDOS DESPOSTE" Style="{StaticResource Etiqueta}"/>
              <TextBox x:Name="txtCerdos" Width="66" Height="27" Background="#0A1128" Foreground="#E8EDFA"
                       BorderBrush="#22305C" FontSize="13" TextAlignment="Right"
                       VerticalContentAlignment="Center" Padding="6,0"/>
              <TextBlock x:Name="txtCerdosHint" Text="" FontSize="10.5" Foreground="#4E5D85"
                         VerticalAlignment="Center" Margin="8,0,14,0"/>
              <TextBlock Text="VARA FRIA" Style="{StaticResource Etiqueta}"/>
              <TextBox x:Name="txtVara" Width="60" Height="27" Background="#0A1128" Foreground="#E8EDFA"
                       BorderBrush="#22305C" FontSize="13" TextAlignment="Right"
                       VerticalContentAlignment="Center" Padding="6,0"/>
              <TextBlock Text="kg" FontSize="10.5" Foreground="#4E5D85"
                         VerticalAlignment="Center" Margin="6,0,14,0"/>
              <Button x:Name="btnCfgZonas" Content="Configurar zonas" Style="{StaticResource BotonRango}"/>
              <Button x:Name="btnExportCuadre" Content="Exportar CSV" Style="{StaticResource BotonRango}"/>
              <Button x:Name="btnRefCuadre" Content="Actualizar ahora" Style="{StaticResource BotonRango}"/>
              <StackPanel Orientation="Horizontal" Margin="16,0,0,0" VerticalAlignment="Center">
                <Ellipse Width="9" Height="9" Fill="#00FFA3"/>
                <TextBlock Text="98-105" FontSize="11" Foreground="#8C9BC4" Margin="4,0,10,0"/>
                <Ellipse Width="9" Height="9" Fill="#A96BFF"/>
                <TextBlock Text="sobre 105" FontSize="11" Foreground="#8C9BC4" Margin="4,0,10,0"/>
                <Ellipse Width="9" Height="9" Fill="#FFA829"/>
                <TextBlock Text="85-97" FontSize="11" Foreground="#8C9BC4" Margin="4,0,10,0"/>
                <Ellipse Width="9" Height="9" Fill="#FF3B6B"/>
                <TextBlock Text="bajo 85" FontSize="11" Foreground="#8C9BC4" Margin="4,0,0,0"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,13,0,0">
            <StackPanel>

              <Border x:Name="panelAvisoCfg" Background="#22FFA829" BorderBrush="#59FFA829" BorderThickness="1"
                      CornerRadius="9" Padding="14,11" Margin="0,0,0,13" Visibility="Collapsed">
                <TextBlock x:Name="txtAvisoCfg" Foreground="#FFE1B0" FontSize="12.5" TextWrapping="Wrap"/>
              </Border>

              <Grid Margin="0,0,0,13">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="400"/>
                </Grid.ColumnDefinitions>

                <Border Grid.Column="0" Style="{StaticResource Caja}" Margin="0,0,6,0">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0">
                      <TextBlock Text="Despiece" Style="{StaticResource TituloCaja}"/>
                      <TextBlock x:Name="txtNotaCerdo" Text="Toca una zona para ver su detalle" Style="{StaticResource NotaCaja}"/>
                      <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Left" VerticalAlignment="Top">
                        <Canvas x:Name="cvCerdo" Width="860" Height="470" Background="Transparent"/>
                      </Viewbox>
                    </StackPanel>
                    <Viewbox Grid.Column="1" Stretch="Uniform" StretchDirection="DownOnly"
                             VerticalAlignment="Top" Margin="10,4,0,0">
                      <Canvas x:Name="cvTarjetas" Width="500" Height="586" Background="Transparent"/>
                    </Viewbox>
                  </Grid>
                </Border>

                <Border Grid.Column="1" Style="{StaticResource Caja}" Margin="6,0,0,0">
                  <StackPanel>
                    <TextBlock x:Name="txtZonaTitulo" Text="Cuadre global" Style="{StaticResource TituloCaja}"/>
                    <TextBlock x:Name="txtZonaNota" Text="" Style="{StaticResource NotaCaja}"/>
                    <TextBlock x:Name="txtZonaPct" Text="-" FontSize="34" FontWeight="SemiBold"
                               Foreground="#00E5FF" Margin="0,0,0,10"/>
                    <DataGrid x:Name="gridMat" AutoGenerateColumns="False" IsReadOnly="True" Height="230"
                              Background="Transparent" BorderThickness="0" RowBackground="Transparent"
                              AlternatingRowBackground="#12193A" Foreground="#E8EDFA" GridLinesVisibility="None"
                              HeadersVisibility="Column" FontSize="11.5" RowHeight="25" CanUserAddRows="False"
                              ColumnHeaderStyle="{StaticResource GridHeader}">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Material"  Binding="{Binding Material}" Width="*"/>
                        <DataGridTextColumn Header="Codigo"    Binding="{Binding Codigo}"   Width="62"/>
                        <DataGridTextColumn Header="Piezas"    Binding="{Binding Piezas}"   Width="58"/>
                        <DataGridTextColumn Header="Cajas"     Binding="{Binding Cajas}"    Width="50"/>
                      </DataGrid.Columns>
                    </DataGrid>
                    <TextBlock x:Name="lblTot" Text="" FontSize="12.5" Margin="0,10,0,0"/>
                    <TextBlock x:Name="lblEsp" Text="" FontSize="12.5" Foreground="#8C9BC4" Margin="0,4,0,0"/>
                    <TextBlock x:Name="lblDif" Text="" FontSize="12.5" FontWeight="SemiBold" Margin="0,4,0,0"/>
                  </StackPanel>
                </Border>
              </Grid>

              <!-- APROVECHAMIENTO: caja propia, ocupa el ancho completo -->
              <Border Style="{StaticResource Caja}" Margin="0,0,0,13">
                <StackPanel>
                  <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Stretch">
                    <Canvas x:Name="cvAprov" Width="1860" Height="470" Background="Transparent"/>
                  </Viewbox>
                </StackPanel>
              </Border>

              <Grid Margin="0,0,0,16">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Border Grid.Column="0" Style="{StaticResource Caja}" Margin="0,0,6,0">
                  <StackPanel>
                    <TextBlock Text="Tendencia diaria" Style="{StaticResource TituloCaja}"/>
                    <TextBlock x:Name="txtPromedios" Text="" Style="{StaticResource NotaCaja}"/>
                    <Viewbox Stretch="Uniform" StretchDirection="DownOnly" HorizontalAlignment="Left">
                      <Canvas x:Name="cvTendencia" Width="700" Height="228" Background="Transparent"/>
                    </Viewbox>
                  </StackPanel>
                </Border>

                <Border Grid.Column="1" Style="{StaticResource Caja}" Margin="6,0,0,0">
                  <StackPanel>
                    <TextBlock Text="Resumen por zona" Style="{StaticResource TituloCaja}"/>
                    <TextBlock Text="Toca una fila para verla en el despiece" Style="{StaticResource NotaCaja}"/>
                    <DataGrid x:Name="gridZonas" AutoGenerateColumns="False" IsReadOnly="True" Height="252"
                              Background="Transparent" BorderThickness="0" RowBackground="Transparent"
                              AlternatingRowBackground="#12193A" Foreground="#E8EDFA" GridLinesVisibility="None"
                              HeadersVisibility="Column" FontSize="11.5" RowHeight="24" CanUserAddRows="False"
                              SelectionMode="Single" SelectionUnit="FullRow"
                              ColumnHeaderStyle="{StaticResource GridHeader}">
                      <DataGrid.Columns>
                        <DataGridTextColumn Header="Zona"     Binding="{Binding Zona}"     Width="*"/>
                        <DataGridTextColumn Header="Piezas"   Binding="{Binding Piezas}"   Width="64"/>
                        <DataGridTextColumn Header="Esperado" Binding="{Binding Esperado}" Width="72"/>
                        <DataGridTextColumn Header="%"        Binding="{Binding Pct}"      Width="64"/>
                        <DataGridTextColumn Header="Estado"   Binding="{Binding Estado}"   Width="64"/>
                      </DataGrid.Columns>
                    </DataGrid>
                  </StackPanel>
                </Border>
              </Grid>

            </StackPanel>
          </ScrollViewer>
        </Grid>
      </TabItem>
      <TabItem Header="Buscar Producto">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- FILTROS -->
          <Border Grid.Row="0" Style="{StaticResource Caja}" Margin="0,13,0,0" Padding="16,11">
            <StackPanel>
              <WrapPanel>
                <TextBlock Text="DESDE" Style="{StaticResource Etiqueta}"/>
                <DatePicker x:Name="dpBusDesde" Margin="0,0,12,0"/>
                <TextBlock Text="HASTA" Style="{StaticResource Etiqueta}"/>
                <DatePicker x:Name="dpBusHasta" Margin="0,0,14,0"/>
                <Button x:Name="btnBusHoy"     Content="Hoy"          Style="{StaticResource BotonRango}"/>
                <Button x:Name="btnBusSemana"  Content="7 dias"       Style="{StaticResource BotonRango}"/>
                <Button x:Name="btnBusMes"     Content="30 dias"      Style="{StaticResource BotonRango}"/>
                <Border Width="1" Background="#22305C" Margin="8,2,12,2"/>
                <Button x:Name="btnBusBuscar"  Content="BUSCAR"       Style="{StaticResource BotonRango}"
                        Background="#1B3F94" Foreground="#E8EDFA" FontWeight="SemiBold"/>
                <Button x:Name="btnBusExport"  Content="Exportar CSV" Style="{StaticResource BotonRango}"/>
              </WrapPanel>

              <Grid Margin="0,12,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="330"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Catalogo: se busca por codigo o por nombre -->
                <StackPanel Grid.Column="0" Margin="0,0,14,0">
                  <TextBlock Text="MATERIAL O PRODUCTO" Style="{StaticResource Etiqueta}" Margin="0,0,0,5"/>
                  <TextBox x:Name="txtBusFiltro" Height="27" Background="#0A1128" Foreground="#E8EDFA"
                           BorderBrush="#22305C" FontSize="12" VerticalContentAlignment="Center" Padding="7,0"/>
                  <TextBlock x:Name="txtBusHintFiltro" Text="escribe codigo o nombre para filtrar la lista"
                             FontSize="10.5" Foreground="#4E5D85" Margin="1,4,0,6"/>
                  <ListBox x:Name="lstBusCatalogo" Height="188" Background="#0A1128" BorderBrush="#22305C"
                           Foreground="#E8EDFA" FontSize="11.5" SelectionMode="Extended"
                           ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
                  <WrapPanel Margin="0,7,0,0">
                    <Button x:Name="btnBusAgregar" Content="Agregar seleccion" Style="{StaticResource BotonRango}"/>
                    <Button x:Name="btnBusTodos"   Content="Todos"             Style="{StaticResource BotonRango}"/>
                  </WrapPanel>
                </StackPanel>

                <!-- Productos ya elegidos para la consulta -->
                <StackPanel Grid.Column="1">
                  <TextBlock Text="PRODUCTOS A CONSULTAR" Style="{StaticResource Etiqueta}" Margin="0,0,0,5"/>
                  <ListBox x:Name="lstBusElegidos" Height="215" Background="#0A1128" BorderBrush="#22305C"
                           Foreground="#E8EDFA" FontSize="11.5" SelectionMode="Extended"
                           ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
                  <WrapPanel Margin="0,7,0,0">
                    <Button x:Name="btnBusQuitar"  Content="Quitar seleccion" Style="{StaticResource BotonRango}"/>
                    <Button x:Name="btnBusLimpiar" Content="Limpiar todo"     Style="{StaticResource BotonRango}"/>
                    <TextBlock x:Name="txtBusConteo" Text="0 productos elegidos" FontSize="11"
                               Foreground="#8C9BC4" VerticalAlignment="Center" Margin="8,0,0,0"/>
                  </WrapPanel>
                </StackPanel>
              </Grid>
            </StackPanel>
          </Border>

          <!-- RESULTADOS -->
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,0">
            <StackPanel Margin="0,13,0,16">

              <!-- Tarjetas de estadistica: siempre de UN producto -->
              <Border Style="{StaticResource Caja}" Padding="15,10" Margin="0,0,0,11">
                <WrapPanel>
                  <TextBlock Text="ESTADISTICA DE" Style="{StaticResource Etiqueta}"/>
                  <TextBlock x:Name="txtBusProdSel" Text="elige un producto" FontSize="13.5"
                             FontWeight="SemiBold" Foreground="#00E5FF" VerticalAlignment="Center"
                             Margin="0,0,14,0"/>
                  <TextBlock x:Name="txtBusProdHint" Text="haz clic en una fila del resumen para cambiar"
                             FontSize="10.5" Foreground="#4E5D85" VerticalAlignment="Center"/>
                </WrapPanel>
              </Border>

              <Grid x:Name="gridBusTarjetas" Margin="0,0,0,13">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Border Grid.Column="0" Style="{StaticResource Tarjeta}" BorderBrush="#3D8BFF">
                  <StackPanel>
                    <TextBlock Text="CAJAS" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bCajas" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#3D8BFF"/>
                    <TextBlock x:Name="bCajasSub" Text="en el rango elegido" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>

                <Border Grid.Column="1" Style="{StaticResource Tarjeta}" BorderBrush="#F7941D">
                  <StackPanel>
                    <TextBlock Text="KILOS TOTALES" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bKilos" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#FFA829"/>
                    <TextBlock x:Name="bKilosSub" Text="suma de peso neto" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>

                <Border Grid.Column="2" Style="{StaticResource Tarjeta}" BorderBrush="#00FFA3">
                  <StackPanel>
                    <TextBlock Text="PROMEDIO" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bProm" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#00FFA3"/>
                    <TextBlock x:Name="bPromSub" Text="peso neto por caja" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>

                <Border Grid.Column="3" Style="{StaticResource Tarjeta}" BorderBrush="#00E5FF">
                  <StackPanel>
                    <TextBlock Text="MINIMO" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bMin" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#00E5FF"/>
                    <TextBlock x:Name="bMinSub" Text="caja mas liviana" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>

                <Border Grid.Column="4" Style="{StaticResource Tarjeta}" BorderBrush="#A96BFF">
                  <StackPanel>
                    <TextBlock Text="MAXIMO" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bMax" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#A96BFF"/>
                    <TextBlock x:Name="bMaxSub" Text="caja mas pesada" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>

                <Border Grid.Column="5" Style="{StaticResource Tarjeta}" BorderBrush="#FF3B6B">
                  <StackPanel>
                    <TextBlock Text="DESVIACION" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bDesv" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#FF3B6B"/>
                    <TextBlock x:Name="bDesvSub" Text="dispersion del peso" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>

                <Border Grid.Column="6" Style="{StaticResource Tarjeta}" BorderBrush="#F7941D" Margin="0">
                  <StackPanel>
                    <TextBlock Text="PESO POR PIEZA" Style="{StaticResource Etiqueta}"/>
                    <TextBlock x:Name="bKgPz" Text="-" FontSize="26" FontWeight="SemiBold" Foreground="#F7941D"/>
                    <TextBlock x:Name="bKgPzSub" Text="kilos netos de cada pieza" FontSize="11" Foreground="#8C9BC4" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>
              </Grid>

              <!-- Resumen por producto -->
              <Border Style="{StaticResource Caja}" Margin="0,0,0,13">
                <StackPanel>
                  <TextBlock Text="Resumen por producto" Style="{StaticResource TituloCaja}"/>
                  <TextBlock x:Name="txtBusNotaResumen" Text="una fila por material consultado"
                             Style="{StaticResource NotaCaja}"/>
                  <DataGrid x:Name="gridBusResumen" Height="182" AutoGenerateColumns="False" IsReadOnly="True"
                            Background="#0A1128" RowBackground="#0A1128" AlternatingRowBackground="#0C1330"
                            Foreground="#E8EDFA" BorderBrush="#22305C" GridLinesVisibility="Horizontal"
                            HorizontalGridLinesBrush="#16204A" HeadersVisibility="Column" FontSize="11.5"
                            CanUserAddRows="False" ColumnHeaderStyle="{StaticResource GridHeader}">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Material"  Binding="{Binding Codigo}"   Width="76"/>
                      <DataGridTextColumn Header="Producto"  Binding="{Binding Producto}" Width="*"/>
                      <DataGridTextColumn Header="Cajas"     Binding="{Binding Cajas}"    Width="62"/>
                      <DataGridTextColumn Header="Kg total"  Binding="{Binding KgTotal}"  Width="86"/>
                      <DataGridTextColumn Header="Promedio"  Binding="{Binding Promedio}" Width="78"/>
                      <DataGridTextColumn Header="Minimo"    Binding="{Binding Minimo}"   Width="72"/>
                      <DataGridTextColumn Header="Maximo"    Binding="{Binding Maximo}"   Width="72"/>
                      <DataGridTextColumn Header="Desv."     Binding="{Binding Desv}"     Width="70"/>
                      <DataGridTextColumn Header="Piezas"    Binding="{Binding Piezas}"   Width="68"/>
                      <DataGridTextColumn Header="Kg/pieza"  Binding="{Binding KgPorPz}"  Width="80"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </StackPanel>
              </Border>

              <!-- Detalle caja por caja -->
              <Border Style="{StaticResource Caja}">
                <StackPanel>
                  <TextBlock Text="Detalle de cajas" Style="{StaticResource TituloCaja}"/>
                  <TextBlock x:Name="txtBusNotaDetalle" Text="una fila por caja etiquetada"
                             Style="{StaticResource NotaCaja}"/>
                  <DataGrid x:Name="gridBusDetalle" Height="300" AutoGenerateColumns="False" IsReadOnly="True"
                            Background="#0A1128" RowBackground="#0A1128" AlternatingRowBackground="#0C1330"
                            Foreground="#E8EDFA" BorderBrush="#22305C" GridLinesVisibility="Horizontal"
                            HorizontalGridLinesBrush="#16204A" HeadersVisibility="Column" FontSize="11.5"
                            CanUserAddRows="False" ColumnHeaderStyle="{StaticResource GridHeader}">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Numero caja"  Binding="{Binding NumeroCaja}" Width="118"/>
                      <DataGridTextColumn Header="Material"     Binding="{Binding Codigo}"     Width="76"/>
                      <DataGridTextColumn Header="Producto"     Binding="{Binding Producto}"   Width="*"/>
                      <DataGridTextColumn Header="Peso neto"    Binding="{Binding PesoNeto}"   Width="82"/>
                      <DataGridTextColumn Header="Piezas"       Binding="{Binding Piezas}"     Width="62"/>
                      <DataGridTextColumn Header="Lote"         Binding="{Binding Lote}"       Width="92"/>
                      <DataGridTextColumn Header="F. desposte"  Binding="{Binding FDesposte}"  Width="94"/>
                      <DataGridTextColumn Header="F. pesaje"    Binding="{Binding FPesaje}"    Width="94"/>
                    </DataGrid.Columns>
                  </DataGrid>
                  <TextBlock x:Name="txtBusTope" Text="" FontSize="10.5" Foreground="#FFA829" Margin="1,7,0,0"/>
                </StackPanel>
              </Border>

            </StackPanel>
          </ScrollViewer>
        </Grid>
      </TabItem>
    </TabControl>

    <!-- BARRA DE ESTADO -->
    <Border Grid.Row="2" Background="#0A1128" BorderBrush="#22305C" BorderThickness="0,1,0,0" Padding="20,8" Margin="0,13,0,0">
      <Grid>
        <TextBlock x:Name="txtEstado" Text="Iniciando..." FontSize="11.5" Foreground="#8C9BC4" HorizontalAlignment="Left"/>
        <TextBlock x:Name="txtBase" Text="" FontSize="11" Foreground="#4E5D85" HorizontalAlignment="Right"/>
      </Grid>
    </Border>
  </Grid>
</Window>
"@

$lector  = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$ventana = [Windows.Markup.XamlReader]::Load($lector)

$dpDesde      = $ventana.FindName("dpDesde")
$dpHasta      = $ventana.FindName("dpHasta")
$cbProducto   = $ventana.FindName("cbProducto")
$chkAuto      = $ventana.FindName("chkAuto")
$btnRefrescar = $ventana.FindName("btnRefrescar")
$btnHoy       = $ventana.FindName("btnHoy")
$btn7         = $ventana.FindName("btn7")
$btn30        = $ventana.FindName("btn30")
$btn90        = $ventana.FindName("btn90")
$cvBarras     = $ventana.FindName("cvBarras")
$cvLinea      = $ventana.FindName("cvLinea")
$cvHist       = $ventana.FindName("cvHist")
$cvMapa       = $ventana.FindName("cvMapa")
$grid         = $ventana.FindName("grid")
$txtEstado    = $ventana.FindName("txtEstado")
$txtBase      = $ventana.FindName("txtBase")
$txtActualizado = $ventana.FindName("txtActualizado")
$txtSubtitulo = $ventana.FindName("txtSubtitulo")
$txtNotaHist  = $ventana.FindName("txtNotaHist")
$panelAviso   = $ventana.FindName("panelAviso")
$txtAviso     = $ventana.FindName("txtAviso")
$vGw   = $ventana.FindName("vGw");   $sGw   = $ventana.FindName("sGw")
$vKg   = $ventana.FindName("vKg");   $sKg   = $ventana.FindName("sKg")
$vCajas= $ventana.FindName("vCajas");$sCajas= $ventana.FindName("sCajas")
$vProm = $ventana.FindName("vProm")
$vBajo = $ventana.FindName("vBajo"); $sBajo = $ventana.FindName("sBajo")
$barraBajo = $ventana.FindName("barraBajo")
$vPlata    = $ventana.FindName("vPlata");   $sPlata = $ventana.FindName("sPlata")
$btnTiposCambio = $ventana.FindName("btnTiposCambio")

$dpFechaCuadre   = $ventana.FindName("dpFechaCuadre")
$txtCerdos       = $ventana.FindName("txtCerdos")
$txtVara         = $ventana.FindName("txtVara")
$txtCerdosHint   = $ventana.FindName("txtCerdosHint")
$btnCfgZonas     = $ventana.FindName("btnCfgZonas")
$btnExportCuadre = $ventana.FindName("btnExportCuadre")
$btnRefCuadre    = $ventana.FindName("btnRefCuadre")
$panelAvisoCfg   = $ventana.FindName("panelAvisoCfg")
$txtAvisoCfg     = $ventana.FindName("txtAvisoCfg")
$cvCerdo         = $ventana.FindName("cvCerdo")
$cvTarjetas      = $ventana.FindName("cvTarjetas")
$cvAprov         = $ventana.FindName("cvAprov")
$txtNotaCerdo    = $ventana.FindName("txtNotaCerdo")
$txtZonaTitulo   = $ventana.FindName("txtZonaTitulo")
$txtZonaNota     = $ventana.FindName("txtZonaNota")
$txtZonaPct      = $ventana.FindName("txtZonaPct")
$gridMat         = $ventana.FindName("gridMat")
$lblTot          = $ventana.FindName("lblTot")
$lblEsp          = $ventana.FindName("lblEsp")
$lblDif          = $ventana.FindName("lblDif")
$cvTendencia     = $ventana.FindName("cvTendencia")
$txtPromedios    = $ventana.FindName("txtPromedios")
$gridZonas       = $ventana.FindName("gridZonas")

$dpBusDesde       = $ventana.FindName("dpBusDesde")
$dpBusHasta       = $ventana.FindName("dpBusHasta")
$btnBusHoy        = $ventana.FindName("btnBusHoy")
$btnBusSemana     = $ventana.FindName("btnBusSemana")
$btnBusMes        = $ventana.FindName("btnBusMes")
$btnBusBuscar     = $ventana.FindName("btnBusBuscar")
$btnBusExport     = $ventana.FindName("btnBusExport")
$txtBusFiltro     = $ventana.FindName("txtBusFiltro")
$txtBusHintFiltro = $ventana.FindName("txtBusHintFiltro")
$lstBusCatalogo   = $ventana.FindName("lstBusCatalogo")
$btnBusAgregar    = $ventana.FindName("btnBusAgregar")
$btnBusTodos      = $ventana.FindName("btnBusTodos")
$lstBusElegidos   = $ventana.FindName("lstBusElegidos")
$btnBusQuitar     = $ventana.FindName("btnBusQuitar")
$btnBusLimpiar    = $ventana.FindName("btnBusLimpiar")
$txtBusConteo     = $ventana.FindName("txtBusConteo")
$bCajas           = $ventana.FindName("bCajas")
$bCajasSub        = $ventana.FindName("bCajasSub")
$bKilos           = $ventana.FindName("bKilos")
$bKilosSub        = $ventana.FindName("bKilosSub")
$bProm            = $ventana.FindName("bProm")
$bPromSub         = $ventana.FindName("bPromSub")
$bMin             = $ventana.FindName("bMin")
$bMinSub          = $ventana.FindName("bMinSub")
$bMax             = $ventana.FindName("bMax")
$bMaxSub          = $ventana.FindName("bMaxSub")
$bDesv            = $ventana.FindName("bDesv")
$bDesvSub         = $ventana.FindName("bDesvSub")
$gridBusResumen   = $ventana.FindName("gridBusResumen")
$txtBusNotaResumen = $ventana.FindName("txtBusNotaResumen")
$gridBusDetalle   = $ventana.FindName("gridBusDetalle")
$txtBusNotaDetalle = $ventana.FindName("txtBusNotaDetalle")
$txtBusTope       = $ventana.FindName("txtBusTope")
$txtBusProdSel    = $ventana.FindName("txtBusProdSel")
$txtBusProdHint   = $ventana.FindName("txtBusProdHint")
$bKgPz            = $ventana.FindName("bKgPz")
$bKgPzSub         = $ventana.FindName("bKgPzSub")

$txtBase.Text = $RutaBase

# ------------------------------------------------------------
#  ESTADO GLOBAL
# ------------------------------------------------------------
$script:Hechos        = @()
$script:ZonaSel        = $null
$script:MatDia         = @()
$script:TendRaw        = @{}
$script:OcupadoCuadre  = $false
$script:SuprimirCuadre = $true
$script:FechaCuadre    = (Get-Date).Date
$script:CfgZonas       = @{}
$script:CerdosPorDia   = @{}
$script:TiposCambio   = Read-TiposCambio
$script:SuprimirBus   = $true
$script:Ocupado       = $false
$script:Suprimir      = $false

function Set-Estado {
    param([string]$Texto, [string]$Color = "#8C9BC4")
    $txtEstado.Text = $Texto
    $txtEstado.Foreground = Pincel $Color
}

function Fmt {
    param([double]$Valor, [int]$Decimales = 0)
    return $Valor.ToString("N$Decimales", [System.Globalization.CultureInfo]::GetCultureInfo("es-CL"))
}

# ------------------------------------------------------------
#  AYUDANTES DE DIBUJO SOBRE CANVAS
# ------------------------------------------------------------
function Add-Texto {
    param($Canvas, [string]$Texto, [double]$X, [double]$Y, [double]$Tamano = 11,
          [string]$Color = "#8C9BC4", [string]$Alinear = "Left", [bool]$Negrita = $false)

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Texto
    $tb.FontSize = $Tamano
    $tb.Foreground = Pincel $Color
    $tb.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    if ($Negrita) { $tb.FontWeight = [System.Windows.FontWeights]::SemiBold }
    $tb.Measure([System.Windows.Size]::new([double]::PositiveInfinity, [double]::PositiveInfinity))
    $ancho = $tb.DesiredSize.Width
    $px = $X
    if ($Alinear -eq "Right")  { $px = $X - $ancho }
    if ($Alinear -eq "Center") { $px = $X - $ancho / 2 }
    [System.Windows.Controls.Canvas]::SetLeft($tb, $px)
    [System.Windows.Controls.Canvas]::SetTop($tb, $Y)
    [void]$Canvas.Children.Add($tb)
    return $tb
}

function Add-Rect {
    param($Canvas, [double]$X, [double]$Y, [double]$Ancho, [double]$Alto,
          $Relleno, [double]$Radio = 3, [string]$Tip = "")

    $r = New-Object System.Windows.Shapes.Rectangle
    $r.Width  = [math]::Max(0.5, $Ancho)
    $r.Height = [math]::Max(0.5, $Alto)
    $r.RadiusX = $Radio; $r.RadiusY = $Radio
    if ($Relleno -is [string]) { $r.Fill = Pincel $Relleno } else { $r.Fill = $Relleno }
    if ($Tip -ne "") { $r.ToolTip = $Tip }
    [System.Windows.Controls.Canvas]::SetLeft($r, $X)
    [System.Windows.Controls.Canvas]::SetTop($r, $Y)
    [void]$Canvas.Children.Add($r)
    return $r
}

function Add-Linea {
    param($Canvas, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2,
          [string]$Color = "#22305C", [double]$Grosor = 1, [bool]$Punteada = $false)

    $l = New-Object System.Windows.Shapes.Line
    $l.X1 = $X1; $l.Y1 = $Y1; $l.X2 = $X2; $l.Y2 = $Y2
    $l.Stroke = Pincel $Color
    $l.StrokeThickness = $Grosor
    if ($Punteada) {
        $dc = New-Object System.Windows.Media.DoubleCollection
        $dc.Add(4); $dc.Add(3)
        $l.StrokeDashArray = $dc
    }
    [void]$Canvas.Children.Add($l)
}

function New-Degradado {
    param([string]$Desde, [string]$Hasta, [bool]$Vertical = $false)
    $g = New-Object System.Windows.Media.LinearGradientBrush
    if ($Vertical) {
        $g.StartPoint = [System.Windows.Point]::new(0, 0)
        $g.EndPoint   = [System.Windows.Point]::new(0, 1)
    } else {
        $g.StartPoint = [System.Windows.Point]::new(0, 0)
        $g.EndPoint   = [System.Windows.Point]::new(1, 0)
    }
    $colorA = [System.Windows.Media.ColorConverter]::ConvertFromString($Desde)
    $colorB = [System.Windows.Media.ColorConverter]::ConvertFromString($Hasta)
    $g.GradientStops.Add([System.Windows.Media.GradientStop]::new($colorA, 0))
    $g.GradientStops.Add([System.Windows.Media.GradientStop]::new($colorB, 1))
    return $g
}

function Add-Vacio {
    param($Canvas, [string]$Texto = "Sin datos en este rango")
    $Canvas.Children.Clear()
    [void](Add-Texto -Canvas $Canvas -Texto $Texto -X ($Canvas.Width / 2) -Y ($Canvas.Height / 2 - 10) `
                     -Tamano 13 -Color "#4E5D85" -Alinear "Center")
}

# ------------------------------------------------------------
#  RESUMEN
# ------------------------------------------------------------
function Get-Resumen {
    param([array]$Hechos)

    $mapa = @{}
    foreach ($h in $Hechos) {
        $clave = [string]$h.Codigo
        if (-not $mapa.ContainsKey($clave)) {
            $mapa[$clave] = [pscustomobject]@{
                Codigo = $h.Codigo; Producto = $h.Producto
                Cajas = 0; KgReal = 0.0; KgEtq = 0.0; Bajo = 0
            }
        }
        $r = $mapa[$clave]
        $r.Cajas  = $r.Cajas + $h.Cajas
        $r.KgReal = $r.KgReal + $h.KgReal
        $r.KgEtq  = $r.KgEtq + $h.KgEtiqueta
        if ($h.Bajo) { $r.Bajo = $r.Bajo + $h.Cajas }
    }

    $salida = @()
    foreach ($r in $mapa.Values) {
        $reg = $r.KgReal - $r.KgEtq
        $gw = 0.0;      if ($r.KgEtq -gt 0)  { $gw = ($reg / $r.KgEtq) * 100 }
        $porCaja = 0.0; if ($r.Cajas -gt 0)  { $porCaja = $reg / $r.Cajas }
        $etqProm = 0.0; if ($r.Cajas -gt 0)  { $etqProm = $r.KgEtq / $r.Cajas }
        $prom = 0.0;    if ($r.Cajas -gt 0)  { $prom = $r.KgReal / $r.Cajas }
        $salida += [pscustomobject]@{
            Codigo = $r.Codigo; Producto = $r.Producto; Cajas = $r.Cajas
            KgReal = $r.KgReal; KgEtq = $r.KgEtq; KgReg = $reg
            Giveaway = $gw; PorCaja = $porCaja
            EtqProm = $etqProm; Promedio = $prom; Bajo = $r.Bajo
        }
    }
    return @($salida)
}

# ------------------------------------------------------------
#  DIBUJO: BARRAS POR PRODUCTO
# ------------------------------------------------------------
function Draw-Barras {
    param([array]$Resumen)

    $cvBarras.Children.Clear()
    if ($Resumen.Count -eq 0) { Add-Vacio -Canvas $cvBarras; return }

    $items = @($Resumen | Sort-Object Giveaway -Descending)
    $izq = 250.0; $der = 78.0; $ancho = $cvBarras.Width
    $altoFila = [math]::Min(42.0, ($cvBarras.Height - 10) / $items.Count)
    $maximo = 1.0
    foreach ($i in $items) { if ([math]::Abs($i.Giveaway) -gt $maximo) { $maximo = [math]::Abs($i.Giveaway) } }
    $escala = ($ancho - $izq - $der) / $maximo
    $degradado = New-Degradado -Desde "#1B3F94" -Hasta "#3D8BFF"

    for ($k = 0; $k -lt $items.Count; $k++) {
        $d = $items[$k]
        $y = $k * $altoFila + 6
        $w = [math]::Max(2.0, [math]::Abs($d.Giveaway) * $escala)
        $altoBarra = [math]::Min(23.0, $altoFila - 8)

        $nombre = $d.Producto
        if ($nombre.Length -gt 32) { $nombre = $nombre.Substring(0, 30) + "..." }
        [void](Add-Texto -Canvas $cvBarras -Texto $nombre -X ($izq - 11) -Y ($y + 3) -Tamano 11 -Color "#8C9BC4" -Alinear "Right")

        $relleno = $degradado
        if ($d.Giveaway -lt 0) { $relleno = $C.NeonRojo }
        elseif ($d.Giveaway -gt $UmbralGiveawayAlto) { $relleno = $C.NeonNaranja }

        $tip = "$($d.Producto)  [$($d.Codigo)]`n$(Fmt $d.Giveaway 2) %  |  $(Fmt $d.KgReg 1) kg regalados  |  $(Fmt $d.PorCaja 3) kg por caja  |  $(Fmt $d.Cajas 0) cajas"
        [void](Add-Rect -Canvas $cvBarras -X $izq -Y $y -Ancho $w -Alto $altoBarra -Relleno $relleno -Radio 4 -Tip $tip)
        [void](Add-Texto -Canvas $cvBarras -Texto "$(Fmt $d.Giveaway 2) %" -X ($izq + $w + 9) -Y ($y + 2) -Tamano 11.5 -Color "#E8EDFA" -Negrita $true)
    }
}

# ------------------------------------------------------------
#  DIBUJO: LINEA MENSUAL
# ------------------------------------------------------------
function Draw-Linea {
    param([array]$Hechos)

    $cvLinea.Children.Clear()
    if ($Hechos.Count -eq 0) { Add-Vacio -Canvas $cvLinea; return }

    $meses = @{}
    foreach ($h in $Hechos) {
        if (-not $meses.ContainsKey($h.Mes)) { $meses[$h.Mes] = [pscustomobject]@{ Real = 0.0; Etq = 0.0 } }
        $m = $meses[$h.Mes]
        $m.Real = $m.Real + $h.KgReal
        $m.Etq  = $m.Etq + $h.KgEtiqueta
    }
    $claves = @($meses.Keys | Sort-Object)
    if ($claves.Count -eq 0) { Add-Vacio -Canvas $cvLinea; return }

    $valores = @()
    foreach ($k in $claves) { $valores += ($meses[$k].Real - $meses[$k].Etq) }

    $ancho = $cvLinea.Width; $alto = $cvLinea.Height
    $izq = 66.0; $der = 16.0; $arr = 14.0; $aba = 36.0

    $maximo = ($valores | Measure-Object -Maximum).Maximum
    $minimo = ($valores | Measure-Object -Minimum).Minimum
    if ($maximo -eq $minimo) { $maximo = $maximo + 1; $minimo = $minimo - 1 }
    if ($minimo -gt 0) { $minimo = 0 }

    $xs = @(); $ys = @()
    for ($i = 0; $i -lt $claves.Count; $i++) {
        if ($claves.Count -eq 1) { $xs += ($izq + ($ancho - $izq - $der) / 2) }
        else { $xs += ($izq + $i * ($ancho - $izq - $der) / ($claves.Count - 1)) }
        $ys += ($arr + ($maximo - $valores[$i]) * ($alto - $arr - $aba) / ($maximo - $minimo))
    }
    $yBase = $arr + ($maximo - $minimo) * ($alto - $arr - $aba) / ($maximo - $minimo)

    for ($g = 0; $g -le 4; $g++) {
        $v = $minimo + ($maximo - $minimo) * $g / 4
        $y = $arr + ($maximo - $v) * ($alto - $arr - $aba) / ($maximo - $minimo)
        Add-Linea -Canvas $cvLinea -X1 $izq -Y1 $y -X2 ($ancho - $der) -Y2 $y -Color "#22305C"
        [void](Add-Texto -Canvas $cvLinea -Texto (Fmt $v 0) -X ($izq - 9) -Y ($y - 8) -Tamano 10 -Color "#8C9BC4" -Alinear "Right")
    }

    if ($claves.Count -gt 1) {
        $area = New-Object System.Windows.Shapes.Polygon
        $puntosArea = New-Object System.Windows.Media.PointCollection
        [void]$puntosArea.Add([System.Windows.Point]::new($xs[0], $yBase))
        for ($i = 0; $i -lt $claves.Count; $i++) {
            [void]$puntosArea.Add([System.Windows.Point]::new($xs[$i], $ys[$i]))
        }
        [void]$puntosArea.Add([System.Windows.Point]::new($xs[$claves.Count - 1], $yBase))
        $area.Points = $puntosArea
        $area.Fill = New-Degradado -Desde "#61FFA829" -Hasta "#00FFA829" -Vertical $true
        [void]$cvLinea.Children.Add($area)

        $poli = New-Object System.Windows.Shapes.Polyline
        $puntos = New-Object System.Windows.Media.PointCollection
        for ($i = 0; $i -lt $claves.Count; $i++) {
            [void]$puntos.Add([System.Windows.Point]::new($xs[$i], $ys[$i]))
        }
        $poli.Points = $puntos
        $poli.Stroke = Pincel $C.NeonNaranja
        $poli.StrokeThickness = 2.5
        $poli.StrokeLineJoin = [System.Windows.Media.PenLineJoin]::Round
        [void]$cvLinea.Children.Add($poli)
    }

    for ($i = 0; $i -lt $claves.Count; $i++) {
        $cx = $xs[$i]
        $cy = $ys[$i]
        $e = New-Object System.Windows.Shapes.Ellipse
        $e.Width = 9; $e.Height = 9
        $e.Fill = Pincel $C.NeonNaranja
        $e.Stroke = Pincel $C.Panel
        $e.StrokeThickness = 2
        $etiquetaMes = $claves[$i].Substring(5, 2) + "-" + $claves[$i].Substring(0, 4)
        $e.ToolTip = "$etiquetaMes`n$(Fmt $valores[$i] 0) kg regalados"
        [System.Windows.Controls.Canvas]::SetLeft($e, $cx - 4.5)
        [System.Windows.Controls.Canvas]::SetTop($e, $cy - 4.5)
        [void]$cvLinea.Children.Add($e)

        if ($claves.Count -le 14 -or ($i % 2) -eq 0) {
            [void](Add-Texto -Canvas $cvLinea -Texto $etiquetaMes -X $cx -Y ($alto - 24) -Tamano 9.5 -Color "#8C9BC4" -Alinear "Center")
        }
    }
}

# ------------------------------------------------------------
#  DIBUJO: HISTOGRAMA DE LA DIFERENCIA (giveaway por caja, en kg)
# ------------------------------------------------------------
function Draw-Histograma {
    param([array]$Hechos, [int]$CodigoElegido = 0)

    $cvHist.Children.Clear()

    if ($Hechos.Count -eq 0) {
        $txtNotaHist.Text = "Sin datos en este rango"
        Add-Vacio -Canvas $cvHist
        return
    }

    $filas = $Hechos
    $titulo = "Todos los productos de peso fijo"
    if ($CodigoElegido -gt 0) {
        $filas = @($Hechos | Where-Object { $_.Codigo -eq $CodigoElegido })
        $elegido = $ProductosGiveaway | Where-Object { $_.Codigo -eq $CodigoElegido }
        $titulo = "$($elegido.Nombre)  [$CodigoElegido]"
    }

    if ($filas.Count -eq 0) {
        $txtNotaHist.Text = "$titulo - sin cajas en este rango"
        Add-Vacio -Canvas $cvHist -Texto "Sin cajas de este producto en el rango"
        return
    }

    $txtNotaHist.Text = "$titulo   -   diferencia entre el peso real y el peso de etiqueta, en intervalos de $(Fmt $PasoBin 2) kg"

    $bins = @{}
    $totalCajas = 0
    foreach ($f in $filas) {
        $bins[$f.Bin] = [int]$bins[$f.Bin] + $f.Cajas
        $totalCajas = $totalCajas + $f.Cajas
    }
    $claves = @($bins.Keys | Sort-Object)

    $ancho = $cvHist.Width; $alto = $cvHist.Height
    $izq = 64.0; $der = 16.0; $arr = 14.0; $aba = 46.0
    $maxN = 1
    foreach ($k in $claves) { if ($bins[$k] -gt $maxN) { $maxN = $bins[$k] } }

    $anchoCelda = ($ancho - $izq - $der) / $claves.Count
    $bw = [math]::Max(2.0, $anchoCelda - 2)

    for ($g = 0; $g -le 4; $g++) {
        $v = $maxN * $g / 4
        $y = $arr + ($maxN - $v) * ($alto - $arr - $aba) / $maxN
        Add-Linea -Canvas $cvHist -X1 $izq -Y1 $y -X2 ($ancho - $der) -Y2 $y -Color "#22305C"
        [void](Add-Texto -Canvas $cvHist -Texto (Fmt $v 0) -X ($izq - 9) -Y ($y - 8) -Tamano 10 -Color "#8C9BC4" -Alinear "Right")
    }

    $xCero = -1.0
    for ($i = 0; $i -lt $claves.Count; $i++) {
        $k = $claves[$i]
        $d0 = $k * $PasoBin
        $d1 = ($k + 1) * $PasoBin
        $n = $bins[$k]
        $x = $izq + $i * $anchoCelda
        $y = $arr + ($maxN - $n) * ($alto - $arr - $aba) / $maxN
        $h = ($alto - $aba) - $y

        if ($d0 -le 0 -and 0 -lt $d1) { $xCero = $x }

        # rojo: la caja pesa MENOS que la etiqueta (incumplimiento)
        # verde: giveaway muy chico, azul: giveaway normal, naranja: giveaway alto
        $color = $C.NeonAzul
        if ($d1 -le 0.0001)      { $color = $C.NeonRojo }
        elseif ($d0 -ge 0.50)    { $color = $C.NeonNaranja }

        $pct = 0.0
        if ($totalCajas -gt 0) { $pct = ($n / $totalCajas) * 100 }
        $tip = "Giveaway entre $(Fmt $d0 2) y $(Fmt $d1 2) kg`n$(Fmt $n 0) cajas ($(Fmt $pct 1) %)"
        [void](Add-Rect -Canvas $cvHist -X $x -Y $y -Ancho $bw -Alto $h -Relleno $color -Radio 2 -Tip $tip)

        if ($claves.Count -le 42 -or ($i % 3) -eq 0) {
            [void](Add-Texto -Canvas $cvHist -Texto (Fmt $d0 2) -X ($x + $bw / 2) -Y ($alto - $aba + 6) -Tamano 9 -Color "#8C9BC4" -Alinear "Center")
        }
    }

    if ($xCero -ge 0) {
        Add-Linea -Canvas $cvHist -X1 $xCero -Y1 $arr -X2 $xCero -Y2 ($alto - $aba) -Color "#00FFA3" -Grosor 2 -Punteada $true
        [void](Add-Texto -Canvas $cvHist -Texto "Peso exacto de etiqueta" -X ($xCero + 7) -Y $arr -Tamano 11 -Color "#00FFA3" -Negrita $true)
    }

    [void](Add-Texto -Canvas $cvHist -Texto "Kilos regalados por caja (Peso Neto - Peso Neto Etiqueta)" -X ($ancho / 2) -Y ($alto - 18) -Tamano 10.5 -Color "#8C9BC4" -Alinear "Center")
}

# ------------------------------------------------------------
#  DIBUJO: MAPA DE CALOR
# ------------------------------------------------------------
function Draw-Mapa {
    param([array]$Hechos)

    $cvMapa.Children.Clear()
    if ($Hechos.Count -eq 0) { Add-Vacio -Canvas $cvMapa; return }

    $celdas = @{}
    $meses = @{}
    $prods = @{}
    foreach ($h in $Hechos) {
        $clave = [string]$h.Codigo + "|" + $h.Mes
        if (-not $celdas.ContainsKey($clave)) { $celdas[$clave] = [pscustomobject]@{ Real = 0.0; Etq = 0.0 } }
        $c = $celdas[$clave]
        $c.Real = $c.Real + $h.KgReal
        $c.Etq  = $c.Etq + $h.KgEtiqueta
        $meses[$h.Mes] = $true
        $prods[[string]$h.Codigo] = $true
    }

    $ms = @($meses.Keys | Sort-Object)
    $ps = @($prods.Keys | Sort-Object)
    $nombrePorCodigo = @{}
    foreach ($pg in $ProductosGiveaway) { $nombrePorCodigo[[string]$pg.Codigo] = $pg.Nombre }
    if ($ms.Count -eq 0 -or $ps.Count -eq 0) { Add-Vacio -Canvas $cvMapa; return }

    $izq = 196.0; $arr = 26.0
    $cw = [math]::Max(30.0, [math]::Min(62.0, ($cvMapa.Width - $izq - 10) / $ms.Count))
    $ch = [math]::Max(20.0, [math]::Min(32.0, ($cvMapa.Height - $arr - 10) / $ps.Count))

    $maxAbs = 1.0
    foreach ($p in $ps) { foreach ($m in $ms) {
        $c = $celdas[$p + "|" + $m]
        if ($null -ne $c -and $c.Etq -gt 0) {
            $v = [math]::Abs((($c.Real - $c.Etq) / $c.Etq) * 100)
            if ($v -gt $maxAbs) { $maxAbs = $v }
        }
    } }

    for ($j = 0; $j -lt $ms.Count; $j++) {
        $etiqueta = $ms[$j].Substring(5, 2) + "-" + $ms[$j].Substring(2, 2)
        [void](Add-Texto -Canvas $cvMapa -Texto $etiqueta -X ($izq + $j * $cw + $cw / 2) -Y ($arr - 17) -Tamano 9 -Color "#8C9BC4" -Alinear "Center")
    }

    for ($i = 0; $i -lt $ps.Count; $i++) {
        $nombre = $nombrePorCodigo[$ps[$i]]
        if ([string]::IsNullOrEmpty($nombre)) { $nombre = $ps[$i] }
        $corto = $nombre
        if ($corto.Length -gt 26) { $corto = $corto.Substring(0, 24) + "..." }
        [void](Add-Texto -Canvas $cvMapa -Texto $corto -X ($izq - 9) -Y ($arr + $i * $ch + ($ch / 2) - 8) -Tamano 10 -Color "#8C9BC4" -Alinear "Right")

        for ($j = 0; $j -lt $ms.Count; $j++) {
            $c = $celdas[$ps[$i] + "|" + $ms[$j]]
            $x = $izq + $j * $cw + 1
            $y = $arr + $i * $ch + 1
            $w = $cw - 2
            $h = $ch - 2

            if ($null -eq $c -or $c.Etq -le 0) {
                [void](Add-Rect -Canvas $cvMapa -X $x -Y $y -Ancho $w -Alto $h -Relleno "#0A1128" -Radio 3)
                continue
            }

            $v = (($c.Real - $c.Etq) / $c.Etq) * 100
            $intensidad = [math]::Min(1.0, [math]::Abs($v) / $maxAbs) * 0.82 + 0.14
            $alfa = [int]([math]::Round($intensidad * 255))
            $hex = "{0:X2}" -f $alfa
            $base = "3D8BFF"
            if ($v -lt 0) { $base = "FF3B6B" } elseif ($v -gt $UmbralGiveawayAlto) { $base = "FFA829" }

            $tip = "$nombre`n$($ms[$j])  |  $(Fmt $v 2) %  |  $(Fmt ($c.Real - $c.Etq) 1) kg"
            [void](Add-Rect -Canvas $cvMapa -X $x -Y $y -Ancho $w -Alto $h -Relleno ("#" + $hex + $base) -Radio 3 -Tip $tip)

            if ($cw -ge 42) {
                [void](Add-Texto -Canvas $cvMapa -Texto (Fmt $v 1) -X ($x + $w / 2) -Y ($y + $h / 2 - 8) -Tamano 9 -Color "#E8EDFA" -Alinear "Center")
            }
        }
    }
}

# ------------------------------------------------------------
#  PINTAR TARJETAS + TABLA
# ------------------------------------------------------------
function Draw-Tarjetas {
    param([array]$Resumen)

    $cajas = 0; $kgReal = 0.0; $kgEtq = 0.0; $bajo = 0
    foreach ($r in $Resumen) {
        $cajas  = $cajas + $r.Cajas
        $kgReal = $kgReal + $r.KgReal
        $kgEtq  = $kgEtq + $r.KgEtq
        $bajo   = $bajo + $r.Bajo
    }
    $reg = $kgReal - $kgEtq
    $gw = 0.0;      if ($kgEtq -gt 0) { $gw = ($reg / $kgEtq) * 100 }
    $porCaja = 0.0; if ($cajas -gt 0) { $porCaja = $reg / $cajas }
    $pctBajo = 0.0; if ($cajas -gt 0) { $pctBajo = ($bajo / $cajas) * 100 }

    $vGw.Text    = "$(Fmt $gw 2) %"
    $vKg.Text    = "$(Fmt $reg 1) kg"
    $sKg.Text    = "$(Fmt $kgReal 1) reales / $(Fmt $kgEtq 1) etiqueta"
    $vCajas.Text = Fmt $cajas 0
    $vProm.Text  = "$(Fmt $porCaja 3) kg"

    # Dinero: cada producto se valoriza con su propio precio y moneda
    $tc = $script:TiposCambio
    $plata = 0.0; $sinPrecio = 0
    foreach ($r in $Resumen) {
        $v = Get-PesosRegalados -Codigo $r.Codigo -KgRegalados $r.KgReg -Tipos $tc
        if ($v -eq 0 -and $r.KgReg -ne 0 -and -not $PreciosGiveaway.ContainsKey($r.Codigo)) { $sinPrecio++ }
        $plata = $plata + $v
    }
    $vPlata.Text = "$" + (Fmt $plata 0)
    if ($sinPrecio -gt 0) {
        $sPlata.Text = "$sinPrecio producto(s) sin precio cargado"
    } else {
        $sPlata.Text = "USD $(Fmt $tc['USD'] 2)  EUR $(Fmt $tc['EUR'] 2)  JPY $(Fmt $tc['JPY'] 2)"
    }

    $vBajo.Text  = Fmt $bajo 0
    $sBajo.Text  = "$(Fmt $pctBajo 2) % - riesgo normativo"

    if ($bajo -gt 0) {
        $vBajo.Foreground = Pincel $C.NeonRojo
        $barraBajo.Background = Pincel $C.NeonRojo
        $txtAviso.Text = "Atencion: $(Fmt $bajo 0) caja(s) con peso real por debajo del peso declarado en la etiqueta. Esto no es giveaway, es un incumplimiento de peso que conviene revisar caja por caja."
        $panelAviso.Visibility = "Visible"
    } else {
        $vBajo.Foreground = Pincel $C.NeonVerde
        $barraBajo.Background = Pincel $C.NeonVerde
        $panelAviso.Visibility = "Collapsed"
    }
}

function Draw-Tabla {
    param([array]$Resumen)

    $filas = @()
    foreach ($r in @($Resumen | Sort-Object KgReg -Descending)) {
        $filas += [pscustomobject]@{
            Producto = "$($r.Producto)  [$($r.Codigo)]"
            Etiqueta = Fmt $r.EtqProm 2
            Cajas    = Fmt $r.Cajas 0
            Promedio = Fmt $r.Promedio 3
            PorCaja  = Fmt $r.PorCaja 3
            Giveaway = "$(Fmt $r.Giveaway 2) %"
            KgReg    = Fmt $r.KgReg 1
            Precio   = Get-DetallePrecio $r.Codigo
            Plata    = "$" + (Fmt (Get-PesosRegalados -Codigo $r.Codigo -KgRegalados $r.KgReg -Tipos $script:TiposCambio) 0)
            Bajo     = Fmt $r.Bajo 0
        }
    }
    $grid.ItemsSource = $filas
}

# ------------------------------------------------------------
#  CUADRE DE PIEZAS - DIBUJO E INTERACCION
# ------------------------------------------------------------
function Add-Path {
    param($Canvas, [string]$Data, $Relleno = $null, [string]$Trazo = '',
          [double]$Grosor = 1.0, [string]$Tag = '', [bool]$Punteada = $false)
    $p = New-Object System.Windows.Shapes.Path
    $p.Data = [System.Windows.Media.Geometry]::Parse($Data)
    if ($null -ne $Relleno) {
        if ($Relleno -is [string]) { $p.Fill = Pincel $Relleno } else { $p.Fill = $Relleno }
    }
    if ($Trazo -ne '') {
        $p.Stroke = Pincel $Trazo
        $p.StrokeThickness = $Grosor
        $p.StrokeLineJoin = 'Round'
        $p.StrokeStartLineCap = 'Round'
        $p.StrokeEndLineCap = 'Round'
    }
    if ($Punteada) {
        $dc = New-Object System.Windows.Media.DoubleCollection
        $dc.Add(5); $dc.Add(4)
        $p.StrokeDashArray = $dc
    }
    if ($Tag -ne '') {
        $p.Tag = $Tag
        $p.Cursor = [System.Windows.Input.Cursors]::Hand
    }
    [void]$Canvas.Children.Add($p)
    return $p
}

function Add-Elipse {
    param($Canvas, [double]$CX, [double]$CY, [double]$RX, [double]$RY,
          $Relleno = $null, [string]$Trazo = '', [double]$Grosor = 1.0)
    $e = New-Object System.Windows.Shapes.Ellipse
    $e.Width = $RX * 2; $e.Height = $RY * 2
    if ($null -ne $Relleno) {
        if ($Relleno -is [string]) { $e.Fill = Pincel $Relleno } else { $e.Fill = $Relleno }
    }
    if ($Trazo -ne '') { $e.Stroke = Pincel $Trazo; $e.StrokeThickness = $Grosor }
    [System.Windows.Controls.Canvas]::SetLeft($e, $CX - $RX)
    [System.Windows.Controls.Canvas]::SetTop($e, $CY - $RY)
    [void]$Canvas.Children.Add($e)
    return $e
}

function New-EfectoGlow {
    $ef = New-Object System.Windows.Media.Effects.DropShadowEffect
    $ef.Color = [System.Windows.Media.ColorConverter]::ConvertFromString("#00E5FF")
    $ef.BlurRadius = 16
    $ef.ShadowDepth = 0
    $ef.Opacity = 0.85
    return $ef
}

# Un solo manejador de clic para zonas, etiquetas y elipse de la cola.
# Segundo clic sobre la misma zona la deselecciona (vuelve al cuadre global).
$script:ClickZona = {
    $zid = [string]$this.Tag
    if ($script:ZonaSel -eq $zid) { $script:ZonaSel = $null } else { $script:ZonaSel = $zid }
    Render-Cuadre
}



function Draw-Historial30 {
    # Barras apiladas por dia. Los cortes van abajo y despojos/decomisos
    # arriba a proposito: si se disparan, sobresalen por el techo.
    param($Cv, $Hist, [double]$X, [double]$Y, [double]$W, [double]$H)

    [void](Add-Rect -Canvas $Cv -X $X -Y $Y -Ancho $W -Alto $H -Relleno $C.Panel -Radio 7)
    [void](Add-Texto -Canvas $Cv -Texto "ULTIMOS 30 DIAS" -X ($X + 18) -Y ($Y + 12) -Tamano 15 -Color $C.Texto -Negrita $true)

    if (-not $Hist -or $Hist.Count -eq 0) {
        [void](Add-Texto -Canvas $Cv -Texto "sin datos en el periodo" -X ($X + 14) -Y ($Y + 32) -Tamano 9.5 -Color '#4E5D85')
        return
    }
    [void](Add-Texto -Canvas $Cv -Texto ("cada barra = 100 % del kilaje de ese dia   -   " + $Hist.Count + " dias con produccion") `
                     -X ($X + 18) -Y ($Y + 34) -Tamano 11 -Color '#4E5D85')

    $segs = New-Object System.Collections.Generic.List[object]
    [void]$segs.Add([pscustomobject]@{ Id='Cortes'; Nom='Cortes'; Color=$C.NeonVerde })
    foreach ($p in $PanelesCuadre) {
        $corto = $p.Nombre
        if ($corto -eq 'DESPOJOS Y DECOMISOS') { $corto = 'Despojos/decom.' }
        elseif ($corto -eq 'TRIMMINGS Y RECORTES') { $corto = 'Recortes' }
        elseif ($corto -eq 'CUERO Y GRASAS') { $corto = 'Cuero/grasa' }
        else { $corto = $corto.Substring(0,1) + $corto.Substring(1).ToLower() }
        [void]$segs.Add([pscustomobject]@{ Id=$p.Id; Nom=$corto; Color=$p.Color })
    }

    $gx = $X + 52.0; $gtop = $Y + 58.0; $gh = $H - 148.0; $gw = $W - 72.0
    for ($p = 0; $p -le 100; $p += 25) {
        $yy = $gtop + $gh * (1 - $p / 100)
        Add-Linea -Canvas $Cv -X1 ($X + 46) -Y1 $yy -X2 ($X + $W - 16) -Y2 $yy -Color '#16204A' -Grosor 0.7
        [void](Add-Texto -Canvas $Cv -Texto ("$p%") -X ($X + 42) -Y ($yy - 8) -Tamano 10 -Color '#4E5D85' -Alinear "Right")
    }

    $n = $Hist.Count
    $bw = $gw / [math]::Max(1, $n)
    $ancho = [math]::Max(3.0, $bw - 6)
    for ($i = 0; $i -lt $n; $i++) {
        $d = $Hist[$i]
        if ($d.Total -le 0) { continue }
        $bxx = $gx + $i * $bw
        $yy = $gtop + $gh
        foreach ($sg in $segs) {
            $kg = [double]$d.($sg.Id)
            $f = $kg / $d.Total
            $hh = $gh * $f
            if ($hh -lt 0.4) { continue }
            $yy = $yy - $hh
            $r = Add-Rect -Canvas $Cv -X $bxx -Y $yy -Ancho $ancho -Alto $hh -Relleno $sg.Color -Radio 0 `
                          -Tip ($d.Dia + "  " + $sg.Nom + ": " + (Fmt $kg 0) + " kg (" + (Fmt ($f * 100) 1) + " %)")
            $r.Opacity = 0.92
            # Con el ancho completo el numero cabe dentro de la franja
            if ($hh -ge 13 -and $ancho -ge 26) {
                [void](Add-Texto -Canvas $Cv -Texto ((Fmt ($f * 100) 0) + "%") -X ($bxx + $ancho / 2) -Y ($yy + $hh / 2 - 8) -Tamano 11 -Color '#0A1128' -Alinear "Center" -Negrita $true)
            }
        }
        if ($n -le 32) {
            $et = $d.Dia
            if ($et.Length -ge 10) { $et = $et.Substring(8, 2) }
            [void](Add-Texto -Canvas $Cv -Texto $et -X ($bxx + $ancho / 2) -Y ($gtop + $gh + 6) -Tamano 10.5 -Color '#6B7BA8' -Alinear "Center")
        }
    }

    # Numeros del periodo: cuanto represento cada parte en los 30 dias
    $totPer = 0.0
    foreach ($d in $Hist) { $totPer = $totPer + $d.Total }
    $ly = $gtop + $gh + 34.0
    $lx = $X + 18.0
    $anchoCel = ($W - 36) / $segs.Count
    foreach ($sg in $segs) {
        $acum = 0.0
        foreach ($d in $Hist) { $acum = $acum + [double]$d.($sg.Id) }
        $pc = 0.0
        if ($totPer -gt 0) { $pc = $acum / $totPer * 100 }
        [void](Add-Rect -Canvas $Cv -X $lx -Y ($ly + 5) -Ancho 12 -Alto 12 -Relleno $sg.Color -Radio 2)
        [void](Add-Texto -Canvas $Cv -Texto $sg.Nom -X ($lx + 18) -Y $ly -Tamano 11.5 -Color $C.Texto2)
        [void](Add-Texto -Canvas $Cv -Texto ((Fmt $pc 1) + " %") -X ($lx + 18) -Y ($ly + 18) -Tamano 20 -Color $sg.Color -Negrita $true)
        [void](Add-Texto -Canvas $Cv -Texto ((Fmt $acum 0) + " kg") -X ($lx + 18) -Y ($ly + 44) -Tamano 10.5 -Color '#4E5D85')
        $lx = $lx + $anchoCel
    }
}

function Draw-Paneles {
    # Cuatro tarjetas a la derecha del cerdo. Miden KG POR CERDO porque
    # son productos a granel: contar "piezas" no tendria sentido.
    param($Cv, $AggR, [int]$Cerdos)

    # Lienzo propio: 2 columnas x 3 filas. Al no compartir Viewbox con el
    # cerdo, estas tarjetas se escalan solas y se ven mucho mas grandes.
    $Cv.Children.Clear()
    $X0 = 8.0; $ANCHO = 236.0; $ALTO = 178.0; $SEP = 15.0
    $col = @(0, 1, 0, 1, 0); $fil = @(0, 0, 1, 1, 2)
    $i = 0
    foreach ($p in $PanelesCuadre) {
        $px = $X0 + $col[$i] * ($ANCHO + $SEP)
        $py = 4.0 + $fil[$i] * ($ALTO + $SEP)
        $d  = $AggR.Pan[$p.Id]

        [void](Add-Rect -Canvas $Cv -X $px -Y $py -Ancho $ANCHO -Alto $ALTO -Relleno $C.Panel -Radio 7)
        [void](Add-Rect -Canvas $Cv -X $px -Y $py -Ancho 3.5 -Alto $ALTO -Relleno $p.Color -Radio 2)

        [void](Add-Texto -Canvas $Cv -Texto $p.Nombre -X ($px + 14) -Y ($py + 10) -Tamano 11.5 -Color $p.Color -Negrita $true)
        [void](Add-Texto -Canvas $Cv -Texto $p.Nota   -X ($px + 14) -Y ($py + 27) -Tamano 9.5  -Color '#4E5D85')

        $kgCerdo = 0.0
        if ($Cerdos -gt 0) { $kgCerdo = $d.Kg / $Cerdos }
        [void](Add-Texto -Canvas $Cv -Texto (Fmt $kgCerdo 2) -X ($px + 14) -Y ($py + 44) -Tamano 24 -Color $p.Color -Negrita $true)
        [void](Add-Texto -Canvas $Cv -Texto ("kg por cerdo - " + (Fmt $d.Kg 0) + " kg") -X ($px + 14) -Y ($py + 74) -Tamano 9.5 -Color $C.Texto2)

        $pct = 0.0
        if ($AggR.KgTotal -gt 0) { $pct = $d.Kg / $AggR.KgTotal * 100 }
        [void](Add-Texto -Canvas $Cv -Texto ((Fmt $pct 1) + " %") -X ($px + $ANCHO - 14) -Y ($py + 48) -Tamano 16 -Color '#9FB0DA' -Alinear "Right" -Negrita $true)
        [void](Add-Texto -Canvas $Cv -Texto "del total kg" -X ($px + $ANCHO - 14) -Y ($py + 70) -Tamano 9 -Color '#4E5D85' -Alinear "Right")

        Add-Linea -Canvas $Cv -X1 ($px + 14) -Y1 ($py + 92) -X2 ($px + $ANCHO - 14) -Y2 ($py + 92) -Color $C.Linea
        $top = @($d.Mats.ToArray() | Sort-Object -Property Kg -Descending | Select-Object -First 3)
        $yy = $py + 97.0
        if ($top.Count -eq 0) {
            [void](Add-Texto -Canvas $Cv -Texto "sin materiales asignados" -X ($px + 14) -Y $yy -Tamano 9.5 -Color '#4E5D85')
        }
        foreach ($m in $top) {
            $nom = $m.Producto
            if ($nom.Length -gt 26) { $nom = $nom.Substring(0, 25) + "." }
            [void](Add-Texto -Canvas $Cv -Texto $nom -X ($px + 14) -Y $yy -Tamano 9.5 -Color '#9FB0DA')
            [void](Add-Texto -Canvas $Cv -Texto ((Fmt $m.Kg 0) + " kg") -X ($px + $ANCHO - 14) -Y $yy -Tamano 9.5 -Color $C.Texto -Alinear "Right")
            $frac = 0
            if ($d.Kg -gt 0) { $frac = $m.Kg / $d.Kg }
            [void](Add-Rect -Canvas $Cv -X ($px + 14) -Y ($yy + 14) -Ancho ($ANCHO - 28) -Alto 2.4 -Relleno '#16204A' -Radio 1)
            $b = Add-Rect -Canvas $Cv -X ($px + 14) -Y ($yy + 14) -Ancho (($ANCHO - 28) * $frac) -Alto 2.4 -Relleno $p.Color -Radio 1
            $b.Opacity = 0.85
            $yy = $yy + 24
        }
        $i = $i + 1
    }

    # ---- Tarjeta de total, en el sexto lugar de la grilla ----
    $tx = $X0 + ($ANCHO + $SEP); $ty = 4.0 + 2 * ($ALTO + $SEP)
    [void](Add-Rect -Canvas $Cv -X $tx -Y $ty -Ancho $ANCHO -Alto $ALTO -Relleno '#101A3D' -Radio 7)
    $bordeTot = New-Object System.Windows.Shapes.Rectangle
    $bordeTot.Width = $ANCHO; $bordeTot.Height = $ALTO; $bordeTot.RadiusX = 7; $bordeTot.RadiusY = 7
    $bordeTot.Stroke = Pincel '#3D8BFF'; $bordeTot.StrokeThickness = 1.4
    $bordeTot.Fill = [System.Windows.Media.Brushes]::Transparent
    [System.Windows.Controls.Canvas]::SetLeft($bordeTot, $tx)
    [System.Windows.Controls.Canvas]::SetTop($bordeTot, $ty)
    [void]$Cv.Children.Add($bordeTot)

    [void](Add-Texto -Canvas $Cv -Texto "TOTAL PRODUCCION" -X ($tx + 14) -Y ($ty + 10) -Tamano 11 -Color '#3D8BFF' -Negrita $true)
    [void](Add-Texto -Canvas $Cv -Texto ("$Cerdos cerdos - vara fria " + (Fmt $script:PesoVara 1) + " kg") -X ($tx + 14) -Y ($ty + 27) -Tamano 9 -Color '#4E5D85')
    $kgCerdoTot = 0.0
    if ($Cerdos -gt 0) { $kgCerdoTot = $AggR.KgTotal / $Cerdos }
    [void](Add-Texto -Canvas $Cv -Texto (Fmt $kgCerdoTot 1) -X ($tx + 14) -Y ($ty + 44) -Tamano 25 -Color $C.Texto -Negrita $true)
    [void](Add-Texto -Canvas $Cv -Texto "kg por cerdo" -X ($tx + 14) -Y ($ty + 76) -Tamano 9 -Color $C.Texto2)
    [void](Add-Texto -Canvas $Cv -Texto (Fmt $AggR.KgTotal 0) -X ($tx + $ANCHO - 14) -Y ($ty + 48) -Tamano 17 -Color $C.NeonVerde -Alinear "Right" -Negrita $true)
    [void](Add-Texto -Canvas $Cv -Texto "kg totales del dia" -X ($tx + $ANCHO - 14) -Y ($ty + 70) -Tamano 8.5 -Color '#4E5D85' -Alinear "Right")
    Add-Linea -Canvas $Cv -X1 ($tx + 14) -Y1 ($ty + 95) -X2 ($tx + $ANCHO - 14) -Y2 ($ty + 95) -Color $C.Linea

    $rec = 0.0
    if ($Cerdos -gt 0 -and $script:PesoVara -gt 0) { $rec = $AggR.KgTotal / ($Cerdos * $script:PesoVara) * 100 }
    $kgSin = 0.0
    foreach ($m in $AggR.Sin) { $kgSin = $kgSin + $m.Kg }
    $colRec = $C.NeonNaranja
    if ($rec -ge 60) { $colRec = $C.NeonVerde }
    $colSin = $C.NeonVerde
    if ($kgSin -gt 0) { $colSin = $C.NeonNaranja }

    [void](Add-Texto -Canvas $Cv -Texto "Recuperado de la vara" -X ($tx + 14) -Y ($ty + 106) -Tamano 9 -Color $C.Texto2)
    [void](Add-Texto -Canvas $Cv -Texto ((Fmt $rec 1) + " %") -X ($tx + $ANCHO - 14) -Y ($ty + 106) -Tamano 10 -Color $colRec -Alinear "Right" -Negrita $true)
    [void](Add-Texto -Canvas $Cv -Texto "Sin clasificar" -X ($tx + 14) -Y ($ty + 127) -Tamano 9 -Color $C.Texto2)
    [void](Add-Texto -Canvas $Cv -Texto ((Fmt $kgSin 0) + " kg") -X ($tx + $ANCHO - 14) -Y ($ty + 127) -Tamano 10 -Color $colSin -Alinear "Right" -Negrita $true)
    [void](Add-Texto -Canvas $Cv -Texto "Cortes comerciales" -X ($tx + 14) -Y ($ty + 148) -Tamano 9 -Color $C.Texto2)
    [void](Add-Texto -Canvas $Cv -Texto ((Fmt $AggR.KgZonas 0) + " kg") -X ($tx + $ANCHO - 14) -Y ($ty + 148) -Tamano 10 -Color '#9FB0DA' -Alinear "Right" -Negrita $true)

}

function Draw-Aprovechamiento {
    # Lienzo aparte que ocupa el ancho completo de la ventana. Antes esto
    # compartia Viewbox con el cerdo y se encogia al 65 %.
    param($Cv, $AggR)
    $Cv.Children.Clear()
    $bx = 8.0; $by = 4.0; $bw = 1844.0
    [void](Add-Rect -Canvas $Cv -X $bx -Y $by -Ancho $bw -Alto 124 -Relleno $C.Panel -Radio 7)
    [void](Add-Texto -Canvas $Cv -Texto ("APROVECHAMIENTO DEL DIA - " + (Fmt $AggR.KgTotal 0) + " kg repartidos") -X ($bx + 18) -Y ($by + 12) -Tamano 15 -Color $C.Texto -Negrita $true)

    $tot = $AggR.KgTotal
    $segs = New-Object System.Collections.Generic.List[object]
    [void]$segs.Add([pscustomobject]@{ Nom='Cortes comerciales'; Kg=$AggR.KgZonas; Color=$C.NeonVerde })
    foreach ($p in $PanelesCuadre) {
        [void]$segs.Add([pscustomobject]@{ Nom=$p.Nombre; Kg=$AggR.Pan[$p.Id].Kg; Color=$p.Color })
    }
    $cx = $bx + 16.0; $usable = $bw - 32
    foreach ($sg in $segs) {
        if ($tot -le 0) { break }
        $w = $usable * ($sg.Kg / $tot)
        if ($w -lt 0.5) { continue }
        $r = Add-Rect -Canvas $Cv -X $cx -Y ($by + 42) -Ancho $w -Alto 38 -Relleno $sg.Color -Radio 0 `
                      -Tip ($sg.Nom + ": " + (Fmt $sg.Kg 0) + " kg")
        $r.Opacity = 0.92
        $txtPct = (Fmt ($sg.Kg / $tot * 100) 1) + " %"
        if ($w -gt 46) {
            # Cabe dentro: texto oscuro sobre el color para maximo contraste
            [void](Add-Texto -Canvas $Cv -Texto $txtPct -X ($cx + $w / 2) -Y ($by + 51) -Tamano 14 -Color '#0A1128' -Alinear "Center" -Negrita $true)
        } else {
            # Franja angosta: el numero va arriba, en su color, para que
            # ninguna parte se quede sin porcentaje.
            [void](Add-Texto -Canvas $Cv -Texto $txtPct -X ($cx + $w / 2) -Y ($by + 24) -Tamano 11 -Color $sg.Color -Alinear "Center" -Negrita $true)
            Add-Linea -Canvas $Cv -X1 ($cx + $w / 2) -Y1 ($by + 40) -X2 ($cx + $w / 2) -Y2 ($by + 42) -Color $sg.Color
        }
        $cx = $cx + $w
    }
    $lx = $bx + 14.0
    foreach ($sg in $segs) {
        [void](Add-Rect -Canvas $Cv -X $lx -Y ($by + 96) -Ancho 11 -Alto 11 -Relleno $sg.Color -Radio 2)
        $t = Add-Texto -Canvas $Cv -Texto $sg.Nom -X ($lx + 16) -Y ($by + 92) -Tamano 11.5 -Color $C.Texto2
        $t.Measure([System.Windows.Size]::new([double]::PositiveInfinity, [double]::PositiveInfinity))
        $lx = $lx + 28 + $t.DesiredSize.Width
    }

    # Grafico de los ultimos 30 dias, debajo de la barra del dia
    Draw-Historial30 -Cv $Cv -Hist $script:HistCuadre -X $bx -Y ($by + 136) -W $bw -H 322
}

function Draw-Cerdo {
    param($AggR, [int]$Cerdos)

    $cv = $cvCerdo
    $cv.Children.Clear()

    # Cuerpo con volumen. La silueta esta calcada de la referencia britanica
    # (2 orejas, 4 patas, cola enroscada). F0 = EvenOdd para el hueco del rizo.
    $grad = New-Degradado -Desde '#1B2650' -Hasta '#0C132C' -Vertical $true
    [void](Add-Path -Canvas $cv -Data $CerdoCuerpo -Relleno $grad)

    # Zonas recortadas a la silueta
    $cz = New-Object System.Windows.Controls.Canvas
    $cz.Width = $cv.Width; $cz.Height = $cv.Height
    $cz.Clip = [System.Windows.Media.Geometry]::Parse($CerdoCuerpo)
    foreach ($z in $ZonasCuadre) {
        if ($z.D.StartsWith('ELIPSE')) { continue }
        $selz = ($script:ZonaSel -eq $z.Id)
        $esFilete = ($z.Id -eq 'filete' -or $z.Id -eq 'baby')   # cortes internos: punteados
        $fill = [System.Windows.Media.Brushes]::Transparent
        if ($selz) { $fill = Pincel '#3D8BFF'; $fill.Opacity = 0.30 }
        $trazo = '#2A3A6B'; $gros = 1.15
        if ($esFilete) { $trazo = '#4C7BD9'; $gros = 1.5 }
        if ($selz)     { $trazo = '#00E5FF'; $gros = 2.4 }
        $p = Add-Path -Canvas $cz -Data $z.D -Relleno $fill -Trazo $trazo -Grosor $gros -Tag $z.Id -Punteada $esFilete
        if ($selz) { $p.Effect = New-EfectoGlow }
        $p.Add_MouseLeftButtonDown($script:ClickZona)
    }
    [void]$cv.Children.Add($cz)

    # Estrias del baby back: se dibujan dentro del recorte de la silueta
    # para que se vean como costillas y no como un rectangulo vacio.
    $cz2 = New-Object System.Windows.Controls.Canvas
    $cz2.Width = $cv.Width; $cz2.Height = $cv.Height
    $cz2.Clip = [System.Windows.Media.Geometry]::Parse($CerdoCuerpo)
    $cz2.IsHitTestVisible = $false
    foreach ($e in $EstriasBaby) {
        [void](Add-Path -Canvas $cz2 -Data $e -Trazo '#3A5490' -Grosor 1.0)
    }
    [void]$cv.Children.Add($cz2)

    # Contorno de la silueta
    [void](Add-Path -Canvas $cv -Data $CerdoCuerpo -Trazo '#7C90BE' -Grosor 1.7)

    # Detalles anatomicos calcados de la referencia
    [void](Add-Path -Canvas $cv -Data 'M174.2,95.7 L179.7,100.1 L184.1,106.7 L188.5,113.3 L195.1,119.9 L200.6,123.2' -Trazo '#4C619E' -Grosor 1.5)
    foreach ($dd in @('M135.7,202.4 L137.9,202.4 L140.1,203.5 L142.3,204.6 L144.5,206.8 L146.7,209.0 L150.0,212.3','M139.0,216.7 L141.2,220.0 L143.4,222.2 L146.7,226.6','M192.9,217.8 L189.6,223.3 L185.2,227.7 L179.7,231.0 L174.2,233.2 L168.7,235.4 L156.6,235.4','M327.1,260.7 L333.7,269.5 L338.1,281.6 L341.4,293.7 L344.7,305.8 L348.0,317.9 L348.0,328.9','M342.5,342.1 L343.6,350.9 L346.9,356.4 L345.8,365.2 L339.2,371.8 L333.7,377.3 L327.1,381.7','M685.7,229.9 L674.7,246.4 L662.6,262.9 L657.1,279.4 L658.2,295.9')) {
        [void](Add-Path -Canvas $cv -Data $dd -Trazo '#33457C' -Grosor 1.4)
    }
    [void](Add-Elipse -Canvas $cv -CX 208.9 -CY 165.0 -RX 6.2 -RY 6.2 -Relleno '#0A1128' -Trazo '#4C619E' -Grosor 1.3)
    [void](Add-Elipse -Canvas $cv -CX 207.3 -CY 163.3 -RX 1.7 -RY 1.7 -Relleno '#8C9BC4')

    # Zona de la cola: elipse sin recorte (la cola sobresale de la silueta)
    foreach ($z in $ZonasCuadre) {
        if (-not $z.D.StartsWith('ELIPSE')) { continue }
        $partes = ($z.D -split ':')[1] -split ','
        $ex = [double]$partes[0]; $ey = [double]$partes[1]
        $ew = [double]$partes[2]; $eh = [double]$partes[3]
        $selz = ($script:ZonaSel -eq $z.Id)
        $e = New-Object System.Windows.Shapes.Ellipse
        $e.Width = $ew; $e.Height = $eh
        if ($selz) {
            $f = Pincel '#3D8BFF'; $f.Opacity = 0.22
            $e.Fill = $f
            $e.Stroke = Pincel '#00E5FF'; $e.StrokeThickness = 2.2
            $e.Effect = New-EfectoGlow
        } else {
            $e.Fill = [System.Windows.Media.Brushes]::Transparent
        }
        $e.Tag = $z.Id
        $e.Cursor = [System.Windows.Input.Cursors]::Hand
        [System.Windows.Controls.Canvas]::SetLeft($e, $ex)
        [System.Windows.Controls.Canvas]::SetTop($e, $ey)
        $e.Add_MouseLeftButtonDown($script:ClickZona)
        [void]$cv.Children.Add($e)
    }

    # Lineas guia y etiquetas (nombre + porcentaje) de las 12 zonas
    foreach ($z in $ZonasCuadre) {
        $dz  = $AggR.Agg[$z.Id]
        $esp = $Cerdos * $z.Mult
        $pct = 0.0
        if ($esp -gt 0) { $pct = $dz.Piezas / $esp * 100 }
        $col  = Get-ColorPct $pct
        $selz = ($script:ZonaSel -eq $z.Id)
        $colNombre = '#9FB0DA'
        if ($selz) { $colNombre = '#E8EDFA' }

        $alinear = 'Center'
        if ($z.Modo -eq 'L') { $alinear = 'Right' }
        if ($z.Modo -eq 'R') { $alinear = 'Left' }

        if ($z.Modo -eq 'T') {
            Add-Linea -Canvas $cv -X1 $z.AX -Y1 $z.AY -X2 $z.LX -Y2 ($z.LY + 40) -Color ($(if ($selz) { '#00E5FF' } else { '#2A3A6B' })) -Grosor 1.1
        }
        if ($z.Modo -eq 'L') {
            Add-Linea -Canvas $cv -X1 $z.AX -Y1 $z.AY -X2 ($z.LX + 8) -Y2 ($z.LY + 15) -Color ($(if ($selz) { '#00E5FF' } else { '#2A3A6B' })) -Grosor 1.1
        }
        if ($z.Modo -eq 'R') {
            Add-Linea -Canvas $cv -X1 $z.AX -Y1 $z.AY -X2 ($z.LX - 8) -Y2 ($z.LY + 15) -Color ($(if ($selz) { '#00E5FF' } else { '#2A3A6B' })) -Grosor 1.1
        }

        $tbN = Add-Texto -Canvas $cv -Texto $z.Nombre -X $z.LX -Y $z.LY -Tamano 12.5 -Color $colNombre -Alinear $alinear -Negrita $selz
        $tbN.Tag = $z.Id; $tbN.Cursor = [System.Windows.Input.Cursors]::Hand
        $tbN.Add_MouseLeftButtonDown($script:ClickZona)

        $tbP = Add-Texto -Canvas $cv -Texto "$(Fmt $pct 0) %" -X $z.LX -Y ($z.LY + 17) -Tamano 16 -Color $col -Alinear $alinear -Negrita $true
        $tbP.Tag = $z.Id; $tbP.Cursor = [System.Windows.Input.Cursors]::Hand
        $tbP.Add_MouseLeftButtonDown($script:ClickZona)
    }

    $txtNotaCerdo.Text = "Porcentajes calculados con $(Fmt $Cerdos 0) cerdos - toca una zona para ver su detalle"
}

function Draw-TendenciaCuadre {
    param($T)

    $cvTendencia.Children.Clear()
    if ($T.Barras.Count -eq 0) {
        Add-Vacio -Canvas $cvTendencia -Texto "Sin dias con produccion en el rango"
        return
    }

    $ancho = $cvTendencia.Width; $alto = $cvTendencia.Height
    $izq = 50.0; $der = 12.0; $arr = 14.0; $aba = 40.0

    $maxV = 110.0
    foreach ($b in $T.Barras) { if ($b.Pct -gt $maxV) { $maxV = $b.Pct } }
    $maxV = $maxV * 1.06

    for ($g = 0; $g -le 4; $g++) {
        $v = $maxV * $g / 4
        $y = $arr + ($maxV - $v) * ($alto - $arr - $aba) / $maxV
        Add-Linea -Canvas $cvTendencia -X1 $izq -Y1 $y -X2 ($ancho - $der) -Y2 $y -Color "#22305C"
        [void](Add-Texto -Canvas $cvTendencia -Texto "$(Fmt $v 0)%" -X ($izq - 7) -Y ($y - 8) -Tamano 9.5 -Color "#8C9BC4" -Alinear "Right")
    }

    # Linea de referencia en el 100 por ciento
    $y100 = $arr + ($maxV - 100) * ($alto - $arr - $aba) / $maxV
    if ($y100 -ge $arr -and $y100 -le ($alto - $aba)) {
        Add-Linea -Canvas $cvTendencia -X1 $izq -Y1 $y100 -X2 ($ancho - $der) -Y2 $y100 -Color "#00FFA3" -Grosor 1.2 -Punteada $true
    }

    $n = $T.Barras.Count
    $celda = ($ancho - $izq - $der) / $n
    $bw = [math]::Max(6.0, $celda - 8)

    for ($i = 0; $i -lt $n; $i++) {
        $b = $T.Barras[$i]
        $x = $izq + $i * $celda + ($celda - $bw) / 2
        $y = $arr + ($maxV - $b.Pct) * ($alto - $arr - $aba) / $maxV
        $h = ($alto - $aba) - $y
        $col = Get-ColorPct $b.Pct
        $tip = "$($b.Dia)`n$(Fmt $b.Pct 1) %"
        [void](Add-Rect -Canvas $cvTendencia -X $x -Y $y -Ancho $bw -Alto $h -Relleno $col -Radio 2 -Tip $tip)
        [void](Add-Texto -Canvas $cvTendencia -Texto (Fmt $b.Pct 0) -X ($x + $bw / 2) -Y ($y - 13) -Tamano 9 -Color "#C6D4F5" -Alinear "Center")
        $dd = $b.Dia.Substring(8, 2) + "-" + $b.Dia.Substring(5, 2)
        [void](Add-Texto -Canvas $cvTendencia -Texto $dd -X ($x + $bw / 2) -Y ($alto - $aba + 6) -Tamano 9 -Color "#8C9BC4" -Alinear "Center")
    }
}

# ------------------------------------------------------------
#  CUADRE DE PIEZAS - RENDER Y ORQUESTACION
# ------------------------------------------------------------
function Render-Cuadre {
    $cerdos = Get-CerdosDeFecha $script:FechaCuadre
    $r = Get-AggZonas -Materiales $script:MatDia -Cfg $script:CfgZonas

    Draw-Cerdo -AggR $r -Cerdos $cerdos
    if ($null -ne $r -and $null -ne $r.Pan) {
        Draw-Paneles -Cv $cvTarjetas -AggR $r -Cerdos $cerdos
        Draw-Aprovechamiento -Cv $cvAprov -AggR $r
    }

    # Panel derecho: zona seleccionada o cuadre global
    if ($script:ZonaSel) {
        $zdef = $null
        foreach ($z in $ZonasCuadre) { if ($z.Id -eq $script:ZonaSel) { $zdef = $z } }
        $dz  = $r.Agg[$zdef.Id]
        $esp = $cerdos * $zdef.Mult
        $pct = 0.0
        if ($esp -gt 0) { $pct = $dz.Piezas / $esp * 100 }

        $txtZonaTitulo.Text = $zdef.Nombre
        $txtZonaNota.Text   = "$($zdef.Mult) pieza(s) por cerdo  -  $(Fmt $cerdos 0) cerdos"
        $txtZonaPct.Text    = "$(Fmt $pct 1) %"
        $txtZonaPct.Foreground = Pincel (Get-ColorPct $pct)

        $filas = @()
        foreach ($m in @($dz.Mats | Sort-Object Piezas -Descending)) {
            $filas += [pscustomobject]@{
                Material = $m.Producto; Codigo = $m.Codigo
                Piezas = Fmt $m.Piezas 0; Cajas = Fmt $m.Cajas 0
            }
        }
        $gridMat.ItemsSource = $filas

        $lblTot.Text = "Total producido: $(Fmt $dz.Piezas 0) piezas"
        $lblEsp.Text = "Esperado: $(Fmt $esp 0)  ($($zdef.Mult) x $(Fmt $cerdos 0) cerdos)"
        $dif = $esp - $dz.Piezas
        if ($dif -gt 0) {
            $lblDif.Text = "Faltan $(Fmt $dif 0) piezas"
            $lblDif.Foreground = Pincel $C.NeonNaranja
        } elseif ($dif -lt 0) {
            $lblDif.Text = "Sobran $(Fmt ([math]::Abs($dif)) 0) piezas (revisar posible doble conteo)"
            $lblDif.Foreground = Pincel $C.NeonVioleta
        } else {
            $lblDif.Text = "Cuadre exacto"
            $lblDif.Foreground = Pincel $C.NeonVerde
        }
    } else {
        $espGlobal = $cerdos * $SumaMultiplicadores
        $pctGlobal = 0.0
        if ($espGlobal -gt 0) { $pctGlobal = $r.TotalAsignadas / $espGlobal * 100 }

        $txtZonaTitulo.Text = "Cuadre global"
        $txtZonaNota.Text   = "12 zonas  -  $(Fmt $cerdos 0) cerdos  -  $($SumaMultiplicadores) piezas esperadas por cerdo"
        $txtZonaPct.Text    = "$(Fmt $pctGlobal 1) %"
        $txtZonaPct.Foreground = Pincel (Get-ColorPct $pctGlobal)

        $filas = @()
        foreach ($m in @($r.Sin | Sort-Object Piezas -Descending)) {
            $filas += [pscustomobject]@{
                Material = $m.Producto; Codigo = $m.Codigo
                Piezas = Fmt $m.Piezas 0; Cajas = Fmt $m.Cajas 0
            }
        }
        $gridMat.ItemsSource = $filas

        $lblTot.Text = "Piezas asignadas: $(Fmt $r.TotalAsignadas 0)"
        $lblEsp.Text = "Esperado global: $(Fmt $espGlobal 0)"
        if ($r.Sin.Count -gt 0) {
            $pSin = 0
            foreach ($m in $r.Sin) { $pSin = $pSin + $m.Piezas }
            $lblDif.Text = "La tabla muestra $(Fmt $r.Sin.Count 0) materiales sin zona ($(Fmt $pSin 0) piezas)"
            $lblDif.Foreground = Pincel $C.NeonNaranja
        } else {
            $lblDif.Text = "Todos los materiales del dia tienen zona"
            $lblDif.Foreground = Pincel $C.NeonVerde
        }
    }

    # Resumen por zona
    $script:SuprimirCuadre = $true
    $filasZ = @()
    foreach ($z in $ZonasCuadre) {
        $dz  = $r.Agg[$z.Id]
        $esp = $cerdos * $z.Mult
        $pct = 0.0
        if ($esp -gt 0) { $pct = $dz.Piezas / $esp * 100 }
        $filasZ += [pscustomobject]@{
            Id = $z.Id; Zona = $z.Nombre
            Piezas = Fmt $dz.Piezas 0; Esperado = Fmt $esp 0
            Pct = "$(Fmt $pct 1) %"; Estado = Get-EstadoPct $pct
        }
    }
    $gridZonas.ItemsSource = $filasZ
    for ($i = 0; $i -lt $filasZ.Count; $i++) {
        if ($filasZ[$i].Id -eq $script:ZonaSel) { $gridZonas.SelectedIndex = $i }
    }
    if (-not $script:ZonaSel) { $gridZonas.SelectedIndex = -1 }
    $script:SuprimirCuadre = $false

    # Aviso de materiales sin zona
    if ($r.Sin.Count -gt 0) {
        $pSin = 0
        foreach ($m in $r.Sin) { $pSin = $pSin + $m.Piezas }
        $txtAvisoCfg.Text = "Hay $(Fmt $r.Sin.Count 0) materiales con piezas y sin zona asignada ($(Fmt $pSin 0) piezas). Usa 'Configurar zonas' para asignarlos; mientras tanto no suman a ningun cuadre. La lista aparece en el panel derecho al no tener zona seleccionada."
        $panelAvisoCfg.Visibility = "Visible"
    } else {
        $panelAvisoCfg.Visibility = "Collapsed"
    }

    # Tendencia
    $zonaTxt = "global"
    if ($script:ZonaSel) {
        foreach ($z in $ZonasCuadre) { if ($z.Id -eq $script:ZonaSel) { $zonaTxt = $z.Nombre } }
    }
    $t = Compute-Tendencia -Raw $script:TendRaw -CfgZ $script:CfgZonas -CerdosDia $script:CerdosPorDia -Fecha $script:FechaCuadre -ZonaSel ([string]$script:ZonaSel)
    $txtPromedios.Text = "Cuadre $zonaTxt  -  Promedio 14 dias: $(Fmt $t.Prom14 1) %  -  Promedio mes $($t.MesTexto): $(Fmt $t.PromMes 1) %"
    Draw-TendenciaCuadre -T $t
}

function Refrescar-Cuadre {
    param([bool]$Silencioso = $false)

    if ($script:OcupadoCuadre) { return }
    $script:OcupadoCuadre = $true
    try {
        if (-not (Test-Path $RutaBase)) {
            Set-Estado "No se encuentra la base en $RutaBase" $C.NeonRojo
            return
        }
        $f = $dpFechaCuadre.SelectedDate
        if ($null -eq $f) { $f = (Get-Date).Date }
        $script:FechaCuadre = $f.Date

        if (-not $Silencioso) { Set-Estado "Cuadre: consultando la base..." }
        $script:MatDia  = Get-MaterialesDia -Fecha $script:FechaCuadre
        $script:TendRaw = Get-TendenciaCuadre -Fecha $script:FechaCuadre
        $script:HistCuadre = Get-HistorialAprovechamiento -Hasta $script:FechaCuadre -Dias 30 -Cfg $script:CfgZonas
        Render-Cuadre

        $totP = 0
        foreach ($m in $script:MatDia) { $totP = $totP + $m.Piezas }
        if ($script:MatDia.Count -eq 0) {
            Set-Estado "Cuadre: sin cajas con piezas el $($script:FechaCuadre.ToString('dd-MM-yyyy'))." $C.NeonNaranja
        } else {
            Set-Estado "Cuadre: $(Fmt $script:MatDia.Count 0) materiales y $(Fmt $totP 0) piezas el $($script:FechaCuadre.ToString('dd-MM-yyyy'))." $C.NeonVerde
        }
    }
    catch {
        Set-Estado "Error cuadre: $($_.Exception.Message)  (linea $($_.InvocationInfo.ScriptLineNumber))" $C.NeonRojo
    }
    finally {
        $script:OcupadoCuadre = $false
    }
}

function Export-Cuadre {
    param([string]$Carpeta = '')
    try {
        if ($script:MatDia.Count -eq 0) {
            Set-Estado "Cuadre: no hay datos que exportar para el dia elegido." $C.NeonNaranja
            return
        }
        if ($Carpeta -eq '') {
            $Carpeta = Join-Path ([Environment]::GetFolderPath('Desktop')) 'CUADRE_PIEZAS'
        }
        New-Item -ItemType Directory -Path $Carpeta -Force | Out-Null
        $cerdos = Get-CerdosDeFecha $script:FechaCuadre
        $csv = Build-CuadreCsv -Materiales $script:MatDia -Cfg $script:CfgZonas -Cerdos $cerdos -Fecha $script:FechaCuadre
        $ruta = Join-Path $Carpeta ("Cuadre_Piezas_" + $script:FechaCuadre.ToString('yyyy-MM-dd') + "_" + (Get-Date).ToString('HHmmss') + ".csv")
        [System.IO.File]::WriteAllText($ruta, $csv, (New-Object System.Text.UTF8Encoding $true))
        Set-Estado "Cuadre exportado: $ruta" $C.NeonVerde
        return $ruta
    }
    catch {
        Set-Estado "Error al exportar: $($_.Exception.Message)" $C.NeonRojo
    }
}

function Show-ConfigZonas {
    try {
        $fDesde = Format-FechaAccess (Get-Date).Date.AddDays(-60)
        $sql = @"
SELECT CodigoProducto, ProductoEspanol, SUM(Piezas) AS Piezas
FROM Cajas
WHERE FechaPesaje >= $fDesde
  AND Usuario <> '$UsuarioExcluido'
  AND NombreCliente = '$Cliente'
GROUP BY CodigoProducto, ProductoEspanol
"@
        $t = Invoke-ConsultaAccess -Sql $sql
    }
    catch {
        Set-Estado "No se pudo leer la lista de materiales: $($_.Exception.Message)" $C.NeonRojo
        return
    }

    $nombrePorId = @{}
    $idPorNombre = @{}
    foreach ($z in $ZonasCuadre) { $nombrePorId[$z.Id] = $z.Nombre; $idPorNombre[$z.Nombre] = $z.Id }
    # Los paneles tambien son destino valido, con prefijo para distinguirlos
    foreach ($p in $PanelesCuadre) {
        $etq = 'Panel: ' + $p.Nombre
        $nombrePorId[$p.Id] = $etq; $idPorNombre[$etq] = $p.Id
    }
    $opciones = New-Object System.Collections.Generic.List[string]
    [void]$opciones.Add('Sin asignar')
    [void]$opciones.Add('Excluido')
    foreach ($z in $ZonasCuadre) { [void]$opciones.Add($z.Nombre) }
    foreach ($p in $PanelesCuadre) { [void]$opciones.Add('Panel: ' + $p.Nombre) }

    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add("Codigo",   [int])
    [void]$dt.Columns.Add("Producto", [string])
    [void]$dt.Columns.Add("Piezas60", [int])
    [void]$dt.Columns.Add("Zona",     [string])
    foreach ($f in ($t.Rows | Sort-Object { [string]$_["ProductoEspanol"] })) {
        $cod = [int]$f["CodigoProducto"]
        $zid = $script:CfgZonas[[string]$cod]
        $ztxt = 'Sin asignar'
        if ($zid -eq 'excluido') { $ztxt = 'Excluido' }
        elseif ($zid -and $nombrePorId.ContainsKey($zid)) { $ztxt = $nombrePorId[$zid] }
        $fila = $dt.NewRow()
        $fila["Codigo"] = $cod
        $fila["Producto"] = ([string]$f["ProductoEspanol"]).Trim()
        $fila["Piezas60"] = [int][math]::Round([double]$f["Piezas"], 0)
        $fila["Zona"] = $ztxt
        $dt.Rows.Add($fila)
    }

    $cfgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Configurar zonas del cuadre" Height="620" Width="820"
        WindowStartupLocation="CenterOwner" Background="#070B18" FontFamily="Segoe UI">
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0" Margin="0,0,0,10">
      <TextBlock Text="Asignacion de materiales a zonas" FontSize="15" FontWeight="SemiBold" Foreground="#E8EDFA"/>
      <TextBlock Text="Materiales con piezas en los ultimos 60 dias (cliente AASA). Elige la zona de cada material; 'Excluido' lo deja fuera del cuadre a proposito." FontSize="11.5" Foreground="#8C9BC4" TextWrapping="Wrap" Margin="0,4,0,0"/>
    </StackPanel>
    <DataGrid x:Name="gridCfg" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False"
              Background="#0E1530" BorderBrush="#22305C" BorderThickness="1" RowBackground="#0E1530"
              AlternatingRowBackground="#12193A" Foreground="#E8EDFA" GridLinesVisibility="None"
              HeadersVisibility="Column" FontSize="12" RowHeight="27"/>
    <Grid Grid.Row="2" Margin="0,12,0,0">
      <TextBlock x:Name="txtCfgInfo" Text="" FontSize="11.5" Foreground="#8C9BC4" VerticalAlignment="Center" HorizontalAlignment="Left"/>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="btnCerrarCfg" Content="Cancelar" Width="110" Height="30" Margin="0,0,10,0"
                Background="#0A1128" Foreground="#8C9BC4" BorderBrush="#22305C"/>
        <Button x:Name="btnGuardarCfg" Content="Guardar" Width="130" Height="30"
                Background="#1B3F94" Foreground="#E8EDFA" BorderBrush="#3D8BFF"/>
      </StackPanel>
    </Grid>
  </Grid>
</Window>
"@
    $lector2 = New-Object System.Xml.XmlNodeReader ([xml]$cfgXaml)
    $vCfg = [Windows.Markup.XamlReader]::Load($lector2)
    $vCfg.Owner = $ventana
    $gridCfg      = $vCfg.FindName("gridCfg")
    $btnGuardarCfg= $vCfg.FindName("btnGuardarCfg")
    $btnCerrarCfg = $vCfg.FindName("btnCerrarCfg")
    $txtCfgInfo   = $vCfg.FindName("txtCfgInfo")

    $colCod = New-Object System.Windows.Controls.DataGridTextColumn
    $colCod.Header = "Codigo"; $colCod.IsReadOnly = $true
    $colCod.Binding = New-Object System.Windows.Data.Binding "Codigo"
    $colCod.Width = New-Object System.Windows.Controls.DataGridLength 76
    [void]$gridCfg.Columns.Add($colCod)

    $colProd = New-Object System.Windows.Controls.DataGridTextColumn
    $colProd.Header = "Material"; $colProd.IsReadOnly = $true
    $colProd.Binding = New-Object System.Windows.Data.Binding "Producto"
    $colProd.Width = New-Object System.Windows.Controls.DataGridLength(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
    [void]$gridCfg.Columns.Add($colProd)

    $colP60 = New-Object System.Windows.Controls.DataGridTextColumn
    $colP60.Header = "Piezas 60d"; $colP60.IsReadOnly = $true
    $colP60.Binding = New-Object System.Windows.Data.Binding "Piezas60"
    $colP60.Width = New-Object System.Windows.Controls.DataGridLength 86
    [void]$gridCfg.Columns.Add($colP60)

    $colZona = New-Object System.Windows.Controls.DataGridComboBoxColumn
    $colZona.Header = "Zona"
    $colZona.ItemsSource = $opciones
    $bz = New-Object System.Windows.Data.Binding "Zona"
    $bz.Mode = [System.Windows.Data.BindingMode]::TwoWay
    $colZona.SelectedItemBinding = $bz
    $colZona.Width = New-Object System.Windows.Controls.DataGridLength 190
    [void]$gridCfg.Columns.Add($colZona)

    $gridCfg.ItemsSource = $dt.DefaultView
    $txtCfgInfo.Text = "$($dt.Rows.Count) materiales"

    $btnCerrarCfg.Add_Click({ $vCfg.Close() }.GetNewClosure())

    # Referencia tomada FUERA del closure a proposito (ver nota de arriba)
    $cfgRef = $script:CfgZonas
    $btnGuardarCfg.Add_Click({
        # Todo el manejador va dentro de try/catch: WPF se traga las
        # excepciones de los eventos y los fallos quedaban invisibles.
        try {
        # Primero la celda y despues la fila: al reves el combo no alcanza
        # a entregar el valor y se guardaba la zona anterior.
        [void]$gridCfg.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        [void]$gridCfg.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
        # Red de seguridad: cierra cualquier edicion que haya quedado abierta.
        foreach ($fila in $dt.Rows) {
            if ($fila.HasVersion([System.Data.DataRowVersion]::Proposed)) { $fila.EndEdit() }
        }
        $cambios = 0
        foreach ($fila in $dt.Rows) {
            $cod = [string][int]$fila["Codigo"]
            $ztxt = [string]$fila["Zona"]
            $antes = $cfgRef[$cod]
            $nuevo = $null
            if ($ztxt -eq 'Sin asignar')          { $nuevo = 'sinasignar' }
            elseif ($ztxt -eq 'Excluido')         { $nuevo = 'excluido' }
            elseif ($idPorNombre.ContainsKey($ztxt)) { $nuevo = $idPorNombre[$ztxt] }
            if ($null -ne $nuevo) {
                if ($antes -ne $nuevo) { $cambios = $cambios + 1 }
                $cfgRef[$cod] = $nuevo
            }
        }
        try {
            Write-MapaJson $RutaConfigZonas $cfgRef
            $vCfg.Close()
            Refrescar-Cuadre
            if ($cambios -gt 0) { Set-Estado "Guardado: $cambios material(es) cambiaron de zona." $C.NeonVerde }
            else { Set-Estado "Guardado sin cambios." $C.Texto2 }
        } catch {
            $txtCfgInfo.Text = "No se pudo guardar: $($_.Exception.Message)"
        }
        } catch {
            $txtCfgInfo.Text = "Error al guardar: $($_.Exception.Message)"
        }
    }.GetNewClosure())

    [void]$vCfg.ShowDialog()
}


# ------------------------------------------------------------
#  ORQUESTACION
# ------------------------------------------------------------
function Refrescar {
    param([bool]$Silencioso = $false)

    if ($script:Ocupado) { return }
    $script:Ocupado = $true
    try {
        if (-not (Test-Path $RutaBase)) {
            Set-Estado "No se encuentra la base en $RutaBase" $C.NeonRojo
            return
        }

        $desde = $dpDesde.SelectedDate
        $hasta = $dpHasta.SelectedDate
        if ($null -eq $desde) { $desde = (Get-Date).Date }
        if ($null -eq $hasta) { $hasta = (Get-Date).Date }
        if ($hasta -lt $desde) { Set-Estado "La fecha Hasta no puede ser anterior a Desde." $C.NeonRojo; return }

        if (-not $Silencioso) { Set-Estado "Consultando la base..." }
        $script:Hechos = Get-DatosGiveaway -Desde $desde -Hasta $hasta

        $codigo = 0
        if ($cbProducto.SelectedIndex -gt 0) { $codigo = $ProductosGiveaway[$cbProducto.SelectedIndex - 1].Codigo }

        $resumen = Get-Resumen -Hechos $script:Hechos

        Draw-Tarjetas   -Resumen $resumen
        Draw-Barras     -Resumen $resumen
        Draw-Linea      -Hechos  $script:Hechos
        Draw-Histograma -Hechos  $script:Hechos -CodigoElegido $codigo
        Draw-Mapa       -Hechos  $script:Hechos
        Draw-Tabla      -Resumen $resumen

        $txtActualizado.Text = "Actualizado " + (Get-Date).ToString("HH:mm:ss")
        $totalCajas = 0
        foreach ($r in $resumen) { $totalCajas = $totalCajas + $r.Cajas }
        if ($totalCajas -eq 0) {
            Set-Estado "No hay cajas de peso fijo en el rango seleccionado." $C.NeonNaranja
        } else {
            Set-Estado "$(Fmt $totalCajas 0) cajas de peso fijo en el rango. Proxima lectura automatica en $SegundosRefresco s." $C.NeonVerde
        }
    }
    catch {
        Set-Estado "Error: $($_.Exception.Message)  (linea $($_.InvocationInfo.ScriptLineNumber))" $C.NeonRojo
    }
    finally {
        $script:Ocupado = $false
    }
}

function Redibujar {
    if ($script:Hechos.Count -eq 0) { return }
    $codigo = 0
    if ($cbProducto.SelectedIndex -gt 0) { $codigo = $ProductosGiveaway[$cbProducto.SelectedIndex - 1].Codigo }
    Draw-Histograma -Hechos $script:Hechos -CodigoElegido $codigo
}

# ------------------------------------------------------------
#  EVENTOS
# ------------------------------------------------------------
$script:Suprimir = $true
[void]$cbProducto.Items.Add("Todos los productos")
foreach ($p in $ProductosGiveaway) { [void]$cbProducto.Items.Add("$($p.Nombre)  [$($p.Codigo)]") }
$cbProducto.SelectedIndex = 0
$dpDesde.SelectedDate = (Get-Date).Date.AddDays(-29)
$dpHasta.SelectedDate = (Get-Date).Date
$script:Suprimir = $false

function Set-Rango {
    param([int]$Dias)
    $script:Suprimir = $true
    $dpDesde.SelectedDate = (Get-Date).Date.AddDays(-$Dias)
    $dpHasta.SelectedDate = (Get-Date).Date
    $script:Suprimir = $false
    Refrescar
}

$dpDesde.Add_SelectedDateChanged({ if (-not $script:Suprimir) { Refrescar } })
$dpHasta.Add_SelectedDateChanged({ if (-not $script:Suprimir) { Refrescar } })
$cbProducto.Add_SelectionChanged({ if (-not $script:Suprimir) { Redibujar } })
$btnRefrescar.Add_Click({ Refrescar })

$btnHoy.Add_Click({ Set-Rango 0 })
$btn7.Add_Click({   Set-Rango 6 })
$btn30.Add_Click({  Set-Rango 29 })
$btn90.Add_Click({  Set-Rango 89 })

# ------------------------------------------------------------
#  EVENTOS - CUADRE DE PIEZAS
# ------------------------------------------------------------
function Initialize-Cuadre {
    Initialize-ConfigZonas
    $script:CerdosPorDia = Read-MapaJson $RutaCerdosDia
    $script:PesoVara = Get-PesoVara
    $script:HistCuadre = @()
    $script:SuprimirCuadre = $true
    $txtVara.Text = [string]$script:PesoVara
    $script:SuprimirCuadre = $false
    $script:SuprimirCuadre = $true
    $dpFechaCuadre.SelectedDate = (Get-Date).Date
    $script:FechaCuadre = (Get-Date).Date
    $txtCerdos.Text = [string](Get-CerdosDeFecha $script:FechaCuadre)
    $script:SuprimirCuadre = $false
}

$dpFechaCuadre.Add_SelectedDateChanged({
    if ($script:SuprimirCuadre) { return }
    $f = $dpFechaCuadre.SelectedDate
    if ($null -eq $f) { return }
    $script:FechaCuadre = $f.Date
    $script:SuprimirCuadre = $true
    $txtCerdos.Text = [string](Get-CerdosDeFecha $f.Date)
    $script:SuprimirCuadre = $false
    $txtCerdosHint.Text = ""
    Refrescar-Cuadre
})

$txtVara.Add_TextChanged({
    if ($script:SuprimirCuadre) { return }
    $v = 0.0
    $txt = ([string]$txtVara.Text).Replace(',', '.')
    if ([double]::TryParse($txt, [System.Globalization.NumberStyles]::Any,
                           [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) -and $v -ge 40 -and $v -le 200) {
        Set-PesoVara $v
        Refrescar-Cuadre
    }
})

$txtCerdos.Add_TextChanged({
    if ($script:SuprimirCuadre) { return }
    $n = 0
    if ([int]::TryParse($txtCerdos.Text, [ref]$n) -and $n -ge 1 -and $n -le 20000) {
        $k = $script:FechaCuadre.ToString('yyyy-MM-dd')
        $script:CerdosPorDia[$k] = $n
        try {
            Write-MapaJson $RutaCerdosDia $script:CerdosPorDia
            $txtCerdosHint.Text = "guardado para este dia"
        } catch {
            $txtCerdosHint.Text = "no se pudo guardar"
        }
        Render-Cuadre
    } else {
        $txtCerdosHint.Text = ""
    }
})

$btnRefCuadre.Add_Click({ Refrescar-Cuadre })
$btnExportCuadre.Add_Click({ [void](Export-Cuadre) })
$btnCfgZonas.Add_Click({ Show-ConfigZonas })

$gridZonas.Add_SelectionChanged({
    if ($script:SuprimirCuadre) { return }
    $it = $gridZonas.SelectedItem
    if ($null -ne $it) {
        $script:ZonaSel = $it.Id
        Render-Cuadre
    }
})


# ------------------------------------------------------------
#  BUSCAR PRODUCTO - INTERFAZ
# ------------------------------------------------------------
$script:BusCatalogo = @()   # catalogo completo del rango
$script:BusElegidos = @()   # productos que el usuario acumulo
$script:BusOcupado  = $false

function Set-BusConteo {
    $n = $script:BusElegidos.Count
    if ($n -eq 1) { $txtBusConteo.Text = "1 producto elegido" }
    else { $txtBusConteo.Text = "$n productos elegidos" }
}

function Fill-BusCatalogo {
    # Vuelve a llenar la lista aplicando el texto del filtro.
    $f = ([string]$txtBusFiltro.Text).Trim().ToLower()
    $lstBusCatalogo.Items.Clear()
    $n = 0
    foreach ($p in $script:BusCatalogo) {
        if ($f -ne "") {
            $hay = ($p.Producto.ToLower().Contains($f)) -or ([string]$p.Codigo).Contains($f)
            if (-not $hay) { continue }
        }
        $it = New-Object System.Windows.Controls.ListBoxItem
        $it.Content = $p.Texto
        $it.Tag = $p
        [void]$lstBusCatalogo.Items.Add($it)
        $n++
    }
    if ($script:BusCatalogo.Count -eq 0) {
        $txtBusHintFiltro.Text = "sin datos en el rango elegido"
    } elseif ($f -eq "") {
        $txtBusHintFiltro.Text = "$n materiales en el rango"
    } else {
        $txtBusHintFiltro.Text = "$n coinciden con el filtro"
    }
}

function Fill-BusElegidos {
    $lstBusElegidos.Items.Clear()
    foreach ($p in $script:BusElegidos) {
        $it = New-Object System.Windows.Controls.ListBoxItem
        $it.Content = $p.Texto
        $it.Tag = $p
        [void]$lstBusElegidos.Items.Add($it)
    }
    Set-BusConteo
}

function Add-BusSeleccion {
    param([object[]]$Items)
    $lista = New-Object System.Collections.Generic.List[object]
    foreach ($e in $script:BusElegidos) { [void]$lista.Add($e) }
    foreach ($it in $Items) {
        $p = $it.Tag
        $rep = $false
        foreach ($e in $lista) {
            if ($e.Codigo -eq $p.Codigo -and $e.Producto -eq $p.Producto) { $rep = $true; break }
        }
        if (-not $rep) { [void]$lista.Add($p) }
    }
    $script:BusElegidos = $lista.ToArray()
    Fill-BusElegidos
}

function Cargar-BusCatalogo {
    # Se relee el catalogo cada vez que cambia el rango de fechas.
    if ($null -eq $dpBusDesde.SelectedDate -or $null -eq $dpBusHasta.SelectedDate) { return }
    $d = [datetime]$dpBusDesde.SelectedDate
    $h = [datetime]$dpBusHasta.SelectedDate
    if ($d -gt $h) { return }
    try {
        $script:BusCatalogo = Get-CatalogoProductos -Desde $d -Hasta $h
        Fill-BusCatalogo
    } catch {
        $script:BusCatalogo = @()
        $lstBusCatalogo.Items.Clear()
        $txtBusHintFiltro.Text = "no se pudo leer el catalogo"
    }
}

function Draw-BusTarjetas {
    # Recibe UNA fila del resumen. Si es $null, deja las tarjetas en blanco.
    param($Fila)
    if ($null -eq $Fila) {
        foreach ($t in @($bCajas, $bKilos, $bProm, $bMin, $bMax, $bDesv, $bKgPz)) { $t.Text = "-" }
        $txtBusProdSel.Text = "elige un producto"
        $bCajasSub.Text = "en el rango elegido"
        return
    }
    $script:BusSel = $Fila
    $txtBusProdSel.Text = "$($Fila.Producto)  [$($Fila.Codigo)]"
    $bCajas.Text = Fmt ([double]$Fila.Cajas) 0
    $bKilos.Text = $Fila.KgTotal + " kg"
    $bProm.Text  = $Fila.Promedio + " kg"
    $bMin.Text   = $Fila.Minimo + " kg"
    $bMax.Text   = $Fila.Maximo + " kg"
    $bDesv.Text  = $Fila.Desv + " kg"
    $bCajasSub.Text = "cajas de este producto"
    $bKilosSub.Text = "" + $Fila.Piezas + " piezas en total"
    $bKgPz.Text = $Fila.KgPorPz
    if ($Fila.KgPorPz -eq "-") {
        $bKgPzSub.Text = "sin piezas declaradas"
    } else {
        $bKgPz.Text = $Fila.KgPorPz + " kg"
        $bKgPzSub.Text = "promedio de " + $Fila.Piezas + " piezas"
    }
}

function Limpiar-BusResultados {
    foreach ($t in @($bCajas, $bKilos, $bProm, $bMin, $bMax, $bDesv)) { $t.Text = "-" }
    $gridBusResumen.ItemsSource = $null
    $gridBusDetalle.ItemsSource = $null
    $txtBusProdSel.Text = "elige un producto"
    $script:BusSel = $null
    $txtBusTope.Text = ""
    $script:BusDet = @()
    $script:BusRes = @()
}

function Buscar-Producto {
    if ($script:BusOcupado) { return }
    if ($null -eq $dpBusDesde.SelectedDate -or $null -eq $dpBusHasta.SelectedDate) {
        Set-Estado "Elige el rango de fechas." $C.NeonNaranja
        return
    }
    $d = [datetime]$dpBusDesde.SelectedDate
    $h = [datetime]$dpBusHasta.SelectedDate
    if ($d -gt $h) {
        Set-Estado "La fecha Desde no puede ser mayor que Hasta." $C.NeonRojo
        return
    }
    if ($script:BusElegidos.Count -eq 0) {
        Set-Estado "Elige al menos un producto." $C.NeonNaranja
        return
    }

    $script:BusOcupado = $true
    Set-Estado "Buscando..." $C.NeonCyan
    try {
        $det = Get-DetalleCajas -Desde $d -Hasta $h -Productos $script:BusElegidos
        $res = Get-ResumenBusqueda -Detalle $det
        $script:BusDet = $det
        $script:BusRes = $res

        $gridBusResumen.ItemsSource = @($res)
        $gridBusDetalle.ItemsSource = @($det)

        if ($det.Count -eq 0) {
            Draw-BusTarjetas $null
            $txtBusNotaResumen.Text = "sin cajas para esos productos en el rango"
            $txtBusNotaDetalle.Text = "sin resultados"
            $txtBusTope.Text = ""
            Set-Estado "Sin resultados." $C.NeonNaranja
        } else {
            # Las tarjetas muestran el producto de mayor volumen; el resto
            # se ve haciendo clic en su fila del resumen.
            $arr = @($res)
            Draw-BusTarjetas $arr[0]
            $script:SuprimirBus = $true
            $gridBusResumen.SelectedIndex = 0
            $script:SuprimirBus = $false

            $nprod = $arr.Count
            $txtBusNotaResumen.Text = "$nprod material(es) - clic en una fila para ver su estadistica arriba"
            $txtBusNotaDetalle.Text = "" + (Fmt ([double]$det.Count) 0) + " caja(s) etiquetada(s)"

            if ($det.Count -ge $TopeDetalle) {
                $txtBusTope.Text = "Se muestran las primeras $TopeDetalle cajas. Acota el rango o los productos para ver el resto."
            } else {
                $txtBusTope.Text = ""
            }
            Set-Estado ("Listo: " + (Fmt ([double]$det.Count) 0) + " cajas.") $C.NeonVerde
        }
    } catch {
        Set-Estado ("Error al buscar: " + $_.Exception.Message) $C.NeonRojo
    } finally {
        $script:BusOcupado = $false
    }
}

function Export-Busqueda {
    if ($script:BusDet.Count -eq 0) {
        Set-Estado "No hay resultados para exportar." $C.NeonNaranja
        return $null
    }
    try {
        $carpeta = Join-Path ([Environment]::GetFolderPath('Desktop')) 'BUSCAR_PRODUCTO'
        if (-not (Test-Path $carpeta)) { [void](New-Item -ItemType Directory -Path $carpeta -Force) }
        $nombre = "Busqueda_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + ".csv"
        $ruta = Join-Path $carpeta $nombre
        $csv = Build-BusquedaCsv -Resumen $script:BusRes -Detalle $script:BusDet `
                                 -Desde ([datetime]$dpBusDesde.SelectedDate) `
                                 -Hasta ([datetime]$dpBusHasta.SelectedDate)
        # BOM para que Excel respete los acentos y el punto y coma
        $enc = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($ruta, $csv, $enc)
        Set-Estado "Exportado a $ruta" $C.NeonVerde
        return $ruta
    } catch {
        Set-Estado ("No se pudo exportar: " + $_.Exception.Message) $C.NeonRojo
        return $null
    }
}

function Initialize-Buscar {
    $script:SuprimirBus = $true
    $dpBusHasta.SelectedDate = (Get-Date).Date
    $dpBusDesde.SelectedDate = (Get-Date).Date.AddDays(-6)
    $script:SuprimirBus = $false
    Limpiar-BusResultados
    Cargar-BusCatalogo
}

# ---- Eventos ----
$txtBusFiltro.Add_TextChanged({ Fill-BusCatalogo })

$dpBusDesde.Add_SelectedDateChanged({ if (-not $script:SuprimirBus) { Cargar-BusCatalogo } })
$dpBusHasta.Add_SelectedDateChanged({ if (-not $script:SuprimirBus) { Cargar-BusCatalogo } })

$btnBusHoy.Add_Click({
    $script:SuprimirBus = $true
    $dpBusDesde.SelectedDate = (Get-Date).Date
    $dpBusHasta.SelectedDate = (Get-Date).Date
    $script:SuprimirBus = $false
    Cargar-BusCatalogo
})
$btnBusSemana.Add_Click({
    $script:SuprimirBus = $true
    $dpBusHasta.SelectedDate = (Get-Date).Date
    $dpBusDesde.SelectedDate = (Get-Date).Date.AddDays(-6)
    $script:SuprimirBus = $false
    Cargar-BusCatalogo
})
$btnBusMes.Add_Click({
    $script:SuprimirBus = $true
    $dpBusHasta.SelectedDate = (Get-Date).Date
    $dpBusDesde.SelectedDate = (Get-Date).Date.AddDays(-29)
    $script:SuprimirBus = $false
    Cargar-BusCatalogo
})

$btnBusAgregar.Add_Click({
    if ($lstBusCatalogo.SelectedItems.Count -eq 0) {
        Set-Estado "Marca uno o mas productos de la lista." $C.NeonNaranja
        return
    }
    $sel = @()
    foreach ($i in $lstBusCatalogo.SelectedItems) { $sel += $i }
    Add-BusSeleccion -Items $sel
})

$btnBusTodos.Add_Click({
    $todos = @()
    foreach ($i in $lstBusCatalogo.Items) { $todos += $i }
    if ($todos.Count -eq 0) { return }
    Add-BusSeleccion -Items $todos
})

# Doble clic en el catalogo tambien agrega
$lstBusCatalogo.Add_MouseDoubleClick({
    if ($null -ne $lstBusCatalogo.SelectedItem) {
        Add-BusSeleccion -Items @($lstBusCatalogo.SelectedItem)
    }
})

$btnBusQuitar.Add_Click({
    if ($lstBusElegidos.SelectedItems.Count -eq 0) { return }
    $quitar = New-Object System.Collections.Generic.List[object]
    foreach ($i in $lstBusElegidos.SelectedItems) { [void]$quitar.Add($i.Tag) }
    $lista = New-Object System.Collections.Generic.List[object]
    foreach ($e in $script:BusElegidos) {
        $fuera = $false
        foreach ($q in $quitar) {
            if ($q.Codigo -eq $e.Codigo -and $q.Producto -eq $e.Producto) { $fuera = $true; break }
        }
        if (-not $fuera) { [void]$lista.Add($e) }
    }
    $script:BusElegidos = $lista.ToArray()
    Fill-BusElegidos
})

$btnBusLimpiar.Add_Click({
    $script:BusElegidos = @()
    Fill-BusElegidos
    Limpiar-BusResultados
    Set-Estado "Seleccion limpiada." $C.Texto2
})

$gridBusResumen.Add_SelectionChanged({
    if ($script:SuprimirBus) { return }
    $it = $gridBusResumen.SelectedItem
    if ($null -ne $it) { Draw-BusTarjetas $it }
})

$btnBusBuscar.Add_Click({ Buscar-Producto })
$btnBusExport.Add_Click({ [void](Export-Busqueda) })



function Show-TiposCambio {
    $xamlTC = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Tipos de cambio" Height="330" Width="430"
        WindowStartupLocation="CenterOwner" Background="#070B18"
        FontFamily="Segoe UI" ResizeMode="NoResize">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0" Margin="0,0,0,16">
      <TextBlock Text="Cuanto vale 1 unidad en pesos chilenos" FontSize="13.5"
                 FontWeight="SemiBold" Foreground="#E8EDFA"/>
      <TextBlock Text="Se usan para valorizar los kilos regalados. Quedan guardados." 
                 FontSize="11" Foreground="#8C9BC4" Margin="0,4,0,0" TextWrapping="Wrap"/>
    </StackPanel>

    <StackPanel Grid.Row="1">
      <Grid Margin="0,0,0,11">
        <Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="USD (dolar)" Foreground="#8C9BC4" FontSize="12.5" VerticalAlignment="Center"/>
        <TextBox Grid.Column="1" x:Name="txtUSD" Height="29" Background="#0A1128" Foreground="#E8EDFA"
                 BorderBrush="#22305C" FontSize="13" VerticalContentAlignment="Center" Padding="7,0"/>
      </Grid>
      <Grid Margin="0,0,0,11">
        <Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="EUR (euro)" Foreground="#8C9BC4" FontSize="12.5" VerticalAlignment="Center"/>
        <TextBox Grid.Column="1" x:Name="txtEUR" Height="29" Background="#0A1128" Foreground="#E8EDFA"
                 BorderBrush="#22305C" FontSize="13" VerticalContentAlignment="Center" Padding="7,0"/>
      </Grid>
      <Grid Margin="0,0,0,11">
        <Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="JPY (yen)" Foreground="#8C9BC4" FontSize="12.5" VerticalAlignment="Center"/>
        <TextBox Grid.Column="1" x:Name="txtJPY" Height="29" Background="#0A1128" Foreground="#E8EDFA"
                 BorderBrush="#22305C" FontSize="13" VerticalContentAlignment="Center" Padding="7,0"/>
      </Grid>
      <TextBlock x:Name="txtAvisoTC" Text="" FontSize="11" Foreground="#FFA829" TextWrapping="Wrap" Margin="0,4,0,0"/>
    </StackPanel>

    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="btnCancelarTC" Content="Cancelar" Width="96" Height="31" Margin="0,0,9,0"
              Background="#0A1128" Foreground="#8C9BC4" BorderBrush="#22305C" FontSize="12.5"/>
      <Button x:Name="btnGuardarTC" Content="Guardar" Width="96" Height="31"
              Background="#1B3F94" Foreground="#E8EDFA" BorderBrush="#3D8BFF" FontSize="12.5"/>
    </StackPanel>
  </Grid>
</Window>
"@
    $lectorTC = New-Object System.Xml.XmlNodeReader ([xml]$xamlTC)
    $vTC = [Windows.Markup.XamlReader]::Load($lectorTC)
    $vTC.Owner = $ventana

    $txtUSD = $vTC.FindName("txtUSD")
    $txtEUR = $vTC.FindName("txtEUR")
    $txtJPY = $vTC.FindName("txtJPY")
    $txtAvisoTC   = $vTC.FindName("txtAvisoTC")
    $btnGuardarTC = $vTC.FindName("btnGuardarTC")
    $btnCancelarTC= $vTC.FindName("btnCancelarTC")

    $txtUSD.Text = [string]$script:TiposCambio['USD']
    $txtEUR.Text = [string]$script:TiposCambio['EUR']
    $txtJPY.Text = [string]$script:TiposCambio['JPY']

    $btnCancelarTC.Add_Click({ $vTC.Close() }.GetNewClosure())

    $btnGuardarTC.Add_Click({
        $nuevos = @{}
        $malo = ''
        foreach ($par in @(@('USD', $txtUSD), @('EUR', $txtEUR), @('JPY', $txtJPY))) {
            $v = 0.0
            $txt = ([string]$par[1].Text).Replace(',', '.')
            if ([double]::TryParse($txt, [System.Globalization.NumberStyles]::Any,
                                   [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v) -and $v -gt 0) {
                $nuevos[$par[0]] = $v
            } else { $malo = $par[0] }
        }
        if ($malo -ne '') {
            $txtAvisoTC.Text = "Revisa el valor de $malo : debe ser un numero mayor que cero."
            return
        }
        # Global en vez de script: por el mismo motivo del closure
        $Global:TiposCambioApp = $nuevos
        Set-TiposCambio $nuevos
        try {
            Write-TiposCambio $nuevos
            $vTC.Close()
            Refrescar
        } catch {
            $txtAvisoTC.Text = "No se pudo guardar: " + $_.Exception.Message
        }
    }.GetNewClosure())

    [void]$vTC.ShowDialog()
}


$btnTiposCambio.Add_Click({ Show-TiposCambio })

# Temporizador: relee la base sola, igual que el patron de KPI_Produccion_App
$temporizador = New-Object System.Windows.Threading.DispatcherTimer
$temporizador.Interval = [TimeSpan]::FromSeconds($SegundosRefresco)
$temporizador.Add_Tick({
    if ($chkAuto.IsChecked -eq $true) {
        Refrescar -Silencioso $true
        if ($script:FechaCuadre -eq (Get-Date).Date) {
            Refrescar-Cuadre -Silencioso $true
        }
    }
})

$ventana.Add_ContentRendered({
    Refrescar
    Initialize-Cuadre
    Refrescar-Cuadre
    Initialize-Buscar
    $temporizador.Start()
})

$ventana.Add_Closed({ $temporizador.Stop() })

[void]$ventana.ShowDialog()
