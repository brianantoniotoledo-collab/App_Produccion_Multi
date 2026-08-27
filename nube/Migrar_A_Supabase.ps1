<#
Migra Base_Produccion.accdb hacia Supabase (Postgres) usando la API REST que
Supabase genera automaticamente (PostgREST) - sin necesidad de instalar
ningun driver de Postgres.

Seguro de re-ejecutar: cajas/despachos/config_app usan upsert (nunca
duplican), estandares_peso/inventario_sap se vacian y recargan completas
(son "fotos", igual que en la app original), y las tablas de ingreso manual
(pedidos, pedidos_combos, programa_semanal) se saltan si ya tienen datos.

Uso:
  .\Migrar_A_Supabase.ps1
  (lee SupabaseUrl y SupabaseKey desde nube\conexion.txt, ver conexion.ejemplo.txt)

  .\Migrar_A_Supabase.ps1 -SupabaseUrl "https://xxxx.supabase.co" -SupabaseKey "eyJ..."
  (o se pasan directo como parametros, sin usar el archivo)

La SupabaseKey es la "service_role" (Project Settings -> API). Es secreta:
nunca se guarda en este script ni se sube al repositorio. "conexion.txt"
esta en .gitignore para que se quede solo en tu disco.
#>

param(
    [string]$RutaAccess = 'C:\Produccion\Base_Produccion.accdb',
    [string]$SupabaseUrl,
    [string]$SupabaseKey,
    [int]$TamanoLote = 500,
    # Limita Cajas a lo producido desde esta fecha (formato yyyy-MM-dd).
    # Sirve para no pasarse del limite de 500 MB del plan gratis de Supabase:
    # el historico completo (~2M filas) no cabe. Vacio = migrar todo.
    [string]$DesdeFecha
)

if (-not $SupabaseUrl -or -not $SupabaseKey) {
    $archivoConexion = Join-Path $PSScriptRoot 'conexion.txt'
    if (-not (Test-Path $archivoConexion)) {
        throw "Faltan -SupabaseUrl/-SupabaseKey y no existe $archivoConexion. Copia conexion.ejemplo.txt como conexion.txt y completa tus datos, o pasa los parametros directamente."
    }
    $valores = @{}
    Get-Content $archivoConexion | Where-Object { $_ -match '=' -and $_ -notmatch '^\s*#' } | ForEach-Object {
        $clave, $valor = $_ -split '=', 2
        $valores[$clave.Trim()] = $valor.Trim()
    }
    if (-not $SupabaseUrl) { $SupabaseUrl = $valores['SupabaseUrl'] }
    if (-not $SupabaseKey) { $SupabaseKey = $valores['SupabaseKey'] }
}
if (-not $SupabaseUrl -or -not $SupabaseKey) {
    throw "conexion.txt no tiene SupabaseUrl y/o SupabaseKey completos."
}

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ArchivoLog = 'C:\Produccion\Migracion_Supabase_log.txt'
$SupabaseUrl = $SupabaseUrl.TrimEnd('/')

function Escribir-Log {
    param([string]$Mensaje)
    $linea = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Mensaje"
    Add-Content -Path $ArchivoLog -Value $linea
    Write-Host $linea
}

function Encabezados {
    param([string]$Prefer)
    $h = @{
        'apikey'        = $SupabaseKey
        'Authorization' = "Bearer $SupabaseKey"
        'Content-Type'  = 'application/json'
    }
    if ($Prefer) { $h['Prefer'] = $Prefer }
    return $h
}

function Probar-Conexion {
    # Verifica credenciales ANTES de leer datos, para no fallar recien despues
    # de procesar miles de filas.
    $largo = $SupabaseKey.Length
    $inicio = $SupabaseKey.Substring(0, [Math]::Min(12, $largo))
    Escribir-Log "Verificando conexion. URL: $SupabaseUrl - llave: $largo caracteres, empieza con '$inicio...'"
    try {
        $uri = "$SupabaseUrl/rest/v1/cajas?select=numero_caja&limit=1"
        Invoke-RestMethod -Uri $uri -Headers (Encabezados) -Method Get | Out-Null
        Escribir-Log 'Conexion verificada correctamente.'
    }
    catch {
        $codigo = $null
        if ($_.Exception.Response) { $codigo = [int]$_.Exception.Response.StatusCode }
        if ($codigo -eq 401) {
            throw @"
Supabase rechazo la llave (401 Invalid API key).

Que revisar, en este orden:
 1. En Supabase: Project Settings -> API Keys. Si ahi ves una llave que
    empieza con 'sb_secret_', usa ESA en conexion.txt (el proyecto usa el
    sistema nuevo de llaves y la JWT larga 'eyJ...' ya no sirve).
 2. Que sea la llave secreta/service_role, no la publica (anon/publishable).
 3. Que en conexion.txt la llave este completa y en UNA sola linea.
"@
        }
        if ($codigo -eq 404) {
            throw "No se encontro la tabla 'cajas' en $SupabaseUrl. Revisa que hayas corrido schema_supabase.sql en el SQL Editor."
        }
        throw "No se pudo conectar a Supabase: $($_.Exception.Message)"
    }
}

function Tabla-TieneFilas {
    param([string]$Tabla)
    $uri = "$SupabaseUrl/rest/v1/$Tabla" + '?select=id&limit=1'
    $resultado = Invoke-RestMethod -Uri $uri -Headers (Encabezados) -Method Get
    return @($resultado).Count -gt 0
}

function Vaciar-Tabla {
    param([string]$Tabla)
    $uri = "$SupabaseUrl/rest/v1/$Tabla" + '?id=gt.0'
    Invoke-RestMethod -Uri $uri -Headers (Encabezados) -Method Delete | Out-Null
}

function Enviar-Lote {
    param(
        [string]$Tabla,
        [System.Collections.Generic.List[object]]$Filas,
        [string]$ClaveConflicto
    )
    $uri = "$SupabaseUrl/rest/v1/$Tabla"
    $prefer = 'return=minimal'
    if ($ClaveConflicto) {
        $uri += "?on_conflict=$ClaveConflicto"
        $prefer = 'resolution=merge-duplicates,return=minimal'
    }
    $cuerpo = $Filas | ConvertTo-Json -Depth 5
    if ($Filas.Count -eq 1) { $cuerpo = "[$cuerpo]" }
    Invoke-RestMethod -Uri $uri -Headers (Encabezados -Prefer $prefer) -Method Post -Body $cuerpo | Out-Null
}

$Tablas = @(
    @{ Access = 'Cajas'; Postgres = 'cajas'; Estrategia = 'upsert'; ClaveConflicto = 'numero_caja'
       Columnas = @(
            @{A='NumeroCaja'; P='numero_caja'}, @{A='NumeroSAG'; P='numero_sag'},
            @{A='PesoNeto'; P='peso_neto'}, @{A='PesoBruto'; P='peso_bruto'},
            @{A='PesoNetoEtiqueta'; P='peso_neto_etiqueta'}, @{A='PesoBrutoEtiqueta'; P='peso_bruto_etiqueta'},
            @{A='Piezas'; P='piezas'}, @{A='CodigoProducto'; P='codigo_producto'},
            @{A='RutCliente'; P='rut_cliente'}, @{A='NombreCliente'; P='nombre_cliente'},
            @{A='FechaDesposte'; P='fecha_desposte'}, @{A='FechaFaena'; P='fecha_faena'},
            @{A='FechaVencimiento'; P='fecha_vencimiento'}, @{A='FechaCongelado'; P='fecha_congelado'},
            @{A='NumeroCombo'; P='numero_combo'}, @{A='GuiaRecepcionGanado'; P='guia_recepcion_ganado'},
            @{A='GuiaRecepcion'; P='guia_recepcion'}, @{A='LoteUnico'; P='lote_unico'},
            @{A='CodigoSAP'; P='codigo_sap'}, @{A='NombreDestare'; P='nombre_destare'},
            @{A='PesoDestare'; P='peso_destare'}, @{A='ProductoEspanol'; P='producto_espanol'},
            @{A='ProductoIngles'; P='producto_ingles'}, @{A='FormatoProducto'; P='formato_producto'},
            @{A='CodEnvase'; P='cod_envase'}, @{A='FormatoEnvase'; P='formato_envase'},
            @{A='Etiqueta'; P='etiqueta'}, @{A='Mercado'; P='mercado'},
            @{A='PaisDestino'; P='pais_destino'}, @{A='PigFarm'; P='pig_farm'},
            @{A='FechaPesaje'; P='fecha_pesaje'}, @{A='NombrePredio'; P='nombre_predio'},
            @{A='FechaTraspasoSAP'; P='fecha_traspaso_sap'}, @{A='FechaAsigInsumos'; P='fecha_asig_insumos'},
            @{A='Usuario'; P='usuario'}, @{A='PC'; P='pc'},
            @{A='ArchivoOrigen'; P='archivo_origen'}, @{A='FechaImportacion'; P='fecha_importacion'}
       )
    },
    @{ Access = 'ConfigApp'; Postgres = 'config_app'; Estrategia = 'upsert'; ClaveConflicto = 'clave'
       Columnas = @( @{A='Clave'; P='clave'}, @{A='Valor'; P='valor'} )
    },
    @{ Access = 'Despachos'; Postgres = 'despachos'; Estrategia = 'upsert'; ClaveConflicto = 'numero_caja'
       Columnas = @(
            @{A='FechaDespacho'; P='fecha_despacho'}, @{A='Cliente'; P='cliente'},
            @{A='SubCliente'; P='sub_cliente'}, @{A='Tipo'; P='tipo'},
            @{A='NumeroCaja'; P='numero_caja'}, @{A='Origen'; P='origen'}, @{A='Creado'; P='creado'}
       )
    },
    @{ Access = 'EstandaresPeso'; Postgres = 'estandares_peso'; Estrategia = 'reemplazar'
       Columnas = @(
            @{A='Producto'; P='producto'}, @{A='CodigoProducto'; P='codigo_producto'},
            @{A='Tipo'; P='tipo'}, @{A='PesoObjetivo'; P='peso_objetivo'},
            @{A='DesvEst'; P='desv_est'}, @{A='MinPeso'; P='min_peso'}, @{A='MaxPeso'; P='max_peso'},
            @{A='NCajas'; P='n_cajas'}, @{A='KgTotal'; P='kg_total'}, @{A='FechaCalculo'; P='fecha_calculo'}
       )
    },
    @{ Access = 'InventarioSAP'; Postgres = 'inventario_sap'; Estrategia = 'reemplazar'
       Columnas = @(
            @{A='NumeroCaja'; P='numero_caja'}, @{A='Kg'; P='kg'},
            @{A='FechaProduccion'; P='fecha_produccion'}, @{A='Descripcion'; P='descripcion'},
            @{A='FechaFoto'; P='fecha_foto'}
       )
    },
    @{ Access = 'Pedidos'; Postgres = 'pedidos'; Estrategia = 'insertar_si_vacia'
       Columnas = @(
            @{A='FechaPedido'; P='fecha_pedido'}, @{A='Cliente'; P='cliente'}, @{A='Producto'; P='producto'},
            @{A='Cantidad'; P='cantidad'}, @{A='Unidad'; P='unidad'}, @{A='Estado'; P='estado'},
            @{A='Creado'; P='creado'}
       )
    },
    @{ Access = 'PedidosCombos'; Postgres = 'pedidos_combos'; Estrategia = 'insertar_si_vacia'
       Columnas = @(
            @{A='Fecha'; P='fecha'}, @{A='Cliente'; P='cliente'}, @{A='CodigoProducto'; P='codigo_producto'},
            @{A='Producto'; P='producto'}, @{A='CombosPedidos'; P='combos_pedidos'},
            @{A='KgPedidos'; P='kg_pedidos'}, @{A='Origen'; P='origen'}, @{A='Creado'; P='creado'}
       )
    },
    @{ Access = 'ProgramaSemanal'; Postgres = 'programa_semanal'; Estrategia = 'insertar_si_vacia'
       Columnas = @(
            @{A='Fecha'; P='fecha'}, @{A='MaterialH'; P='material_h'}, @{A='Producto'; P='producto'},
            @{A='PedidoKG'; P='pedido_kg'}, @{A='PedidoCajas'; P='pedido_cajas'}, @{A='Formato'; P='formato'},
            @{A='Cliente'; P='cliente'}, @{A='Destino'; P='destino'}, @{A='Estado'; P='estado'},
            @{A='FechaCarga'; P='fecha_carga'}
       )
    }
)

Escribir-Log 'Inicio de migracion a Supabase.'
Probar-Conexion

$conexionAccess = New-Object System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$RutaAccess;")
$conexionAccess.Open()

try {
    foreach ($tabla in $Tablas) {
        Escribir-Log "Tabla $($tabla.Access) -> $($tabla.Postgres) (estrategia: $($tabla.Estrategia))"

        if ($tabla.Estrategia -eq 'insertar_si_vacia' -and (Tabla-TieneFilas -Tabla $tabla.Postgres)) {
            Escribir-Log "  Se omite: $($tabla.Postgres) ya tiene datos (evita duplicar historico manual)."
            continue
        }
        if ($tabla.Estrategia -eq 'reemplazar') {
            Vaciar-Tabla -Tabla $tabla.Postgres
            Escribir-Log "  $($tabla.Postgres) vaciada antes de recargar."
        }

        $comando = $conexionAccess.CreateCommand()
        $consulta = "SELECT * FROM [$($tabla.Access)]"
        if ($DesdeFecha -and $tabla.Access -eq 'Cajas') {
            # SQL de Access: las fechas literales van entre almohadillas y en mm/dd/yyyy.
            $fecha = [datetime]::ParseExact($DesdeFecha, 'yyyy-MM-dd', $null)
            $consulta += " WHERE FechaDesposte >= #$($fecha.ToString('MM/dd/yyyy'))#"
            Escribir-Log "  Filtrando Cajas desde $DesdeFecha."
        }
        $comando.CommandText = $consulta
        $lector = $comando.ExecuteReader()

        $claveConflicto = $null
        if ($tabla.Estrategia -eq 'upsert') { $claveConflicto = $tabla.ClaveConflicto }

        $lote = New-Object System.Collections.Generic.List[object]
        $total = 0
        while ($lector.Read()) {
            $fila = [ordered]@{}
            foreach ($col in $tabla.Columnas) {
                $valor = $lector[$col.A]
                if ($valor -is [DBNull]) { $valor = $null }
                elseif ($valor -is [datetime]) { $valor = $valor.ToString('o') }
                $fila[$col.P] = $valor
            }
            $lote.Add($fila)
            if ($lote.Count -ge $TamanoLote) {
                Enviar-Lote -Tabla $tabla.Postgres -Filas $lote -ClaveConflicto $claveConflicto
                $total += $lote.Count
                $lote.Clear()
                Escribir-Log "  $($tabla.Postgres): $total filas enviadas..."
            }
        }
        if ($lote.Count -gt 0) {
            Enviar-Lote -Tabla $tabla.Postgres -Filas $lote -ClaveConflicto $claveConflicto
            $total += $lote.Count
        }
        $lector.Close()
        Escribir-Log "  $($tabla.Postgres) completa: $total filas."
    }
    Escribir-Log 'Migracion finalizada con exito.'
}
catch {
    Escribir-Log "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    $conexionAccess.Close()
}
