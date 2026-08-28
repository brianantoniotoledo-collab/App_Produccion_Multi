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
    [string]$DesdeFecha,
    # Modo sincronizacion: sube solo lo que llego nuevo al Access desde la
    # ultima vez, en vez de recorrer todo. Es el modo para la tarea programada
    # que corre cada 20 minutos detras del importador.
    [switch]$SoloNuevos,
    # Sube unicamente los archivos JSON de configuracion, sin abrir el Access.
    # Sirve para hacerlo desde un PC que no tiene la base (ej. el notebook):
    # los JSON estan en el repositorio, en app_original.
    [switch]$SoloConfiguraciones,
    # Carpeta donde buscar los JSON. Por defecto, la del propio Access.
    [string]$CarpetaJson
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

# Modo de autenticacion que resulto funcionar; lo fija Probar-Conexion.
$script:ModoAuth = 'ambos'

function Encabezados {
    param([string]$Prefer, [string]$Modo)
    if (-not $Modo) { $Modo = $script:ModoAuth }
    # Content-Type NO va aca: Windows PowerShell lo ignora si se pasa dentro de
    # -Headers y manda el cuerpo como formulario, lo que hace que PostgREST
    # responda "Empty or invalid json". Se pasa con -ContentType en Enviar-Lote.
    $h = @{}
    switch ($Modo) {
        # Las llaves nuevas (sb_secret_...) no son JWT: mandarlas tambien en
        # Authorization puede hacer que el servidor intente validarlas como JWT
        # y las rechace. Las legacy (eyJ...) si funcionan en ambos encabezados.
        'apikey' { $h['apikey'] = $SupabaseKey }
        'bearer' { $h['Authorization'] = "Bearer $SupabaseKey" }
        default  { $h['apikey'] = $SupabaseKey; $h['Authorization'] = "Bearer $SupabaseKey" }
    }
    if ($Prefer) { $h['Prefer'] = $Prefer }
    return $h
}

function Probar-Conexion {
    # Verifica credenciales ANTES de leer datos, y prueba las tres formas de
    # autenticar para no depender de adivinar cual espera este proyecto.
    $largo = $SupabaseKey.Length
    $inicio = $SupabaseKey.Substring(0, [Math]::Min(14, $largo))
    $final = $SupabaseKey.Substring([Math]::Max(0, $largo - 4))
    Escribir-Log "URL: $SupabaseUrl"
    Escribir-Log "Llave: $largo caracteres, de '$inicio' a '...$final'"

    if ($SupabaseKey -match '[^\x21-\x7E]') {
        throw @"
La llave contiene caracteres invalidos (espacios, puntos suspensivos o los
puntitos de enmascarado). Eso pasa cuando se copia el texto oculto en pantalla
en vez de usar el boton de copiar.

Solucion: en Supabase -> Project Settings -> API Keys -> Secret keys, usa el
BOTON DE COPIAR (icono de dos hojitas) y pega eso en conexion.txt.
"@
    }

    $uriLectura = "$SupabaseUrl/rest/v1/cajas?select=numero_caja&limit=1"
    $uriEscritura = "$SupabaseUrl/rest/v1/cajas"
    # Escritura de prueba con un arreglo vacio: PostgREST la acepta y no inserta
    # nada, pero falla con 401 si el modo no tiene permiso de escribir.
    $cuerpoVacio = [System.Text.Encoding]::UTF8.GetBytes('[]')

    $errores = @()
    # 'ambos' primero: el encabezado apikey solo identifica al proyecto, el rol
    # con el que actua la base viene de Authorization. Sin el, Supabase trata la
    # peticion como visitante anonimo: deja leer pero no escribir, y la
    # migracion moria recien al enviar el primer lote.
    foreach ($modo in @('ambos', 'bearer', 'apikey')) {
        try {
            Invoke-RestMethod -Uri $uriLectura -Headers (Encabezados -Modo $modo) -Method Get | Out-Null
            Invoke-RestMethod -Uri $uriEscritura -Headers (Encabezados -Modo $modo -Prefer 'return=minimal') `
                -Method Post -ContentType 'application/json; charset=utf-8' -Body $cuerpoVacio | Out-Null
            $script:ModoAuth = $modo
            Escribir-Log "Conexion verificada (lectura y escritura) con modo de autenticacion: $modo."
            return
        }
        catch {
            $codigo = $null
            if ($_.Exception.Response) { $codigo = [int]$_.Exception.Response.StatusCode }
            if ($codigo -eq 404) {
                throw "No se encontro la tabla 'cajas' en $SupabaseUrl. Revisa que hayas corrido schema_supabase.sql en el SQL Editor de Supabase."
            }
            $errores += "  modo '$modo': HTTP $codigo"
        }
    }

    throw @"
Supabase acepto la llave para leer pero no para escribir, con las tres formas
de autenticar:
$($errores -join "`n")

Casi siempre significa que la llave es de solo lectura. Usa la llave SECRETA:
Project Settings -> API Keys -> pestana "Legacy anon, service_role API keys"
-> fila service_role -> Reveal -> boton de copiar. La llave 'anon' no sirve
para escribir.

La llave leida de conexion.txt tiene $largo caracteres ('$inicio...$final').
Si ese largo se ve corto comparado con lo que muestra Supabase, quedo cortada
al copiarla: usa el BOTON DE COPIAR y pegala en UNA sola linea.
"@
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

function Obtener-MaximoEnNube {
    # Devuelve el valor mas alto de una columna de fecha en la nube, para saber
    # desde donde seguir subiendo sin repetir lo ya cargado.
    param([string]$Tabla, [string]$Columna)
    $uri = "$SupabaseUrl/rest/v1/$Tabla" + "?select=$Columna&order=$Columna.desc&limit=1"
    $resultado = Invoke-RestMethod -Uri $uri -Headers (Encabezados) -Method Get
    $fila = @($resultado)[0]
    if (-not $fila -or -not $fila.$Columna) { return $null }
    return [datetime]$fila.$Columna
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
    # Se envia como bytes UTF-8 explicitos: si se pasa el texto directamente,
    # Windows PowerShell lo codifica en Latin-1 y rompe las tildes y la ñ de
    # los nombres de producto.
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($cuerpo)
    Invoke-RestMethod -Uri $uri -Headers (Encabezados -Prefer $prefer) -Method Post `
        -ContentType 'application/json; charset=utf-8' -Body $bytes | Out-Null
}

function Vaciar-Con-Filtro {
    # PostgREST exige un filtro para borrar. Se usa la propia llave primaria
    # con "no es nulo", que siempre es verdadera y borra la tabla completa.
    param([string]$Tabla, [string]$ColumnaLlave)
    $uri = "$SupabaseUrl/rest/v1/$Tabla" + "?$ColumnaLlave=not.is.null"
    Invoke-RestMethod -Uri $uri -Headers (Encabezados) -Method Delete | Out-Null
}

function Sincronizar-Configuraciones {
    # Config_Zonas.json, Cerdos_Por_Dia.json, Peso_Vara.json y
    # Tipos_Cambio.json viven como archivos sueltos en C:\Produccion y NO
    # estan en el Access. La app web los necesita, y como se editan desde la
    # app, deben quedar en la nube para que el cambio se vea en todos lados.
    $carpeta = $CarpetaJson
    if (-not $carpeta) { $carpeta = Split-Path -Parent $RutaAccess }
    Escribir-Log "Configuraciones: leyendo JSON desde $carpeta"

    $rutaZonas = Join-Path $carpeta 'Config_Zonas.json'
    if (Test-Path $rutaZonas) {
        $mapa = Get-Content $rutaZonas -Raw -Encoding UTF8 | ConvertFrom-Json
        $filas = New-Object System.Collections.Generic.List[object]
        foreach ($p in $mapa.PSObject.Properties) {
            $filas.Add([ordered]@{ codigo_producto = [string]$p.Name; zona = [string]$p.Value })
        }
        Vaciar-Con-Filtro -Tabla 'config_zonas' -ColumnaLlave 'codigo_producto'
        if ($filas.Count -gt 0) { Enviar-Lote -Tabla 'config_zonas' -Filas $filas }
        Escribir-Log "  config_zonas: $($filas.Count) asignaciones."
    }
    else { Escribir-Log "  Config_Zonas.json no encontrado en $carpeta (se omite)." }

    $rutaCerdos = Join-Path $carpeta 'Cerdos_Por_Dia.json'
    if (Test-Path $rutaCerdos) {
        $mapa = Get-Content $rutaCerdos -Raw -Encoding UTF8 | ConvertFrom-Json
        $filas = New-Object System.Collections.Generic.List[object]
        foreach ($p in $mapa.PSObject.Properties) {
            $filas.Add([ordered]@{ fecha = [string]$p.Name; cerdos = [int]$p.Value })
        }
        Vaciar-Con-Filtro -Tabla 'cerdos_por_dia' -ColumnaLlave 'fecha'
        if ($filas.Count -gt 0) { Enviar-Lote -Tabla 'cerdos_por_dia' -Filas $filas }
        Escribir-Log "  cerdos_por_dia: $($filas.Count) fechas."
    }
    else { Escribir-Log "  Cerdos_Por_Dia.json no encontrado en $carpeta (se omite)." }

    # Valores sueltos: peso de vara y tipos de cambio, a la tabla parametros.
    $parametros = New-Object System.Collections.Generic.List[object]

    $rutaVara = Join-Path $carpeta 'Peso_Vara.json'
    if (Test-Path $rutaVara) {
        $j = Get-Content $rutaVara -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($j.PesoVara) { $parametros.Add([ordered]@{ clave = 'peso_vara'; valor = [double]$j.PesoVara }) }
    }

    $rutaTc = Join-Path $carpeta 'Tipos_Cambio.json'
    if (Test-Path $rutaTc) {
        $j = Get-Content $rutaTc -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($m in @('USD', 'EUR', 'JPY')) {
            if ($j.$m) { $parametros.Add([ordered]@{ clave = "tc_$($m.ToLower())"; valor = [double]$j.$m }) }
        }
    }

    if ($parametros.Count -gt 0) {
        Enviar-Lote -Tabla 'parametros' -Filas $parametros -ClaveConflicto 'clave'
        Escribir-Log "  parametros: $($parametros.Count) valores."
    }
    else { Escribir-Log "  Sin Peso_Vara.json ni Tipos_Cambio.json en $carpeta (se omiten)." }
}

$Tablas = @(
    # IncrementalAccess/IncrementalPg: columna de marca de tiempo que permite
    # subir solo lo nuevo en modo -SoloNuevos. Las tablas sin esta columna son
    # chicas y en ese modo se recargan completas.
    @{ Access = 'Cajas'; Postgres = 'cajas'; Estrategia = 'upsert'; ClaveConflicto = 'numero_caja'
       IncrementalAccess = 'FechaImportacion'; IncrementalPg = 'fecha_importacion'
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
       IncrementalAccess = 'Creado'; IncrementalPg = 'creado'
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

if ($SoloConfiguraciones) {
    Escribir-Log 'Inicio: solo configuraciones (sin abrir el Access).'
} elseif ($SoloNuevos) {
    Escribir-Log 'Inicio de sincronizacion (solo lo nuevo).'
} else {
    Escribir-Log 'Inicio de migracion a Supabase.'
}
Probar-Conexion

# Solo configuraciones: no se toca el Access, asi se puede correr desde un PC
# que no tiene la base (los JSON viven en el repositorio).
if ($SoloConfiguraciones) {
    try {
        Sincronizar-Configuraciones
        Escribir-Log 'Configuraciones subidas con exito.'
    }
    catch {
        Escribir-Log "ERROR: $($_.Exception.Message)"
        throw
    }
    return
}

$conexionAccess = New-Object System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$RutaAccess;")
$conexionAccess.Open()

try {
    foreach ($tabla in $Tablas) {
        # En modo sincronizacion las tablas de ingreso manual se recargan
        # completas (son chicas y el Access manda), en vez de omitirse.
        $estrategia = $tabla.Estrategia
        if ($SoloNuevos -and $estrategia -eq 'insertar_si_vacia') { $estrategia = 'reemplazar' }

        $usaIncremental = $SoloNuevos -and $tabla.IncrementalAccess
        if ($usaIncremental) { $estrategia = 'upsert' }

        Escribir-Log "Tabla $($tabla.Access) -> $($tabla.Postgres) (estrategia: $estrategia)"

        if ($estrategia -eq 'insertar_si_vacia' -and (Tabla-TieneFilas -Tabla $tabla.Postgres)) {
            Escribir-Log "  Se omite: $($tabla.Postgres) ya tiene datos (evita duplicar historico manual)."
            continue
        }
        if ($estrategia -eq 'reemplazar') {
            Vaciar-Tabla -Tabla $tabla.Postgres
            Escribir-Log "  $($tabla.Postgres) vaciada antes de recargar."
        }

        $comando = $conexionAccess.CreateCommand()
        $consulta = "SELECT * FROM [$($tabla.Access)]"
        if ($usaIncremental) {
            # Solo lo que entro al Access despues de lo ultimo que ya esta en la
            # nube. Se resta un minuto de holgura por si el importador escribio
            # filas mientras corria la sincronizacion anterior; el upsert se
            # encarga de que un solape no duplique nada.
            $ultimo = Obtener-MaximoEnNube -Tabla $tabla.Postgres -Columna $tabla.IncrementalPg
            if ($ultimo) {
                $corte = $ultimo.AddMinutes(-1)
                $consulta += " WHERE [$($tabla.IncrementalAccess)] > #$($corte.ToString('MM/dd/yyyy HH:mm:ss'))#"
                Escribir-Log "  Subiendo lo posterior a $($corte.ToString('dd-MM-yyyy HH:mm:ss'))."
            }
            else {
                Escribir-Log "  La nube esta vacia: se sube todo lo que haya."
            }
        }
        elseif ($DesdeFecha -and $tabla.Access -eq 'Cajas') {
            # SQL de Access: las fechas literales van entre almohadillas y en mm/dd/yyyy.
            $fecha = [datetime]::ParseExact($DesdeFecha, 'yyyy-MM-dd', $null)
            $consulta += " WHERE FechaDesposte >= #$($fecha.ToString('MM/dd/yyyy'))#"
            Escribir-Log "  Filtrando Cajas desde $DesdeFecha."
        }
        $comando.CommandText = $consulta
        $lector = $comando.ExecuteReader()

        $claveConflicto = $null
        if ($estrategia -eq 'upsert') { $claveConflicto = $tabla.ClaveConflicto }

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

    Sincronizar-Configuraciones

    Escribir-Log 'Migracion finalizada con exito.'
}
catch {
    Escribir-Log "ERROR: $($_.Exception.Message)"
    throw
}
finally {
    $conexionAccess.Close()
}
