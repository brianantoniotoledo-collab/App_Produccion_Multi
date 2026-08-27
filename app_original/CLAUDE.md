# Proyecto: Sistema de Revisión de Producción — Comafri

Contexto para trabajar sobre los scripts de `C:\Produccion`.
Autor del proyecto: Brian Toledo Lucero (área Exportación, Comafri / Agrícola Industrial Lo Valledor AASA S.A.).

---

## 1. Qué es esto

Sistema semi-automatizado de control de producción de desposte de cerdo. Tres capas:

1. **Captura**: un importador programado lleva los Excel exportados del ERP a una base Access central.
2. **Consolidación**: histórico completo (~1.95M registros) migrado desde la base legada de la empresa.
3. **Aplicación de escritorio** (PowerShell + WPF) con KPIs en tiempo real, pedidos, combos, despachos y existencia pendiente.

Todo corre local en el PC de Brian (Windows, dominio `AASA-DOM`, usuario `brian.toledo`).

---

## 2. Convenciones y preferencias de trabajo

- **Idioma**: español (Chile). Nombres de variables y comentarios en español, sin tildes en el código para evitar problemas de codificación.
- **Pedir confirmación antes de modificar scripts existentes.** Brian explícitamente pidió avisar y esperar OK antes de aplicar cambios de código.
- **Entregar soluciones terminadas**, no kits de configuración. Se descartó Power BI por esta razón: se prefieren apps que se abran con doble clic y funcionen.
- **Iterar con correcciones puntuales**, no reescrituras completas.
- **Verificar con datos reales antes de entregar.** Varios bugs se detectaron reproduciendo el error de forma aislada. No asumir que un cambio de código resuelve el problema sin probarlo.
- Patrón estándar del área: PowerShell + Windows Forms/WPF para apps de escritorio; SharePoint + Power Query para archivos compartidos.

---

## 3. Archivos en `C:\Produccion`

| Archivo | Rol |
|---|---|
| `Importador_Produccion.ps1` | Importador automático ERP → Access. Corre cada 20 min por Programador de tareas. |
| `Migrar_Historico.ps1` | Migración única del histórico legado. **Ya ejecutado, no volver a correr.** |
| `Reparar_FechaPesaje.ps1` | Reparación puntual de `FechaPesaje`. **Ya ejecutado.** |
| `KPI_Produccion_App.ps1` | Aplicación principal (WPF). ~2.000 líneas. |
| `KPI_Produccion.vbs` | Lanzador sin ventana de consola. Apunta al nombre exacto `KPI_Produccion_App.ps1`. |
| `Base_Produccion.accdb` | Base de datos activa. |
| `BASE RCP.accdb` | Base legada, solo respaldo. No se toca. |
| `LOGO-AASA.jpg`, `AGROSUPER-PNG.png` | Logos que la app muestra según el filtro de cliente. |
| `Descargas_ERP\` | Carpeta vigilada. Brian suelta aquí los Excel del ERP. Subcarpetas `Procesados\` y `Con_Error\`. |
| `*_log.txt` | Logs del importador, migración y reparación. |

Proyecto hermano (no tocar sin motivo): `Etiquetas_SAG_App.ps1` — app de etiquetas SAG, ya completa.

---

## 4. Base de datos

`Base_Produccion.accdb`, motor **Microsoft Access (ACE.OLEDB.12.0)**.

### Tabla `Cajas` (principal)

36 campos que replican la consulta "Detalle Cajas Etiquetadas" del ERP:

`NumeroCaja` (texto, índice único), `NumeroSAG`, `PesoNeto`, `PesoBruto`, `PesoNetoEtiqueta`, `PesoBrutoEtiqueta`, `Piezas`, `CodigoProducto`, `RutCliente`, `NombreCliente`, `FechaDesposte`, `FechaFaena`, `FechaVencimiento`, `FechaCongelado`, `NumeroCombo`, `GuiaRecepcionGanado`, `GuiaRecepcion`, `LoteUnico`, `CodigoSAP`, `NombreDestare`, `PesoDestare`, `ProductoEspanol`, `ProductoIngles`, `FormatoProducto`, `CodEnvase`, `FormatoEnvase`, `Etiqueta`, `Mercado`, `PaisDestino`, `PigFarm`, `FechaPesaje`, `NombrePredio`, `FechaTraspasoSAP`, `FechaAsigInsumos`, `Usuario`, `PC`
\+ `ArchivoOrigen`, `FechaImportacion` (trazabilidad propia).

### Tablas creadas por la app

- `Pedidos` — pedidos manuales (los que llegan por correo en texto libre).
- `EstandaresPeso` — estándares de peso por producto. **Agrupada por `ProductoEspanol`, NO por código** (muchos registros tienen `CodigoProducto` = 0 o vacío).
- `ProgramaSemanal` — programa semanal leído de la hoja `Base.` del Excel de programación.
- `PedidosCombos` — pedidos de combos de Modinger y P.F.
- `Despachos` — folios despachados por cliente. Índice único por `NumeroCaja`.
- `ConfigApp` — clave/valor. Hoy guarda `CorteExistencia`.

---

## 5. Reglas de negocio (críticas)

- **Excluir siempre** `Usuario = 'COM_DSK_SUBPROD'` de los KPIs: no pasa por desposte.
- **Solo cuentan dos clientes** en los KPIs: `AASA PORK LIMITADA` y `AGROSUPER COMERCIAL. DE ALIMENTOS LTDA.` Este filtro **no** aplica a Buscar Producto ni a Pedidos.
- **Combo vs. caja**: `FormatoEnvase = 'COMBO CARTON'` identifica combos; el resto son cajas.
- **Fresco vs. congelado**: se determina por `DateDiff('d', FechaDesposte, FechaVencimiento)`. Fresco ≈ 12 días, congelado ≈ 2 años. Umbral en código: `$DiasFrescoMax = 180`.
- **Turno**: desposte tiene solo Turno A (06:00–16:00). No hay turno B en este proceso.
- **Hígados**: nunca se etiquetan en desposte. Excluirlos siempre de los pedidos de combos.
- **Despacho ≠ etiquetado**: lo producido un día puede despacharse otro (ej. producción de sábado despachada el lunes). Por eso el cumplimiento de combos se cuenta **por fecha de despacho**, no por `FechaPesaje`.
- **SAP manda sobre la app para existencia real**: SAP es el sistema de facturación y movimientos reales de la empresa, más autoritativo que el ERP de producción. Cuando hay una foto SAP cargada en Inventario, la Existencia Online se **limita** a lo que figura en SAP — todo lo que la app cree vivo pero SAP ya movió/facturó se rebaja automáticamente (no queda como "pendiente" fantasma). Ejemplo validado: 1000 cajas producidas, 0 despachadas, foto SAP con 100 folios → existencia real queda en 100, se rebajan 900.

### Pesos fijos (para giveaway)

```
130374 = 15.0 kg   TIRA DE LOMO DE CERDO CONGELADO
130154 =  5.0 kg   FILETE CERDO C/CABEZA CONGELADO VP UE AF
130051 =  5.0 kg   PLATEADA DE LOMO CONGELADO VP
```

Todo el resto usa **promedio + desviación estándar** calculado sobre los últimos 6 meses de la propia base.

### Catálogo de combos

**Modinger** (15 códigos):
`130119, 130123, 130124, 130127, 130128, 130139, 130176, 130181, 130183, 130189, 130210, 130211, 130240, 130251, 130252`

**P.F.** (4 códigos):
`130122, 130186, 130200, 130212`

Modinger envía cantidades **en combos**; P.F. envía **kilos**, que se dividen por `$DivisorPF = 850` (peso promedio real del combo P.F., no es fijo: varía 840–867).

Advertencia: en la hoja "Modinger" del Excel, la columna `Hanna` **ha tenido códigos equivocados**. El parser prioriza esa columna solo si el código existe en el catálogo; si no, deduce por nombre.

---

## 6. Módulos de la app

1. **KPIs** — rango de fechas configurable (dos calendarios), filtro de cliente con logo dinámico, cajas vs. combos separados, ritmo kg/hora, mix export/nacional, giveaway, % fuera de rango, brecha producción→etiquetado, cumplimiento de pedidos con semáforo 90/95/100%.
2. **Programa Semanal** — carga la hoja `Base.` del Excel de programación (>200 productos × 6 días; la mayoría en cero, hay que filtrar `Pedido KG > 0` y descartar la fila TOTAL final que duplica la suma).
3. **Combos PF / Modinger** — carga por **pegado de texto** (ver §7). Avisos de "PEDIDO COMPLETO" y "EXCEDIDO" por producto y por día. El avance se cuenta por fecha de **despacho**, no de etiquetado.
4. **Despachos** — carga del Excel de andén (una hoja por cliente; detecta subclientes por los títulos de columna, ej. Steve / Carpo / Don Chicharrón dentro de "Huella Austral") o ingreso manual pegando folios. Distingue Nacional / Exportación. Folios aún no importados quedan como PENDIENTE y se resuelven solos.
5. **Existencia Online** — producido − despachado − rebajado por SAP = existencia real, desde una fecha de corte configurable (inicial `2026-07-23`). Si hay una foto SAP cargada (módulo Inventario), esta **manda**: solo cuenta como existencia lo que SAP confirma, y se recalcula sola al cargar una foto nueva.
6. **Inventario** — carga la foto de existencia SAP (Excel) y cruza contra la base y los despachos en 4 grupos: existencia confirmada (a contar físico), SAP existe/app ya despachó (posible error), en SAP pero no en la base (folio no importado aún), y por rebajar (viva en la app pero SAP ya la movió). Exportable a Excel al escritorio, una hoja por grupo.
7. **Pedidos** — carga manual de los pedidos que llegan por correo.
8. **Buscar Producto** — estadísticas de peso por producto (promedio, mín, máx, desviación).

Identidad visual: azul `#1B3F94`, naranja `#F7941D`.

---

## 7. Trampas conocidas (leer antes de tocar código)

Cada una de estas costó tiempo real de depuración:

- **`TryParseExact` con arreglo de formatos**: hay que tiparlo explícitamente como `[string[]]`. Con `@(...)` queda `object[]`, .NET resuelve la sobrecarga equivocada y **falla en silencio**, devolviendo siempre `false`. Esto dejó en blanco 1.9M de fechas.
- **Precedencia de la coma en PowerShell**: `@($x[0] + $v, $a, $b)` se interpreta como `$x[0] + ($v, $a, $b)` y lanza `op_Addition`. Calcular la suma en su propia línea antes de construir el arreglo.
- **`Get-ChildItem -Include`** exige comodín en la ruta: `Join-Path $carpeta '*'`. Sin eso devuelve cero resultados sin error.
- **Diccionarios `[ordered]@{}` con claves enteras**: el indexador se vuelve ambiguo (¿índice o clave?). Usar arreglos de hashtables: `@(@{ Cod = 130119; Nom = "..." })`.
- **Comparaciones con `StartsWith` contra texto vacío** dan siempre verdadero. Validar longitud mínima antes de comparar nombres.
- **Automatización COM de Excel es frágil**: Vista Protegida bloquea la apertura invisible; los procesos `EXCEL.EXE` quedan huérfanos si no se liberan **todos** los objetos (hojas → colección → libro → aplicación) con `FinalReleaseComObject` + `GC::Collect`. Donde se pueda, preferir pegado de texto.
- **Access tiene límite duro de 2 GB** por archivo `.accdb`. La base ya lo alcanzó una vez; se resolvió con "Compactar y reparar". **Monitorear**: con ~2M de registros creciendo, en algún momento habrá que migrar de motor.
- **Conexión a Access**: abrir y cerrar por operación, nunca mantenerla viva. Una conexión persistente en la app bloquea al importador que escribe cada 20 min.
- **SQL de Access, no T-SQL**: usar `IIF`, `NOW()`, `Date()`, `DateValue`, `DateAdd`, `DateDiff`, `DatePart`, `STDEV`, `FIRST`. Parámetros OleDb son **posicionales** (`?`), el nombre no importa pero deben ir en orden.
- **Tarea programada**: no se puede usar "ejecutar aunque el usuario no haya iniciado sesión" (la política de dominio rechaza la contraseña). Queda en modo sesión interactiva, lo que basta para el uso diario.

---

## 8. Estado y pendientes

**Funcionando:**
- Importador cada 20 min, autónomo.
- Histórico consolidado y reparado.
- App con los 8 módulos (KPIs, Programa Semanal, Combos PF/Modinger, Despachos, Existencia Online, Inventario, Pedidos, Buscar Producto).

**Pendiente:**
- Acceso de solo lectura (`db_datareader`) al SQL Server del ERP: `172.16.0.188:1433`. Solicitado a informática. Si llega, permite conexión ODBC directa y elimina la descarga manual.
- KPI de **rendimiento de desposte** (kg canal vs. kg producto terminado): requiere la tabla `Animalesfaenados` del ERP, hoy inaccesible.
- Multiusuario (jefatura, sala de proceso): la carpeta de SharePoint disponible sincroniza con OneDrive y eso choca con archivos `.accdb`. Alternativa evaluada: exportar a Excel plano en esa carpeta.
- Monitoreo del límite de 2 GB de Access.
- Mejoras propuestas a Inventario sin aplicar aún: cruce por lote/material, antigüedad en stock, comparación kg app vs. kg SAP, resumen por producto, histórico de fotos SAP.

## 8.5 Decisión de arquitectura pendiente: modularizar el archivo

`KPI_Produccion_App.ps1` ya tiene ~2.400 líneas y 8 módulos compartiendo variables globales y funciones. Brian lo notó: el archivo "se enredó" al abarcar tanto en un solo lugar (ej. el bug de `op_Addition` apareció "en combos" pero en realidad estaba en una función de resumen compartida).

**Decisión tomada con Brian: separar en ARCHIVOS, no en apps independientes.** Sigue siendo una sola ventana con las mismas 8 pestañas — el usuario no nota ningún cambio. Lo que cambia es que el código vive en módulos que la app principal importa, en vez de todo en un solo archivo.

Apps independientes (una ventana por módulo) **no** se justifican hoy porque Brian es el único usuario navegando entre todo — tener 8 ventanas sueltas sería más fricción, no menos. Reconsiderar esto solo si algún día un perfil de usuario distinto (ej. jefatura) debe ver un subconjunto y nada más del resto por tema de acceso a datos.

**Estructura de archivos propuesta** (a implementar en Claude Code, ver `Guia_Migracion_Claude_Code.md`):

```
C:\Produccion\KPI_App\
├── App.ps1              (punto de entrada: arma la ventana XAML, importa los demas, orquesta navegacion)
├── Db.ps1                (conexion, Tabla/Escalar/Ejecutar, esquema de tablas, migraciones de version)
├── Catalogos.ps1         (clientes AASA/Agrosuper, catalogo Modinger/PF, pesos fijos, colores de marca)
├── Modulos\
│   ├── Kpis.ps1
│   ├── ProgramaSemanal.ps1
│   ├── Combos.ps1
│   ├── Despachos.ps1
│   ├── ExistenciaOnline.ps1
│   ├── Inventario.ps1
│   ├── Pedidos.ps1
│   └── BuscarProducto.ps1
└── KPI_Produccion.vbs   (lanzador, sin cambios)
```

Reglas para el refactor:
- **Primero separar sin cambiar comportamiento.** Ninguna función ni cálculo debe cambiar de resultado durante la migración — es reorganización pura, no una reescritura. Verificar cada módulo probándolo igual que hoy (ver §9) antes de dar el refactor por bueno.
- Las funciones de `Db.ps1` (`Tabla`, `Escalar`, `Ejecutar`) son las que más módulos comparten — no duplicarlas en cada archivo.
- El patrón de conexión abrir/cerrar por operación (§7) se mantiene igual, solo cambia de archivo.
- Cada módulo debe poder probarse de forma aislada cuando tenga sentido (ej. `Procesar-PegadoCombos` en `Combos.ps1` no depende de la interfaz, se puede probar con texto de ejemplo sin abrir la ventana).

---

## 9. Cómo probar cambios

Verificación de sintaxis sin ejecutar la app:

```powershell
$errores = $null; $tokens = $null
[System.Management.Automation.Language.Parser]::ParseFile(
    'C:\Produccion\KPI_Produccion_App.ps1', [ref]$tokens, [ref]$errores) | Out-Null
if ($errores.Count -eq 0) { 'OK' } else { $errores | ForEach-Object { "Linea $($_.Extent.StartLineNumber): $($_.Message)" } }
```

Ejecutar la app con consola visible (para ver errores):

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Produccion\KPI_Produccion_App.ps1"
```

Ejecutar el importador a mano:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Produccion\Importador_Produccion.ps1"
```

Antes de probar cargas de Excel: cerrar Access y revisar que no queden procesos `EXCEL.EXE` huérfanos en el Administrador de tareas.
