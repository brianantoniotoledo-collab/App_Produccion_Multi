# App_Produccion_Multi — acceso multi-dispositivo

Contexto para retomar este repositorio en una sesión nueva.
Autor: Brian Toledo Lucero — área Exportación, Comafri (Agrícola Industrial Lo Valledor AASA S.A.).

---

## 1. Qué es este repositorio y qué NO es

Este repo resuelve **un problema puntual**: que la app de producción, que hoy solo corre en el PC del trabajo de Brian, se pueda ver desde otros computadores y desde el celular.

**No es el repositorio de la app.** La app (`Produccion_App.ps1`) y todo su contexto de desarrollo viven en **otro proyecto de Claude**, aparte de este.

### Regla que no se rompe

> **`app_original/Produccion_App.ps1` es una COPIA de solo lectura. No se modifica aquí, nunca.**

Está en el repo para poder leer sus cálculos y portarlos con exactitud, y como respaldo fuera del PC del trabajo. Cualquier cambio a la app se hace en el otro proyecto de Claude. Si se edita en los dos lados, quedan dos versiones divergentes del mismo script — que es justo lo que hay que evitar.

Lo mismo aplica a `Config_Zonas.json`, `Cerdos_Por_Dia.json`, `Peso_Vara.json` y `Produccion_App.vbs`: son copias de referencia.

---

## 2. La app original (contexto mínimo para entender el resto)

`Produccion_App.ps1` — PowerShell + WPF, ~3.700 líneas, corre local contra `C:\Produccion\Base_Produccion.accdb` (Access, ~1,6 GB, ~1,95M filas).

Tres pestañas:

| Pestaña | Qué hace |
|---|---|
| **Giveaway** | Kilos regalados (`PesoNeto − PesoNetoEtiqueta`) en 5 productos de peso fijo. Precios en EUR/USD/JPY y tipos de cambio editables → traduce los kilos a pesos chilenos. |
| **Cuadre de Piezas** | 14 zonas anatómicas del cerdo con su silueta, piezas por cerdo, y 5 paneles laterales en kg/cerdo. Cerdos faenados × peso de vara para el aprovechamiento. |
| **Buscar Producto** | Estadísticas de peso por producto, exportables. |

Ojo: existe también un `KPI_Produccion_App.ps1` (más viejo, 8 módulos). **La app vigente es `Produccion_App`**, no esa.

### Reglas de negocio que la versión web replica

Verificadas leyendo el código, no asumidas:

- Cliente único: `AASA PORK LIMITADA` (la app de KPIs usaba dos clientes; esta usa uno)
- Todo se mide por **`FechaPesaje`**, no por `FechaDesposte`
- La estación `COM_DSK_SUBPROD` marca subproducto y **manda sobre la zona configurada**
- Los paneles laterales acumulan **kilos**; las zonas del cerdo acumulan **piezas**
- Zonas marcadas `excluido` no suman a ningún cuadre
- Esperado por zona = cerdos del día × piezas por cerdo de esa zona
- La tendencia excluye subproductos; el cuadre del día no
- Cortes de color: <85% rojo, 85–98% naranja, 98–105% verde, >105% violeta

Las configuraciones (zonas, cerdos por día, peso de vara, tipos de cambio) **no están en el Access**: viven como JSON sueltos en `C:\Produccion`.

---

## 3. Arquitectura decidida

Surgieron **dos necesidades distintas**, y se resuelven distinto:

### A. PCs de la planta (prioridad — lo pidió el jefe de producción)

Carpeta compartida en el servidor + la app de escritorio en cada equipo.

```
\\servidor\Produccion\Base_Produccion.accdb   ← una sola base
   ← escribe:  importador (PC de Brian, cada 20 min)
   ← leen:     PCs de planta (app de escritorio)
```

Funciona porque la app abre y cierra la conexión **por operación**: un escritor y varios lectores es el escenario sano de Access. Los problemas aparecen con varios escribiendo, o con OneDrive/SharePoint sincronizando el `.accdb`.

Pendiente para esto: carpeta compartida real (no OneDrive), motor ACE OLEDB en cada PC, e índices en el Access sobre `FechaPesaje` y `NombreCliente`. **El cambio de rutas en la app se hace en el otro proyecto de Claude, no aquí.**

### B. Acceso remoto de Brian (notebook personal, celular)

Copia en la nube (Supabase/Postgres) + app web. Es lo único que funciona fuera de la red de la planta.

```
Access  →  script de sincronización  →  Supabase  →  app web
```

El origen es una pieza reemplazable: hoy el Access se alimenta de un Excel descargado a mano del ERP; si algún día llega el acceso al SQL Server (`172.16.0.188:1433`, solicitado y pendiente), se cambia solo el cargador y el resto no se entera.

---

## 4. Estado actual

| Pieza | Estado |
|---|---|
| Migración Access → Supabase | Funcionando. Migró ~18 meses. |
| Sincronización incremental | Escrita y probada, **falta programarla** en el Programador de tareas |
| Esquema y consultas en la nube | Creado |
| App web del Cuadre | Publicada en Cloudflare Pages (subida manual de la carpeta `web`) |
| Configuraciones (JSON) en la nube | Subidas |
| Automatización de descarga del ERP | **Abandonada** — ver §6 |
| Despliegue en PCs de planta | Pendiente, requiere estar en el trabajo |

### Problema abierto: espacio en Supabase

La base llegó a **705 MB contra el límite de 500 MB** del plan gratis. Consecuencia: las consultas del Cuadre dan `57014 statement timeout`, probablemente porque los índices no alcanzaron a crearse por falta de espacio.

Con la arquitectura definida esto se destraba: la web pasó a ser **solo el acceso remoto de Brian**, no la herramienta principal. Con 3–6 meses de datos sobra y entra holgado en el plan gratis. Los PC de planta ven el histórico completo desde el Access, que no tiene ese límite.

`nube/espacio.sql` trae el diagnóstico y las opciones de limpieza (recortar meses, soltar índices no usados, soltar columnas que la web no muestra). **Correr el diagnóstico antes de borrar nada.**

---

## 5. Archivos del repositorio

```
app_original/     COPIA de referencia. NO MODIFICAR.
importador/       Automatización de descarga del ERP (abandonada) + inspector de esquema
nube/             Migración, sincronización, esquemas SQL, diagnósticos
web/              App web del Cuadre de Piezas
```

- `nube/Migrar_A_Supabase.ps1` — hace las tres cosas: migración inicial (`-DesdeFecha`), sincronización incremental (`-SoloNuevos`) y solo configuraciones (`-SoloConfiguraciones`, sin abrir el Access)
- `nube/conexion.txt` — **no está en el repo** (`.gitignore`). Lleva la llave secreta y vive solo en cada PC. Si un PC nuevo falla con 401, es porque le falta este archivo.
- `web/config.js` — sí lleva la llave publicable, que es de solo lectura por diseño

---

## 6. Aprendizajes que costaron tiempo real

No repetir estos errores:

- **Trabajar sobre el archivo real, no sobre una descripción.** La automatización del ERP falló cuatro veces seguidas por asumir cosas del proceso en vez de mirarlo. Se abandonó.
- **El menú del ERP es invisible para UI Automation.** El proceso se llama `ProjectDesposte_Insumos` (no `SPC`), pero su menú está dibujado a mano y las herramientas de accesibilidad no lo ven. Por eso se descartó automatizar la descarga.
- **Verificar escritura, no solo lectura.** El encabezado `apikey` de Supabase identifica el proyecto; el rol viene de `Authorization`. Sin él, la base atiende como anónimo: deja leer pero no escribir, y la migración moría recién al enviar el primer lote.
- **En este proyecto la llave `sb_secret_...` es rechazada**; la que funciona es la legacy `service_role` (empieza con `eyJ`, ~219 caracteres).
- **Windows PowerShell ignora `Content-Type` dentro de `-Headers`** y manda el cuerpo como formulario → `PGRST102 Empty or invalid json`. Va como parámetro `-ContentType`, y el cuerpo como bytes UTF-8 explícitos (si no, se rompen las tildes y la ñ).
- **Los índices se crean para las columnas que se consultan.** El esquema traía índice por `fecha_desposte`, pero el Cuadre filtra por `fecha_pesaje` + `nombre_cliente`.
- **Estimar tamaño de base es poco fiable.** Se estimó 0,34 GB para 18 meses y resultaron 705 MB. Medir con `pg_total_relation_size`, no calcular.
- **La silueta del cerdo ya estaba en formato SVG** dentro del `.ps1`. Se extrajo con un script a `web/anatomia.js` en vez de transcribirla. Si cambia en la app, se regenera — no se edita a mano.

---

## 7. Cómo trabaja Brian

- Español (Chile). Código y comentarios en español, sin tildes en el código.
- **Confirmar antes de modificar cualquier script existente.**
- Prefiere soluciones terminadas por sobre kits de configuración, y explicaciones paso a paso cuando se trata de herramientas nuevas (GitHub, Supabase, Cloudflare).
- Corregirle terminología y redacción en español cuando corresponda.
- Verificar con datos reales antes de dar algo por resuelto.
