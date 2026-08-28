# Cuadre de Piezas — versión web

La misma pestaña "Cuadre de Piezas" de `Produccion_App.ps1`, pero en el navegador: se abre desde cualquier PC y desde el celular.

## Pasos (una sola vez)

### 1. Crear las tablas y consultas en Supabase

En Supabase → **SQL Editor** → **New query**, pega el contenido de `nube/schema_cuadre.sql` y dale **Run**.

Eso crea:
- `config_zonas`, `cerdos_por_dia`, `parametros` — la configuración que hoy vive en los JSON de `C:\Produccion`
- Las consultas agregadas (`materiales_dia`, `tendencia_cuadre`, etc.), que agrupan en el servidor para no traer cientos de miles de filas al teléfono
- **Seguridad de solo lectura**: la página lleva la llave pública, y con estas políticas esa llave solo puede leer. Cualquier intento de escribir o borrar lo rechaza la base.

### 2. Subir la configuración de los JSON

En el PC del trabajo, doble clic en `nube\Sincronizar_Nube.vbs` (o corre el `.bat` de migración). Ahora también sube `Config_Zonas.json`, `Cerdos_Por_Dia.json` y `Peso_Vara.json`.

Revisa `C:\Produccion\Migracion_Supabase_log.txt`: debe decir algo como `config_zonas: 105 asignaciones`.

### 3. Poner la llave pública

Supabase → **Project Settings → API Keys** → pestaña **"Publishable and secret API keys"** → **Publishable key** → botón de copiar.

Abre `web/config.js` y reemplaza `PEGAR_AQUI_LA_LLAVE_PUBLICABLE` por esa llave.

Esta llave **sí** puede ir en el código: está diseñada para ser pública y solo permite leer. La secreta (`sb_secret_...`) nunca va aquí — esa se queda en `conexion.txt`, que no se sube.

### 4. Publicar la página

La forma más simple, gratis y sin instalar nada — **Cloudflare Pages**:

1. Entra a [pages.cloudflare.com](https://pages.cloudflare.com) y crea una cuenta.
2. **Create a project** → **Connect to Git** → autoriza GitHub y elige `App_Produccion_Multi`.
3. Configura:
   - **Production branch**: `claude/app-multi-device-access-7a4g5o`
   - **Build command**: *(déjalo vacío)*
   - **Build output directory**: `web`
4. **Save and Deploy**.

En un par de minutos te entrega una dirección tipo `https://app-produccion-multi.pages.dev`. Ábrela en el celular y agrégala a la pantalla de inicio.

Cada vez que subamos cambios al repositorio, la página se actualiza sola.

## Qué muestra

- **Silueta del cerdo** con las 14 zonas pintadas según su porcentaje de cuadre. Los mismos cortes de color del original: rojo bajo 85%, naranja 85–98%, verde 98–105%, violeta sobre 105%.
- **Toca una zona** para ver sus materiales, piezas producidas, esperadas y la diferencia.
- **Paneles** de cuero, recortes, huesos, despojos y subproductos en kg por cerdo.
- **Resumen por zona** y **tendencia de 14 días**, que cambia según la zona seleccionada.
- Se refresca sola cada 90 segundos, igual que la app de escritorio.

## Reglas replicadas del original

Verificadas contra `Produccion_App.ps1`:

- Cliente único: `AASA PORK LIMITADA`
- Todo se mide por **fecha de pesaje**, no fecha de desposte
- La estación `COM_DSK_SUBPROD` marca subproducto y **manda sobre la configuración de zona**: si un material sale de esa estación, va a subproductos aunque su código tenga otra zona asignada
- Los paneles laterales acumulan **kilos**; las zonas del cerpo acumulan **piezas**
- Zonas marcadas como `excluido` no suman a ningún cuadre
- Esperado por zona = cerdos del día × piezas por cerdo de esa zona
- La tendencia excluye la estación de subproductos (el cuadre del día no)

## Diferencias respecto a la app de escritorio

- **La configuración de zonas todavía se edita en el PC**, no desde la web. La página lee la configuración pero no la escribe (por seguridad, con llave de solo lectura). Si quieres editar zonas o peso de vara desde el celular, hay que agregar autenticación — dímelo y lo hacemos.
- El original mostraba "12 zonas" en el resumen global; son **14**. Se corrigió en la versión web.

## Si algo falla

La página muestra el error en pantalla, en español:

| Mensaje | Solución |
|---|---|
| "Falta la llave publicable" | Paso 3: falta pegar la llave en `web/config.js` |
| Error `404` | Falta correr `nube/schema_cuadre.sql` (paso 1) |
| Error `500` con `statement timeout` | Faltan los índices: corre `nube/indices.sql` en el SQL Editor |
| Error `401` | La llave de `config.js` está mal copiada |
| Zonas todas en 0% | Falta el paso 2: la configuración de zonas no está en la nube |
