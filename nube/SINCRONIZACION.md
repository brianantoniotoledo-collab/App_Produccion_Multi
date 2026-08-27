# Sincronización automática Access → nube

## Qué cambia en tu día a día

**Nada.** Tu rutina se mantiene igual:

1. Descargas la producción del ERP y dejas el Excel en `Descargas_ERP` *(igual que hoy)*
2. `Importador_Produccion.ps1` lo carga al Access cada 20 minutos *(igual que hoy)*
3. **Nuevo:** `Sincronizar_Nube.vbs` sube a Supabase solo lo que llegó nuevo *(automático, no lo tocas)*

El paso 3 corre solo, sin ventanas ni avisos. Puedes seguir cargando el Excel varias veces al día como siempre.

## Cómo detecta "lo nuevo"

No vuelve a subir las filas que ya están arriba. Antes de cada corrida le pregunta a Supabase cuál es el registro más reciente que tiene y solo lee del Access lo posterior a eso:

| Tabla | Cómo detecta lo nuevo |
|---|---|
| `Cajas` | por `FechaImportacion` (la marca que pone tu importador) |
| `Despachos` | por `Creado` |
| Resto (pedidos, combos, programa semanal, estándares, inventario SAP) | son chicas: se recargan completas cada vez |

Toma un minuto de holgura hacia atrás por si el importador escribió filas justo mientras corría la sincronización anterior. Si eso provoca un solape, no importa: la carga usa *upsert*, así que un folio repetido se actualiza en vez de duplicarse.

## Configurar la tarea programada (una sola vez)

1. Abre **Programador de tareas** → *Crear tarea básica*
2. Nombre: `Sincronizar Nube`
3. Desencadenador: *Diariamente*, y luego marca repetir **cada 20 minutos**
4. Acción: *Iniciar un programa*
   - Programa: `wscript.exe`
   - Argumentos: `"C:\Produccion\App_Produccion_Multi\nube\Sincronizar_Nube.vbs"`
     *(ajusta la ruta a donde tengas clonado el repositorio)*
5. En **Configuración**, marca *"Ejecutar solo si el usuario inició sesión"* — igual que tu importador

**Recomendación:** haz que dispare unos 5 minutos después de `Importador_Produccion.ps1`, para que siempre encuentre los datos ya cargados en el Access.

## Verificar que está funcionando

- Abre `C:\Produccion\Migracion_Supabase_log.txt` — cada corrida deja una línea con la hora y cuántas filas subió.
- Si no hubo producción nueva, verás la tabla con 0 filas. Eso es normal, no es un error.

## Correrla a mano cuando quieras

Doble clic en `Sincronizar_Nube.vbs`. No abre ninguna ventana; revisa el log para ver el resultado.

## Si algo falla

El log guarda el error completo. Los más probables:

| Mensaje | Qué significa |
|---|---|
| `Supabase acepto la llave para leer pero no para escribir` | La llave de `conexion.txt` es de solo lectura. Usa la `service_role`. |
| `No se encontro la tabla ... ` | Falta correr `schema_supabase.sql` en el SQL Editor de Supabase. |
| Error de ACE.OLEDB | El Access está abierto en modo exclusivo, o falta el motor de Access en ese PC. |

Un detalle importante: si la sincronización falla una vez, **no se pierde nada**. La corrida siguiente detecta que esos datos aún no están arriba y los sube igual.
