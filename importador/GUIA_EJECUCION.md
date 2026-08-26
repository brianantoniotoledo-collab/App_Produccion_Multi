# Guía de ejecución — Descarga_Automatica_ERP.ps1

## 1. Antes de la primera prueba

- **SPC.exe** debe estar abierto y logeado, como lo dejas siempre.
- Cierra cualquier Excel que tengas abierto con datos sin guardar — el script manda `Ctrl+S` y teclas a la ventana activa, y si Excel de otra cosa queda al frente, le va a escribir a esa ventana por error.
- Verifica que PowerShell te deje ejecutar scripts (una sola vez, como administrador):

  ```powershell
  Get-ExecutionPolicy
  ```

  Si te devuelve `Restricted`, ejecuta:

  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

## 2. Primera prueba manual (con consola visible)

Ábrela así para ver todo lo que hace en vivo, línea por línea:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Produccion\importador\Descarga_Automatica_ERP.ps1"
```

**Mientras corre: no toques el mouse ni el teclado.** El script simula clics y escritura sobre la ventana de SPC.exe y luego sobre Excel — si le quitas el foco a esas ventanas mientras trabaja, se puede desordenar todo.

Deberías ver en pantalla, en este orden:
1. Se abre el menú `Consultas_Produccion` y luego `Detalle Cajas Etiquetadas`.
2. Se escriben las dos fechas y se aprieta `Buscar`.
3. Carga la grilla y se aprieta `Excel`.
4. Se abre un Excel nuevo con los datos, se guarda solo en `Descargas_ERP` y se cierra esa ventana de Excel.

## 3. Verificar el resultado

- Revisa que apareció un archivo nuevo en `C:\Produccion\Descargas_ERP\`, con nombre tipo `Produccion_ERP_20260826_170500.xlsx`.
- Abre `C:\Produccion\Descarga_Automatica_ERP_log.txt` y confirma que la última línea dice `Descarga automatica finalizada con exito.`
- Corre `Importador_Produccion.ps1` (o espera su ciclo de 20 min) y confirma que procesó ese archivo y lo movió a `Procesados\`.

## 4. Programarlo automático (Programador de tareas)

Mismo mecanismo que ya usas para `Importador_Produccion.ps1`, solo que este debe disparar **unos 5 minutos antes** en cada ciclo:

1. Abre **Programador de tareas** → *Crear tarea básica*.
2. Nombre: `Descarga Automatica ERP`.
3. Desencadenador: *Diariamente*, repetir cada 20 minutos (igual que el importador), pero con hora de inicio 5 minutos antes que la tarea del importador.
4. Acción: *Iniciar un programa*.
   - Programa: `powershell.exe`
   - Argumentos: `-ExecutionPolicy Bypass -File "C:\Produccion\importador\Descarga_Automatica_ERP.ps1"`
5. En **Configuración**, marca *"Ejecutar solo si el usuario inició sesión"* — igual que el importador, no puede correr en segundo plano sin sesión porque necesita ver la pantalla para manejar las ventanas.

## 5. Si algo falla

Abre `Descarga_Automatica_ERP_log.txt` y busca la línea que empieza con `ERROR:`. Las causas más probables de una primera corrida:

| Mensaje de error | Causa probable |
|---|---|
| `SPC.exe no esta abierto...` | Se cerró la sesión del ERP o se reinició el PC. Vuelve a loguearte y corre de nuevo. |
| `No se encontro el elemento 'Consultas_Produccion'...` | El menú tardó más de lo esperado en aparecer, o el PC estaba ocupado con otra ventana al frente. |
| `No aparecio la ventana 'Guardar como'...` | Excel tardó más de lo esperado en abrirse, o algo bloqueó el diálogo (otra ventana emergente de Excel). |
| `No se genero el archivo esperado en...` | El guardado falló silenciosamente — revisa si quedó un Excel abierto sin cerrar. |

Cuando pase alguno de estos, pásame el mensaje exacto de esa línea `ERROR:` (y en qué paso ibas al mirar la pantalla, si alcanzaste a ver algo) y lo ajustamos — es normal que la primera corrida contra el SPC.exe real necesite algún retoque de tiempos.
