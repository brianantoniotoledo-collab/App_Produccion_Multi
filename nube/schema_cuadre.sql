-- ============================================================
--  Cuadre de Piezas: configuracion y consultas agregadas
--  Correr en Supabase -> SQL Editor, DESPUES de schema_supabase.sql
-- ============================================================

-- ------------------------------------------------------------
--  1. Configuracion que hoy vive en archivos JSON de C:\Produccion
--     Se sube a la nube para que editar el peso de vara o una zona
--     desde el celular quede reflejado en todos los dispositivos.
-- ------------------------------------------------------------

-- Config_Zonas.json: codigo de producto -> zona anatomica
CREATE TABLE IF NOT EXISTS config_zonas (
    codigo_producto  TEXT PRIMARY KEY,
    zona             TEXT NOT NULL
);

-- Cerdos_Por_Dia.json: fecha -> cerdos faenados
CREATE TABLE IF NOT EXISTS cerdos_por_dia (
    fecha   DATE PRIMARY KEY,
    cerdos  INTEGER NOT NULL
);

-- Peso_Vara.json y Tipos_Cambio.json: valores sueltos clave/valor
CREATE TABLE IF NOT EXISTS parametros (
    clave  TEXT PRIMARY KEY,
    valor  DOUBLE PRECISION NOT NULL
);

-- ------------------------------------------------------------
--  2. Consultas agregadas
--     Equivalen a las de la app de escritorio, traducidas de SQL de
--     Access a Postgres. Se agrupan en el servidor para no traer
--     cientos de miles de filas al navegador ni al telefono.
--
--     Reglas replicadas tal cual de Produccion_App.ps1:
--       - cliente unico: 'AASA PORK LIMITADA'
--       - la estacion COM_DSK_SUBPROD marca subproducto (es_sub),
--         no se excluye en el cuadre (si en la tendencia)
--       - todo se mide por fecha_pesaje, no por fecha_desposte
-- ------------------------------------------------------------

CREATE OR REPLACE FUNCTION materiales_dia(p_fecha DATE)
RETURNS TABLE (
    codigo_producto   DOUBLE PRECISION,
    producto_espanol  TEXT,
    es_sub            INTEGER,
    piezas            BIGINT,
    kg                DOUBLE PRECISION,
    cajas             BIGINT
)
LANGUAGE sql STABLE AS $$
    SELECT c.codigo_producto,
           TRIM(c.producto_espanol)::TEXT,
           CASE WHEN c.usuario = 'COM_DSK_SUBPROD' THEN 1 ELSE 0 END,
           COALESCE(SUM(c.piezas), 0)::BIGINT,
           COALESCE(SUM(c.peso_neto), 0),
           COUNT(*)::BIGINT
    FROM cajas c
    WHERE c.fecha_pesaje >= p_fecha
      AND c.fecha_pesaje <  p_fecha + 1
      AND c.nombre_cliente = 'AASA PORK LIMITADA'
    GROUP BY c.codigo_producto, TRIM(c.producto_espanol),
             CASE WHEN c.usuario = 'COM_DSK_SUBPROD' THEN 1 ELSE 0 END;
$$;

-- Piezas por dia y producto, para la barra de tendencia.
-- Aca SI se excluye la estacion de subproductos (igual que el original).
CREATE OR REPLACE FUNCTION tendencia_cuadre(p_desde DATE, p_hasta DATE)
RETURNS TABLE (
    dia              TEXT,
    codigo_producto  DOUBLE PRECISION,
    piezas           BIGINT
)
LANGUAGE sql STABLE AS $$
    SELECT TO_CHAR(c.fecha_pesaje, 'YYYY-MM-DD'),
           c.codigo_producto,
           COALESCE(SUM(c.piezas), 0)::BIGINT
    FROM cajas c
    WHERE c.piezas > 0
      AND c.fecha_pesaje >= p_desde
      AND c.fecha_pesaje <  p_hasta + 1
      AND c.usuario <> 'COM_DSK_SUBPROD'
      AND c.nombre_cliente = 'AASA PORK LIMITADA'
    GROUP BY TO_CHAR(c.fecha_pesaje, 'YYYY-MM-DD'), c.codigo_producto;
$$;

-- Kilos por dia y producto para el historial de aprovechamiento.
CREATE OR REPLACE FUNCTION historial_aprovechamiento(p_desde DATE, p_hasta DATE)
RETURNS TABLE (
    dia               TEXT,
    codigo_producto   DOUBLE PRECISION,
    producto_espanol  TEXT,
    es_sub            INTEGER,
    kg                DOUBLE PRECISION
)
LANGUAGE sql STABLE AS $$
    SELECT TO_CHAR(c.fecha_pesaje, 'YYYY-MM-DD'),
           c.codigo_producto,
           TRIM(c.producto_espanol)::TEXT,
           CASE WHEN c.usuario = 'COM_DSK_SUBPROD' THEN 1 ELSE 0 END,
           COALESCE(SUM(c.peso_neto), 0)
    FROM cajas c
    WHERE c.fecha_pesaje >= p_desde
      AND c.fecha_pesaje <  p_hasta + 1
      AND c.nombre_cliente = 'AASA PORK LIMITADA'
    GROUP BY TO_CHAR(c.fecha_pesaje, 'YYYY-MM-DD'), c.codigo_producto,
             TRIM(c.producto_espanol),
             CASE WHEN c.usuario = 'COM_DSK_SUBPROD' THEN 1 ELSE 0 END
    ORDER BY 1;
$$;

-- Ultima fecha con produccion cargada: la app web abre en ese dia.
CREATE OR REPLACE FUNCTION ultima_fecha_produccion()
RETURNS DATE
LANGUAGE sql STABLE AS $$
    SELECT MAX(fecha_pesaje)::DATE
    FROM cajas
    WHERE nombre_cliente = 'AASA PORK LIMITADA';
$$;

-- ------------------------------------------------------------
--  3. Acceso de solo lectura para la app web
--     La pagina lleva la llave publica (anon), nunca la secreta.
--     Con RLS activo y estas politicas, esa llave solo puede LEER:
--     cualquier intento de escribir o borrar es rechazado por la base.
-- ------------------------------------------------------------
ALTER TABLE cajas            ENABLE ROW LEVEL SECURITY;
ALTER TABLE config_zonas     ENABLE ROW LEVEL SECURITY;
ALTER TABLE cerdos_por_dia   ENABLE ROW LEVEL SECURITY;
ALTER TABLE parametros       ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS lectura_publica ON cajas;
CREATE POLICY lectura_publica ON cajas          FOR SELECT TO anon USING (true);
DROP POLICY IF EXISTS lectura_publica ON config_zonas;
CREATE POLICY lectura_publica ON config_zonas   FOR SELECT TO anon USING (true);
DROP POLICY IF EXISTS lectura_publica ON cerdos_por_dia;
CREATE POLICY lectura_publica ON cerdos_por_dia FOR SELECT TO anon USING (true);
DROP POLICY IF EXISTS lectura_publica ON parametros;
CREATE POLICY lectura_publica ON parametros     FOR SELECT TO anon USING (true);

-- El script de sincronizacion usa la llave secreta (service_role), que
-- pasa por encima de RLS: seguir escribiendo desde el PC no se ve afectado.

-- ------------------------------------------------------------
--  4. Indices
--     Sin ellos cada consulta recorre las ~670.000 filas de cajas y
--     Postgres la cancela por tiempo (error 57014). Ver nube/indices.sql,
--     que es el mismo bloque para correrlo por separado.
-- ------------------------------------------------------------
CREATE INDEX IF NOT EXISTS idx_cajas_cliente_pesaje ON cajas (nombre_cliente, fecha_pesaje);
CREATE INDEX IF NOT EXISTS idx_cajas_pesaje_desc    ON cajas (fecha_pesaje DESC);
CREATE INDEX IF NOT EXISTS idx_cajas_codigo_pesaje  ON cajas (codigo_producto, fecha_pesaje);
ANALYZE cajas;
