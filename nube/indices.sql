-- ============================================================
--  Indices para el Cuadre de Piezas
--  Correr en Supabase -> SQL Editor. Es seguro repetirlo.
--
--  Sin estos indices, cada consulta recorre las ~670.000 filas de cajas
--  y Postgres la cancela por tiempo (error 57014, statement timeout).
--  Todas las consultas del Cuadre filtran por nombre_cliente + fecha_pesaje,
--  asi que un indice compuesto por ese par las cubre todas.
-- ============================================================

CREATE INDEX IF NOT EXISTS idx_cajas_cliente_pesaje
    ON cajas (nombre_cliente, fecha_pesaje);

-- Para ultima_fecha_produccion(): busca el maximo, asi que conviene
-- descendente para que lo encuentre en la primera fila del indice.
CREATE INDEX IF NOT EXISTS idx_cajas_pesaje_desc
    ON cajas (fecha_pesaje DESC);

-- El giveaway y buscar producto filtran ademas por codigo de producto.
CREATE INDEX IF NOT EXISTS idx_cajas_codigo_pesaje
    ON cajas (codigo_producto, fecha_pesaje);

-- Deja las estadisticas al dia para que el planificador elija bien.
ANALYZE cajas;
