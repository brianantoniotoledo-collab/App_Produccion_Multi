-- ============================================================
--  Diagnostico y limpieza de espacio en Supabase
--  Correr en Supabase -> SQL Editor
-- ============================================================

-- ------------------------------------------------------------
--  1. DIAGNOSTICO: que esta ocupando el espacio
--     Corre solo esta consulta primero, antes de borrar nada.
-- ------------------------------------------------------------
SELECT
    relname                                        AS tabla,
    pg_size_pretty(pg_total_relation_size(c.oid))  AS total,
    pg_size_pretty(pg_relation_size(c.oid))        AS solo_datos,
    pg_size_pretty(pg_indexes_size(c.oid))         AS solo_indices,
    (SELECT COUNT(*) FROM pg_index i WHERE i.indrelid = c.oid) AS n_indices
FROM pg_class c
JOIN pg_namespace n ON n.oid = c.relnamespace
WHERE n.nspname = 'public' AND c.relkind = 'r'
ORDER BY pg_total_relation_size(c.oid) DESC;

-- Rango de fechas y cuantas filas hay por mes en cajas:
SELECT TO_CHAR(fecha_pesaje, 'YYYY-MM') AS mes, COUNT(*) AS filas
FROM cajas
GROUP BY 1 ORDER BY 1;


-- ============================================================
--  2. OPCIONES PARA LIBERAR ESPACIO
--     Descomenta y corre SOLO la que elijas.
-- ============================================================

-- ------------------------------------------------------------
--  OPCION A - Recortar el historico (la mas efectiva y gratis)
--  El Cuadre usa el dia actual y una tendencia de 14 dias, asi que
--  con 6 meses sobra. El historico completo sigue intacto en el
--  Access del PC: esto solo recorta la copia de la nube.
--
--  Cambia la fecha si quieres conservar mas o menos.
-- ------------------------------------------------------------
-- DELETE FROM cajas WHERE fecha_pesaje < '2026-03-01';
-- VACUUM FULL cajas;   -- devuelve el espacio al disco (tarda un rato)
-- ANALYZE cajas;

-- ------------------------------------------------------------
--  OPCION B - Soltar el indice que el Cuadre no usa
--  idx_cajas_fecha_desposte se creo pensando en la app de KPIs.
--  El Cuadre trabaja por fecha_pesaje, asi que hoy no aporta.
--  Libera poco, pero es inmediato y sin perder datos.
-- ------------------------------------------------------------
-- DROP INDEX IF EXISTS idx_cajas_fecha_desposte;

-- ------------------------------------------------------------
--  OPCION C - Soltar columnas que la app web no usa
--  Estas columnas ocupan espacio por cada una de las ~670.000
--  filas y ninguna pantalla web las muestra hoy. Siguen estando
--  en el Access.
--  OJO: si mas adelante las necesitas, hay que volver a migrar.
-- ------------------------------------------------------------
-- ALTER TABLE cajas
--     DROP COLUMN IF EXISTS producto_ingles,
--     DROP COLUMN IF EXISTS guia_recepcion_ganado,
--     DROP COLUMN IF EXISTS nombre_destare,
--     DROP COLUMN IF EXISTS peso_destare,
--     DROP COLUMN IF EXISTS pig_farm,
--     DROP COLUMN IF EXISTS nombre_predio,
--     DROP COLUMN IF EXISTS archivo_origen,
--     DROP COLUMN IF EXISTS pc;
-- VACUUM FULL cajas;
-- ANALYZE cajas;

-- ------------------------------------------------------------
--  Despues de cualquier limpieza, verifica el resultado:
-- ------------------------------------------------------------
-- SELECT pg_size_pretty(pg_database_size(current_database())) AS base_completa;
