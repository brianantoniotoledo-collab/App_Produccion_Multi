-- Esquema de la base en la nube (Supabase / Postgres), generado a partir del
-- esquema real de Base_Produccion.accdb (ver importador/Inspeccionar_Esquema.ps1).
-- Nombres en snake_case porque Postgres pliega a minusculas los identificadores
-- sin comillas; usar CamelCase obligaria a entrecomillar en cada consulta.
--
-- Los campos que en Access son "Number - Double" pero parecen codigos enteros
-- (codigo_producto, codigo_sap, numero_combo, cod_envase) se mantienen como
-- DOUBLE PRECISION a proposito: asi estan guardados de verdad en el origen.

CREATE TABLE cajas (
    numero_caja            VARCHAR(100) PRIMARY KEY,
    numero_sag             VARCHAR(100),
    peso_neto              DOUBLE PRECISION,
    peso_bruto             DOUBLE PRECISION,
    peso_neto_etiqueta     DOUBLE PRECISION,
    peso_bruto_etiqueta    DOUBLE PRECISION,
    piezas                 DOUBLE PRECISION,
    codigo_producto        DOUBLE PRECISION,
    rut_cliente            VARCHAR(100),
    nombre_cliente         VARCHAR(100),
    fecha_desposte         TIMESTAMP,
    fecha_faena            TIMESTAMP,
    fecha_vencimiento      TIMESTAMP,
    fecha_congelado        TIMESTAMP,
    numero_combo           DOUBLE PRECISION,
    guia_recepcion_ganado  VARCHAR(100),
    guia_recepcion         VARCHAR(100),
    lote_unico             VARCHAR(100),
    codigo_sap             DOUBLE PRECISION,
    nombre_destare         VARCHAR(100),
    peso_destare           DOUBLE PRECISION,
    producto_espanol       VARCHAR(100),
    producto_ingles        VARCHAR(100),
    formato_producto       VARCHAR(100),
    cod_envase             DOUBLE PRECISION,
    formato_envase         VARCHAR(100),
    etiqueta               VARCHAR(100),
    mercado                VARCHAR(100),
    pais_destino           VARCHAR(100),
    pig_farm               VARCHAR(100),
    fecha_pesaje           TIMESTAMP,
    nombre_predio          VARCHAR(100),
    fecha_traspaso_sap     TIMESTAMP,
    fecha_asig_insumos     TIMESTAMP,
    usuario                VARCHAR(100),
    pc                     VARCHAR(100),
    archivo_origen         VARCHAR(150),
    fecha_importacion      TIMESTAMP
);

CREATE TABLE config_app (
    clave   VARCHAR(50) PRIMARY KEY,
    valor   VARCHAR(100)
);

CREATE TABLE despachos (
    id               SERIAL PRIMARY KEY,
    fecha_despacho   TIMESTAMP,
    cliente          VARCHAR(60),
    sub_cliente      VARCHAR(60),
    tipo             VARCHAR(15),
    numero_caja      VARCHAR(30) UNIQUE,
    origen           VARCHAR(20),
    creado           TIMESTAMP
);

CREATE TABLE estandares_peso (
    id               SERIAL PRIMARY KEY,
    producto         VARCHAR(150),
    codigo_producto  DOUBLE PRECISION,
    tipo             VARCHAR(10),
    peso_objetivo    DOUBLE PRECISION,
    desv_est         DOUBLE PRECISION,
    min_peso         DOUBLE PRECISION,
    max_peso         DOUBLE PRECISION,
    n_cajas          DOUBLE PRECISION,
    kg_total         DOUBLE PRECISION,
    fecha_calculo    TIMESTAMP
);

CREATE TABLE inventario_sap (
    id                SERIAL PRIMARY KEY,
    numero_caja       VARCHAR(30),
    kg                DOUBLE PRECISION,
    fecha_produccion  TIMESTAMP,
    descripcion       VARCHAR(120),
    fecha_foto        TIMESTAMP
);

CREATE TABLE pedidos (
    id            SERIAL PRIMARY KEY,
    fecha_pedido  TIMESTAMP,
    cliente       VARCHAR(100),
    producto      VARCHAR(150),
    cantidad      DOUBLE PRECISION,
    unidad        VARCHAR(10),
    estado        VARCHAR(20),
    creado        TIMESTAMP
);

CREATE TABLE pedidos_combos (
    id               SERIAL PRIMARY KEY,
    fecha            TIMESTAMP,
    cliente          VARCHAR(30),
    codigo_producto  DOUBLE PRECISION,
    producto         VARCHAR(150),
    combos_pedidos   DOUBLE PRECISION,
    kg_pedidos       DOUBLE PRECISION,
    origen           VARCHAR(20),
    creado           TIMESTAMP
);

CREATE TABLE programa_semanal (
    id            SERIAL PRIMARY KEY,
    fecha         TIMESTAMP,
    material_h    DOUBLE PRECISION,
    producto      VARCHAR(150),
    pedido_kg     DOUBLE PRECISION,
    pedido_cajas  DOUBLE PRECISION,
    formato       VARCHAR(30),
    cliente       VARCHAR(60),
    destino       VARCHAR(40),
    estado        VARCHAR(30),
    fecha_carga   TIMESTAMP
);

CREATE INDEX idx_cajas_fecha_desposte ON cajas (fecha_desposte);
CREATE INDEX idx_despachos_fecha ON despachos (fecha_despacho);
