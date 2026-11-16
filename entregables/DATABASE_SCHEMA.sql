-- =====================================================================
-- ATLAS COLOMBIA 2027 · Base maestra de inteligencia electoral
-- Motor: PostgreSQL 15+ con PostGIS 3+
-- Capas: raw -> staging -> clean (dimensiones/hechos observados) -> analytics (indicadores)
--        -> predictive (pronósticos versionados) -> presentation (vistas para Excel/tableros)
-- Carga inicial: CSV de data/clean/ generados por etl/01_build_data.py (COPY ... CSV HEADER).
-- =====================================================================
CREATE EXTENSION IF NOT EXISTS postgis;
CREATE SCHEMA IF NOT EXISTS raw;
CREATE SCHEMA IF NOT EXISTS staging;
CREATE SCHEMA IF NOT EXISTS clean;
CREATE SCHEMA IF NOT EXISTS analytics;
CREATE SCHEMA IF NOT EXISTS predictive;
CREATE SCHEMA IF NOT EXISTS presentation;

-- ---------------------------------------------------------------- FUENTES Y LOG
CREATE TABLE clean.dim_fuentes (
    fuente_id            text PRIMARY KEY,
    institucion          text NOT NULL,
    nombre_documento     text NOT NULL,
    url                  text,
    tipo_fuente          text NOT NULL,
    fecha_publicacion    text,
    fecha_consulta       date NOT NULL,
    cobertura_geografica text,
    nivel_confiabilidad  text CHECK (nivel_confiabilidad IN ('Alta','Media-alta','Media','Baja','Según insumos')),
    observaciones        text
);

CREATE TABLE clean.log_actualizaciones (
    log_id         bigserial PRIMARY KEY,
    tabla          text NOT NULL,
    operacion      text NOT NULL CHECK (operacion IN ('INSERT','UPDATE','DELETE','RELOAD')),
    registros      integer,
    fuente_id      text REFERENCES clean.dim_fuentes,
    ejecutado_por  text,
    ejecutado_en   timestamptz NOT NULL DEFAULT now(),
    detalle        text
);

-- ---------------------------------------------------------------- TERRITORIO
CREATE TABLE clean.dim_departamentos (
    departamento_id    char(2) PRIMARY KEY,                  -- DIVIPOLA
    nombre_departamento text NOT NULL,
    tipo               text NOT NULL CHECK (tipo IN ('Departamento','Distrito Capital')),
    capital_divipola   char(5),
    geom               geometry(MultiPolygon, 4326)
);

CREATE TABLE clean.dim_municipios (
    municipio_id            text PRIMARY KEY,                 -- 'MUN-' || codigo_divipola
    codigo_divipola         char(5) NOT NULL UNIQUE,
    nombre_municipio        text NOT NULL,
    departamento_id         char(2) NOT NULL REFERENCES clean.dim_departamentos,
    tipo_entidad_territorial text NOT NULL,                   -- Municipio | Distrito | Área no municipalizada | Isla
    es_distrito_especial    text,
    denominacion_distrito   text,
    elige_alcalde           boolean NOT NULL,
    categoria_municipal     text,                             -- Ley 617/2000 (pendiente de fuente CGN)
    lat                     numeric(9,6),
    lon                     numeric(9,6),
    area_km2_aprox          numeric,
    geom                    geometry(MultiPolygon, 4326),
    vigente_desde           date,
    vigente_hasta           date,
    CHECK (left(codigo_divipola, 2) = departamento_id)
);
ALTER TABLE clean.dim_departamentos ADD FOREIGN KEY (capital_divipola) REFERENCES clean.dim_municipios (codigo_divipola) DEFERRABLE INITIALLY DEFERRED;
CREATE INDEX ON clean.dim_municipios (departamento_id);
CREATE INDEX ON clean.dim_municipios USING gist (geom);

CREATE TABLE clean.dim_distritos (          -- localidades de Bogotá, comunas, distritos internos
    distrito_id     text PRIMARY KEY,
    municipio_id    text NOT NULL REFERENCES clean.dim_municipios,
    tipo            text NOT NULL,          -- Localidad | Comuna | Corregimiento | Zona
    codigo          text,
    nombre          text NOT NULL,
    geom            geometry(MultiPolygon, 4326)
);
CREATE TABLE clean.dim_localidades (LIKE clean.dim_distritos INCLUDING ALL);  -- alias semántico solicitado

CREATE TABLE clean.dim_circunscripciones (
    circunscripcion_id text PRIMARY KEY,
    dignidad           text NOT NULL,
    nivel              text NOT NULL,       -- Departamental | Municipal | Local
    departamento_id    char(2) REFERENCES clean.dim_departamentos,
    municipio_id       text REFERENCES clean.dim_municipios,
    distrito_id        text REFERENCES clean.dim_distritos
);

CREATE TABLE clean.dim_puestos_votacion (
    puesto_id       text PRIMARY KEY,
    municipio_id    text NOT NULL REFERENCES clean.dim_municipios,
    zona            text,
    codigo_puesto   text NOT NULL,
    nombre_puesto   text,
    direccion       text,
    geom            geometry(Point, 4326),  -- sólo con geocodificación verificada
    fuente_id       text REFERENCES clean.dim_fuentes,
    UNIQUE (municipio_id, zona, codigo_puesto)
);
CREATE TABLE clean.dim_mesas_votacion (
    mesa_id         text PRIMARY KEY,
    puesto_id       text NOT NULL REFERENCES clean.dim_puestos_votacion,
    numero_mesa     integer NOT NULL,
    anio            smallint NOT NULL,
    UNIQUE (puesto_id, numero_mesa, anio)
);

-- ---------------------------------------------------------------- ORGANIZACIONES Y PERSONAS
CREATE TABLE clean.dim_partidos (            -- partidos y movimientos con personería (histórica o vigente)
    organizacion_id   text PRIMARY KEY,      -- 'ORG-' || clave canónica
    org_key           text NOT NULL UNIQUE,
    nombre_oficial    text NOT NULL,
    sigla             text,
    tipo              text NOT NULL,
    personeria_juridica text,
    estado_legal_2026 text,
    fecha_reconocimiento date,
    ambito            text,
    orientacion_politica_documentada text,  -- sólo con fuente
    fuente_id         text
);
CREATE TABLE clean.hist_partidos (           -- fusiones, cambios de nombre, pérdida de personería
    organizacion_id   text REFERENCES clean.dim_partidos,
    fecha             date NOT NULL,
    evento            text NOT NULL,         -- Reconocimiento | Cambio de nombre | Fusión | Pérdida de personería | Recuperación
    organizacion_relacionada text,
    fuente_id         text REFERENCES clean.dim_fuentes,
    PRIMARY KEY (organizacion_id, fecha, evento)
);
CREATE TABLE clean.dim_movimientos (         -- coaliciones, GSC y organizaciones locales, una por territorio y elección
    organizacion_id   text PRIMARY KEY,      -- 'LOC-' || hash
    nombre            text NOT NULL,
    tipo              text NOT NULL,
    ambito            text,
    dignidad          text,
    departamento_id   char(2) REFERENCES clean.dim_departamentos,
    codigo_divipola   char(5),
    primer_anio       smallint,
    fuente_id         text
);
CREATE VIEW clean.dim_grupos_significativos_ciudadanos AS
    SELECT * FROM clean.dim_movimientos WHERE tipo = 'Grupo significativo de ciudadanos';

CREATE TABLE clean.dim_coaliciones (
    coalicion_id       text PRIMARY KEY,
    nombre_coalicion   text NOT NULL,
    dignidad           text NOT NULL,
    anio               smallint NOT NULL,
    departamento_id    char(2) REFERENCES clean.dim_departamentos,
    codigo_divipola    char(5),
    candidato_respaldado text,
    fecha_conformacion date,
    estado             text,
    fuente_id          text
);
CREATE TABLE clean.bridge_coalicion_organizacion (   -- integrantes (un partido puede aliarse distinto en cada municipio)
    coalicion_id     text REFERENCES clean.dim_coaliciones,
    organizacion_id  text,                          -- ORG-* o LOC-*
    rol              text DEFAULT 'Integrante',     -- Integrante | Avalista
    fuente_id        text,
    PRIMARY KEY (coalicion_id, organizacion_id)
);

CREATE TABLE clean.dim_candidatos (
    candidato_id    text PRIMARY KEY,               -- hash del nombre normalizado (no es identidad verificada)
    nombre_completo text NOT NULL,
    documento_hash  text,                           -- si se obtiene cédula, guardar sólo hash
    observaciones   text
);

CREATE TABLE clean.dim_dignidades (
    dignidad      text PRIMARY KEY CHECK (dignidad IN ('ALC','GOB','CON','ASA','JAL')),
    nombre        text NOT NULL,
    nivel         text NOT NULL
);
INSERT INTO clean.dim_dignidades VALUES
 ('ALC','Alcaldía municipal/distrital','Municipal'), ('GOB','Gobernación','Departamental'),
 ('CON','Concejo municipal/distrital','Municipal'), ('ASA','Asamblea departamental','Departamental'),
 ('JAL','Juntas Administradoras Locales','Local');

CREATE TABLE clean.dim_elecciones (
    eleccion_id     text PRIMARY KEY,               -- ELEC-<dignidad>-<año>
    anio            smallint NOT NULL,
    fecha           date NOT NULL,
    dignidad        text NOT NULL REFERENCES clean.dim_dignidades,
    tipo            text NOT NULL,                  -- Ordinaria territorial | Atípica
    periodo         text,
    estado          text
);

-- ---------------------------------------------------------------- HECHOS OBSERVADOS (separados por dignidad)
CREATE TABLE clean.fact_candidaturas (
    candidatura_id   text PRIMARY KEY,
    eleccion_id      text NOT NULL REFERENCES clean.dim_elecciones,
    candidato_id     text NOT NULL REFERENCES clean.dim_candidatos,
    municipio_id     text REFERENCES clean.dim_municipios,
    departamento_id  char(2) NOT NULL REFERENCES clean.dim_departamentos,
    organizacion_id  text,                          -- partido, movimiento o GSC
    coalicion_id     text REFERENCES clean.dim_coaliciones,
    tipo_apoyo       text,                          -- Aval individual | Coalición | GSC | Movimiento
    numero_lista     text,
    estado_candidatura text NOT NULL CHECK (estado_candidatura IN ('Actor político identificado','Aspirante','Precandidato',
                         'Candidato anunciado','Candidato inscrito','Candidato en firme','Candidato revocado','Candidato retirado')),
    fecha_inscripcion date,
    es_alcalde_actual boolean,
    plan_gobierno_url text,
    fuente_id        text NOT NULL,
    fecha_verificacion date NOT NULL
);

CREATE TABLE clean.fact_resultados_alcaldia (
    resultado_id     text PRIMARY KEY,
    eleccion_id      text NOT NULL REFERENCES clean.dim_elecciones,
    municipio_id     text NOT NULL,                 -- puede apuntar a códigos históricos suprimidos
    codigo_divipola  char(5) NOT NULL,
    departamento_id  char(2) NOT NULL REFERENCES clean.dim_departamentos,
    candidato_id     text NOT NULL REFERENCES clean.dim_candidatos,
    organizacion_id  text NOT NULL,
    votos            integer NOT NULL CHECK (votos >= 0),
    votos_validos    integer,
    porcentaje_votos numeric(6,4) CHECK (porcentaje_votos BETWEEN 0 AND 1),
    posicion         smallint,
    ganador          boolean NOT NULL,
    etapa_resultado  text NOT NULL DEFAULT 'Resultado declarado'
                     CHECK (etapa_resultado IN ('Preconteo','Escrutinio','Resultado declarado','Resultado definitivo','No declarado')),
    fuente_id        text NOT NULL,
    estado_validacion text
);
CREATE INDEX ON clean.fact_resultados_alcaldia (eleccion_id, codigo_divipola);
CREATE INDEX ON clean.fact_resultados_alcaldia (organizacion_id);
CREATE INDEX ON clean.fact_resultados_alcaldia (candidato_id);

CREATE TABLE clean.fact_resultados_gobernacion (LIKE clean.fact_resultados_alcaldia INCLUDING ALL);
ALTER TABLE clean.fact_resultados_gobernacion ALTER COLUMN municipio_id DROP NOT NULL, ALTER COLUMN codigo_divipola DROP NOT NULL;
CREATE TABLE clean.fact_resultados_concejo  (LIKE clean.fact_resultados_alcaldia INCLUDING ALL);
ALTER TABLE clean.fact_resultados_concejo ADD COLUMN curules smallint, ADD COLUMN voto_preferente boolean;
CREATE TABLE clean.fact_resultados_asamblea (LIKE clean.fact_resultados_concejo INCLUDING ALL);
CREATE TABLE clean.fact_resultados_jal      (LIKE clean.fact_resultados_concejo INCLUDING ALL);
ALTER TABLE clean.fact_resultados_jal ADD COLUMN distrito_id text REFERENCES clean.dim_distritos;

CREATE TABLE clean.fact_resultados_mesa (
    mesa_resultado_id text PRIMARY KEY,
    eleccion_id      text NOT NULL REFERENCES clean.dim_elecciones,
    mesa_id          text NOT NULL REFERENCES clean.dim_mesas_votacion,
    candidato_id     text REFERENCES clean.dim_candidatos,   -- NULL para blanco/nulo/no marcado
    tipo_voto        text NOT NULL CHECK (tipo_voto IN ('CANDIDATO','LISTA','BLANCO','NULOS','NO_MARCADOS')),
    organizacion_id  text,
    votos            integer NOT NULL CHECK (votos >= 0),
    etapa_resultado  text NOT NULL CHECK (etapa_resultado IN ('Preconteo','Escrutinio','Resultado declarado','Resultado definitivo')),
    formulario       text,                                   -- E-14 | E-24 | E-26 ...
    id_formulario    text,
    fuente_id        text NOT NULL
);
CREATE INDEX ON clean.fact_resultados_mesa (eleccion_id, mesa_id);

CREATE TABLE clean.fact_participacion (
    dignidad         text NOT NULL REFERENCES clean.dim_dignidades,
    anio             smallint NOT NULL,
    departamento_id  char(2) NOT NULL REFERENCES clean.dim_departamentos,
    codigo_divipola  char(5),                                -- NULL = total departamental
    censo_electoral  integer,
    sufragantes      integer,
    validos          integer,
    votos_candidatos integer,
    votos_blancos    integer,
    votos_nulos      integer,
    votos_no_marcados integer,
    participacion    numeric(6,4),
    abstencion       numeric(6,4),
    fuente_id        text NOT NULL,
    nota             text,
    UNIQUE (dignidad, anio, departamento_id, codigo_divipola),
    CHECK (validos IS NULL OR votos_candidatos IS NULL OR votos_candidatos <= validos)
);

CREATE TABLE clean.fact_autoridades (          -- alcaldes y gobernadores por período (histórico de titulares)
    autoridad_id     text PRIMARY KEY,
    dignidad         text NOT NULL CHECK (dignidad IN ('ALC','GOB')),
    candidato_id     text NOT NULL REFERENCES clean.dim_candidatos,
    municipio_id     text REFERENCES clean.dim_municipios,
    departamento_id  char(2) NOT NULL REFERENCES clean.dim_departamentos,
    organizacion_id  text,
    fecha_inicio     date NOT NULL,
    fecha_fin        date,
    eleccion_origen  text REFERENCES clean.dim_elecciones,
    condicion        text NOT NULL,                -- Alcalde actual | Encargado | Elegido en atípica | ...
    fuente_id        text NOT NULL,
    fecha_verificacion date NOT NULL
);

CREATE TABLE clean.fact_encuestas (
    encuesta_id      text NOT NULL,
    ambito           text NOT NULL CHECK (ambito IN ('Nacional','Departamental','Municipal')),
    departamento_id  char(2) REFERENCES clean.dim_departamentos,
    municipio_id     text REFERENCES clean.dim_municipios,
    empresa_encuestadora text NOT NULL,
    fecha_inicio_campo date NOT NULL,
    fecha_fin_campo  date NOT NULL,
    fecha_publicacion date,
    tamano_muestra   integer,
    universo         text,
    metodologia      text,
    margen_error     numeric(5,2),
    nivel_confianza  numeric(5,2),
    candidato_id     text REFERENCES clean.dim_candidatos,
    organizacion_id  text,
    opcion           text NOT NULL,               -- nombre de la opción tal como figura
    intencion_voto   numeric(5,2),
    fuente_id        text NOT NULL,
    url              text,
    estado_verificacion text NOT NULL,
    PRIMARY KEY (encuesta_id, opcion),
    CHECK (fecha_fin_campo >= fecha_inicio_campo)
);

CREATE TABLE clean.fact_variables_territoriales (   -- formato largo: una fila por variable, territorio y año
    codigo_divipola  char(5) NOT NULL,
    variable         text NOT NULL,
    anio_referencia  smallint NOT NULL,
    valor            numeric,
    unidad           text NOT NULL,
    cobertura        text,
    metodologia      text,
    fuente_id        text NOT NULL REFERENCES clean.dim_fuentes,
    fecha_actualizacion date,
    PRIMARY KEY (codigo_divipola, variable, anio_referencia)
);
CREATE TABLE clean.fact_gestion_municipal (LIKE clean.fact_variables_territoriales INCLUDING ALL);  -- MDM, FUT, CUIPO, desempeño fiscal

CREATE TABLE clean.fact_noticias (
    noticia_id       text PRIMARY KEY,
    fecha            date,
    municipio_id     text REFERENCES clean.dim_municipios,
    departamento_id  char(2) REFERENCES clean.dim_departamentos,
    persona_relacionada text,
    organizacion_id  text,
    titulo           text NOT NULL,
    resumen          text,
    categoria        text NOT NULL,
    fuente_id        text,
    url              text,
    estado_verificacion text NOT NULL CHECK (estado_verificacion IN ('Confirmada (fuente oficial)','Información periodística','Declaración','Rumor','Por verificar'))
);

CREATE TABLE clean.dim_prospectos (
    prospecto_id     text PRIMARY KEY,
    nombre           text NOT NULL,
    codigo_territorio text NOT NULL,              -- 5 dígitos municipio o 2 dígitos departamento
    dignidad         text NOT NULL,
    partido_mencionado text,
    cargo_o_trayectoria text,
    estado           text NOT NULL,
    motivo_inclusion text,
    prioridad_investigacion text,
    fuente_id        text NOT NULL,
    fecha_actualizacion date NOT NULL
);

-- ---------------------------------------------------------------- ANALÍTICA (calculada, recalculable por ciclo)
CREATE TABLE analytics.fact_indicadores_electorales (
    dignidad         text NOT NULL,
    anio             smallint NOT NULL,
    codigo_divipola  char(5) NOT NULL,
    n_candidatos     smallint,
    organizacion_ganadora text,
    fuerza_ganadora  text,
    pct_ganador      numeric(6,4),
    margen_pts       numeric(6,2),
    nep              numeric(6,3),
    fragmentacion_rae numeric(6,4),
    concentracion_top2 numeric(6,4),
    participacion    numeric(6,4),
    var_participacion_pts numeric(6,2),
    volatilidad_pedersen numeric(6,2),
    alternancia      text,
    calculado_en     timestamptz NOT NULL DEFAULT now(),
    version_reglas   text NOT NULL DEFAULT 'v1',
    PRIMARY KEY (dignidad, anio, codigo_divipola, version_reglas)
);
CREATE TABLE analytics.config_umbrales (
    umbral      text PRIMARY KEY,                 -- U_MUY_COMPETITIVO, U_COMPETITIVO, U_MODERADO, U_NEP_FRAGMENTADO, U_SWING
    valor       numeric NOT NULL,
    descripcion text
);
INSERT INTO analytics.config_umbrales VALUES
 ('U_MUY_COMPETITIVO',5,'Margen máximo muy competitivo (pts)'), ('U_COMPETITIVO',10,'Margen máximo competitivo (pts)'),
 ('U_MODERADO',20,'Margen máximo moderado (pts)'), ('U_NEP_FRAGMENTADO',3.5,'NEP mínimo fragmentado'), ('U_SWING',10,'Margen máximo para swing (pts)');
CREATE TABLE analytics.fact_clasificacion_territorial (
    codigo_divipola  char(5) NOT NULL,
    ciclo            smallint NOT NULL,
    dignidad         text NOT NULL,
    clasificacion_bastion text,
    fuerza_dominante text,
    clasificacion_competitividad text,
    version_reglas   text NOT NULL,
    PRIMARY KEY (codigo_divipola, ciclo, dignidad, version_reglas)
);

-- ---------------------------------------------------------------- PREDICTIVO (nunca se sobrescribe)
CREATE TABLE predictive.fact_pronosticos (
    forecast_id      text PRIMARY KEY,
    dignidad         text NOT NULL CHECK (dignidad IN ('ALC','GOB')),   -- modelos separados por dignidad
    municipio_id     text,
    departamento_id  char(2) NOT NULL,
    fecha_corte      date NOT NULL,
    fecha_generacion timestamptz NOT NULL DEFAULT now(),
    version_modelo   text NOT NULL,
    candidato_id     text,
    organizacion_id  text,
    coalicion_id     text,
    porcentaje_estimado numeric(6,4),
    probabilidad_victoria numeric(6,4),
    probabilidad_retencion_fuerza numeric(6,4),
    intervalo_inferior numeric(6,4),
    intervalo_superior numeric(6,4),
    margen_estimado  numeric(6,2),
    clasificacion_competitividad text,
    nivel_confianza  text,
    escenario        text NOT NULL,
    tipo_dato        text NOT NULL DEFAULT 'Estimación estadística' CHECK (tipo_dato IN ('Estimación estadística','Escenario hipotético')),
    variables_utilizadas text,
    fuente_datos     text,
    observaciones    text
);
CREATE INDEX ON predictive.fact_pronosticos (version_modelo, fecha_corte);
CREATE RULE no_update_pronosticos AS ON UPDATE TO predictive.fact_pronosticos DO INSTEAD NOTHING;
CREATE RULE no_delete_pronosticos AS ON DELETE TO predictive.fact_pronosticos DO INSTEAD NOTHING;

-- ---------------------------------------------------------------- PRESENTACIÓN
CREATE VIEW presentation.v_municipios_maestro AS
SELECT m.municipio_id, m.codigo_divipola, m.nombre_municipio, d.nombre_departamento, m.tipo_entidad_territorial,
       a.candidato_id AS alcalde_candidato_id, c.nombre_completo AS alcalde_actual, a.organizacion_id AS organizacion_gobernante
FROM clean.dim_municipios m
JOIN clean.dim_departamentos d USING (departamento_id)
LEFT JOIN clean.fact_autoridades a ON a.municipio_id = m.municipio_id AND a.dignidad = 'ALC' AND a.fecha_fin IS NULL
LEFT JOIN clean.dim_candidatos c ON c.candidato_id = a.candidato_id;
