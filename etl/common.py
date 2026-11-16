"""Utilidades compartidas del ETL de ATLAS COLOMBIA 2027."""
import hashlib, re, unicodedata, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, 'raw')
CLEAN = os.path.join(ROOT, 'data', 'clean')
OUT = os.path.join(ROOT, 'entregables')
FECHA_CORTE = '2026-11-15'
ELECCION_OBJETIVO = '2027-10-31'


def sin_tildes(s):
    if s is None:
        return ''
    s = str(s)
    for bad, good in (('Ã¡', 'á'), ('Ã©', 'é'), ('Ã\xad', 'í'), ('Ã³', 'ó'), ('Ãº', 'ú'), ('Ã±', 'ñ')):
        s = s.replace(bad, good)
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if not unicodedata.combining(c))


def norm(s):
    """Mayúsculas, sin tildes, sin signos, espacios simples."""
    s = sin_tildes(s).upper().replace('�', '')
    s = re.sub(r'[^A-Z0-9 ]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def name_key(nombre):
    """Clave de persona independiente del orden nombres/apellidos (se usa sólo para enlazar candidaturas
    entre años; es una coincidencia por nombre, no una identidad verificada)."""
    toks = sorted(t for t in norm(nombre).split() if len(t) > 1)
    return ' '.join(toks)


def hid(prefix, *parts, n=10):
    h = hashlib.sha1('|'.join(str(p) for p in parts).encode('utf-8')).hexdigest()[:n].upper()
    return f'{prefix}-{h}'


# --- Organizaciones nacionales: clave canónica, nombre de referencia y sigla -------------------
# El orden importa (NUEVO LIBERALISMO antes que LIBERAL, etc.).
ORG_RULES = [
    ('NUEVO_LIBERALISMO', r'NUEVO LIBERALISMO', 'Partido Nuevo Liberalismo', 'NL'),
    ('LIBERAL', r'PARTIDO LIBERAL|^LIBERAL COLOMBIANO|^PARTIDO LIBERAL', 'Partido Liberal Colombiano', 'PLC'),
    ('CONSERVADOR', r'CONSERVADOR', 'Partido Conservador Colombiano', 'PCC'),
    ('LA_U', r'PARTIDO DE LA U\b|UNIDAD NACIONAL|UNION POR LA GENTE|PARTIDO SOCIAL DE UNIDAD|^PARTIDO DE LA U$|^LA U$', 'Partido de la Unión por la Gente (La U)', 'U'),
    ('CAMBIO_RADICAL', r'CAMBIO RADICAL', 'Partido Cambio Radical', 'CR'),
    ('CENTRO_DEMOCRATICO', r'CENTRO DEMOCRATICO', 'Partido Centro Democrático', 'CD'),
    ('ALIANZA_VERDE', r'ALIANZA VERDE|^PARTIDO VERDE$|PARTIDO VERDE OPCION CENTRO', 'Partido Alianza Verde', 'AV'),
    ('PACTO_HISTORICO', r'PACTO HISTORICO', 'Pacto Histórico', 'PH'),
    ('COLOMBIA_HUMANA_UP', r'COLOMBIA HUMANA|UNION PATRIOTICA', 'Colombia Humana – Unión Patriótica', 'CH-UP'),
    ('POLO', r'POLO DEMOCRATICO', 'Polo Democrático Alternativo', 'PDA'),
    ('COMUNES', r'\bCOMUNES\b|FUERZA ALTERNATIVA REVOLUCIONARIA', 'Partido Comunes (antes FARC)', 'COMUNES'),
    ('MIRA', r'\bMIRA\b', 'Partido Político MIRA', 'MIRA'),
    ('CJL', r'COLOMBIA JUSTA', 'Partido Colombia Justa Libres', 'CJL'),
    ('ASI', r'ALIANZA SOCIAL INDEPENDIENTE|ALIANZA SOCIAL INDIGENA|\bASI\b', 'Alianza Social Independiente', 'ASI'),
    ('AICO', r'AUTORIDADES INDIGENAS', 'Movimiento Autoridades Indígenas de Colombia', 'AICO'),
    ('MAIS', r'ALTERNATIVO INDIGENA', 'Movimiento Alternativo Indígena y Social', 'MAIS'),
    ('OPCION_CIUDADANA', r'OPCION CIUDADANA|INTEGRACION NACIONAL|CONVERGENCIA CIUDADANA|ALAS EQUIPO', 'Opción Ciudadana / PIN / Convergencia (linaje)', 'OC'),
    ('COLOMBIA_RENACIENTE', r'COLOMBIA RENACIENTE|AFROVIDES', 'Partido Colombia Renaciente', 'CR-AFRO'),
    ('ADA', r'PARTIDO ADA|ALIANZA DEMOCRATICA AMPLIA', 'Alianza Democrática Amplia (ADA)', 'ADA'),
    ('CREEMOS', r'\bCREEMOS\b', 'Partido Creemos', 'CREEMOS'),
    ('FUERZA_DE_LA_PAZ', r'FUERZA DE LA PAZ', 'Partido La Fuerza de la Paz', 'FDP'),
    ('EN_MARCHA', r'EN MARCHA', 'Agrupación Política En Marcha', 'EM'),
    ('DIGNIDAD', r'DIGNIDAD', 'Dignidad & Compromiso', 'DyC'),
    ('SALVACION_NACIONAL', r'SALVACION NACIONAL', 'Movimiento de Salvación Nacional', 'MSN'),
    ('GENTE_EN_MOVIMIENTO', r'GENTE EN MOVIMIENTO', 'Partido Gente en Movimiento', 'GEM'),
    ('VERDE_OXIGENO', r'VERDE OXIGENO|OXIGENO VERDE', 'Partido Verde Oxígeno', 'VO'),
    ('DEMOCRATA', r'PARTIDO DEMOCRATA', 'Partido Demócrata Colombiano', 'PDC'),
    ('ECOLOGISTA', r'ECOLOGISTA', 'Partido Ecologista Colombiano', 'PEC'),
    ('FUERZA_CIUDADANA', r'FUERZA CIUDADANA', 'Movimiento Fuerza Ciudadana', 'FC'),
    ('LIGA', r'LIGA DE GOBERNANTES', 'Liga de Gobernantes Anticorrupción', 'LIGA'),
    ('INDEPENDIENTES', r'^INDEPENDIENTES$|PARTIDO INDEPENDIENTES', 'Partido Independientes', 'IND'),
    ('DEFENSORES_PATRIA', r'DEFENSORES DE LA PATRIA', 'Defensores de la Patria', 'DP'),
    ('PRE', r'REIVINDICACION ETNICA', 'Partido de Reivindicación Étnica (PRE)', 'PRE'),
    ('MSN_AFRO', r'MOVIMIENTO SOCIAL INDEPENDIENTE|MOVIMIENTO NACIONAL AFROCOLOMBIANO', 'Movimientos afro / sociales (varios)', 'AFRO'),
]
ORG_NAME = {k: n for k, _, n, _ in ORG_RULES}
ORG_SIGLA = {k: s for k, _, _, s in ORG_RULES}

# Partidos con personería jurídica vigente según El Espectador, 05-08-2026 (21 organizaciones).
VIGENTES_2026 = {'PACTO_HISTORICO', 'CENTRO_DEMOCRATICO', 'LIBERAL', 'CONSERVADOR', 'LA_U', 'CAMBIO_RADICAL',
                 'ALIANZA_VERDE', 'AICO', 'MAIS', 'MIRA', 'NUEVO_LIBERALISMO', 'DIGNIDAD', 'SALVACION_NACIONAL',
                 'DEFENSORES_PATRIA', 'CJL', 'ADA', 'LIGA', 'EN_MARCHA', 'ASI', 'COLOMBIA_RENACIENTE', 'DEMOCRATA'}

SPECIAL = {'VOTOS EN BLANCO': 'BLANCO', 'VOTOS NULOS': 'NULOS', 'VOTOS NO MARCADOS': 'NO_MARCADOS',
           'TARJETAS NO MARCADAS': 'NO_MARCADOS', 'CANDIDATOS TOTALES': 'TOTAL'}


def org_tipo(nombre):
    n = norm(nombre)
    if not n:
        return 'Sin dato'
    if n.startswith('COALICION') or ' COALICION ' in f' {n} ':
        return 'Coalición'
    if n.startswith('GSC') or 'GRUPO SIGNIFICATIVO' in n or n.startswith('G S C'):
        return 'Grupo significativo de ciudadanos'
    if org_key(nombre):
        return 'Partido / movimiento con personería (histórica o vigente)'
    if n.startswith('MOVIMIENTO'):
        return 'Movimiento político o social'
    if n.startswith('PARTIDO'):
        return 'Partido (no catalogado)'
    return 'Organización local / sin clasificar'


def org_key(nombre):
    """Clave canónica de un partido nacional; None si es coalición, GSC u organización local."""
    n = norm(nombre)
    if not n or n.startswith('COALICION') or n.startswith('GSC') or 'GRUPO SIGNIFICATIVO' in n:
        return None
    for k, rx, _, _ in ORG_RULES:
        if re.search(rx, n):
            return k
    return None


def fuerza(nombre):
    """Etiqueta analítica de la organización ganadora: el partido canónico si existe; si no, el tipo."""
    k = org_key(nombre)
    if k:
        return k
    t = org_tipo(nombre)
    return {'Coalición': 'COALICION', 'Grupo significativo de ciudadanos': 'GSC'}.get(t, 'OTRA_LOCAL')

# Regiones naturales (asignación por departamento; algunos departamentos abarcan más de una región natural).
REGIONES = {'Caribe': ['08', '13', '20', '23', '44', '47', '70'], 'Insular': ['88'], 'Pacífica': ['19', '27', '52', '76'],
            'Andina': ['05', '11', '15', '17', '25', '41', '54', '63', '66', '68', '73'], 'Orinoquía': ['50', '81', '85', '99'],
            'Amazonía': ['18', '86', '91', '94', '95', '97']}
REGION_OF = {d: r for r, ds in REGIONES.items() for d in ds}
