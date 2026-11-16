"""ATLAS COLOMBIA 2027 – Paso 1: RAW -> CLEAN.

Lee las fuentes descargadas en raw/ y deja tablas normalizadas en data/clean/*.csv.
Uso: python etl/01_build_data.py
"""
import json, os, re, difflib, math
import numpy as np
import pandas as pd
from common import *

os.makedirs(CLEAN, exist_ok=True)
LOG = []
def log(*a):
    s = ' '.join(str(x) for x in a); print(s, flush=True); LOG.append(s)
def save(df, name):
    df.to_csv(os.path.join(CLEAN, name + '.csv'), index=False, encoding='utf-8-sig'); log(f'  {name}: {len(df):,} filas')

ANIOS = [2003, 2007, 2011, 2015, 2019, 2023]
FECHAS = {2003: '2003-10-26', 2007: '2007-10-28', 2011: '2011-10-30', 2015: '2015-10-25', 2019: '2019-10-27', 2023: '2023-10-29'}

# =============================================================== 1. GEOGRAFÍA
dv = pd.DataFrame(json.load(open(os.path.join(RAW, 'divipola.json'), encoding='utf-8')))
dv['lat'] = dv.latitud.str.replace(',', '.').astype(float)
dv['lon'] = dv.longitud.str.replace(',', '.').astype(float)
DEP_NAME_FIX = {'BOGOTÁ, D.C.': 'BOGOTÁ, D.C.', 'ARCHIPIÉLAGO DE SAN ANDRÉS, PROVIDENCIA Y SANTA CATALINA': 'SAN ANDRÉS, PROVIDENCIA Y SANTA CATALINA'}
dv['dpto'] = dv.dpto.replace(DEP_NAME_FIX)

# Distritos: lista preliminar documentada; se marca para verificación normativa.
DISTRITOS = {'11001': 'Distrito Capital (Constitución, art. 322)', '08001': 'Distrito Especial, Industrial y Portuario',
             '13001': 'Distrito Turístico y Cultural', '47001': 'Distrito Turístico, Cultural e Histórico',
             '76109': 'Distrito Especial, Industrial, Portuario, Biodiverso y Ecoturístico',
             '44001': 'Distrito Especial, Turístico y Cultural', '13468': 'Distrito Especial, Turístico, Cultural e Histórico',
             '05837': 'Distrito Especial Portuario, Logístico, Industrial, Turístico y Comercial',
             '76001': 'Distrito Especial, Deportivo, Cultural, Turístico, Empresarial y de Servicios',
             '05001': 'Distrito Especial de Ciencia, Tecnología e Innovación', '68081': 'Distrito Especial Portuario, Biodiverso, Industrial y Turístico',
             '52835': 'Distrito Especial, Industrial, Portuario, Biodiverso y Ecoturístico'}

# Población DANE
pp = pd.read_excel(os.path.join(RAW, 'dane_proy_area.xlsx'), sheet_name=2, header=7, dtype={'DP': str, 'MPIO': str})
pp.columns = ['dp', 'dpnom', 'mpio', 'mpnom', 'anio', 'area', 'pob']
pp = pp.dropna(subset=['anio']); pp['anio'] = pp.anio.astype(int)
pp = pp[pp.anio.isin([2018, 2026, 2027])]
pp['area'] = pp.area.map({'Total': 'tot', 'Cabecera Municipal': 'cab', 'Centros Poblados y Rural Disperso': 'rur'})
pw = pp.pivot_table(index='mpio', columns=['area', 'anio'], values='pob', aggfunc='sum')
pw.columns = [f'pob_{a}_{y}' for a, y in pw.columns]
pw = pw.reset_index()

# NBI 2018
nbi = pd.read_excel(os.path.join(RAW, 'dane_nbi_2021.xlsx'), sheet_name='Municipios', header=None, skiprows=10, dtype={0: str, 2: str})
nbi = nbi[[0, 2, 4, 5]].dropna(subset=[2]); nbi.columns = ['dp', 'mp', 'nbi_2018', 'miseria_2018']
nbi['mpio'] = nbi.dp.str.zfill(2) + nbi.mp.str.zfill(3)
nbi = nbi[nbi.mpio.str.fullmatch(r'\d{5}')][['mpio', 'nbi_2018', 'miseria_2018']]
nbi[['nbi_2018', 'miseria_2018']] = nbi[['nbi_2018', 'miseria_2018']].apply(pd.to_numeric, errors='coerce')

# IPM 2018
ipm = pd.read_excel(os.path.join(RAW, 'dane_ipm_mun.xlsx'), sheet_name='IPM_Municipios', header=None, skiprows=16, dtype={0: str, 1: str})
ipm = ipm[[1, 3, 4, 5]]; ipm.columns = ['mpio', 'ipm_2018', 'ipm_cab_2018', 'ipm_rur_2018']
ipm = ipm[ipm.mpio.astype(str).str.fullmatch(r'\d{5}')]
ipm[['ipm_2018', 'ipm_cab_2018', 'ipm_rur_2018']] = ipm[['ipm_2018', 'ipm_cab_2018', 'ipm_rur_2018']].apply(pd.to_numeric, errors='coerce')

# MDM 2020
mdm = pd.DataFrame(json.load(open(os.path.join(RAW, 'mdm2020.json'), encoding='utf-8')))
mg = mdm[mdm.indicador == 'Grupo de dotaciones iniciales'][['codigo_entidad', 'etiqueta']].rename(columns={'etiqueta': 'mdm_grupo_2020'})
ms = mdm[mdm.indicador == 'MDM'][['codigo_entidad', 'dato']].rename(columns={'dato': 'mdm_puntaje_2020'})
mdmw = mg.merge(ms, on='codigo_entidad', how='outer').rename(columns={'codigo_entidad': 'mpio'})
mdmw['mdm_puntaje_2020'] = pd.to_numeric(mdmw.mdm_puntaje_2020, errors='coerce')

# Superficie y geometría (MGN vía Esri, simplificada)
import geopandas as gpd
g = gpd.read_file(os.path.join(RAW, 'mun_2023.geojson'))
g = g.rename(columns={'codigoMunicipalMGNCompleto': 'mpio'})[['mpio', 'geometry']]
g['area_km2_aprox'] = g.to_crs(9377).area / 1e6
os.makedirs(OUT, exist_ok=True)
gg = g.merge(dv[['cod_mpio', 'nom_mpio', 'dpto']], left_on='mpio', right_on='cod_mpio', how='left').drop(columns='cod_mpio')
gg.to_file(os.path.join(OUT, 'MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson'), driver='GeoJSON')
area = g[['mpio', 'area_km2_aprox']]

mun = dv.rename(columns={'cod_mpio': 'codigo_divipola', 'nom_mpio': 'nombre_municipio', 'cod_dpto': 'departamento_id', 'dpto': 'nombre_departamento'})
mun['municipio_id'] = 'MUN-' + mun.codigo_divipola
mun['region'] = mun.departamento_id.map(REGION_OF)
mun['tipo_entidad_territorial'] = mun.tipo_municipio.map({'Municipio': 'Municipio', 'Área no municipalizada': 'Área no municipalizada', 'Isla': 'Isla (sin alcaldía)'}).fillna(mun.tipo_municipio)
mun.loc[mun.codigo_divipola.isin(DISTRITOS), 'tipo_entidad_territorial'] = 'Distrito'
mun['es_distrito_especial'] = np.where(mun.codigo_divipola.isin(DISTRITOS), 'Sí (verificar norma)', 'No')
mun['denominacion_distrito'] = mun.codigo_divipola.map(DISTRITOS).fillna('')
mun['elige_alcalde'] = np.where(mun.tipo_municipio == 'Municipio', 'Sí', 'No')
mun = (mun.merge(pw, left_on='codigo_divipola', right_on='mpio', how='left').drop(columns='mpio')
          .merge(area, left_on='codigo_divipola', right_on='mpio', how='left').drop(columns='mpio')
          .merge(nbi, left_on='codigo_divipola', right_on='mpio', how='left').drop(columns='mpio')
          .merge(ipm, left_on='codigo_divipola', right_on='mpio', how='left').drop(columns='mpio')
          .merge(mdmw, left_on='codigo_divipola', right_on='mpio', how='left').drop(columns='mpio'))
mun['densidad_2026'] = (mun.pob_tot_2026 / mun.area_km2_aprox).round(1)
mun['pct_rural_2026'] = (mun.pob_rur_2026 / mun.pob_tot_2026).round(4)
CODES = set(mun.codigo_divipola)
NAME = dict(zip(mun.codigo_divipola, mun.nombre_municipio))
DEPN = dict(zip(mun.departamento_id, mun.nombre_departamento))

dep = (mun.groupby(['departamento_id', 'nombre_departamento'])
          .agg(n_municipios=('elige_alcalde', lambda s: (s == 'Sí').sum()), n_areas_no_municipalizadas=('tipo_municipio', lambda s: (s == 'Área no municipalizada').sum()),
               pob_tot_2026=('pob_tot_2026', 'sum'), pob_tot_2027=('pob_tot_2027', 'sum'), area_km2_aprox=('area_km2_aprox', 'sum'))
          .reset_index())
CAPITAL = {'11': '11001', '88': '88001'}
dep['capital_divipola'] = dep.departamento_id.map(lambda d: CAPITAL.get(d, d + '001'))
dep.loc[dep.departamento_id == '27', 'capital_divipola'] = '27001'
dep['capital'] = dep.capital_divipola.map(NAME)
dep['tipo'] = np.where(dep.departamento_id == '11', 'Distrito Capital', 'Departamento')
dep['region'] = dep.departamento_id.map(REGION_OF)
log('GEO', len(mun), 'entidades;', (mun.elige_alcalde == 'Sí').sum(), 'eligen alcalde;', len(dep), 'departamentos + DC')

# =============================================================== 2. NOMBRES DE PARTIDO (CEDAE)
def read_cedae(path, usecols=None):
    return pd.read_csv(path, encoding='utf-8', encoding_errors='replace', low_memory=False, usecols=usecols)

code_name = {}
for y in [2003, 2007, 2011, 2015, 2019]:
    f = os.path.join(RAW, f'{y}_concejo.csv')
    if not os.path.exists(f): continue
    c = read_cedae(f, ['codigo_partido', 'codigo_lista', 'primer_apellido', 'nombres'])
    c = c[(c.codigo_lista == 0) & c.primer_apellido.isna() & c.codigo_partido.notna() & c.nombres.notna()]  # fila de voto por la lista = nombre del partido
    for k, v in c.groupby('codigo_partido').nombres.agg(lambda s: s.value_counts().index[0]).items():
        code_name.setdefault(int(k), v)
log('Códigos de partido con nombre (concejo):', len(code_name))

# Esri 2019 para nombrar organizaciones de los 3 primeros puestos
e19 = pd.DataFrame([f['attributes'] for f in json.load(open(os.path.join(RAW, 'esri_alc_2019.json'), encoding='utf-8'))['features']])
esri19 = {}
for _, r in e19.iterrows():
    for c, p in (('candidatoGanador', 'partidoGanador'), ('segundoPuesto', 'partidoSegundoPuesto'), ('tercerPuesto', 'partidoTercerPuesto')):
        if r[c] and r[p] and norm(r[c]) not in SPECIAL:
            esri19[(r['codigoMunicipalMGNCompleto'], name_key(r[c]))] = r[p]

# =============================================================== 3. RESULTADOS HISTÓRICOS (CEDAE)
def cedae(dign, y):
    f = os.path.join(RAW, f'{y}_{"alcaldia" if dign == "ALC" else "gobernacion"}.csv')
    d = read_cedae(f)
    d['codigo_divipola'] = d.codmpio.astype(int).astype(str).str.zfill(5)
    d['nombre_candidato'] = (d.nombres.fillna('') + ' ' + d.primer_apellido.fillna('') + ' ' + d.segundo_apellido.fillna('')).str.replace(r'\s+', ' ', regex=True).str.strip()
    d['tipo_fila'] = np.select([d.codigo_lista == 999, d.codigo_lista == 998, d.codigo_lista == 997], ['BLANCO', 'NULOS', 'NO_MARCADOS'], 'CANDIDATO')
    d['anio'] = y
    return d

hist, part = [], []
sin_nombre = set()
for dign in ['ALC', 'GOB']:
    for y in [2003, 2007, 2011, 2015, 2019]:
        d = cedae(dign, y)
        fuera = set(d.codigo_divipola) - CODES
        if fuera: log(f'  {dign} {y}: códigos fuera de DIVIPOLA vigente:', sorted(fuera)[:10])
        cand = d[d.tipo_fila == 'CANDIDATO'].copy()
        def orgname(r):
            if dign == 'ALC' and y == 2019:
                n = esri19.get((r.codigo_divipola, name_key(r.nombre_candidato)))
                if n: return n, 'F-ESRI-2019'
            if pd.notna(r.codigo_partido) and int(r.codigo_partido) in code_name:
                return code_name[int(r.codigo_partido)], 'F-CEDAE'
            return None, None
        cand['organizacion'] = None; cand['fuente_nombre_org'] = None
        cp = cand.codigo_partido.fillna(0).astype(int).map(code_name)
        cand.loc[cp.notna(), 'organizacion'] = cp[cp.notna()]; cand.loc[cp.notna(), 'fuente_nombre_org'] = 'F-CEDAE'
        if dign == 'ALC' and y == 2019:
            en = pd.Series([esri19.get((c, name_key(n))) for c, n in zip(cand.codigo_divipola, cand.nombre_candidato)], index=cand.index)
            cand.loc[en.notna(), 'organizacion'] = en[en.notna()]; cand.loc[en.notna(), 'fuente_nombre_org'] = 'F-ESRI-2019' 
        miss = cand.organizacion.isna()
        sin_nombre |= set(cand.loc[miss, 'codigo_partido'].dropna().astype(int))
        cand.loc[miss, 'organizacion'] = 'SIN NOMBRE EN FUENTE (código ' + cand.loc[miss, 'codigo_partido'].fillna(0).astype(int).astype(str) + ')'
        hist.append(cand.assign(dignidad=dign))
        sp = d.pivot_table(index='codigo_divipola', columns='tipo_fila', values='votos', aggfunc='sum').fillna(0)
        cz = d.groupby('codigo_divipola').censoe_total.max() if 'censoe_total' in d else pd.Series(dtype=float)
        sp['censo_electoral'] = cz
        sp['anio'] = y; sp['dignidad'] = dign
        part.append(sp.reset_index())
H = pd.concat(hist, ignore_index=True)
P = pd.concat(part, ignore_index=True)
log('Histórico CEDAE: filas candidato', len(H), '| organizaciones sin nombre:', len(sin_nombre))

# =============================================================== 4. ALCALDÍA 2023 (Esri) + Wikipedia
e23 = pd.DataFrame([f['attributes'] for f in json.load(open(os.path.join(RAW, 'esri_alc_2023.json'), encoding='utf-8'))['features']])
e23['codigo_divipola'] = e23.codigoMunicipalMGNCompleto
wa = pd.read_csv(os.path.join(RAW, 'alc2023_wiki.csv'))
wa['dkey'] = wa.departamento.map(norm); wa['mkey'] = wa.municipio.map(norm)
mun['dkey'] = mun.nombre_departamento.map(norm); mun['mkey'] = mun.nombre_municipio.map(norm)
wiki_by = {}
for _, r in wa.iterrows():
    if pd.isna(r.alcalde): continue
    wiki_by.setdefault(name_key(r.alcalde), []).append(r)
rows23 = []
# Registros Esri cuyo "ganador" es blanco/nulo/no marcado: dato inconsistente en la fuente -> se descartan y se usa Wikipedia si existe
e23['esri_inconsistente'] = e23.candidatoGanador.map(lambda x: bool(x) and norm(x) in SPECIAL).astype(bool)
INCONS = e23[e23.esri_inconsistente].codigo_divipola.tolist()
wk_mun = {}
for _, r in wa.iterrows():
    if pd.isna(r.alcalde): continue
    wk_mun[(r.dkey, r.mkey)] = r
for cod in INCONS:
    dn, mn = norm(DEPN.get(cod[:2], '')), norm(NAME.get(cod, ''))
    r = wk_mun.get((dn, mn))
    if r is None:
        cand = [v for (d, m), v in wk_mun.items() if d[:5] == dn[:5] and difflib.SequenceMatcher(None, m, mn).ratio() > 0.85]
        r = cand[0] if cand else None
    if r is not None and pd.notna(r.votos):
        org = r.organizacion if pd.notna(r.organizacion) else 'SIN NOMBRE EN FUENTE'
        rows23.append(dict(anio=2023, dignidad='ALC', codigo_divipola=cod, nombre_candidato=r.alcalde, organizacion=org, votos=r.votos, posicion=1, ganador=1,
                           validos_estimados=round(r.votos / (r.pct_wiki / 100)) if pd.notna(r.pct_wiki) and r.pct_wiki else None,
                           integrantes_coalicion=r.integrantes if pd.notna(r.integrantes) else '', fuente_nombre_org='F-WIKI-2023', fuente_id='F-WIKI-2023'))
log('Alcaldía 2023: registros Esri inconsistentes', len(INCONS), '; repuestos con Wikipedia', len(rows23))
for _, r in e23[~e23.esri_inconsistente].iterrows():
    cod = r.codigo_divipola
    wk = None
    if pd.notna(r.candidatoGanador):
        cands = wiki_by.get(name_key(r.candidatoGanador), [])
        wk = cands[0] if cands else None
    pct_w = float(wk.pct_wiki) / 100 if wk is not None and pd.notna(wk.pct_wiki) else None
    validos_est = round(r.votosGanador / pct_w) if pct_w and r.votosGanador else None
    for pos, (c, p, v) in enumerate((('candidatoGanador', 'partidoGanador', 'votosGanador'), ('segundoPuesto', 'partidoSegundoPuesto', 'votosSegundo'),
                                     ('tercerPuesto', 'partidoTercerPuesto', 'votosTercero')), 1):
        if pd.isna(r[c]) or not r[c] or norm(r[c]) in SPECIAL: continue
        rows23.append(dict(anio=2023, dignidad='ALC', codigo_divipola=cod, nombre_candidato=r[c], organizacion=r[p], votos=r[v],
                           posicion=(1 if pos == 1 else np.nan), ganador=int(pos == 1), validos_estimados=validos_est,
                           integrantes_coalicion=(wk.integrantes if (pos == 1 and wk is not None and pd.notna(wk.integrantes)) else ''),
                           fuente_nombre_org='F-ESRI-2023', fuente_id='F-ESRI-2023'))
A23 = pd.DataFrame(rows23)
A23['organizacion'] = A23.organizacion.fillna('SIN NOMBRE EN FUENTE')
log('Alcaldía 2023 (Esri top-3):', A23.codigo_divipola.nunique(), 'municipios;', A23.ganador.sum(), 'ganadores;',
    'con % de Wikipedia:', A23[A23.ganador == 1].validos_estimados.notna().sum())

# =============================================================== 5. GOBERNACIÓN 2023 (Wikipedia, por departamento)
wg = pd.read_csv(os.path.join(RAW, 'gob2023_wiki.csv'))
dep['dkey'] = dep.nombre_departamento.map(norm)
dk = {k: i for k, i in zip(dep.dkey, dep.departamento_id)}
def dep_id(name):
    k = norm(name)
    if k in dk: return dk[k]
    m = difflib.get_close_matches(k, dk.keys(), n=1, cutoff=0.6)
    return dk[m[0]] if m else None
wg['departamento_id'] = wg.departamento.map(dep_id)
assert wg.departamento_id.notna().all(), wg[wg.departamento_id.isna()].departamento.unique()

# =============================================================== 6. TABLA DE CANDIDATURAS UNIFICADA (FACT_RESULTADOS)
H['posicion'] = H.groupby(['dignidad', 'anio', 'codigo_divipola']).votos.rank(ascending=False, method='first').astype(int)
H['ganador'] = (H.posicion == 1).astype(int)
H['fuente_id'] = 'F-CEDAE'
cols = ['fuente_id', 'anio', 'dignidad', 'codigo_divipola', 'nombre_candidato', 'organizacion', 'codigo_partido', 'votos', 'posicion', 'ganador', 'fuente_nombre_org']
R = pd.concat([H[cols].assign(nivel='Municipio', estado_validacion='Oficial consolidado (CEDAE)'),
               A23.assign(nivel='Municipio', estado_validacion=np.where(A23.fuente_id == 'F-WIKI-2023', 'Secundaria (Wikipedia) – sólo ganador; Esri inconsistente', 'Secundaria – 3 registros por municipio; orden 2.º/3.º no confiable; corte no declarado'))],
              ignore_index=True)
R = R[R.dignidad.eq('GOB') | R.codigo_divipola.isin(CODES) | True]
R['departamento_id'] = R.codigo_divipola.str[:2]
R['municipio_id'] = 'MUN-' + R.codigo_divipola
R['eleccion_id'] = 'ELEC-' + R.dignidad + '-' + R.anio.astype(str)
R['nombre_municipio'] = R.codigo_divipola.map(NAME).fillna('(código no vigente)')
R['nombre_departamento'] = R.departamento_id.map(DEPN)
R['nombre_candidato'] = R.nombre_candidato.fillna('(sin nombre en fuente)')
R['person_key'] = R.nombre_candidato.map(name_key)
R['candidato_id'] = R.person_key.map(lambda k: hid('CAN', k))
R['org_key'] = R.organizacion.map(org_key)
R['tipo_organizacion'] = R.organizacion.map(org_tipo)
R['fuerza'] = R.organizacion.map(fuerza)
R['organizacion_id'] = np.where(R.org_key.notna(), 'ORG-' + R.org_key.fillna(''), R.apply(lambda r: hid('LOC', norm(r.organizacion), r.codigo_divipola if r.dignidad == 'ALC' else r.departamento_id, r.anio), axis=1))
# votos válidos por municipio (candidatos + blanco) para porcentaje
Pi = P.set_index(['dignidad', 'anio', 'codigo_divipola'])
P['validos'] = P.CANDIDATO + P.BLANCO
P['sufragantes'] = P.CANDIDATO + P.BLANCO + P.NULOS + P.NO_MARCADOS
val = P.set_index(['dignidad', 'anio', 'codigo_divipola']).validos
R = R.merge(val.rename('votos_validos').reset_index(), on=['dignidad', 'anio', 'codigo_divipola'], how='left')
m23 = R.anio.eq(2023)
R.loc[m23, 'votos_validos'] = R.loc[m23, 'codigo_divipola'].map(A23.groupby('codigo_divipola').validos_estimados.first())
R['porcentaje_votos'] = (R.votos / R.votos_validos).round(4)
R['tipo_dato_porcentaje'] = np.where(m23 & R.votos_validos.notna(), 'Derivado: válidos estimados con % del ganador (Wikipedia)', np.where(m23, 'Sin dato', 'Observado'))
R = R.sort_values(['dignidad', 'anio', 'codigo_divipola', 'posicion']).reset_index(drop=True)
R['resultado_id'] = [hid('RES', a, b, c, d, e, f, g) for a, b, c, d, e, f, g in zip(R.dignidad, R.anio, R.codigo_divipola, R.posicion, R.person_key, R.organizacion, R.votos)]

# Gobernación 2023 a nivel departamento (Wikipedia)
G23 = wg[wg.tipo == 'CANDIDATO'].copy()
tot = wg.pivot_table(index='departamento_id', columns='tipo', values='votos', aggfunc='sum')
G23 = G23.assign(anio=2023, dignidad='GOB', nivel='Departamento', codigo_divipola='', nombre_candidato=G23.candidato,
                 organizacion=G23.partido, ganador=G23.electo, fuente_id='F-WIKI-2023',
                 estado_validacion='Secundaria (Wikipedia) – contrastar con escrutinio oficial')
G23['votos_validos'] = G23.departamento_id.map(tot.VALIDOS)
G23['porcentaje_votos'] = (G23.votos / G23.votos_validos).round(4)
G23['tipo_dato_porcentaje'] = 'Observado (fuente secundaria)'
for c in ['municipio_id', 'nombre_municipio']: G23[c] = ''
G23['nombre_departamento'] = G23.departamento_id.map(DEPN)
G23['eleccion_id'] = 'ELEC-GOB-2023'
G23['person_key'] = G23.nombre_candidato.map(name_key); G23['candidato_id'] = G23.person_key.map(lambda k: hid('CAN', k))
G23['org_key'] = G23.organizacion.map(org_key); G23['tipo_organizacion'] = G23.organizacion.map(org_tipo); G23['fuerza'] = G23.organizacion.map(fuerza)
G23['organizacion_id'] = np.where(G23.org_key.notna(), 'ORG-' + G23.org_key.fillna(''), [hid('LOC', norm(o), d, 2023) for o, d in zip(G23.organizacion, G23.departamento_id)])
G23['resultado_id'] = [hid('RES', 'GOB', 2023, d, p, k) for d, p, k in zip(G23.departamento_id, G23.posicion, G23.person_key)]
G23['fuente_nombre_org'] = 'F-WIKI-2023'

RES_COLS = ['resultado_id', 'eleccion_id', 'anio', 'dignidad', 'nivel', 'departamento_id', 'nombre_departamento', 'municipio_id', 'codigo_divipola',
            'nombre_municipio', 'candidato_id', 'nombre_candidato', 'organizacion_id', 'organizacion', 'tipo_organizacion', 'fuerza', 'codigo_partido',
            'votos', 'votos_validos', 'porcentaje_votos', 'tipo_dato_porcentaje', 'posicion', 'ganador', 'fuente_id', 'fuente_nombre_org', 'estado_validacion']
RA = R[R.dignidad == 'ALC'][RES_COLS]
RG = pd.concat([R[R.dignidad == 'GOB'], G23], ignore_index=True)[RES_COLS]
save(RA, 'fact_resultados_alcaldia'); save(RG, 'fact_resultados_gobernacion')

# Participación
P['departamento_id'] = P.codigo_divipola.str[:2]
P['municipio_id'] = 'MUN-' + P.codigo_divipola
P['nombre_municipio'] = P.codigo_divipola.map(NAME).fillna('(código no vigente)')
P['participacion'] = (P.sufragantes / P.censo_electoral).where(P.censo_electoral > 0).round(4)
P['abstencion'] = (1 - P.participacion).round(4)
P = P.rename(columns={'CANDIDATO': 'votos_candidatos', 'BLANCO': 'votos_blancos', 'NULOS': 'votos_nulos', 'NO_MARCADOS': 'votos_no_marcados'})
P['fuente_id'] = 'F-CEDAE'
P['nota'] = np.where(P.censo_electoral.isna(), 'Censo electoral no disponible en la fuente para esta dignidad/año', '')
PCOLS = ['dignidad', 'anio', 'departamento_id', 'municipio_id', 'codigo_divipola', 'nombre_municipio', 'censo_electoral', 'sufragantes', 'validos',
         'votos_candidatos', 'votos_blancos', 'votos_nulos', 'votos_no_marcados', 'participacion', 'abstencion', 'fuente_id', 'nota']
# Gobernación 2023 por departamento
pg = tot.reset_index().rename(columns={'VALIDOS': 'validos', 'BLANCO': 'votos_blancos', 'NULOS': 'votos_nulos', 'NO_MARCADOS': 'votos_no_marcados', 'ESCRUTADOS': 'sufragantes'})
pg = pg.assign(dignidad='GOB', anio=2023, municipio_id='', codigo_divipola='', nombre_municipio='(total departamental)', censo_electoral=np.nan,
               votos_candidatos=pg.validos - pg.votos_blancos.fillna(0), participacion=np.nan, abstencion=np.nan, fuente_id='F-WIKI-2023',
               nota='Nivel departamento; censo electoral 2023 pendiente de fuente oficial')
P = pd.concat([P[PCOLS], pg[PCOLS]], ignore_index=True)
save(P, 'fact_participacion')

# =============================================================== 7. ELECCIONES, ORGANIZACIONES, CANDIDATOS
EL = []
for d, nombre in (('ALC', 'Alcaldía municipal/distrital'), ('GOB', 'Gobernación departamental')):
    for y in ANIOS:
        EL.append(dict(eleccion_id=f'ELEC-{d}-{y}', anio=y, fecha=FECHAS[y], dignidad=d, dignidad_nombre=nombre, tipo='Ordinaria territorial',
                       periodo=f'{y + 1}-{y + 4}', fuente_resultados={2023: 'F-ESRI-2023' if d == 'ALC' else 'F-WIKI-2023'}.get(y, 'F-CEDAE'),
                       cobertura=('Todos los candidatos por municipio' if y < 2023 else ('3 primeros por municipio' if d == 'ALC' else 'Todos los candidatos, total departamental')),
                       estado='Histórica'))
    EL.append(dict(eleccion_id=f'ELEC-{d}-2027', anio=2027, fecha=ELECCION_OBJETIVO, dignidad=d, dignidad_nombre=nombre, tipo='Ordinaria territorial',
                   periodo='2028-2031', fuente_resultados='', cobertura='', estado='Convocada (Res. 12338 de 2026)'))
for d, nombre in (('CON', 'Concejo municipal/distrital'), ('ASA', 'Asamblea departamental'), ('JAL', 'Juntas Administradoras Locales')):
    EL.append(dict(eleccion_id=f'ELEC-{d}-2027', anio=2027, fecha=ELECCION_OBJETIVO, dignidad=d, dignidad_nombre=nombre, tipo='Ordinaria territorial',
                   periodo='2028-2031', fuente_resultados='', cobertura='Sin resultados cargados (fuera del alcance de esta versión)', estado='Convocada (Res. 12338 de 2026)'))
EL = pd.DataFrame(EL); save(EL, 'dim_elecciones')

ALLR = pd.concat([RA, RG], ignore_index=True)
orgs = ALLR.groupby('organizacion_id').agg(nombre=('organizacion', lambda s: s.value_counts().index[0]), tipo=('tipo_organizacion', 'first'),
                                            fuerza=('fuerza', 'first'), primer_anio=('anio', 'min'), ultimo_anio=('anio', 'max'),
                                            candidaturas=('resultado_id', 'count'), victorias=('ganador', 'sum')).reset_index()
nat = orgs[orgs.organizacion_id.str.startswith('ORG-')].copy()
nat['org_key'] = nat.organizacion_id.str[4:]
# completar catálogo con partidos vigentes 2026 sin resultados
for k in VIGENTES_2026 - set(nat.org_key):
    nat = pd.concat([nat, pd.DataFrame([dict(organizacion_id='ORG-' + k, org_key=k, tipo='Partido / movimiento con personería (histórica o vigente)', fuerza=k, candidaturas=0, victorias=0)])])
nat['nombre_oficial'] = nat.org_key.map(ORG_NAME)
nat['sigla'] = nat.org_key.map(ORG_SIGLA)
nat['estado_legal_2026'] = np.where(nat.org_key.isin(VIGENTES_2026), 'Con personería jurídica (El Espectador, 05-08-2026)', 'No figura en el listado 2026 consultado – verificar CNE')
nat['ambito'] = 'Nacional'
nat['orientacion_politica_documentada'] = 'Pendiente de fuente'
vic = ALLR[(ALLR.ganador == 1) & ALLR.organizacion_id.str.startswith('ORG-')].pivot_table(index='organizacion_id', columns=['dignidad', 'anio'], values='resultado_id', aggfunc='count').fillna(0).astype(int)
vic.columns = [f'victorias_{d.lower()}_{a}' for d, a in vic.columns]
nat = nat.merge(vic, left_on='organizacion_id', right_index=True, how='left')
nat['fuente_id'] = 'F-CEDAE;F-ESRI-2023;F-WIKI-2023;F-NEWS-PART'
save(nat.sort_values('victorias', ascending=False), 'dim_partidos')
loc = orgs[~orgs.organizacion_id.str.startswith('ORG-')].copy()
loc['ambito'] = np.where(loc.organizacion_id.isin(RG.organizacion_id), 'Departamental', 'Municipal')
terr = ALLR.groupby('organizacion_id').agg(departamento_id=('departamento_id', 'first'), codigo_divipola=('codigo_divipola', 'first'), dignidad=('dignidad', 'first')).reset_index()
loc = loc.merge(terr, on='organizacion_id', how='left')
loc['nombre_territorio'] = np.where(loc.codigo_divipola.fillna('') != '', loc.codigo_divipola.map(NAME), loc.departamento_id.map(DEPN))
loc['fuente_id'] = np.where(loc.primer_anio == 2023, np.where(loc.dignidad == 'GOB', 'F-WIKI-2023', 'F-ESRI-2023'), 'F-CEDAE')
save(loc.sort_values(['primer_anio', 'departamento_id']), 'dim_movimientos_coaliciones_gsc')

# Coaliciones 2023 con integrantes (alcaldía: Wikipedia por ganador; gobernación: texto entre paréntesis)
co = []
for _, r in A23[A23.organizacion.map(org_tipo) == 'Coalición'].iterrows():
    co.append(dict(coalicion_id=hid('LOC', norm(r.organizacion), r.codigo_divipola, 2023), nombre_coalicion=r.organizacion, dignidad='ALC', anio=2023,
                   departamento_id=r.codigo_divipola[:2], codigo_divipola=r.codigo_divipola, nombre_territorio=NAME.get(r.codigo_divipola),
                   candidato_respaldado=r.nombre_candidato, resultado=('Ganó' if r.ganador else f'Puesto {r.posicion}'),
                   organizaciones_integrantes=r.integrantes_coalicion or 'Pendiente de fuente', fuente_id='F-ESRI-2023' + (';F-WIKI-2023' if r.integrantes_coalicion else ''),
                   estado='Histórica (2023)'))
for _, r in G23[G23.tipo_organizacion == 'Coalición'].iterrows():
    mm = re.search(r'\((.*)\)\s*$', r.organizacion)
    co.append(dict(coalicion_id=r.organizacion_id, nombre_coalicion=re.sub(r'\s*\(.*$', '', r.organizacion), dignidad='GOB', anio=2023, departamento_id=r.departamento_id,
                   codigo_divipola='', nombre_territorio=DEPN[r.departamento_id], candidato_respaldado=r.nombre_candidato,
                   resultado='Ganó' if r.ganador else f'Puesto {int(r.posicion)}', organizaciones_integrantes=(mm.group(1).strip() if mm else 'Pendiente de fuente'),
                   fuente_id='F-WIKI-2023', estado='Histórica (2023)'))
CO = pd.DataFrame(co)
CO['integrantes_canonicos'] = CO.organizaciones_integrantes.map(lambda s: ', '.join(sorted({k for k in (org_key(x) for x in re.split(r'[;,]| y ', s)) if k})))
save(CO, 'fact_coaliciones')

cand = ALLR.groupby('candidato_id').agg(nombre_completo=('nombre_candidato', lambda s: max(s.dropna().astype(str), key=len, default='(sin nombre)')), participaciones=('resultado_id', 'count'),
                                        victorias=('ganador', 'sum'), anios=('anio', lambda s: ', '.join(str(x) for x in sorted(set(s)))),
                                        dignidades=('dignidad', lambda s: ', '.join(sorted(set(s))))).reset_index()
save(cand, 'dim_candidatos_historicos')

# =============================================================== 8. ALCALDES Y GOBERNADORES ACTUALES (electos 2023)
prev = RA[RA.anio < 2023].groupby('candidato_id').apply(lambda d: '; '.join(f'{r.anio} {r.nombre_municipio} ({"ganó" if r.ganador else f"puesto {int(r.posicion)}"}, {r.organizacion[:40]})' for r in d.itertuples()), include_groups=False)
w = RA[(RA.anio == 2023) & (RA.ganador == 1)].copy()
w2 = RA[(RA.anio == 2023) & (RA.posicion == 2)].set_index('codigo_divipola')
w['segundo'] = w.codigo_divipola.map(w2.nombre_candidato); w['votos_segundo'] = w.codigo_divipola.map(w2.votos)

ALC = pd.DataFrame(dict(
    alcalde_id=[hid('ALC', c, 2024) for c in w.codigo_divipola], nombre_completo=w.nombre_candidato, candidato_id=w.candidato_id, municipio_id=w.municipio_id,
    region=w.departamento_id.map(REGION_OF),
    codigo_divipola=w.codigo_divipola, nombre_municipio=w.nombre_municipio, departamento_id=w.departamento_id, nombre_departamento=w.nombre_departamento,
    partido_o_movimiento=w.organizacion, tipo_organizacion=w.tipo_organizacion, fuerza=w.fuerza,
    integrantes_coalicion=w.codigo_divipola.map(A23[A23.ganador == 1].set_index('codigo_divipola').integrantes_coalicion).fillna(''),
    fecha_inicio_gestion='2024-01-01', fecha_fin_gestion='2027-12-31', eleccion_origen='ELEC-ALC-2023', votos_obtenidos=w.votos,
    porcentaje_votos=w.porcentaje_votos, posicion_electoral=1,
    condicion_actual='Alcalde electo 2023 – verificar continuidad (atípicas, renuncias, nulidades)', reeleccion_inmediata='No permitida (art. 314 C.P.)',
    antecedentes_electorales=w.candidato_id.map(prev).fillna('Sin candidaturas a alcaldía 2003-2019 con el mismo nombre'),
    fuente_id='F-ESRI-2023', fecha_verificacion=FECHA_CORTE))
save(ALC, 'alcaldes_2024_2027')

gw = G23[G23.ganador == 1].copy(); g2 = G23[G23.posicion == 2].set_index('departamento_id')
GOBS = pd.DataFrame(dict(
    gobernador_id=[hid('GOB', d, 2024) for d in gw.departamento_id], nombre_completo=gw.nombre_candidato, candidato_id=gw.candidato_id,
    departamento_id=gw.departamento_id, nombre_departamento=gw.nombre_departamento, partido_o_coalicion=gw.organizacion, tipo_organizacion=gw.tipo_organizacion,
    fuerza=gw.fuerza, fecha_inicio_gestion='2024-01-01', fecha_fin_gestion='2027-12-31', eleccion_origen='ELEC-GOB-2023',
    votos_obtenidos=gw.votos, porcentaje_votos=gw.porcentaje_votos, segundo_lugar=gw.departamento_id.map(g2.nombre_candidato),
    margen_pts=((gw.votos - gw.departamento_id.map(g2.votos)) / gw.votos_validos * 100).round(2).values,
    condicion_actual=np.where(gw.departamento_id == '88', 'Elección anulada por doble militancia (El Nuevo Siglo) – verificar titular actual',
                              'Gobernador electo 2023 – verificar continuidad'),
    fuente_id='F-WIKI-2023', fecha_verificacion=FECHA_CORTE))
save(GOBS, 'gobernadores_2024_2027')

# =============================================================== 9. INDICADORES POR MUNICIPIO Y ELECCIÓN (alcaldía)
def indic(d):
    anio = d.name[0]
    v = d.sort_values('votos', ascending=False)
    vv = d.votos_validos.iloc[0]
    s = v.votos / v.votos.sum()
    out = dict(n_candidatos=len(v) if anio < 2023 else np.nan,
               ganador=v.nombre_candidato.iloc[0], organizacion_ganadora=v.organizacion.iloc[0], fuerza_ganadora=v.fuerza.iloc[0],
               tipo_org_ganadora=v.tipo_organizacion.iloc[0], votos_ganador=v.votos.iloc[0],
               pct_ganador=round(v.votos.iloc[0] / vv, 4) if vv else np.nan,
               segundo=v.nombre_candidato.iloc[1] if len(v) > 1 else '', organizacion_segundo=v.organizacion.iloc[1] if len(v) > 1 else '',
               pct_segundo=round(v.votos.iloc[1] / vv, 4) if (len(v) > 1 and vv) else np.nan)
    out['margen_pts'] = round((out['pct_ganador'] - (out['pct_segundo'] if len(v) > 1 else 0)) * 100, 2) if vv else np.nan
    if anio == 2023:  # el 2.º/3.º de la fuente 2023 no es confiable: no se calcula margen
        out.update(segundo='', organizacion_segundo='', pct_segundo=np.nan, margen_pts=np.nan)
    out.update(nep=np.nan, fragmentacion_rae=np.nan, concentracion_top2=np.nan)
    if anio < 2023:
        out['nep'] = round(1 / (s ** 2).sum(), 3)
        out['fragmentacion_rae'] = round(1 - (s ** 2).sum(), 4)
        out['concentracion_top2'] = round(s.iloc[:2].sum(), 4)
    return pd.Series(out)
IND = RA.groupby(['anio', 'codigo_divipola']).apply(indic, include_groups=False).reset_index()
pa = P[P.dignidad == 'ALC'].set_index(['anio', 'codigo_divipola'])
IND = IND.merge(pa.participacion.reset_index(), on=['anio', 'codigo_divipola'], how='left')
IND = IND.sort_values(['codigo_divipola', 'anio'])
IND['fuerza_anterior'] = IND.groupby('codigo_divipola').fuerza_ganadora.shift()
IND['participacion_anterior'] = IND.groupby('codigo_divipola').participacion.shift()
PARTY_LIKE = lambda f: isinstance(f, str) and f not in ('COALICION', 'GSC', 'OTRA_LOCAL')
IND['alternancia'] = [('Sin ciclo previo' if not isinstance(p, str) else ('Indeterminada (coalición/GSC/local)' if not (PARTY_LIKE(f) and PARTY_LIKE(p)) else ('Continuidad' if f == p else 'Alternancia')))
                      for f, p in zip(IND.fuerza_ganadora, IND.fuerza_anterior)]
IND['var_participacion_pts'] = ((IND.participacion - IND.participacion_anterior) * 100).round(2)
# Volatilidad de Pedersen por fuerza (todas las coaliciones/GSC agrupadas en su tipo) entre elecciones consecutivas con datos completos
sh = RA[RA.anio < 2023].groupby(['codigo_divipola', 'anio', 'fuerza']).votos.sum()
sh = (sh / sh.groupby(level=[0, 1]).transform('sum')).unstack('fuerza').fillna(0)
ped = {}
for c, dd in sh.groupby(level=0):
    dd = dd.droplevel(0).sort_index()
    for a0, a1 in zip(dd.index[:-1], dd.index[1:]):
        ped[(a1, c)] = round(0.5 * (dd.loc[a1] - dd.loc[a0]).abs().sum() * 100, 2)
IND = IND.merge(pd.Series(ped, name='volatilidad_pedersen').rename_axis(['anio', 'codigo_divipola']).reset_index(), on=['anio', 'codigo_divipola'], how='left')
IND['departamento_id'] = IND.codigo_divipola.str[:2]
IND['municipio_id'] = 'MUN-' + IND.codigo_divipola
IND['nombre_municipio'] = IND.codigo_divipola.map(NAME).fillna('(código no vigente)')
IND['fuente_id'] = 'F-ATLAS;' + np.where(IND.anio == 2023, 'F-ESRI-2023', 'F-CEDAE')
save(IND, 'fact_indicadores_alcaldia')

# =============================================================== 10. BASTIONES Y SWING (alcaldía, por municipio)
piv = IND.pivot(index='codigo_divipola', columns='anio', values='fuerza_ganadora')
mg_ = IND.pivot(index='codigo_divipola', columns='anio', values='margen_pts')
nep_ = IND.pivot(index='codigo_divipola', columns='anio', values='nep')
B = pd.DataFrame(index=piv.index)
for y in ANIOS:
    B[f'fuerza_{y}'] = piv.get(y); B[f'margen_{y}'] = mg_.get(y)
B['nep_2019'] = nep_.get(2019)
def bast(r):
    seq = [r.get(f'fuerza_{y}') for y in [2007, 2011, 2015, 2019, 2023]]
    last = seq[-1]
    party = [s for s in seq if PARTY_LIKE(s)]
    cnt = pd.Series(party).value_counts()
    top, n = (cnt.index[0], cnt.iloc[0]) if len(cnt) else (None, 0)
    if PARTY_LIKE(last) and all(PARTY_LIKE(x) and x == last for x in seq[-3:]): return last, 'Bastión consolidado', int((pd.Series(seq) == last).sum())
    if PARTY_LIKE(last) and (pd.Series(seq) == last).sum() >= 3: return last, 'Bastión político', int((pd.Series(seq) == last).sum())
    if top and n >= 3 and last != top: return top, 'Bastión en declive', int(n)
    return (top or ''), 'Sin bastión partidario', int(n)
bb = B.apply(bast, axis=1, result_type='expand'); B['fuerza_dominante'], B['clasificacion_bastion'], B['victorias_fuerza_2007_2023'] = bb[0], bb[1], bb[2]
B['cambio_2019_2023'] = [('Indeterminado (coalición/GSC/local)' if not (PARTY_LIKE(a) and PARTY_LIKE(b)) else ('Sí' if a != b else 'No')) for a, b in zip(B.fuerza_2019, B.fuerza_2023)]
B = B.reset_index()
B['municipio_id'] = 'MUN-' + B.codigo_divipola; B['nombre_municipio'] = B.codigo_divipola.map(NAME).fillna('(código no vigente)'); B['departamento_id'] = B.codigo_divipola.str[:2]
B['nombre_departamento'] = B.departamento_id.map(DEPN)
save(B, 'fact_bastiones_alcaldia')

# =============================================================== 11. MODELO BASE: retención de la fuerza ganadora
# Unidad: municipio×ciclo con ganador de partido canónico. y = la misma fuerza gana el ciclo siguiente.
# Variable: racha (la fuerza también había ganado el ciclo anterior). No se usa el margen 2023: el orden 2.º/3.º de la fuente Esri no es confiable.
T = IND[['codigo_divipola', 'anio', 'fuerza_ganadora']].copy()
T['previa'] = T.groupby('codigo_divipola').fuerza_ganadora.shift(1)
T['siguiente'] = T.groupby('codigo_divipola').fuerza_ganadora.shift(-1)
T['anio_sig'] = T.groupby('codigo_divipola').anio.shift(-1)
T = T[T.fuerza_ganadora.map(PARTY_LIKE) & T.siguiente.notna()]
T['y'] = (T.fuerza_ganadora == T.siguiente).astype(int)
racha = lambda f, p: 'Repite (también ganó el ciclo previo)' if f == p else 'Primera victoria (o ciclo previo distinto)'
T['tramo'] = [racha(f, p) for f, p in zip(T.fuerza_ganadora, T.previa)]
train, test = T[T.anio_sig < 2023], T[T.anio_sig == 2023]
k = 2  # suavizado bayesiano hacia la tasa global
g0 = train.y.mean()
rate = train.groupby('tramo').y.agg(['sum', 'count'])
rate['p'] = (rate['sum'] + k * g0) / (rate['count'] + k)
test = test.assign(p=test.tramo.map(rate.p).astype(float))
brier = ((test.p - test.y) ** 2).mean(); brier0 = ((g0 - test.y) ** 2).mean()
log(f'Modelo base: tasa global retención (train →2019) {g0:.3f}; backtest 2019→2023 n={len(test)} Brier {brier:.4f} vs constante {brier0:.4f}')
full_g = T.y.mean(); rf = T.groupby('tramo').y.agg(['sum', 'count']); rf['p'] = (rf['sum'] + k * full_g) / (rf['count'] + k)
CAL = rf.reset_index().rename(columns={'tramo': 'condicion', 'sum': 'retenciones', 'count': 'casos', 'p': 'prob_retencion'})
CAL['tasa_global'] = round(full_g, 4)
CAL['backtest_2019_2023'] = [f'n={len(test)}; Brier={brier:.4f}; Brier constante={brier0:.4f}'] * len(CAL)
save(CAL, 'modelo_base_calibracion')
cur = IND[IND.anio == 2023].copy()
cur['previa'] = cur.codigo_divipola.map(IND[IND.anio == 2019].set_index('codigo_divipola').fuerza_ganadora)
cur['tramo'] = [racha(f, p) for f, p in zip(cur.fuerza_ganadora, cur.previa)]
FC = pd.DataFrame(dict(
    forecast_id=[hid('FC', c, 'v0.1', FECHA_CORTE) for c in cur.codigo_divipola], municipio_id=cur.municipio_id, codigo_divipola=cur.codigo_divipola,
    nombre_municipio=cur.nombre_municipio, departamento_id=cur.departamento_id, fecha_corte=FECHA_CORTE, fecha_generacion=FECHA_CORTE,
    version_modelo='v0.1-base-retencion', dignidad='ALC', fuerza_incumbente=cur.fuerza_ganadora, organizacion_incumbente=cur.organizacion_ganadora,
    fuerza_2019=cur.previa, condicion_modelo=cur.tramo,
    probabilidad_retencion_fuerza=np.where(cur.fuerza_ganadora.map(PARTY_LIKE), cur.tramo.map(rf.p).astype(float).round(3), np.nan),
    escenario='Base: persistencia histórica de la fuerza ganadora', candidato_id='', porcentaje_estimado=np.nan, intervalo_inferior=np.nan, intervalo_superior=np.nan,
    nivel_confianza='Bajo (sólo historia de la fuerza; sin candidatos ni encuestas)',
    variables_utilizadas='fuerza ganadora 2023 y 2019; tasa empírica de retención 2003-2023 según racha',
    fuente_datos='F-CEDAE;F-ESRI-2023;F-ATLAS',
    observaciones=np.where(cur.fuerza_ganadora.map(PARTY_LIKE), 'Estimación estadística, no resultado. No hay reelección inmediata: mide si la misma organización retiene la alcaldía.',
                           'Ganador 2023 es coalición/GSC/organización local: sin continuidad partidaria medible.')))
save(FC, 'fact_pronosticos')

# =============================================================== 12. CATÁLOGO MAESTRO DE MUNICIPIOS (para Excel)
last = IND[IND.anio == 2023].set_index('codigo_divipola')
mun['alcalde_actual'] = mun.codigo_divipola.map(ALC.set_index('codigo_divipola').nombre_completo)
mun['partido_gobernante'] = mun.codigo_divipola.map(ALC.set_index('codigo_divipola').partido_o_movimiento)
mun['tipo_org_gobernante'] = mun.codigo_divipola.map(ALC.set_index('codigo_divipola').tipo_organizacion)
mun['fuerza_gobernante'] = mun.codigo_divipola.map(ALC.set_index('codigo_divipola').fuerza)
mun['votos_ganador_2023'] = mun.codigo_divipola.map(last.votos_ganador)
mun['pct_ganador_2023'] = mun.codigo_divipola.map(last.pct_ganador)
mun['margen_2019_pts'] = mun.codigo_divipola.map(IND[IND.anio == 2019].set_index('codigo_divipola').margen_pts)
mun['fuerza_2019'] = mun.codigo_divipola.map(IND[IND.anio == 2019].set_index('codigo_divipola').fuerza_ganadora)
c19 = P[(P.dignidad == 'ALC') & (P.anio == 2019)].set_index('codigo_divipola')
mun['censo_electoral_2019'] = mun.codigo_divipola.map(c19.censo_electoral)
mun['participacion_alc_2019'] = mun.codigo_divipola.map(c19.participacion)
pros = pd.read_csv(os.path.join(ROOT, 'etl', 'manual', 'prospectos_2027.csv'), dtype={'cod_divipola': str})
pros['cod_divipola'] = pros.cod_divipola.str.zfill(5).where(pros.cod_divipola.str.len() > 2, pros.cod_divipola.str.zfill(2))
alc_pros = pros[pros.dignidad.str.startswith('Alcald')]
mun['actores_2027_identificados'] = mun.codigo_divipola.map(alc_pros.groupby('cod_divipola').nombre.agg(lambda s: '; '.join(dict.fromkeys(s)))).fillna('')
mun['n_aspirantes_registrados_2027'] = mun.codigo_divipola.map(alc_pros.groupby('cod_divipola').size()).fillna(0).astype(int)
mun['n_candidaturas_inscritas_2027'] = 0
mun['estado_informacion'] = np.select(
    [mun.elige_alcalde == 'No', mun.alcalde_actual.isna()],
    ['No elige alcalde', 'Sin resultado 2023 en fuente'], 'Resultados 2003-2019 completos + 2023 top-3')
mun['ultima_actualizacion'] = FECHA_CORTE
save(mun.drop(columns=['dkey', 'mkey', 'latitud', 'longitud', 'tipo_municipio']), 'dim_municipios')
dep['gobernador_actual'] = dep.departamento_id.map(GOBS.set_index('departamento_id').nombre_completo)
dep['partido_o_coalicion_gobernador'] = dep.departamento_id.map(GOBS.set_index('departamento_id').partido_o_coalicion)
dep['margen_gob_2023_pts'] = dep.departamento_id.map(GOBS.set_index('departamento_id').margen_pts)
dep['alcaldias_2023_partido'] = dep.departamento_id.map(ALC[ALC.tipo_organizacion.str.startswith('Partido')].groupby('departamento_id').size()).fillna(0).astype(int)
dep['alcaldias_2023_coalicion'] = dep.departamento_id.map(ALC[ALC.tipo_organizacion == 'Coalición'].groupby('departamento_id').size()).fillna(0).astype(int)
dep['alcaldias_2023_otras'] = dep.n_municipios - dep.alcaldias_2023_partido - dep.alcaldias_2023_coalicion
fr = IND[IND.anio == 2019].groupby('departamento_id').nep.mean().round(2)
dep['nep_medio_alc_2019'] = dep.departamento_id.map(fr)
nbid = pd.read_excel(os.path.join(RAW, 'dane_nbi_2021.xlsx'), sheet_name='Departamento', header=None, skiprows=10, dtype={0: str})[[0, 2]].dropna()
nbid.columns = ['d', 'nbi']; nbid = nbid[nbid.d.str.fullmatch(r'\d{2}')]
dep['nbi_2018'] = dep.departamento_id.map(nbid.set_index('d').nbi)
save(dep.drop(columns=['dkey']), 'dim_departamentos')
save(pros, 'prospectos_2027')
log('Listo.')
open(os.path.join(CLEAN, '_log_etl.txt'), 'w', encoding='utf-8').write('\n'.join(LOG))
