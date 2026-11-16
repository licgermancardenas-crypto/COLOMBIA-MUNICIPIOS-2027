"""ATLAS COLOMBIA 2027 – Paso 4: datos del mapa web interactivo (entregables/MAPA_INTERACTIVO/).

Departamentos -> municipios -> sectores censales DANE 2018 (urbanos y rurales).
Los datos se escriben como archivos .js (window.ATLAS_*) para que el HTML funcione abierto desde el disco (file://),
sin servidor. Los sectores se parten por departamento y se cargan al entrar a cada uno.
Uso: python etl/04_build_webmap.py   (requiere raw/sectores/ de etl/05_download_sectores.py y data/clean/)
"""
import os, json, math
import numpy as np, pandas as pd, geopandas as gpd
from common import *
from maps_common import enrich, COL, SW, BA, IPMC, lab

WEB = os.path.join(OUT, 'MAPA_INTERACTIVO'); DATA = os.path.join(WEB, 'data'); os.makedirs(os.path.join(DATA, 'sectores'), exist_ok=True)
rd = lambda n: pd.read_csv(os.path.join(CLEAN, n + '.csv'), dtype={'codigo_divipola': str, 'departamento_id': str, 'cod_divipola': str}, low_memory=False)
mun, dep, B, GOB, IND, RA, FC, PROS, ALC = (rd(n) for n in ('dim_municipios', 'dim_departamentos', 'fact_bastiones_alcaldia', 'gobernadores_2024_2027',
                                                             'fact_indicadores_alcaldia', 'fact_resultados_alcaldia', 'fact_pronosticos', 'prospectos_2027', 'alcaldes_2024_2027'))
M = enrich(mun, B)
def clean(v):
    if v is None: return None
    if isinstance(v, (float, np.floating)): return None if math.isnan(v) else round(float(v), 4)
    if isinstance(v, (np.integer,)): return int(v)
    return v
def js(name, var, obj):
    with open(os.path.join(DATA, name), 'w', encoding='utf-8') as fh:
        fh.write(f'window.{var}=' + json.dumps(obj, ensure_ascii=False, separators=(',', ':')) + ';')
    print(f'  {name}: {os.path.getsize(os.path.join(DATA, name)) / 1e6:.2f} MB')

# ---------------------------------------------------------------- municipios (geometría + ficha)
g = gpd.read_file(os.path.join(OUT, 'MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson'))[['mpio', 'geometry']]
g['geometry'] = g.geometry.make_valid().buffer(0).simplify(0.0015, preserve_topology=True)
g = g.merge(M, left_on='mpio', right_on='codigo_divipola')
pr_mun = PROS[PROS.cod_divipola.str.len() == 5].groupby('cod_divipola').apply(lambda d: d[['nombre', 'dignidad', 'partido_mencionado', 'estado']].to_dict('records'), include_groups=False).to_dict()
pr_dep = PROS[PROS.cod_divipola.str.len() == 2].groupby('cod_divipola').apply(lambda d: d[['nombre', 'dignidad', 'partido_mencionado', 'estado']].to_dict('records'), include_groups=False).to_dict()
hist = IND.groupby('codigo_divipola').apply(lambda d: [[int(r.anio), r.ganador, r.organizacion_ganadora, r.fuerza_ganadora, clean(r.pct_ganador), clean(r.margen_pts), clean(r.participacion)]
                                                      for r in d.sort_values('anio').itertuples()], include_groups=False).to_dict()
r19 = RA[RA.anio == 2019].sort_values('posicion').groupby('codigo_divipola').apply(lambda d: [[r.nombre_candidato, r.organizacion, int(r.votos), clean(r.porcentaje_votos)] for r in d.head(6).itertuples()], include_groups=False).to_dict()
fc = FC.set_index('codigo_divipola').probabilidad_retencion_fuerza.to_dict()
alc = ALC.set_index('codigo_divipola')
feats = []
for r in g.itertuples():
    c = r.codigo_divipola
    p = dict(c=c, n=r.nombre_municipio, d=r.nombre_departamento, di=r.departamento_id, rg=REGION_OF.get(r.departamento_id), t=r.tipo_entidad_territorial,
             al=clean(r.alcalde_actual), org=clean(r.partido_gobernante), fz=clean(r.fuerza_gobernante), tor=clean(r.tipo_org_gobernante),
             coa=clean(alc.integrantes_coalicion.get(c)) if c in alc.index else None, v23=clean(r.votos_ganador_2023), p23=clean(r.pct_ganador_2023),
             f19=clean(r.fuerza_2019), m19=clean(r.margen_2019_pts), sw=r.swing, ba=clean(r.clasificacion_bastion), fd=clean(r.fuerza_dominante),
             pr=clean(fc.get(c)), pob=clean(r.pob_tot_2026), rur=clean(r.pct_rural_2026), ipm=clean(r.ipm_2018), nbi=clean(r.nbi_2018), mdm=clean(r.mdm_grupo_2020),
             ce=clean(r.censo_electoral_2019), pa=clean(r.participacion_alc_2019), iq=clean(r.ipm_q),
             h=hist.get(c, []), c19=r19.get(c, []), act=pr_mun.get(c, []))
    feats.append(dict(type='Feature', properties=p, geometry=json.loads(gpd.GeoSeries([r.geometry]).to_json())['features'][0]['geometry']))
for f_ in feats:   # redondear coordenadas
    def rr(x): return [rr(y) for y in x] if isinstance(x[0], list) else [round(x[0], 4), round(x[1], 4)]
    f_['geometry']['coordinates'] = rr(f_['geometry']['coordinates'])
js('municipios.js', 'ATLAS_MUN', dict(type='FeatureCollection', features=feats))

# ---------------------------------------------------------------- departamentos
dg = gpd.GeoDataFrame(g[['departamento_id', 'geometry']], crs=4326).dissolve(by='departamento_id', as_index=False)
dg['geometry'] = dg.geometry.buffer(0.002).buffer(-0.002).simplify(0.01, preserve_topology=True)
dg = dg.merge(dep, on='departamento_id').merge(GOB[['departamento_id', 'nombre_completo', 'partido_o_coalicion', 'fuerza', 'porcentaje_votos', 'margen_pts', 'condicion_actual']], on='departamento_id', how='left')
cnt = M.groupby(['departamento_id', 'fuerza_gobernante']).size().unstack(fill_value=0)
dfe = []
for r in dg.itertuples():
    cc = cnt.loc[r.departamento_id].sort_values(ascending=False) if r.departamento_id in cnt.index else pd.Series(dtype=int)
    p = dict(c=r.departamento_id, n=r.nombre_departamento, rg=REGION_OF.get(r.departamento_id), cap=r.capital, nm=int(r.n_municipios), pob=clean(r.pob_tot_2026), gob=clean(r.nombre_completo),
             org=clean(r.partido_o_coalicion), fz=clean(r.fuerza), pg=clean(r.porcentaje_votos), mg=clean(r.margen_pts), cond=clean(r.condicion_actual),
             nbi=clean(r.nbi_2018), alc=[[k, int(v)] for k, v in cc.items() if v > 0], act=pr_dep.get(r.departamento_id, []))
    geom = json.loads(gpd.GeoSeries([r.geometry]).to_json())['features'][0]['geometry']
    def rr(x): return [rr(y) for y in x] if isinstance(x[0], list) else [round(x[0], 4), round(x[1], 4)]
    geom['coordinates'] = rr(geom['coordinates'])
    dfe.append(dict(type='Feature', properties=p, geometry=geom))
js('departamentos.js', 'ATLAS_DEP', dict(type='FeatureCollection', features=dfe))

# ---------------------------------------------------------------- leyendas / temas
THEMES = {
    'fz': dict(t='Fuerza que gobierna (alcaldía 2023)', colors={k: v for k, v in COL.items()}, labels={k: lab(k) for k in COL}),
    'sw': dict(t='Competitividad / swing', colors=SW, labels={k: k for k in SW}),
    'ba': dict(t='Bastiones 2007-2023', colors=BA, labels={k: k for k in BA}),
    'iq': dict(t='Pobreza multidimensional (IPM 2018, quintiles)', colors=IPMC, labels={k: k for k in IPMC}),
}
js('temas.js', 'ATLAS_TEMAS', dict(themes=THEMES, corte=FECHA_CORTE, regiones=REGIONES))

# ---------------------------------------------------------------- sectores censales 2018 (por departamento)
SP = os.path.join(RAW, 'sectores')
s = gpd.read_file(os.path.join(SP, 'sectores_personas.geojson'))
vv = pd.DataFrame(json.load(open(os.path.join(SP, 'sectores_vivienda.json'), encoding='utf-8'))).drop_duplicates('ID_UNIFICADO').set_index('ID_UNIFICADO')
s = s.join(vv, on='ID_UNIFICADO', rsuffix='_v')
print('sectores', len(s), 'con vivienda', s.TOTAL_VIVIENDAS.notna().sum())
s['geometry'] = s.geometry.make_valid()
s['km2'] = s.to_crs(9377).area / 1e6
N = lambda c: pd.to_numeric(s[c], errors='coerce').fillna(0)
pop = N('SEXO_TOTAL'); pd0 = pop.replace(0, np.nan)
viv = N('TOTAL_VIVIENDAS').replace(0, np.nan)
est = s[[f'ESTRATO_{i}' for i in range(1, 7)]].apply(pd.to_numeric, errors='coerce').fillna(0)
s['v_pob'] = pop
s['v_den'] = (pop / s.km2.replace(0, np.nan)).round(1)
s['v_may'] = ((N('EDAD_60_64') + N('EDAD_65_69') + N('TOTAL_MAYORES_70')) / pd0).round(4)
s['v_jov'] = ((N('EDAD_15_19') + N('EDAD_20_24') + N('EDAD_25_29')) / pd0).round(4)
s['v_sup'] = ((N('NIVEL_EDUC_UNIVERSITARIA') + N('NIVEL_EDUC_ESP_MAES_DOC') + N('NIVEL_EDUC_TECNICA_TECNOLOGO')) / pd0).round(4)
s['v_ind'] = (N('GRUPO_ETNICO_INDIGENA') / pd0).round(4)
s['v_afr'] = ((N('GRUPO_ETNICO_NEGRO') + N('GRUPO_ETNICO_RAIZAL') + N('GRUPO_ETNICO_PALANQUERO')) / pd0).round(4)
s['v_mig'] = (N('LUG_NAC_OTRO_PAIS') / pd0).round(4)
s['v_anf'] = (N('ALFABETA_NO') / pd0).round(4)
s['v_est'] = np.where(est.sum(axis=1) > 0, est.values.argmax(axis=1) + 1, np.nan)
s['v_acu'] = (N('ACUEDUCTO_NO') / viv).round(4)
s['v_int'] = (N('INTERNET_SI') / viv).round(4)
s['v_viv'] = viv
s['mpio'] = s.ID_UNIFICADO.astype(str).str[:5]
s['dpto'] = s.ID_UNIFICADO.astype(str).str[:2]
s['tipo'] = np.where(s.CLASE.astype(str).str.lower().str.startswith('urb'), 'Urbano', 'Rural')
s['geometry'] = s.geometry.simplify(0.0002, preserve_topology=True)
print('códigos de municipio desde ID_UNIFICADO (muestra):', s[['ID_UNIFICADO', 'mpio', 'DEPTO', 'MPIO', 'CLASE', 'tipo']].head(3).to_dict('records'))
VARS = ['v_pob', 'v_den', 'v_may', 'v_jov', 'v_sup', 'v_ind', 'v_afr', 'v_mig', 'v_anf', 'v_est', 'v_acu', 'v_int', 'v_viv']
idx = {}
for d_, part in s.groupby('dpto'):
    fs = []
    for r in part.itertuples():
        if r.geometry is None or r.geometry.is_empty: continue
        geom = json.loads(gpd.GeoSeries([r.geometry]).to_json())['features'][0]['geometry']
        if geom['type'] not in ('Polygon', 'MultiPolygon'):
            continue
        def rr(x): return [rr(y) for y in x] if isinstance(x[0], list) else [round(x[0], 5), round(x[1], 5)]
        geom['coordinates'] = rr(geom['coordinates'])
        p = dict(id=str(r.ID_UNIFICADO), m=r.mpio, ti=r.tipo, km2=round(r.km2, 3), **{k: clean(getattr(r, k)) for k in VARS})
        fs.append(dict(type='Feature', properties=p, geometry=geom))
    fn = f'sectores/s_{d_}.js'
    with open(os.path.join(DATA, fn), 'w', encoding='utf-8') as fh:
        fh.write(f'window.ATLAS_SEC=window.ATLAS_SEC||{{}};window.ATLAS_SEC["{d_}"]=' + json.dumps(dict(type='FeatureCollection', features=fs), ensure_ascii=False, separators=(',', ':')) + ';')
    idx[d_] = len(fs)
print('sectores por departamento:', idx, 'tamaño total', round(sum(os.path.getsize(os.path.join(DATA, 'sectores', f)) for f in os.listdir(os.path.join(DATA, 'sectores'))) / 1e6, 1), 'MB')

# ---------------------------------------------------------------- accesos directos para Excel (sin depender del #fragmento)
IR = os.path.join(WEB, 'ir'); os.makedirs(IR, exist_ok=True)
def redirect(name, frag, title):
    with open(os.path.join(IR, name), 'w', encoding='utf-8') as fh:
        fh.write(f'<!doctype html><meta charset="utf-8"><title>{title}</title>'
                 f'<script>location.replace("../ATLAS_MAPA_COLOMBIA_2027.html#{frag}")</script>'
                 f'<a href="../ATLAS_MAPA_COLOMBIA_2027.html#{frag}">Abrir {title} en el mapa</a>')
for c, n in zip(mun.codigo_divipola, mun.nombre_municipio): redirect(f'{c}.html', f'm={c}', n)
for d, n in zip(dep.departamento_id, dep.nombre_departamento): redirect(f'd{d}.html', f'd={d}', n)
print('accesos directos:', len(os.listdir(IR)))

# ---------------------------------------------------------------- puestos de votación (escrutinio legislativas 8-mar-2026)
PU = os.path.join(RAW, 'puestos')
LISTAS26 = [  # (patrón en el nombre de la lista ganadora, etiqueta, color)
    ('PACTO HISTORICO', 'Pacto Histórico', '#8E44AD'), ('CENTRO DEMOCRATICO', 'Centro Democrático', '#1B4F72'),
    ('PARTIDO DE LA U', 'La U', '#F39C12'), ('UNION POR LA GENTE', 'La U', '#F39C12'), ('LIBERAL', 'Liberal', '#C0392B'),
    ('CONSERVADOR', 'Conservador', '#2471A3'), ('ALIANZA POR COLOMBIA', 'Alianza por Colombia', '#27AE60'),
    ('CAMBIO RADICAL', 'Cambio Radical – Alma', '#5DADE2'), ('FRENTE AMPLIO', 'Frente Amplio Unitario', '#D81B60'),
    ('MINGA', 'Unidad en Minga', '#6D4C41'), ('ALTERNATIVO INDIGENA', 'MAIS', '#A0522D'), ('AHORA COLOMBIA', 'Ahora Colombia', '#1565C0'),
    ('AUTORIDADES INDIGENAS', 'AICO', '#D35400'), ('FUERZA CIUDADANA', 'Fuerza Ciudadana', '#00897B'), ('CREEMOS', 'Creemos', '#00BCD4'),
    ('SALVACION NACIONAL', 'Salvación Nacional', '#283593')]
def lista(n):
    k = norm(n)
    if not k or k == 'CANDIDATOS TOTALES': return None
    for pat, lb, _ in LISTAS26:
        if pat in k: return lb
    return 'Otras listas'
PCOL = {lb: c for _, lb, c in LISTAS26}; PCOL['Otras listas'] = '#9E9E9E'
def load(c):
    d = gpd.read_file(os.path.join(PU, f'puestos_{c}_2026.geojson'))
    return d.drop_duplicates('codigoPuesto')
sen, cam = load('senado'), load('camara')
cam = cam.set_index('codigoPuesto')[['Partido_Ganador', 'Votos_Ganador', 'votos_validos']]
pu = sen.join(cam, on='codigoPuesto', rsuffix='_cam')
pu = pu[pu.geometry.notna() & ~pu.geometry.is_empty]
votantes = pu.votos_validos.fillna(0) + pu.votos_nulos.fillna(0) + pu.votos_no_marcados.fillna(0)
pu['part'] = (votantes / pu.potencial.replace(0, np.nan)).round(4)
pu['ls'] = pu.Partido_Ganador.map(lista); pu['lc'] = pu.Partido_Ganador_cam.map(lista)
pu['com'] = pu.comuna_localidad.fillna('').astype(str).str.replace(r'^\d+', '', regex=True).str.strip()
cnt_p = {}
os.makedirs(os.path.join(DATA, 'puestos'), exist_ok=True)
for d_, part in pu.groupby(pu.codigoMunicipioCompleto.str[:2]):
    fs = [dict(type='Feature', geometry=dict(type='Point', coordinates=[round(r.geometry.x, 5), round(r.geometry.y, 5)]),
               properties=dict(id=r.codigoPuesto, m=r.codigoMunicipioCompleto, n=r.nombrePuesto, esp=int('PUESTO CENSO' in norm(r.nombrePuesto)), dir=r.direccionPuesto, com=r.com, cat=r.categoria,
                               pot=clean(r.potencial), h=clean(r.hombres), mu=clean(r.mujeres), part=clean(r.part), vv=clean(r.votos_validos),
                               ls=r.ls, lsv=clean(r.Votos_Ganador), lc=r.lc, lcv=clean(r.Votos_Ganador_cam), vvc=clean(r.votos_validos_cam)))
          for r in part.itertuples()]
    with open(os.path.join(DATA, 'puestos', f'p_{d_}.js'), 'w', encoding='utf-8') as fh:
        fh.write(f'window.ATLAS_PUE=window.ATLAS_PUE||{{}};window.ATLAS_PUE["{d_}"]=' + json.dumps(dict(type='FeatureCollection', features=fs), ensure_ascii=False, separators=(',', ':')) + ';')
    cnt_p[d_] = len(fs)
js('listas2026.js', 'ATLAS_LISTAS', PCOL)
print('puestos', len(pu), 'en', len(cnt_p), 'departamentos;', round(sum(os.path.getsize(os.path.join(DATA, 'puestos', f)) for f in os.listdir(os.path.join(DATA, 'puestos'))) / 1e6, 1), 'MB')
