import json, urllib.request, urllib.parse, time, sys
B='https://ags.esri.co/arcgis/rest/services/LivingAtlas'
P=('ID_UNIFICADO,DEPTO,MPIO,CLASE,U_SECT_URB,U_SECT_RUR,UA2_CPOB,SEXO_M,SEXO_H,SEXO_TOTAL,EDAD_0_4,EDAD_5_9,EDAD_10_14,EDAD_15_19,EDAD_20_24,EDAD_25_29,'
   'EDAD_30_34,EDAD_35_39,EDAD_40_44,EDAD_45_49,EDAD_50_54,EDAD_55_59,EDAD_60_64,EDAD_65_69,TOTAL_MAYORES_70,GRUPO_ETNICO_INDIGENA,GRUPO_ETNICO_NEGRO,'
   'GRUPO_ETNICO_RAIZAL,GRUPO_ETNICO_PALANQUERO,LUG_NAC_OTRO_PAIS,LUG_NAC_OTRO_MPO,ALFABETA_NO,NIVEL_EDUC_UNIVERSITARIA,NIVEL_EDUC_ESP_MAES_DOC,'
   'NIVEL_EDUC_TECNICA_TECNOLOGO,NIVEL_EDUC_NINGUNO')
V=('ID_UNIFICADO,TOTAL_VIVIENDAS,ESTRATO_1,ESTRATO_2,ESTRATO_3,ESTRATO_4,ESTRATO_5,ESTRATO_6,ESTRATO_SIN_ESTRATO,ACUEDUCTO_SI,ACUEDUCTO_NO,'
   'ALCANTARILLADO_SI,ALCANTARILLADO_NO,INTERNET_SI,INTERNET_NO,ENERGIA_ELECTRICA_SI,ENERGIA_ELECTRICA_NO,PISO_EN_TIERRA')
def get(url, params):
    for k in range(5):
        try: return json.load(urllib.request.urlopen(url + '?' + urllib.parse.urlencode(params), timeout=300))
        except Exception as e: print('retry', k, e, flush=True); time.sleep(5)
    raise SystemExit('fallo')
feats=[]
for off in range(0, 14000, 1000):
    d=get(f'{B}/Censo_personas_sectores_2018/MapServer/0/query', dict(where='1=1', outFields=P, returnGeometry='true', outSR=4326,
        maxAllowableOffset=0.0002, geometryPrecision=5, f='geojson', resultOffset=off, resultRecordCount=1000, orderByFields='OBJECTID'))
    n=len(d.get('features',[])); feats+=d.get('features',[]); print('personas', off, n, flush=True)
    if n==0: break
import os; os.makedirs(os.path.join(os.path.dirname(__file__), '..', 'raw', 'sectores'), exist_ok=True); os.chdir(os.path.join(os.path.dirname(__file__), '..', 'raw', 'sectores'))
json.dump({'type':'FeatureCollection','features':feats}, open('sectores_personas.geojson','w',encoding='utf-8'))
rows=[]
for off in range(0, 14000, 2000):
    d=get(f'{B}/Censo_sectores_2018/MapServer/1/query', dict(where='1=1', outFields=V, returnGeometry='false', f='json', resultOffset=off, resultRecordCount=2000, orderByFields='OBJECTID'))
    n=len(d.get('features',[])); rows+=[f['attributes'] for f in d.get('features',[])]; print('vivienda', off, n, flush=True)
    if n==0: break
json.dump(rows, open('sectores_vivienda.json','w',encoding='utf-8'))
print('total', len(feats), len(rows))
