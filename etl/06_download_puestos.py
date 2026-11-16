"""Descarga los puestos de votación (escrutinio Senado y Cámara 8-mar-2026, Esri Colombia) a raw/puestos/."""
import json, os, time, urllib.request, urllib.parse
B = 'https://ags.esri.co/arcgis/rest/services/Elecciones_8_Marzo_2026'
OUTD = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'raw', 'puestos'); os.makedirs(OUTD, exist_ok=True)
def get(url, params):
    for k in range(5):
        try: return json.load(urllib.request.urlopen(url + '?' + urllib.parse.urlencode(params), timeout=300))
        except Exception as e: print('reintento', k, e, flush=True); time.sleep(5)
    raise SystemExit('fallo')
for corp in ('Senado', 'Camara'):
    feats = []
    for off in range(0, 20000, 2000):
        d = get(f'{B}/Escrutinios_{corp}_2026/MapServer/2/query', dict(where='1=1', outFields='*', returnGeometry='true', outSR=4326, geometryPrecision=6,
                f='geojson', resultOffset=off, resultRecordCount=2000, orderByFields='OBJECTID'))
        n = len(d.get('features', [])); feats += d.get('features', []); print(corp, off, n, flush=True)
        if n < 2000: break
    json.dump({'type': 'FeatureCollection', 'features': feats}, open(os.path.join(OUTD, f'puestos_{corp.lower()}_2026.geojson'), 'w', encoding='utf-8'))
    print(corp, 'total', len(feats))
