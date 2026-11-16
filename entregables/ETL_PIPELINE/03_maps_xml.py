"""ATLAS COLOMBIA 2027 – Paso 3: inyecta los mapas interactivos en las hojas MAPA del Excel.

Cada municipio (o departamento) es una forma DrawingML con su contorno real (geometría MGN simplificada),
color, tooltip con sus datos e hipervínculo a su fila en MUNICIPIOS / DEPARTAMENTOS. Se escribe el XML del .xlsx
directamente: no necesita Excel instalado.
Uso: python etl/03_maps_xml.py   (después de 02_build_excel.py)
"""
import os, re, math, shutil, zipfile, time
from xml.sax.saxutils import escape
import numpy as np, pandas as pd, geopandas as gpd
from shapely.geometry import box, MultiPolygon, Polygon
from common import *
from maps_common import enrich, maps_spec, COL, SIN, lab

T0 = time.time()
def log(*a): print(f'[{time.time() - T0:6.1f}s]', *a, flush=True)
XLSX = os.path.join(OUT, 'BASE_MAESTRA_ATLAS_COLOMBIA_2027.xlsx')
EMU = 12700  # EMU por punto
rd = lambda n: pd.read_csv(os.path.join(CLEAN, n + '.csv'), dtype={'codigo_divipola': str, 'departamento_id': str}, low_memory=False)
mun, dep, B, GOB = rd('dim_municipios'), rd('dim_departamentos'), rd('fact_bastiones_alcaldia'), rd('gobernadores_2024_2027')
ROW_MUN = {c: 5 + i for i, c in enumerate(mun.codigo_divipola)}   # mismo orden que la tabla tMUNICIPIOS
ROW_DEP = {d: 5 + i for i, d in enumerate(dep.departamento_id)}
M = enrich(mun, B)
SPEC = maps_spec(M, GOB)

def f(x, pct=False):
    if x is None or (isinstance(x, float) and math.isnan(x)): return 's/d'
    return f'{x:.1%}' if pct else f'{x:,.0f}'.replace(',', '.')
def s_(x): return x if isinstance(x, str) and x else 's/d'
def tip_mun(r):
    t = (f'{r.nombre_municipio} ({r.nombre_departamento}) · Alcalde: {s_(r.alcalde_actual)} · {lab(r.fuerza_gobernante) if isinstance(r.fuerza_gobernante, str) else "s/d"}'
         f' · Margen 2019: {"s/d" if pd.isna(r.margen_2019) else f"{r.margen_2019:.1f} pts"} · {r.swing} · {s_(r.clasificacion_bastion)}'
         f' · Pob. 2026: {f(r.pob_tot_2026)} · IPM: {f(r.ipm_2018 / 100 if pd.notna(r.ipm_2018) else np.nan, True)}')
    return t[:255]

# ---------------------------------------------------------------- geometrías
g = gpd.read_file(os.path.join(OUT, 'MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson'))[['mpio', 'geometry']]
g['geometry'] = g.geometry.make_valid().buffer(0)
g = g.merge(M, left_on='mpio', right_on='codigo_divipola')
def rings(geom, tol, min_area):
    if geom is None or geom.is_empty: return []
    geom = geom.simplify(tol, preserve_topology=True)
    polys = [p for p in (geom.geoms if hasattr(geom, 'geoms') else [geom]) if isinstance(p, Polygon)]
    return [list(p.exterior.coords) for p in polys if p.area >= min_area and len(p.exterior.coords) >= 4]
MAIN = (-79.1, -4.3, -66.8, 12.6)
ANDES = (-77.3, 2.3, -72.6, 8.6)
ISL = (-81.8, 12.45, -81.3, 13.45)
def clip(gdf, bb):
    x = gdf[gdf.intersects(box(*bb))].copy(); x['geometry'] = x.geometry.intersection(box(*bb)); return x[~x.geometry.is_empty]
gm = clip(g, MAIN); gm['r'] = gm.geometry.apply(lambda x: rings(x, 0.010, 0.0003))
ga = clip(g, ANDES); ga['r'] = ga.geometry.apply(lambda x: rings(x, 0.004, 0.00004))
gs = clip(g, ISL); gs['r'] = gs.geometry.apply(lambda x: rings(x, 0.0008, 0.00001))
dg = g[['departamento_id', 'geometry']].dissolve(by='departamento_id', as_index=False)
dg['geometry'] = dg.geometry.buffer(0.003).buffer(-0.003)
dg = clip(dg, MAIN); dg['r'] = dg.geometry.apply(lambda x: rings(x, 0.012, 0.002))
dg = dg.merge(dep, on='departamento_id').merge(GOB[['departamento_id', 'nombre_completo', 'partido_o_coalicion', 'fuerza', 'porcentaje_votos']], on='departamento_id', how='left')
log('geometrías', len(gm), len(ga), len(gs), len(dg))

# ---------------------------------------------------------------- DrawingML
class Drawing:
    def __init__(self):
        self.parts, self.rels, self.n = [], [], 1
    def _rel(self, target):
        rid = f'rIdH{len(self.rels) + 1}'
        mode = '' if target.startswith('#') else ' TargetMode="External"'
        self.rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{escape(target, {chr(34): "&quot;"})}"{mode}/>')
        return rid
    def poly(self, rings_, transform, color, link=None, tip='', name='p', lw=0.25, line='FFFFFF'):
        pts = [[transform(x, y) for x, y in r] for r in rings_]
        if not pts: return
        xs = [p[0] for r in pts for p in r]; ys = [p[1] for r in pts for p in r]
        x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
        w, h = max(int((x1 - x0) * EMU), 1), max(int((y1 - y0) * EMU), 1)
        paths = []
        for r in pts:
            seg = [f'<a:pt x="{int((x - x0) * EMU)}" y="{int((y - y0) * EMU)}"/>' for x, y in r]
            paths.append(f'<a:path w="{w}" h="{h}"><a:moveTo>{seg[0]}</a:moveTo>' + ''.join(f'<a:lnTo>{p}</a:lnTo>' for p in seg[1:]) + '<a:close/></a:path>')
        if color.upper().lstrip('#') == 'FFFFFF' and line == 'FFFFFF': line = 'BFBFBF'
        hl = f'<a:hlinkClick r:id="{self._rel(link)}" tooltip="{escape(tip, {chr(34): "&quot;"})}"/>' if link else ''
        self.n += 1
        self.parts.append(
            f'<xdr:absoluteAnchor><xdr:pos x="{int(x0 * EMU)}" y="{int(y0 * EMU)}"/><xdr:ext cx="{w}" cy="{h}"/>'
            f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self.n}" name="{escape(name)}">{hl}</xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="{int(x0 * EMU)}" y="{int(y0 * EMU)}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>{"".join(paths)}</a:pathLst></a:custGeom>'
            f'<a:solidFill><a:srgbClr val="{color.lstrip("#").upper()}"/></a:solidFill>'
            f'<a:ln w="{int(lw * EMU)}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:absoluteAnchor>')
    def rect(self, x, y, w, h, line='A6A6A6', dash=True, text=None, size=8, bold=False, fill=None):
        self.n += 1
        fl = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
        ln = '<a:ln><a:noFill/></a:ln>' if line is None else f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill>{"<a:prstDash val=\"dash\"/>" if dash else ""}</a:ln>'
        tx = ''
        if text:
            tx = (f'<xdr:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:pPr algn="l"/>'
                  f'<a:r><a:rPr lang="es-CO" sz="{size * 100}" b="{int(bold)}"><a:solidFill><a:srgbClr val="404040"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>'
                  f'<a:t>{escape(text)}</a:t></a:r></a:p></xdr:txBody>')
        self.parts.append(
            f'<xdr:absoluteAnchor><xdr:pos x="{int(x * EMU)}" y="{int(y * EMU)}"/><xdr:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/>'
            f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self.n}" name="r{self.n}"/><xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/><a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fl}{ln}</xdr:spPr>{tx}</xdr:sp><xdr:clientData/></xdr:absoluteAnchor>')
    def xml(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + ''.join(self.parts) + '</xdr:wsDr>')
    def rels_xml(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(self.rels) + '</Relationships>')

TOP, MW, AX, AW = 52, 470, 500, 470
def tf(bb, ox, oy, width):
    sc = width / (bb[2] - bb[0])
    return (lambda x, y: (ox + (x - bb[0]) * sc, oy + (bb[3] - y) * sc)), (bb[3] - bb[1]) * sc

def municipal(name):
    sp, d = SPEC[name], Drawing()
    t_main, h_main = tf(MAIN, 10, TOP, MW); t_and, h_and = tf(ANDES, AX, TOP, AW); t_isl, h_isl = tf(ISL, 14, TOP + 4, 60)
    for frame, rows in ((t_main, gm), (t_and, ga), (t_isl, gs)):
        lw = 0.25 if frame is t_main else 0.5
        for r in rows.itertuples():
            d.poly(r.r, frame, sp['color'](r), f"{MAP_URL}/ir/{r.codigo_divipola}.html", tip_mun(r)[:205] + ' · Clic: abre el mapa en línea', f'{r.codigo_divipola}', lw=lw)
    sc = MW / (MAIN[2] - MAIN[0])
    d.rect(10 + (ANDES[0] - MAIN[0]) * sc, TOP + (MAIN[3] - ANDES[3]) * sc, (ANDES[2] - ANDES[0]) * sc, (ANDES[3] - ANDES[1]) * sc, line='7F7F7F')
    d.rect(AX, TOP, AW, h_and, line='A6A6A6')
    d.rect(AX, TOP + h_and + 3, AW, 14, line=None, text='Zona andina ampliada (Antioquia, Eje Cafetero, Cundinamarca, Boyacá, Santanderes, Tolima, Huila, Valle)', size=7)
    d.rect(10, TOP, 70, h_isl + 10, line='A6A6A6')
    d.rect(10, TOP + h_isl + 12, 90, 12, line=None, text='San Andrés y Providencia', size=7)
    return d

def gobernaciones():
    sp, d = SPEC['MAPA GOBERNACIONES'], Drawing()
    t_main, h_main = tf(MAIN, 10, TOP, 640)
    for r in dg.itertuples():
        fz = r.fuerza if isinstance(r.fuerza, str) else None
        tip = (f'{r.nombre_departamento} · Gobernador: {r.nombre_completo if isinstance(r.nombre_completo, str) else "— (Bogotá D.C.: alcalde mayor)"}'
               f' · {r.partido_o_coalicion if isinstance(r.partido_o_coalicion, str) else ""}')[:240]
        if pd.notna(r.porcentaje_votos): tip += f' · {r.porcentaje_votos:.1%}'
        d.poly(r.r, t_main, COL.get(fz, '#D5D8DC') if fz else SIN, f"{MAP_URL}/ir/d{r.departamento_id}.html", tip[:200] + ' · Clic: abre el mapa en línea', r.departamento_id, lw=0.75)
    for r in dg.itertuples():
        x, y = t_main(*r.geometry.representative_point().coords[0])
        d.rect(x - 40, y - 6, 80, 12, line=None, text=r.nombre_departamento.title()[:18], size=7, bold=True)
    return d

DRAW = {n: municipal(n) for n in ('MAPA POLÍTICO', 'MAPA SWING', 'MAPA BASTIONES', 'MAPA TERRITORIO')}
DRAW['MAPA GOBERNACIONES'] = gobernaciones()
log('drawings', {k: v.n for k, v in DRAW.items()})

# ---------------------------------------------------------------- inyección en el .xlsx
zin = zipfile.ZipFile(XLSX)
files = {n: zin.read(n) for n in zin.namelist()}; zin.close()
wbx = files['xl/workbook.xml'].decode('utf-8'); wrel = files['xl/_rels/workbook.xml.rels'].decode('utf-8')
rid_of = {m.group(1): m.group(2) for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx)}
tgt = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', wrel)}
tgt.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', wrel)})
ct = files['[Content_Types].xml'].decode('utf-8')
for i, (name, d) in enumerate(DRAW.items(), start=1):
    sheet = 'xl/' + tgt[rid_of[escape(name)]].lstrip('/').replace('xl/', '')
    dname = f'xl/drawings/drawingMapa{i}.xml'
    files[dname] = d.xml().encode('utf-8')
    files[f'xl/drawings/_rels/drawingMapa{i}.xml.rels'] = d.rels_xml().encode('utf-8')
    srel = sheet.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    if srel in files:
        sr = files[srel].decode('utf-8')
        sr = sr.replace('</Relationships>', f'<Relationship Id="rIdMapa" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawingMapa{i}.xml"/></Relationships>')
    else:
        sr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              f'<Relationship Id="rIdMapa" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawingMapa{i}.xml"/></Relationships>')
    files[srel] = sr.encode('utf-8')
    sx = files[sheet].decode('utf-8')
    assert '<drawing ' not in sx, 'la hoja ya tiene dibujo'
    if 'xmlns:r=' not in sx[:600]:
        sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    # orden del esquema: ... pageMargins, pageSetup, headerFooter, drawing, legacyDrawing, tableParts
    for anchor in ('<legacyDrawing', '<tableParts', '<extLst', '</worksheet>'):
        if anchor in sx:
            sx = sx.replace(anchor, '<drawing r:id="rIdMapa"/>' + anchor, 1); break
    files[sheet] = sx.encode('utf-8')
    ct = ct.replace('</Types>', f'<Override PartName="/{dname}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
files['[Content_Types].xml'] = ct.encode('utf-8')
tmp = XLSX + '.tmp'
with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, b in files.items(): z.writestr(n, b)
shutil.move(tmp, XLSX)
log('OK', XLSX, round(os.path.getsize(XLSX) / 1e6, 1), 'MB')
