"""ATLAS COLOMBIA 2027 – Paso 2: CLEAN -> BASE_MAESTRA_ATLAS_COLOMBIA_2027.xlsx (+ diccionario, validación, mapa).
Uso: python etl/02_build_excel.py
"""
import os, math, datetime as dt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter as L
from common import *
from dictionary import D as DICT

NAVY, TEAL, ORANGE, GREY, WHITE = '1F3864', '00695C', 'F4B183', 'F2F2F2', 'FFFFFF'
RED, AMBER, GREEN = 'F8CBAD', 'FFE699', 'C6EFCE'
FONT = 'Arial'
rd = lambda n, **k: pd.read_csv(os.path.join(CLEAN, n + '.csv'), dtype={'codigo_divipola': str, 'departamento_id': str, 'capital_divipola': str, 'cod_divipola': str}, low_memory=False, **k)

mun, dep = rd('dim_municipios'), rd('dim_departamentos')
RA, RG, P = rd('fact_resultados_alcaldia'), rd('fact_resultados_gobernacion'), rd('fact_participacion')
EL, PAR, LOC, CO = rd('dim_elecciones'), rd('dim_partidos'), rd('dim_movimientos_coaliciones_gsc'), rd('fact_coaliciones')
ALC, GOB, IND, BAS = rd('alcaldes_2024_2027'), rd('gobernadores_2024_2027'), rd('fact_indicadores_alcaldia'), rd('fact_bastiones_alcaldia')
FC, CAL, PROS = rd('fact_pronosticos'), rd('modelo_base_calibracion'), rd('prospectos_2027')
FUE = pd.read_csv(os.path.join(ROOT, 'etl', 'manual', 'fuentes.csv'))

wb = Workbook(); wb.remove(wb.active)
SHEETS = ['INICIO', 'LEEME', 'LISTAS', 'GEO', 'DEPARTAMENTOS', 'MUNICIPIOS', 'FICHA', 'ALCALDES', 'GOBERNADORES', 'CANDIDATOS 2027', 'PARTIDOS',
          'MOVIMIENTOS', 'COALICIONES', 'ELECCIONES', 'DATOS VOTO', 'DATOS VOTO GOB', 'RESULTADOS MESA A MESA', 'HISTORIA', 'PARTICIPACION',
          'TERRITORIO', 'ENCUESTAS', 'INDICADORES', 'BASTIONES', 'SWING', 'SIMULADOR', 'PRONOSTICO', 'PROSPECTOS', 'NOTICIAS', 'FUENTES',
          'CALIDAD DATOS', 'DASHBOARD', 'MAPA POLÍTICO', 'MAPA GOBERNACIONES', 'MAPA SWING', 'MAPA BASTIONES', 'MAPA TERRITORIO', 'GEO FICHA']
DESC = {
    'INICIO': 'Portada, navegación y cómo usar el archivo', 'LEEME': 'Diccionario de datos, reglas y fórmulas', 'LISTAS': 'Catálogos, valores permitidos y umbrales configurables',
    'GEO': 'DIVIPOLA: códigos, coordenadas y jerarquía', 'DEPARTAMENTOS': '32 departamentos + Bogotá D.C.', 'MUNICIPIOS': 'Catálogo maestro (1.122 entidades)', 'FICHA': 'Ficha de un municipio: elegilo de la lista', 'GEO FICHA': 'Coordenadas para el mapa de la FICHA (auxiliar)',
    'ALCALDES': 'Alcaldes electos 2023 (período 2024-2027)', 'GOBERNADORES': 'Gobernadores electos 2023', 'CANDIDATOS 2027': 'Aspirantes y candidaturas 2027 con estado',
    'PARTIDOS': 'Partidos nacionales: estado legal y presencia', 'MOVIMIENTOS': 'Coaliciones, GSC y organizaciones locales (todas las elecciones)',
    'COALICIONES': 'Coaliciones 2023 y sus integrantes', 'ELECCIONES': 'Elecciones 2003-2027 y calendario 2027', 'DATOS VOTO': 'Resultados de alcaldía por candidato 2003-2023',
    'DATOS VOTO GOB': 'Resultados de gobernación por candidato 2003-2023', 'RESULTADOS MESA A MESA': 'Estructura para mesa a mesa (pendiente de carga)',
    'HISTORIA': 'Ganador y margen por municipio 2003-2023', 'PARTICIPACION': 'Censo, votos y participación', 'TERRITORIO': 'Población, pobreza y desempeño municipal',
    'ENCUESTAS': 'Registro de encuestas (pendiente)', 'INDICADORES': 'Margen, NEP, fragmentación, volatilidad, alternancia', 'BASTIONES': 'Bastiones por fuerza política',
    'SWING': 'Competitividad y swing 2023 (umbrales en LISTAS)', 'SIMULADOR': 'Mové el voto y mirá qué alcaldías cambian', 'PRONOSTICO': 'Modelo base v0.1 (estimaciones, no resultados)',
    'PROSPECTOS': 'Actores políticos mencionados para 2027', 'NOTICIAS': 'Monitoreo de novedades', 'FUENTES': 'Catálogo de fuentes', 'CALIDAD DATOS': 'Controles de calidad automáticos',
    'DASHBOARD': 'Cobertura, poder territorial y calidad', 'MAPA POLÍTICO': 'Mapa: fuerza que gobierna cada municipio (2024-2027)',
    'MAPA GOBERNACIONES': 'Mapa: quién gobierna cada departamento', 'MAPA SWING': 'Mapa: competitividad y swing por municipio',
    'MAPA BASTIONES': 'Mapa: bastiones partidarios 2007-2023', 'MAPA TERRITORIO': 'Mapa: pobreza multidimensional (IPM 2018)'}
WS = {}
thin = Side(style='thin', color='D9D9D9')

def sheet(name):
    ws = wb.create_sheet(name); WS[name] = ws
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = {'INICIO': ORANGE, 'DASHBOARD': ORANGE, **{m: 'C55A11' for m in SHEETS if m.startswith('MAPA')}, 'PRONOSTICO': '7030A0', 'SIMULADOR': '7030A0'}.get(name, NAVY if SHEETS.index(name) < 13 else TEAL)
    return ws

def header(ws, title, sub, width=14):
    ws['A1'] = title; ws['A1'].font = Font(name=FONT, size=16, bold=True, color=WHITE)
    ws['A2'] = sub; ws['A2'].font = Font(name=FONT, size=9, italic=True, color='404040')
    for c in range(1, width + 1):
        ws.cell(1, c).fill = PatternFill('solid', fgColor=NAVY)
    ws.row_dimensions[1].height = 28
    h = ws.cell(1, width); h.value = '← INICIO'; h.hyperlink = "#'INICIO'!A1"
    h.font = Font(name=FONT, size=9, bold=True, color=ORANGE, underline='single'); h.alignment = Alignment(horizontal='right')

def fmt_for(col):
    c = col.lower()
    if c.startswith(('pct', 'porcentaje', 'participacion', 'abstencion', 'probabilidad', 'concentracion', 'fragmentacion')) or c in ('tasa_global', 'prob_retencion', 'pct_rural_2026'):
        return '0.0%'
    if c.startswith(('votos', 'pob_', 'censo', 'sufragantes', 'validos', 'margen_votos', 'n_', 'retenciones', 'casos', 'victorias', 'candidaturas')) or c in ('votos', 'validos'):
        return '#,##0'
    if c.startswith(('margen', 'nep', 'volatilidad', 'var_', 'densidad', 'area', 'nbi', 'ipm', 'miseria', 'mdm_puntaje')):
        return '#,##0.00'
    return None

def table(ws, df, name, row=4, col=1, style='TableStyleMedium2', widths=None, wrap_header=True, freeze=True):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].where(df[c].notna(), None)
    cols = list(df.columns)
    for j, c in enumerate(cols):
        cell = ws.cell(row, col + j, c); cell.font = Font(name=FONT, bold=True, color=WHITE, size=9)
        cell.alignment = Alignment(wrap_text=wrap_header, vertical='center')
    vals = df.to_numpy(dtype=object)
    for i, r in enumerate(vals, start=1):
        for j, v in enumerate(r):
            if isinstance(v, float) and math.isnan(v): v = None
            elif isinstance(v, (np.integer,)): v = int(v)
            elif isinstance(v, (np.floating,)): v = float(v)
            ws.cell(row + i, col + j, v)
    last = row + max(len(df), 1)
    ref = f'{L(col)}{row}:{L(col + len(cols) - 1)}{last}'
    t = Table(displayName=name, ref=ref); t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)
    for j, c in enumerate(cols):
        f = fmt_for(c)
        if f:
            for r in range(row + 1, last + 1): ws.cell(r, col + j).number_format = f
        q = df[c].dropna().astype(str).str.len().quantile(0.9) if df[c].notna().any() else 10
        w = (widths or {}).get(c) or min(max(len(str(c)) + 2, min(int(q) + 2, 45)), 45)
        ws.column_dimensions[L(col + j)].width = max(w, 9)
    ws.row_dimensions[row].height = 30
    if freeze and ws.freeze_panes is None: ws.freeze_panes = ws.cell(row + 1, col + (2 if len(cols) > 6 else 0))
    return ref, last, {c: L(col + j) for j, c in enumerate(cols)}

def banner(ws, cell, text, color=AMBER, merge_to=None):
    ws[cell] = text; ws[cell].font = Font(name=FONT, bold=True, size=10, color='7F6000'); ws[cell].fill = PatternFill('solid', fgColor=color)
    ws[cell].alignment = Alignment(wrap_text=True, vertical='center')
    if merge_to: ws.merge_cells(f'{cell}:{merge_to}'); ws.row_dimensions[ws[cell].row].height = 32

for s in SHEETS: sheet(s)

# =============================================================== LISTAS (catálogos + umbrales)
ws = WS['LISTAS']; header(ws, 'LISTAS · Catálogos y valores permitidos', 'Cada columna es un catálogo usado en validaciones. Los umbrales (columna amarilla) son configurables: cambiarlos recalcula SWING y SIMULADOR.', 16)
CATS = {
    'estado_candidatura': ['Actor político identificado', 'Aspirante', 'Precandidato', 'Candidato anunciado', 'Candidato inscrito', 'Candidato en firme', 'Candidato revocado', 'Candidato retirado'],
    'dignidad': ['Alcaldía', 'Gobernación', 'Concejo', 'Asamblea', 'JAL'],
    'tipo_organizacion': ['Partido político', 'Movimiento político', 'Grupo significativo de ciudadanos', 'Coalición', 'Movimiento social', 'Organización local / sin clasificar'],
    'tipo_apoyo': ['Aval individual', 'Coalición', 'Grupo significativo de ciudadanos', 'Movimiento político', 'Movimiento social'],
    'categoria_noticia': ['Candidaturas', 'Avales', 'Coaliciones', 'Alianzas', 'Rupturas políticas', 'Campaña', 'Encuestas', 'Debates', 'Decisiones del CNE', 'Decisiones judiciales', 'Gestión municipal', 'Conflictos políticos'],
    'estado_verificacion': ['Confirmada (fuente oficial)', 'Información periodística', 'Declaración', 'Rumor', 'Por verificar'],
    'condicion_autoridad': ['Alcalde actual', 'Alcalde electo', 'Exalcalde', 'Candidato a la reelección (no permitida inmediata)', 'Posible sucesor', 'Actor político sin candidatura confirmada'],
    'clasificacion_territorio': ['Bastión consolidado', 'Bastión político', 'Bastión en declive', 'Territorio competitivo', 'Territorio swing', 'Territorio disputado', 'Territorio fragmentado', 'Territorio de alta incertidumbre'],
    'etapa_resultado': ['Preconteo', 'Escrutinio', 'Resultado declarado', 'Resultado definitivo'],
    'tipo_dato': ['Resultado oficial', 'Resultado histórico', 'Estimación estadística', 'Escenario hipotético'],
    'nivel_confiabilidad': ['Alta', 'Media-alta', 'Media', 'Baja'],
    'fuerza': sorted(set(IND.fuerza_ganadora.dropna())) + ['TODAS'],
}
for j, (k, vals) in enumerate(CATS.items(), start=1):
    ws.cell(4, j, k).font = Font(name=FONT, bold=True, color=WHITE, size=9); ws.cell(4, j).fill = PatternFill('solid', fgColor=NAVY)
    for i, v in enumerate(vals, start=5): ws.cell(i, j, v).font = Font(name=FONT, size=9)
    ws.column_dimensions[L(j)].width = 30
    wb.defined_names[f'L_{k}'] = DefinedName(f'L_{k}', attr_text=f"LISTAS!${L(j)}$5:${L(j)}${4 + len(vals)}")
UC = len(CATS) + 2
ws.cell(4, UC, 'UMBRALES (editables)').font = Font(name=FONT, bold=True, color=WHITE); ws.cell(4, UC).fill = PatternFill('solid', fgColor=TEAL)
ws.cell(4, UC + 1, 'Valor').font = Font(name=FONT, bold=True, color=WHITE); ws.cell(4, UC + 1).fill = PatternFill('solid', fgColor=TEAL)
UMB = [('U_MUY_COMPETITIVO', 'Margen máximo "muy competitivo" (pts)', 5), ('U_COMPETITIVO', 'Margen máximo "competitivo" (pts)', 10),
       ('U_MODERADO', 'Margen máximo "moderado" (pts)', 20), ('U_NEP_FRAGMENTADO', 'NEP mínimo para "fragmentado"', 3.5),
       ('U_SWING', 'Margen máximo para "swing" con cambio de fuerza (pts)', 10)]
for i, (n, lab, v) in enumerate(UMB, start=5):
    ws.cell(i, UC, lab).font = Font(name=FONT, size=9); c = ws.cell(i, UC + 1, v); c.fill = PatternFill('solid', fgColor='FFF2CC'); c.font = Font(name=FONT, bold=True)
    wb.defined_names[n] = DefinedName(n, attr_text=f"LISTAS!${L(UC + 1)}${i}")
ws.column_dimensions[L(UC)].width = 48

def dv_list(ws, rng, listname):
    d = DataValidation(type='list', formula1=f'L_{listname}', allow_blank=True); d.error = 'Valor fuera del catálogo LISTAS'; d.showErrorMessage = True
    ws.add_data_validation(d); d.add(rng)

# =============================================================== GEO
ws = WS['GEO']; header(ws, 'GEO · DIVIPOLA y relaciones geográficas', 'Fuente: DANE DIVIPOLA (datos.gov.co). Geometría municipal MGN simplificada en entregables/MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson. Localidades, comunas, corregimientos, puestos y mesas: pendiente de carga.', 12)
geo = mun[['municipio_id', 'codigo_divipola', 'nombre_municipio', 'departamento_id', 'nombre_departamento', 'tipo_entidad_territorial', 'es_distrito_especial', 'elige_alcalde', 'lat', 'lon', 'area_km2_aprox']].copy()
geo.insert(2, 'codigo_geografico', geo.codigo_divipola)
geo['geometria_geografica'] = 'GeoJSON MGN (codigo = ' + geo.codigo_divipola + ')'
table(ws, geo, 'tGEO')

# =============================================================== DEPARTAMENTOS
ws = WS['DEPARTAMENTOS']; header(ws, 'DEPARTAMENTOS · 32 departamentos + Bogotá D.C.', 'Población DANE 2026/2027; gobernador electo 2023 (Wikipedia, contrastar con escrutinio); composición de alcaldías 2023; NEP medio de alcaldías 2019 (CEDAE).', 14)
dcols = ['departamento_id', 'nombre_departamento', 'region', 'tipo', 'capital_divipola', 'capital', 'n_municipios', 'n_areas_no_municipalizadas', 'pob_tot_2026', 'pob_tot_2027',
         'area_km2_aprox', 'nbi_2018', 'gobernador_actual', 'partido_o_coalicion_gobernador', 'margen_gob_2023_pts', 'alcaldias_2023_partido', 'alcaldias_2023_coalicion',
         'alcaldias_2023_otras', 'nep_medio_alc_2019']
_, last, C = table(ws, dep[dcols], 'tDEPARTAMENTOS', widths={'partido_o_coalicion_gobernador': 50})
ws.conditional_formatting.add(f'{C["margen_gob_2023_pts"]}5:{C["margen_gob_2023_pts"]}{last}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))

# =============================================================== MUNICIPIOS (maestro)
ws = WS['MUNICIPIOS']; header(ws, 'MUNICIPIOS · Catálogo maestro', 'Una fila por entidad DIVIPOLA. Clave: municipio_id / codigo_divipola. Alcalde y organización = electos 2023 (verificar cambios posteriores). Filtrá por departamento, tipo o estado.', 14)
mcols = ['municipio_id', 'codigo_divipola', 'nombre_municipio', 'departamento_id', 'nombre_departamento', 'region', 'tipo_entidad_territorial', 'es_distrito_especial', 'elige_alcalde',
         'mdm_grupo_2020', 'area_km2_aprox', 'pob_tot_2026', 'densidad_2026', 'pob_cab_2026', 'pob_rur_2026', 'censo_electoral_2019', 'participacion_alc_2019',
         'alcalde_actual', 'partido_gobernante', 'tipo_org_gobernante', 'fuerza_gobernante', 'votos_ganador_2023', 'pct_ganador_2023', 'fuerza_2019', 'margen_2019_pts',
         'actores_2027_identificados', 'n_aspirantes_registrados_2027', 'n_candidaturas_inscritas_2027', 'estado_informacion', 'ultima_actualizacion']
mun['municipio_clave'] = mun.nombre_municipio + ' · ' + mun.nombre_departamento
mcols = mcols + ['municipio_clave']
_, last, C = table(ws, mun[mcols], 'tMUNICIPIOS', widths={'partido_gobernante': 38, 'alcalde_actual': 32, 'municipio_clave': 40})
MUN_CLAVE_COL, MUN_LAST = C['municipio_clave'], last
ws.conditional_formatting.add(f'{C["margen_2019_pts"]}5:{C["margen_2019_pts"]}{last}', CellIsRule(operator='lessThan', formula=['U_MUY_COMPETITIVO'], fill=PatternFill('solid', fgColor=RED)))
ws.conditional_formatting.add(f'{C["estado_informacion"]}5:{C["estado_informacion"]}{last}', FormulaRule(formula=[f'LEFT({C["estado_informacion"]}5,3)="Sin"'], fill=PatternFill('solid', fgColor=AMBER)))

# =============================================================== ALCALDES / GOBERNADORES
ws = WS['ALCALDES']; header(ws, 'ALCALDES · Período 2024-2027 (electos el 29-oct-2023)', 'Fuente principal: servicio Esri con datos de la Registraduría (3 primeros); 48 municipios repuestos con Wikipedia por registros inconsistentes. No hay reelección inmediata (art. 314 C.P.). Antecedentes = coincidencia por nombre.', 14)
_, last, C = table(ws, ALC, 'tALCALDES', widths={'antecedentes_electorales': 60, 'partido_o_movimiento': 38, 'integrantes_coalicion': 40, 'condicion_actual': 40})
dv_list(ws, f'{C["condicion_actual"]}5:{C["condicion_actual"]}{last}', 'condicion_autoridad')
ws = WS['GOBERNADORES']; header(ws, 'GOBERNADORES · Período 2024-2027', 'Fuente: Wikipedia (resultados completos por departamento); contrastar con escrutinio oficial. San Andrés: elección anulada por doble militancia según prensa — verificar titular.', 14)
table(ws, GOB, 'tGOBERNADORES', widths={'partido_o_coalicion': 60, 'condicion_actual': 45})

# =============================================================== CANDIDATOS 2027
ws = WS['CANDIDATOS 2027']; header(ws, 'CANDIDATOS 2027 · Aspiraciones y candidaturas', 'Sólo aspirantes/precandidatos con manifestación pública. Inscripción oficial: 30-jun al 31-jul-2027; lista oficial: 8-ago-2027. Nadie figura todavía como inscrito. Los "nombres que suenan" están en PROSPECTOS.', 16)
c27 = PROS[PROS.estado.isin(['Aspirante', 'Precandidato', 'Candidato anunciado'])].copy()
CAND27 = pd.DataFrame(dict(candidato_id=[hid('C27', n, d) for n, d in zip(c27.nombre, c27.dignidad)], nombre_completo=c27.nombre, territorio=c27.territorio,
                           codigo_divipola=c27.cod_divipola, dignidad=c27.dignidad, partido_mencionado=c27.partido_mencionado, coalicion_id='', grupo_significativo_id='',
                           numero_lista_o_identificador='', estado_candidatura=c27.estado, fecha_inscripcion='', aval_partidario='Pendiente', es_alcalde_actual='No',
                           fue_candidato_anteriormente='Ver PROSPECTOS / DATOS VOTO', trayectoria_politica=c27.cargo_o_trayectoria, plan_gobierno_url='', fotografia_url='',
                           fuente=c27.fuente_id, descripcion_fuente=c27.descripcion_fuente, fecha_verificacion=FECHA_CORTE))
banner(ws, 'A3', f'Corte {FECHA_CORTE}: {len(CAND27)} aspiraciones con manifestación pública; 0 candidaturas inscritas. Agregar filas al final de la tabla; el estado se elige de la lista.', merge_to='L3')
_, last, C = table(ws, CAND27, 'tCANDIDATOS2027', row=5, widths={'descripcion_fuente': 60, 'trayectoria_politica': 40})
dv_list(ws, f'{C["estado_candidatura"]}6:{C["estado_candidatura"]}{last + 200}', 'estado_candidatura')
for est, col in (('Candidato inscrito', GREEN), ('Candidato en firme', GREEN), ('Candidato revocado', RED), ('Candidato retirado', RED), ('Aspirante', AMBER), ('Precandidato', AMBER)):
    ws.conditional_formatting.add(f'{C["estado_candidatura"]}6:{C["estado_candidatura"]}{last + 200}', CellIsRule(operator='equal', formula=[f'"{est}"'], fill=PatternFill('solid', fgColor=col)))

# =============================================================== PARTIDOS / MOVIMIENTOS / COALICIONES
ws = WS['PARTIDOS']; header(ws, 'PARTIDOS · Organizaciones nacionales', 'Estado legal 2026: listado de 21 organizaciones con personería (El Espectador, 05-08-2026) – verificar contra CNE. Victorias = alcaldías/gobernaciones ganadas con aval propio (no incluye coaliciones). Orientación: pendiente de fuente.', 14)
pc = ['organizacion_id', 'org_key', 'nombre_oficial', 'sigla', 'nombre', 'tipo', 'ambito', 'estado_legal_2026', 'orientacion_politica_documentada', 'primer_anio', 'ultimo_anio', 'candidaturas', 'victorias'] + sorted(c for c in PAR.columns if c.startswith('victorias_')) + ['fuente_id']
PAR = PAR.rename(columns={'nombre': 'nombre_en_fuente'}); pc[pc.index('nombre')] = 'nombre_en_fuente'
_, last, C = table(ws, PAR[pc], 'tPARTIDOS', widths={'nombre_oficial': 40, 'nombre_en_fuente': 40, 'estado_legal_2026': 45})
ws.conditional_formatting.add(f'{C["estado_legal_2026"]}5:{C["estado_legal_2026"]}{last}', FormulaRule(formula=[f'LEFT({C["estado_legal_2026"]}5,3)="Con"'], fill=PatternFill('solid', fgColor=GREEN)))
ws = WS['MOVIMIENTOS']; header(ws, 'MOVIMIENTOS · Coaliciones, GSC y organizaciones locales', 'Organizaciones que no son partidos nacionales, una por territorio y elección. "SIN NOMBRE EN FUENTE" = la fuente CEDAE no trae el nombre de ese código.', 14)
LOC = LOC.rename(columns={'nombre': 'nombre_organizacion'})
table(ws, LOC[['organizacion_id', 'nombre_organizacion', 'tipo', 'ambito', 'dignidad', 'departamento_id', 'codigo_divipola', 'nombre_territorio', 'primer_anio', 'candidaturas', 'victorias', 'fuente_id']], 'tMOVIMIENTOS', widths={'nombre_organizacion': 50})
ws = WS['COALICIONES']; header(ws, 'COALICIONES · Elecciones 2023', 'Integrantes según Wikipedia; "Pendiente de fuente" cuando no se pudo documentar. Los votos de una coalición NO se reparten entre sus partidos.', 14)
table(ws, CO, 'tCOALICIONES', widths={'organizaciones_integrantes': 70, 'nombre_coalicion': 45})

# =============================================================== ELECCIONES (+ calendario)
ws = WS['ELECCIONES']; header(ws, 'ELECCIONES · 2003-2027', 'Elecciones territoriales ordinarias. 2027: Resolución 12338 de 2026 de la Registraduría (31-oct-2027).', 12)
_, last, _ = table(ws, EL, 'tELECCIONES', freeze=False)
CALEN = pd.DataFrame([
    ('2026-10-28', 'Resolución 12343: requisitos y número mínimo de firmas'), ('2026-10-31', 'Inicio inscripción de cédulas y registro de comités inscriptores (GSC, movimientos sociales, voto en blanco)'),
    ('2027-06-30', 'Inicio de inscripción de candidatos; cierre del registro de comités; suspensión de cédulas de primera vez'),
    ('2027-07-31', 'Cierre de inscripción de candidatos y entrega de firmas; inicio de propaganda en espacio público'), ('2027-08-02', 'Inicio del plazo de modificación de candidaturas (hasta 06-ago)'),
    ('2027-08-05', 'Inicio de publicidad en medios'), ('2027-08-08', 'Publicación de la lista oficial de candidatos inscritos'), ('2027-08-31', 'Cierre de inscripción de cédulas; inicio de espacios gratuitos en radio y TV'),
    ('2027-09-30', 'Censo electoral definitivo'), ('2027-10-16', 'Plazo de sorteo y publicación de jurados'), ('2027-10-20', 'Último plazo de reemplazo por muerte o incapacidad'),
    ('2027-10-31', 'ELECCIONES TERRITORIALES'), ('2027-11-02', 'Inicio de escrutinios generales')], columns=['fecha', 'hito'])
CALEN['fuente_id'] = 'F-REG-CAL2027;F-NEWS-CAL'
CALEN['dias_restantes'] = None
r0 = last + 3; ws.cell(r0 - 1, 1, 'CALENDARIO ELECTORAL 2027').font = Font(name=FONT, bold=True, size=12, color=NAVY)
_, l2, C2 = table(ws, CALEN, 'tCALENDARIO', row=r0, style='TableStyleMedium7', widths={'hito': 90}, freeze=False)
for r in range(r0 + 1, l2 + 1):
    ws[f'{C2["fecha"]}{r}'].value = dt.date.fromisoformat(ws[f'{C2["fecha"]}{r}'].value); ws[f'{C2["fecha"]}{r}'].number_format = 'dd-mmm-yyyy'
    ws[f'{C2["dias_restantes"]}{r}'] = f'={C2["fecha"]}{r}-TODAY()'; ws[f'{C2["dias_restantes"]}{r}'].number_format = '#,##0'

# =============================================================== DATOS VOTO
ws = WS['DATOS VOTO']; header(ws, 'DATOS VOTO · Alcaldía por candidato', '2003-2019: CEDAE (todos los candidatos, oficial consolidado). 2023: tres primeros (Esri) o sólo ganador (Wikipedia). Las coaliciones conservan su nombre; no se reparten votos.', 14)
RA['clave'] = RA.codigo_divipola.astype(str) + '|' + RA.anio.astype(str) + '|' + RA.posicion.map(lambda x: '' if pd.isna(x) else str(int(x)))
table(ws, RA.drop(columns=['nivel']), 'tDATOSVOTO', widths={'organizacion': 40, 'nombre_candidato': 32, 'estado_validacion': 40})
ws = WS['DATOS VOTO GOB']; header(ws, 'DATOS VOTO GOB · Gobernación por candidato', '2003-2019: CEDAE por municipio (sumar por departamento para el total). 2023: Wikipedia, total departamental. Tabla separada de alcaldías: denominadores propios.', 14)
table(ws, RG, 'tDATOSVOTOGOB', widths={'organizacion': 45, 'nombre_candidato': 32, 'estado_validacion': 40})

# =============================================================== MESA A MESA (estructura)
ws = WS['RESULTADOS MESA A MESA']; header(ws, 'RESULTADOS MESA A MESA · Estructura', 'Sin datos cargados. Ver procedimiento de carga en README (ETL_PIPELINE/03_mesa_a_mesa.md).', 12)
banner(ws, 'A3', 'PENDIENTE: la Registraduría publica E-14 y escrutinio mesa a mesa (2023) en sus portales. No se descargaron en esta versión. No se inventan filas ni correspondencias mesa–geometría.', merge_to='N3')
mm = pd.DataFrame(columns=['mesa_resultado_id', 'eleccion_id', 'dignidad', 'departamento_id', 'municipio_id', 'codigo_divipola', 'zona', 'puesto_codigo', 'puesto_nombre', 'mesa',
                           'candidato_id', 'organizacion_id', 'votos', 'etapa_resultado', 'formulario', 'id_formulario', 'fuente_id', 'fecha_carga'])
_, last, C = table(ws, mm, 'tMESAS', row=5)
dv_list(ws, f'{C["etapa_resultado"]}6:{C["etapa_resultado"]}5000', 'etapa_resultado')

# =============================================================== HISTORIA
ws = WS['HISTORIA']; header(ws, 'HISTORIA · Alcaldías 2003-2023', 'Organización ganadora y margen (pts) por ciclo. Fuerza = partido canónico o COALICION/GSC/OTRA_LOCAL. 2003: la fuente cubre 909 municipios.', 14)
hw = IND.pivot(index='codigo_divipola', columns='anio', values='organizacion_ganadora')
hw.columns = [f'org_{c}' for c in hw.columns]
hm = IND.pivot(index='codigo_divipola', columns='anio', values='margen_pts'); hm.columns = [f'margen_{c}' for c in hm.columns]
hp = IND.pivot(index='codigo_divipola', columns='anio', values='participacion'); hp.columns = [f'participacion_{c}' for c in hp.columns]
H = mun[['codigo_divipola', 'nombre_municipio', 'nombre_departamento']].set_index('codigo_divipola').join([hw, hm, hp], how='inner').reset_index()
H = H[['codigo_divipola', 'nombre_municipio', 'nombre_departamento'] + [f'{p}_{y}' for y in [2003, 2007, 2011, 2015, 2019, 2023] for p in ('org', 'margen') if f'{p}_{y}' in H] + [c for c in H if c.startswith('participacion_')]]
H = H.drop(columns=[c for c in H if c in ('participacion_2023', 'margen_2023')])
_, last, C = table(ws, H, 'tHISTORIA', widths={f'org_{y}': 28 for y in range(2003, 2024)})
for y in [2003, 2007, 2011, 2015, 2019, 2023]:
    if f'margen_{y}' in C:
        ws.conditional_formatting.add(f'{C[f"margen_{y}"]}5:{C[f"margen_{y}"]}{last}', ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=15, mid_color='FFEB84', end_type='num', end_value=40, end_color='63BE7B'))

# =============================================================== PARTICIPACION
ws = WS['PARTICIPACION']; header(ws, 'PARTICIPACIÓN · Alcaldía y gobernación', 'Participación = sufragantes / censo electoral. Censo disponible en CEDAE desde 2007 (archivo de alcaldía). Gobernación 2023: totales departamentales (Wikipedia).', 14)
_, last, C = table(ws, P, 'tPARTICIPACION', widths={'nota': 50})
ws.conditional_formatting.add(f'{C["participacion"]}5:{C["participacion"]}{last}', CellIsRule(operator='greaterThan', formula=['1'], fill=PatternFill('solid', fgColor=RED)))

# =============================================================== TERRITORIO
ws = WS['TERRITORIO']; header(ws, 'TERRITORIO · Población, pobreza y desempeño', 'Año de referencia en el nombre de cada columna. Población: DANE proyecciones (act. 30-07-2025). NBI e IPM: CNPV 2018. MDM: DNP 2020. Superficie aproximada desde geometría MGN.', 14)
tcols = ['municipio_id', 'codigo_divipola', 'nombre_municipio', 'nombre_departamento', 'pob_tot_2018', 'pob_tot_2026', 'pob_tot_2027', 'pob_cab_2026', 'pob_rur_2026', 'pct_rural_2026',
         'area_km2_aprox', 'densidad_2026', 'nbi_2018', 'miseria_2018', 'ipm_2018', 'ipm_cab_2018', 'ipm_rur_2018', 'mdm_grupo_2020', 'mdm_puntaje_2020']
T = mun[tcols].copy(); T['pct_crecimiento_2018_2026'] = (T.pob_tot_2026 / T.pob_tot_2018 - 1).round(4)
_, last, C = table(ws, T, 'tTERRITORIO')
for c in ('nbi_2018', 'ipm_2018'):
    ws.conditional_formatting.add(f'{C[c]}5:{C[c]}{last}', ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B'))
banner(ws, 'A3', 'Pendientes con fuente identificada: TerriData (DNP), FUT/CHIP y CUIPO (ingresos, gastos, transferencias, dependencia fiscal), estructura por edad y sexo (DANE). Ver FUENTES.', merge_to='N3')

# =============================================================== ENCUESTAS
ws = WS['ENCUESTAS']; header(ws, 'ENCUESTAS · Registro de encuestas publicadas', 'Registrar la fecha de campo, no sólo la de publicación. No extrapolar encuestas nacionales a municipios.', 14)
banner(ws, 'A3', 'Sin encuestas registradas al corte. No se cargó ninguna encuesta municipal o departamental con ficha técnica verificada. Completar desde el registro de encuestas del CNE.', merge_to='N3')
enc = pd.DataFrame(columns=['encuesta_id', 'ambito', 'departamento_id', 'municipio_id', 'empresa_encuestadora', 'fecha_inicio_campo', 'fecha_fin_campo', 'fecha_publicacion',
                            'tamano_muestra', 'universo', 'metodologia', 'margen_error', 'nivel_confianza', 'candidato', 'partido', 'intencion_voto', 'indecisos', 'blanco',
                            'no_sabe_no_responde', 'fuente', 'url', 'estado_verificacion', 'antiguedad_dias'])
_, last, C = table(ws, enc, 'tENCUESTAS', row=5)
dv_list(ws, f'{C["estado_verificacion"]}6:{C["estado_verificacion"]}2000', 'estado_verificacion')
for r in range(6, 300): ws[f'{C["antiguedad_dias"]}{r}'] = f'=IF({C["fecha_fin_campo"]}{r}="","",TODAY()-{C["fecha_fin_campo"]}{r})'

# =============================================================== INDICADORES
ws = WS['INDICADORES']; header(ws, 'INDICADORES · Alcaldía por municipio y elección', 'Margen, NEP (Laakso-Taagepera), fragmentación de Rae, concentración, volatilidad de Pedersen, alternancia y variación de participación. Fórmulas en LEEME. 2023: sin NEP (fuente con 3 candidatos).', 14)
ic = ['municipio_id', 'codigo_divipola', 'nombre_municipio', 'departamento_id', 'anio', 'n_candidatos', 'ganador', 'organizacion_ganadora', 'fuerza_ganadora', 'tipo_org_ganadora',
      'votos_ganador', 'pct_ganador', 'segundo', 'organizacion_segundo', 'pct_segundo', 'margen_pts', 'nep', 'fragmentacion_rae', 'concentracion_top2', 'participacion',
      'var_participacion_pts', 'volatilidad_pedersen', 'fuerza_anterior', 'alternancia', 'fuente_id']
IND['clave'] = IND.codigo_divipola.astype(str) + '|' + IND.anio.astype(str); ic = ic + ['clave']
_, last, C = table(ws, IND[ic], 'tINDICADORES', widths={'organizacion_ganadora': 36, 'organizacion_segundo': 30})
col = L(len(ic) + 1); ws[f'{col}4'] = 'competitividad (fórmula)'; ws[f'{col}4'].font = Font(name=FONT, bold=True, size=9)
M, N = C['margen_pts'], C['nep']
for r in range(5, last + 1):
    ws[f'{col}{r}'] = (f'=IF({M}{r}="","Sin dato",IF(AND({N}{r}<>"",{N}{r}>=U_NEP_FRAGMENTADO),"Territorio fragmentado",'
                       f'IF({M}{r}<U_MUY_COMPETITIVO,"Muy competitivo",IF({M}{r}<U_COMPETITIVO,"Competitivo",IF({M}{r}<U_MODERADO,"Moderado","Dominante")))))')
ws.column_dimensions[col].width = 22
ws.conditional_formatting.add(f'{M}5:{M}{last}', ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=15, mid_color='FFEB84', end_type='num', end_value=40, end_color='63BE7B'))

# =============================================================== BASTIONES
ws = WS['BASTIONES']; header(ws, 'BASTIONES · Fuerzas dominantes por municipio', 'Consolidado: misma fuerza en 2015, 2019 y 2023. Político: misma fuerza ≥3 de 5 ciclos (2007-2023) incl. 2023. En declive: ≥3 victorias 2007-2019 pero perdió 2023. Sólo partidos canónicos (las coaliciones no forman bastión partidario).', 14)
bc = ['municipio_id', 'codigo_divipola', 'nombre_municipio', 'nombre_departamento', 'clasificacion_bastion', 'fuerza_dominante', 'victorias_fuerza_2007_2023'] + [f'fuerza_{y}' for y in [2003, 2007, 2011, 2015, 2019, 2023]] + ['cambio_2019_2023']
ws.freeze_panes = 'C5'
_, last, C = table(ws, BAS[bc], 'tBASTIONES')
for k, colr in (('Bastión consolidado', '63BE7B'), ('Bastión político', GREEN), ('Bastión en declive', AMBER)):
    ws.conditional_formatting.add(f'{C["clasificacion_bastion"]}5:{C["clasificacion_bastion"]}{last}', CellIsRule(operator='equal', formula=[f'"{k}"'], fill=PatternFill('solid', fgColor=colr)))
# resumen por fuerza
bsum = BAS[BAS.clasificacion_bastion != 'Sin bastión partidario'].pivot_table(index='fuerza_dominante', columns='clasificacion_bastion', values='codigo_divipola', aggfunc='count', fill_value=0).reset_index()
c0 = len(bc) + 2; ws.cell(3, c0, 'Resumen por fuerza').font = Font(name=FONT, bold=True, color=NAVY)
table(ws, bsum, 'tBASTIONES_RES', row=4, col=c0, style='TableStyleLight9')

# =============================================================== SWING (fórmulas con umbrales)
ws = WS['SWING']; header(ws, 'SWING · Competitividad y cambios de fuerza', 'Fórmulas con umbrales de LISTAS. Márgenes: 2015 y 2019 (oficiales, CEDAE); 2023 sólo aporta quién ganó (la fuente no da un 2.º confiable). Swing = cambió la fuerza 2019→2023 y el margen 2019 era < U_SWING.', 16)
S = BAS[['municipio_id', 'codigo_divipola', 'nombre_municipio', 'nombre_departamento', 'fuerza_2015', 'margen_2015', 'fuerza_2019', 'margen_2019', 'fuerza_2023', 'nep_2019', 'cambio_2019_2023']].copy()
S = S[S.fuerza_2023.notna()]
S['clasificacion'] = None
_, last, C = table(ws, S, 'tSWING')
M15, M19, F23, NP, CB, CL = (C[k] for k in ('margen_2015', 'margen_2019', 'fuerza_2023', 'nep_2019', 'cambio_2019_2023', 'clasificacion'))
for r in range(5, last + 1):
    ws[f'{CL}{r}'] = (f'=IF({M19}{r}="","Sin dato",IF(AND({CB}{r}="Sí",{M19}{r}<U_SWING),"Territorio swing",'
                      f'IF(AND(OR({F23}{r}="COALICION",{F23}{r}="GSC",{F23}{r}="OTRA_LOCAL"),{M19}{r}<U_COMPETITIVO),"Territorio de alta incertidumbre",'
                      f'IF(AND({M15}{r}<>"",{M15}{r}<U_COMPETITIVO,{M19}{r}<U_COMPETITIVO),"Territorio disputado",'
                      f'IF({M19}{r}<U_MUY_COMPETITIVO,"Territorio competitivo",IF(AND({NP}{r}<>"",{NP}{r}>=U_NEP_FRAGMENTADO),"Territorio fragmentado","Estable"))))))')
ws.column_dimensions[CL].width = 30
for k, colr in (('Territorio swing', 'F4B183'), ('Territorio disputado', RED), ('Territorio competitivo', AMBER), ('Territorio de alta incertidumbre', 'D9D2E9')):
    ws.conditional_formatting.add(f'{CL}5:{CL}{last}', CellIsRule(operator='equal', formula=[f'"{k}"'], fill=PatternFill('solid', fgColor=colr)))
c0 = len(S.columns) + 2
ws.cell(4, c0, 'Clasificación').font = Font(name=FONT, bold=True, color=WHITE); ws.cell(4, c0).fill = PatternFill('solid', fgColor=TEAL)
ws.cell(4, c0 + 1, 'Municipios').font = Font(name=FONT, bold=True, color=WHITE); ws.cell(4, c0 + 1).fill = PatternFill('solid', fgColor=TEAL)
for i, k in enumerate(['Territorio swing', 'Territorio disputado', 'Territorio competitivo', 'Territorio de alta incertidumbre', 'Territorio fragmentado', 'Estable', 'Sin dato'], start=5):
    ws.cell(i, c0, k); ws.cell(i, c0 + 1, f'=COUNTIF({CL}$5:{CL}${last},{L(c0)}{i})')
ws.column_dimensions[L(c0)].width = 32

# =============================================================== SIMULADOR
ws = WS['SIMULADOR']; header(ws, 'SIMULADOR · ¿Qué alcaldías cambian?', 'Base: elección 2019 (último resultado oficial completo). Elegí una fuerza y un swing (pts): si la fuerza ganó, su margen baja 2×swing; "TODAS" afecta a todos los ganadores. Escenario hipotético, no pronóstico.', 14)
ws['A4'] = 'Fuerza afectada'; ws['B4'] = 'TODAS'; ws['A5'] = 'Swing en contra del incumbente (pts)'; ws['B5'] = 3
for c in ('B4', 'B5'):
    ws[c].fill = PatternFill('solid', fgColor='FFF2CC'); ws[c].font = Font(name=FONT, bold=True, size=12)
for c in ('A4', 'A5', 'A7', 'A8', 'A9'): ws[c].font = Font(name=FONT, bold=True)
dv_list(ws, 'B4', 'fuerza')
ws.column_dimensions['A'].width = 38; ws.column_dimensions['B'].width = 22
SIM = IND[IND.anio == 2019][['codigo_divipola', 'nombre_municipio', 'departamento_id', 'organizacion_ganadora', 'fuerza_ganadora', 'margen_pts']].dropna(subset=['margen_pts']).copy()
SIM['margen_simulado'] = None; SIM['cambia'] = None
ws.freeze_panes = 'A12'
_, last, C = table(ws, SIM, 'tSIMULADOR', row=11, widths={'organizacion_ganadora': 36})
FZ, MG, MS, CA = C['fuerza_ganadora'], C['margen_pts'], C['margen_simulado'], C['cambia']
for r in range(12, last + 1):
    ws[f'{MS}{r}'] = f'=IF(OR($B$4="TODAS",{FZ}{r}=$B$4),{MG}{r}-2*$B$5,{MG}{r})'; ws[f'{MS}{r}'].number_format = '0.00'
    ws[f'{CA}{r}'] = f'=IF({MS}{r}<0,"CAMBIA","Retiene")'
ws.conditional_formatting.add(f'{CA}12:{CA}{last}', CellIsRule(operator='equal', formula=['"CAMBIA"'], fill=PatternFill('solid', fgColor=RED)))
ws['A7'] = 'Alcaldías simuladas (margen 2019)'; ws['B7'] = f'=COUNTA({MS}12:{MS}{last})'
ws['A8'] = 'Alcaldías que cambian'; ws['B8'] = f'=COUNTIF({CA}12:{CA}{last},"CAMBIA")'
ws['A9'] = '% que cambia'; ws['B9'] = '=IF(B7=0,0,B8/B7)'; ws['B9'].number_format = '0.0%'

# =============================================================== PRONOSTICO
ws = WS['PRONOSTICO']; header(ws, 'PRONÓSTICO · Modelo base v0.1 (retención de la fuerza)', 'ESTIMACIONES, NO RESULTADOS. P(la misma fuerza retiene la alcaldía) según la tasa histórica 2003-2023 de retención, separando fuerzas que repiten de las que ganaron por primera vez (suavizado bayesiano). Cada versión se agrega con su forecast_id; no se sobrescribe.', 16)
cal_txt = CAL.backtest_2019_2023.iloc[0]
banner(ws, 'A3', f'Backtest (entrenado 2003→2019, evaluado 2019→2023): {cal_txt}. Mejora pequeña sobre la constante: usar como línea de base, no como predicción de candidatos.', merge_to='N3')
ws.freeze_panes = 'D6'
_, last, C = table(ws, FC, 'tPRONOSTICO', row=5, widths={'observaciones': 60, 'organizacion_incumbente': 36})
ws.conditional_formatting.add(f'{C["probabilidad_retencion_fuerza"]}6:{C["probabilidad_retencion_fuerza"]}{last}', ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF', end_type='num', end_value=0.5, end_color='8EA9DB'))
c0 = len(FC.columns) + 2; ws.cell(4, c0, 'Calibración por tramo de margen').font = Font(name=FONT, bold=True, color=NAVY)
table(ws, CAL, 'tCALIBRACION', row=5, col=c0, style='TableStyleLight9')

# =============================================================== PROSPECTOS / NOTICIAS / FUENTES
ws = WS['PROSPECTOS']; header(ws, 'PROSPECTOS · Actores políticos 2027', 'Nombres mencionados en prensa. NO son candidaturas inscritas. Prioridad de investigación editable.', 14)
PR = PROS.rename(columns={'nombre': 'nombre', 'cod_divipola': 'codigo_territorio'}).copy()
PR.insert(0, 'prospecto_id', [hid('PRO', n, t, d) for n, t, d in zip(PR.nombre, PR.territorio, PR.dignidad)])
PR['prioridad_investigacion'] = np.where(PR.territorio.str.contains('Bogotá|Medellín|Cali|Barranquilla'), 'Alta', 'Media')
PR['fecha_actualizacion'] = FECHA_CORTE
PR['clave'] = PR.codigo_territorio.astype(str) + '|' + (PR.groupby('codigo_territorio').cumcount() + 1).astype(str)
_, last, C = table(ws, PR, 'tPROSPECTOS', widths={'descripcion_fuente': 60, 'cargo_o_trayectoria': 40})
dv_list(ws, f'{C["estado"]}5:{C["estado"]}{last + 200}', 'estado_candidatura')
ws = WS['NOTICIAS']; header(ws, 'NOTICIAS · Monitoreo político', 'Diferenciar información confirmada, declaraciones, prensa y rumores (columna estado_verificacion).', 12)
NEWS = pd.DataFrame([
    ('2026-10-28', '', '', 'Registraduría', '', '', 'Registraduría fija el calendario electoral 2027 (Res. 12338): elecciones el 31-oct-2027', 'Inscripción de candidatos 30-jun a 31-jul-2027; censo definitivo 30-sep-2027.', 'Decisiones del CNE', 'F-REG-CAL2027', 'Confirmada (fuente oficial)'),
    ('2026-08-05', '', '', '', '', '', '21 organizaciones con personería jurídica tras el nuevo Congreso', 'Defensores de la Patria, el último en obtenerla; Pacto Histórico, la mayor bancada.', 'Avales', 'F-NEWS-PART', 'Información periodística'),
    ('2026-09-24', '11001', '11', 'Claudia López', 'Pacto Histórico', '', 'Pacto Histórico perfila a Claudia López (Bogotá), Francia Márquez (Cali) y Gustavo Bolívar (Cundinamarca)', 'Versiones sobre la baraja del Pacto para 2027.', 'Candidaturas', 'F-NEWS-01', 'Información periodística'),
    ('2026-09-25', '11001', '11', 'Camilo Romero', '', '', 'Camilo Romero se sumaría a la contienda por la Alcaldía de Bogotá', 'También Carlos Carrillo, Ovidio Claros; nombres para Cali, Valle, Cundinamarca y Medellín.', 'Candidaturas', 'F-NEWS-03', 'Información periodística'),
    ('2026-10', '', '', '', '', '', 'Exalcaldes y exgobernadores que buscarían repetir en 2027', 'Peñalosa, López, Pumarejo, Gaviria y exgobernadores de 12 departamentos.', 'Candidaturas', 'F-NEWS-02', 'Información periodística'),
    ('2026-10', '05001', '05', 'Manuel Villa', 'Creemos', '', 'Manuel Villa renuncia y se suma a los precandidatos para 2027', 'Titular enlazado en El Colombiano; leer la nota completa.', 'Candidaturas', 'F-NEWS-04', 'Por verificar'),
], columns=['fecha', 'codigo_divipola', 'departamento_id', 'persona_relacionada', 'partido', 'coalicion', 'titulo', 'resumen', 'categoria', 'fuente_id', 'estado_verificacion'])
NEWS.insert(0, 'noticia_id', [hid('NOT', t) for t in NEWS.titulo]); NEWS['url'] = NEWS.fuente_id.map(FUE.set_index('fuente_id').url)
_, last, C = table(ws, NEWS, 'tNOTICIAS', widths={'titulo': 60, 'resumen': 60, 'url': 40})
dv_list(ws, f'{C["categoria"]}5:{C["categoria"]}{last + 500}', 'categoria_noticia'); dv_list(ws, f'{C["estado_verificacion"]}5:{C["estado_verificacion"]}{last + 500}', 'estado_verificacion')
ws = WS['FUENTES']; header(ws, 'FUENTES · Catálogo', 'Cada registro de la base lleva un fuente_id de esta tabla.', 10)
_, last, C = table(ws, FUE, 'tFUENTES', widths={'observaciones': 70, 'url': 50, 'nombre_documento': 50})
for r in range(5, last + 1):
    u = ws[f'{C["url"]}{r}'].value
    if u: ws[f'{C["url"]}{r}'].hyperlink = u; ws[f'{C["url"]}{r}'].font = Font(name=FONT, color='0563C1', underline='single', size=9)

# =============================================================== CALIDAD DATOS (controles)
chk = []
def add(control, total, ok, nivel, detalle):
    chk.append(dict(control=control, registros=total, correctos=ok, pendientes_o_errores=total - ok, pct_ok=(ok / total if total else None), nivel=nivel, detalle=detalle))
eleg = mun[mun.elige_alcalde == 'Sí']
add('Cobertura DIVIPOLA: municipios que eligen alcalde con alcalde 2023 identificado', len(eleg), eleg.alcalde_actual.notna().sum(), 'Advertencia',
    'Faltan: ' + ', '.join(eleg[eleg.alcalde_actual.isna()].nombre_municipio.head(30)) + ' (Esri inconsistente sin reposición en Wikipedia, o municipio nuevo: Nuevo Belén de Bajirá).')
codes = set(mun.codigo_divipola)
for nm, df in (('alcaldía', RA), ('gobernación', RG[RG.nivel == 'Municipio'])):
    add(f'Integridad territorial: resultados de {nm} con código DIVIPOLA vigente', len(df), df.codigo_divipola.isin(codes).sum(), 'Advertencia',
        'Códigos históricos no vigentes: ' + ', '.join(sorted(set(df.codigo_divipola) - codes)) + ' (municipios suprimidos o recodificados).')
add('Duplicados de resultado_id (alcaldía + gobernación)', len(RA) + len(RG), len(RA) + len(RG) - pd.concat([RA.resultado_id, RG.resultado_id]).duplicated().sum(), 'Error', 'Debe ser 0 duplicados.')
pp = P[(P.dignidad == 'ALC') & P.censo_electoral.notna()]
add('Consistencia de participación (0-100 %) alcaldía 2007-2019', len(pp), pp.participacion.between(0, 1).sum(), 'Advertencia', 'Participación >100 % indica censo/sufragantes inconsistentes en la fuente.')
s = RA[RA.anio < 2023].groupby(['anio', 'codigo_divipola']).agg(sv=('votos', 'sum'), vv=('votos_validos', 'first'))
add('Coherencia de votos: suma de candidatos ≤ válidos (2003-2019)', len(s), (s.sv <= s.vv).sum(), 'Error', 'válidos = candidatos + blanco.')
g = RG[RG.anio == 2023].groupby('departamento_id').agg(sv=('votos', 'sum'), vv=('votos_validos', 'first'))
gb = P[(P.dignidad == 'GOB') & (P.anio == 2023)].set_index('departamento_id').votos_blancos.fillna(0)
add('Gobernación 2023: candidatos + blanco = válidos (Wikipedia)', len(g), int(((g.sv + g.index.map(gb)) - g.vv).abs().le(5).sum()), 'Advertencia', 'Tolerancia ±5 votos.')
for y in [2003, 2007, 2011, 2015, 2019, 2023]:
    d = RA[RA.anio == y]
    add(f'Organizaciones identificadas por nombre – alcaldía {y}', len(d), (~d.organizacion.str.startswith('SIN NOMBRE')).sum(), 'Advertencia', 'CEDAE no trae nombre de partido; se infiere del concejo del mismo año y de Esri 2019.')
d23 = RA[RA.anio == 2023]
add('Porcentajes 2023 calculables (válidos estimados)', len(d23), d23.porcentaje_votos.notna().sum(), 'Advertencia', 'Sólo donde Wikipedia trae el % del ganador.')
add('Alcaldía 2023: 2.º y 3.º lugar confiables', len(d23), int(d23.ganador.sum()), 'Advertencia', 'En Esri el 2.º/3.º promedia 15 % (vs 34 % en 2019) y los 3 registros suman ~66 % de válidos: no se usan para márgenes; sólo cuenta el ganador.')
add('Registros Esri 2023 inconsistentes repuestos', 70, 48, 'Advertencia', '70 municipios con "Votos en blanco/nulos/no marcados" como ganador en Esri; 48 repuestos con Wikipedia.')
add('Preconteo vs escrutinio diferenciado (2023 marcado como corte no declarado)', len(d23), int(d23.estado_validacion.str.contains('corte no declarado|Wikipedia').sum()), 'Advertencia', 'La fuente 2023 no declara el corte: ningún resultado 2023 se presenta como definitivo (ver estado_validacion).')
add('Candidaturas 2027 vinculadas a territorio existente', len(PROS), PROS.cod_divipola.isin(codes | set(dep.departamento_id)).sum(), 'Error', 'cod_divipola de 5 dígitos (municipio) o 2 (departamento).')
add('Fuentes con fecha de consulta', len(FUE), FUE.fecha_consulta.notna().sum(), 'Advertencia', 'Todas consultadas al corte.')
add('Organizaciones de resultados presentes en catálogos (FK)', len(RA) + len(RG), int(pd.concat([RA.organizacion_id, RG.organizacion_id]).isin(set(PAR.organizacion_id) | set(LOC.organizacion_id)).sum()), 'Error', 'PARTIDOS + MOVIMIENTOS.')
add('Encuestas registradas', 1, 0, 'Pendiente', 'Sin encuestas cargadas.')
add('Resultados mesa a mesa', 1, 0, 'Pendiente', 'Estructura creada; carga pendiente.')
CHK = pd.DataFrame(chk)
ws = WS['CALIDAD DATOS']; header(ws, 'CALIDAD DE DATOS · Controles', f'Generado por el ETL el {FECHA_CORTE}. Nivel: Error (debe corregirse), Advertencia (limitación conocida), Pendiente (sin datos).', 10)
_, last, C = table(ws, CHK, 'tCALIDAD', widths={'control': 70, 'detalle': 90})
ws.conditional_formatting.add(f'{C["pct_ok"]}5:{C["pct_ok"]}{last}', ColorScaleRule(start_type='num', start_value=0, start_color='F8696B', mid_type='num', mid_value=0.8, mid_color='FFEB84', end_type='num', end_value=1, end_color='63BE7B'))
CHK.to_csv(os.path.join(CLEAN, 'calidad_controles.csv'), index=False, encoding='utf-8-sig')

# =============================================================== MAPA (PNG) para DASHBOARD
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, geopandas as gpd
from matplotlib.patches import Patch
COL = {'LIBERAL': '#C0392B', 'CONSERVADOR': '#2471A3', 'LA_U': '#F39C12', 'CAMBIO_RADICAL': '#5DADE2', 'CENTRO_DEMOCRATICO': '#1B4F72', 'ALIANZA_VERDE': '#27AE60',
       'PACTO_HISTORICO': '#8E44AD', 'COLOMBIA_HUMANA_UP': '#A569BD', 'POLO': '#F4D03F', 'MAIS': '#A0522D', 'AICO': '#D35400', 'ASI': '#16A085', 'COALICION': '#7F8C8D', 'FUERZA_DE_LA_PAZ': '#E91E63', 'CREEMOS': '#00BCD4',
       'EN_MARCHA': '#CDDC39', 'DEMOCRATA': '#795548', 'DIGNIDAD': '#FF7043', 'ADA': '#9C27B0', 'COLOMBIA_RENACIENTE': '#3F51B5', 'OTRA_LOCAL': '#D5D8DC'}
gm = gpd.read_file(os.path.join(OUT, 'MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson'))
gm = gm.merge(mun[['codigo_divipola', 'fuerza_gobernante']], left_on='mpio', right_on='codigo_divipola', how='left')
gm['c'] = gm.fuerza_gobernante.map(lambda f: COL.get(f, '#D5D8DC' if isinstance(f, str) else '#FFFFFF'))
fig, ax = plt.subplots(figsize=(8.4, 7.4), dpi=110)
gm.plot(color=gm.c, ax=ax, edgecolor='#BBBBBB', linewidth=0.06); ax.set_axis_off(); ax.set_xlim(-80, -66.8); ax.set_ylim(-4.3, 12.6)
top = mun.fuerza_gobernante.value_counts()
lab = lambda k: {'OTRA_LOCAL': 'Organización local', 'COALICION': 'Coalición', 'LA_U': 'La U', 'GSC': 'GSC'}.get(k, k.replace('_', ' ').title())
hand = [Patch(color=COL.get(k, '#D5D8DC'), label=f'{lab(k)} ({v})') for k, v in top.head(16).items()] + [Patch(facecolor='white', edgecolor='#BBBBBB', label='Sin dato / no elige')]
ax.legend(handles=hand, loc='upper left', bbox_to_anchor=(1.0, 0.95), fontsize=7, frameon=False, title='Fuerza ganadora 2023', title_fontsize=8)
ax.set_title('Alcaldías 2024-2027 por fuerza ganadora', fontsize=10, color='#1F3864', loc='left')
fig.text(0.01, 0.01, 'Fuente: Esri/Registraduría y Wikipedia; geometría MGN simplificada. Colores = convención analítica de Atlas.', fontsize=5.5, color='#595959')
png = os.path.join(CLEAN, 'mapa_fuerza_2023.png'); fig.savefig(png, bbox_inches='tight'); plt.close(fig)

# =============================================================== DASHBOARD
ws = WS['DASHBOARD']; header(ws, 'DASHBOARD · Atlas Colombia 2027', f'Cobertura territorial, poder político y estado de la base. Corte {FECHA_CORTE}. Elecciones: 31-oct-2027.', 16)
K = [('Entidades DIVIPOLA', "=COUNTA(tMUNICIPIOS[codigo_divipola])"), ('Municipios que eligen alcalde', '=COUNTIF(tMUNICIPIOS[elige_alcalde],"Sí")'),
     ('Alcaldes 2023 identificados', '=COUNTA(tALCALDES[alcalde_id])'), ('Gobernaciones', '=COUNTA(tGOBERNADORES[gobernador_id])'),
     ('Resultados alcaldía (filas)', '=COUNTA(tDATOSVOTO[resultado_id])'), ('Resultados gobernación (filas)', '=COUNTA(tDATOSVOTOGOB[resultado_id])'),
     ('Aspirantes 2027 registrados', '=COUNTA(tCANDIDATOS2027[candidato_id])'), ('Candidaturas inscritas 2027', '=COUNTIF(tCANDIDATOS2027[estado_candidatura],"Candidato inscrito")+COUNTIF(tCANDIDATOS2027[estado_candidatura],"Candidato en firme")'),
     ('Prospectos (actores)', '=COUNTA(tPROSPECTOS[prospecto_id])'), ('Encuestas', '=COUNTA(tENCUESTAS[encuesta_id])'),
     ('Días a la elección', '=DATE(2027,10,31)-TODAY()'), ('Controles de calidad con error', '=COUNTIFS(tCALIDAD[nivel],"Error",tCALIDAD[pendientes_o_errores],">0")')]
for i, (lab, f) in enumerate(K):
    r, c = 4 + (i // 4) * 3, 1 + (i % 4) * 3
    ws.cell(r, c, lab).font = Font(name=FONT, size=9, color='595959')
    v = ws.cell(r + 1, c, f); v.font = Font(name=FONT, size=20, bold=True, color=NAVY); v.number_format = '#,##0'
    for cc in range(c, c + 2):
        for rr in (r, r + 1): ws.cell(rr, cc).fill = PatternFill('solid', fgColor=GREY)
# tabla de alcaldías 2023 por fuerza (fórmulas)
r0 = 14; ws.cell(r0 - 1, 1, 'Alcaldías 2024-2027 por fuerza ganadora').font = Font(name=FONT, bold=True, color=NAVY, size=11)
fz = mun.fuerza_gobernante.value_counts().index.tolist()
ws.cell(r0, 1, 'fuerza'); ws.cell(r0, 2, 'alcaldías'); ws.cell(r0, 3, '% del total')
for i, f in enumerate(fz, start=r0 + 1):
    ws.cell(i, 1, f); ws.cell(i, 2, f'=COUNTIF(tMUNICIPIOS[fuerza_gobernante],A{i})'); ws.cell(i, 3, f'=B{i}/SUM($B${r0 + 1}:$B${r0 + len(fz)})'); ws.cell(i, 3).number_format = '0.0%'
t = Table(displayName='tDASH_FUERZA', ref=f'A{r0}:C{r0 + len(fz)}'); t.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True); ws.add_table(t)
ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Alcaldías por fuerza (2023)'; ch.legend = None
ch.add_data(Reference(ws, min_col=2, min_row=r0, max_row=r0 + len(fz)), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + len(fz)))
ch.height, ch.width = 9, 13; ch.y_axis.majorGridlines = None
ws.add_chart(ch, 'E13')
# participación media por año
r1 = r0 + len(fz) + 3; ws.cell(r1 - 1, 1, 'Participación media en alcaldía (municipios con censo)').font = Font(name=FONT, bold=True, color=NAVY, size=11)
ws.cell(r1, 1, 'anio'); ws.cell(r1, 2, 'participacion_media')
for i, y in enumerate([2007, 2011, 2015, 2019], start=r1 + 1):
    ws.cell(i, 1, y); ws.cell(i, 2, f'=AVERAGEIFS(tPARTICIPACION[participacion],tPARTICIPACION[anio],A{i},tPARTICIPACION[dignidad],"ALC")'); ws.cell(i, 2).number_format = '0.0%'
lc = LineChart(); lc.title = 'Participación media – alcaldía'; lc.style = 12; lc.legend = None
lc.add_data(Reference(ws, min_col=2, min_row=r1, max_row=r1 + 4), titles_from_data=True); lc.set_categories(Reference(ws, min_col=1, min_row=r1 + 1, max_row=r1 + 4))
lc.height, lc.width = 7, 13; lc.y_axis.number_format = '0%'
ws.add_chart(lc, f'E{r1 - 1}')
# calidad resumida
r2 = r1 + 7; ws.cell(r2 - 1, 1, 'Estado de calidad').font = Font(name=FONT, bold=True, color=NAVY, size=11)
for i, nv in enumerate(['Error', 'Advertencia', 'Pendiente'], start=r2):
    ws.cell(i, 1, nv); ws.cell(i, 2, f'=COUNTIFS(tCALIDAD[nivel],A{i},tCALIDAD[pendientes_o_errores],">0")')
img = XLImage(png); img.width, img.height = img.width * 0.75, img.height * 0.75
ws.add_image(img, 'M4')
for c in range(1, 12): ws.column_dimensions[L(c)].width = 14
ws.column_dimensions['A'].width = 26

import geopandas as gpd
# =============================================================== GEO FICHA (coordenadas para el mapa dinámico)
from openpyxl.chart import ScatterChart, Series
from openpyxl.chart.marker import Marker
gf_ = gpd.read_file(os.path.join(OUT, 'MUNICIPIOS_COLOMBIA_MGN_simplificado.geojson'))[['mpio', 'geometry']]
gf_['geometry'] = gf_.geometry.make_valid().buffer(0)
BBX = (-79.1, -4.3, -66.8, 12.6)
def _ext(geom, tol):
    geom = geom.simplify(tol, preserve_topology=True)
    ps = [p for p in (geom.geoms if hasattr(geom, 'geoms') else [geom]) if p.geom_type == 'Polygon']
    return [list(p.exterior.coords) for p in ps]
deps_ = gf_.assign(d=gf_.mpio.str[:2]).dissolve(by='d')
DROWS = []
for gm_ in deps_.geometry:
    for ring in _ext(gm_.buffer(0.003).buffer(-0.003), 0.03):
        if min(x for x, y in ring) < BBX[0] or len(ring) < 4: continue
        DROWS += [(round(x, 4), round(y, 4)) for x, y in ring] + [(None, None)]
MROWS = []
for cod, gm_ in zip(gf_.mpio, gf_.geometry):
    rings_ = sorted(_ext(gm_, 0.006), key=len, reverse=True)[:3]
    for ring in rings_:
        if min(x for x, y in ring) < BBX[0]: continue
        MROWS += [(cod, round(x, 4), round(y, 4)) for x, y in ring] + [(cod, None, None)]
ws = WS['GEO FICHA']; header(ws, 'GEO FICHA · Coordenadas del mapa de la FICHA', 'Hoja auxiliar: contornos departamentales (columnas A-B) y municipales (D-H). Las columnas G-H muestran sólo el municipio elegido en FICHA.', 10)
ws['A4'], ws['B4'] = 'lon_depto', 'lat_depto'
for i, (x, y) in enumerate(DROWS, start=5): ws.cell(i, 1, x); ws.cell(i, 2, y)
ND = 4 + len(DROWS)
for j, h in enumerate(['codigo_divipola', 'lon', 'lat', 'lon_sel', 'lat_sel'], start=4): ws.cell(4, j, h)
for i, (c_, x, y) in enumerate(MROWS, start=5):
    ws.cell(i, 4, c_); ws.cell(i, 5, x); ws.cell(i, 6, y)
    ws.cell(i, 7, f'=IF(AND($D{i}=FICHA!$C$5,E{i}<>""),E{i},NA())'); ws.cell(i, 8, f'=IF(AND($D{i}=FICHA!$C$5,F{i}<>""),F{i},NA())')
NM = 4 + len(MROWS)
ws['J4'], ws['K4'] = 'lon_centro', 'lat_centro'
ws['J5'] = '=IFERROR(INDEX(tGEO[lon],MATCH(FICHA!$C$5,tGEO[codigo_divipola],0)),NA())'
ws['K5'] = '=IFERROR(INDEX(tGEO[lat],MATCH(FICHA!$C$5,tGEO[codigo_divipola],0)),NA())'
for c_ in 'ABDEFGHJK': ws.column_dimensions[c_].width = 14
ws.sheet_state = 'visible'

# =============================================================== FICHA
ws = WS['FICHA']; header(ws, 'FICHA · Municipio', 'Elegí un municipio de la lista (celda amarilla). Todo se calcula con fórmulas sobre las tablas de la base.', 12)
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 30; ws.column_dimensions['C'].width = 58
for c_, w_ in zip('DEFGHI', [16, 30, 14, 12, 12, 10]): ws.column_dimensions[c_].width = w_
ws['B4'] = 'ELEGÍ UN MUNICIPIO ►'; ws['B4'].font = Font(name=FONT, bold=True, color=NAVY, size=11)
c = ws['C4']; c.value = 'MEDELLÍN · ANTIOQUIA'; c.fill = PatternFill('solid', fgColor='FFF2CC'); c.font = Font(name=FONT, size=15, bold=True, color=NAVY)
c.border = Border(left=Side(style='medium', color=ORANGE), right=Side(style='medium', color=ORANGE), top=Side(style='medium', color=ORANGE), bottom=Side(style='medium', color=ORANGE))
ws.row_dimensions[4].height = 26
dvf = DataValidation(type='list', formula1=f"MUNICIPIOS!${MUN_CLAVE_COL}$5:${MUN_CLAVE_COL}${MUN_LAST}", allow_blank=False)
dvf.error = 'Elegí un municipio de la lista (formato: NOMBRE · DEPARTAMENTO)'; dvf.showErrorMessage = True; ws.add_data_validation(dvf); dvf.add('C4')
ws['D4'] = '◄ desplegá la lista o escribí NOMBRE · DEPARTAMENTO'; ws['D4'].font = Font(name=FONT, italic=True, size=9, color='595959')
ws['B5'] = 'Código DIVIPOLA'; ws['C5'] = '=IFERROR(INDEX(tMUNICIPIOS[codigo_divipola],MATCH($C$4,tMUNICIPIOS[municipio_clave],0)),"")'
def lk(tbl, col, key='$C$5', keycol='codigo_divipola', fmt_=None):
    x = f'INDEX({tbl}[{col}],MATCH({key},{tbl}[{keycol}],0))'
    return f'=IFERROR(IF({x}="","",{x}),"")'
DEPK = 'LEFT($C$5,2)'
campos = [
    ('DATOS GENERALES', None, None),
    ('Departamento', lk('tMUNICIPIOS', 'nombre_departamento'), None), ('Tipo de entidad', lk('tMUNICIPIOS', 'tipo_entidad_territorial'), None),
    ('Distrito especial', lk('tMUNICIPIOS', 'es_distrito_especial'), None),
    ('Población 2026 (DANE)', lk('tMUNICIPIOS', 'pob_tot_2026'), '#,##0'), ('Población rural 2026', lk('tTERRITORIO', 'pct_rural_2026'), '0.0%'),
    ('Densidad 2026 (hab/km², aprox.)', lk('tMUNICIPIOS', 'densidad_2026'), '#,##0.0'),
    ('Pobreza multidimensional (IPM 2018, %)', lk('tTERRITORIO', 'ipm_2018'), '0.0'), ('NBI 2018 (%)', lk('tTERRITORIO', 'nbi_2018'), '0.0'),
    ('Grupo de capacidades (MDM 2020)', lk('tMUNICIPIOS', 'mdm_grupo_2020'), None), ('Puntaje MDM 2020', lk('tTERRITORIO', 'mdm_puntaje_2020'), '0.0'),
    ('Censo electoral 2019', lk('tMUNICIPIOS', 'censo_electoral_2019'), '#,##0'), ('Participación alcaldía 2019', lk('tMUNICIPIOS', 'participacion_alc_2019'), '0.0%'),
    ('PODER LOCAL', None, None),
    ('Alcalde 2024-2027', lk('tMUNICIPIOS', 'alcalde_actual'), None), ('Organización', lk('tMUNICIPIOS', 'partido_gobernante'), None),
    ('Tipo / fuerza', f'=IFERROR(INDEX(tMUNICIPIOS[tipo_org_gobernante],MATCH($C$5,tMUNICIPIOS[codigo_divipola],0))&" · "&INDEX(tMUNICIPIOS[fuerza_gobernante],MATCH($C$5,tMUNICIPIOS[codigo_divipola],0)),"")', None),
    ('Integrantes de la coalición', lk('tALCALDES', 'integrantes_coalicion'), None),
    ('Votos del ganador 2023', lk('tMUNICIPIOS', 'votos_ganador_2023'), '#,##0'), ('% del ganador 2023', lk('tMUNICIPIOS', 'pct_ganador_2023'), '0.0%'),
    ('Antecedentes del alcalde', lk('tALCALDES', 'antecedentes_electorales'), None),
    ('Gobernador del departamento', lk('tGOBERNADORES', 'nombre_completo', DEPK, 'departamento_id'), None),
    ('Organización del gobernador', lk('tGOBERNADORES', 'partido_o_coalicion', DEPK, 'departamento_id'), None),
    ('LECTURA ELECTORAL', None, None),
    ('Fuerza y margen 2019', f'=IFERROR(INDEX(tSWING[fuerza_2019],MATCH($C$5,tSWING[codigo_divipola],0))&" · "&FIXED(INDEX(tSWING[margen_2019],MATCH($C$5,tSWING[codigo_divipola],0)),1)&" pts","")', None),
    ('Clasificación (SWING)', lk('tSWING', 'clasificacion'), None), ('Bastión', lk('tBASTIONES', 'clasificacion_bastion'), None),
    ('Fuerza dominante 2007-2023', lk('tBASTIONES', 'fuerza_dominante'), None),
    ('P(la misma fuerza retiene) · modelo v0.1', lk('tPRONOSTICO', 'probabilidad_retencion_fuerza'), '0%'),
    ('Actores 2027 en prensa', lk('tMUNICIPIOS', 'actores_2027_identificados'), None),
]
r_ = 7
for lab_, f_, nf in campos:
    if f_ is None:
        ws.cell(r_, 2, lab_).font = Font(name=FONT, bold=True, color=WHITE, size=9)
        for cc in (2, 3): ws.cell(r_, cc).fill = PatternFill('solid', fgColor=TEAL)
    else:
        a_ = ws.cell(r_, 2, lab_); a_.font = Font(name=FONT, bold=True, size=9); a_.fill = PatternFill('solid', fgColor=GREY); a_.alignment = Alignment(vertical='top', wrap_text=True)
        v_ = ws.cell(r_, 3, f_); v_.font = Font(name=FONT, size=10); v_.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        if nf: v_.number_format = nf
        a_.border = v_.border = Border(bottom=thin)
    r_ += 1
for rr in (27,):  # antecedentes puede ser largo
    pass
ws['B5'].font = Font(name=FONT, size=9, color='595959'); ws['C5'].font = Font(name=FONT, size=9, color='595959'); ws['C5'].alignment = Alignment(horizontal='left')
# enlaces
r_ += 1
ws.cell(r_, 2, 'Enlaces').font = Font(name=FONT, bold=True, color=NAVY)
links = ['=HYPERLINK("#\'MUNICIPIOS\'!C"&(MATCH($C$4,tMUNICIPIOS[municipio_clave],0)+4),"► Ir a su fila en MUNICIPIOS")',
         '=HYPERLINK("#\'TERRITORIO\'!C"&(MATCH($C$5,tTERRITORIO[codigo_divipola],0)+4),"► Ver su fila en TERRITORIO")',
         '=HYPERLINK("https://news.google.com/search?q=%22"&SUBSTITUTE(INDEX(tMUNICIPIOS[nombre_municipio],MATCH($C$5,tMUNICIPIOS[codigo_divipola],0))," ","+")&"%22+alcald%C3%ADa+2027&hl=es-419&gl=CO","► Buscar noticias en Google News")',
         '=HYPERLINK("https://www.google.com/maps/search/"&SUBSTITUTE($C$4," ","+")&"+Colombia","► Ver en Google Maps")',
         f'=HYPERLINK("{MAP_URL}/ir/"&$C$5&".html","► Abrir en el MAPA INTERACTIVO en línea (sectores censales y puestos)")']
for k, f_ in enumerate(links):
    c = ws.cell(r_ + k, 3, f_); c.font = Font(name=FONT, color='0563C1', underline='single', size=9)
r_ += len(links) + 1

def section(r0, title, heads, widths_=None):
    ws.cell(r0, 2, title).font = Font(name=FONT, bold=True, size=12, color=NAVY)
    for j, h in enumerate(heads, start=2):
        cc = ws.cell(r0 + 1, j, h); cc.font = Font(name=FONT, bold=True, color=WHITE, size=9); cc.fill = PatternFill('solid', fgColor=NAVY)
    return r0 + 2
# historia
r0 = section(r_, 'Historia de la alcaldía 2003-2023', ['Año', 'Ganador', 'Organización', 'Fuerza', '% ganador', 'Margen (pts)', 'Participación', 'NEP'])
for k, y in enumerate([2003, 2007, 2011, 2015, 2019, 2023]):
    rr = r0 + k; key = f'$C$5&"|{y}"'
    ws.cell(rr, 2, y).font = Font(name=FONT, bold=True, size=9)
    for j, col_ in enumerate(['ganador', 'organizacion_ganadora', 'fuerza_ganadora', 'pct_ganador', 'margen_pts', 'participacion', 'nep'], start=3):
        x = f'INDEX(tINDICADORES[{col_}],MATCH({key},tINDICADORES[clave],0))'
        c = ws.cell(rr, j, f'=IFERROR(IF({x}="","",{x}),"")'); c.font = Font(name=FONT, size=9)
        c.number_format = {'pct_ganador': '0.0%', 'participacion': '0.0%', 'margen_pts': '0.0', 'nep': '0.00'}.get(col_, 'General')
        c.alignment = Alignment(wrap_text=col_ in ('ganador', 'organizacion_ganadora'), vertical='top')
    if k % 2: 
        for j in range(2, 10): ws.cell(rr, j).fill = PatternFill('solid', fgColor='F8F8F8')
ws.cell(r0 + 6, 2, '2023: sólo el ganador es confiable en la fuente (sin margen ni NEP).').font = Font(name=FONT, italic=True, size=8, color='7F7F7F')
r_ = r0 + 8
# candidatos 2019
r0 = section(r_, 'Alcaldía 2019 · resultados por candidato (oficial, CEDAE)', ['Puesto', 'Candidato', 'Organización', 'Fuerza', 'Votos', '% válidos'])
for k in range(1, 9):
    rr = r0 + k - 1; key = f'$C$5&"|2019|{k}"'
    ws.cell(rr, 2, f'=IF(ISNUMBER(MATCH({key},tDATOSVOTO[clave],0)),{k},"")').font = Font(name=FONT, size=9)
    for j, col_ in enumerate(['nombre_candidato', 'organizacion', 'fuerza', 'votos', 'porcentaje_votos'], start=3):
        c = ws.cell(rr, j, f'=IFERROR(INDEX(tDATOSVOTO[{col_}],MATCH({key},tDATOSVOTO[clave],0)),"")'); c.font = Font(name=FONT, size=9)
        c.number_format = {'votos': '#,##0', 'porcentaje_votos': '0.0%'}.get(col_, 'General')
ws.conditional_formatting.add(f'B{r0}:G{r0}', FormulaRule(formula=[f'$B${r0}=1'], fill=PatternFill('solid', fgColor=GREEN)))
r_ = r0 + 9
# actores 2027
r0 = section(r_, 'Actores 2027 mencionados en prensa (no son candidaturas inscritas)', ['Ámbito', 'Nombre', 'Dignidad', 'Partido mencionado', 'Estado', 'Fuente'])
for k in range(1, 7):
    for amb, keyb in (('Municipio', '$C$5'), ('Departamento', DEPK)):
        pass
rr = r0
for amb, keyb, n_ in (('Municipio', '$C$5', 6), ('Departamento', DEPK, 6)):
    for k in range(1, n_ + 1):
        key = f'{keyb}&"|{k}"'
        ws.cell(rr, 2, f'=IF(ISNUMBER(MATCH({key},tPROSPECTOS[clave],0)),"{amb}","")').font = Font(name=FONT, size=9)
        for j, col_ in enumerate(['nombre', 'dignidad', 'partido_mencionado', 'estado', 'fuente_id'], start=3):
            x = f'INDEX(tPROSPECTOS[{col_}],MATCH({key},tPROSPECTOS[clave],0))'
            ws.cell(rr, j, f'=IFERROR(IF({x}="","",{x}),"")').font = Font(name=FONT, size=9)
        rr += 1
ws.cell(rr, 2, 'Para agregar actores, sumá filas en PROSPECTOS (la clave territorio|n se completa con el mismo formato).').font = Font(name=FONT, italic=True, size=8, color='7F7F7F')

# mapa dinámico
ch = ScatterChart(); ch.scatterStyle = 'line'; ch.legend = None; ch.title = 'Ubicación en Colombia'; ch.style = 2
def _ser(xc, yc, r1, r2, color, w, name):
    s_ = Series(Reference(WS['GEO FICHA'], min_col=yc, min_row=r1, max_row=r2), Reference(WS['GEO FICHA'], min_col=xc, min_row=r1, max_row=r2), title=name)
    s_.marker = Marker(symbol='none'); s_.smooth = False
    s_.graphicalProperties.line.solidFill = color; s_.graphicalProperties.line.width = w
    return s_
ch.series.append(_ser(1, 2, 5, ND, 'A6A6A6', 6000, 'Departamentos'))
ch.series.append(_ser(7, 8, 5, NM, 'ED7D31', 28000, 'Municipio'))
cen = Series(Reference(WS['GEO FICHA'], min_col=11, min_row=5, max_row=5), Reference(WS['GEO FICHA'], min_col=10, min_row=5, max_row=5), title='Centro')
cen.marker = Marker(symbol='circle', size=9); cen.marker.graphicalProperties.solidFill = 'ED7D31'; cen.marker.graphicalProperties.line.solidFill = 'FFFFFF'
cen.graphicalProperties.line.noFill = True
ch.series.append(cen)
ch.x_axis.scaling.min, ch.x_axis.scaling.max = BBX[0], BBX[2]; ch.y_axis.scaling.min, ch.y_axis.scaling.max = BBX[1], BBX[3]
ch.x_axis.delete = True; ch.y_axis.delete = True; ch.x_axis.majorGridlines = None; ch.y_axis.majorGridlines = None
ch.display_blanks = 'gap'
ch.height, ch.width = 15.5, 11.8
ws.add_chart(ch, 'K4')
ws.freeze_panes = 'A6'

# =============================================================== HOJAS MAPA (títulos y leyendas; las formas las inyecta 03_maps_xml.py)
from maps_common import enrich, maps_spec
MM = enrich(mun, BAS)
SPEC = maps_spec(MM, GOB)
LC = 23  # columna W: a la derecha del mapa
for name, sp in SPEC.items():
    ws = WS[name]; header(ws, sp['title'], sp['sub'], 30)
    ws.cell(4, LC, sp['ltitle']).font = Font(name=FONT, bold=True, size=11, color=NAVY)
    for i, (lb, colr) in enumerate(sp['legend'], start=5):
        c = ws.cell(i, LC); c.fill = PatternFill('solid', fgColor=colr.lstrip('#')); c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws.cell(i, LC + 1, lb).font = Font(name=FONT, size=9)
    ws.column_dimensions[L(LC)].width = 3; ws.column_dimensions[L(LC + 1)].width = 34
    nr = 5 + len(sp['legend']) + 1
    tip = ('Pasá el mouse sobre un departamento para ver gobernador y organización; clic: abre el MAPA INTERACTIVO en línea en ese departamento.' if name == 'MAPA GOBERNACIONES' else
           'Pasá el mouse sobre un municipio para ver alcalde, fuerza, margen 2019, swing, bastión, población e IPM. '
           'El recuadro de la derecha amplía la zona andina; arriba a la izquierda, San Andrés y Providencia. Clic en un municipio: abre el MAPA INTERACTIVO en línea (navegador) con su ficha, sus sectores censales y sus puestos de votación.')
    c = ws.cell(nr, LC, tip); c.font = Font(name=FONT, size=8, italic=True, color='595959'); c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=nr, start_column=LC, end_row=nr + 4, end_column=LC + 1)
    if name == 'MAPA GOBERNACIONES':
        g2 = GOB.sort_values('nombre_departamento')[['nombre_departamento', 'nombre_completo', 'partido_o_coalicion']].copy()
        g2['partido_o_coalicion'] = g2.partido_o_coalicion.str.split('(').str[0].str.strip()
        table(ws, g2.rename(columns={'nombre_completo': 'gobernador', 'partido_o_coalicion': 'organizacion'}), 'tMAPA_GOB', row=nr + 6, col=LC + 1,
              style='TableStyleLight9', widths={'nombre_departamento': 22, 'gobernador': 32, 'organizacion': 50}, freeze=False)

# =============================================================== LEEME
ws = WS['LEEME']; header(ws, 'LEEME · Diccionario, reglas y fórmulas', 'Descripción de cada columna por hoja, fórmulas de los indicadores y reglas de clasificación.', 6)
rows = []
TABMAP = {'MUNICIPIOS': mcols, 'ALCALDES': list(ALC.columns), 'GOBERNADORES': list(GOB.columns), 'DATOS VOTO': list(RA.columns), 'PARTICIPACION': list(P.columns),
          'TERRITORIO': tcols, 'INDICADORES': ic, 'BASTIONES': bc, 'PRONOSTICO': list(FC.columns), 'GEO': list(geo.columns), 'DEPARTAMENTOS': dcols}
for sh, cols in TABMAP.items():
    for c in cols: rows.append((sh, c, DICT.get(c, 'Ver encabezado de la hoja.'), fmt_for(c) or 'texto'))
DDF = pd.DataFrame(rows, columns=['hoja', 'columna', 'descripcion', 'formato'])
FORM = pd.DataFrame([
    ('Margen (pts)', '(votos 1.º − votos 2.º) / votos válidos × 100', 'INDICADORES, HISTORIA, SWING'),
    ('NEP (Laakso-Taagepera)', '1 / Σ pᵢ², pᵢ = votos del candidato i / votos a candidatos', 'INDICADORES'),
    ('Fragmentación de Rae', '1 − Σ pᵢ²', 'INDICADORES'), ('Volatilidad de Pedersen', '½ Σ |p_f,t − p_f,t−1| × 100, por fuerza (coaliciones/GSC agrupadas por tipo)', 'INDICADORES'),
    ('Participación', 'sufragantes / censo electoral; sufragantes = candidatos + blanco + nulos + no marcados', 'PARTICIPACION'),
    ('Competitividad (fórmula Excel)', 'Sin dato / Territorio fragmentado (NEP ≥ U_NEP_FRAGMENTADO) / Muy competitivo (< U_MUY_COMPETITIVO) / Competitivo / Moderado / Dominante', 'INDICADORES'),
    ('Clasificación swing (fórmula Excel)', 'Swing (cambió fuerza 19→23 y margen 2019 < U_SWING) → Alta incertidumbre (ganó coalición/GSC/local en 2023 y margen 2019 < U_COMPETITIVO) → Disputado (2015 y 2019 < U_COMPETITIVO) → Competitivo (2019 < U_MUY_COMPETITIVO) → Fragmentado → Estable', 'SWING'),
    ('Bastiones', 'Consolidado: misma fuerza 2015-2019-2023. Político: ≥3 de 5 (2007-2023) incl. 2023. En declive: ≥3 victorias pero perdió 2023', 'BASTIONES'),
    ('Modelo base v0.1', 'P(retención) = (retenciones_condición + 2·tasa_global) / (casos_condición + 2); condición = la fuerza repite o es primera victoria', 'PRONOSTICO'),
    ('Simulador', 'margen_simulado = margen_2019 − 2 × swing (si la fuerza coincide o "TODAS"); cambia si < 0', 'SIMULADOR'),
    ('Identificadores', 'municipio_id = MUN-DIVIPOLA; candidato_id = hash del nombre normalizado (orden de tokens indiferente); organizacion_id = ORG-partido o LOC-hash(nombre, territorio, año)', 'Todas'),
], columns=['concepto', 'formula_o_regla', 'hojas'])
ws['A3'] = ('Principios: no se inventan datos; coaliciones ≠ partidos; alcaldes ≠ candidatos; alcaldía, gobernación y corporaciones en tablas separadas; '
            'preconteo ≠ escrutinio; datos observados, indicadores y estimaciones separados; cada fila tiene fuente_id.')
ws['A3'].alignment = Alignment(wrap_text=True); ws.merge_cells('A3:F3'); ws.row_dimensions[3].height = 40
_, last, _ = table(ws, FORM, 'tFORMULAS', row=5, widths={'formula_o_regla': 100, 'concepto': 30, 'hojas': 30}, freeze=False)
table(ws, DDF, 'tDICCIONARIO', row=last + 3, widths={'descripcion': 100, 'columna': 30, 'hoja': 18}, freeze=False)

# =============================================================== INICIO
ws = WS['INICIO']
for c in range(1, 9):
    for r in range(1, 6): ws.cell(r, c).fill = PatternFill('solid', fgColor=NAVY)
ws['B2'] = 'ATLAS COLOMBIA'; ws['B2'].font = Font(name=FONT, size=26, bold=True, color=WHITE)
ws['B3'] = 'Base maestra de inteligencia electoral · Elecciones territoriales 31-oct-2027 (período 2028-2031)'; ws['B3'].font = Font(name=FONT, size=12, color=ORANGE, bold=True)
ws['B4'] = f'Municipios. Territorio. Poder. Inteligencia.   ·   Corte de datos: {FECHA_CORTE}'; ws['B4'].font = Font(name=FONT, size=10, italic=True, color=WHITE)
guide = ['CÓMO USAR ESTE ARCHIVO',
         '1. Empezá por DASHBOARD y las hojas MAPA (pasá el mouse sobre un municipio para ver sus datos; clic para ir a su fila). MUNICIPIOS es el catálogo maestro.',
         '2. Resultados: DATOS VOTO (alcaldía) y DATOS VOTO GOB (gobernación) por candidato. HISTORIA resume ganador y margen 2003-2023.',
         '3. Análisis: INDICADORES, BASTIONES, SWING. Los umbrales están en LISTAS (celdas amarillas) y recalculan las fórmulas.',
         '4. SIMULADOR: elegí fuerza y swing; PRONOSTICO guarda estimaciones del modelo (nunca resultados observados).',
         '5. 2027: CANDIDATOS 2027 (aspiraciones con estado) y PROSPECTOS (nombres en prensa). Agregá filas al final de cada tabla; los estados se eligen de listas.',
         '6. Cada fila tiene fuente_id (ver FUENTES). Las limitaciones conocidas están en CALIDAD DATOS.',
         'LIMITACIONES PRINCIPALES',
         '• 2023 alcaldía: ganador por municipio (Esri/Registraduría y Wikipedia). El 2.º y 3.º de esa fuente no son confiables: no hay márgenes 2023; SWING y SIMULADOR usan 2019. Pendiente: escrutinio oficial 2023.',
         '• 2003-2019: oficial (CEDAE-Registraduría). El nombre del partido no viene en la fuente: se infirió por código; el resto figura como "SIN NOMBRE EN FUENTE".',
         '• Concejos, asambleas, JAL y mesa a mesa: estructura prevista, sin datos cargados. Encuestas: ninguna verificada al corte.',
         '• Ningún aspirante está inscrito todavía (inscripción: 30-jun al 31-jul-2027).']
for i, g_ in enumerate(guide, start=7):
    c = ws.cell(i, 2, g_); c.font = Font(name=FONT, size=10, bold=g_.isupper(), color=NAVY if g_.isupper() else '262626')
r = 7 + len(guide) + 1
ws.cell(r, 2, 'NAVEGACIÓN').font = Font(name=FONT, size=11, bold=True, color=NAVY)
for i, s in enumerate(SHEETS[1:], start=1):
    rr, cc = r + 1 + (i - 1) % 15, 2 + ((i - 1) // 15) * 3
    c = ws.cell(rr, cc, s); c.hyperlink = f"#'{s}'!A1"; c.font = Font(name=FONT, size=10, bold=True, color='0563C1', underline='single')
    ws.cell(rr, cc + 1, DESC[s]).font = Font(name=FONT, size=9, color='595959')
ws.column_dimensions['A'].width = 3; ws.column_dimensions['B'].width = 26; ws.column_dimensions['C'].width = 48; ws.column_dimensions['E'].width = 26; ws.column_dimensions['F'].width = 48

# Fuentes de letra base
for w_ in wb.worksheets:
    w_.sheet_view.zoomScale = 90
wb.active = 0
path = os.path.join(OUT, 'BASE_MAESTRA_ATLAS_COLOMBIA_2027.xlsx')
wb.save(path); print('OK', path, round(os.path.getsize(path) / 1e6, 1), 'MB')

# =============================================================== ENTREGABLES AUXILIARES
with pd.ExcelWriter(os.path.join(OUT, 'DATA_DICTIONARY_ATLAS_COLOMBIA.xlsx')) as xw:
    DDF.to_excel(xw, sheet_name='DICCIONARIO', index=False); FORM.to_excel(xw, sheet_name='FORMULAS', index=False)
    pd.DataFrame([(s, DESC[s]) for s in SHEETS], columns=['hoja', 'funcion']).to_excel(xw, sheet_name='HOJAS', index=False)
with pd.ExcelWriter(os.path.join(OUT, 'DATA_VALIDATION_REPORT.xlsx')) as xw:
    CHK.to_excel(xw, sheet_name='CONTROLES', index=False)
    pd.DataFrame(dict(resumen=['Controles', 'Con errores', 'Advertencias con pendientes', 'Pendientes (sin datos)', 'Fuentes verificadas', 'Última actualización', 'Cobertura alcaldes 2023'],
                      valor=[len(CHK), int(((CHK.nivel == 'Error') & (CHK.pendientes_o_errores > 0)).sum()), int(((CHK.nivel == 'Advertencia') & (CHK.pendientes_o_errores > 0)).sum()),
                             int((CHK.nivel == 'Pendiente').sum()), len(FUE), FECHA_CORTE, f'{CHK.pct_ok.iloc[0]:.1%}'])).to_excel(xw, sheet_name='RESUMEN', index=False)
    open(os.path.join(CLEAN, '_log_etl.txt'), encoding='utf-8').read() and pd.read_csv(os.path.join(CLEAN, '_log_etl.txt'), sep='\t', header=None, names=['log_etl']).to_excel(xw, sheet_name='LOG_ETL', index=False)
mun.drop(columns=[c for c in mun.columns if c.startswith(('alcalde', 'partido', 'tipo_org', 'fuerza', 'votos', 'pct_g', 'margen', 'actores', 'n_'))], errors='ignore').to_csv(os.path.join(OUT, 'CATALOGOS_DIVIPOLA.csv'), index=False, encoding='utf-8-sig')
pd.concat([PAR.assign(nivel_catalogo='Partido nacional'), LOC.rename(columns={'nombre_organizacion': 'nombre_en_fuente'}).assign(nivel_catalogo='Coalición/GSC/local')], ignore_index=True).to_csv(os.path.join(OUT, 'CATALOGOS_POLITICOS.csv'), index=False, encoding='utf-8-sig')
pd.concat([RA, RG], ignore_index=True).to_csv(os.path.join(OUT, 'RESULTADOS_HISTORICOS.csv'), index=False, encoding='utf-8-sig')
print('Entregables auxiliares OK')
