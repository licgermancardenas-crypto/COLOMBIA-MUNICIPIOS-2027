"""Clasificaciones y colores compartidos por las hojas MAPA (02_build_excel.py arma títulos y leyendas;
03_maps_xml.py dibuja las formas)."""
import numpy as np, pandas as pd

COL = {'LIBERAL': '#C0392B', 'CONSERVADOR': '#2471A3', 'LA_U': '#F39C12', 'CAMBIO_RADICAL': '#5DADE2', 'CENTRO_DEMOCRATICO': '#1B4F72',
       'ALIANZA_VERDE': '#27AE60', 'PACTO_HISTORICO': '#8E44AD', 'COLOMBIA_HUMANA_UP': '#A569BD', 'POLO': '#F4D03F', 'MAIS': '#A0522D',
       'AICO': '#D35400', 'ASI': '#16A085', 'COALICION': '#7F8C8D', 'FUERZA_DE_LA_PAZ': '#E91E63', 'CREEMOS': '#00BCD4', 'EN_MARCHA': '#CDDC39',
       'DEMOCRATA': '#795548', 'DIGNIDAD': '#FF7043', 'ADA': '#9C27B0', 'COLOMBIA_RENACIENTE': '#3F51B5', 'NUEVO_LIBERALISMO': '#E74C3C',
       'GENTE_EN_MOVIMIENTO': '#FFB300', 'FUERZA_CIUDADANA': '#00897B', 'INDEPENDIENTES': '#607D8B', 'ECOLOGISTA': '#8BC34A',
       'CJL': '#4DB6AC', 'SALVACION_NACIONAL': '#283593', 'LIGA': '#AFB42B', 'MIRA': '#1565C0', 'PRE': '#6D4C41',
       'OPCION_CIUDADANA': '#558B2F', 'COMUNES': '#B71C1C', 'DEFENSORES_PATRIA': '#0D47A1', 'VERDE_OXIGENO': '#9CCC65', 'GSC': '#BDBDBD', 'OTRA_LOCAL': '#D5D8DC'}
SIN = '#FFFFFF'
def lab(k):
    return {'OTRA_LOCAL': 'Organización local', 'COALICION': 'Coalición', 'LA_U': 'La U', 'GSC': 'GSC'}.get(k, k if k in ('MAIS', 'AICO', 'ASI', 'ADA', 'CJL', 'MIRA', 'PRE') else k.replace('_', ' ').title())

SW = {'Territorio swing': '#E67E22', 'Territorio disputado': '#C0392B', 'Territorio competitivo': '#F1C40F', 'Territorio de alta incertidumbre': '#8E44AD',
      'Territorio fragmentado': '#5DADE2', 'Estable': '#A9DFBF', 'Sin dato': SIN}
BA = {'Bastión consolidado': '#1E8449', 'Bastión político': '#82E0AA', 'Bastión en declive': '#F5B041', 'Sin bastión partidario': '#E5E7E9'}
QL = ['Q1 · IPM más bajo', 'Q2', 'Q3', 'Q4', 'Q5 · IPM más alto']
IPMC = dict(zip(QL, ['#FEF0D9', '#FDCC8A', '#FC8D59', '#E34A33', '#B30000']))
U = dict(muy=5, comp=10, nep=3.5, swing=10)   # umbrales por defecto de LISTAS


def enrich(mun, B):
    """Agrega a dim_municipios las clasificaciones que pintan los mapas."""
    m = mun.merge(B[['codigo_divipola', 'clasificacion_bastion', 'fuerza_dominante', 'margen_2015', 'margen_2019', 'nep_2019', 'cambio_2019_2023']],
                  on='codigo_divipola', how='left')
    def swing(r):
        m19, m15 = r.margen_2019, r.margen_2015
        if pd.isna(m19): return 'Sin dato'
        if r.cambio_2019_2023 == 'Sí' and m19 < U['swing']: return 'Territorio swing'
        if r.fuerza_gobernante in ('COALICION', 'GSC', 'OTRA_LOCAL') and m19 < U['comp']: return 'Territorio de alta incertidumbre'
        if pd.notna(m15) and m15 < U['comp'] and m19 < U['comp']: return 'Territorio disputado'
        if m19 < U['muy']: return 'Territorio competitivo'
        if pd.notna(r.nep_2019) and r.nep_2019 >= U['nep']: return 'Territorio fragmentado'
        return 'Estable'
    m['swing'] = m.apply(swing, axis=1)
    m['ipm_q'] = pd.qcut(m.ipm_2018, 5, labels=QL).astype(object)
    return m


def maps_spec(m, gob):
    """Para cada hoja MAPA: título, subtítulo, función de color y leyenda [(etiqueta, color)]."""
    cnt = lambda col, k: int((m[col] == k).sum())
    fz = m.fuerza_gobernante.value_counts()
    gf = gob.fuerza.value_counts()
    return {
        'MAPA POLÍTICO': dict(
            title='MAPA POLÍTICO · Quién gobierna cada municipio (2024-2027)',
            sub='Color = fuerza del alcalde electo el 29-oct-2023 (partido, coalición u organización local). Fuente: Esri/Registraduría y Wikipedia. Colores: convención analítica de Atlas.',
            color=lambda r: COL.get(r.fuerza_gobernante, '#D5D8DC') if isinstance(r.fuerza_gobernante, str) else SIN,
            legend=[(f'{lab(k)} ({v})', COL.get(k, '#D5D8DC')) for k, v in fz.items()] + [('Sin dato / no elige alcalde', SIN)], ltitle='Fuerza ganadora 2023'),
        'MAPA SWING': dict(
            title='MAPA SWING · Competitividad y cambios de fuerza',
            sub='Misma regla que la hoja SWING con los umbrales por defecto (5 / 10 / 3,5 / 10). Si cambiás los umbrales, la hoja SWING se recalcula; el mapa se regenera con el ETL.',
            color=lambda r: SW[r.swing], legend=[(f'{k} ({cnt("swing", k)})', c) for k, c in SW.items()], ltitle='Clasificación'),
        'MAPA BASTIONES': dict(
            title='MAPA BASTIONES · Fuerzas dominantes 2007-2023',
            sub='Consolidado: misma fuerza en 2015, 2019 y 2023. Político: ≥3 de 5 ciclos incl. 2023. En declive: ≥3 victorias pero perdió 2023. Las coaliciones no forman bastión partidario.',
            color=lambda r: BA.get(r.clasificacion_bastion, SIN) if isinstance(r.clasificacion_bastion, str) else SIN,
            legend=[(f'{k} ({cnt("clasificacion_bastion", k)})', c) for k, c in BA.items()] + [('Sin dato', SIN)], ltitle='Clasificación'),
        'MAPA TERRITORIO': dict(
            title='MAPA TERRITORIO · Pobreza multidimensional (IPM 2018)',
            sub='Quintiles del IPM municipal total (DANE, CNPV 2018). Q5 = mayor incidencia de pobreza multidimensional.',
            color=lambda r: IPMC.get(r.ipm_q, SIN) if isinstance(r.ipm_q, str) else SIN,
            legend=[(f'{k} ({cnt("ipm_q", k)})', c) for k, c in IPMC.items()] + [('Sin dato', SIN)], ltitle='IPM total 2018'),
        'MAPA GOBERNACIONES': dict(
            title='MAPA GOBERNACIONES · Quién gobierna cada departamento (2024-2027)',
            sub='Color = organización del gobernador electo en 2023. Las coaliciones no se atribuyen a un partido. Fuente: Wikipedia; contrastar con escrutinio.',
            color=None, legend=[(f'{lab(k)} ({v})', COL.get(k, '#D5D8DC')) for k, v in gf.items()] + [('Bogotá D.C. (alcalde mayor)', SIN)],
            ltitle='Organización del gobernador'),
    }
