"""Parsea la página de Wikipedia 'Elecciones regionales de Colombia de 2023' (markdown vía Firecrawl).
Salida: raw/gob2023_wiki.csv (todos los candidatos a gobernación) y raw/alc2023_wiki.csv (alcaldes electos con partidos de la coalición)."""
import re, csv, sys, os
RAW = sys.argv[1] if len(sys.argv) > 1 else 'raw'
m = open(os.path.join(RAW, 'wiki_reg2023.md'), encoding='utf-8').read()

NOTE = re.compile(r'\\[\d+\\]')
def unlink(s):
    """[texto](url "título") -> texto, respetando paréntesis anidados en la URL."""
    out, i = [], 0
    while i < len(s):
        if s[i] == '[':
            j = s.find('](', i)
            if j != -1 and '\n' not in s[i:j]:
                k, depth = j + 2, 1
                while k < len(s) and depth:
                    depth += {'(': 1, ')': -1}.get(s[k], 0); k += 1
                txt = s[i + 1:j]
                out.append('' if NOTE.fullmatch(txt) else txt); i = k; continue
        out.append(s[i]); i += 1
    return ''.join(out)
def clean(s):
    s = re.sub(r'!\[[^\]]*\]\([^)]*\)', '', s)
    s = unlink(s)
    s = NOTE.sub('', s)
    s = re.sub(r'\\?\[\d+\\?\]', '', s).replace('\\', '')
    s = s.replace('**', '').replace('<br>', ' | ')
    s = re.sub(r'\s+', ' ', s).strip(' ,|')
    return s
def num(s):
    s = clean(s).replace('.', '').replace(' ', '')
    return int(s) if s.isdigit() else None

g0, g1 = m.index('## Resultados a Gobernador'), m.index('## Alcaldes electos')
a1 = m.index('## Asambleas Departamentales')
secs = lambda txt, off: [(mm.start() + off, clean(mm.group(1))) for mm in re.finditer(r'^###\s+(.*)$', txt, re.M)]

gob = []
gs = secs(m[g0:g1], g0) + [(g1, None)]
for (p, dep), (q, _) in zip(gs, gs[1:]):
    body = m[p:q].split('\n')
    pos = 0
    for i, ln in enumerate(body):
        cells = [c.strip() for c in ln.split('|')]
        # fila candidato: |  | Nombre | Partido | Votos | |
        if ln.startswith('|  | ') and len(cells) >= 6 and num(cells[4]) is not None:
            pct = None
            for ln2 in body[i+1:i+4]:
                mm = re.search(r'(\d+[.,]\d+)%', ln2)
                if mm: pct = float(mm.group(1).replace(',', '.')); break
            pos += 1
            gob.append(dict(departamento=dep, tipo='CANDIDATO', candidato=clean(cells[2]), partido=clean(cells[3]),
                            votos=num(cells[4]), pct_wiki=pct, posicion=pos, electo=int('**' in cells[2])))
        elif ln.startswith('|  | ') and len(cells) >= 5 and num(cells[3]) is not None and 'blanco' in ln.lower():
            gob.append(dict(departamento=dep, tipo='BLANCO', candidato='VOTOS EN BLANCO', partido='', votos=num(cells[3])))
        elif re.match(r'^\| (\*\*)?Total de votos válidos', ln):
            gob.append(dict(departamento=dep, tipo='VALIDOS', candidato='TOTAL VÁLIDOS', partido='', votos=num(cells[2])))
        elif re.match(r'^\| \[?Votos nulos', ln):
            gob.append(dict(departamento=dep, tipo='NULOS', candidato='VOTOS NULOS', partido='', votos=num(cells[2])))
        elif re.match(r'^\| Tarjetas no marcadas', ln):
            gob.append(dict(departamento=dep, tipo='NO_MARCADOS', candidato='TARJETAS NO MARCADAS', partido='', votos=num(cells[2])))
        elif re.match(r'^\| (\*\*)?Total de votos escrutados', ln):
            gob.append(dict(departamento=dep, tipo='ESCRUTADOS', candidato='TOTAL ESCRUTADOS', partido='', votos=num(cells[2])))

alc = []
as_ = secs(m[g1:a1], g1) + [(a1, None)]
for (p, dep), (q, _) in zip(as_, as_[1:]):
    for ln in m[p:q].split('\n'):
        if not ln.startswith('| ') or ln.startswith('| :') or 'Municipio' in ln[:30]: continue
        cells = [c.strip() for c in ln.split('|')][1:-1]
        if len(cells) < 5: continue
        raw_party = cells[2]
        parts = [clean(x) for x in raw_party.split('<br>') if clean(x)]
        pct = re.search(r'(\d+[.,]\d+)', cells[4])
        alc.append(dict(departamento=re.sub(r'\s*\[.*$', '', dep), municipio=clean(cells[0]), alcalde=clean(cells[1]),
                        organizacion=parts[0] if parts else '', integrantes=' ; '.join(parts[1:]),
                        votos=num(cells[3]), pct_wiki=float(pct.group(1).replace(',', '.')) if pct else None))

def dump(rows, f):
    keys = sorted({k for r in rows for k in r}, key=lambda k: list(rows[0]).index(k) if k in rows[0] else 99)
    with open(os.path.join(RAW, f), 'w', newline='', encoding='utf-8') as fh:
        w = csv.DictWriter(fh, keys); w.writeheader(); w.writerows(rows)
dump(gob, 'gob2023_wiki.csv'); dump(alc, 'alc2023_wiki.csv')
print('gobernación filas', len(gob), 'deptos', len({r['departamento'] for r in gob}), 'electos', sum(r.get('electo', 0) for r in gob))
print('alcaldes', len(alc))
