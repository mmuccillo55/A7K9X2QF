"""
parse.py — RACKS_C12.xlsx → connections.json

Reglas:
- Cada hoja de equipo central es la fuente de verdad de SUS conexiones.
- add() rechaza cables duplicados (incluyendo el cable visto desde el otro extremo).
- No hay normalización de nombres: se usan los nombres exactos del xlsx.
- La hoja DB se ignora (es solo referencia).
"""

import pandas as pd, json, sys, os

XLSX = 'RACKS_C12.xlsx'
if not os.path.exists(XLSX):
    print(f"ERROR: no se encontró {XLSX}", file=sys.stderr); sys.exit(1)

sheets = pd.read_excel(XLSX, sheet_name=None, header=None)
sheets.pop('DB', None)

# ── helpers ──────────────────────────────────────────────────────────────────
def clean(v):
    if pd.isna(v): return None
    s = str(v).strip()
    return s if s not in ('', 'nan', 'NaN', '0', '0.0') else None

def sg(row, idx):
    try: return clean(row[idx])
    except (KeyError, IndexError): return None

SKIP = {'?', 'N/C', 'VACANTE', 'N/A'}
def skip(v): return not v or v in SKIP

# Variantes reales que aún existen en el xlsx con distinta capitalización
_EQ_NORM = {
    'blackmagic multiview 16': 'Blackmagic MultiView 16',
}
def norm_eq(eq):
    if not eq: return eq
    return _EQ_NORM.get(eq.lower(), eq)

_seen = set()
connections = []

def add(src_rack, src_eq, src_port, dst_rack, dst_eq, dst_port, label='', notes=''):
    src_eq = norm_eq(src_eq); dst_eq = norm_eq(dst_eq)
    if skip(src_rack) or skip(dst_rack): return
    if skip(src_eq)   or skip(dst_eq):  return
    src_port = str(src_port or '').strip()
    dst_port = str(dst_port or '').strip()
    key = (src_rack, src_eq, src_port, dst_rack, dst_eq, dst_port)
    rev = (dst_rack, dst_eq, dst_port, src_rack, src_eq, src_port)
    if key in _seen or rev in _seen: return
    _seen.add(key)
    connections.append({
        'src_rack': src_rack, 'src_eq': src_eq, 'src_port': src_port,
        'dst_rack': dst_rack, 'dst_eq': dst_eq, 'dst_port': dst_port,
        'label': str(label or '').strip(),
        'notes': str(notes or '').strip(),
    })

# ── Rack C.3 AJA KUMO 6464 ───────────────────────────────────────────────────
# Inputs:  0=in_n  1=rotulo  2=src_rack  3=src_eq  4=slot  5=src_port  6=notas
# Outputs: 9=out_n 10=rotulo 11=dst_rack 12=dst_eq 13=slot 14=dst_port 15=notas
df = sheets['Rack C.3 AJA KUMO 6464']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = sg(row, 0)
    ro, eo, po = sg(row, 2), sg(row, 3), sg(row, 5)
    lbl, nota = sg(row, 1), sg(row, 6)
    if in_n and not skip(ro) and ro != 'K':
        add(ro, eo, po or 'Out', 'C.3', 'AJA KUMO 6464', f'In {in_n}', lbl or eo or '', nota or '')

    out_n = sg(row, 9)
    rd, ed, pd_ = sg(row, 11), sg(row, 12), sg(row, 14)
    lbl2, nota2 = sg(row, 10), sg(row, 15)
    if out_n and not skip(rd):
        add('C.3', 'AJA KUMO 6464', f'Out {out_n}', rd, ed, pd_ or 'In', lbl2 or ed or '', nota2 or '')

# ── Rack C.4 Blackmagic MultiView ────────────────────────────────────────────
# Inputs:  0=in_n  1=src_rack  2=src_eq  3=src_port  4=notas
# Outputs: 7=out_n 8=dst_rack  9=dst_eq  11=dst_port 12=notas
df = sheets['Rack C.4 Blackmagic Multiview 1']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = sg(row, 0); ro, eo, po = sg(row, 1), sg(row, 2), sg(row, 3); nota = sg(row, 4)
    if in_n and not skip(ro):
        add(ro, eo, po or 'Out', 'C.4', 'Blackmagic MultiView 16', f'In {in_n}', '', nota or '')
    out_n = sg(row, 7); rd, ed, pd_ = sg(row, 8), sg(row, 9), sg(row, 11); nota2 = sg(row, 12)
    if out_n and not skip(rd):
        add('C.4', 'Blackmagic MultiView 16', f'Out {out_n}', rd, ed, pd_ or 'SDI In', '', nota2 or '')
    if i == 19:
        lbl = sg(row, 7); rd2, ed2, pd2 = sg(row, 8), sg(row, 9), sg(row, 11); nota3 = sg(row, 12)
        if lbl and not skip(rd2):
            add('C.4', 'Blackmagic MultiView 16', 'SDI OUT', rd2, ed2, pd2 or 'In', '', nota3 or '')

# ── Rack C.5 Blackmagic 12 x 12 ──────────────────────────────────────────────
# Inputs:  0=in_n  1=src_rack  2=src_eq  3=src_port  4=notas
# Outputs: 7=out_n 8=dst_rack  9=dst_eq  10=dst_port 11=notas
df = sheets['Rack C.5 Blackmagic 12 x 12']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = sg(row, 0); ro, eo, po = sg(row, 1), sg(row, 2), sg(row, 3); nota = sg(row, 4)
    if in_n and not skip(ro):
        add(ro, eo, po or 'Out', 'C.5', 'Blackmagic 12 x 12', f'In {in_n}', nota or '', '')
    out_n = sg(row, 7); rd, ed, pd_ = sg(row, 8), sg(row, 9), sg(row, 10); nota2 = sg(row, 11)
    if out_n and not skip(rd):
        add('C.5', 'Blackmagic 12 x 12', f'Out {out_n}', rd, ed, pd_ or 'In', '', nota2 or '')

# ── Rack C.6 Patch Panel Inputs ──────────────────────────────────────────────
# Cols: 0=patch_n  1=src_rack  2=src_eq  3=src_port  4=notas
df = sheets['Rack C.6 Patch panel Inputs']
for i in range(2, len(df)):
    row = df.iloc[i]
    pn = sg(row, 0); ro = sg(row, 1); eo = sg(row, 2); po = sg(row, 3); nota = sg(row, 4)
    if pn and not skip(ro):
        add(ro, eo or 'Ext', po or f'Out {pn}', 'C.6', 'Patch Panel', f'Port {pn}', '', nota or '')

# ── Rack C.7 Patch Panel Outputs ─────────────────────────────────────────────
# Cols: 0=patch_n  1=src_rack  2=src_eq  3=src_port  4=notas
df = sheets['Rack C.7 Patch panel Outputs']
for i in range(2, len(df)):
    row = df.iloc[i]
    pn = sg(row, 0); ro = sg(row, 1); eo = sg(row, 2); po = sg(row, 3); nota = sg(row, 4)
    if pn and not skip(ro) and eo:
        add(ro, eo, po or f'Out {pn}', 'C.7', 'Patch Panel', f'Port {pn}', nota or '', '')

# ── Rack D.Back ───────────────────────────────────────────────────────────────
# Cols: 0=src_rack 1=src_eq 2=slot 3=src_port 4=dst_rack 5=dst_eq 6=slot 7=dst_port 8=notas
df = sheets['Rack D.Back']
for i in range(2, len(df)):
    row = df.iloc[i]
    ro, eo, po = sg(row, 0), sg(row, 1), sg(row, 3)
    rd, ed, pd_ = sg(row, 4), sg(row, 5), sg(row, 7)
    nota = sg(row, 8)
    add(ro, eo, po or 'Out', rd, ed, pd_ or 'In', '', nota or '')

# ── Rack D.7 openGear X ──────────────────────────────────────────────────────
# Cols: 0=placa 1=slot_type 2=puerto_o 3=dst_rack 4=dst_eq 5=dst_slot 6=dst_port 7=notas
df = sheets['Rack D.7 openGear X']
cp = None; cs = None
for i in range(2, len(df)):
    row = df.iloc[i]
    p = sg(row, 0)
    if p:
        try: cp = int(float(p))
        except: cp = p
    s = sg(row, 1)
    if s: cs = s
    po = sg(row, 2); rd, ed, pd_ = sg(row, 3), sg(row, 4), sg(row, 6); nota = sg(row, 7)
    if po and not skip(rd) and rd != 'D.Back':
        sl = f'Slot {cp} ({cs})' if cs else f'Slot {cp}'
        add('D.7', f'openGear X {sl}', po, rd, ed, pd_ or 'In', '', nota or '')

# ── Rack D.9 openGear X ──────────────────────────────────────────────────────
df = sheets['Rack D.9 openGear X']
cp = None; cs = None
for i in range(2, len(df)):
    row = df.iloc[i]
    p = sg(row, 0)
    if p:
        try: cp = int(float(p))
        except: cp = p
    s = sg(row, 1)
    if s: cs = s
    po = sg(row, 2); rd, ed, pd_ = sg(row, 3), sg(row, 4), sg(row, 6); nota = sg(row, 7)
    if po and not skip(rd):
        sl = f'Slot {cp} ({cs})' if cs else f'Slot {cp}'
        add('D.9', f'openGear X {sl}', po, rd, ed, pd_ or 'In', '', nota or '')

# ── Rack F.7 FOR A HVS-390HS ─────────────────────────────────────────────────
# Inputs:  0=in_n  1=src_rack  2=src_eq  3=slot  4=src_port  5=notas
# Outputs: 7=out_n 8=dst_rack  9=dst_eq  10=dst_port 11=notas
df = sheets['Rack F.7 FOR A HVS-390HS']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = sg(row, 0); ro, eo, po = sg(row, 1), sg(row, 2), sg(row, 4); nota = sg(row, 5)
    if in_n and not skip(ro):
        add(ro, eo, po or 'Out', 'F.7', 'FOR A HVS-390HS', f'In {in_n}', '', nota or '')
    out_n = sg(row, 7); rd, ed, pd_ = sg(row, 8), sg(row, 9), sg(row, 10); nota2 = sg(row, 11)
    if out_n and not skip(rd):
        add('F.7', 'FOR A HVS-390HS', str(out_n), rd, ed, pd_ or 'In', '', nota2 or '')

# ── Output ───────────────────────────────────────────────────────────────────
with open('connections.json', 'w', encoding='utf-8') as f:
    json.dump(connections, f, ensure_ascii=False, indent=2)
print(f"OK: {len(connections)} conexiones → connections.json")