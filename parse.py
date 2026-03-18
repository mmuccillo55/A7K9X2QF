import pandas as pd
import json
import sys
import os

XLSX = 'RACKS_C12.xlsx'

if not os.path.exists(XLSX):
    print(f"ERROR: no se encontró {XLSX}", file=sys.stderr)
    sys.exit(1)

all_sheets = pd.read_excel(XLSX, sheet_name=None, header=None)

def clean(v):
    if pd.isna(v): return None
    s = str(v).strip()
    return s if s not in ('', 'nan', 'NaN') else None

connections = []

def add(src_rack, src_eq, src_port, dst_rack, dst_eq, dst_port, label='', notes=''):
    if not src_rack or not dst_rack: return
    if src_rack in ('?', 'N/C') or dst_rack in ('?', 'N/C', 'VACANTE'): return
    connections.append({
        'src_rack': src_rack, 'src_eq': src_eq or '',  'src_port': src_port or '',
        'dst_rack': dst_rack, 'dst_eq': dst_eq or '',  'dst_port': dst_port or '',
        'label': label, 'notes': notes
    })

# ── Rack A ───────────────────────────────────────────────────────────────────
df = all_sheets['Rack A']
for i in range(2, len(df)):
    row = df.iloc[i]
    r_o, eq_o, p_o = clean(row[0]), clean(row[1]), clean(row[3])
    r_d, eq_d, p_d = clean(row[4]), clean(row[5]), clean(row[7])
    nota = clean(row[8])
    if r_o and r_d:
        add(r_o, eq_o, p_o, r_d, eq_d, p_d, eq_o or '', nota or '')

# ── Rack B ───────────────────────────────────────────────────────────────────
df = all_sheets['Rack B']
for i in range(2, len(df)):
    row = df.iloc[i]
    r_d, eq_d, p_d = clean(row[0]), clean(row[1]), clean(row[2])
    r_o, eq_o, p_o = clean(row[3]), clean(row[4]), clean(row[5])
    nota = clean(row[6])
    if r_o and r_d:
        add(r_o, eq_o, p_o, r_d, eq_d, p_d, eq_d or '', nota or '')

# ── Rack C.3 AJA KUMO 6464 ───────────────────────────────────────────────────
df = all_sheets['Rack C.3 AJA KUMO 6464']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n   = clean(row[0]);  rotulo = clean(row[1])
    r_o, eq_o, sl_o, p_o = clean(row[2]), clean(row[3]), clean(row[4]), clean(row[5])
    nota   = clean(row[6])
    if in_n and r_o and r_o not in ('?', 'N/C', 'K'):
        port_o = p_o or (f"Slot {sl_o}" if sl_o else 'Out')
        add(r_o, eq_o, port_o, 'C.3', 'AJA KUMO 6464', f'In {in_n}', rotulo or eq_o or '', nota or '')
    out_n  = clean(row[9]);  rotulo2 = clean(row[10])
    r_d, eq_d, sl_d, p_d = clean(row[11]), clean(row[12]), clean(row[13]), clean(row[14])
    nota2  = clean(row[15])
    if out_n and r_d and r_d not in ('?', 'VACANTE', 'N/C'):
        port_d = p_d or (f"Slot {sl_d}" if sl_d else 'In')
        add('C.3', 'AJA KUMO 6464', f'Out {out_n}', r_d, eq_d, port_d, rotulo2 or eq_d or '', nota2 or '')

# ── Rack C.4 Blackmagic MultiView ────────────────────────────────────────────
df = all_sheets['Rack C.4 Blackmagic Multiview 1']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = clean(row[0])
    r_o, eq_o, p_o, nota = clean(row[1]), clean(row[2]), clean(row[3]), clean(row[4])
    if in_n and r_o and r_o not in ('?',):
        add(r_o, eq_o, p_o or 'Out', 'C.4', 'Blackmagic MultiView 16', f'In {in_n}', '', nota or '')
    out_n = clean(row[7])
    r_d, eq_d, p_d, nota2 = clean(row[8]), clean(row[9]), clean(row[11]), clean(row[12])
    if out_n and r_d and r_d not in ('?', 'N/C'):
        add('C.4', 'Blackmagic MultiView 16', f'Out {out_n}', r_d, eq_d, p_d or 'SDI In', '', nota2 or '')
# SDI OUT fila especial (fila 19)
row = df.iloc[19]
r_d, eq_d, p_d, nota = clean(row[8]), clean(row[9]), clean(row[11]), clean(row[12])
if r_d:
    add('C.4', 'Blackmagic MultiView 16', 'SDI OUT', r_d, eq_d, p_d or 'In', '', nota or '')

# ── Rack C.5 Blackmagic 12x12 ────────────────────────────────────────────────
df = all_sheets['Rack C.5 Blackmagic 12 x 12']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = clean(row[0])
    r_o, eq_o, p_o, nota = clean(row[1]), clean(row[2]), clean(row[3]), clean(row[4])
    if in_n and r_o and r_o not in ('?', 'N/C'):
        add(r_o, eq_o, p_o or 'Out', 'C.5', 'Blackmagic 12x12', f'In {in_n}', nota or '', '')
    out_n = clean(row[7])
    r_d, eq_d, p_d, nota2 = clean(row[8]), clean(row[9]), clean(row[10]), clean(row[11])
    if out_n and r_d and r_d not in ('?', 'N/C'):
        add('C.5', 'Blackmagic 12x12', f'Out {out_n}', r_d, eq_d, p_d or 'In', '', nota2 or '')

# ── Rack C.6 Patch Panel Inputs ──────────────────────────────────────────────
df = all_sheets['Rack C.6 Patch panel Inputs']
for i in range(2, len(df)):
    row = df.iloc[i]
    patch_n = clean(row[0])
    r_o, eq_o, p_o, nota = clean(row[1]), clean(row[2]), clean(row[3]), clean(row[4])
    if patch_n and r_o and r_o not in ('N/C', 'C'):
        add(r_o, 'Patch Panel (Inputs)', f'Port {patch_n}', 'C.3', 'AJA KUMO 6464', p_o or f'In {patch_n}', '', nota or '')

# ── Rack C.7 Patch Panel Outputs ─────────────────────────────────────────────
df = all_sheets['Rack C.7 Patch panel Outputs']
for i in range(2, len(df)):
    row = df.iloc[i]
    patch_n = clean(row[0])
    r_o, eq_o, p_o, nota = clean(row[1]), clean(row[2]), clean(row[3]), clean(row[4])
    if patch_n and eq_o and r_o not in ('N/C',):
        add('C.3', 'AJA KUMO 6464', p_o or f'Out {patch_n}', 'C.7', 'Patch Panel (Outputs)', f'Port {patch_n}', nota or '', '')

# ── Rack D.7 openGear X ──────────────────────────────────────────────────────
df = all_sheets['Rack D.7 openGear X']
current_placa = None
current_slot_type = None
for i in range(2, len(df)):
    row = df.iloc[i]
    placa = clean(row[0])
    if placa:
        try: current_placa = int(float(placa))
        except: pass
    slot_type = clean(row[1])
    if slot_type: current_slot_type = slot_type
    puerto_o = clean(row[2])
    r_d, eq_d, sl_d, p_d, nota = clean(row[3]), clean(row[4]), clean(row[5]), clean(row[6]), clean(row[7])
    if puerto_o and r_d and r_d not in ('N/C', '?', 'D.Back'):
        slot_label = f"Slot {current_placa} ({current_slot_type})" if current_slot_type else f"Slot {current_placa}"
        add('D.7', f'openGear X {slot_label}', puerto_o, r_d, eq_d, p_d or 'In', '', nota or '')

# ── Rack F.7 FOR-A HVS-390HS ─────────────────────────────────────────────────
df = all_sheets['Rack F.7 FOR A HVS-390HS']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = clean(row[0])
    r_o, eq_o, p_o, nota = clean(row[1]), clean(row[2]), clean(row[4]), clean(row[5])
    if in_n and r_o and r_o not in ('?',):
        add(r_o, eq_o, p_o or 'Out', 'F.7', 'FOR A HVS-390HS', f'In {in_n}', '', nota or '')
    out_n = clean(row[8])
    r_d, eq_d, p_d, nota2 = clean(row[9]), clean(row[10]), clean(row[11]), clean(row[12])
    if out_n and r_d and r_d not in ('?', 'N/C'):
        add('F.7', 'FOR A HVS-390HS', f'{out_n}', r_d, eq_d, p_d or 'In', '', nota2 or '')

# ── Output ───────────────────────────────────────────────────────────────────
with open('connections.json', 'w', encoding='utf-8') as f:
    json.dump(connections, f, ensure_ascii=False, indent=2)

print(f"OK: {len(connections)} conexiones → connections.json")
