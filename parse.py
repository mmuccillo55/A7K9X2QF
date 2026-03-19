import pandas as pd
import json
import sys
import os

XLSX = 'RACKS_C12.xlsx'

if not os.path.exists(XLSX):
    print(f"ERROR: no se encontró {XLSX}", file=sys.stderr)
    sys.exit(1)

# Ignorar la hoja "DB" — solo es referencia de nombres
all_sheets = pd.read_excel(XLSX, sheet_name=None, header=None)
all_sheets.pop('DB', None)

# ── Helpers ──────────────────────────────────────────────────────────────────

def clean(v):
    if pd.isna(v): return None
    s = str(v).strip()
    return s if s not in ('', 'nan', 'NaN') else None

def safeget(row, idx):
    """Obtiene clean(row[idx]) sin explotar si la columna no existe."""
    try:
        return clean(row[idx])
    except (KeyError, IndexError):
        return None

SKIP = {'?', 'N/C', 'VACANTE', 'N/A'}

def skip(v):
    return not v or v in SKIP

connections = []

def add(src_rack, src_eq, src_port, dst_rack, dst_eq, dst_port, label='', notes=''):
    if skip(src_rack) or skip(dst_rack): return
    if skip(src_eq) or skip(dst_eq): return
    connections.append({
        'src_rack': src_rack,
        'src_eq':   src_eq   or '',
        'src_port': src_port or '',
        'dst_rack': dst_rack,
        'dst_eq':   dst_eq   or '',
        'dst_port': dst_port or '',
        'label':    label    or '',
        'notes':    notes    or '',
    })

# ── Rack A ───────────────────────────────────────────────────────────────────
# Cols: 0=src_rack 1=src_eq 2=src_slot 3=src_port 4=dst_rack 5=dst_eq 6=dst_slot 7=dst_port 8=notas
df = all_sheets['Rack A']
for i in range(2, len(df)):
    row = df.iloc[i]
    r_o, eq_o, p_o = safeget(row,0), safeget(row,1), safeget(row,3)
    r_d, eq_d, p_d = safeget(row,4), safeget(row,5), safeget(row,7)
    nota            = safeget(row,8)
    add(r_o, eq_o, p_o, r_d, eq_d, p_d, eq_o or '', nota or '')

# ── Rack B ───────────────────────────────────────────────────────────────────
# Cols: 0=dst_rack 1=dst_eq 2=dst_port 3=src_rack 4=src_eq 5=src_port 6=notas
df = all_sheets['Rack B']
for i in range(2, len(df)):
    row = df.iloc[i]
    r_d, eq_d, p_d = safeget(row,0), safeget(row,1), safeget(row,2)
    r_o, eq_o, p_o = safeget(row,3), safeget(row,4), safeget(row,5)
    nota            = safeget(row,6)
    add(r_o, eq_o, p_o, r_d, eq_d, p_d, eq_d or '', nota or '')

# ── Rack C.3 AJA KUMO 6464 ───────────────────────────────────────────────────
# Cols INPUTS  (0-6):  0=in_n 1=rotulo 2=src_rack 3=src_eq 4=src_slot 5=src_port 6=notas
# Cols OUTPUTS (9-15): 9=out_n 10=rotulo 11=dst_rack 12=dst_eq 13=dst_slot 14=dst_port 15=notas
df = all_sheets['Rack C.3 AJA KUMO 6464']
for i in range(2, len(df)):
    row = df.iloc[i]
    # Inputs → KUMO
    in_n   = safeget(row, 0);  rotulo = safeget(row, 1)
    r_o, eq_o, sl_o, p_o = safeget(row,2), safeget(row,3), safeget(row,4), safeget(row,5)
    nota   = safeget(row, 6)
    if in_n and not skip(r_o) and r_o != 'K':
        port_o = p_o or (f"Slot {sl_o}" if sl_o else 'Out')
        add(r_o, eq_o, port_o, 'C.3', 'AJA KUMO 6464', f'In {in_n}', rotulo or eq_o or '', nota or '')
    # KUMO → Outputs
    out_n  = safeget(row, 9);  rotulo2 = safeget(row, 10)
    r_d, eq_d, sl_d, p_d = safeget(row,11), safeget(row,12), safeget(row,13), safeget(row,14)
    nota2  = safeget(row, 15)
    if out_n and not skip(r_d):
        port_d = p_d or (f"Slot {sl_d}" if sl_d else 'In')
        add('C.3', 'AJA KUMO 6464', f'Out {out_n}', r_d, eq_d, port_d, rotulo2 or eq_d or '', nota2 or '')

# ── Rack C.4 Blackmagic MultiView ────────────────────────────────────────────
# Cols INPUTS  (0-4):  0=in_n 1=src_rack 2=src_eq 3=src_port 4=notas
# Cols OUTPUTS (7-12): 7=out_n 8=dst_rack 9=dst_eq 10=dst_slot 11=dst_port 12=notas
# Fila 19 especial: SDI OUT en col 7, destino en cols 8-12
df = all_sheets['Rack C.4 Blackmagic Multiview 1']
for i in range(2, len(df)):
    row = df.iloc[i]
    # Inputs → MultiView
    in_n = safeget(row, 0)
    r_o, eq_o, p_o = safeget(row,1), safeget(row,2), safeget(row,3)
    nota = safeget(row, 4)
    if in_n and not skip(r_o):
        add(r_o, eq_o, p_o or 'Out', 'C.4', 'Blackmagic MultiView 16', f'In {in_n}', '', nota or '')
    # MultiView → Outputs
    out_n = safeget(row, 7)
    r_d   = safeget(row, 8)
    eq_d  = safeget(row, 9)
    p_d   = safeget(row, 11)
    nota2 = safeget(row, 12)
    if out_n and not skip(r_d):
        add('C.4', 'Blackmagic MultiView 16', f'Out {out_n}', r_d, eq_d, p_d or 'SDI In', '', nota2 or '')
    # Fila 19 especial: SDI OUT
    if i == 19:
        sdi_label = safeget(row, 7)   # 'SDI OUT'
        r_d2  = safeget(row, 8)
        eq_d2 = safeget(row, 9)
        p_d2  = safeget(row, 11)
        nota3 = safeget(row, 12)
        if sdi_label and not skip(r_d2):
            add('C.4', 'Blackmagic MultiView 16', 'SDI OUT', r_d2, eq_d2, p_d2 or 'In', '', nota3 or '')

# ── Rack C.5 Blackmagic 12x12 ────────────────────────────────────────────────
# Cols INPUTS  (0-4):  0=in_n 1=src_rack 2=src_eq 3=src_port 4=notas
# Cols OUTPUTS (7-11): 7=out_n 8=dst_rack 9=dst_eq 10=dst_port 11=notas
df = all_sheets['Rack C.5 Blackmagic 12 x 12']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = safeget(row, 0)
    r_o, eq_o, p_o = safeget(row,1), safeget(row,2), safeget(row,3)
    nota = safeget(row, 4)
    if in_n and not skip(r_o):
        add(r_o, eq_o, p_o or 'Out', 'C.5', 'Blackmagic 12x12', f'In {in_n}', nota or '', '')
    out_n = safeget(row, 7)
    r_d, eq_d, p_d = safeget(row,8), safeget(row,9), safeget(row,10)
    nota2 = safeget(row, 11)
    if out_n and not skip(r_d):
        add('C.5', 'Blackmagic 12x12', f'Out {out_n}', r_d, eq_d, p_d or 'In', '', nota2 or '')

# ── Rack C.6 Patch Panel Inputs ──────────────────────────────────────────────
# Cols: 0=patch_n 1=src_rack 2=src_eq 3=src_port 4=notas
# Semántica: el patch recibe de rack externo y conecta al KUMO
df = all_sheets['Rack C.6 Patch panel Inputs']
for i in range(2, len(df)):
    row = df.iloc[i]
    patch_n = safeget(row, 0)
    r_o     = safeget(row, 1)
    eq_o    = safeget(row, 2)
    p_o     = safeget(row, 3)
    nota    = safeget(row, 4)
    if patch_n and not skip(r_o):
        add(r_o, eq_o or 'Patch Panel (Inputs)', f'Port {patch_n}',
            'C.3', 'AJA KUMO 6464', p_o or f'In {patch_n}', '', nota or '')

# ── Rack C.7 Patch Panel Outputs ─────────────────────────────────────────────
# Cols: 0=patch_n 1=src_rack 2=src_eq 3=src_port 4=notas
# Semántica: KUMO → Patch Panel Outputs
df = all_sheets['Rack C.7 Patch panel Outputs']
for i in range(2, len(df)):
    row = df.iloc[i]
    patch_n = safeget(row, 0)
    r_o     = safeget(row, 1)
    eq_o    = safeget(row, 2)
    p_o     = safeget(row, 3)
    nota    = safeget(row, 4)
    if patch_n and not skip(r_o) and eq_o:
        add(r_o, eq_o, p_o or f'Out {patch_n}',
            'C.7', 'Patch Panel (Outputs)', f'Port {patch_n}', nota or '', '')

# ── Rack D.Back ───────────────────────────────────────────────────────────────
# Cols: 0=src_rack 1=src_eq 2=src_slot 3=src_port 4=dst_rack 5=dst_eq 6=dst_slot 7=dst_port 8=notas
df = all_sheets['Rack D.Back']
for i in range(2, len(df)):
    row = df.iloc[i]
    r_o, eq_o, p_o = safeget(row,0), safeget(row,1), safeget(row,3)
    r_d, eq_d, p_d = safeget(row,4), safeget(row,5), safeget(row,7)
    nota            = safeget(row,8)
    add(r_o, eq_o, p_o or 'Out', r_d, eq_d, p_d or 'In', '', nota or '')

# ── Rack D.7 openGear X ──────────────────────────────────────────────────────
# Cols: 0=placa 1=slot_type 2=puerto_o 3=dst_rack 4=dst_eq 5=dst_slot 6=dst_port 7=notas
# Las filas de placa/slot_type se heredan hacia abajo mientras no aparezca nuevo valor
df = all_sheets['Rack D.7 openGear X']
current_placa     = None
current_slot_type = None
for i in range(2, len(df)):
    row = df.iloc[i]
    placa = safeget(row, 0)
    if placa:
        try: current_placa = int(float(placa))
        except: current_placa = placa
    slot_type = safeget(row, 1)
    if slot_type: current_slot_type = slot_type
    puerto_o = safeget(row, 2)
    r_d, eq_d, p_d = safeget(row,3), safeget(row,4), safeget(row,6)
    nota            = safeget(row,7)
    if puerto_o and not skip(r_d) and r_d != 'D.Back':
        slot_label = (f"Slot {current_placa} ({current_slot_type})"
                      if current_slot_type else f"Slot {current_placa}")
        add('D.7', f'openGear X {slot_label}', puerto_o,
            r_d, eq_d, p_d or 'In', '', nota or '')

# ── Rack D.9 openGear X ──────────────────────────────────────────────────────
# Misma estructura que D.7
df = all_sheets['Rack D.9 openGear X']
current_placa     = None
current_slot_type = None
for i in range(2, len(df)):
    row = df.iloc[i]
    placa = safeget(row, 0)
    if placa:
        try: current_placa = int(float(placa))
        except: current_placa = placa
    slot_type = safeget(row, 1)
    if slot_type: current_slot_type = slot_type
    puerto_o = safeget(row, 2)
    r_d, eq_d, p_d = safeget(row,3), safeget(row,4), safeget(row,6)
    nota            = safeget(row,7)
    if puerto_o and not skip(r_d):
        slot_label = (f"Slot {current_placa} ({current_slot_type})"
                      if current_slot_type else f"Slot {current_placa}")
        add('D.9', f'openGear X {slot_label}', puerto_o,
            r_d, eq_d, p_d or 'In', '', nota or '')

# ── Rack F.7 FOR-A HVS-390HS ─────────────────────────────────────────────────
# Cols INPUTS  (0-5):  0=in_n 1=src_rack 2=src_eq 3=src_slot 4=src_port 5=notas
# Cols OUTPUTS (7-11): 7=out_n 8=dst_rack 9=dst_eq 10=dst_port 11=notas
df = all_sheets['Rack F.7 FOR A HVS-390HS']
for i in range(2, len(df)):
    row = df.iloc[i]
    in_n = safeget(row, 0)
    r_o, eq_o, p_o = safeget(row,1), safeget(row,2), safeget(row,4)
    nota = safeget(row, 5)
    if in_n and not skip(r_o):
        add(r_o, eq_o, p_o or 'Out', 'F.7', 'FOR A HVS-390HS', f'In {in_n}', '', nota or '')
    out_n = safeget(row, 7)
    r_d, eq_d, p_d = safeget(row,8), safeget(row,9), safeget(row,10)
    nota2 = safeget(row, 11)
    if out_n and not skip(r_d):
        add('F.7', 'FOR A HVS-390HS', str(out_n), r_d, eq_d, p_d or 'In', '', nota2 or '')

# ── Output ───────────────────────────────────────────────────────────────────
with open('connections.json', 'w', encoding='utf-8') as f:
    json.dump(connections, f, ensure_ascii=False, indent=2)

print(f"OK: {len(connections)} conexiones → connections.json")
