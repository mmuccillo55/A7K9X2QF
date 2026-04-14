"""
parse.py — Ext1.xlsx -> ext1.json
"""

import pandas as pd
import json
import sys
import os

# ── Configuración ─────────────────────────────────────────────────────────────

XLSX_FILE = sys.argv[1] if len(sys.argv) > 1 else 'Ext1.xlsx'
JSON_FILE = sys.argv[2] if len(sys.argv) > 2 else 'ext1.json'

THRU_MARCA = 'T'

COL = {
    'src_rack':     'Rack (Origen)',
    'src_equipo':   'Equipo (Origen)',
    'src_slot':     'Slot (Origen)',
    'src_placa':    'Placa (Origen)',
    'src_puerto':   'Puerto (Origen)',
    'thru':         'Thru',
    'thru_entrada': 'Puerto Thru',
    'dst_rack':     'Rack (Destino)',
    'dst_equipo':   'Equipo (Destino)',
    'dst_slot':     'Slot (Destino)',
    'dst_placa':    'Placa (Destino)',   # ← NUEVO
    'dst_puerto':   'Puerto (Destino)',
    'rotulo':       'Rótulo',
    'notas':        'Notas',
    'disp_ext':     'Disp. Ext.',
}

COL_DB = {
    'rack':   'Rack',
    'equipo': 'Equipo',
    'ip':     'IP',
    'puerto': 'Puerto',   # ← NUEVO
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def cv(v):
    if pd.isna(v):
        return ''
    
    # Si es número
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))  # 1021.0 → "1021"
        return str(v)
    
    s = str(v).strip()
    return '' if s in ('nan', 'NaN') else s


def parse_ip(v):
    s = cv(v)
    if s == '':
        return None
    return s


def make_endpoint(rack, equipo, slot, placa, puerto):
    return {
        'rack':   cv(rack),
        'equipo': cv(equipo),
        'slot':   cv(slot),
        'placa':  cv(placa),
        'puerto': cv(puerto),
    }


def check_columns(df, required, sheet_name):
    missing = [col for col in required.values() if col not in df.columns]
    if missing:
        print(f"ERROR: en la hoja '{sheet_name}' faltan columnas: {missing}", file=sys.stderr)
        sys.exit(1)


def endpoint_to_key(ep):
    return f"{ep['rack']}/{ep['slot']}/{ep['placa']}/{ep['puerto']}"


# ── Main ──────────────────────────────────────────────────────────────────────

if not os.path.exists(XLSX_FILE):
    print(f"ERROR: no se encontró '{XLSX_FILE}'", file=sys.stderr)
    sys.exit(1)

sheets = pd.read_excel(XLSX_FILE, sheet_name=None)

for sheet in ('Conexiones', 'DB'):
    if sheet not in sheets:
        print(f"ERROR: no se encontró la hoja '{sheet}'", file=sys.stderr)
        sys.exit(1)

df_con = sheets['Conexiones']
df_db = sheets['DB']

check_columns(df_con, COL, 'Conexiones')
check_columns(df_db, COL_DB, 'DB')

# ── Construir nodos desde DB ──────────────────────────────────────────────────

nodos = []
nodos_index = {}

for _, row in df_db.iterrows():
    rack = cv(row[COL_DB['rack']])
    equipo = cv(row[COL_DB['equipo']])
    ip = parse_ip(row[COL_DB['ip']])  # ← FALTABA ESTO
    puerto = cv(row[COL_DB['puerto']]) if COL_DB['puerto'] in df_db.columns else ''
    
    if not rack and not equipo:
        continue
    
    nodo = {
        'rack':   rack,
        'equipo': equipo,
        'ip':     ip,
        'management': {
            'puerto': puerto
        }
    }
    
    nodos.append(nodo)
    nodos_index[rack] = nodo

# ── Construir conexiones ──────────────────────────────────────────────────────

conexiones = []
warnings = []

for idx, row in df_con.iterrows():
    fila = idx + 2
    
    # Origen
    src_rack = cv(row[COL['src_rack']])
    src_equipo = cv(row[COL['src_equipo']])
    src_slot = cv(row[COL['src_slot']])
    src_placa = cv(row[COL['src_placa']])
    src_puerto = cv(row[COL['src_puerto']])
    
    # Thru
    tiene_thru = cv(row[COL['thru']]) == THRU_MARCA
    thru_entrada = cv(row[COL['thru_entrada']])
    
    # Destino
    dst_rack = cv(row[COL['dst_rack']])
    dst_equipo = cv(row[COL['dst_equipo']])
    dst_slot = cv(row[COL['dst_slot']])
    dst_placa = cv(row[COL['dst_placa']])
    dst_puerto = cv(row[COL['dst_puerto']])
    
    # Metadata
    rotulo = cv(row[COL['rotulo']])
    notas = cv(row[COL['notas']])
    disp_ext = cv(row[COL['disp_ext']]) if COL['disp_ext'] in df_con.columns else ''
    
    if not src_rack and not src_equipo and not src_puerto:
        continue
    
    # Validaciones
    if src_rack not in ('', 'N/D'):
        if src_rack not in nodos_index:
            warnings.append(f"fila {fila}: origen '{src_rack} / {src_equipo}' no está en DB")
    
    if dst_rack not in ('', 'N/C', 'N/D') and dst_equipo not in ('', 'N/D'):
        if dst_rack not in nodos_index:
            warnings.append(f"fila {fila}: destino '{dst_rack} / {dst_equipo}' no está en DB")
    
    # ── Arco externo ───────────────────────────────────────────────────────────
    
    arco_externo = {
        'src': make_endpoint(src_rack, src_equipo, src_slot, src_placa, src_puerto),
        'dst': make_endpoint(dst_rack, dst_equipo, dst_slot, dst_placa, dst_puerto),
        'rotulo': rotulo,
        'notas': notas,
        'disp_ext': disp_ext,
    }
    conexiones.append(arco_externo)
    
    # ── Arco interno (Thru) ───────────────────────────────────────────────────
    
    if tiene_thru:
        if not thru_entrada:
            warnings.append(f"fila {fila}: Thru=T pero 'Puerto Thru' está vacío")
        else:
            arco_interno = {
                'src': make_endpoint(src_rack, src_equipo, src_slot, src_placa, thru_entrada),
                'dst': make_endpoint(src_rack, src_equipo, src_slot, src_placa, src_puerto),
                'rotulo': '',
                'notas': '',
                'disp_ext': '',
                'es_thru_interno': True,
            }
            conexiones.append(arco_interno)

# ── Stats ─────────────────────────────────────────────────────────────────────

arcos_thru = sum(1 for c in conexiones if c.get('es_thru_interno', False))
arcos_externos = len(conexiones) - arcos_thru
arcos_con_destino = sum(1 for c in conexiones if c['dst']['puerto'] != '' and not c.get('es_thru_interno', False))
arcos_sin_destino = arcos_externos - arcos_con_destino

stats = {
    'nodos': len(nodos),
    'arcos_totales': len(conexiones),
    'arcos_externos': arcos_externos,
    'arcos_thru_internos': arcos_thru,
    'arcos_con_destino': arcos_con_destino,
    'arcos_sin_destino': arcos_sin_destino,
}

# ── Output ────────────────────────────────────────────────────────────────────

proyecto = os.path.splitext(os.path.basename(XLSX_FILE))[0]

output = {
    'proyecto': proyecto,
    'nodos': nodos,
    'conexiones': conexiones,
    'stats': stats,
}

with open(JSON_FILE, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

# ── Resumen ───────────────────────────────────────────────────────────────────

print(f"OK: '{XLSX_FILE}' -> '{JSON_FILE}'")
print(f"  nodos:               {stats['nodos']}")
print(f"  arcos totales:       {stats['arcos_totales']}")
print(f"    - externos:        {stats['arcos_externos']}")
print(f"    - thru internos:   {stats['arcos_thru_internos']}")

if warnings:
    print(f"\nWARNINGS ({len(warnings)}):")
    for w in warnings:
        print(f"  ! {w}")
