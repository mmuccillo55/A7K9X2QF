"""
parse.py — Conexiones.xlsx -> connections.json

Uso:
    python parse.py [archivo.xlsx] [salida.json]

Defaults:
    archivo : Conexiones.xlsx
    salida  : connections.json

El parser es completamente genérico: no hardcodea nombres de equipos,
racks, placas ni puertos. Todo lo lee del xlsx.

Hojas requeridas:
    - Conexiones : tabla de cableado
    - DB         : inventario de equipos con IPs

Columnas requeridas en Conexiones (por nombre, el orden no importa):
    Rack (Origen), Equipo (Origen), Slot (Origen), Placa (Origen),
    Puerto (Origen), Thru, Puerto Thru,
    Rack (Destino), Equipo (Destino), Slot (Destino), Puerto (Destino),
    Rótulo, Notas

Columnas requeridas en DB (por nombre):
    Rack, Equipo, IP
"""

import pandas as pd
import json
import sys
import os

# ── Configuración ─────────────────────────────────────────────────────────────

XLSX_FILE = sys.argv[1] if len(sys.argv) > 1 else 'Conexiones.xlsx'
JSON_FILE = sys.argv[2] if len(sys.argv) > 2 else 'connections.json'

# Valores que significan "sin destino / desconocido"
SIN_DESTINO = {'', 'N/D', 'N/C'}

# Valor en columna Thru que activa el arco interno
THRU_VALOR = 'T'

# Columnas esperadas
COL = {
    'src_rack':    'Rack (Origen)',
    'src_equipo':  'Equipo (Origen)',
    'src_slot':    'Slot (Origen)',
    'src_placa':   'Placa (Origen)',
    'src_puerto':  'Puerto (Origen)',
    'thru':        'Thru',
    'thru_entrada':'Puerto Thru',
    'dst_rack':    'Rack (Destino)',
    'dst_equipo':  'Equipo (Destino)',
    'dst_slot':    'Slot (Destino)',
    'dst_puerto':  'Puerto (Destino)',
    'rotulo':      'Rótulo',
    'notas':       'Notas',
    'disp_ext':    'Disp. Ext.',
}

COL_DB = {
    'rack':   'Rack',
    'equipo': 'Equipo',
    'ip':     'IP',
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def cv(v):
    """Limpia un valor de celda. Devuelve string o ''."""
    if pd.isna(v):
        return ''
    s = str(v).strip()
    return '' if s in ('nan', 'NaN') else s


def parse_ip(v):
    """
    Convierte el valor de IP del DB:
      - celda vacía / NaN  -> None  (sin red)
      - 'N/D'                -> 'N/D'   (pendiente)
      - cualquier otro     -> string tal cual
    """
    s = cv(v)
    if s == '':
        return None
    return s


def sin_destino(v):
    return cv(v) in SIN_DESTINO


def make_endpoint(rack, equipo, slot, placa, puerto):
    return {
        'rack':   cv(rack),
        'equipo': cv(equipo),
        'slot':   cv(slot),
        'placa':  cv(placa),
        'puerto': cv(puerto),
    }

# ── Validación de columnas ────────────────────────────────────────────────────

def check_columns(df, required, sheet_name):
    missing = [col for col in required.values() if col not in df.columns]
    if missing:
        print(f"ERROR: en la hoja '{sheet_name}' faltan columnas: {missing}", file=sys.stderr)
        sys.exit(1)

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
df_db  = sheets['DB']

check_columns(df_con, COL, 'Conexiones')
check_columns(df_db,  COL_DB, 'DB')

# ── Construir nodos desde DB ──────────────────────────────────────────────────

nodos = []
db_index = {}  # (rack, equipo) -> ip

for _, row in df_db.iterrows():
    rack   = cv(row[COL_DB['rack']])
    equipo = cv(row[COL_DB['equipo']])
    ip     = parse_ip(row[COL_DB['ip']])

    if not rack and not equipo:
        continue

    nodo = {
        'rack':   rack,
        'equipo': equipo,
        'ip':     ip,
    }
    nodos.append(nodo)
    db_index[(rack, equipo)] = ip

# ── Construir conexiones desde hoja Conexiones ────────────────────────────────

conexiones = []
warnings   = []

def add(src_rack, src_equipo, src_slot, src_placa, src_puerto,
        dst_rack, dst_equipo, dst_slot, dst_placa, dst_puerto,
        thru, thru_entrada, rotulo, notas, es_thru=False, disp_ext=''):

    conexiones.append({
        'src': make_endpoint(src_rack, src_equipo, src_slot, src_placa, src_puerto),
        'dst': make_endpoint(dst_rack, dst_equipo, dst_slot, dst_placa, dst_puerto),
        'thru':         thru,
        'thru_entrada': cv(thru_entrada),
        'rotulo':       cv(rotulo),
        'notas':        cv(notas),
        'es_thru':      es_thru,
        'disp_ext':     cv(disp_ext),
    })


for idx, row in df_con.iterrows():
    fila = idx + 2  # número de fila en Excel (1-indexed + header)

    src_rack   = cv(row[COL['src_rack']])
    src_equipo = cv(row[COL['src_equipo']])
    src_slot   = cv(row[COL['src_slot']])
    src_placa  = cv(row[COL['src_placa']])
    src_puerto = cv(row[COL['src_puerto']])

    thru        = cv(row[COL['thru']]) == THRU_VALOR
    thru_entrada = cv(row[COL['thru_entrada']])

    dst_rack   = cv(row[COL['dst_rack']])
    dst_equipo = cv(row[COL['dst_equipo']])
    dst_slot   = cv(row[COL['dst_slot']])
    dst_puerto = cv(row[COL['dst_puerto']])

    rotulo   = cv(row[COL['rotulo']])
    notas    = cv(row[COL['notas']])
    disp_ext = cv(row[COL['disp_ext']]) if COL['disp_ext'] in df_con.columns else ''

    # Si no hay nada útil en la fila, saltar
    if not src_rack and not src_equipo and not src_puerto:
        continue

    # Warning si src no está en DB
    if src_rack not in ('', 'N/D') and src_equipo not in ('', 'N/D'):
        if (src_rack, src_equipo) not in db_index:
            warnings.append(f"fila {fila}: origen '{src_rack} / {src_equipo}' no está en DB")

    # Warning si dst no está en DB (solo si tiene destino conocido)
    if not sin_destino(dst_rack) and dst_equipo not in ('', 'N/D'):
        if (dst_rack, dst_equipo) not in db_index:
            warnings.append(f"fila {fila}: destino '{dst_rack} / {dst_equipo}' no está en DB")

    # Si destino es N/C o N/D, mantener en rack pero limpiar otros campos
    # Si destino está completamente vacío, dejar todo vacío
    if dst_rack in ('N/C', 'N/D'):
        # Mantener el valor N/C o N/D en dst_rack
        # Limpiar los otros campos
        dst_equipo = dst_slot = dst_puerto = ''
    elif sin_destino(dst_rack):
        # Si es otro valor en SIN_DESTINO (ej: ''), limpiar todo
        dst_rack = dst_equipo = dst_slot = dst_puerto = ''

    # Conexión principal (el cable externo)
    add(
        src_rack, src_equipo, src_slot, src_placa, src_puerto,
        dst_rack, dst_equipo, dst_slot, '',         dst_puerto,
        False, '', rotulo, notas, disp_ext=disp_ext
    )

    # Arco interno thru: Puerto Thru -> Puerto Origen (dentro del mismo equipo)
    if thru:
        if not thru_entrada:
            warnings.append(f"fila {fila}: Thru=T pero 'Puerto Thru' está vacío — arco interno no generado")
        else:
            add(
                src_rack, src_equipo, src_slot, src_placa, thru_entrada,
                src_rack, src_equipo, src_slot, src_placa, src_puerto,
                False, '', '', '', es_thru=True
            )

# ── Estadísticas ──────────────────────────────────────────────────────────────

total         = len(conexiones)
thru_count    = sum(1 for c in conexiones if c['es_thru'])
con_destino   = sum(1 for c in conexiones if c['dst']['rack'] != '' and not c['es_thru'])
sin_dst_count = sum(1 for c in conexiones if c['dst']['rack'] == '' and not c['es_thru'])

stats = {
    'nodos':          len(nodos),
    'conexiones':     total - thru_count,
    'arcos_thru':     thru_count,
    'con_destino':    con_destino,
    'sin_destino':    sin_dst_count,
}

# ── Output ────────────────────────────────────────────────────────────────────

proyecto = os.path.splitext(os.path.basename(XLSX_FILE))[0]

output = {
    'proyecto':    proyecto,
    'nodos':       nodos,
    'conexiones':  conexiones,
    'stats':       stats,
}

with open(JSON_FILE, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

# ── Resumen en stdout ─────────────────────────────────────────────────────────

print(f"OK: '{XLSX_FILE}' -> '{JSON_FILE}'")
print(f"  nodos:          {stats['nodos']}")
print(f"  conexiones:     {stats['conexiones']}  ({stats['con_destino']} con destino, {stats['sin_destino']} sin destino)")
print(f"  arcos thru:     {stats['arcos_thru']}")

if warnings:
    print(f"\n  WARNINGS ({len(warnings)}):")
    for w in warnings:
        print(f"    ! {w}")
