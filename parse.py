"""
parse.py — Conexiones.xlsx -> connections.json

Uso:
    python parse.py [archivo.xlsx] [salida.json]

Defaults:
    archivo : Conexiones.xlsx
    salida  : connections.json

Parser completamente genérico basado en la estructura de cableado broadcast.

Lógica:
  - Cada fila en "Conexiones" = una conexión origen -> destino
  - Si hay Thru=T, se generan DOS entradas en el JSON:
    1. Arco interno: (thru_entrada -> src_puerto) dentro del mismo equipo
    2. Arco externo: (src_puerto -> dst_puerto) hacia el destino
  - Rótulo y Notas solo en arcos externos
  - N/C = puerto no conectado
  - N/D = pendiente de documentar

Hojas requeridas:
    - Conexiones : tabla de cableado
    - DB         : inventario de equipos

Columnas requeridas en Conexiones:
    Rack (Origen), Equipo (Origen), Slot (Origen), Placa (Origen),
    Puerto (Origen), Thru, Puerto Thru,
    Rack (Destino), Equipo (Destino), Slot (Destino), Puerto (Destino),
    Rótulo, Notas

Columnas requeridas en DB:
    Rack, Equipo, IP
"""

import pandas as pd
import json
import sys
import os

# ── Configuración ─────────────────────────────────────────────────────────────

XLSX_FILE = sys.argv[1] if len(sys.argv) > 1 else 'Conexiones.xlsx'
JSON_FILE = sys.argv[2] if len(sys.argv) > 2 else 'connections.json'

SIN_DESTINO = {'', 'N/D', 'N/C'}
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
    'dst_puerto':   'Puerto (Destino)',
    'rotulo':       'Rótulo',
    'notas':        'Notas',
    'disp_ext':     'Disp. Ext.',
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
    """Convierte IP del DB: vacío -> None, N/D -> 'N/D', otro -> string."""
    s = cv(v)
    if s == '':
        return None
    return s


def sin_destino(v):
    """¿Es un valor que significa 'sin destino'?"""
    return cv(v) in SIN_DESTINO


def make_endpoint(rack, equipo, slot, placa, puerto):
    """Crea un endpoint estructurado."""
    return {
        'rack':   cv(rack),
        'equipo': cv(equipo),
        'slot':   cv(slot),
        'placa':  cv(placa),
        'puerto': cv(puerto),
    }


def check_columns(df, required, sheet_name):
    """Valida que existan las columnas requeridas."""
    missing = [col for col in required.values() if col not in df.columns]
    if missing:
        print(f"ERROR: en la hoja '{sheet_name}' faltan columnas: {missing}", file=sys.stderr)
        sys.exit(1)


def endpoint_to_key(ep):
    """Convierte un endpoint a una clave única para búsquedas."""
    return f"{ep['rack']}.{ep['pos']}/{ep['slot']}/{ep['placa']}/{ep['puerto']}"


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
nodos_index = {}  # rack -> nodo info

for _, row in df_db.iterrows():
    rack = cv(row[COL_DB['rack']])
    equipo = cv(row[COL_DB['equipo']])
    ip = parse_ip(row[COL_DB['ip']])
    
    if not rack and not equipo:
        continue
    
    nodo = {
        'rack':   rack,
        'equipo': equipo,
        'ip':     ip,
    }
    nodos.append(nodo)
    nodos_index[rack] = nodo

# ── Construir conexiones desde hoja Conexiones ────────────────────────────────

conexiones = []
warnings = []

for idx, row in df_con.iterrows():
    fila = idx + 2  # número de fila en Excel (1-indexed + header)
    
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
    dst_puerto = cv(row[COL['dst_puerto']])
    
    # Metadata
    rotulo = cv(row[COL['rotulo']])
    notas = cv(row[COL['notas']])
    disp_ext = cv(row[COL['disp_ext']]) if COL['disp_ext'] in df_con.columns else ''
    
    # Si no hay nada útil en origen, saltar fila
    if not src_rack and not src_equipo and not src_puerto:
        continue
    
    # Validar que origen esté en DB (si no es N/D)
    if src_rack not in ('', 'N/D'):
        if src_rack not in nodos_index:
            warnings.append(f"fila {fila}: origen '{src_rack} / {src_equipo}' no está en DB")
    
    # Validar que destino esté en DB (si tiene destino conocido y no es N/D)
    if not sin_destino(dst_rack) and dst_equipo not in ('', 'N/D'):
        if dst_rack not in nodos_index:
            warnings.append(f"fila {fila}: destino '{dst_rack} / {dst_equipo}' no está en DB")
    
    # ── Arco externo (la conexión principal: src -> dst) ───────────────────────
    
    # Limpiar destino si es N/C o N/D
    if dst_rack in ('N/C', 'N/D'):
        # Mantener el valor para que sea claro en el JSON
        dst_equipo_clean = ''
        dst_slot_clean = ''
        dst_puerto_clean = ''
    elif sin_destino(dst_rack):
        # Si es otro valor en SIN_DESTINO, limpiar todo
        dst_rack = ''
        dst_equipo_clean = ''
        dst_slot_clean = ''
        dst_puerto_clean = ''
    else:
        dst_equipo_clean = dst_equipo
        dst_slot_clean = dst_slot
        dst_puerto_clean = dst_puerto
    
    arco_externo = {
        'src': make_endpoint(src_rack, src_equipo, src_slot, src_placa, src_puerto),
        'dst': make_endpoint(dst_rack, dst_equipo_clean, dst_slot_clean, '', dst_puerto_clean),
        'rotulo': rotulo,
        'notas': notas,
        'disp_ext': disp_ext,
    }
    conexiones.append(arco_externo)
    
    # ── Arco interno (Thru) ───────────────────────────────────────────────────
    
    if tiene_thru:
        if not thru_entrada:
            warnings.append(f"fila {fila}: Thru=T pero 'Puerto Thru' está vacío — arco interno no generado")
        else:
            # Arco interno: thru_entrada -> src_puerto (dentro del mismo equipo)
            arco_interno = {
                'src': make_endpoint(src_rack, src_equipo, src_slot, src_placa, thru_entrada),
                'dst': make_endpoint(src_rack, src_equipo, src_slot, src_placa, src_puerto),
                'rotulo': '',
                'notas': '',
                'disp_ext': '',
                'es_thru_interno': True,
            }
            conexiones.append(arco_interno)

# ── Estadísticas ──────────────────────────────────────────────────────────────

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

# ── Resumen en stdout ─────────────────────────────────────────────────────────

print(f"OK: '{XLSX_FILE}' -> '{JSON_FILE}'")
print(f"  nodos:               {stats['nodos']}")
print(f"  arcos totales:       {stats['arcos_totales']}")
print(f"    - externos:        {stats['arcos_externos']}  ({stats['arcos_con_destino']} con destino, {stats['arcos_sin_destino']} sin destino)")
print(f"    - thru internos:   {stats['arcos_thru_internos']}")

if warnings:
    print(f"\n  WARNINGS ({len(warnings)}):")
    for w in warnings:
        print(f"    ! {w}")
