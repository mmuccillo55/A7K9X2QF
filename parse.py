import pandas as pd
import json
import sys
import os

# ── Config ────────────────────────────────────────────────────────────────────

if len(sys.argv) < 3:
    print("Use: python parse.py <file.xlsx> <out.json>", file=sys.stderr)
    sys.exit(1)

XLSX_FILE = sys.argv[1]
JSON_FILE = sys.argv[2]

THRU_FLAG = 'T'
MRG_FLAG  = 'M'
SC_FLAG   = 'SC'

COL = {
    'src_rack':   'Rack (Origen)',
    'src_equip':  'Equipo (Origen)',
    'src_slot':   'Slot (Origen)',
    'src_board':  'Placa (Origen)',
    'src_port':   'Puerto (Origen)',
    'link_type':  'Link',
    'link_ports': 'Puerto Link',
    'dst_rack':   'Rack (Destino)',
    'dst_equip':  'Equipo (Destino)',
    'dst_slot':   'Slot (Destino)',
    'dst_board':  'Placa (Destino)',
    'dst_port':   'Puerto (Destino)',
    'label':      'Rótulo',
    'notes':      'Notas',
    'ext_device': 'Disp. Ext.',
}

COL_DB = {
    'rack':  'Rack',
    'equip': 'Equipo',
    'ip':    'IP',
    'type':  'Tipo',
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def cv(v):
    if pd.isna(v):
        return ''
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))  # 1021.0 → "1021"
        return str(v)
    s = str(v).strip()
    return '' if s in ('nan', 'NaN') else s


def parse_ip(v):
    s = cv(v)
    return None if s == '' else s


def make_endpoint(rack, equip, slot, board, port):
    return {
        'rack':  cv(rack),
        'equip': cv(equip),
        'slot':  cv(slot),
        'board': cv(board),
        'port':  cv(port),
    }


def check_columns(df, required, sheet_name):
    missing = [col for col in required.values() if col not in df.columns]
    if missing:
        print(f"ERROR: missing columns in sheet '{sheet_name}': {missing}", file=sys.stderr)
        sys.exit(1)


def endpoint_to_key(ep):
    return f"{ep['rack']}/{ep['slot']}/{ep['board']}/{ep['port']}"


# ── Main ──────────────────────────────────────────────────────────────────────

if not os.path.exists(XLSX_FILE):
    print(f"ERROR: '{XLSX_FILE}' was not found", file=sys.stderr)
    sys.exit(1)

sheets = pd.read_excel(XLSX_FILE, sheet_name=None)

for sheet in ('Conexiones', 'DB'):
    if sheet not in sheets:
        print(f"ERROR: sheet '{sheet}' was not found", file=sys.stderr)
        sys.exit(1)

df_con = sheets['Conexiones']
df_db  = sheets['DB']

check_columns(df_con, COL, 'Conexiones')
check_columns(df_db, COL_DB, 'DB')

# ── Build nodes from DB ───────────────────────────────────────────────────────

nodes = []
nodes_index = {}

for _, row in df_db.iterrows():
    rack      = cv(row[COL_DB['rack']])
    equip     = cv(row[COL_DB['equip']])
    ip        = parse_ip(row[COL_DB['ip']])
    node_type = cv(row[COL_DB['type']]) if COL_DB['type'] in df_db.columns else ''

    if not rack and not equip:
        continue

    node = {
        'rack':  rack,
        'equip': equip,
        'ip':    ip,
        'type':  node_type or None,
    }

    nodes.append(node)
    nodes_index[rack] = node

# ── Build connections ─────────────────────────────────────────────────────────

connections = []
warnings    = []

for idx, row in df_con.iterrows():
    line_no = idx + 2

    src_rack  = cv(row[COL['src_rack']])
    src_equip = cv(row[COL['src_equip']])
    src_slot  = cv(row[COL['src_slot']])
    src_board = cv(row[COL['src_board']])
    src_port  = cv(row[COL['src_port']])

    link_type  = cv(row[COL['link_type']])
    link_ports = cv(row[COL['link_ports']])
    is_thru = link_type == THRU_FLAG
    is_mrg  = link_type == MRG_FLAG
    is_sc   = link_type == SC_FLAG

    dst_rack  = cv(row[COL['dst_rack']])
    dst_equip = cv(row[COL['dst_equip']])
    dst_slot  = cv(row[COL['dst_slot']])
    dst_board = cv(row[COL['dst_board']])
    dst_port  = cv(row[COL['dst_port']])

    label      = cv(row[COL['label']])
    notes      = cv(row[COL['notes']])
    ext_device = cv(row[COL['ext_device']]) if COL['ext_device'] in df_con.columns else ''

    if not src_rack and not src_equip and not src_port:
        continue

    # Validate source
    if src_rack not in ('', 'N/D'):
        if src_rack not in nodes_index:
            warnings.append(f"row {line_no}: source '{src_rack} / {src_equip}' not found in DB")

    # Validate destination
    if dst_rack not in ('', 'N/C', 'N/D') and dst_equip not in ('', 'N/D'):
        if dst_rack not in nodes_index:
            warnings.append(f"row {line_no}: destination '{dst_rack} / {dst_equip}' not found in DB")

    # ── External edge ─────────────────────────────────────────────────────────

    edge = {
        'src':        make_endpoint(src_rack, src_equip, src_slot, src_board, src_port),
        'dst':        make_endpoint(dst_rack, dst_equip, dst_slot, dst_board, dst_port),
        'label':      label,
        'notes':      notes,
        'ext_device': ext_device,
    }
    if is_sc:
        edge['is_sc'] = True
    connections.append(edge)

    # ── Internal edge (Thru) ──────────────────────────────────────────────────

    if is_thru:
        if not link_ports:
            warnings.append(f"row {line_no}: link_type=T but 'link_ports' is empty")
        else:
            connections.append({
                'src':        make_endpoint(src_rack, src_equip, src_slot, src_board, link_ports),
                'dst':        make_endpoint(src_rack, src_equip, src_slot, src_board, src_port),
                'label':      '',
                'notes':      '',
                'ext_device': '',
                'is_thru':    True,
            })

    if is_mrg:
        if not link_ports:
            warnings.append(f"row {line_no}: link_type=M but 'link_ports' is empty")
        else:
            for port in [p.strip() for p in link_ports.split(',')]:
                connections.append({
                    'src':        make_endpoint(src_rack, src_equip, src_slot, src_board, port),
                    'dst':        make_endpoint(src_rack, src_equip, src_slot, src_board, src_port),
                    'label':      '',
                    'notes':      '',
                    'ext_device': '',
                    'is_thru':    True,
                    'is_mrg':     True,
                })

# ── Stats ─────────────────────────────────────────────────────────────────────

thru_edges    = sum(1 for c in connections if c.get('is_thru', False))
merge_edges   = sum(1 for c in connections if c.get('is_mrg', False))
sc_edges      = sum(1 for c in connections if c.get('is_sc', False))
ext_edges     = len(connections) - thru_edges
linked_edges  = sum(1 for c in connections if c['dst']['port'] != '' and not c.get('is_thru', False))
dangling_edges = ext_edges - linked_edges

stats = {
    'nodes':       len(nodes),
    'total_edges': len(connections),
    'ext_edges':   ext_edges,
    'thru_edges':  thru_edges,
    'linked':      linked_edges,
    'dangling':    dangling_edges,
    'sc_edges':    sc_edges,
}

# ── Output ────────────────────────────────────────────────────────────────────

project = os.path.splitext(os.path.basename(XLSX_FILE))[0]

output = {
    'project':     project,
    'nodes':       nodes,
    'connections': connections,
    'stats':       stats,
}

with open(JSON_FILE, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

# ── Summary ───────────────────────────────────────────────────────────────────

print(f"  nodes:              {stats['nodes']}")
print(f"  total edges:        {stats['total_edges']}")
print(f"    - external:       {stats['ext_edges']}")
print(f"    - thru (internal):{stats['thru_edges']}")
print(f"    - same connector: {stats['sc_edges']}")

if warnings:
    print(f"\nWARNINGS ({len(warnings)}):")
    for w in warnings:
        print(f"  ! {w}")