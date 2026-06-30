# -*- coding: utf-8 -*-
import sys, os, json, re, webbrowser, io, time
from datetime import datetime

# Fix Windows stdout encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl no instalado. Ejecuta: pip install openpyxl")
    sys.exit(1)

DIR         = r'C:\Users\franc\OneDrive\Documentos\KeyCarbon\claude MONCAR RMH'
BASE_DIR    = DIR + r'\ARCHIVOS BASE'
VENTAS_PATH = BASE_DIR + r'\VENTAS.xlsx'
MC_PATH     = BASE_DIR + r'\STOCK_MONCAR.xlsx'
RMH_PATH    = BASE_DIR + r'\STOCK_RMH.xlsx'
OUT_PATH    = DIR + r'\dashboard_live.html'

# ─── helpers ────────────────────────────────────────────────────────────────
def load_wb(path, data_only=False):
    """Carga un workbook con reintentos para manejar locks de Excel/OneDrive."""
    for intento in range(4):
        try:
            with open(path, 'rb') as f:
                data = f.read()
            return openpyxl.load_workbook(io.BytesIO(data), data_only=data_only)
        except PermissionError:
            if intento < 3:
                print(f"  Archivo en uso, esperando... (intento {intento+1}/3)")
                time.sleep(2)
            else:
                raise PermissionError(
                    f"No se puede leer '{os.path.basename(path)}'. "
                    "Cerrá el archivo en Excel antes de actualizar."
                )

def s(val):
    return str(val).strip() if val is not None else ""

def n(val):
    return float(val) if isinstance(val, (int, float)) else 0.0

# ─── read VENTAS ─────────────────────────────────────────────────────────────
def read_ventas(path):
    print("Leyendo VENTAS...")
    wb = load_wb(path, data_only=True)

    # ── LP sheet (precios)
    precios = {}
    if "LP" in wb.sheetnames:
        ws_lp = wb["LP"]
        for r in ws_lp.iter_rows(min_row=4, values_only=True):
            cod = s(r[0]) if len(r) > 0 else ""
            precio = n(ws_lp.cell(ws_lp.min_row + list(ws_lp.iter_rows(min_row=4, values_only=True)).index(r), 5).value) if False else 0.0
            if cod:
                try:
                    precio = n(r[4]) if len(r) > 4 else 0.0
                except Exception:
                    precio = 0.0
                precios[cod] = precio

    # ── Ventas sheet
    ws = wb["Ventas"] if "Ventas" in wb.sheetnames else wb.active
    raw = []
    _empty_streak = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Saltar filas vacías rápidamente (Excel suele tener miles de filas formateadas vacías)
        if row[0] is None and (len(row) <= 9 or row[9] is None):
            _empty_streak += 1
            if _empty_streak > 20:
                break
            continue
        _empty_streak = 0
        if not row or len(row) < 11:
            continue
        REMITO    = s(row[0])
        VENDEDOR  = s(row[1])
        # FECHA
        fecha_raw = row[2]
        if isinstance(fecha_raw, datetime):
            FECHA = fecha_raw.strftime("%Y-%m-%d")
        else:
            FECHA = s(fecha_raw)
        LINEA_raw = s(row[3])
        COD       = s(row[4])
        # Índices reales verificados contra el archivo (no hay columna TIPO separada)
        CANTIDAD  = n(row[5])
        CLIENTE   = s(row[6])
        com_raw   = n(row[7])
        COMISION  = com_raw if 0 < com_raw <= 1 else 0.30
        PRECIO    = n(row[8])
        SUBTOTAL  = n(row[9])
        MES       = s(row[10])
        COMENTARIOS = s(row[11]) if len(row) > 11 else ""
        CMV       = n(row[12]) if len(row) > 12 else 0.0
        MARGEN    = n(row[13]) if len(row) > 13 else 0.0
        MARCA     = s(row[14]) if len(row) > 14 else ""
        TIPO_CLI  = s(row[15]) if len(row) > 15 else ""
        ZONA      = s(row[16]) if len(row) > 16 else ""
        LOCALIDAD = s(row[17]) if len(row) > 17 else ""
        PROVINCIA = s(row[18]) if len(row) > 18 else ""
        _desc_raw = s(row[19]) if len(row) > 19 else ""
        DESC      = "" if _desc_raw.startswith('DESCRIPCION') else _desc_raw

        # Filters
        if SUBTOTAL <= 0:
            continue
        if not re.match(r'^\d{4}-\d{2}$', MES):
            continue
        if not COD or COD.startswith('COD'):
            continue

        # LINEA derive
        if LINEA_raw in ('MONCAR', 'RMH'):
            LINEA = LINEA_raw
        else:
            LINEA = 'RMH' if COD.startswith('RM') else 'MONCAR'

        raw.append({
            'REMITO': REMITO, 'VENDEDOR': VENDEDOR, 'FECHA': FECHA,
            'LINEA': LINEA, 'COD': COD,
            'CANTIDAD': CANTIDAD, 'CLIENTE': CLIENTE, 'COMISION': COMISION,
            'PRECIO': PRECIO, 'SUBTOTAL': SUBTOTAL, 'MES': MES,
            'COMENTARIOS': COMENTARIOS, 'CMV': CMV, 'MARGEN': MARGEN,
            'MARCA': MARCA, 'TIPO_CLI': TIPO_CLI, 'ZONA': ZONA,
            'LOCALIDAD': LOCALIDAD, 'PROVINCIA': PROVINCIA, 'DESC': DESC,
        })

    months     = sorted(set(r['MES'] for r in raw))
    vendedores = sorted(set(r['VENDEDOR'] for r in raw if r['VENDEDOR']))
    print(f"  -> {len(raw)} filas válidas, {len(months)} meses")
    return raw, months, vendedores, precios

# ─── read STOCK ──────────────────────────────────────────────────────────────
def read_stock_mc(path):
    print("Leyendo STOCK MONCAR...")
    wb = load_wb(path, data_only=True)

    # Leer valores base de Productos como respaldo (cuando Stock aun no fue calculado por Excel)
    base = {}
    if "Productos" in wb.sheetnames:
        for row in wb["Productos"].iter_rows(min_row=4, values_only=True):
            if row[0]:
                cod = str(row[0]).strip()
                base[cod] = {
                    'NO_EMB': n(row[3]), 'EMB_DEP': n(row[4]),
                    'PREM_DEP': n(row[5]), 'JUANI': n(row[6]),
                }

    ws = wb["Stock"] if "Stock" in wb.sheetnames else wb.active
    rows = []
    sin_cache = 0
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row: continue
        COD = s(row[0]) if len(row) > 0 else ""
        if not COD or not (COD.startswith('CC') or COD.startswith('CF') or COD.startswith('CE')):
            continue
        # Si los valores son None (formulas no calculadas aun por Excel) usar base de Productos
        no_emb  = row[2] if len(row) > 2 and row[2] is not None else None
        emb_dep = row[3] if len(row) > 3 and row[3] is not None else None
        if no_emb is None and emb_dep is None and COD in base:
            sin_cache += 1
            b = base[COD]
            no_emb_v  = b['NO_EMB'];  emb_dep_v = b['EMB_DEP']
            prem_dep_v = b['PREM_DEP']; juani_v  = b['JUANI']
            vendible_v = emb_dep_v + prem_dep_v + juani_v
            total_v    = no_emb_v + vendible_v
        else:
            no_emb_v   = n(no_emb);   emb_dep_v  = n(emb_dep)
            prem_dep_v = n(row[4]) if len(row) > 4 else 0
            juani_v    = n(row[5]) if len(row) > 5 else 0
            vendible_v = n(row[6]) if len(row) > 6 else 0
            total_v    = n(row[7]) if len(row) > 7 else 0
        rows.append({
            'COD': COD,
            'MARCA':    s(row[1]) if len(row) > 1 else "",
            'NO_EMB':   no_emb_v,
            'EMB_DEP':  emb_dep_v,
            'PREM_DEP': prem_dep_v,
            'JUANI':    juani_v,
            'VENDIBLE': vendible_v,
            'TOTAL':    total_v,
            'PREST':    n(row[8]) if len(row) > 8 else 0,
            'DEBO':     n(row[9]) if len(row) > 9 else 0,
        })
    if sin_cache:
        print(f"  -> AVISO: {sin_cache} productos sin cache (aun no guardado en Excel) — usando valores base de Productos")
    print(f"  -> {len(rows)} productos MONCAR")
    return rows

def read_stock_rmh(path):
    print("Leyendo STOCK RMH...")
    wb = load_wb(path, data_only=True)
    ws = wb["Stock"] if "Stock" in wb.sheetnames else wb.active
    rows = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row: continue
        COD = s(row[0]) if len(row) > 0 else ""
        if not COD or not COD.startswith('RM'):
            continue
        rows.append({
            'COD':     COD,
            'DESC':    s(row[1]) if len(row) > 1 else "",
            'NO_EMB':  n(row[2]) if len(row) > 2 else 0,
            'EMB_DEP': n(row[3]) if len(row) > 3 else 0,
            'CABA':    n(row[4]) if len(row) > 4 else 0,
            'MAXI':    n(row[5]) if len(row) > 5 else 0,
            'VENDIBLE':n(row[6]) if len(row) > 6 else 0,
            'TOTAL':   n(row[7]) if len(row) > 7 else 0,
            'PREST':   n(row[8]) if len(row) > 8 else 0,
            'DEBO':    n(row[9]) if len(row) > 9 else 0,
        })
    print(f"  -> {len(rows)} productos RMH")
    return rows

# ─── main ────────────────────────────────────────────────────────────────────
def main():
    raw, months, vendedores, precios = read_ventas(VENTAS_PATH)

    stock_mc_rows = []
    if os.path.exists(MC_PATH):
        try:
            stock_mc_rows = read_stock_mc(MC_PATH)
        except PermissionError as e:
            print(f"  ERROR: {e}")
        except Exception as e:
            print(f"  ADVERTENCIA stock MONCAR: {e}")
    else:
        print("  STOCK_MONCAR.xlsx no encontrado — stock MC vacío")

    stock_rmh_rows = []
    if os.path.exists(RMH_PATH):
        try:
            stock_rmh_rows = read_stock_rmh(RMH_PATH)
        except Exception as e:
            print(f"  ADVERTENCIA stock RMH: {e}")
    else:
        print("  STOCK_RMH.xlsx no encontrado — stock RMH vacío")

    data = {
        'ventas': {
            'raw': raw,
            'months': months,
            'vendedores': vendedores,
        },
        'stock_mc':  {'rows': stock_mc_rows},
        'stock_rmh': {'rows': stock_rmh_rows},
        'precios':   precios,
        'generated_at': datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

    json_str = json.dumps(data, ensure_ascii=False)
    html = TEMPLATE.replace('/*INJECT_DATA*/', f'window.DATA = {json_str};')

    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"✓ Dashboard generado: {OUT_PATH}")
    if '--no-browser' not in sys.argv:
        print("✓ Abriendo en el navegador...")
        webbrowser.open('file:///' + OUT_PATH.replace('\\', '/'))

# ─── HTML TEMPLATE ────────────────────────────────────────────────────────────
TEMPLATE = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MONCAR + RMH — Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0a0c10;--surface:#111419;--surface2:#181c23;
  --border:#1f2430;--border2:#2a3040;
  --text:#e8eaf0;--muted:#6b7280;
  --accent:#00d4aa;--accent2:#3b82f6;--accent3:#f59e0b;--accent4:#ef4444;
  --green:#22c55e;--red:#f43f5e;
}
body.light{
  --bg:#f0f3f9;--surface:#ffffff;--surface2:#f4f7fc;
  --border:#dde3ef;--border2:#c5cfe3;
  --text:#1a2332;--muted:#5a6a80;
  --accent:#00957a;--accent2:#1d4ed8;--accent3:#b45309;--accent4:#b91c1c;
  --green:#15803d;--red:#dc2626;
}
body.light .card{box-shadow:0 1px 6px rgba(0,0,0,.08)}
body.light nav{box-shadow:0 1px 4px rgba(0,0,0,.1)}
body.light select{color:var(--text)}
body{background:var(--bg);color:var(--text);font-family:'Space Grotesk',sans-serif;min-height:100vh}
/* NAV */
nav{position:sticky;top:0;z-index:100;background:var(--surface);border-bottom:1px solid var(--border);
    display:flex;align-items:center;padding:0 24px;height:56px;gap:8px}
.logo{font-weight:700;font-size:18px;color:var(--accent);white-space:nowrap;margin-right:16px;letter-spacing:-0.5px}
.logo span{color:var(--accent2)}
.tabs{display:flex;gap:2px;flex:1}
.tab{padding:6px 14px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;
     color:var(--muted);border:none;background:none;transition:all .2s}
.tab:hover{color:var(--text);background:var(--border)}
.tab.active{color:var(--accent);background:rgba(0,212,170,.1)}
.nav-right{display:flex;align-items:center;gap:12px;margin-left:auto}
select{background:var(--surface2);color:var(--text);border:1px solid var(--border2);
       border-radius:6px;padding:5px 10px;font-size:13px;cursor:pointer}
.badge{background:var(--surface2);border:1px solid var(--border2);border-radius:6px;
       padding:5px 12px;font-size:12px;color:var(--muted);white-space:nowrap;font-family:'DM Mono',monospace}
.theme-btn{background:none;border:1px solid var(--border2);border-radius:6px;color:var(--muted);
  cursor:pointer;padding:5px 8px;font-size:17px;line-height:1;font-family:inherit;transition:all .2s}
.theme-btn:hover{color:var(--text);border-color:var(--accent)}
.pl-pct{font-size:10px;color:var(--muted);margin-left:7px;opacity:.75}
/* LAYOUT */
.page{display:none;padding:24px;max-width:1600px;margin:0 auto}
.page.active{display:block}
h2{font-size:22px;font-weight:700;margin-bottom:20px;color:var(--text)}
h3{font-size:15px;font-weight:600;margin-bottom:14px;color:var(--text)}
/* KPI */
.kpi-row{display:grid;gap:12px;margin-bottom:16px}
.kpi-row.cols4{grid-template-columns:repeat(4,1fr)}
.kpi-row.cols2{grid-template-columns:repeat(2,1fr)}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px}
.kpi-label{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);margin-bottom:8px}
.kpi-val{font-size:26px;font-weight:700;font-family:'DM Mono',monospace;line-height:1}
.kpi-sub{font-size:12px;color:var(--muted);margin-top:6px}
.kpi.green .kpi-val{color:var(--green)}
.kpi.red   .kpi-val{color:var(--red)}
.kpi.amber .kpi-val{color:var(--accent3)}
.kpi.blue  .kpi-val{color:var(--accent2)}
.kpi.accent .kpi-val{color:var(--accent)}
/* GRID */
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.grid-2.left-wide{grid-template-columns:2fr 1fr}
.card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:20px}
.card.full{grid-column:1/-1}
/* CHART containers */
.chart-wrap{position:relative}
.chart-wrap canvas{max-width:100%}
/* PILLS */
.pills{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px}
.pill{padding:5px 12px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;
      border:1px solid var(--border2);color:var(--muted);background:var(--surface2);transition:all .2s}
.pill:hover{color:var(--text)}
.pill.active{background:rgba(0,212,170,.15);border-color:var(--accent);color:var(--accent)}
.pill.active.blue{background:rgba(59,130,246,.15);border-color:var(--accent2);color:var(--accent2)}
/* TABLE */
.tbl-wrap{overflow-x:auto;max-height:420px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{position:sticky;top:0;background:var(--surface2);color:var(--muted);
         font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;
         padding:8px 12px;text-align:left;border-bottom:1px solid var(--border2)}
tbody tr{border-bottom:1px solid var(--border)}
tbody tr:hover{background:var(--surface2)}
tbody td{padding:8px 12px;color:var(--text)}
td.num{font-family:'DM Mono',monospace;text-align:center}
th.num{text-align:center}
td.cod{font-family:'DM Mono',monospace;font-size:12px;color:var(--accent)}
.badge-linea{display:inline-block;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:700}
.badge-mc{background:rgba(0,212,170,.15);color:var(--accent)}
.badge-rmh{background:rgba(59,130,246,.15);color:var(--accent2)}
.mini-bar-wrap{display:flex;align-items:center;gap:8px}
.mini-bar{height:6px;border-radius:3px;background:var(--accent);min-width:2px}
/* ROW TINTS */
tr.row-danger td{background:rgba(244,63,94,.06)}
tr.row-warn td{background:rgba(245,158,11,.06)}
/* STOCK TOGGLE */
.stock-toggle{display:flex;gap:8px;margin-bottom:16px}
.stk-btn{padding:7px 20px;border-radius:8px;cursor:pointer;font-weight:600;font-size:13px;
         border:1px solid var(--border2);color:var(--muted);background:var(--surface2);transition:all .2s}
.stk-btn.active{background:rgba(0,212,170,.15);border-color:var(--accent);color:var(--accent)}
/* EMPTY */
.empty{text-align:center;padding:60px 20px;color:var(--muted);font-size:14px}
/* WATERFALL */
.wf-pos{fill:var(--green)}
.wf-neg{fill:var(--red)}
/* NOTE */
.note{font-size:11px;color:var(--muted);margin-top:10px;padding-top:10px;border-top:1px solid var(--border)}
/* PERIOD SELECTOR */
.period-bar{display:flex;align-items:center;gap:12px;margin-bottom:20px;flex-wrap:wrap}
.period-tabs{display:flex;gap:2px;background:var(--surface2);border-radius:8px;padding:3px}
.period-tab{padding:5px 14px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;
  border:none;background:transparent;color:var(--muted);font-family:inherit;transition:all .15s;letter-spacing:.03em}
.period-tab.active{background:var(--surface);color:var(--text);box-shadow:0 1px 4px rgba(0,0,0,.5)}
.period-nav{display:flex;align-items:center;gap:6px}
.period-arrow{background:none;border:none;color:var(--muted);cursor:pointer;font-size:22px;
  padding:2px 6px;line-height:1;font-family:inherit;transition:color .15s}
.period-arrow:hover{color:var(--text)}
.period-nav-label{font-size:15px;font-weight:600;color:var(--text);min-width:130px;text-align:center}
.compare-wrap{margin-left:auto;display:flex;align-items:center;gap:8px;font-size:12px;color:var(--muted)}
.compare-select{background:var(--surface2);border:1px solid var(--border2);color:var(--text);
  font-family:inherit;font-size:12px;padding:5px 10px;border-radius:6px;cursor:pointer;outline:none}
.custom-range{display:none;align-items:center;gap:8px;margin-bottom:16px;flex-wrap:wrap}
.custom-range.show{display:flex}
.date-inp{background:var(--surface2);border:1px solid var(--border2);color:var(--text);
  font-family:inherit;font-size:12px;padding:5px 10px;border-radius:6px;outline:none}
.kpi.highlight{background:rgba(0,212,170,.09);border-color:rgba(0,212,170,.35)}
.kpi.highlight .kpi-val{color:var(--accent)}
.kpi-change{font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;
  margin-left:6px;vertical-align:middle;display:inline-block}
.kpi-change.pos{background:rgba(34,197,94,.15);color:#22c55e}
.kpi-change.neg{background:rgba(244,63,94,.15);color:#f43f5e}
/* ANALYTICS CONTROLS BAR */
#analytics-controls{background:var(--surface);border-bottom:1px solid var(--border);padding:10px 24px 10px}
#analytics-controls .period-bar{margin-bottom:0}
#analytics-controls .custom-range{margin-top:8px;margin-bottom:0}
/* SCROLLBAR */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--bg)}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
/* responsive */
@media(max-width:900px){
  .kpi-row.cols4{grid-template-columns:repeat(2,1fr)}
  .grid-2,.grid-2.left-wide{grid-template-columns:1fr}
}
</style>
</head>
<body>

<!-- NAV -->
<nav>
  <div class="logo">MONCAR<span>+RMH</span></div>
  <div class="tabs">
    <button class="tab active" onclick="showPage('overview')">Overview</button>
    <button class="tab" onclick="showPage('productos')">Productos</button>
    <button class="tab" onclick="showPage('marcas')">Marcas</button>
    <button class="tab" onclick="showPage('clientes')">Clientes</button>
    <button class="tab" onclick="showPage('stock')">Stock</button>
    <button class="tab" onclick="showPage('pl')">P&amp;L</button>
  </div>
  <div class="nav-right">
    <select id="sel-vendedor" onchange="onVendedorChange()">
      <option value="ALL">Todos los vendedores</option>
    </select>
    <div class="badge" id="badge-info">Cargando...</div>
    <button class="theme-btn" id="theme-toggle" onclick="toggleTheme()" title="Cambiar tema">☀</button>
  </div>
</nav>

<!-- ── Controles compartidos: período + línea (visible en Overview, Productos, Marcas, Clientes) ── -->
<div id="analytics-controls">
  <div class="period-bar">
    <div class="period-tabs">
      <button class="period-tab active" data-mode="mes">Mes</button>
      <button class="period-tab" data-mode="3m">3 Meses</button>
      <button class="period-tab" data-mode="6m">6 Meses</button>
      <button class="period-tab" data-mode="año">Año</button>
      <button class="period-tab" data-mode="custom">Personalizado</button>
    </div>
    <div class="period-nav">
      <button class="period-arrow" id="btn-prev">&#8249;</button>
      <div class="period-nav-label" id="period-label">—</div>
      <button class="period-arrow" id="btn-next">&#8250;</button>
    </div>
    <div class="stock-toggle" style="margin:0 0 0 12px">
      <button class="stk-btn btn-linea btn-linea-all active" onclick="setLinea('ALL')">Ambas</button>
      <button class="stk-btn btn-linea btn-linea-moncar" onclick="setLinea('MONCAR')">MONCAR</button>
      <button class="stk-btn btn-linea btn-linea-rmh" onclick="setLinea('RMH')">RMH</button>
    </div>
    <div class="compare-wrap">
      <span>Comparar con:</span>
      <select class="compare-select" id="compare-sel">
        <option value="none">Sin comparación</option>
        <option value="prev">Período anterior</option>
        <option value="prev_year">Mismo período año anterior</option>
      </select>
    </div>
  </div>
  <div class="custom-range" id="custom-range">
    <span style="color:var(--muted);font-size:12px">Desde:</span>
    <input type="date" class="date-inp" id="custom-start">
    <span style="color:var(--muted);font-size:12px">Hasta:</span>
    <input type="date" class="date-inp" id="custom-end">
    <button class="stk-btn active" style="padding:5px 14px;font-size:12px" onclick="applyCustom()">Aplicar</button>
  </div>
</div>

<!-- PAGES -->
<div id="page-overview" class="page active">

  <!-- KPIs -->
  <div class="kpi-row cols4" id="ov-kpi1"></div>
  <div class="kpi-row cols4" id="ov-kpi2"></div>

  <!-- Gráficos -->
  <div class="grid-2">
    <div class="card">
      <div class="card-title"><span class="dot"></span>Facturación mensual · MONCAR vs RMH</div>
      <div class="chart-wrap" style="height:260px"><canvas id="ch-ov-fact"></canvas></div>
    </div>
    <div class="card">
      <div class="card-title"><span class="dot amber"></span>Unidades por mes</div>
      <div class="chart-wrap" style="height:260px"><canvas id="ch-ov-units"></canvas></div>
    </div>
  </div>

  <!-- Top 10 -->
  <div class="grid-2">
    <div class="card">
      <div class="card-title"><span class="dot blue"></span>Top 10 Productos · período seleccionado</div>
      <div class="tbl-wrap"><table id="tbl-ov-prod"><thead><tr>
        <th>COD</th><th>MARCA</th><th>LÍNEA</th><th class="num">FACTURACIÓN</th><th style="min-width:70px"></th>
      </tr></thead><tbody></tbody></table></div>
    </div>
    <div class="card">
      <div class="card-title"><span class="dot red"></span>Top 10 Clientes · período seleccionado</div>
      <div class="tbl-wrap"><table id="tbl-ov-cli"><thead><tr>
        <th>#</th><th>CLIENTE</th><th class="num">FACTURACIÓN</th><th class="num">UDS</th>
      </tr></thead><tbody></tbody></table></div>
    </div>
  </div>

</div>

<div id="page-productos" class="page">
  <h2>Productos</h2>

  <!-- Buscador de producto -->
  <div class="card" style="margin-bottom:16px;padding:14px 18px">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <div style="position:relative;flex:1;min-width:180px">
        <input id="prod-search" type="text" placeholder="Buscar por código o descripción..."
               autocomplete="off"
               style="width:100%;box-sizing:border-box;padding:8px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--fg);font-family:inherit;font-size:13px;outline:none"
               oninput="renderProdDropdown()"
               onfocus="renderProdDropdown()"
               onkeydown="if(event.key==='Enter')searchProd()">
        <div id="prod-dd" onmousedown="event.preventDefault()" style="display:none;position:absolute;top:38px;left:0;right:0;background:var(--surface);border:1px solid var(--border);border-radius:6px;max-height:200px;overflow-y:auto;z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.3)"></div>
      </div>
      <button class="stk-btn" onclick="searchProd()" style="white-space:nowrap">Buscar</button>
      <button class="stk-btn active" onclick="selectProd(null)" style="white-space:nowrap">Ver todos</button>
    </div>
    <div id="prod-badge" style="display:none;margin-top:10px;font-size:13px">
      Mostrando: <strong id="prod-badge-cod" style="color:var(--primary);font-family:DM Mono"></strong>
      <span id="prod-badge-desc" style="color:var(--muted)"></span>
    </div>
  </div>

  <!-- Vista detalle (1 producto seleccionado) -->
  <div id="prod-detail" style="display:none">
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px">
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Facturación</div>
        <div id="pd-rev" style="font-size:22px;font-weight:700;color:var(--primary);margin-top:4px">—</div>
      </div>
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Unidades</div>
        <div id="pd-uds" style="font-size:22px;font-weight:700;margin-top:4px">—</div>
      </div>
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Clientes</div>
        <div id="pd-cli" style="font-size:22px;font-weight:700;margin-top:4px">—</div>
      </div>
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Precio prom.</div>
        <div id="pd-prom" style="font-size:22px;font-weight:700;margin-top:4px">—</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="card">
        <h3>Evolución mensual (facturación + unidades)</h3>
        <div class="chart-wrap" style="height:250px"><canvas id="ch-pd-evol"></canvas></div>
      </div>
      <div class="card">
        <h3>Top clientes</h3>
        <div class="chart-wrap" style="height:250px"><canvas id="ch-pd-cli"></canvas></div>
      </div>
    </div>
    <div class="card" style="margin-top:16px">
      <h3>Detalle por cliente</h3>
      <div class="tbl-wrap"><table id="tbl-pd-cli"><thead><tr>
        <th>#</th><th>CLIENTE</th><th class="num">FACTURACIÓN</th><th class="num">UDS</th><th class="num">PRECIO PROM.</th>
      </tr></thead><tbody></tbody></table></div>
    </div>
  </div>

  <!-- Vista general (default: top 15 + evolución) -->
  <div id="prod-overview">
    <div class="grid-2">
      <div class="card">
        <h3>Top 15 por Facturación</h3>
        <div class="chart-wrap" style="height:340px"><canvas id="ch-prod-rev"></canvas></div>
      </div>
      <div class="card">
        <h3>Top 15 por Unidades</h3>
        <div class="chart-wrap" style="height:340px"><canvas id="ch-prod-units"></canvas></div>
      </div>
    </div>
    <div class="card" style="margin-top:16px">
      <h3>Evolución Top 8 Productos</h3>
      <div class="chart-wrap" style="height:280px"><canvas id="ch-prod-evol"></canvas></div>
    </div>
  </div>

  <!-- Tabla resumen (siempre visible, clickeable) -->
  <div class="card" style="margin-top:16px">
    <h3>Resumen de todos los productos del período</h3>
    <div class="tbl-wrap"><table id="tbl-prod"><thead><tr>
      <th>#</th><th>CÓDIGO</th><th>DESCRIPCIÓN</th><th>LÍNEA</th>
      <th class="num">FACTURACIÓN</th><th class="num">UDS</th>
      <th class="num">CLIENTES</th><th class="num">PRECIO PROM.</th><th class="num">% TOTAL</th>
    </tr></thead><tbody></tbody></table></div>
  </div>
</div>

<div id="page-marcas" class="page">
  <h2>Marcas</h2>

  <!-- Buscador de marca -->
  <div class="card" style="margin-bottom:16px;padding:14px 18px">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <div style="position:relative;flex:1;min-width:180px">
        <input id="marc-search" type="text" placeholder="Buscar marca..."
               autocomplete="off"
               style="width:100%;box-sizing:border-box;padding:8px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--fg);font-family:inherit;font-size:13px;outline:none"
               oninput="renderMarcDropdown()"
               onfocus="renderMarcDropdown()"
               onkeydown="if(event.key==='Enter')searchMarc()">
        <div id="marc-dd" onmousedown="event.preventDefault()" style="display:none;position:absolute;top:38px;left:0;right:0;background:var(--surface);border:1px solid var(--border);border-radius:6px;max-height:200px;overflow-y:auto;z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.3)"></div>
      </div>
      <button class="stk-btn" onclick="searchMarc()" style="white-space:nowrap">Buscar</button>
      <button class="stk-btn active" onclick="selectMarc(null)" style="white-space:nowrap">Ver todas</button>
    </div>
    <div id="marc-badge" style="display:none;margin-top:10px;font-size:13px">
      Mostrando: <strong id="marc-badge-nom" style="color:var(--primary)"></strong>
    </div>
  </div>

  <!-- Vista detalle (1 marca seleccionada) -->
  <div id="marc-detail" style="display:none">
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px">
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Facturación</div>
        <div id="md-rev" style="font-size:22px;font-weight:700;color:var(--primary);margin-top:4px">—</div>
      </div>
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Unidades</div>
        <div id="md-uds" style="font-size:22px;font-weight:700;margin-top:4px">—</div>
      </div>
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Productos distintos</div>
        <div id="md-prods" style="font-size:22px;font-weight:700;margin-top:4px">—</div>
      </div>
      <div class="card" style="padding:14px 16px">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)">Clientes</div>
        <div id="md-clis" style="font-size:22px;font-weight:700;margin-top:4px">—</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="card">
        <h3>Evolución mensual</h3>
        <div class="chart-wrap" style="height:250px"><canvas id="ch-md-evol"></canvas></div>
      </div>
      <div class="card">
        <h3>Top productos de esta marca</h3>
        <div class="chart-wrap" style="height:250px"><canvas id="ch-md-prods"></canvas></div>
      </div>
    </div>
    <div class="card" style="margin-top:16px">
      <h3>Clientes que compraron esta marca</h3>
      <div class="tbl-wrap"><table id="tbl-md-cli"><thead><tr>
        <th>#</th><th>CLIENTE</th><th class="num">FACTURACIÓN</th><th class="num">UDS</th><th class="num">% DE LA MARCA</th>
      </tr></thead><tbody></tbody></table></div>
    </div>
  </div>

  <!-- Vista general (default) -->
  <div id="marc-overview">
    <div class="grid-2">
      <div class="card">
        <h3>Facturación por Marca</h3>
        <div class="chart-wrap" style="height:320px"><canvas id="ch-marc-rev"></canvas></div>
      </div>
      <div class="card">
        <h3>Unidades por Marca</h3>
        <div class="chart-wrap" style="height:320px"><canvas id="ch-marc-units"></canvas></div>
      </div>
    </div>
    <div class="card" style="margin-top:16px">
      <h3>Evolución Top 6 Marcas</h3>
      <div class="chart-wrap" style="height:280px"><canvas id="ch-marc-evol"></canvas></div>
    </div>
  </div>

  <!-- Tabla resumen (siempre visible, clickeable) -->
  <div class="card" style="margin-top:16px">
    <h3>Resumen por Marca</h3>
    <div class="tbl-wrap"><table id="tbl-marc"><thead><tr>
      <th>#</th><th>MARCA</th><th class="num">FACTURACIÓN</th><th class="num">UDS</th><th class="num">MARGEN BRUTO</th><th class="num">% TOTAL</th>
    </tr></thead><tbody></tbody></table></div>
  </div>
</div>

<div id="page-clientes" class="page">
  <h2>Clientes</h2>
  <div class="grid-2">
    <div class="card">
      <h3>Facturación por Cliente (Top 15)</h3>
      <div class="chart-wrap" style="height:340px"><canvas id="ch-cli-rev"></canvas></div>
    </div>
    <div class="card">
      <h3>Unidades por Cliente (Top 15)</h3>
      <div class="chart-wrap" style="height:340px"><canvas id="ch-cli-units"></canvas></div>
    </div>
  </div>
  <div class="card">
    <h3>Detalle de Clientes</h3>
    <div class="tbl-wrap"><table id="tbl-cli"><thead><tr>
      <th>#</th><th>CLIENTE</th><th class="num">FACTURACIÓN</th><th class="num">UDS</th><th class="num">% TOTAL</th><th>PROVINCIA</th>
    </tr></thead><tbody></tbody></table></div>
  </div>
</div>

<div id="page-stock" class="page">
  <h2>Stock</h2>
  <div class="stock-toggle">
    <button class="stk-btn active" id="btn-mc" onclick="setStock('mc')">MONCAR</button>
    <button class="stk-btn" id="btn-rmh" onclick="setStock('rmh')">RMH</button>
  </div>
  <div id="stock-body"></div>
</div>

<div id="page-pl" class="page">
  <h2>P&amp;L — Estado de Resultados</h2>
  <div class="pills" id="pills-pl"></div>
  <div class="grid-2">
    <div class="card" id="pl-table-wrap">
      <h3>Estado de Resultados</h3>
      <div id="pl-table"></div>
    </div>
    <div class="card">
      <h3>Waterfall P&amp;L</h3>
      <div class="chart-wrap" style="height:320px"><canvas id="ch-pl-wf"></canvas></div>
    </div>
  </div>
  <div class="card" style="margin-top:16px">
    <h3>Facturación / CMV / Comisión mensual</h3>
    <div class="chart-wrap" style="height:260px"><canvas id="ch-pl-stack"></canvas></div>
  </div>
  <div class="card" style="margin-top:16px">
    <h3>Margen Bruto % mensual</h3>
    <div class="chart-wrap" style="height:220px"><canvas id="ch-pl-mg"></canvas></div>
  </div>
</div>

<script>
/*INJECT_DATA*/

// ── Data aliases ─────────────────────────────────────────────────────────────
const V       = window.DATA.ventas;
const MC_ROWS = window.DATA.stock_mc.rows;
const RMH_ROWS= window.DATA.stock_rmh.rows;
const LP      = window.DATA.precios;

// ── State ────────────────────────────────────────────────────────────────────
let curPage = 'overview';
const F = { vendedor:'ALL', linea:'ALL', pl:'ALL', stock:'mc' };

// ── Chart registry ────────────────────────────────────────────────────────────
const charts = {};
function mkChart(id, cfg) {
  if(charts[id]) { charts[id].destroy(); delete charts[id]; }
  const el = document.getElementById(id);
  if(!el) return;
  charts[id] = new Chart(el, cfg);
}

// ── Palette ──────────────────────────────────────────────────────────────────
const PALETTE = [
  '#00d4aa','#3b82f6','#f59e0b','#ef4444','#a78bfa','#34d399',
  '#f97316','#06b6d4','#e879f9','#84cc16','#fb7185','#38bdf8',
  '#fbbf24','#4ade80','#c084fc','#f472b6'
];

// ── Formatters ────────────────────────────────────────────────────────────────
function fmt(v) {
  return '$' + Math.round(v).toLocaleString('es-AR');
}
function fmtN(v) {
  return Math.round(v).toLocaleString('es-AR');
}
function pct(v) {
  return (v*100).toFixed(1) + '%';
}
function pctDiff(v) {
  if(!isFinite(v)) return '—';
  const s = v>=0 ? '+' : '';
  return s + (v*100).toFixed(1) + '%';
}

// ── Filter helpers ────────────────────────────────────────────────────────────
function filterVentas(page) {
  let rows = V.raw;
  if(F.vendedor !== 'ALL') rows = rows.filter(r => r.VENDEDOR === F.vendedor);
  const mes = F[page];
  if(mes && mes !== 'ALL') rows = rows.filter(r => r.MES === mes);
  return rows;
}

function getMonthly(rows) {
  const m = {};
  rows.forEach(r => {
    if(!m[r.MES]) m[r.MES] = {revenue:0,cmv:0,units:0,comision:0,margen:0,mc_rev:0,rmh_rev:0};
    m[r.MES].revenue  += r.SUBTOTAL;
    m[r.MES].cmv      += r.CMV;
    m[r.MES].units    += r.CANTIDAD;
    m[r.MES].comision += r.SUBTOTAL * r.COMISION;
    m[r.MES].margen   += r.MARGEN;
    if(r.LINEA==='MONCAR') m[r.MES].mc_rev += r.SUBTOTAL;
    else m[r.MES].rmh_rev += r.SUBTOTAL;
  });
  return m;
}

// ── Chart defaults ────────────────────────────────────────────────────────────
Chart.defaults.font.family = "'Space Grotesk', sans-serif";
Chart.defaults.font.size = 12;

const cssVar = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const GC = () => cssVar('--border');
const TC = () => cssVar('--muted');

let tooltipStyle = {};

function applyChartTheme() {
  const light = document.body.classList.contains('light');
  Chart.defaults.color = TC();
  Chart.defaults.borderColor = GC();
  tooltipStyle = {
    backgroundColor: light ? '#ffffff' : '#181c23',
    borderColor: GC(),
    borderWidth: 1,
    titleColor: cssVar('--text'),
    bodyColor: TC(),
    padding: 10,
    cornerRadius: 8,
    displayColors: true,
    boxPadding: 4
  };
}

function toggleTheme() {
  const isLight = document.body.classList.toggle('light');
  document.getElementById('theme-toggle').textContent = isLight ? '🌙' : '☀';
  try { localStorage.setItem('moncar_theme', isLight ? 'light' : 'dark'); } catch(e) {}
  applyChartTheme();
  renderPage(curPage);
}

function barOpts(horizontal=false, stacked=false) {
  return {
    type: horizontal ? 'bar' : 'bar',
    options:{
      indexAxis: horizontal ? 'y' : 'x',
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:stacked, labels:{color:TC(),boxWidth:12,padding:16}}, tooltip:{...tooltipStyle} },
      scales:{
        x:{ stacked, grid:{color:GC()}, ticks:{color:TC()} },
        y:{ stacked, grid:{color:horizontal?'#1f2430':'#1f2430'}, ticks:{color:TC(),maxTicksLimit:12} }
      }
    }
  };
}

function lineOpts(filled=false) {
  return {
    type:'line',
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:true, labels:{color:TC(),boxWidth:12,padding:16}}, tooltip:{...tooltipStyle} },
      scales:{
        x:{ grid:{color:GC()}, ticks:{color:TC()} },
        y:{ grid:{color:GC()}, ticks:{color:TC()} }
      },
      elements:{ line:{tension:0.4}, point:{radius:3,hoverRadius:5} }
    }
  };
}

// ── Pills builder ─────────────────────────────────────────────────────────────
function mkPills(containerId, page) {
  const months = (function(){
    let rows = V.raw;
    if(F.vendedor !== 'ALL') rows = rows.filter(r => r.VENDEDOR === F.vendedor);
    return [...new Set(rows.map(r=>r.MES))].sort();
  })();
  const el = document.getElementById(containerId);
  if(!el) return;
  let html = `<span class="pill${F[page]==='ALL'?' active':''}" onclick="setPill('${page}','ALL')">Todos</span>`;
  months.forEach(m => {
    html += `<span class="pill${F[page]===m?' active':''}" onclick="setPill('${page}','${m}')">${m}</span>`;
  });
  el.innerHTML = html;
}

function setPill(page, val) {
  F[page] = val;
  renderPage(page);
}

// ─────────────────────────────────────────────────────────────────────────────
// PERIOD STATE
// ─────────────────────────────────────────────────────────────────────────────
const PERIOD = { mode:'mes', current:null, compare:'none', cStart:null, cEnd:null };

function initPeriod() {
  if(V.months.length>0) PERIOD.current = V.months[V.months.length-1];
  updatePeriodUI();
}

// Navega un mes hacia atrás (usado por múltiples modos)
function prevMes(m) {
  const[y,mo]=m.split('-').map(Number);
  return mo===1?`${y-1}-12`:`${y}-${String(mo-1).padStart(2,'0')}`;
}
function nextMes(m) {
  const[y,mo]=m.split('-').map(Number);
  return mo===12?`${y+1}-01`:`${y}-${String(mo+1).padStart(2,'0')}`;
}

// Devuelve array de N meses terminando en endMonth (inclusive)
function lastNMonths(endMonth, n) {
  const ms=[]; let c=endMonth;
  for(let i=0;i<n;i++){ms.unshift(c);c=prevMes(c);}
  return ms;
}

function periodLabel(mode, cur) {
  if(!cur) return 'Todos';
  const MESES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  if(mode==='mes')  { const[y,m]=cur.split('-'); return MESES[+m-1]+' '+y; }
  if(mode==='año')  return cur;
  if(mode==='semana') { const[y,w]=cur.split('-W'); return `Semana ${w}, ${y}`; }
  if(mode==='3m') { const ms=lastNMonths(cur,3); return `${periodLabel('mes',ms[0])} – ${periodLabel('mes',ms[2])}`; }
  if(mode==='6m') { const ms=lastNMonths(cur,6); return `${periodLabel('mes',ms[0])} – ${periodLabel('mes',ms[5])}`; }
  if(mode==='custom') return PERIOD.cStart&&PERIOD.cEnd ? `${PERIOD.cStart} → ${PERIOD.cEnd}` : 'Personalizado';
  return cur;
}

function prevPer(mode,cur) {
  if(mode==='mes'||mode==='3m'||mode==='6m') return prevMes(cur);
  if(mode==='año')  return String(+cur-1);
  if(mode==='semana') { const[y,w]=cur.split('-W').map(Number); return w===1?`${y-1}-W52`:`${y}-W${String(w-1).padStart(2,'0')}`; }
  return cur;
}
function nextPer(mode,cur) {
  if(mode==='mes'||mode==='3m'||mode==='6m') return nextMes(cur);
  if(mode==='año')  return String(+cur+1);
  if(mode==='semana') { const[y,w]=cur.split('-W').map(Number); return w>=52?`${y+1}-W01`:`${y}-W${String(w+1).padStart(2,'0')}`; }
  return cur;
}

function rowsForPeriod(mode, cur, cStart, cEnd) {
  const base = F.vendedor==='ALL' ? V.raw : V.raw.filter(r=>r.VENDEDOR===F.vendedor);
  if(!cur && mode!=='custom') return base;
  if(mode==='mes')    return base.filter(r=>r.MES===cur);
  if(mode==='año')    return base.filter(r=>r.MES&&r.MES.startsWith(cur));
  if(mode==='3m')     { const ms=lastNMonths(cur,3); return base.filter(r=>ms.includes(r.MES)); }
  if(mode==='6m')     { const ms=lastNMonths(cur,6); return base.filter(r=>ms.includes(r.MES)); }
  if(mode==='semana') {
    const[wy,wn]=cur.split('-W').map(Number);
    return base.filter(r=>{
      if(!r.FECHA)return false;
      const d=new Date(r.FECHA+'T00:00:00');
      const jan4=new Date(d.getFullYear(),0,4);
      const w=Math.ceil(((d-jan4)/86400000+jan4.getDay()+1)/7);
      return w===wn&&d.getFullYear()===wy;
    });
  }
  if(mode==='custom') return (!cStart||!cEnd)?base:base.filter(r=>r.FECHA&&r.FECHA>=cStart&&r.FECHA<=cEnd);
  return base;
}

// Filas del período inmediatamente anterior (siempre, independiente del dropdown comparar)
function prevPeriodRows() {
  if(!PERIOD.current||PERIOD.mode==='custom') return null;
  // Para 3m/6m: desplaza el fin del período N meses atrás
  const n = PERIOD.mode==='3m'?3 : PERIOD.mode==='6m'?6 : 1;
  let prevEnd = PERIOD.current;
  for(let i=0;i<n;i++) prevEnd = prevMes(prevEnd);
  return rowsForPeriod(PERIOD.mode, prevEnd, null, null);
}

function compareRows() {
  if(PERIOD.compare==='none'||!PERIOD.current) return null;
  if(PERIOD.compare==='prev') return prevPeriodRows();
  if(PERIOD.compare==='prev_year') {
    if(PERIOD.mode==='mes') { const[y,m]=PERIOD.current.split('-'); return rowsForPeriod('mes',`${+y-1}-${m}`,null,null); }
    if(PERIOD.mode==='año') return rowsForPeriod('año',String(+PERIOD.current-1),null,null);
    if(PERIOD.mode==='3m'||PERIOD.mode==='6m') {
      const[y,m]=PERIOD.current.split('-').map(Number);
      return rowsForPeriod(PERIOD.mode,`${y-1}-${String(m).padStart(2,'0')}`,null,null);
    }
  }
  return null;
}

function applyLineaFilter(rows) {
  return F.linea === 'ALL' ? rows : rows.filter(r => r.LINEA === F.linea);
}

function rowsFiltered() {
  return applyLineaFilter(rowsForPeriod(PERIOD.mode, PERIOD.current, PERIOD.cStart, PERIOD.cEnd));
}

function setLinea(v) {
  F.linea = v;
  document.querySelectorAll('.btn-linea').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.btn-linea-' + v.toLowerCase()).forEach(b => b.classList.add('active'));
  renderPage(curPage);
}

function updatePeriodUI() {
  document.getElementById('period-label').textContent = periodLabel(PERIOD.mode, PERIOD.current);
  document.querySelectorAll('.period-tab').forEach(t=>t.classList.toggle('active',t.dataset.mode===PERIOD.mode));
  const cr=document.getElementById('custom-range');
  if(cr) cr.classList.toggle('show',PERIOD.mode==='custom');
  renderPage(curPage);
}

function applyCustom() {
  PERIOD.cStart=document.getElementById('custom-start').value;
  PERIOD.cEnd=document.getElementById('custom-end').value;
  updatePeriodUI();
}

// Init period event listeners (called after DOM ready)
function bindPeriod() {
  document.querySelectorAll('.period-tab').forEach(t=>t.addEventListener('click',()=>{
    PERIOD.mode=t.dataset.mode;
    const last = V.months.length>0 ? V.months[V.months.length-1] : null;
    if(PERIOD.mode==='mes')  PERIOD.current = last;
    else if(PERIOD.mode==='3m'||PERIOD.mode==='6m') PERIOD.current = last;
    else if(PERIOD.mode==='año') { const yrs=[...new Set(V.months.map(m=>m.split('-')[0]))].sort(); PERIOD.current=yrs[yrs.length-1]||null; }
    else if(PERIOD.mode==='semana') PERIOD.current=null;
    else PERIOD.current=null;
    updatePeriodUI();
  }));
  document.getElementById('btn-prev').addEventListener('click',()=>{ if(PERIOD.mode!=='custom'&&PERIOD.current){PERIOD.current=prevPer(PERIOD.mode,PERIOD.current);updatePeriodUI();} });
  document.getElementById('btn-next').addEventListener('click',()=>{ if(PERIOD.mode!=='custom'&&PERIOD.current){PERIOD.current=nextPer(PERIOD.mode,PERIOD.current);updatePeriodUI();} });
  document.getElementById('compare-sel').addEventListener('change',function(){PERIOD.compare=this.value;renderPage(curPage);});
}

// ─────────────────────────────────────────────────────────────────────────────
// OVERVIEW
// ─────────────────────────────────────────────────────────────────────────────
function calcKPIs(rows) {
  const rev  = rows.reduce((a,r)=>a+r.SUBTOTAL,0);
  const cmv  = rows.reduce((a,r)=>a+r.CMV,0);
  const com  = rows.reduce((a,r)=>a+r.SUBTOTAL*r.COMISION,0);
  const uds  = rows.reduce((a,r)=>a+r.CANTIDAD,0);
  const ord  = new Set(rows.map(r=>r.REMITO).filter(Boolean)).size;
  const net  = rev - cmv - com;
  return { rev, cmv, com, uds, ord, net,
    margin: rev>0 ? net/rev : 0,
    ticket: ord>0 ? rev/ord : 0,
    upt:    ord>0&&uds>0 ? uds/ord : null };
}

function changeBadge(cur, prev) {
  if(!prev||prev===0) return '';
  const v=(cur-prev)/prev;
  const s=v>=0?'+':'';
  return `<span class="kpi-change ${v>=0?'pos':'neg'}">${s}${(v*100).toFixed(1)}%</span>`;
}

function renderOverview() {
  const rows = rowsFiltered();
  const cmpRaw = compareRows();
  const cmp  = cmpRaw ? applyLineaFilter(cmpRaw) : null;
  const cur  = calcKPIs(rows);
  const prev = cmp ? calcKPIs(cmp) : null;

  // KPI fila 1: Facturación | AOI Resultado Neto (highlight) | COGS | Margen Neto
  document.getElementById('ov-kpi1').innerHTML = [
    kpi('Facturación', fmt(cur.rev), `${cur.ord} pedidos · ${fmtN(cur.uds)} uds`, 'green', changeBadge(cur.rev,prev?.rev)),
    kpi('AOI — Resultado Neto', fmt(cur.net), `${pct(cur.margin)} del total`, 'highlight', changeBadge(cur.net,prev?.net)),
    kpi('COGS', fmt(cur.cmv), `${pct(cur.rev>0?cur.cmv/cur.rev:0)} de la facturación`, 'red', changeBadge(cur.cmv,prev?.cmv)),
    kpi('Margen Neto', pct(cur.margin), `Bruto: ${pct(cur.rev>0?(cur.rev-cur.cmv)/cur.rev:0)}`, 'blue', changeBadge(cur.margin,prev?.margin)),
  ].join('');

  // Evolución automática vs período anterior (independiente del dropdown "Comparar con")
  const prevAutoRaw = prevPeriodRows();
  const prevAutoR = prevAutoRaw ? applyLineaFilter(prevAutoRaw) : null;
  const prevAuto  = prevAutoR ? calcKPIs(prevAutoR) : null;
  const evoPct = prevAuto&&prevAuto.rev>0 ? (cur.rev-prevAuto.rev)/prevAuto.rev : null;
  const evoVal = evoPct!==null ? pctDiff(evoPct) : '—';
  const evoSub = evoPct!==null
    ? `${fmt(prevAuto.rev)} → ${fmt(cur.rev)}`
    : 'sin datos del período anterior';
  const evoColor = evoPct===null?'muted':evoPct>=0?'green':'red';

  // KPI fila 2: Órdenes | Ticket | Evolución vs ant. | UPT
  document.getElementById('ov-kpi2').innerHTML = [
    kpi('Órdenes', fmtN(cur.ord), cur.ord>0?`${fmtN(cur.uds)} uds totales`:'—', 'green', changeBadge(cur.ord,prev?.ord)),
    kpi('Ticket Promedio', fmt(cur.ticket), 'por orden', 'amber', changeBadge(cur.ticket,prev?.ticket)),
    kpi('Evolución vs per. ant.', evoVal, evoSub, evoColor, ''),
    kpi('UPT', cur.upt!==null ? cur.upt.toFixed(1) : '—', cur.upt!==null?'uds por pedido':'sin datos de unidades', 'blue', ''),
  ].join('');

  // Gráficos — siempre sobre todos los meses (contexto histórico)
  const allRowsBase = F.vendedor==='ALL' ? V.raw : V.raw.filter(r=>r.VENDEDOR===F.vendedor);
  const allRows = applyLineaFilter(allRowsBase);
  const monthly  = getMonthly(allRows);
  const months   = Object.keys(monthly).sort();

  mkChart('ch-ov-fact', {
    type:'bar',
    data:{ labels:months.map(shortMes), datasets:[
      {label:'MONCAR',data:months.map(m=>monthly[m].mc_rev), backgroundColor:'rgba(0,212,170,.75)',borderRadius:3},
      {label:'RMH',   data:months.map(m=>monthly[m].rmh_rev),backgroundColor:'rgba(59,130,246,.75)',borderRadius:3},
    ]},
    options:{ responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,labels:{color:TC(),boxWidth:12}},tooltip:{...tooltipStyle,callbacks:{label:c=>`${c.dataset.label}: ${fmt(c.raw)}`}}},
      scales:{x:{stacked:true,grid:{color:GC()},ticks:{color:TC()}},y:{stacked:true,grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}}} }
  });

  mkChart('ch-ov-units', {
    type:'bar',
    data:{ labels:months.map(shortMes), datasets:[
      {label:'Unidades',data:months.map(m=>monthly[m].units),backgroundColor:'rgba(245,158,11,.7)',borderRadius:3}
    ]},
    options:{ responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{...tooltipStyle}},
      scales:{x:{grid:{color:GC()},ticks:{color:TC()}},y:{grid:{color:GC()},ticks:{color:TC()}}} }
  });

  // Top 10 productos (período seleccionado)
  const prodMap={};
  rows.forEach(r=>{ if(!prodMap[r.COD])prodMap[r.COD]={cod:r.COD,marca:r.MARCA,linea:r.LINEA,rev:0,uds:0}; prodMap[r.COD].rev+=r.SUBTOTAL; prodMap[r.COD].uds+=r.CANTIDAD; });
  const top10p=Object.values(prodMap).sort((a,b)=>b.rev-a.rev).slice(0,10);
  const maxR=top10p[0]?.rev||1;
  document.querySelector('#tbl-ov-prod tbody').innerHTML=top10p.map(p=>`<tr>
    <td class="cod">${p.cod}</td><td>${p.marca||'—'}</td>
    <td><span class="badge-linea ${p.linea==='MONCAR'?'badge-mc':'badge-rmh'}">${p.linea}</span></td>
    <td class="num">${fmt(p.rev)}</td>
    <td><div class="mini-bar-wrap"><div class="mini-bar" style="width:${Math.round(p.rev/maxR*80)}px;background:${p.linea==='MONCAR'?'var(--accent)':'var(--accent2)'}"></div></div></td>
  </tr>`).join('');

  // Top 10 clientes (período seleccionado)
  const cliMap={};
  rows.forEach(r=>{ if(!cliMap[r.CLIENTE])cliMap[r.CLIENTE]={cli:r.CLIENTE,rev:0,uds:0}; cliMap[r.CLIENTE].rev+=r.SUBTOTAL; cliMap[r.CLIENTE].uds+=r.CANTIDAD; });
  const top10c=Object.values(cliMap).sort((a,b)=>b.rev-a.rev).slice(0,10);
  document.querySelector('#tbl-ov-cli tbody').innerHTML=top10c.map((c,i)=>`<tr>
    <td class="num" style="color:var(--muted)">${i+1}</td><td>${c.cli||'—'}</td>
    <td class="num">${fmt(c.rev)}</td><td class="num">${fmtN(c.uds)}</td>
  </tr>`).join('');
}

function kpi(label, val, sub, color, change='') {
  return `<div class="kpi ${color}"><div class="kpi-label">${label}</div><div class="kpi-val">${val}${change}</div><div class="kpi-sub">${sub}</div></div>`;
}
// alias para compatibilidad con otros render functions
function kpiCard(label,val,sub,color){return kpi(label,val,sub,color);}

// ─────────────────────────────────────────────────────────────────────────────
// PRODUCTOS
// ─────────────────────────────────────────────────────────────────────────────
let PROD_SEL = null;

function selectProd(cod) {
  PROD_SEL = cod;
  const inp = document.getElementById('prod-search');
  if(inp) inp.value = cod || '';
  const dd = document.getElementById('prod-dd');
  if(dd) dd.style.display = 'none';
  renderProductos();
}

function searchProd() {
  const q = (document.getElementById('prod-search').value||'').trim().toUpperCase();
  if(!q){ selectProd(null); return; }
  const rows = rowsFiltered();
  const codes = [...new Set(rows.map(r=>r.COD))];
  const exact = codes.find(c=>c.toUpperCase()===q);
  if(exact){ selectProd(exact); return; }
  const partial = codes.filter(c=>c.toUpperCase().includes(q));
  if(partial.length===1){ selectProd(partial[0]); return; }
  renderProdDropdown();
}

function renderProdDropdown() {
  const inp = document.getElementById('prod-search');
  const dd  = document.getElementById('prod-dd');
  if(!inp || !dd) return;
  const q = inp.value.trim().toUpperCase();
  const rows = rowsFiltered();
  const seen = {};
  rows.forEach(r=>{ if(!seen[r.COD]) seen[r.COD]={cod:r.COD, desc:r.DESC||''}; });
  const all = Object.values(seen).sort((a,b)=>a.cod.localeCompare(b.cod));
  const filtered = q ? all.filter(p=>p.cod.toUpperCase().includes(q)||(p.desc&&p.desc.toUpperCase().includes(q))) : all;
  if(!filtered.length){ dd.style.display='none'; return; }
  dd.innerHTML = filtered.slice(0,25).map(p=>
    `<div style="padding:7px 12px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--border)"
          onclick="selectProd(${JSON.stringify(p.cod)})">
       <strong style="font-family:DM Mono">${p.cod}</strong>
       ${p.desc ? `<span style="color:var(--muted);font-size:11px;margin-left:8px">${p.desc}</span>` : ''}
     </div>`
  ).join('');
  dd.style.display = '';
}

function _buildProdMap(rows) {
  const m = {};
  rows.forEach(r => {
    if(!m[r.COD]) m[r.COD]={cod:r.COD,desc:r.DESC||'',linea:r.LINEA,marca:r.MARCA||'',rev:0,uds:0,clis:new Set()};
    m[r.COD].rev += r.SUBTOTAL;
    m[r.COD].uds += r.CANTIDAD;
    m[r.COD].clis.add(r.CLIENTE);
  });
  return m;
}

function renderProductos() {
  const rows = rowsFiltered();
  const prodMap = _buildProdMap(rows);
  const sorted = Object.values(prodMap).sort((a,b)=>b.rev-a.rev);
  const totalRev = sorted.reduce((a,p)=>a+p.rev, 0);

  // Si producto seleccionado no tiene datos en el período, resetear
  if(PROD_SEL && !prodMap[PROD_SEL]) {
    PROD_SEL = null;
    const inp = document.getElementById('prod-search');
    if(inp) inp.value = '';
  }

  // Badge
  const badge = document.getElementById('prod-badge');
  if(badge) {
    if(PROD_SEL && prodMap[PROD_SEL]) {
      const p = prodMap[PROD_SEL];
      badge.style.display = '';
      document.getElementById('prod-badge-cod').textContent = PROD_SEL;
      document.getElementById('prod-badge-desc').textContent = p.desc ? ` — ${p.desc}` : '';
    } else {
      badge.style.display = 'none';
    }
  }

  // Toggle vistas
  const detailEl  = document.getElementById('prod-detail');
  const overviewEl = document.getElementById('prod-overview');
  if(PROD_SEL && prodMap[PROD_SEL]) {
    if(detailEl)  detailEl.style.display  = '';
    if(overviewEl) overviewEl.style.display = 'none';
    _renderProdDetail(rows, prodMap[PROD_SEL]);
  } else {
    if(detailEl)  detailEl.style.display  = 'none';
    if(overviewEl) overviewEl.style.display = '';
    _renderProdOverview(rows, sorted);
  }

  // Tabla resumen (siempre)
  _renderProdTable(sorted, totalRev);
}

function _renderProdDetail(rows, p) {
  const detRows = rows.filter(r=>r.COD===p.cod);

  // KPIs
  document.getElementById('pd-rev').textContent  = fmt(p.rev);
  document.getElementById('pd-uds').textContent  = p.uds.toLocaleString('es-AR');
  document.getElementById('pd-cli').textContent  = p.clis.size;
  const prom = p.uds > 0 ? p.rev / p.uds : 0;
  document.getElementById('pd-prom').textContent = prom > 0 ? fmt(prom) : '—';

  // Meses del período (todos, para que el eje sea completo)
  const months = [...new Set(rows.map(r=>r.MES))].sort();
  const mRev={}, mUds={};
  months.forEach(m=>{mRev[m]=0;mUds[m]=0;});
  detRows.forEach(r=>{ mRev[r.MES]=(mRev[r.MES]||0)+r.SUBTOTAL; mUds[r.MES]=(mUds[r.MES]||0)+r.CANTIDAD; });

  mkChart('ch-pd-evol',{
    type:'line',
    data:{
      labels:months.map(shortMes),
      datasets:[
        {label:'Facturación',data:months.map(m=>mRev[m]),borderColor:PALETTE[0],backgroundColor:'transparent',yAxisID:'yR',tension:0.4,borderWidth:2,pointRadius:4,pointBackgroundColor:PALETTE[0]},
        {label:'Unidades',  data:months.map(m=>mUds[m]),borderColor:PALETTE[2],backgroundColor:'transparent',yAxisID:'yU',tension:0.4,borderWidth:2,pointRadius:4,pointBackgroundColor:PALETTE[2]}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,labels:{color:TC(),boxWidth:10,font:{size:11}}},tooltip:{...tooltipStyle,callbacks:{label:ctx=>ctx.dataset.yAxisID==='yR'?fmt(ctx.raw):ctx.raw+' uds'}}},
      scales:{
        x:{grid:{color:GC()},ticks:{color:TC()}},
        yR:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)},position:'left'},
        yU:{grid:{color:'none'},ticks:{color:PALETTE[2]},position:'right',beginAtZero:true}
      }
    }
  });

  // Top clientes chart
  const cliMap={};
  detRows.forEach(r=>{
    if(!cliMap[r.CLIENTE]) cliMap[r.CLIENTE]={cli:r.CLIENTE,rev:0,uds:0};
    cliMap[r.CLIENTE].rev+=r.SUBTOTAL; cliMap[r.CLIENTE].uds+=r.CANTIDAD;
  });
  const cliSorted = Object.values(cliMap).sort((a,b)=>b.rev-a.rev);
  const top10 = cliSorted.slice(0,10);

  mkChart('ch-pd-cli',{
    type:'bar',
    data:{
      labels:top10.map(c=>c.cli),
      datasets:[{label:'Facturación',data:top10.map(c=>c.rev),backgroundColor:'rgba(59,130,246,.7)',borderRadius:4}]
    },
    options:{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{...tooltipStyle,callbacks:{label:ctx=>fmt(ctx.raw)}}},
      scales:{x:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}},y:{grid:{color:'none'},ticks:{color:TC()}}}
    }
  });

  // Tabla detalle clientes
  const tbody = document.querySelector('#tbl-pd-cli tbody');
  if(tbody) tbody.innerHTML = cliSorted.map((c,i)=>`
    <tr>
      <td>${i+1}</td><td>${c.cli}</td>
      <td class="num">${fmt(c.rev)}</td>
      <td class="num">${c.uds}</td>
      <td class="num">${c.uds>0?fmt(c.rev/c.uds):'—'}</td>
    </tr>`).join('');
}

function _renderProdOverview(rows, sorted) {
  const top15r = sorted.slice(0,15);
  const top15u = [...sorted].sort((a,b)=>b.uds-a.uds).slice(0,15);

  mkChart('ch-prod-rev',{
    type:'bar',
    data:{
      labels:top15r.map(p=>p.cod),
      datasets:[{label:'Facturación',data:top15r.map(p=>p.rev),backgroundColor:top15r.map(p=>p.linea==='MONCAR'?'rgba(0,212,170,.7)':'rgba(59,130,246,.7)'),borderRadius:4}]
    },
    options:{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{...tooltipStyle,callbacks:{label:ctx=>fmt(ctx.raw)}}},
      scales:{x:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}},y:{grid:{color:'none'},ticks:{color:TC(),font:{family:'DM Mono',size:11}}}}
    }
  });

  mkChart('ch-prod-units',{
    type:'bar',
    data:{
      labels:top15u.map(p=>p.cod),
      datasets:[{label:'Unidades',data:top15u.map(p=>p.uds),backgroundColor:top15u.map(p=>p.linea==='MONCAR'?'rgba(245,158,11,.7)':'rgba(168,85,247,.7)'),borderRadius:4}]
    },
    options:{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{...tooltipStyle}},
      scales:{x:{grid:{color:GC()},ticks:{color:TC()}},y:{grid:{color:'none'},ticks:{color:TC(),font:{family:'DM Mono',size:11}}}}
    }
  });

  const top8 = sorted.slice(0,8);
  const months = [...new Set(rows.map(r=>r.MES))].sort();
  const evolData={};
  top8.forEach(p=>{evolData[p.cod]={};months.forEach(m=>evolData[p.cod][m]=0);});
  rows.forEach(r=>{if(evolData[r.COD]) evolData[r.COD][r.MES]=(evolData[r.COD][r.MES]||0)+r.SUBTOTAL;});

  mkChart('ch-prod-evol',{
    type:'line',
    data:{
      labels:months.map(shortMes),
      datasets:top8.map((p,i)=>({
        label:p.cod,data:months.map(m=>evolData[p.cod][m]),
        borderColor:PALETTE[i],backgroundColor:'transparent',
        pointBackgroundColor:PALETTE[i],tension:0.4,borderWidth:2,pointRadius:3
      }))
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,labels:{color:TC(),boxWidth:12,font:{size:11}}},tooltip:{...tooltipStyle,callbacks:{label:ctx=>`${ctx.dataset.label}: ${fmt(ctx.raw)}`}}},
      scales:{x:{grid:{color:GC()},ticks:{color:TC()}},y:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}}}
    }
  });
}

function _renderProdTable(sorted, totalRev) {
  const tbody = document.querySelector('#tbl-prod tbody');
  if(!tbody) return;
  tbody.innerHTML = sorted.map((p,i)=>{
    const prom = p.uds>0 ? p.rev/p.uds : 0;
    const pct  = totalRev>0 ? (p.rev/totalRev*100).toFixed(1)+'%' : '—';
    const lBg  = p.linea==='MONCAR' ? 'rgba(0,212,170,.15)' : 'rgba(59,130,246,.15)';
    const lFg  = p.linea==='MONCAR' ? 'var(--primary)' : '#60a5fa';
    const isSelected = PROD_SEL===p.cod;
    return `<tr onclick="selectProd(${JSON.stringify(p.cod)})" style="cursor:pointer${isSelected?';background:rgba(255,255,255,.05)':''}">
      <td>${i+1}</td>
      <td><strong style="font-family:DM Mono">${p.cod}</strong></td>
      <td style="max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:11px;color:var(--muted)">${p.desc}</td>
      <td><span style="font-size:11px;background:${lBg};padding:2px 7px;border-radius:10px;color:${lFg}">${p.linea||''}</span></td>
      <td class="num">${fmt(p.rev)}</td>
      <td class="num">${p.uds}</td>
      <td class="num">${p.clis.size}</td>
      <td class="num">${prom>0?fmt(prom):'—'}</td>
      <td class="num">${pct}</td>
    </tr>`;
  }).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// MARCAS
// ─────────────────────────────────────────────────────────────────────────────
// ── MARCAS: selector y detalle
let MARCA_SEL = null;

function selectMarc(nombre) {
  MARCA_SEL = nombre;
  const inp = document.getElementById('marc-search');
  if(inp) inp.value = nombre || '';
  const dd = document.getElementById('marc-dd');
  if(dd) dd.style.display = 'none';
  renderMarcas();
}

function searchMarc() {
  const q = (document.getElementById('marc-search').value||'').trim().toUpperCase();
  if(!q){ selectMarc(null); return; }
  const rows = rowsFiltered();
  const names = [...new Set(rows.map(r=>r.MARCA||'(sin marca)'))];
  const exact = names.find(n=>n.toUpperCase()===q);
  if(exact){ selectMarc(exact); return; }
  const partial = names.filter(n=>n.toUpperCase().includes(q));
  if(partial.length===1){ selectMarc(partial[0]); return; }
  renderMarcDropdown();
}

function renderMarcDropdown() {
  const inp = document.getElementById('marc-search');
  const dd  = document.getElementById('marc-dd');
  if(!inp || !dd) return;
  const q = inp.value.trim().toUpperCase();
  const rows = rowsFiltered();
  const names = [...new Set(rows.map(r=>r.MARCA||'(sin marca)'))].sort();
  const filtered = q ? names.filter(n=>n.toUpperCase().includes(q)) : names;
  if(!filtered.length){ dd.style.display='none'; return; }
  dd.innerHTML = filtered.slice(0,25).map(n=>
    `<div style="padding:7px 12px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--border)"
          onclick="selectMarc(${JSON.stringify(n)})">${n}</div>`
  ).join('');
  dd.style.display = '';
}

function renderMarcas() {
  const rows = rowsFiltered();

  const marcaMap = {};
  rows.forEach(r => {
    const m = r.MARCA || '(sin marca)';
    if(!marcaMap[m]) marcaMap[m]={marca:m,rev:0,uds:0,cmv:0,prods:new Set(),clis:new Set()};
    marcaMap[m].rev += r.SUBTOTAL;
    marcaMap[m].uds += r.CANTIDAD;
    marcaMap[m].cmv += r.CMV;
    marcaMap[m].prods.add(r.COD);
    marcaMap[m].clis.add(r.CLIENTE);
  });
  const sorted = Object.values(marcaMap).sort((a,b)=>b.rev-a.rev);
  const totalRev = sorted.reduce((a,m)=>a+m.rev,0);

  // Si la marca seleccionada no tiene datos, resetear
  if(MARCA_SEL && !marcaMap[MARCA_SEL]) {
    MARCA_SEL = null;
    const inp = document.getElementById('marc-search');
    if(inp) inp.value = '';
  }

  // Badge
  const badge = document.getElementById('marc-badge');
  if(badge) {
    if(MARCA_SEL) {
      badge.style.display = '';
      document.getElementById('marc-badge-nom').textContent = MARCA_SEL;
    } else {
      badge.style.display = 'none';
    }
  }

  // Toggle vistas
  const detailEl   = document.getElementById('marc-detail');
  const overviewEl = document.getElementById('marc-overview');
  if(MARCA_SEL && marcaMap[MARCA_SEL]) {
    if(detailEl)  detailEl.style.display  = '';
    if(overviewEl) overviewEl.style.display = 'none';
    _renderMarcDetail(rows, marcaMap[MARCA_SEL]);
  } else {
    if(detailEl)  detailEl.style.display  = 'none';
    if(overviewEl) overviewEl.style.display = '';
    _renderMarcOverview(rows, sorted);
  }

  // Tabla resumen siempre
  const tbody = document.querySelector('#tbl-marc tbody');
  if(tbody) tbody.innerHTML = sorted.map((m,i)=>`<tr onclick="selectMarc(${JSON.stringify(m.marca)})" style="cursor:pointer${MARCA_SEL===m.marca?';background:rgba(255,255,255,.05)':''}">
    <td class="num" style="color:var(--muted)">${i+1}</td>
    <td><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${PALETTE[i%PALETTE.length]};margin-right:8px"></span>${m.marca}</td>
    <td class="num">${fmt(m.rev)}</td>
    <td class="num">${fmtN(m.uds)}</td>
    <td class="num" style="color:var(--green)">${fmt(m.rev-m.cmv)}</td>
    <td class="num" style="color:var(--muted)">${pct(totalRev>0?m.rev/totalRev:0)}</td>
  </tr>`).join('');
}

function _renderMarcDetail(rows, m) {
  const detRows = rows.filter(r=>(r.MARCA||'(sin marca)')===m.marca);

  document.getElementById('md-rev').textContent   = fmt(m.rev);
  document.getElementById('md-uds').textContent   = m.uds.toLocaleString('es-AR');
  document.getElementById('md-prods').textContent = m.prods.size;
  document.getElementById('md-clis').textContent  = m.clis.size;

  // Evolución mensual
  const months = [...new Set(rows.map(r=>r.MES))].sort();
  const mRev={}, mUds={};
  months.forEach(mes=>{mRev[mes]=0;mUds[mes]=0;});
  detRows.forEach(r=>{ mRev[r.MES]=(mRev[r.MES]||0)+r.SUBTOTAL; mUds[r.MES]=(mUds[r.MES]||0)+r.CANTIDAD; });

  mkChart('ch-md-evol',{
    type:'line',
    data:{
      labels:months.map(shortMes),
      datasets:[
        {label:'Facturación',data:months.map(mes=>mRev[mes]),borderColor:PALETTE[0],backgroundColor:'transparent',yAxisID:'yR',tension:0.4,borderWidth:2,pointRadius:4,pointBackgroundColor:PALETTE[0]},
        {label:'Unidades',  data:months.map(mes=>mUds[mes]),borderColor:PALETTE[2],backgroundColor:'transparent',yAxisID:'yU',tension:0.4,borderWidth:2,pointRadius:4,pointBackgroundColor:PALETTE[2]}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,labels:{color:TC(),boxWidth:10,font:{size:11}}},tooltip:{...tooltipStyle,callbacks:{label:ctx=>ctx.dataset.yAxisID==='yR'?fmt(ctx.raw):ctx.raw+' uds'}}},
      scales:{
        x:{grid:{color:GC()},ticks:{color:TC()}},
        yR:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)},position:'left'},
        yU:{grid:{color:'none'},ticks:{color:PALETTE[2]},position:'right',beginAtZero:true}
      }
    }
  });

  // Top productos de esta marca
  const prodMap={};
  detRows.forEach(r=>{ if(!prodMap[r.COD]) prodMap[r.COD]={cod:r.COD,rev:0,uds:0}; prodMap[r.COD].rev+=r.SUBTOTAL; prodMap[r.COD].uds+=r.CANTIDAD; });
  const topProds = Object.values(prodMap).sort((a,b)=>b.rev-a.rev).slice(0,12);

  mkChart('ch-md-prods',{
    type:'bar',
    data:{
      labels:topProds.map(p=>p.cod),
      datasets:[{label:'Facturación',data:topProds.map(p=>p.rev),backgroundColor:'rgba(245,158,11,.7)',borderRadius:4}]
    },
    options:{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{...tooltipStyle,callbacks:{label:ctx=>fmt(ctx.raw)}}},
      scales:{x:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}},y:{grid:{color:'none'},ticks:{color:TC(),font:{family:'DM Mono',size:11}}}}
    }
  });

  // Tabla clientes
  const cliMap={};
  detRows.forEach(r=>{ if(!cliMap[r.CLIENTE]) cliMap[r.CLIENTE]={cli:r.CLIENTE,rev:0,uds:0}; cliMap[r.CLIENTE].rev+=r.SUBTOTAL; cliMap[r.CLIENTE].uds+=r.CANTIDAD; });
  const cliSorted = Object.values(cliMap).sort((a,b)=>b.rev-a.rev);
  const tbody = document.querySelector('#tbl-md-cli tbody');
  if(tbody) tbody.innerHTML = cliSorted.map((c,i)=>`<tr>
    <td>${i+1}</td><td>${c.cli}</td>
    <td class="num">${fmt(c.rev)}</td>
    <td class="num">${c.uds}</td>
    <td class="num">${pct(m.rev>0?c.rev/m.rev:0)}</td>
  </tr>`).join('');
}

function _renderMarcOverview(rows, sorted) {
  // Dynamic chart height so all brands fit
  const marcaH = Math.max(220, sorted.length * 34 + 50) + 'px';
  const revEl = document.querySelector('#ch-marc-rev');
  const udsEl = document.querySelector('#ch-marc-units');
  if(revEl) revEl.parentElement.style.height = marcaH;
  if(udsEl) udsEl.parentElement.style.height = marcaH;

  mkChart('ch-marc-rev',{
    type:'bar',
    data:{labels:sorted.map(m=>m.marca),datasets:[{label:'Facturación',data:sorted.map(m=>m.rev),backgroundColor:sorted.map((_,i)=>PALETTE[i%PALETTE.length]),borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{...tooltipStyle,callbacks:{label:ctx=>fmt(ctx.raw)}}},scales:{x:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}},y:{grid:{color:'none'},ticks:{color:TC()}}}}
  });
  mkChart('ch-marc-units',{
    type:'bar',
    data:{labels:sorted.map(m=>m.marca),datasets:[{label:'Unidades',data:sorted.map(m=>m.uds),backgroundColor:sorted.map((_,i)=>PALETTE[i%PALETTE.length]+'bb'),borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{...tooltipStyle}},scales:{x:{grid:{color:GC()},ticks:{color:TC()}},y:{grid:{color:'none'},ticks:{color:TC()}}}}
  });

  const top6 = sorted.slice(0,6);
  const months = [...new Set(rows.map(r=>r.MES))].sort();
  const evolData={};
  top6.forEach(m=>{evolData[m.marca]={};months.forEach(mes=>evolData[m.marca][mes]=0);});
  rows.forEach(r=>{ const m=r.MARCA||'(sin marca)'; if(evolData[m]) evolData[m][r.MES]=(evolData[m][r.MES]||0)+r.SUBTOTAL; });
  mkChart('ch-marc-evol',{
    type:'line',
    data:{labels:months.map(shortMes),datasets:top6.map((m,i)=>({label:m.marca,data:months.map(mes=>evolData[m.marca][mes]),borderColor:PALETTE[i],backgroundColor:'transparent',pointBackgroundColor:PALETTE[i],tension:0.4,borderWidth:2,pointRadius:3}))},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,labels:{color:TC(),boxWidth:12}},tooltip:{...tooltipStyle,callbacks:{label:ctx=>`${ctx.dataset.label}: ${fmt(ctx.raw)}`}}},scales:{x:{grid:{color:GC()},ticks:{color:TC()}},y:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}}}}
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// CLIENTES
// ─────────────────────────────────────────────────────────────────────────────
function renderClientes() {
  const rows = rowsFiltered();

  const cliMap = {};
  rows.forEach(r => {
    const k = r.CLIENTE||'(sin nombre)';
    if(!cliMap[k]) cliMap[k]={cli:k,rev:0,uds:0,prov:r.PROVINCIA};
    cliMap[k].rev += r.SUBTOTAL;
    cliMap[k].uds += r.CANTIDAD;
    if(!cliMap[k].prov && r.PROVINCIA) cliMap[k].prov = r.PROVINCIA;
  });
  const sorted = Object.values(cliMap).sort((a,b)=>b.rev-a.rev);
  const totalRev = sorted.reduce((a,c)=>a+c.rev,0);
  const top15 = sorted.slice(0,15);

  mkChart('ch-cli-rev', {
    type:'bar',
    data:{
      labels: top15.map(c=>truncate(c.cli,20)),
      datasets:[{label:'Facturación', data:top15.map(c=>c.rev), backgroundColor:PALETTE.slice(0,15), borderRadius:4}]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, tooltip:{...tooltipStyle,callbacks:{label:ctx=>fmt(ctx.raw)}} },
      scales:{ x:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}}, y:{grid:{color:'none'},ticks:{color:TC(),font:{size:11}}} }
    }
  });

  mkChart('ch-cli-units', {
    type:'bar',
    data:{
      labels: [...sorted].sort((a,b)=>b.uds-a.uds).slice(0,15).map(c=>truncate(c.cli,20)),
      datasets:[{label:'Unidades', data:[...sorted].sort((a,b)=>b.uds-a.uds).slice(0,15).map(c=>c.uds), backgroundColor:PALETTE.slice(0,15).map(x=>x+'99'), borderRadius:4}]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, tooltip:{...tooltipStyle} },
      scales:{ x:{grid:{color:GC()},ticks:{color:TC()}}, y:{grid:{color:'none'},ticks:{color:TC(),font:{size:11}}} }
    }
  });

  const tbody = document.querySelector('#tbl-cli tbody');
  tbody.innerHTML = sorted.map((c,i)=>`<tr>
    <td class="num" style="color:var(--muted)">${i+1}</td>
    <td>${c.cli}</td>
    <td class="num">${fmt(c.rev)}</td>
    <td class="num">${fmtN(c.uds)}</td>
    <td class="num" style="color:var(--muted)">${pct(totalRev>0?c.rev/totalRev:0)}</td>
    <td style="color:var(--muted);font-size:12px">${c.prov||'—'}</td>
  </tr>`).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// EMBOLSADO PRIORITY
// ─────────────────────────────────────────────────────────────────────────────
function urgColor(w) {
  if(w===null||w===undefined||!isFinite(w)||w>52) return 'var(--muted)';
  if(w<1)  return 'var(--red)';
  if(w<2)  return '#f97316';
  if(w<4)  return 'var(--accent3)';
  return 'var(--green)';
}
function urgLabel(w) { return (w===null||w===undefined||!isFinite(w)||w>52) ? '∞' : w.toFixed(1)+' sem'; }

function calcEmbolsado(sRows, salesMap, avgFactor, isMC) {
  return sRows.map(r => {
    const avgM = salesMap[r.COD] ? salesMap[r.COD]/avgFactor : 0;
    if(r.NO_EMB<=0) return null;
    const avgW   = avgM>0 ? avgM/4.33 : 0;
    const embDep = isMC ? r.EMB_DEP+r.PREM_DEP : r.EMB_DEP;
    const noEmb  = r.NO_EMB;
    const totalS = isMC
      ? noEmb+r.EMB_DEP+r.PREM_DEP+r.JUANI
      : noEmb+r.EMB_DEP+r.MAXI;
    const wEmb  = avgW>0 ? embDep/avgW : Infinity;
    const wMaxi = (!isMC&&avgW>0) ? (r.MAXI||0)/avgW : null;
    const mTot  = avgM>0 ? totalS/avgM : Infinity;
    const toEmb = avgM>0 ? Math.min(Math.max(0, Math.ceil(avgM*1.5-embDep)), noEmb) : 0;
    const buy   = isFinite(mTot)&&mTot<2;
    const buyQty= buy ? Math.max(0, Math.ceil(3*avgM-totalS)) : 0;
    return {
      cod:r.COD, lbl:isMC?r.MARCA:r.DESC,
      embDep, noEmb, maxi:isMC?null:r.MAXI,
      avgM, wEmb, wMaxi, mTot, toEmb, buy, buyQty
    };
  }).filter(Boolean).sort((a,b)=>a.wEmb-b.wEmb);
}

function renderEmbolsadoCard(topEmb, isMC, topN) {
  if(!topEmb.length) return '<div class="card" style="margin-bottom:16px"><div class="empty">Sin datos de venta en los últimos 6 meses para calcular prioridad de embolsado.</div></div>';
  const mFmt = m => isFinite(m)&&m<=36 ? m.toFixed(1)+'m' : '∞';
  const rows = topEmb.map(r => {
    const maxiTd  = !isMC ? '<td class="num">'+fmtN(r.maxi||0)+'</td>' : '';
    const maxiSem = !isMC ? '<td class="num" style="color:'+urgColor(r.wMaxi)+'">'+urgLabel(r.wMaxi)+'</td>' : '';
    const rc = r.wEmb<1 ? 'row-danger' : r.wEmb<3 ? 'row-warn' : '';
    return '<tr class="'+rc+'">'
      +'<td class="cod">'+r.cod+'</td>'
      +'<td style="font-size:12px">'+(r.lbl||'—')+'</td>'
      +'<td class="num">'+fmtN(r.embDep)+'</td>'
      +'<td class="num">'+fmtN(r.noEmb)+'</td>'
      +maxiTd
      +'<td class="num" style="color:var(--muted)">'+r.avgM.toFixed(1)+'</td>'
      +'<td class="num" style="font-weight:700;color:'+urgColor(r.wEmb)+'">'+urgLabel(r.wEmb)+'</td>'
      +maxiSem
      +'<td class="num" style="color:var(--accent);font-weight:600">'+(r.toEmb>0?fmtN(r.toEmb):'—')+'</td>'
      +'<td style="font-size:12px;color:'+(r.buy?'var(--red)':'var(--muted)')+'">'+(r.buy?'SI — '+fmtN(r.buyQty)+' uds':'—')+'</td>'
      +'</tr>';
  }).join('');
  const mxTh  = !isMC ? '<th class="num">Maxi</th>' : '';
  const mxSTh = !isMC ? '<th class="num">Sem. Maxi</th>' : '';
  const lCol  = isMC  ? 'MARCA' : 'DESCRIPCIÓN';
  const noteMaxi = !isMC
    ? ' &nbsp;|&nbsp; <strong>Sem. Maxi</strong>: semanas de stock de Maxi ÷ ventas — cuándo te va a pedir repuesto. CABA excluida (sin vendedor activo).'
    : '';
  return '<div class="card" style="margin-bottom:16px">'
    +'<h3>Prioridad de Embolsado — Top '+topN+' '+(isMC?'MONCAR':'RMH')+'</h3>'
    +'<p style="font-size:12px;color:var(--muted);margin-bottom:12px;margin-top:-6px">Ordenado por urgencia: menos semanas de embolsado en depósito = embolsá primero. Velocidad calculada sobre los últimos 6 meses de ventas.</p>'
    +'<div class="tbl-wrap"><table>'
    +'<thead><tr>'
    +'<th>COD</th><th>'+lCol+'</th>'
    +'<th class="num">Emb. Dep.</th><th class="num">Sin Emb.</th>'
    +mxTh
    +'<th class="num">Vtas/mes</th><th class="num">Sem. Emb. Dep.</th>'
    +mxSTh
    +'<th class="num">Embolsar ahora</th><th>Comprar?</th>'
    +'</tr></thead>'
    +'<tbody>'+rows+'</tbody>'
    +'</table></div>'
    +'<div class="note">'
    +'<strong>Velocidad</strong>: promedio de los últimos 6 meses de ventas. '
    +'<strong>Sem. Emb. Dep.</strong>: semanas de embolsado en depósito ÷ ventas/sem. '
    +'Rojo = urgente (&lt;1s) · naranja (&lt;2s) · amarillo (&lt;4s) · verde (≥4s). ∞ = sin ventas recientes.'
    +noteMaxi
    +' &nbsp;|&nbsp; <strong>Embolsar ahora</strong>: uds para alcanzar 1.5 meses en depósito (limitado por sin embolsar disponible).'
    +'</div>'
    +'</div>';
}

// ─────────────────────────────────────────────────────────────────────────────
// STOCK
// ─────────────────────────────────────────────────────────────────────────────
function setStock(which) {
  F.stock = which;
  document.getElementById('btn-mc').classList.toggle('active', which==='mc');
  document.getElementById('btn-rmh').classList.toggle('active', which==='rmh');
  renderStock();
}

function renderStock() {
  const isMC = F.stock === 'mc';
  const sRows = isMC ? MC_ROWS : RMH_ROWS;
  const body = document.getElementById('stock-body');

  if(!sRows || sRows.length===0) {
    body.innerHTML = `<div class="empty">No se encontró archivo de stock para ${isMC?'MONCAR':'RMH'}.<br>Asegurate de que el archivo ${isMC?'STOCK_MONCAR':'STOCK_RMH'}.xlsx esté en la misma carpeta.</div>`;
    return;
  }

  // Calc avg monthly sales last 6 months (from ventas)
  const ventasLinea = V.raw.filter(r=>r.LINEA===(isMC?'MONCAR':'RMH'));
  const allMonths = [...new Set(ventasLinea.map(r=>r.MES))].sort();
  const last6 = allMonths.slice(-6);
  const salesMap = {};
  ventasLinea.filter(r=>last6.includes(r.MES)).forEach(r=>{
    salesMap[r.COD] = (salesMap[r.COD]||0) + r.CANTIDAD;
  });
  const avgFactor = last6.length > 0 ? last6.length : 1;

  // Embolsado priority
  const embRows = calcEmbolsado(sRows, salesMap, avgFactor, isMC);
  const topN    = 10;
  const topEmb  = embRows.slice(0, topN);

  // KPIs
  const totalDep   = sRows.reduce((a,r)=>a+r.TOTAL,0);
  const totalVend  = sRows.reduce((a,r)=>a+r.VENDIBLE,0);
  const agotados   = sRows.filter(r=>r.TOTAL<=0).length;
  const valorEst   = sRows.reduce((a,r)=>a+(LP[r.COD]||0)*r.TOTAL,0);

  // Rotación avg
  const totalSold = Object.values(salesMap).reduce((a,v)=>a+v,0);
  const rotAvg = totalDep > 0 ? (totalSold/avgFactor)/totalDep : 0;

  let kpiExtra = '';
  if(last6.length>0) {
    kpiExtra = kpiCard('Rotación Avg (últ. 6m)', (rotAvg*100).toFixed(1)+'%', `uds/mes vendidas vs stock actual`, 'accent');
  }

  // Top 15 by TOTAL
  const top15 = [...sRows].sort((a,b)=>b.TOTAL-a.TOTAL).slice(0,15);

  // Donut distribution
  const noEmb   = sRows.reduce((a,r)=>a+r.NO_EMB,0);
  const embDep  = sRows.reduce((a,r)=>a+r.EMB_DEP,0);
  const col3    = isMC ? sRows.reduce((a,r)=>a+r.PREM_DEP,0) : sRows.reduce((a,r)=>a+r.CABA,0);
  const col4    = isMC ? sRows.reduce((a,r)=>a+r.JUANI,0)    : sRows.reduce((a,r)=>a+r.MAXI,0);
  const col3lbl = isMC ? 'Prem. Dep' : 'CABA';
  const col4lbl = isMC ? 'Juani'     : 'Maxi';

  // Pre-compute comprar card (avoids body.innerHTML += re-parse)
  const comprarHtml = renderComprarCard(calcComprar(sRows, salesMap, avgFactor, isMC), isMC);

  // Sort detail table: stock desc (zeros last)
  const sRowsSorted = [...sRows].sort((a,b)=>a.COD.localeCompare(b.COD));

  body.innerHTML = `
    <div class="kpi-row cols4" style="margin-bottom:12px">
      ${kpiCard('Total en Depósito', fmtN(totalDep), `en ${sRows.length} productos`, 'blue')}
      ${kpiCard('Unidades Vendibles', fmtN(totalVend), `de ${fmtN(totalDep)} totales`, 'green')}
      ${kpiCard('Productos Agotados', fmtN(agotados), `con stock = 0`, 'red')}
      ${kpiCard('Valor Estimado', fmt(valorEst), 'al precio de lista', 'amber')}
    </div>
    ${kpiExtra ? `<div class="kpi-row cols2" style="max-width:600px;margin-bottom:16px">${kpiExtra}</div>` : ''}
    ` + renderEmbolsadoCard(topEmb, isMC, topN) + `
    <div class="grid-2" style="margin-bottom:16px">
      <div class="card">
        <h3>Top 15 por Stock Total</h3>
        <div class="chart-wrap" style="height:320px"><canvas id="ch-stk-top"></canvas></div>
      </div>
      <div class="card">
        <h3>Distribución por Depósito</h3>
        <div class="chart-wrap" style="height:320px"><canvas id="ch-stk-donut"></canvas></div>
      </div>
    </div>
    <div class="card">
      <h3>Detalle Completo de Stock — ${isMC?'MONCAR':'RMH'}</h3>
      <div class="tbl-wrap">
        <table id="tbl-stk">
          <thead><tr>
            <th>COD</th>
            <th>${isMC?'MARCA':'DESCRIPCIÓN'}</th>
            <th class="num">NO EMB</th>
            <th class="num">EMB DEP</th>
            <th class="num">${col3lbl}</th>
            <th class="num">${col4lbl}</th>
            <th class="num">TOTAL</th>
            <th class="num">VENDIBLE</th>
            <th class="num">PRECIO</th>
            <th class="num">VALOR EST</th>
            <th class="num">MESES STOCK</th>
          </tr></thead>
          <tbody>
            ${sRowsSorted.map(r=>{
              const precio = LP[r.COD]||0;
              const valorRow = precio * r.TOTAL;
              const avgSales = salesMap[r.COD] ? salesMap[r.COD]/avgFactor : 0;
              const mesesStock = avgSales>0 ? (r.TOTAL/avgSales).toFixed(1) : '∞';
              const tintClass = r.TOTAL<=0 ? 'row-danger' : r.TOTAL<5 ? 'row-warn' : '';
              const col3val = isMC ? r.PREM_DEP : r.CABA;
              const col4val = isMC ? r.JUANI    : r.MAXI;
              return `<tr class="${tintClass}">
                <td class="cod">${r.COD}</td>
                <td style="font-size:12px">${isMC?r.MARCA:r.DESC}</td>
                <td class="num">${fmtN(r.NO_EMB)}</td>
                <td class="num">${fmtN(r.EMB_DEP)}</td>
                <td class="num">${fmtN(col3val)}</td>
                <td class="num">${fmtN(col4val)}</td>
                <td class="num" style="font-weight:600">${fmtN(r.TOTAL)}</td>
                <td class="num">${fmtN(r.VENDIBLE)}</td>
                <td class="num" style="color:var(--muted)">${precio>0?fmt(precio):'—'}</td>
                <td class="num" style="color:var(--accent3)">${valorRow>0?fmt(valorRow):'—'}</td>
                <td class="num" style="color:${mesesStock==='∞'?'var(--muted)':parseFloat(mesesStock)<2?'var(--red)':parseFloat(mesesStock)<4?'var(--accent3)':'var(--green)'}">${mesesStock}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
      <div class="note">Tabla ordenada por código (A→Z). Valores al último guardado en Excel. Precio de lista tomado de hoja LP en VENTAS.xlsx.</div>
    </div>` + comprarHtml;

  // Charts (after DOM render)
  requestAnimationFrame(()=>{
    mkChart('ch-stk-top', {
      type:'bar',
      data:{
        labels: top15.map(r=>r.COD),
        datasets:[{
          label:'Stock Total', data:top15.map(r=>r.TOTAL),
          backgroundColor:'rgba(59,130,246,.7)', borderRadius:4
        }]
      },
      options:{
        indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{display:false}, tooltip:{...tooltipStyle} },
        scales:{ x:{grid:{color:GC()},ticks:{color:TC()}}, y:{grid:{color:'none'},ticks:{color:TC(),font:{family:'DM Mono',size:11}}} }
      }
    });

    mkChart('ch-stk-donut', {
      type:'doughnut',
      data:{
        labels:['No Emb', 'Emb Dep', col3lbl, col4lbl],
        datasets:[{
          data:[noEmb, embDep, col3, col4],
          backgroundColor:['rgba(0,212,170,.8)','rgba(59,130,246,.8)','rgba(245,158,11,.8)','rgba(168,85,247,.8)'],
          borderWidth:1, borderColor:'#181c23', hoverOffset:4
        }]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        plugins:{
          legend:{display:true,position:'bottom',labels:{color:TC(),padding:16,boxWidth:12}},
          tooltip:{...tooltipStyle,callbacks:{label:ctx=>`${ctx.label}: ${fmtN(ctx.raw)}`}}
        },
        cutout:'60%'
      }
    });
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// COMPRAR
// ─────────────────────────────────────────────────────────────────────────────
function calcComprar(sRows, salesMap, avgFactor, isMC) {
  var minBuy   = isMC ? 100 : 20;
  var multiple = isMC ? 100 : 5;
  var targetM  = 8;
  var result   = [];
  for (var i = 0; i < sRows.length; i++) {
    var r      = sRows[i];
    var avgM   = salesMap[r.COD] ? salesMap[r.COD] / avgFactor : 0;
    var totalS = isMC
      ? r.NO_EMB + r.EMB_DEP + r.PREM_DEP + r.JUANI
      : r.NO_EMB + r.EMB_DEP + (r.CABA || 0) + (r.MAXI || 0);
    var mTot   = avgM > 0 ? totalS / avgM : Infinity;
    if (isFinite(mTot) && mTot >= targetM) continue;
    if (avgM <= 0 && totalS > 0) continue;
    var needed  = avgM > 0 ? Math.max(0, targetM * avgM - totalS) : 0;
    var rounded = needed > 0 ? Math.ceil(needed / multiple) * multiple : minBuy;
    var buyQty  = Math.max(minBuy, rounded);
    result.push({ cod: r.COD, lbl: isMC ? r.MARCA : r.DESC, totalS: totalS, avgM: avgM, mTot: mTot, buyQty: buyQty });
  }
  result.sort(function(a, b) { return a.mTot - b.mTot; });
  return result;
}

function exportComprar(isMC) {
  var sRows = isMC ? MC_ROWS : RMH_ROWS;
  var ventasLinea = [];
  for (var i = 0; i < V.raw.length; i++) {
    if (V.raw[i].LINEA === (isMC ? 'MONCAR' : 'RMH')) ventasLinea.push(V.raw[i]);
  }
  var allMonths = [];
  for (var i = 0; i < ventasLinea.length; i++) {
    if (allMonths.indexOf(ventasLinea[i].MES) < 0) allMonths.push(ventasLinea[i].MES);
  }
  allMonths.sort();
  var last6     = allMonths.slice(-6);
  var salesMap  = {};
  for (var i = 0; i < ventasLinea.length; i++) {
    var rv = ventasLinea[i];
    if (last6.indexOf(rv.MES) >= 0) salesMap[rv.COD] = (salesMap[rv.COD] || 0) + rv.CANTIDAD;
  }
  var avgFactor = last6.length > 0 ? last6.length : 1;
  var compRows  = calcComprar(sRows, salesMap, avgFactor, isMC);
  var linea     = isMC ? 'MONCAR_Carbones' : 'RMH_Llaves';
  var fecha     = new Date().toISOString().slice(0, 10);
  var lCol      = isMC ? 'MARCA' : 'DESCRIPCION';
  var bom       = String.fromCharCode(0xFEFF);
  var csv       = bom + 'COD;' + lCol + ';STOCK_ACTUAL;VTAS_MES;MESES_RESTANTES;COMPRAR_UDS\\n';
  var totalUds  = 0;
  for (var i = 0; i < compRows.length; i++) {
    var row   = compRows[i];
    var mTotS = isFinite(row.mTot) && row.mTot <= 99 ? row.mTot.toFixed(1) : '999';
    csv += row.cod + ';"' + (row.lbl || '') + '";' + Math.round(row.totalS) + ';' + row.avgM.toFixed(1) + ';' + mTotS + ';' + row.buyQty + '\\n';
    totalUds += row.buyQty;
  }
  csv += ';;;"TOTAL";;' + totalUds + '\\n';
  var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href     = url;
  a.download = 'Orden_Compra_' + linea + '_' + fecha + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function renderComprarCard(compRows, isMC) {
  if (!compRows.length) return '';
  var mFmt = function(m) { return isFinite(m) && m <= 36 ? m.toFixed(1) + 'm' : '∞'; };
  var mColor = function(m) {
    if (!isFinite(m)) return 'var(--muted)';
    if (m < 1) return 'var(--red)';
    if (m < 3) return '#f97316';
    if (m < 5) return 'var(--accent3)';
    return 'var(--green)';
  };
  var lCol   = isMC ? 'MARCA' : 'DESCRIPCIÓN';
  var linea  = isMC ? 'MONCAR — Carbones' : 'RMH — Llaves';
  var minStr = isMC ? 'Mínimo 100 uds · múltiplos de 100' : 'Mínimo 20 uds · múltiplos de 5';
  var totalUds = 0;
  for (var i = 0; i < compRows.length; i++) totalUds += compRows[i].buyQty;
  var rows = '';
  for (var i = 0; i < compRows.length; i++) {
    var r  = compRows[i];
    var rc = r.mTot < 1 ? 'row-danger' : r.mTot < 3 ? 'row-warn' : '';
    rows += '<tr class="' + rc + '">'
      + '<td class="cod">' + r.cod + '</td>'
      + '<td style="font-size:12px">' + (r.lbl || '&mdash;') + '</td>'
      + '<td class="num">' + fmtN(r.totalS) + '</td>'
      + '<td class="num" style="color:var(--muted)">' + r.avgM.toFixed(1) + '</td>'
      + '<td class="num" style="color:' + mColor(r.mTot) + ';font-weight:700">' + mFmt(r.mTot) + '</td>'
      + '<td class="num" style="color:var(--accent2);font-weight:700">' + fmtN(r.buyQty) + '</td>'
      + '</tr>';
  }
  return '<div class="card" style="margin-top:16px;border-color:rgba(59,130,246,.3)">'
    + '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">'
    + '<h3 style="margin-bottom:0">Orden de compra sugerida &mdash; ' + linea + '</h3>'
    + '<button onclick="exportComprar(' + isMC + ')" style="background:rgba(59,130,246,.15);border:1px solid rgba(59,130,246,.4);color:var(--accent2);border-radius:6px;padding:6px 16px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit">Exportar a Excel</button>'
    + '</div>'
    + '<p style="font-size:12px;color:var(--muted);margin-bottom:12px">'
    + 'Productos con menos de 8 meses de stock. Pedido para cubrir 8 meses '
    + '(4&ndash;5 meses de tr&aacute;nsito + 4&ndash;5 hasta el pr&oacute;ximo). '
    + minStr + '. Velocidad: promedio &uacute;ltimos 6 meses. '
    + '<strong>Total pedido: ' + fmtN(totalUds) + ' uds.</strong>'
    + '</p>'
    + '<div class="tbl-wrap"><table>'
    + '<thead><tr>'
    + '<th>COD</th><th>' + lCol + '</th>'
    + '<th class="num">Stock actual</th>'
    + '<th class="num">Vtas/mes</th>'
    + '<th class="num">Meses restantes</th>'
    + '<th class="num">Comprar (uds)</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody>'
    + '</table></div>'
    + '<div class="note">'
    + '<strong>Stock actual</strong>: suma de todos los dep&oacute;sitos y vendedores. '
    + '<strong>Comprar</strong>: uds para alcanzar 8 meses de stock, redondeado al m&uacute;ltiplo, m&iacute;nimo aplicado. '
    + 'Rojo &lt;1m &middot; naranja &lt;3m &middot; amarillo &lt;5m.'
    + '</div>'
    + '</div>';
}

// ─────────────────────────────────────────────────────────────────────────────
// P&L
// ─────────────────────────────────────────────────────────────────────────────
function renderPL() {
  mkPills('pills-pl','pl');
  const rows = filterVentas('pl');
  const monthly = getMonthly(rows);
  const months = Object.keys(monthly).sort();

  const total_rev  = rows.reduce((a,r)=>a+r.SUBTOTAL,0);
  const total_cmv  = rows.reduce((a,r)=>a+r.CMV,0);
  const total_com  = rows.reduce((a,r)=>a+r.SUBTOTAL*r.COMISION,0);
  const mb         = total_rev - total_cmv;
  const mb_pct     = total_rev>0 ? mb/total_rev : 0;
  const com_admin  = total_rev * 0.20;
  const resultado  = mb - total_com - com_admin;

  // MONCAR vs RMH split
  const mc_rev = rows.filter(r=>r.LINEA==='MONCAR').reduce((a,r)=>a+r.SUBTOTAL,0);
  const rmh_rev = rows.filter(r=>r.LINEA==='RMH').reduce((a,r)=>a+r.SUBTOTAL,0);
  const mc_cmv = rows.filter(r=>r.LINEA==='MONCAR').reduce((a,r)=>a+r.CMV,0);
  const rmh_cmv = rows.filter(r=>r.LINEA==='RMH').reduce((a,r)=>a+r.CMV,0);

  const rp = (v) => total_rev>0 ? '<span class="pl-pct">'+pct(Math.abs(v)/total_rev)+'</span>' : '';

  // PL table
  const plRows = [
    {label:'Facturación Total',     val:total_rev, style:'color:var(--green);font-weight:700',  badge:'<span class="pl-pct">100%</span>'},
    {label:'  — MONCAR',            val:mc_rev,    style:'color:var(--muted);font-size:12px',   badge:rp(mc_rev)},
    {label:'  — RMH',               val:rmh_rev,   style:'color:var(--muted);font-size:12px',   badge:rp(rmh_rev)},
    {label:'CMV (Costo Mercadería)', val:-total_cmv,style:'color:var(--red)',                   badge:rp(total_cmv)},
    {label:'Margen Bruto',          val:mb,        style:`color:${mb>=0?'var(--green)':'var(--red)'};font-weight:700;border-top:1px solid var(--border2)`, badge:rp(mb)},
    {label:'% Margen Bruto',        val:null, pct:mb_pct, style:'color:var(--accent)',          badge:''},
  ]; // Comisión y Resultado Neto se renderizan debajo con el detalle expandible

  // Comisión por vendedor
  const comPorVend = {};
  rows.forEach(r=>{ if(r.VENDEDOR){ comPorVend[r.VENDEDOR]=(comPorVend[r.VENDEDOR]||0)+r.SUBTOTAL*r.COMISION; } });
  const comVendRows = Object.entries(comPorVend).sort((a,b)=>b[1]-a[1]);

  const comDetRow = `
    <tr id="com-header" onclick="document.getElementById('com-detail').classList.toggle('hidden')" style="cursor:pointer">
      <td style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">
        Comisión Vendedores <span style="font-size:10px;opacity:.6">▼ ver detalle</span>
      </td>
      <td class="num" style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">${fmt(-total_com)}${rp(total_com)}</td>
    </tr>
    <tbody id="com-detail">
      ${comVendRows.map(([v,c])=>`
      <tr>
        <td style="padding:7px 12px 7px 28px;color:var(--accent3);font-size:12px;border-bottom:1px solid var(--border);opacity:.85">↳ ${v}</td>
        <td class="num" style="padding:7px 12px;color:var(--accent3);font-size:12px;border-bottom:1px solid var(--border);opacity:.85">${fmt(-c)}${rp(c)}</td>
      </tr>`).join('')}
      <tr>
        <td style="padding:8px 12px 8px 28px;color:var(--accent3);font-size:12px;font-weight:700;border-bottom:1px solid var(--border2)">Subtotal comisiones</td>
        <td class="num" style="padding:8px 12px;color:var(--accent3);font-size:12px;font-weight:700;border-bottom:1px solid var(--border2)">${fmt(-total_com)}${rp(total_com)}</td>
      </tr>
    </tbody>`;

  document.getElementById('pl-table').innerHTML = `
    <style>#com-detail.hidden{display:none}</style>
    <table style="width:100%">
      ${plRows.map(row=>`<tr>
        <td style="padding:10px 12px;${row.style||''};border-bottom:1px solid var(--border)">${row.label}</td>
        <td class="num" style="padding:10px 12px;${row.style||''};border-bottom:1px solid var(--border)">
          ${row.pct!==undefined ? pct(row.pct) : fmt(row.val||0)}${row.badge||''}
        </td>
      </tr>`).join('')}
      ${comDetRow}
      <tr>
        <td style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">Comisión administrativa (Franco 20%)</td>
        <td class="num" style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">${fmt(-com_admin)}${rp(com_admin)}</td>
      </tr>
      <tr>
        <td style="padding:10px 12px;color:${resultado>=0?'var(--green)':'var(--red)'};font-weight:700;font-size:16px;border-top:1px solid var(--border2)">Resultado Neto</td>
        <td class="num" style="padding:10px 12px;color:${resultado>=0?'var(--green)':'var(--red)'};font-weight:700;font-size:16px;border-top:1px solid var(--border2)">${fmt(resultado)}${rp(resultado)}</td>
      </tr>
    </table>`;

  // Waterfall
  const wfLabels = ['Facturación','CMV','Margen Bruto','Com. Vendedores','Com. Admin','Resultado'];
  const wfVals   = [total_rev, -total_cmv, mb, -total_com, -com_admin, resultado];
  const wfColors = [
    'rgba(0,212,170,.8)',
    'rgba(244,63,94,.8)',
    resultado>=0?'rgba(34,197,94,.6)':'rgba(244,63,94,.6)',
    'rgba(245,158,11,.8)',
    'rgba(168,85,247,.8)',
    resultado>=0?'rgba(34,197,94,.9)':'rgba(244,63,94,.9)'
  ];
  mkChart('ch-pl-wf', {
    type:'bar',
    data:{ labels:wfLabels, datasets:[{label:'P&L', data:wfVals, backgroundColor:wfColors, borderRadius:6}] },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, tooltip:{...tooltipStyle,callbacks:{label:ctx=>fmt(ctx.raw)}} },
      scales:{ x:{grid:{color:GC()},ticks:{color:TC()}}, y:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}} }
    }
  });

  // Stacked monthly
  mkChart('ch-pl-stack', {
    type:'bar',
    data:{
      labels: months.map(shortMes),
      datasets:[
        {label:'Facturación',  data:months.map(m=>monthly[m].revenue),          backgroundColor:'rgba(0,212,170,.6)', borderRadius:2},
        {label:'CMV',          data:months.map(m=>monthly[m].cmv),              backgroundColor:'rgba(244,63,94,.6)', borderRadius:2},
        {label:'Com. Vend.',   data:months.map(m=>monthly[m].comision),         backgroundColor:'rgba(245,158,11,.6)', borderRadius:2},
        {label:'Com. Admin',   data:months.map(m=>monthly[m].revenue*0.20),     backgroundColor:'rgba(168,85,247,.6)', borderRadius:2},
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:true,labels:{color:TC(),boxWidth:12}}, tooltip:{...tooltipStyle,callbacks:{label:ctx=>`${ctx.dataset.label}: ${fmt(ctx.raw)}`}} },
      scales:{ x:{grid:{color:GC()},ticks:{color:TC()}}, y:{grid:{color:GC()},ticks:{color:TC(),callback:v=>'$'+formatK(v)}} }
    }
  });

  // Margen % line
  const mgPct = months.map(m=>{
    const r = monthly[m].revenue;
    const c = monthly[m].cmv;
    return r>0 ? ((r-c)/r*100) : 0;
  });
  mkChart('ch-pl-mg', {
    type:'line',
    data:{
      labels: months.map(shortMes),
      datasets:[{
        label:'Margen Bruto %', data:mgPct,
        borderColor:'var(--accent)', backgroundColor:'rgba(0,212,170,.08)',
        fill:true, tension:0.4, borderWidth:2, pointRadius:4,
        pointBackgroundColor:'var(--accent)'
      }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, tooltip:{...tooltipStyle,callbacks:{label:ctx=>ctx.raw.toFixed(1)+'%'}} },
      scales:{
        x:{grid:{color:GC()},ticks:{color:TC()}},
        y:{grid:{color:GC()},ticks:{color:TC(),callback:v=>v.toFixed(0)+'%'}}
      }
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTER
// ─────────────────────────────────────────────────────────────────────────────
function renderPage(name) {
  if(name==='overview')  renderOverview();
  if(name==='productos') renderProductos();
  if(name==='marcas')    renderMarcas();
  if(name==='clientes')  renderClientes();
  if(name==='stock')     renderStock();
  if(name==='pl')        renderPL();
}

function showPage(name) {
  curPage = name;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  const tabs = document.querySelectorAll('.tab');
  const order = ['overview','productos','marcas','clientes','stock','pl'];
  const idx = order.indexOf(name);
  if(idx>=0 && tabs[idx]) tabs[idx].classList.add('active');
  const ac = document.getElementById('analytics-controls');
  if(ac) ac.style.display = ['overview','productos','marcas','clientes'].indexOf(name) >= 0 ? '' : 'none';
  renderPage(name);
}

function onVendedorChange() {
  F.vendedor = document.getElementById('sel-vendedor').value;
  renderPage(curPage);
}

// ─────────────────────────────────────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────────────────────────────────────
function shortMes(m) {
  if(!m) return '';
  const parts = m.split('-');
  const meses = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  return meses[parseInt(parts[1]||0)] + ' ' + (parts[0]||'').slice(2);
}

function formatK(v) {
  if(Math.abs(v)>=1000000) return (v/1000000).toFixed(1)+'M';
  if(Math.abs(v)>=1000)    return (v/1000).toFixed(0)+'k';
  return v;
}

function truncate(str, n) {
  return str && str.length>n ? str.slice(0,n-1)+'…' : (str||'');
}

// ─────────────────────────────────────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────────────────────────────────────
(function init() {
  // Restore saved theme
  try {
    if(localStorage.getItem('moncar_theme') === 'light') {
      document.body.classList.add('light');
      document.getElementById('theme-toggle').textContent = '🌙';
    }
  } catch(e) {}
  applyChartTheme();

  // Populate vendedor select
  const sel = document.getElementById('sel-vendedor');
  V.vendedores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });

  // Badge
  const totalReg = V.raw.length;
  document.getElementById('badge-info').textContent = `${window.DATA.generated_at} · ${totalReg} reg.`;

  // Cerrar dropdowns al hacer click fuera
  document.addEventListener('click', function(e){
    const checks = [['prod-dd','prod-search'],['marc-dd','marc-search']];
    checks.forEach(([ddId,inpId])=>{
      const dd=document.getElementById(ddId), inp=document.getElementById(inpId);
      if(dd && inp && !dd.contains(e.target) && e.target!==inp) dd.style.display='none';
    });
  });

  // Bind period selector and initialize
  bindPeriod();
  initPeriod();
})();
</script>
</body>
</html>"""

if __name__ == '__main__':
    main()
