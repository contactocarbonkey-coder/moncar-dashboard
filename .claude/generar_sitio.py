# -*- coding: utf-8 -*-
"""
generar_sitio.py — genera sitio/index.html con homepage + catalogo MONCAR + RMH.
Ejecutar cada vez que se actualicen precios o stock.
"""
import os, json
import openpyxl

DIR      = os.path.dirname(os.path.abspath(__file__))
OUT_DIR  = os.path.join(DIR, 'sitio')
OUT_PATH = os.path.join(OUT_DIR, 'index.html')

LP_PATH        = os.path.join(DIR, 'LP 10 CARBONES 01.04.2026.xlsx')
STOCK_MC_PATH  = os.path.join(DIR, 'STOCK_MONCAR.xlsx')
STOCK_RMH_PATH = os.path.join(DIR, 'STOCK_RMH.xlsx')

os.makedirs(OUT_DIR, exist_ok=True)

# Marcas principales que se muestran como tarjetas en la home
MARCAS_DESTACADAS = [
    'Bosch', 'De Walt', 'Makita', 'Black Y Decker',
    'Metabo', 'Hitachi', 'Stanley', 'Gamma',
    'Milwaukee', 'Skil',
]

# Categorias: keyword en descripcion UPPERCASE -> nombre
CATEGORIAS_KW = [
    ('Amoladoras',         ['AMOLADORA', 'AMOL.', 'AMOL ', 'ESMERIL']),
    ('Taladros',           ['TALADRO', 'TAL.', 'TAL ']),
    ('Lijadoras',          ['LIJADORA', 'LIJADORAS', 'LIJA', 'CEPILLO', 'CEPILLOS', 'PULIDORA']),
    ('Sierras',            ['SIERRA', 'CALADORA', 'TRONZADORA']),
    ('Lavarropas',         ['LAVARROPAS', 'LAVADORA']),
    ('Hidrolavadoras',     ['HIDROLAVADORA', 'HIDROLAV']),
    ('Aspiradoras',        ['ASPIRADORA', 'ASPIRADORAS', 'LUSTRA', 'BARRE']),
    ('Electrodomesticos',  ['BATIDORA', 'LICUADORA', 'PICADORA', 'SECADOR',
                            'VAPORIZADOR', 'ENCERADORAS', 'AFEITADORA', 'VENTILADOR',
                            'MAQUINA DE COSER', 'COSER']),
    ('Equipos de Pintura', ['PINTURA', 'COMPRESOR', 'BORDEADORA']),
]

CATEGORIA_OTROS = 'Otros'

def detectar_categoria(desc, cod, marca_brand=''):
    if cod.startswith('CE'):
        return 'Electrodomesticos'
    d = desc.upper()
    m = marca_brand.upper()
    for cat, kws in CATEGORIAS_KW:
        for kw in kws:
            if kw in d or kw in m:
                return cat
    return CATEGORIA_OTROS

# ---------------------------------------------------------------------------
# Lectores de Excel
# ---------------------------------------------------------------------------
def leer_stock_moncar():
    wb = openpyxl.load_workbook(STOCK_MC_PATH, data_only=True)
    ws = wb['Productos']
    stock = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        cod = row[0]
        if not cod or not str(cod).startswith(('CC', 'CF', 'CE')):
            continue
        cod = str(cod).strip()
        vals = [v for v in row[3:7] if isinstance(v, (int, float))]
        stock[cod] = sum(vals) > 0
    return stock

def leer_rmh():
    wb = openpyxl.load_workbook(STOCK_RMH_PATH, data_only=True)
    ws = wb['Productos']
    productos = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        cod = row[0]
        if not cod or not str(cod).strip().startswith('RM'):
            continue
        cod    = str(cod).strip()
        desc   = str(row[1]).strip() if row[1] else ''
        precio = row[2] if isinstance(row[2], (int, float)) else 0
        vals   = [v for v in row[3:7] if isinstance(v, (int, float))]
        cat    = detectar_categoria(desc, cod)
        productos.append({
            'cod':   cod,
            'linea': 'RMH',
            'marca': 'RMH',
            'cat':   cat,
            'desc':  desc,
            'dims':  '',
            'precio': int(round(precio)),
            'stock': sum(vals) > 0,
        })
    return productos

def _es_separador_marca(row):
    return (row[0] is None
            and row[6] == 0
            and isinstance(row[5], str)
            and len(row[5]) <= 40)

def _dims(esp, anch, larg):
    partes = []
    for v in (esp, anch, larg):
        if v is not None and str(v).strip() not in ('', 'None', 'DIAM'):
            partes.append(str(v).strip())
    return 'x'.join(partes) if partes else ''

def leer_moncar_lp(stock_map):
    wb = openpyxl.load_workbook(LP_PATH, data_only=True)
    productos = []

    ws = wb['Precios de Carbones']
    marca_actual = 'MONCAR'
    for row in ws.iter_rows(min_row=8, values_only=True):
        if all(v is None for v in row):
            continue
        if _es_separador_marca(row):
            marca_actual = str(row[5]).strip().title()
            continue
        cod = row[0]
        if not cod or not str(cod).strip().startswith(('CF', 'CC')):
            continue
        cod    = str(cod).strip()
        precio = row[6] if isinstance(row[6], (int, float)) and row[6] > 0 else 0
        desc   = str(row[5]).strip() if row[5] else ''
        dims   = _dims(row[2], row[3], row[4])
        cat    = detectar_categoria(desc, cod, marca_actual)
        productos.append({
            'cod':   cod,
            'linea': 'MONCAR',
            'marca': marca_actual,
            'cat':   cat,
            'desc':  desc,
            'dims':  dims,
            'precio': int(round(precio)),
            'stock': stock_map.get(cod, False),
        })

    ws2 = wb['Carbones Electrodomesticos']
    for row in ws2.iter_rows(min_row=2, values_only=True):
        cod = row[0]
        if not cod or not str(cod).strip().startswith('CE'):
            continue
        cod    = str(cod).strip()
        precio = row[7] if isinstance(row[7], (int, float)) and row[7] > 0 else 0
        desc   = str(row[6]).strip() if row[6] else ''
        dims   = _dims(row[3], row[4], row[5])
        productos.append({
            'cod':   cod,
            'linea': 'MONCAR',
            'marca': 'Electrodomesticos',
            'cat':   'Electrodomesticos',
            'desc':  desc,
            'dims':  dims,
            'precio': int(round(precio)),
            'stock': stock_map.get(cod, False),
        })

    return productos

# ---------------------------------------------------------------------------
# Generador HTML
# ---------------------------------------------------------------------------
def generar_html(productos):
    # Categorias con conteo
    from collections import Counter
    cat_count = Counter(p['cat'] for p in productos)
    cats_ordenadas = [c for c, _ in sorted(cat_count.items(), key=lambda x: -x[1])]

    # Marcas MONCAR con conteo
    marca_count = Counter(
        p['marca'] for p in productos
        if p['linea'] == 'MONCAR' and p['marca'] not in ('Moncar', 'Electrodomesticos')
    )
    top_marcas = [m for m in MARCAS_DESTACADAS if m in marca_count]
    # Agregar marcas con muchos productos que no esten en MARCAS_DESTACADAS
    for m, cnt in sorted(marca_count.items(), key=lambda x: -x[1]):
        if m not in top_marcas and cnt >= 5:
            top_marcas.append(m)

    CAT_CONFIG = {
        'Amoladoras':         {'grad': 'linear-gradient(135deg,#FF6B00,#cc3300)', 'icon': '⚙️'},
        'Taladros':           {'grad': 'linear-gradient(135deg,#2563EB,#1e40af)', 'icon': '🔩'},
        'Lijadoras':          {'grad': 'linear-gradient(135deg,#0d9488,#0f766e)', 'icon': '🔧'},
        'Sierras':            {'grad': 'linear-gradient(135deg,#16a34a,#15803d)', 'icon': '🪚'},
        'Lavarropas':         {'grad': 'linear-gradient(135deg,#7c3aed,#6d28d9)', 'icon': '🫧'},
        'Hidrolavadoras':     {'grad': 'linear-gradient(135deg,#0284c7,#0369a1)', 'icon': '💧'},
        'Aspiradoras':        {'grad': 'linear-gradient(135deg,#dc2626,#b91c1c)', 'icon': '🌀'},
        'Electrodomesticos':  {'grad': 'linear-gradient(135deg,#d97706,#b45309)', 'icon': '🔌'},
        'Equipos de Pintura': {'grad': 'linear-gradient(135deg,#db2777,#be185d)', 'icon': '🎨'},
        'Otros':              {'grad': 'linear-gradient(135deg,#374151,#1f2937)', 'icon': '🛠️'},
    }

    def cat_cards():
        cards = []
        for cat in cats_ordenadas:
            cfg   = CAT_CONFIG.get(cat, CAT_CONFIG['Otros'])
            count = cat_count[cat]
            cards.append(
                f'<div class="cat-card" style="background:{cfg["grad"]}" '
                f'onclick="irACat(\'{cat}\')">'
                f'<div class="cat-icon">{cfg["icon"]}</div>'
                f'<div class="cat-name">{cat}</div>'
                f'<div class="cat-count">{count} productos</div>'
                f'</div>'
            )
        return '\n'.join(cards)

    def marca_cards():
        MARCA_COLORS = [
            '#2563EB','#16a34a','#dc2626','#7c3aed',
            '#d97706','#0284c7','#db2777','#374151',
            '#FF6B00','#0d9488',
        ]
        cards = []
        for i, m in enumerate(top_marcas):
            color = MARCA_COLORS[i % len(MARCA_COLORS)]
            cards.append(
                f'<div class="marca-card" style="border-color:{color};color:{color}" '
                f'onclick="irAMarca(\'{m}\')">{m}</div>'
            )
        return '\n'.join(cards)

    productos_json = json.dumps(productos, ensure_ascii=False)

    return f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MONCAR RMH — Repuestos para Herramientas</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#0f0f0f;--surface:#1a1a1a;--surface2:#222;--border:#2a2a2a;
  --orange:#FF6B00;--orange2:#e05e00;
  --text:#f0f0f0;--muted:#888;--white:#fff;
  --green:#22c55e;--red:#ef4444;
}}
html{{scroll-behavior:smooth}}
body{{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}}

/* NAV */
nav{{position:fixed;top:0;left:0;right:0;z-index:100;background:rgba(15,15,15,.95);backdrop-filter:blur(10px);border-bottom:1px solid var(--border);height:60px;display:flex;align-items:center;padding:0 24px;gap:32px}}
.nav-logo{{font-size:20px;font-weight:900;letter-spacing:1px;cursor:pointer}}
.nav-logo span{{color:var(--orange)}}
.nav-links{{display:flex;gap:24px;flex:1}}
.nav-links a{{color:var(--muted);text-decoration:none;font-size:14px;font-weight:500;cursor:pointer;transition:color .15s}}
.nav-links a:hover,.nav-links a.active{{color:var(--white)}}
.nav-right{{display:flex;gap:10px;align-items:center;margin-left:auto}}
.btn-wa-nav{{background:#25D366;color:#fff;border:none;padding:8px 16px;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;text-decoration:none;display:flex;align-items:center;gap:5px}}
.btn-cart-nav{{background:var(--orange);color:#fff;border:none;padding:8px 16px;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;gap:6px}}
.btn-cart-nav:hover{{background:var(--orange2)}}
#cart-count{{background:#fff;color:var(--orange);border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:900}}

/* SECTIONS */
section{{padding-top:60px}}
.page{{display:none}}
.page.active{{display:block}}

/* ========== INICIO ========== */
/* HERO */
.hero{{min-height:420px;background:linear-gradient(135deg,#0f0f0f 0%,#1a0a00 50%,#0f0f0f 100%);display:flex;align-items:center;justify-content:center;text-align:center;padding:60px 20px;position:relative;overflow:hidden}}
.hero::before{{content:'';position:absolute;inset:0;background:radial-gradient(ellipse at 60% 50%,rgba(255,107,0,.15) 0%,transparent 70%)}}
.hero-content{{position:relative;z-index:1;max-width:680px}}
.hero-eyebrow{{display:inline-block;background:rgba(255,107,0,.15);color:var(--orange);border:1px solid rgba(255,107,0,.3);border-radius:20px;padding:5px 16px;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:16px}}
.hero h1{{font-size:clamp(32px,5vw,52px);font-weight:900;line-height:1.1;margin-bottom:16px}}
.hero h1 span{{color:var(--orange)}}
.hero p{{color:var(--muted);font-size:17px;margin-bottom:28px;line-height:1.5}}
.hero-btns{{display:flex;gap:12px;justify-content:center;flex-wrap:wrap}}
.btn-primary{{background:var(--orange);color:#fff;border:none;padding:14px 28px;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer}}
.btn-primary:hover{{background:var(--orange2)}}
.btn-secondary{{background:transparent;color:var(--text);border:1px solid var(--border);padding:14px 28px;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer}}
.btn-secondary:hover{{border-color:var(--orange);color:var(--orange)}}

/* SECTION HEADERS */
.sec{{padding:48px 24px}}
.sec-header{{margin-bottom:28px}}
.sec-header h2{{font-size:clamp(22px,4vw,30px);font-weight:900}}
.sec-header p{{color:var(--muted);font-size:14px;margin-top:4px}}

/* CATEGORIAS */
.cat-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:14px}}
.cat-card{{border-radius:12px;padding:24px 16px;cursor:pointer;transition:transform .15s,box-shadow .15s;text-align:center}}
.cat-card:hover{{transform:translateY(-3px);box-shadow:0 8px 24px rgba(0,0,0,.4)}}
.cat-icon{{font-size:36px;margin-bottom:10px}}
.cat-name{{color:#fff;font-size:15px;font-weight:700;text-transform:uppercase;letter-spacing:.5px}}
.cat-count{{color:rgba(255,255,255,.7);font-size:12px;margin-top:4px}}

/* MARCAS */
.marca-grid{{display:flex;flex-wrap:wrap;gap:10px}}
.marca-card{{border:2px solid;border-radius:8px;padding:10px 20px;font-size:14px;font-weight:700;cursor:pointer;background:transparent;transition:background .15s}}
.marca-card:hover{{background:rgba(255,255,255,.06)}}

/* STATS */
.stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:16px;padding:0 24px 48px}}
.stat{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:20px;text-align:center}}
.stat-num{{font-size:32px;font-weight:900;color:var(--orange)}}
.stat-lbl{{font-size:13px;color:var(--muted);margin-top:4px}}

/* NOSOTROS */
.nosotros{{max-width:720px;padding:48px 24px}}
.nosotros h2{{font-size:28px;font-weight:900;margin-bottom:16px}}
.nosotros p{{color:var(--muted);line-height:1.7;margin-bottom:12px}}
.contact-cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-top:28px}}
.contact-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:20px}}
.contact-card h3{{font-size:13px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}}
.contact-card a{{color:var(--orange);text-decoration:none;font-size:15px;font-weight:600}}

/* ========== PRODUCTOS ========== */
.prod-header{{padding:24px 24px 0;border-bottom:1px solid var(--border)}}
.prod-header h2{{font-size:22px;font-weight:800;margin-bottom:4px}}
.prod-header p{{color:var(--muted);font-size:13px;margin-bottom:16px}}
.filters{{display:flex;flex-wrap:wrap;gap:10px;padding:16px 24px;border-bottom:1px solid var(--border);align-items:center}}
.search-wrap{{position:relative;flex:1;min-width:200px}}
.search-wrap input{{width:100%;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:9px 36px 9px 12px;color:var(--text);font-size:14px}}
.search-wrap input:focus{{outline:none;border-color:var(--orange)}}
.search-wrap input::placeholder{{color:var(--muted)}}
.search-clear{{position:absolute;right:10px;top:50%;transform:translateY(-50%);background:none;border:none;color:var(--muted);cursor:pointer;font-size:15px;display:none}}
.filter-group{{display:flex;gap:6px;flex-wrap:wrap}}
.filt{{background:var(--surface);border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 12px;font-size:12px;cursor:pointer;transition:.15s;white-space:nowrap}}
.filt:hover,.filt.active{{background:var(--orange);border-color:var(--orange);color:#fff}}
.filt.active-blue{{background:#2563EB;border-color:#2563EB;color:#fff}}
.toggle-stock{{display:flex;align-items:center;gap:6px;font-size:13px;color:var(--muted);cursor:pointer;white-space:nowrap}}
.toggle-stock input{{accent-color:var(--orange)}}
.active-filters{{padding:8px 24px;display:flex;gap:8px;flex-wrap:wrap;display:none}}
.active-tag{{background:rgba(255,107,0,.15);border:1px solid rgba(255,107,0,.3);color:var(--orange);border-radius:20px;padding:3px 12px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:6px;cursor:pointer}}
.grid-wrap{{padding:20px 24px}}
.results-info{{font-size:13px;color:var(--muted);margin-bottom:14px}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:14px}}
.card{{background:var(--surface2);border:1px solid var(--border);border-radius:10px;overflow:hidden;display:flex;flex-direction:column;transition:border-color .15s}}
.card:hover{{border-color:var(--orange)}}
.card-icon{{background:var(--surface);display:flex;align-items:center;justify-content:center;padding:20px;height:100px}}
.card-body{{padding:12px;flex:1;display:flex;flex-direction:column;gap:5px}}
.card-badges{{display:flex;gap:5px;flex-wrap:wrap}}
.badge{{display:inline-block;font-size:10px;border-radius:4px;padding:2px 6px;font-weight:700}}
.badge-linea{{background:rgba(255,107,0,.15);color:var(--orange)}}
.badge-cat{{background:rgba(255,255,255,.07);color:var(--muted)}}
.card-cod{{font-size:11px;font-weight:700;color:var(--orange);letter-spacing:.5px;margin-top:2px}}
.card-desc{{font-size:13px;color:var(--text);line-height:1.35;flex:1}}
.card-sub{{font-size:11px;color:var(--muted)}}
.card-precio{{font-size:17px;font-weight:800;color:var(--white);margin-top:4px}}
.card-stock{{font-size:11px;font-weight:600;padding:2px 8px;border-radius:4px;align-self:flex-start}}
.stock-ok{{background:rgba(34,197,94,.15);color:var(--green)}}
.stock-no{{background:rgba(239,68,68,.12);color:var(--red)}}
.btn-add{{margin-top:8px;background:var(--orange);color:#fff;border:none;border-radius:6px;padding:8px;font-size:13px;font-weight:600;cursor:pointer;width:100%;transition:.15s}}
.btn-add:hover{{background:var(--orange2)}}
.btn-add:disabled{{background:#333;color:var(--muted);cursor:not-allowed}}
.no-results{{text-align:center;padding:60px 20px;color:var(--muted)}}

/* CART SIDEBAR */
.cart-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:200}}
.cart-panel{{position:fixed;top:0;right:-420px;width:min(420px,100vw);height:100%;background:var(--surface);border-left:1px solid var(--border);z-index:201;display:flex;flex-direction:column;transition:right .25s ease}}
.cart-panel.open{{right:0}}
.cart-header{{padding:18px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between}}
.cart-header h2{{font-size:16px;font-weight:700}}
.btn-close-cart{{background:none;border:none;color:var(--muted);font-size:22px;cursor:pointer}}
.cart-items{{flex:1;overflow-y:auto;padding:12px}}
.cart-empty{{text-align:center;padding:40px 20px;color:var(--muted)}}
.cart-item{{display:flex;gap:10px;align-items:flex-start;padding:10px 0;border-bottom:1px solid var(--border)}}
.cart-item-info{{flex:1}}
.cart-item-cod{{font-size:11px;color:var(--orange);font-weight:700}}
.cart-item-desc{{font-size:13px;margin:2px 0}}
.cart-item-precio{{font-size:14px;font-weight:700}}
.cart-item-qty{{display:flex;align-items:center;gap:6px;margin-top:6px}}
.cart-item-qty button{{background:var(--border);border:none;color:var(--text);width:24px;height:24px;border-radius:4px;cursor:pointer;font-size:15px}}
.cart-item-qty span{{font-size:14px;font-weight:700;min-width:20px;text-align:center}}
.btn-remove-item{{background:none;border:none;color:var(--red);cursor:pointer;font-size:18px;padding:0 4px}}
.cart-footer{{padding:16px 20px;border-top:1px solid var(--border)}}
.cart-total{{display:flex;justify-content:space-between;font-size:16px;font-weight:700;margin-bottom:12px}}
.btn-wa-cart{{background:#25D366;color:#fff;border:none;border-radius:8px;padding:12px;font-size:15px;font-weight:700;cursor:pointer;width:100%;margin-bottom:8px}}
.btn-email-cart{{background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:12px;font-size:14px;font-weight:600;cursor:pointer;width:100%}}
.btn-email-cart:hover{{border-color:var(--orange)}}

/* FOOTER */
footer{{padding:32px 24px;border-top:1px solid var(--border);display:flex;flex-wrap:wrap;gap:24px;justify-content:space-between;color:var(--muted);font-size:13px}}
footer .ft-brand{{font-size:18px;font-weight:900;color:var(--text)}}
footer .ft-brand span{{color:var(--orange)}}
footer a{{color:var(--orange);text-decoration:none}}
.ft-copy{{width:100%;text-align:center;padding-top:16px;border-top:1px solid var(--border);color:var(--muted);font-size:12px}}

@media(max-width:640px){{
  nav{{padding:0 14px;gap:16px}}
  .nav-links{{gap:14px}}
  .sec{{padding:32px 16px}}
  .cat-grid{{grid-template-columns:repeat(2,1fr)}}
  .filters{{padding:12px 16px}}
  .grid-wrap{{padding:16px}}
  .grid{{grid-template-columns:repeat(2,1fr);gap:10px}}
}}
</style>
</head>
<body>

<!-- NAV -->
<nav>
  <div class="nav-logo" onclick="irAInicio()">MONCAR<span> RMH</span></div>
  <div class="nav-links">
    <a id="nav-inicio" class="active" onclick="irAInicio()">Inicio</a>
    <a id="nav-productos" onclick="irAProductos()">Productos</a>
    <a id="nav-nosotros" onclick="irANosotros()">Nosotros</a>
  </div>
  <div class="nav-right">
    <a class="btn-wa-nav" href="https://wa.me/5491136267009" target="_blank">&#128172; WhatsApp</a>
    <button class="btn-cart-nav" onclick="abrirCart()">
      Carrito <span id="cart-count">0</span>
    </button>
  </div>
</nav>

<!-- ===== PAGINA: INICIO ===== -->
<div id="page-inicio" class="page active">

  <!-- HERO -->
  <section class="hero">
    <div class="hero-content">
      <div class="hero-eyebrow">Repuestos originales</div>
      <h1>Carbones y repuestos<br>para tus <span>herramientas</span></h1>
      <p>Carbones electricos MONCAR y interruptores RMH.<br>Envios a todo el pais.</p>
      <div class="hero-btns">
        <button class="btn-primary" onclick="irAProductos()">Ver catalogo completo</button>
        <a class="btn-secondary" href="https://wa.me/5491136267009" target="_blank" style="display:inline-flex;align-items:center;gap:6px;text-decoration:none;padding:14px 28px;border-radius:8px;border:1px solid var(--border);color:var(--text);font-size:15px;font-weight:600">Consultar por WhatsApp</a>
      </div>
    </div>
  </section>

  <!-- STATS -->
  <div class="stats">
    <div class="stat"><div class="stat-num">{len([p for p in productos if p['linea']=='MONCAR'])}</div><div class="stat-lbl">Carbones MONCAR</div></div>
    <div class="stat"><div class="stat-num">{len([p for p in productos if p['linea']=='RMH'])}</div><div class="stat-lbl">Interruptores RMH</div></div>
    <div class="stat"><div class="stat-num">{sum(1 for p in productos if p['stock'])}</div><div class="stat-lbl">Productos en stock</div></div>
    <div class="stat"><div class="stat-num">{len(set(p['marca'] for p in productos))}</div><div class="stat-lbl">Marcas compatibles</div></div>
  </div>

  <!-- CATEGORIAS -->
  <div class="sec">
    <div class="sec-header">
      <h2>Compra por categoria</h2>
      <p>Encontra rapido el repuesto para tu herramienta</p>
    </div>
    <div class="cat-grid">
      {cat_cards()}
    </div>
  </div>

  <!-- MARCAS -->
  <div class="sec" style="padding-top:0">
    <div class="sec-header">
      <h2>Compra por marca</h2>
      <p>Compatible con las principales marcas del mercado</p>
    </div>
    <div class="marca-grid">
      {marca_cards()}
      <div class="marca-card" style="border-color:#888;color:#888" onclick="irAProductos()">+ Ver todas</div>
    </div>
  </div>

</div><!-- /page-inicio -->

<!-- ===== PAGINA: PRODUCTOS ===== -->
<div id="page-productos" class="page">

  <div class="prod-header">
    <h2>Catalogo de productos</h2>
    <p>Carbones electricos MONCAR e interruptores RMH</p>
  </div>

  <div class="filters">
    <div class="search-wrap">
      <input type="text" id="search" placeholder="Buscar por codigo o descripcion..." oninput="filtrar()">
      <button class="search-clear" id="clear-btn" onclick="limpiarBusqueda()">&#x2715;</button>
    </div>
    <div class="filter-group" id="linea-group">
      <button class="filt active" data-val="" onclick="setFiltro(this,'linea')">Todos</button>
      <button class="filt" data-val="MONCAR" onclick="setFiltro(this,'linea')">MONCAR</button>
      <button class="filt" data-val="RMH" onclick="setFiltro(this,'linea')">RMH</button>
    </div>
    <label class="toggle-stock">
      <input type="checkbox" id="solo-stock" onchange="filtrar()">
      Solo en stock
    </label>
  </div>

  <div class="active-filters" id="active-filters">
    <!-- tags de filtros activos -->
  </div>

  <div class="grid-wrap">
    <div class="results-info" id="results-info"></div>
    <div class="grid" id="grid"></div>
    <div class="no-results" id="no-results" style="display:none">
      <p style="font-size:18px;margin-bottom:8px">Sin resultados</p>
      <span>Proba con otro termino o quitá los filtros</span>
    </div>
  </div>

</div><!-- /page-productos -->

<!-- ===== PAGINA: NOSOTROS ===== -->
<div id="page-nosotros" class="page">
  <div class="nosotros">
    <h2>Sobre MONCAR RMH</h2>
    <p>Somos una empresa familiar especializada en repuestos electricos para herramientas. Comercializamos carbones electricos de la linea MONCAR y los interruptores RMH para las principales marcas del mercado.</p>
    <p>Con anos de experiencia en el rubro, brindamos atencion personalizada y envios a todo el pais. Cada repuesto que vendemos pasa por un control de calidad para asegurar el mejor funcionamiento de tus herramientas.</p>
    <div class="contact-cards">
      <div class="contact-card">
        <h3>WhatsApp</h3>
        <a href="https://wa.me/5491136267009" target="_blank">+54 9 11 3626-7009</a>
      </div>
      <div class="contact-card">
        <h3>Email</h3>
        <a href="mailto:contacto.carbonkey@gmail.com">contacto.carbonkey@gmail.com</a>
      </div>
    </div>
  </div>
</div><!-- /page-nosotros -->

<!-- FOOTER -->
<footer>
  <div>
    <div class="ft-brand">MONCAR<span> RMH</span></div>
    <div style="margin-top:8px">Carbones electricos e interruptores para herramientas</div>
  </div>
  <div>
    <div>WhatsApp: <a href="https://wa.me/5491136267009">+54 9 11 3626-7009</a></div>
    <div style="margin-top:4px">Email: <a href="mailto:contacto.carbonkey@gmail.com">contacto.carbonkey@gmail.com</a></div>
  </div>
  <div class="ft-copy">© 2025 MONCAR RMH — Todos los derechos reservados</div>
</footer>

<!-- CART SIDEBAR -->
<div class="cart-overlay" id="overlay" onclick="cerrarCart()"></div>
<div class="cart-panel" id="cart-panel">
  <div class="cart-header">
    <h2>Tu pedido</h2>
    <button class="btn-close-cart" onclick="cerrarCart()">&#x2715;</button>
  </div>
  <div class="cart-items" id="cart-items"></div>
  <div class="cart-footer">
    <div class="cart-total"><span>Total estimado</span><span id="cart-total">$0</span></div>
    <button class="btn-wa-cart" onclick="pedirPorWA()">&#128172; Pedir por WhatsApp</button>
    <button class="btn-email-cart" onclick="pedirPorEmail()">Pedir por Email</button>
  </div>
</div>

<script>
const PRODUCTOS = {productos_json};

const SVG_MC = `<svg width="56" height="56" viewBox="0 0 56 56" fill="none">
  <circle cx="28" cy="28" r="22" fill="#FF6B00" opacity=".1"/>
  <circle cx="28" cy="28" r="10" fill="none" stroke="#FF6B00" stroke-width="2.5" opacity=".6"/>
  <rect x="11" y="26.5" width="34" height="3" rx="1.5" fill="#FF6B00" opacity=".4"/>
  <rect x="26.5" y="11" width="3" height="34" rx="1.5" fill="#FF6B00" opacity=".4"/>
  <circle cx="28" cy="28" r="3.5" fill="#FF6B00" opacity=".8"/>
</svg>`;
const SVG_RMH = `<svg width="56" height="56" viewBox="0 0 56 56" fill="none">
  <rect x="12" y="16" width="32" height="24" rx="4" fill="#FF6B00" opacity=".1" stroke="#FF6B00" stroke-width="1.5" opacity=".4"/>
  <rect x="18" y="22" width="20" height="12" rx="2" fill="#FF6B00" opacity=".2"/>
  <circle cx="28" cy="28" r="3" fill="#FF6B00" opacity=".7"/>
  <path d="M22 28h3M31 28h3" stroke="#FF6B00" stroke-width="2" stroke-linecap="round" opacity=".6"/>
</svg>`;

let cart = JSON.parse(localStorage.getItem('moncar_cart') || '[]');
let filtroLinea = '';
let filtroCat = '';
let filtroMarca = '';

function fmt(n) {{
  if (!n) return 'Consultar';
  return '$' + Math.round(n).toLocaleString('es-AR');
}}

// NAV
function setNavActive(id) {{
  ['inicio','productos','nosotros'].forEach(n => {{
    const el = document.getElementById('nav-'+n);
    if (el) el.classList.toggle('active', n === id);
  }});
}}

function irAInicio() {{
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-inicio').classList.add('active');
  setNavActive('inicio');
  window.scrollTo(0,0);
}}

function irAProductos(cat, marca) {{
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-productos').classList.add('active');
  setNavActive('productos');
  window.scrollTo(0,0);
  if (cat !== undefined) {{
    filtroCat = cat || '';
    filtroMarca = '';
  }} else if (marca !== undefined) {{
    filtroMarca = marca || '';
    filtroCat = '';
  }}
  filtrar();
}}

function irANosotros() {{
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-nosotros').classList.add('active');
  setNavActive('nosotros');
  window.scrollTo(0,0);
}}

function irACat(cat) {{ irAProductos(cat, undefined); }}
function irAMarca(m) {{ irAProductos(undefined, m); }}

// FILTROS
function setFiltro(btn, tipo) {{
  const group = tipo === 'linea' ? 'linea-group' : null;
  if (group) {{
    document.querySelectorAll('#'+group+' .filt').forEach(b => b.classList.remove('active'));
  }}
  btn.classList.add('active');
  if (tipo === 'linea') filtroLinea = btn.dataset.val;
  filtrar();
}}

function filtrar() {{
  const q = document.getElementById('search').value.toLowerCase().trim();
  const soloStock = document.getElementById('solo-stock').checked;
  document.getElementById('clear-btn').style.display = q ? 'block' : 'none';

  // Active filter tags
  const af = document.getElementById('active-filters');
  const tags = [];
  if (filtroCat) tags.push(`<div class="active-tag" onclick="filtroCat='';filtrar()">📂 ${{filtroCat}} &#x2715;</div>`);
  if (filtroMarca) tags.push(`<div class="active-tag" onclick="filtroMarca='';filtrar()">🏷️ ${{filtroMarca}} &#x2715;</div>`);
  af.innerHTML = tags.join('');
  af.style.display = tags.length ? 'flex' : 'none';

  const visibles = PRODUCTOS.filter(p => {{
    if (filtroLinea && p.linea !== filtroLinea) return false;
    if (filtroCat && p.cat !== filtroCat) return false;
    if (filtroMarca && p.marca.toLowerCase() !== filtroMarca.toLowerCase()) return false;
    if (soloStock && !p.stock) return false;
    if (q) {{
      const hay = (p.cod + ' ' + p.desc + ' ' + p.marca + ' ' + p.cat).toLowerCase().includes(q);
      if (!hay) return false;
    }}
    return true;
  }});

  renderGrid(visibles);
  document.getElementById('results-info').textContent =
    visibles.length + ' producto' + (visibles.length !== 1 ? 's' : '');
  document.getElementById('no-results').style.display = visibles.length ? 'none' : 'block';
  document.getElementById('grid').style.display = visibles.length ? 'grid' : 'none';
}}

function renderGrid(lista) {{
  const g = document.getElementById('grid');
  g.innerHTML = lista.map(p => {{
    const svg = p.linea === 'RMH' ? SVG_RMH : SVG_MC;
    const stockCls = p.stock ? 'stock-ok' : 'stock-no';
    const stockTxt = p.stock ? 'En stock' : 'Sin stock';
    const enCart = cart.find(c => c.cod === p.cod);
    const btnTxt = p.stock ? (enCart ? 'Agregar mas' : 'Agregar') : 'Sin stock';
    const marcaTxt = p.marca && p.marca !== 'RMH' && p.marca !== 'Moncar' ? p.marca : '';
    return `<div class="card">
      <div class="card-icon">${{svg}}</div>
      <div class="card-body">
        <div class="card-badges">
          <span class="badge badge-linea">${{p.linea}}</span>
          <span class="badge badge-cat">${{p.cat}}</span>
        </div>
        <div class="card-cod">${{p.cod}}</div>
        <div class="card-desc">${{p.desc}}</div>
        ${{marcaTxt ? `<div class="card-sub">${{marcaTxt}}</div>` : ''}}
        ${{p.dims ? `<div class="card-sub">${{p.dims}} mm</div>` : ''}}
        <div class="card-precio">${{fmt(p.precio)}}</div>
        <span class="card-stock ${{stockCls}}">${{stockTxt}}</span>
        <button class="btn-add" ${{p.stock ? '' : 'disabled'}} onclick="agregarAlCarrito('${{p.cod}}')">${{btnTxt}}</button>
      </div>
    </div>`;
  }}).join('');
}}

function limpiarBusqueda() {{
  document.getElementById('search').value = '';
  filtrar();
}}

// CART
function agregarAlCarrito(cod) {{
  const p = PRODUCTOS.find(x => x.cod === cod);
  if (!p || !p.stock) return;
  const ex = cart.find(c => c.cod === cod);
  if (ex) {{ ex.qty++; }} else {{ cart.push({{cod:p.cod, desc:p.desc, precio:p.precio, qty:1}}); }}
  guardarCart();
  filtrar();
  abrirCart();
}}

function cambiarQty(cod, delta) {{
  const ex = cart.find(c => c.cod === cod);
  if (!ex) return;
  ex.qty += delta;
  if (ex.qty <= 0) cart = cart.filter(c => c.cod !== cod);
  guardarCart();
}}

function quitarItem(cod) {{
  cart = cart.filter(c => c.cod !== cod);
  guardarCart();
}}

function guardarCart() {{
  localStorage.setItem('moncar_cart', JSON.stringify(cart));
  renderCart();
  document.getElementById('cart-count').textContent = cart.reduce((s,c) => s+c.qty, 0);
}}

function renderCart() {{
  const el = document.getElementById('cart-items');
  if (!cart.length) {{
    el.innerHTML = '<div class="cart-empty">El carrito esta vacio</div>';
    document.getElementById('cart-total').textContent = '$0';
    return;
  }}
  el.innerHTML = cart.map(c => `
    <div class="cart-item">
      <div class="cart-item-info">
        <div class="cart-item-cod">${{c.cod}}</div>
        <div class="cart-item-desc">${{c.desc}}</div>
        <div class="cart-item-precio">${{fmt(c.precio)}}</div>
        <div class="cart-item-qty">
          <button onclick="cambiarQty('${{c.cod}}',-1);renderCart();filtrar()">-</button>
          <span>${{c.qty}}</span>
          <button onclick="cambiarQty('${{c.cod}}',1);renderCart();filtrar()">+</button>
        </div>
      </div>
      <button class="btn-remove-item" onclick="quitarItem('${{c.cod}}');renderCart();filtrar()">&#x2715;</button>
    </div>`).join('');
  const total = cart.reduce((s,c) => s + c.precio*c.qty, 0);
  document.getElementById('cart-total').textContent = fmt(total);
}}

function abrirCart() {{
  document.getElementById('cart-panel').classList.add('open');
  document.getElementById('overlay').style.display = 'block';
  renderCart();
}}

function cerrarCart() {{
  document.getElementById('cart-panel').classList.remove('open');
  document.getElementById('overlay').style.display = 'none';
}}

function pedirPorWA() {{
  if (!cart.length) return;
  let msg = 'Hola! Quiero consultar disponibilidad de los siguientes productos:\\n\\n';
  cart.forEach(c => {{ msg += `- ${{c.cod}} x${{c.qty}}\\n`; }});
  msg += '\\nMuchas gracias!';
  window.open('https://wa.me/5491136267009?text=' + encodeURIComponent(msg), '_blank');
}}

function pedirPorEmail() {{
  if (!cart.length) return;
  let body = 'Hola, quiero consultar disponibilidad de los siguientes productos:\\n\\n';
  cart.forEach(c => {{ body += `- ${{c.cod}} ${{c.desc}} x${{c.qty}}\\n`; }});
  body += '\\nMuchas gracias!';
  window.location.href = 'mailto:contacto.carbonkey@gmail.com?subject=Consulta%20de%20productos&body=' + encodeURIComponent(body);
}}

// INIT
guardarCart();
filtrar();
</script>
</body>
</html>'''

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    print("Leyendo stock MONCAR...")
    stock_mc = leer_stock_moncar()

    print("Leyendo productos MONCAR desde LP...")
    productos_mc = leer_moncar_lp(stock_mc)
    print(f"  {len(productos_mc)} productos MONCAR")

    print("Leyendo productos RMH...")
    productos_rmh = leer_rmh()
    print(f"  {len(productos_rmh)} productos RMH")

    todos = productos_mc + productos_rmh
    print(f"Total: {len(todos)} productos")

    # Mostrar distribucion por categoria
    from collections import Counter
    cats = Counter(p['cat'] for p in todos)
    for cat, n in sorted(cats.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {n}")

    print("\nGenerando HTML...")
    html = generar_html(todos)

    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\nSitio generado: {OUT_PATH}")
    kb = len(html) // 1024
    print(f"Tamano: {kb} KB")
