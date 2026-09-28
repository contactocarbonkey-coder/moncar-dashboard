# -*- coding: utf-8 -*-
"""
generar_sitio_drive.py — genera sitio/index.html (catalogo MONCAR + RMH) leyendo
precios y stock desde los mismos Google Sheets que usa el dashboard (VENTAS,
STOCK_MONCAR, STOCK_RMH), en vez de los Excel locales viejos.

Copia de generar_sitio.py (version Excel, sin uso activo) — mismo HTML/JS,
distinta fuente de datos. Dimensiones (esp/anch/larg/terminal) de los carbones
MONCAR siguen viniendo de LP/LP 10 CARBONES 01.04.2026.xlsx porque ese detalle
todavia no esta migrado a Sheets (ver pendiente en CLAUDE.md) — no cambia
seguido, así que alcanza con leerlo del archivo local.
"""
import sys, os, json, csv, io, time
import urllib.request, urllib.parse, urllib.error
from collections import Counter

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

DIR      = os.path.dirname(os.path.abspath(__file__))
OUT_DIR  = os.path.join(DIR, 'sitio')
OUT_PATH = os.path.join(OUT_DIR, 'index.html')
IMG_DIR  = os.path.join(OUT_DIR, 'imagenes')
LP_DIM_PATH = os.path.join(os.path.dirname(DIR), 'LP', 'LP 10 CARBONES 01.04.2026.xlsx')

os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

VENTAS_SHEET_ID    = '1K8wcDYl5E6hf7B-Go-R-3BPnCISNFKibnEV2C4AmfOE'
STOCK_MC_SHEET_ID  = '1u9xG2uyeT3V8osBi9Qz0xofDCPmFmhOXlu85c1OIEQ8'
STOCK_RMH_SHEET_ID = '1Msl0CC5abCoIaM0MpeCtYOoHVHI0mjPlfElc5x8ytJE'

MARCAS_CONFIG = {
    'Bosch':          {'bg': '#007BC0', 'text': '#ffffff'},
    'De Walt':        {'bg': '#FFCD00', 'text': '#000000'},
    'Makita':         {'bg': '#00A8E0', 'text': '#ffffff'},
    'Black Y Decker': {'bg': '#E31837', 'text': '#ffffff'},
    'Metabo':         {'bg': '#003082', 'text': '#ffffff'},
    'Hitachi':        {'bg': '#E4002B', 'text': '#ffffff'},
    'Stanley':        {'bg': '#FFC72C', 'text': '#000000'},
    'Gamma':          {'bg': '#FF6B00', 'text': '#ffffff'},
    'Milwaukee':      {'bg': '#CC0000', 'text': '#ffffff'},
    'Skil':           {'bg': '#003087', 'text': '#ffffff'},
    'Dremel':         {'bg': '#1BA0E1', 'text': '#ffffff'},
    'Einhell':        {'bg': '#E20000', 'text': '#ffffff'},
}

# Iconos: la mayoria son SVG reales (fill="currentColor", toman color/tamano
# de .cat-icon via CSS) sacados de Material Symbols / MDI / Game-Icons via
# Iconify — mucho mas especificos que los emoji viejos. Para Amoladoras,
# Electrodomesticos y Otros no existe un icono de herramienta especifico en
# ningun set libre (se busco a proposito: "angle grinder", "grinding wheel",
# "disc sander" no dieron resultados) — se usa un icono generico de
# Font Awesome (clase fa-solid) en su lugar, por eso se ve distinto (<i> vs
# <svg>). cat_cards() inserta el valor de 'icon' tal cual, sin importar el tipo.
_SVG_TALADROS = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M6 19h6v-1H6zm.75-9h4.5q.3 0 .525-.225T12 9.25t-.225-.525t-.525-.225h-4.5q-.3 0-.525.225T6 9.25t.225.525t.525.225m0-2.5h4.5q.3 0 .525-.225T12 6.75t-.225-.525T11.25 6h-4.5q-.3 0-.525.225T6 6.75t.225.525t.525.225M16 11V9h2V7h-2V5h2q.825 0 1.413.588T20 7h2q.425 0 .713.288T23 8t-.288.713T22 9h-2q0 .825-.587 1.413T18 11zm-4 5h-2v-5h4V5H6q-.825 0-1.412.588T4 7v2q0 .825.588 1.413T6 11h2v5H6v-3q-1.65 0-2.825-1.175T2 9V7q0-1.65 1.175-2.825T6 3h8q.825 0 1.413.588T16 5v6q0 .825-.587 1.413T14 13h-2zm-6.5 5q-.625 0-1.062-.437T4 19.5v-2q0-.625.438-1.062T5.5 16h7q.625 0 1.063.438T14 17.5v2q0 .625-.437 1.063T12.5 21zm6.5-2H6z"/></svg>'
_SVG_SIERRAS  = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M13 11a1 1 0 0 1 1 1a1 1 0 0 1-1 1a1 1 0 0 1-1-1a1 1 0 0 1 1-1M7.86 6.25A7 7 0 0 1 13 4c3.5 0 6.44 2.61 6.93 6H22v2h-6a3 3 0 0 0-3-3a3 3 0 0 0-3 3H2v-2h.05c.2-2.27 1.09-4.34 2.45-6zM6.73 7.89L5.06 6.77c-.53.98-.88 2.07-1 3.23h2.01c.11-.75.33-1.46.66-2.11m.67 7.51L6 14h5.79c.24.42.71.7 1.21.7s.97-.28 1.21-.7H20v1.4c-1.61-.98-1.54.35-1.54.35v1.96l-1.96 1.96c-.5-1.75-1.4-.77-1.4-.77l-1.4 1.4h-2.8c.98-1.61-.35-1.54-.35-1.54H8.59L6.63 16.8c1.75-.49.77-1.4.77-1.4"/></svg>'
_SVG_COSER    = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><path fill="currentColor" d="m61.67 119.8l-12.69 68.3c-5.47 29.6-1.02 60.1 12.69 86.8h51.03V119.8zm68.63 0v115.6h2.9c6.2 0 11.2-5.1 11.2-11.3V119.8zm31.8 0v104.3c0 15.9-12.9 28.8-28.9 28.8h-2.9v81.7h14.9v-59.7h19.6s7.6-15.3 13.6-27.7c4.7-9.2 14.1-15.1 24.5-15.1h80.2c15 0 27.2 12.1 27.2 27.1v91.5H25.95l10.57 68.5H455.1l10.5-63.1l-21.1-236.3zm310.1 23.4c-4.9.4-8.5 4.7-8 9.5l4.3 47.4c.4 4.8 4.7 8.4 9.5 8c4.9-.4 8.4-4.8 8-9.6l-4.2-47.3c-.4-4.5-4.2-8-8.7-8zm-86.8.1c21.7 0 39.2 17.5 39.2 39.2s-17.5 39.2-39.2 39.2s-39.2-17.5-39.2-39.2s17.5-39.2 39.2-39.2m0 17.5c-11.9 0-21.7 9.8-21.7 21.7c0 12 9.8 21.7 21.7 21.7s21.7-9.7 21.7-21.7c0-11.9-9.8-21.7-21.7-21.7m0 89.9c26.6 0 48.2 21.7 48.2 48.2c0 26.6-21.6 48.2-48.2 48.2s-48.2-21.6-48.2-48.2c0-26.5 21.6-48.2 48.2-48.2m0 17.6c-16.9 0-30.6 13.7-30.6 30.6s13.7 30.5 30.6 30.5s30.5-13.6 30.5-30.5s-13.6-30.6-30.5-30.6M92.71 287.8v17.6h20.09v-17.6zM36.52 438.9v26.5H454.9v-26.5z"/></svg>'
_SVG_ASPIRAD  = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M1.875 21.125Q1 20.25 1 19t.875-2.125T4 16t2.125.875T7 19t-.875 2.125T4 22t-2.125-.875m2.838-1.412Q5 19.425 5 19t-.288-.712T4 18t-.712.288T3 19t.288.713T4 20t.713-.288M8 22q.35-.45.563-.937T8.9 20H11v-7q0-.825-.587-1.412T9 11H4v3q-.525 0-1.037.1T2 14.4V9h3V5.6q0-1.925 1.338-3.262T9.6 1q1.4 0 2.55.763t1.7 2.037L20.675 20H23v2h-7v-2h2.5L12.025 4.6q-.3-.725-.962-1.162T9.6 3q-1.1 0-1.85.75T7 5.6V9h2q1.65 0 2.825 1.175T13 13v9zm-.5-6.5"/></svg>'
_SVG_HIDRO    = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 17h4.1q-.425-.425-.787-.925T5.675 15H3zm12.538-1.463Q17 14.076 17 12t-1.463-3.537T12 7T8.463 8.463T7 12t1.463 3.538T12 17t3.538-1.463M18.325 9H21V7h-4.1q.425.425.788.925T18.325 9M1 20v-8h2v1h2.075q-.05-.25-.062-.488T5 12q0-2.925 2.038-4.962T12 5h9V4h2v8h-2v-1h-2.075q.05.25.063.488T19 12q0 2.925-2.037 4.963T12 19H3v1zm2-3v-2zm18-8V7zm-10.412 5.413Q10 13.825 10 13q0-.575.238-1.137t.912-1.613L12 9l.85 1.25q.675 1.05.913 1.613T14 13q0 .825-.587 1.413T12 15t-1.412-.587"/></svg>'
_SVG_LAVARR   = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M14.83 11.17a4.01 4.01 0 0 1 0 5.66a4.01 4.01 0 0 1-5.66 0zM6 2h12a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2m1 2a1 1 0 0 0-1 1a1 1 0 0 0 1 1a1 1 0 0 0 1-1a1 1 0 0 0-1-1m3 0a1 1 0 0 0-1 1a1 1 0 0 0 1 1a1 1 0 0 0 1-1a1 1 0 0 0-1-1m2 4a6 6 0 0 0-6 6a6 6 0 0 0 6 6a6 6 0 0 0 6-6a6 6 0 0 0-6-6"/></svg>'

CAT_CONFIG = {
    'Amoladoras':         {'icon': '<i class="fa-solid fa-gear"></i>'},
    'Taladros':           {'icon': _SVG_TALADROS},
    'Lijadoras':          {'icon': '<i class="fa-solid fa-layer-group"></i>'},
    'Sierras':            {'icon': _SVG_SIERRAS},
    'Lavarropas':         {'icon': _SVG_LAVARR},
    'Hidrolavadoras':     {'icon': _SVG_HIDRO},
    'Aspiradoras':        {'icon': _SVG_ASPIRAD},
    'Electrodomesticos':  {'icon': '<i class="fa-solid fa-plug"></i>'},
    'Maq. de Coser':      {'icon': _SVG_COSER},
    'Otros':              {'icon': '<i class="fa-solid fa-toolbox"></i>'},
}

EXCLUIR_MARCAS = {'Moncar','Electrodomesticos','Cobreados','Porta Carbones',
                  'Power Tools','Experto','Bta','Neo','Lazo Asincron','Versa',
                  'Lavarropas','Aspiradora','Ventilador','Batidora',
                  'Maquina De Coser','Lustra Aspiradora','Hidrolavadora',
                  'Enceradoras','Afeitadoras','Bordeadora','Barre Alfombras',
                  'Salkor','Stingray','Dover','Gladiator','Stayer','Peugeot',
                  'Dowen Paggio','Martins','Agro','Argentec','Adiabatic'}

# ---------------------------------------------------------------------------
# Helpers Google Sheets (mismo patron que generar_dashboard_drive.py)
# ---------------------------------------------------------------------------
def fetch_sheet_csv(sheet_id, sheet_name):
    url = ('https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&sheet=%s'
           % (sheet_id, urllib.parse.quote(sheet_name)))
    for intento in range(4):
        try:
            with urllib.request.urlopen(url, timeout=20) as resp:
                text = resp.read().decode('utf-8')
            return list(csv.reader(io.StringIO(text)))
        except Exception as e:
            if intento < 3:
                print(f"  Error de red leyendo '{sheet_name}', reintentando... ({e})")
                time.sleep(2)
            else:
                raise RuntimeError(f"No se pudo descargar la hoja '{sheet_name}': {e}")

def s(val):
    return str(val).strip() if val is not None else ""

def n(val):
    """Numero formato AR (punto=miles, coma=decimal) -> float."""
    txt = s(val).replace('%', '')
    if not txt:
        return 0.0
    neg = txt.startswith('-')
    if neg:
        txt = txt[1:]
    txt = txt.replace('.', '').replace(',', '.')
    try:
        f = float(txt)
    except ValueError:
        return 0.0
    return -f if neg else f

# ---------------------------------------------------------------------------
# Deteccion de categoria (igual a generar_sitio.py)
# ---------------------------------------------------------------------------
def detectar_categoria(desc, cod, marca_brand=''):
    d = desc.upper()
    m = marca_brand.upper()

    if cod.startswith('CE'):
        if 'HIDRO' in d:                              return 'Hidrolavadoras'
        if 'LAVARROPAS' in d or 'WHIRLPOOL' in d:    return 'Lavarropas'
        if 'COSER' in d or 'SINGER' in d or 'NECCHI' in d: return 'Maq. de Coser'
        if 'ASPIRADORA' in d or 'TURBION' in d or 'BARRE' in d or 'LUSTRA' in d: return 'Aspiradoras'
        if 'SIERRA' in d or 'LAZO' in d:             return 'Sierras'
        return 'Electrodomesticos'

    if 'HIDROLAVADORA' in m or 'HIDROLAV' in m or 'HIDROLAVADORA' in d: return 'Hidrolavadoras'
    if 'LAVARROPAS' in m or 'LAVARROPAS' in d:       return 'Lavarropas'
    if 'LAVADORA' in d and 'HIDRO' not in d:          return 'Lavarropas'
    if 'ASPIRADORA' in m or 'ASPIRADORAS' in m or 'ASPIRADORA' in d: return 'Aspiradoras'
    if 'BARRE ALFOMBRA' in m or 'BARRE ALFOMBRA' in d: return 'Aspiradoras'
    if 'AMOLADORA' in d or 'AMOL.' in d or 'ESMERIL' in d: return 'Amoladoras'
    if 'AMOLADORA' in m or 'AMOL.' in m:             return 'Amoladoras'
    if 'TALADRO' in d or 'TAL.' in d or 'TAL ' in d: return 'Taladros'
    if 'LIJADORA' in d or 'LIJADORAS' in d or 'LIJA' in d or 'PULIDORA' in d or 'CEPILLO' in d: return 'Lijadoras'
    if 'SIERRA' in d or 'CALADORA' in d or 'TRONZADORA' in d: return 'Sierras'
    return 'Otros'

# ---------------------------------------------------------------------------
# Lectores Google Sheets
# ---------------------------------------------------------------------------
def leer_stock_moncar():
    """COD -> {stock: bool, marca: str}. Layout confirmado en vivo:
    fila0=TOTALES, fila1=header, desde fila2: col0=COD,col1=MARCA,col2=PRECIO(cache),
    col3:7=NO_EMB/EMB_DEP/PREM_DEP/CABA (stock = suma > 0)."""
    print("Leyendo STOCK MONCAR desde Google Sheets...")
    result = {}
    for row in fetch_sheet_csv(STOCK_MC_SHEET_ID, 'Productos')[2:]:
        if not row or not s(row[0]) or not s(row[0]).startswith(('CC', 'CF', 'CE')):
            continue
        cod   = s(row[0])
        marca = s(row[1]).title() if len(row) > 1 and s(row[1]) else 'Moncar'
        vals  = [n(v) for v in row[3:7]] if len(row) > 3 else []
        result[cod] = {'stock': sum(vals) > 0, 'marca': marca}
    print(f"  -> {len(result)} productos MONCAR (stock)")
    return result

def leer_stock_rmh():
    """COD -> {stock: bool}. fila0=TOTALES, fila1=header, desde fila2:
    col0=COD,col1=DESC,col2=PRECIO(cache),col3:7=stock (suma > 0)."""
    print("Leyendo STOCK RMH desde Google Sheets...")
    result = {}
    for row in fetch_sheet_csv(STOCK_RMH_SHEET_ID, 'Productos')[2:]:
        if not row or not s(row[0]) or not s(row[0]).startswith('RM'):
            continue
        cod  = s(row[0])
        vals = [n(v) for v in row[3:7]] if len(row) > 3 else []
        result[cod] = {'stock': sum(vals) > 0}
    print(f"  -> {len(result)} productos RMH (stock)")
    return result

def leer_imagenes():
    """COD -> nombre de archivo en sitio/imagenes/ (foto real del producto,
    extraida del PDF 'CATALOGO CARBONES MONCAR' con extraer_imagenes_catalogo.py).
    Si un producto no tiene foto todavia, el sitio sigue usando el icono SVG
    generico como fallback (ver renderGrid en el JS)."""
    mapa = {}
    if os.path.isdir(IMG_DIR):
        for fname in os.listdir(IMG_DIR):
            cod = os.path.splitext(fname)[0].replace('_', ' ')
            mapa[cod] = fname
    print(f"  -> {len(mapa)} productos con foto real (sitio/imagenes/)")
    return mapa

def leer_dimensiones_moncar():
    """esp/anch/larg/terminal por COD, del Excel local (no migrado a Sheets,
    ver pendiente en CLAUDE.md). Si el archivo no esta o cambio de formato,
    se sigue sin dimensiones en vez de romper la generacion del sitio."""
    dims = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(LP_DIM_PATH, data_only=True)

        def val(v):
            try:
                return float(str(v).strip())
            except Exception:
                return None

        ws = wb['Precios de Carbones']
        for row in ws.iter_rows(min_row=8, values_only=True):
            cod = row[0]
            if not cod or not str(cod).strip().startswith(('CF', 'CC')):
                continue
            cod  = str(cod).strip()
            esp, anch, larg = val(row[2]), val(row[3]), val(row[4])
            term = int(row[1]) if isinstance(row[1], (int, float)) else None
            d = 'x'.join(str(v) for v in [row[2], row[3], row[4]]
                         if v is not None and str(v).strip() not in ('', 'None', 'DIAM'))
            dims[cod] = {'esp': esp, 'anch': anch, 'larg': larg, 'terminal': term, 'dims': d}

        ws2 = wb['Carbones Electrodomesticos']
        for row in ws2.iter_rows(min_row=2, values_only=True):
            cod = row[0]
            if not cod or not str(cod).strip().startswith('CE'):
                continue
            cod  = str(cod).strip()
            esp, anch, larg = val(row[3]), val(row[4]), val(row[5])
            d = 'x'.join(str(v) for v in [row[3], row[4], row[5]]
                         if v is not None and str(v).strip() not in ('', 'None'))
            dims[cod] = {'esp': esp, 'anch': anch, 'larg': larg, 'terminal': None, 'dims': d}

        print(f"  -> {len(dims)} productos con dimensiones (LP local)")
    except Exception as e:
        print(f"  AVISO: no se pudieron leer dimensiones de '{LP_DIM_PATH}': {e}")
    return dims

def leer_moncar():
    """Productos MONCAR desde LP MONCAR (precios/marca/desc siempre actualizados)
    + stock desde STOCK_MONCAR + dimensiones desde el Excel local. Layout de LP
    MONCAR confirmado en vivo: fila0=header, desde fila1: col0=COD,col1=LINEA,
    col2=MARCA,col3=DESCRIPCION,col4=PRECIO vigente."""
    print("Leyendo LP MONCAR desde Google Sheets...")
    stock_map = leer_stock_moncar()
    dims_map  = leer_dimensiones_moncar()
    img_map   = leer_imagenes()

    productos = []
    for row in fetch_sheet_csv(VENTAS_SHEET_ID, 'LP MONCAR')[1:]:
        if len(row) < 5:
            continue
        cod = s(row[0])
        if not cod or not cod.startswith(('CF', 'CC', 'CE')):
            continue
        marca_lp = s(row[2]).title() if s(row[2]) else 'Moncar'
        desc     = s(row[3])
        precio   = n(row[4])

        info  = stock_map.get(cod, {})
        marca = info.get('marca', marca_lp) if info else marca_lp
        cat   = detectar_categoria(desc, cod, marca)
        d     = dims_map.get(cod, {})

        productos.append({
            'cod':      cod,
            'linea':    'MONCAR',
            'marca':    marca,
            'cat':      cat,
            'desc':     desc,
            'dims':     d.get('dims', ''),
            'esp':      d.get('esp'),
            'anch':     d.get('anch'),
            'larg':     d.get('larg'),
            'terminal': d.get('terminal'),
            'precio':   int(round(precio)),
            'stock':    info.get('stock', False),
            'img':      img_map.get(cod, ''),
        })
    print(f"  -> {len(productos)} productos MONCAR")
    return productos

def leer_rmh():
    """Productos RMH desde LP RMH + stock desde STOCK_RMH. Layout de LP RMH
    confirmado en vivo: fila0=header, desde fila1: col0=COD,col1=LINEA,
    col2=MARCA,col3=DESCRIPCION,col4=precio anterior,col5=PRECIO vigente."""
    print("Leyendo LP RMH desde Google Sheets...")
    stock_map = leer_stock_rmh()
    img_map   = leer_imagenes()

    productos = []
    for row in fetch_sheet_csv(VENTAS_SHEET_ID, 'LP RMH')[1:]:
        if len(row) < 6:
            continue
        cod = s(row[0])
        if not cod or not cod.startswith('RM'):
            continue
        desc   = s(row[3])
        precio = n(row[5])
        info   = stock_map.get(cod, {})
        cat    = detectar_categoria(desc, cod)

        productos.append({
            'cod':      cod,
            'linea':    'RMH',
            'marca':    'RMH',
            'cat':      cat,
            'desc':     desc,
            'dims':     '',
            'esp':      None,
            'anch':     None,
            'larg':     None,
            'terminal': None,
            'precio':   int(round(precio)),
            'stock':    info.get('stock', False),
            'img':      img_map.get(cod, ''),
        })
    print(f"  -> {len(productos)} productos RMH")
    return productos

# ---------------------------------------------------------------------------
# Generador HTML (igual a generar_sitio.py, salvo el bloque de ruteo por zona
# — ver nota en confirmarPedido: la version vieja mandaba el pedido a
# "hermanos" con telefonos ficticios que no correspondian a nadie del equipo
# real; ahora siempre va al WhatsApp/email real de la empresa)
# ---------------------------------------------------------------------------
def generar_html(productos):
    cat_count = Counter(p['cat'] for p in productos)
    cats_orden = [c for c, _ in sorted(cat_count.items(), key=lambda x: -x[1])]

    marca_count = Counter(
        p['marca'] for p in productos
        if p['linea'] == 'MONCAR' and p['marca'] not in EXCLUIR_MARCAS
    )
    marcas_destacadas = [(m, cnt) for m, cnt in marca_count.items() if m in MARCAS_CONFIG]
    marcas_destacadas.sort(key=lambda x: -x[1])

    esps  = sorted(set(p['esp']  for p in productos if p['esp']  is not None))
    anchs = sorted(set(p['anch'] for p in productos if p['anch'] is not None))
    largs = sorted(set(p['larg'] for p in productos if p['larg'] is not None))
    terms = sorted(set(p['terminal'] for p in productos if p['terminal'] is not None))

    def cat_cards():
        cards = []
        for cat in cats_orden:
            cfg   = CAT_CONFIG.get(cat, CAT_CONFIG['Otros'])
            count = cat_count[cat]
            cards.append(
                f'<div class="cat-card" onclick="irACat(\'{cat}\')">'
                f'<div class="cat-icon">{cfg["icon"]}</div>'
                f'<div class="cat-name">{cat}</div>'
                f'<div class="cat-count">{count} productos</div>'
                f'</div>'
            )
        return '\n'.join(cards)

    def marca_cards():
        cards = []
        for m, cnt in marcas_destacadas:
            cards.append(
                f'<div class="marca-card" onclick="irAMarca(\'{m}\')">'
                f'<div class="marca-name">{m}</div>'
                f'<div class="marca-sub">{cnt} carbones</div>'
                f'</div>'
            )
        return '\n'.join(cards)

    def sidebar_marcas():
        opts = []
        for m, cnt in sorted(marca_count.items(), key=lambda x: -x[1]):
            opts.append(
                f'<label class="sb-check">'
                f'<input type="checkbox" value="{m}" onchange="filtrar()">'
                f'<span>{m}</span><span class="sb-cnt">{cnt}</span>'
                f'</label>'
            )
        return '\n'.join(opts)

    def sidebar_cats():
        opts = []
        for cat in cats_orden:
            opts.append(
                f'<label class="sb-check">'
                f'<input type="checkbox" value="{cat}" onchange="filtrar()">'
                f'<span>{cat}</span><span class="sb-cnt">{cat_count[cat]}</span>'
                f'</label>'
        )
        return '\n'.join(opts)

    def opt(vals):
        return '\n'.join(f'<option value="{v}">{v}</option>' for v in vals)

    productos_json = json.dumps(productos, ensure_ascii=False)

    return f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MONCAR RMH — Repuestos para Herramientas</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#f0f0f0;--white:#fff;--surface:#fff;--border:#e0e0e0;
  --black:#111111;--black2:#000000;--grey-dark:#3a3a3a;
  --orange:#FF5A00;--orange2:#e04e00;
  --text:#111;--muted:#666;
  --green:#16a34a;--red:#dc2626;
  --nav-h:104px;--sb-w:260px;
  --font-head:'Oswald',sans-serif;--font-body:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
}}
html{{scroll-behavior:smooth}}
body{{background:var(--bg);color:var(--text);font-family:var(--font-body)}}

/* NAV */
nav{{position:fixed;top:0;left:0;right:0;z-index:100;background:var(--black2);height:var(--nav-h);display:flex;align-items:center;padding:0 28px;box-shadow:0 1px 4px rgba(0,0,0,.3)}}
.nav-logo{{display:flex;align-items:center;cursor:pointer;margin-right:36px}}
.nav-wordmark{{font-family:var(--font-head);font-size:28px;font-weight:700;letter-spacing:.5px;color:#fff;text-transform:uppercase}}
.nav-wordmark span{{color:var(--orange)}}
.nav-links{{display:flex;gap:2px}}
.nav-link{{font-family:var(--font-head);color:rgba(255,255,255,.75);font-size:17px;font-weight:600;letter-spacing:.3px;text-transform:uppercase;cursor:pointer;padding:10px 16px;border-radius:6px;border:none;background:none;transition:.15s}}
.nav-link:hover{{color:#fff;background:rgba(255,255,255,.08)}}
.nav-link.active{{color:var(--orange);font-weight:700}}
.nav-right{{display:flex;gap:10px;align-items:center;margin-left:auto}}
.btn-wa-nav{{font-family:var(--font-head);background:#25D366;color:#fff;border:none;padding:11px 18px;border-radius:7px;font-size:14px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;cursor:pointer;text-decoration:none;display:flex;align-items:center;gap:6px}}
.btn-cart-nav{{font-family:var(--font-head);background:var(--orange);color:#fff;border:none;padding:11px 18px;border-radius:7px;font-size:14px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;cursor:pointer;display:flex;align-items:center;gap:6px}}
.btn-cart-nav:hover{{background:var(--orange2)}}
#cart-count{{background:#fff;color:var(--orange);border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:900}}

/* PAGES */
.page{{display:none;padding-top:var(--nav-h)}}
.page.active{{display:block}}

/* ===== INICIO ===== */
.hero{{background:#0a0a0a;padding:90px 40px;position:relative;overflow:hidden}}
.hero::before{{content:'';position:absolute;inset:0;background:
    radial-gradient(ellipse 800px 550px at 0% 0%,rgba(255,90,0,.28) 0%,transparent 60%),
    radial-gradient(ellipse 1000px 750px at 100% 100%,rgba(255,90,0,.32) 0%,transparent 60%)}}
.hero-inner{{position:relative;z-index:1;max-width:1100px;margin:0 auto;display:flex;align-items:center;justify-content:center;gap:56px;text-align:left}}
.hero-logo{{flex-shrink:0}}
.hero-logo-frame{{display:inline-block;border:2px solid rgba(255,90,0,.45);border-radius:18px;padding:20px;background:rgba(0,0,0,.3)}}
.hero-logo img{{display:block;width:clamp(200px,22vw,320px);height:clamp(200px,22vw,320px);border-radius:8px;object-fit:cover}}
.hero-text{{flex:0 1 auto;max-width:640px;min-width:0}}
.hero-badge{{display:inline-block;background:rgba(255,107,0,.15);color:#FF6B00;border:1px solid rgba(255,107,0,.35);border-radius:20px;padding:6px 18px;font-size:13px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:22px}}
.hero h1{{font-family:var(--font-head);font-size:clamp(36px,4.2vw,64px);font-weight:700;text-transform:uppercase;line-height:1.08;color:#fff;margin-bottom:20px}}
.hero h1 em{{color:var(--orange);font-style:normal}}
.hero p{{color:rgba(255,255,255,.7);font-size:19px;margin-bottom:34px;line-height:1.6;max-width:640px}}
.hero-btns{{display:flex;gap:16px;justify-content:flex-start;flex-wrap:wrap}}
.btn-hero-p{{font-family:var(--font-head);background:var(--orange);color:#fff;border:none;padding:17px 36px;border-radius:8px;font-size:16px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;cursor:pointer}}
.btn-hero-p:hover{{background:var(--orange2)}}
.btn-hero-s{{font-family:var(--font-head);color:#fff;border:1px solid rgba(255,255,255,.3);padding:17px 36px;border-radius:8px;font-size:16px;font-weight:600;text-transform:uppercase;letter-spacing:.3px;cursor:pointer;text-decoration:none;display:inline-block;background:transparent}}
.btn-hero-s:hover{{border-color:rgba(255,255,255,.6)}}

.sec{{padding:52px 28px}}
.sec-alt{{background:var(--white)}}
.sec-head{{margin-bottom:28px}}
.sec-head h2{{font-family:var(--font-head);font-size:clamp(22px,3vw,28px);font-weight:700;text-transform:uppercase;letter-spacing:.3px}}
.sec-head p{{color:var(--muted);font-size:14px;margin-top:5px}}

/* MARCA CARDS — grilla fija de 6 columnas (12 marcas destacadas = 2 filas
   exactas). Si el numero de marcas cambia en el futuro y deja de ser
   multiplo de 6, hay que revisar esto a mano. */
.marca-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:14px}}
.marca-card{{background:var(--white);border:1px solid var(--border);border-bottom:3px solid var(--border);border-radius:8px;padding:22px 16px;cursor:pointer;transition:transform .15s,box-shadow .15s,border-color .15s;text-align:center}}
.marca-card:hover{{transform:translateY(-3px);box-shadow:0 8px 20px rgba(0,0,0,.12);border-bottom-color:var(--orange)}}
.marca-name{{font-family:var(--font-head);font-size:16px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;color:var(--black)}}
.marca-sub{{font-size:12px;margin-top:4px;color:var(--muted)}}

/* CAT CARDS — grilla fija de 5 columnas (10 categorias = 2 filas exactas) */
.cat-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px}}
.cat-card{{background:var(--white);border:1px solid var(--border);border-radius:8px;padding:22px 16px;cursor:pointer;transition:transform .15s,box-shadow .15s,border-color .15s;text-align:center}}
.cat-card:hover{{transform:translateY(-3px);box-shadow:0 8px 20px rgba(0,0,0,.12);border-color:var(--orange)}}
.cat-icon{{font-size:30px;color:var(--orange);margin-bottom:10px}}
.cat-icon svg{{width:1em;height:1em;display:inline-block;vertical-align:middle}}
.cat-name{{font-family:var(--font-head);color:var(--black);font-size:14px;font-weight:700;text-transform:uppercase;letter-spacing:.5px}}
.cat-count{{color:var(--muted);font-size:12px;margin-top:3px}}

/* NOSOTROS */
.nosotros-wrap{{max-width:900px;margin:0 auto}}
.nos-hero{{background:linear-gradient(135deg,#111,#000);border-radius:16px;padding:48px 40px;margin-bottom:36px;text-align:center}}
.nos-hero h2{{font-family:var(--font-head);font-size:30px;font-weight:700;text-transform:uppercase;color:#fff;margin-bottom:12px}}
.nos-hero h2 span{{color:var(--orange)}}
.nos-hero p{{color:rgba(255,255,255,.65);font-size:16px;line-height:1.7;max-width:560px;margin:0 auto}}
.nos-cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:28px}}
.n-card{{background:var(--white);border:1px solid var(--border);border-radius:8px;padding:22px}}
.n-icon{{font-size:28px;color:var(--orange);margin-bottom:10px}}
.n-card h3{{font-family:var(--font-head);font-size:14px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;margin-bottom:6px}}
.n-card p{{font-size:13px;color:var(--muted);line-height:1.6}}
.contact-box{{background:var(--white);border:1px solid var(--border);border-radius:8px;padding:24px}}
.contact-box h3{{font-family:var(--font-head);font-size:15px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;margin-bottom:16px}}
.c-row{{display:flex;align-items:center;gap:14px;padding:11px 0;border-bottom:1px solid var(--border)}}
.c-row:last-child{{border-bottom:none}}
.c-ico{{width:38px;height:38px;border-radius:9px;background:var(--bg);color:var(--orange);display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0}}
.c-lbl{{font-size:11px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.5px}}
.c-val{{font-size:14px;font-weight:700}}
.c-val a{{color:var(--text);text-decoration:none}}
.c-val a:hover{{color:var(--orange)}}

/* ===== PRODUCTOS - layout con sidebar ===== */
.prod-layout{{display:flex;align-items:flex-start;min-height:calc(100vh - var(--nav-h))}}

/* SIDEBAR */
.sidebar{{width:var(--sb-w);flex-shrink:0;background:var(--white);border-right:1px solid var(--border);min-height:calc(100vh - var(--nav-h));position:sticky;top:var(--nav-h);max-height:calc(100vh - var(--nav-h));overflow-y:auto;padding:20px 16px}}
.sb-title{{font-family:var(--font-head);font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--black);margin-bottom:12px;margin-top:4px}}
.sb-section{{margin-bottom:20px;border-bottom:1px solid var(--border);padding-bottom:18px}}
.sb-section:last-child{{border-bottom:none}}
.sb-check{{display:flex;align-items:center;gap:7px;font-size:13px;cursor:pointer;padding:3px 0;color:var(--text)}}
.sb-check input{{accent-color:var(--orange);width:14px;height:14px;cursor:pointer}}
.sb-check span:first-of-type{{flex:1}}
.sb-cnt{{color:var(--muted);font-size:11px}}
.sb-linea{{display:flex;gap:6px;flex-wrap:wrap}}
.sb-linea-btn{{font-family:var(--font-head);border:1px solid var(--border);background:var(--white);color:var(--muted);border-radius:6px;padding:6px 12px;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.2px;cursor:pointer;transition:.15s}}
.sb-linea-btn:hover{{background:var(--bg)}}
.sb-linea-btn.active{{background:var(--orange);border-color:var(--orange);color:#fff}}
.sb-range{{display:flex;gap:8px;align-items:center}}
.sb-range input{{width:80px;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:5px 8px;font-size:13px;color:var(--text)}}
.sb-range span{{color:var(--muted);font-size:13px}}
.sb-select{{width:100%;background:var(--bg);border:1px solid var(--border);border-radius:6px;padding:6px 8px;font-size:13px;color:var(--text);margin-bottom:6px}}
.sb-clear{{font-family:var(--font-head);width:100%;background:none;border:1px solid var(--border);border-radius:7px;padding:8px;font-size:13px;font-weight:600;text-transform:uppercase;letter-spacing:.3px;color:var(--muted);cursor:pointer;margin-top:4px}}
.sb-clear:hover{{border-color:var(--orange);color:var(--orange)}}

/* MAIN CONTENT */
.prod-main{{flex:1;min-width:0}}
.prod-topbar{{padding:16px 20px;background:var(--white);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:12px}}
.search-wrap{{position:relative;flex:1}}
.search-wrap input{{width:100%;background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:9px 36px 9px 12px;color:var(--text);font-size:14px}}
.search-wrap input:focus{{outline:none;border-color:var(--orange)}}
.search-wrap input::placeholder{{color:var(--muted)}}
.search-clear{{position:absolute;right:10px;top:50%;transform:translateY(-50%);background:none;border:none;color:var(--muted);cursor:pointer;font-size:15px;display:none}}
.active-filters{{padding:8px 20px;background:var(--bg);display:none;gap:8px;flex-wrap:wrap}}
.active-tag{{background:rgba(255,107,0,.1);border:1px solid rgba(255,107,0,.25);color:var(--orange);border-radius:20px;padding:3px 12px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:5px;cursor:pointer}}
.grid-wrap{{padding:20px}}
.results-info{{font-size:13px;color:var(--muted);margin-bottom:14px}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(195px,1fr));gap:13px}}
.no-results{{text-align:center;padding:60px 20px;color:var(--muted)}}

/* PRODUCT CARD */
.card{{background:var(--white);border:1px solid var(--border);border-radius:10px;overflow:hidden;display:flex;flex-direction:column;transition:border-color .15s,box-shadow .15s}}
.card:hover{{border-color:var(--orange);box-shadow:0 4px 16px rgba(255,107,0,.1)}}
.card-icon{{background:var(--bg);display:flex;align-items:center;justify-content:center;padding:10px;height:120px}}
.card-icon img{{max-width:100%;max-height:100%;object-fit:contain}}
.card-body{{padding:11px;flex:1;display:flex;flex-direction:column;gap:4px}}
.card-badges{{display:flex;gap:4px;flex-wrap:wrap}}
.badge{{display:inline-block;font-size:10px;border-radius:4px;padding:2px 6px;font-weight:700}}
.bl{{background:rgba(255,107,0,.1);color:var(--orange)}}
.bc{{background:#f0f0f0;color:var(--muted)}}
.card-cod{{font-family:var(--font-head);font-size:11px;font-weight:700;color:var(--orange);letter-spacing:.5px;margin-top:2px}}
.card-desc{{font-size:13px;line-height:1.35;flex:1}}
.card-sub{{font-size:11px;color:var(--muted)}}
.card-precio{{font-family:var(--font-head);font-size:17px;font-weight:700;margin-top:4px;color:var(--black)}}
.card-qty{{display:flex;align-items:center;gap:6px;margin-top:6px}}
.card-qty button{{background:var(--bg);border:1px solid var(--border);width:26px;height:26px;border-radius:5px;cursor:pointer;font-size:15px;font-weight:700;color:var(--text);line-height:1}}
.card-qty button:hover{{border-color:var(--orange);color:var(--orange)}}
.card-qty input{{width:40px;text-align:center;border:1px solid var(--border);border-radius:5px;padding:4px 2px;font-size:13px;font-family:var(--font-body);color:var(--text)}}
.btn-add{{font-family:var(--font-head);margin-top:7px;background:var(--orange);color:#fff;border:none;border-radius:6px;padding:9px;font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;cursor:pointer;width:100%}}
.btn-add:hover{{background:var(--orange2)}}
.btn-cargar-mas{{font-family:var(--font-head);background:var(--white);color:var(--black);border:1px solid var(--border);border-radius:6px;padding:11px 34px;font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;cursor:pointer}}
.btn-cargar-mas:hover{{border-color:var(--orange);color:var(--orange)}}

/* CART */
.cart-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:200}}
.cart-panel{{position:fixed;top:0;right:-420px;width:min(420px,100vw);height:100%;background:var(--white);border-left:1px solid var(--border);z-index:201;display:flex;flex-direction:column;transition:right .25s;box-shadow:-4px 0 20px rgba(0,0,0,.1)}}
.cart-panel.open{{right:0}}
.cart-header{{padding:18px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between}}
.cart-header h2{{font-size:16px;font-weight:800}}
.btn-xc{{background:none;border:none;color:var(--muted);font-size:22px;cursor:pointer;line-height:1}}
.cart-items{{flex:1;overflow-y:auto;padding:12px}}
.cart-empty{{text-align:center;padding:40px 20px;color:var(--muted)}}
.ci{{display:flex;gap:10px;align-items:flex-start;padding:10px 0;border-bottom:1px solid var(--border)}}
.ci-info{{flex:1}}
.ci-cod{{font-size:11px;color:var(--orange);font-weight:700}}
.ci-desc{{font-size:13px;margin:2px 0}}
.ci-precio{{font-size:14px;font-weight:700}}
.ci-qty{{display:flex;align-items:center;gap:6px;margin-top:5px}}
.ci-qty button{{background:#f0f0f0;border:none;width:24px;height:24px;border-radius:4px;cursor:pointer;font-size:15px}}
.ci-qty span{{font-size:14px;font-weight:700;min-width:20px;text-align:center}}
.btn-rm{{background:none;border:none;color:var(--red);cursor:pointer;font-size:17px;padding:0 4px;line-height:1}}
.cart-footer{{padding:16px 20px;border-top:1px solid var(--border)}}
.cart-total{{display:flex;justify-content:space-between;font-size:16px;font-weight:800;margin-bottom:14px}}
.btn-wa-c{{background:#25D366;color:#fff;border:none;border-radius:8px;padding:12px;font-size:15px;font-weight:700;cursor:pointer;width:100%;margin-bottom:8px}}
.btn-em-c{{background:var(--white);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:12px;font-size:14px;font-weight:600;cursor:pointer;width:100%}}
.btn-em-c:hover{{border-color:var(--orange);color:var(--orange)}}

/* FOOTER */
footer{{background:var(--black2);border-top:3px solid var(--orange);padding:36px 28px 24px}}
.fi{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:24px;max-width:1000px;margin:0 auto;align-items:start}}
.f-logo{{display:flex;align-items:center;gap:12px;margin-bottom:8px}}
.f-logo img{{height:52px;width:52px;border-radius:6px;object-fit:cover}}
.fb{{font-family:var(--font-head);font-size:18px;font-weight:700;text-transform:uppercase;color:#fff}}
.fb span{{color:var(--orange)}}
.fd{{font-size:13px;color:rgba(255,255,255,.6);line-height:1.5}}
.fc h4{{font-family:var(--font-head);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:rgba(255,255,255,.5);margin-bottom:10px}}
.fc a{{display:block;font-size:13px;color:rgba(255,255,255,.85);text-decoration:none;margin-bottom:5px;cursor:pointer}}
.fc a:hover{{color:var(--orange)}}
.ft{{text-align:center;padding-top:18px;margin-top:20px;border-top:1px solid rgba(255,255,255,.1);font-size:12px;color:rgba(255,255,255,.4)}}

/* MOBILE SIDEBAR TOGGLE */
.sb-toggle{{font-family:var(--font-head);display:none;background:var(--orange);color:#fff;border:none;padding:8px 16px;border-radius:7px;font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.3px;cursor:pointer}}
.sidebar-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:50}}

@media(max-width:900px){{
  .marca-grid{{grid-template-columns:repeat(3,1fr)}}
  .cat-grid{{grid-template-columns:repeat(3,1fr)}}
}}
@media(max-width:860px){{
  .sidebar{{position:fixed;top:var(--nav-h);left:-280px;z-index:51;transition:left .25s;max-height:100vh}}
  .sidebar.open{{left:0}}
  .sidebar-overlay.open{{display:block}}
  .sb-toggle{{display:flex;align-items:center;gap:6px}}
  .prod-main{{width:100%}}
}}
@media(max-width:760px){{
  .hero{{padding:48px 20px}}
  .hero-inner{{flex-direction:column;text-align:center;gap:28px}}
  .hero-logo-frame{{padding:12px}}
  .hero-logo img{{width:150px;height:150px}}
  .hero-badge,.hero p{{margin-left:auto;margin-right:auto}}
  .hero-btns{{justify-content:center}}
}}
@media(max-width:640px){{
  :root{{--nav-h:76px}}
  nav{{padding:0 14px;gap:6px}}
  .nav-logo{{margin-right:10px}}
  .nav-wordmark{{font-size:17px}}
  .nav-links{{gap:0}}
  .nav-link{{padding:8px 8px;font-size:12px}}
  .btn-wa-nav span{{display:none}}
  .btn-cart-nav span:first-child{{display:none}}
  .sec{{padding:36px 16px}}
  .marca-grid,.cat-grid{{grid-template-columns:repeat(2,1fr)}}
  .nos-hero{{padding:32px 20px}}
  .grid{{grid-template-columns:repeat(2,1fr);gap:10px}}
  .grid-wrap{{padding:14px}}
  .fi{{text-align:center}}
  .f-logo{{justify-content:center}}
}}

/* FORM PEDIDO MODAL */
.fmodal-ov{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:400;align-items:center;justify-content:center;padding:20px}}
.fmodal-ov.open{{display:flex}}
.fmodal{{background:var(--white);border-radius:14px;padding:28px;width:min(480px,100%);max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.25)}}
.fmodal-title{{font-size:19px;font-weight:900;margin-bottom:5px}}
.fmodal-sub{{font-size:13px;color:var(--muted);margin-bottom:22px;line-height:1.5}}
.fg{{margin-bottom:14px}}
.fg label{{display:block;font-size:13px;font-weight:700;margin-bottom:5px}}
.fg input{{width:100%;border:1px solid var(--border);border-radius:7px;padding:10px 12px;font-size:14px;color:var(--text);background:var(--bg)}}
.fg input:focus{{outline:none;border-color:var(--orange)}}
.fg input.ferr{{border-color:var(--red)}}
.fg-note{{font-size:12px;color:var(--muted);margin-top:3px}}
.fmodal-acts{{display:flex;gap:10px;margin-top:22px}}
.btn-cfm{{flex:1;background:var(--orange);color:#fff;border:none;border-radius:8px;padding:12px;font-size:15px;font-weight:700;cursor:pointer}}
.btn-cfm:hover{{background:var(--orange2)}}
.btn-cnc{{background:#f0f0f0;color:var(--muted);border:none;border-radius:8px;padding:12px 16px;font-size:14px;font-weight:600;cursor:pointer}}
</style>
</head>
<body>

<nav>
  <div class="nav-logo" onclick="irAInicio()">
    <span class="nav-wordmark">MONCAR<span> RMH</span></span>
  </div>
  <div class="nav-links">
    <button class="nav-link active" id="nav-inicio" onclick="irAInicio()">Inicio</button>
    <button class="nav-link" id="nav-productos" onclick="irAProductos()">Productos</button>
    <button class="nav-link" id="nav-nosotros" onclick="irANosotros()">Nosotros</button>
  </div>
  <div class="nav-right">
    <button class="btn-wa-nav" onclick="abrirConsultaWA()">&#128172; <span>WhatsApp</span></button>
    <button class="btn-cart-nav" onclick="abrirCart()"><span>Carrito</span> <span id="cart-count">0</span></button>
  </div>
</nav>

<!-- ===== INICIO ===== -->
<div id="page-inicio" class="page active">
  <div class="hero">
    <div class="hero-inner">
      <div class="hero-logo"><div class="hero-logo-frame"><img src="logo-hero.png" alt="MONCAR RMH"></div></div>
      <div class="hero-text">
        <div class="hero-badge">Repuestos originales</div>
        <h1>Carbones y repuestos<br>para tus <em>herramientas</em></h1>
        <p>Carbones electricos MONCAR e interruptores RMH.<br>Compatible con las principales marcas. Envios a todo el pais.</p>
        <div class="hero-btns">
          <button class="btn-hero-p" onclick="irAProductos()">Ver catalogo completo</button>
          <button class="btn-hero-s" onclick="abrirConsultaWA()">Consultar por WhatsApp</button>
        </div>
      </div>
    </div>
  </div>

  <div class="sec">
    <div class="sec-head"><h2>Compra por marca</h2><p>Compatible con las principales marcas del mercado</p></div>
    <div class="marca-grid">{marca_cards()}</div>
  </div>

  <div class="sec sec-alt">
    <div class="sec-head"><h2>Compra por categoria</h2><p>Encontra rapido el repuesto para tu herramienta</p></div>
    <div class="cat-grid">{cat_cards()}</div>
  </div>
</div>

<!-- ===== PRODUCTOS ===== -->
<div id="page-productos" class="page">
  <div class="prod-layout">

    <!-- SIDEBAR -->
    <div class="sidebar" id="sidebar">
      <div class="sb-section">
        <div class="sb-title">Linea</div>
        <div class="sb-linea">
          <button class="sb-linea-btn active" data-val="" onclick="setLinea(this)">Todos</button>
          <button class="sb-linea-btn" data-val="MONCAR" onclick="setLinea(this)">MONCAR</button>
          <button class="sb-linea-btn" data-val="RMH" onclick="setLinea(this)">RMH</button>
        </div>
      </div>

      <div class="sb-section">
        <div class="sb-title">Categoria</div>
        <div id="sb-cats">{sidebar_cats()}</div>
      </div>

      <div class="sb-section">
        <div class="sb-title">Marca</div>
        <div id="sb-marcas">{sidebar_marcas()}</div>
      </div>

      <div class="sb-section">
        <div class="sb-title">Precio</div>
        <div class="sb-range">
          <input type="number" id="precio-min" placeholder="Min" value="" oninput="filtrar()">
          <span>—</span>
          <input type="number" id="precio-max" placeholder="Max" value="" oninput="filtrar()">
        </div>
      </div>

      <div class="sb-section">
        <div class="sb-title">Dimensiones (mm)</div>
        <select class="sb-select" id="fil-esp" onchange="filtrar()">
          <option value="">Espesor: todos</option>
          {opt(esps)}
        </select>
        <select class="sb-select" id="fil-anch" onchange="filtrar()">
          <option value="">Ancho: todos</option>
          {opt(anchs)}
        </select>
        <select class="sb-select" id="fil-larg" onchange="filtrar()">
          <option value="">Largo: todos</option>
          {opt(largs)}
        </select>
      </div>

      <div class="sb-section">
        <div class="sb-title">Terminal</div>
        <select class="sb-select" id="fil-term" onchange="filtrar()">
          <option value="">Todos los terminales</option>
          {opt(terms)}
        </select>
      </div>

      <button class="sb-clear" onclick="limpiarFiltros()">Limpiar filtros</button>
    </div><!-- /sidebar -->

    <div class="sidebar-overlay" id="sb-overlay" onclick="cerrarSidebar()"></div>

    <!-- MAIN -->
    <div class="prod-main">
      <div class="prod-topbar">
        <button class="sb-toggle" onclick="abrirSidebar()">&#9776; Filtros</button>
        <div class="search-wrap">
          <input type="text" id="search" placeholder="Buscar codigo, descripcion, marca..." oninput="filtrar()">
          <button class="search-clear" id="clear-btn" onclick="limpiarBusqueda()">&#x2715;</button>
        </div>
      </div>
      <div class="active-filters" id="active-filters"></div>
      <div class="grid-wrap">
        <div class="results-info" id="results-info"></div>
        <div class="grid" id="grid"></div>
        <div class="no-results" id="no-results" style="display:none">
          <p style="font-size:18px;margin-bottom:8px">Sin resultados</p>
          <span>Proba ajustando los filtros</span>
        </div>
        <div id="cargar-mas-wrap" style="display:none;text-align:center;margin-top:24px">
          <div id="cargar-mas-info" style="font-size:12px;color:var(--muted);margin-bottom:10px"></div>
          <button class="btn-cargar-mas" onclick="cargarMas()">Cargar mas</button>
        </div>
      </div>
    </div>

  </div>
</div>

<!-- ===== NOSOTROS ===== -->
<div id="page-nosotros" class="page">
  <div class="sec">
    <div class="nosotros-wrap">
      <div class="nos-hero">
        <h2>Sobre <span>MONCAR RMH</span></h2>
        <p>Empresa familiar especializada en repuestos electricos para herramientas. Carbones electricos MONCAR e interruptores RMH compatibles con las principales marcas del mercado.</p>
      </div>
      <div class="nos-cards">
        <div class="n-card"><div class="n-icon"><i class="fa-solid fa-bolt"></i></div><h3>Experiencia en el rubro</h3><p>Anos de trayectoria en la venta de repuestos electricos para herramientas profesionales y de uso domestico.</p></div>
        <div class="n-card"><div class="n-icon"><i class="fa-solid fa-tags"></i></div><h3>Multimarca</h3><p>Compatible con Bosch, DeWalt, Makita, Black & Decker, Metabo, Hitachi, Stanley, Gamma, Milwaukee y muchas mas.</p></div>
        <div class="n-card"><div class="n-icon"><i class="fa-solid fa-truck-fast"></i></div><h3>Envios a todo el pais</h3><p>Despachamos a cualquier punto del pais. Consulta por tu zona y los tiempos de entrega.</p></div>
        <div class="n-card"><div class="n-icon"><i class="fa-solid fa-headset"></i></div><h3>Atencion personalizada</h3><p>Te ayudamos a encontrar el repuesto exacto para tu herramienta. Consulta sin compromiso por WhatsApp.</p></div>
      </div>
      <div class="contact-box">
        <h3>Contacto</h3>
        <div class="c-row">
          <div class="c-ico"><i class="fa-brands fa-whatsapp"></i></div>
          <div><div class="c-lbl">WhatsApp</div><div class="c-val"><a href="https://wa.me/5491136267009" target="_blank">+54 9 11 3626-7009</a></div></div>
        </div>
        <div class="c-row">
          <div class="c-ico"><i class="fa-solid fa-envelope"></i></div>
          <div><div class="c-lbl">Email</div><div class="c-val"><a href="mailto:contacto.carbonkey@gmail.com">contacto.carbonkey@gmail.com</a></div></div>
        </div>
        <div class="c-row">
          <div class="c-ico"><i class="fa-solid fa-boxes-stacked"></i></div>
          <div><div class="c-lbl">Catalogo</div><div class="c-val"><a onclick="irAProductos()" style="cursor:pointer;color:var(--orange)">{len(productos)} productos disponibles</a></div></div>
        </div>
      </div>
    </div>
  </div>
</div>

<footer>
  <div class="fi">
    <div>
      <div class="f-logo"><img src="logo-footer.png" alt="MONCAR RMH"><div class="fb">MONCAR<span> RMH</span></div></div>
      <div class="fd">Carbones electricos e interruptores para herramientas. Envios a todo el pais.</div>
    </div>
    <div class="fc"><h4>Navegacion</h4><a onclick="irAInicio()">Inicio</a><a onclick="irAProductos()">Productos</a><a onclick="irANosotros()">Nosotros</a></div>
    <div class="fc"><h4>Contacto</h4><a href="https://wa.me/5491136267009" target="_blank">WhatsApp: +54 9 11 3626-7009</a><a href="mailto:contacto.carbonkey@gmail.com">contacto.carbonkey@gmail.com</a></div>
  </div>
  <div class="ft">© 2026 MONCAR RMH — Todos los derechos reservados</div>
</footer>

<!-- FORM PEDIDO -->
<div class="fmodal-ov" id="fmodal-ov">
  <div class="fmodal">
    <div class="fmodal-title">Antes de escribirnos...</div>
    <div class="fmodal-sub">Nos manejamos por zonas — con estos datos te derivamos directo con quien te puede atender.</div>
    <div class="fg">
      <label>Nombre *</label>
      <input type="text" id="f-nombre" placeholder="Tu nombre completo">
    </div>
    <div class="fg">
      <label>Nombre del negocio</label>
      <input type="text" id="f-negocio" placeholder="Ferreteria, taller, bulonera, local... (opcional)">
      <div class="fg-note">Opcional si sos particular</div>
    </div>
    <div class="fg">
      <label>Localidad / Ciudad *</label>
      <input type="text" id="f-localidad" placeholder="Ej: Mendoza, Mar del Plata...">
    </div>
    <div class="fmodal-acts">
      <button class="btn-cnc" onclick="cerrarFormModal()">Cancelar</button>
      <button class="btn-cfm" onclick="confirmarPedido()">Enviar pedido &#8594;</button>
    </div>
  </div>
</div>

<!-- CART -->
<div class="cart-overlay" id="overlay" onclick="cerrarCart()"></div>
<div class="cart-panel" id="cart-panel">
  <div class="cart-header"><h2>Tu pedido</h2><button class="btn-xc" onclick="cerrarCart()">&#x2715;</button></div>
  <div class="cart-items" id="cart-items"></div>
  <div class="cart-footer">
    <div class="cart-total"><span>Total estimado</span><span id="cart-total">$0</span></div>
    <button class="btn-wa-c" onclick="pedirPorWA()">&#128172; Pedir por WhatsApp</button>
    <button class="btn-em-c" onclick="pedirPorEmail()">Pedir por Email</button>
  </div>
</div>

<script>
const PRODUCTOS = {productos_json};
const CONTACTO_EMAIL = 'contacto.carbonkey@gmail.com';
// Equipo real (sacado de la lista de precios RMH). El negocio se maneja por
// zonas, cada uno tiene su zona — pero todavia no tenemos cargado el mapa de
// "que zona le corresponde a quien", asi que por ahora TODO pedido cae en
// Franco (default) hasta que se defina el resto. Cuando este el mapa, cambiar
// elegirDestino() para que devuelva el contacto correcto segun localidad.
const CONTACTOS = {{
  FRANCO: '5491136267009',
  ANDI:   '5491139219269',
  DAVID:  '5491150225627',
  MAXI:   '5491144712072',
}};
function elegirDestino(localidad){{
  return CONTACTOS.FRANCO;
}}

const SVG_MC=`<svg width="50" height="50" viewBox="0 0 50 50" fill="none"><circle cx="25" cy="25" r="17" fill="none" stroke="#FF6B00" stroke-width="1.8" opacity=".2"/><circle cx="25" cy="25" r="9" fill="none" stroke="#FF6B00" stroke-width="1.8" opacity=".45"/><circle cx="25" cy="25" r="3" fill="#FF6B00" opacity=".8"/><line x1="9" y1="25" x2="16" y2="25" stroke="#FF6B00" stroke-width="1.8" stroke-linecap="round" opacity=".4"/><line x1="34" y1="25" x2="41" y2="25" stroke="#FF6B00" stroke-width="1.8" stroke-linecap="round" opacity=".4"/><line x1="25" y1="9" x2="25" y2="16" stroke="#FF6B00" stroke-width="1.8" stroke-linecap="round" opacity=".4"/><line x1="25" y1="34" x2="25" y2="41" stroke="#FF6B00" stroke-width="1.8" stroke-linecap="round" opacity=".4"/></svg>`;
const SVG_RMH=`<svg width="50" height="50" viewBox="0 0 50 50" fill="none"><rect x="9" y="15" width="32" height="20" rx="4" stroke="#FF6B00" stroke-width="1.8" opacity=".3" fill="none"/><rect x="15" y="20" width="20" height="10" rx="2" fill="#FF6B00" opacity=".1"/><circle cx="25" cy="25" r="3" fill="#FF6B00" opacity=".7"/><line x1="18" y1="25" x2="22" y2="25" stroke="#FF6B00" stroke-width="1.8" stroke-linecap="round" opacity=".5"/><line x1="28" y1="25" x2="32" y2="25" stroke="#FF6B00" stroke-width="1.8" stroke-linecap="round" opacity=".5"/></svg>`;

let cart = JSON.parse(localStorage.getItem('moncar_cart')||'[]');
let filtroLinea='';
let visiblesActuales=[];
let pagMostrados=0;
const PAG_SIZE=40;
function qid(cod){{ return 'qty-'+cod.replace(/[^A-Za-z0-9]/g,'_'); }}

function fmt(n){{ return n ? '$'+Math.round(n).toLocaleString('es-AR') : 'Consultar'; }}

function setNavActive(id){{
  ['inicio','productos','nosotros'].forEach(n=>document.getElementById('nav-'+n).classList.toggle('active',n===id));
}}
function setHash(h){{ if(location.hash.slice(1)!==h) location.hash=h; }}
function irAInicio(){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-inicio').classList.add('active');
  setNavActive('inicio'); window.scrollTo(0,0);
  setHash('inicio');
}}
function irAProductos(cat,marca){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-productos').classList.add('active');
  setNavActive('productos');
  document.querySelectorAll('#sb-cats input,#sb-marcas input').forEach(cb=>cb.checked=false);
  if(cat){{
    const cb=document.querySelector(`#sb-cats input[value="${{cat}}"]`);
    if(cb) cb.checked=true;
  }}
  if(marca){{
    const cb=document.querySelector(`#sb-marcas input[value="${{marca}}"]`);
    if(cb) cb.checked=true;
  }}
  filtrar(); window.scrollTo(0,0);
  if(cat) setHash('productos/cat/'+encodeURIComponent(cat));
  else if(marca) setHash('productos/marca/'+encodeURIComponent(marca));
  else setHash('productos');
}}
function irANosotros(){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-nosotros').classList.add('active');
  setNavActive('nosotros'); window.scrollTo(0,0);
  setHash('nosotros');
}}
function irACat(c){{ irAProductos(c,null); }}
function irAMarca(m){{ irAProductos(null,m); }}

// Ruteo liviano por hash: permite compartir un link directo a Productos /
// una categoria / una marca, y que el boton "atras" del navegador funcione
// entre las 3 paginas (no es SSR real, sigue siendo todo la misma pagina —
// para que Google indexe productos individuales hay que ir a un rediseño
// mas grande, ver charla con Franco 2026-09-24).
function aplicarHash(){{
  const h=location.hash.slice(1);
  if(!h||h==='inicio'){{ irAInicio(); return; }}
  if(h==='nosotros'){{ irANosotros(); return; }}
  if(h==='productos'){{ irAProductos(); return; }}
  let m=h.match(/^productos\\/cat\\/(.+)$/);
  if(m){{ irAProductos(decodeURIComponent(m[1]),null); return; }}
  m=h.match(/^productos\\/marca\\/(.+)$/);
  if(m){{ irAProductos(null,decodeURIComponent(m[1])); return; }}
  irAInicio();
}}
window.addEventListener('hashchange', aplicarHash);

function setLinea(btn){{
  document.querySelectorAll('.sb-linea-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  filtroLinea=btn.dataset.val;
  filtrar();
}}

function filtrar(){{
  const q=document.getElementById('search').value.toLowerCase().trim();
  document.getElementById('clear-btn').style.display=q?'block':'none';

  const selCats=[...document.querySelectorAll('#sb-cats input:checked')].map(c=>c.value);
  const selMarcas=[...document.querySelectorAll('#sb-marcas input:checked')].map(c=>c.value);
  const pMin=parseFloat(document.getElementById('precio-min').value)||0;
  const pMax=parseFloat(document.getElementById('precio-max').value)||Infinity;
  const filEsp=document.getElementById('fil-esp').value;
  const filAnch=document.getElementById('fil-anch').value;
  const filLarg=document.getElementById('fil-larg').value;
  const filTerm=document.getElementById('fil-term').value;

  const af=document.getElementById('active-filters');
  const tags=[];
  if(filtroLinea) tags.push(`<div class="active-tag" onclick="document.querySelector('.sb-linea-btn[data-val=\\'\\']').click()">Linea: ${{filtroLinea}} ✕</div>`);
  selCats.forEach(c=>tags.push(`<div class="active-tag" onclick="uncheckCat('${{c}}')">📂 ${{c}} ✕</div>`));
  selMarcas.forEach(m=>tags.push(`<div class="active-tag" onclick="uncheckMarca('${{m}}')">🏷️ ${{m}} ✕</div>`));
  if(pMin>0||pMax<Infinity) tags.push(`<div class="active-tag" onclick="limpiarPrecio()">Precio ✕</div>`);
  af.innerHTML=tags.join('');
  af.style.display=tags.length?'flex':'none';

  const visibles=PRODUCTOS.filter(p=>{{
    if(filtroLinea&&p.linea!==filtroLinea) return false;
    if(selCats.length&&!selCats.includes(p.cat)) return false;
    if(selMarcas.length&&!selMarcas.includes(p.marca)) return false;
    if(p.precio>0&&(p.precio<pMin||p.precio>pMax)) return false;
    if(filEsp&&String(p.esp)!==filEsp) return false;
    if(filAnch&&String(p.anch)!==filAnch) return false;
    if(filLarg&&String(p.larg)!==filLarg) return false;
    if(filTerm&&String(p.terminal)!==filTerm) return false;
    if(q&&!(p.cod+' '+p.desc+' '+p.marca+' '+p.cat).toLowerCase().includes(q)) return false;
    return true;
  }});

  // Orden por defecto: marca alfabetico, despues codigo (para que quede estable)
  visibles.sort((a,b)=> a.marca.localeCompare(b.marca) || a.cod.localeCompare(b.cod));

  visiblesActuales=visibles;
  pagMostrados=PAG_SIZE;
  document.getElementById('results-info').textContent=visibles.length+' producto'+(visibles.length!==1?'s':'');
  document.getElementById('no-results').style.display=visibles.length?'none':'block';
  document.getElementById('grid').style.display=visibles.length?'grid':'none';
  renderGridPage();
}}

function renderGridPage(){{
  renderGrid(visiblesActuales.slice(0,pagMostrados));
  const wrap=document.getElementById('cargar-mas-wrap');
  if(visiblesActuales.length>pagMostrados){{
    wrap.style.display='block';
    document.getElementById('cargar-mas-info').textContent=
      `Mostrando ${{Math.min(pagMostrados,visiblesActuales.length)}} de ${{visiblesActuales.length}}`;
  }} else {{
    wrap.style.display='none';
  }}
}}

function cargarMas(){{
  pagMostrados+=PAG_SIZE;
  renderGridPage();
}}

function renderGrid(lista){{
  document.getElementById('grid').innerHTML=lista.map(p=>{{
    const svg=p.linea==='RMH'?SVG_RMH:SVG_MC;
    const icon=p.img?`<img src="imagenes/${{p.img}}" alt="${{p.cod}}" loading="lazy">`:svg;
    const enCart=cart.find(c=>c.cod===p.cod);
    const btnTxt=enCart?'Agregar mas':'Agregar';
    const marcaTxt=(p.marca&&!['RMH','Moncar','Electrodomesticos'].includes(p.marca))?p.marca:'';
    return `<div class="card">
      <div class="card-icon">${{icon}}</div>
      <div class="card-body">
        <div class="card-badges"><span class="badge bl">${{p.linea}}</span><span class="badge bc">${{p.cat}}</span></div>
        <div class="card-cod">${{p.cod}}</div>
        <div class="card-desc">${{p.desc}}</div>
        ${{marcaTxt?`<div class="card-sub">${{marcaTxt}}</div>`:''}}
        ${{p.dims?`<div class="card-sub">${{p.dims}} mm</div>`:''}}
        <div class="card-precio">${{fmt(p.precio)}}</div>
        <div class="card-qty">
          <button onclick="cambiarQtyCard('${{p.cod}}',-1)">-</button>
          <input type="number" min="1" value="1" id="${{qid(p.cod)}}">
          <button onclick="cambiarQtyCard('${{p.cod}}',1)">+</button>
        </div>
        <button class="btn-add" onclick="agregarAlCarrito('${{p.cod}}')">${{btnTxt}}</button>
      </div></div>`;
  }}).join('');
}}

function cambiarQtyCard(cod,delta){{
  const inp=document.getElementById(qid(cod));
  if(!inp) return;
  let v=(parseInt(inp.value)||1)+delta;
  if(v<1) v=1;
  inp.value=v;
}}

function limpiarBusqueda(){{ document.getElementById('search').value=''; filtrar(); }}
function limpiarPrecio(){{ document.getElementById('precio-min').value=''; document.getElementById('precio-max').value=''; filtrar(); }}
function uncheckCat(v){{ const cb=document.querySelector(`#sb-cats input[value="${{v}}"]`); if(cb)cb.checked=false; filtrar(); }}
function uncheckMarca(v){{ const cb=document.querySelector(`#sb-marcas input[value="${{v}}"]`); if(cb)cb.checked=false; filtrar(); }}
function limpiarFiltros(){{
  document.querySelectorAll('#sb-cats input,#sb-marcas input').forEach(cb=>cb.checked=false);
  document.getElementById('precio-min').value='';
  document.getElementById('precio-max').value='';
  document.getElementById('fil-esp').value='';
  document.getElementById('fil-anch').value='';
  document.getElementById('fil-larg').value='';
  document.getElementById('fil-term').value='';
  document.getElementById('search').value='';
  document.querySelector('.sb-linea-btn[data-val=""]').click();
}}

function abrirSidebar(){{ document.getElementById('sidebar').classList.add('open'); document.getElementById('sb-overlay').classList.add('open'); }}
function cerrarSidebar(){{ document.getElementById('sidebar').classList.remove('open'); document.getElementById('sb-overlay').classList.remove('open'); }}

// CART
function agregarAlCarrito(cod){{
  const p=PRODUCTOS.find(x=>x.cod===cod); if(!p) return;
  const qtyInp=document.getElementById(qid(cod));
  const cant=qtyInp?Math.max(1,parseInt(qtyInp.value)||1):1;
  const ex=cart.find(c=>c.cod===cod);
  if(ex) ex.qty+=cant; else cart.push({{cod:p.cod,desc:p.desc,precio:p.precio,qty:cant}});
  guardarCart(); renderGridPage(); abrirCart();
}}
function cambiarQty(cod,delta){{
  const ex=cart.find(c=>c.cod===cod); if(!ex) return;
  ex.qty+=delta; if(ex.qty<=0) cart=cart.filter(c=>c.cod!==cod);
  guardarCart();
}}
function quitarItem(cod){{ cart=cart.filter(c=>c.cod!==cod); guardarCart(); }}
function guardarCart(){{
  localStorage.setItem('moncar_cart',JSON.stringify(cart));
  renderCart();
  document.getElementById('cart-count').textContent=cart.reduce((s,c)=>s+c.qty,0);
}}
function renderCart(){{
  const el=document.getElementById('cart-items');
  if(!cart.length){{ el.innerHTML='<div class="cart-empty">El carrito esta vacio</div>'; document.getElementById('cart-total').textContent='$0'; return; }}
  el.innerHTML=cart.map(c=>`<div class="ci">
    <div class="ci-info">
      <div class="ci-cod">${{c.cod}}</div>
      <div class="ci-desc">${{c.desc}}</div>
      <div class="ci-precio">${{fmt(c.precio)}}</div>
      <div class="ci-qty">
        <button onclick="cambiarQty('${{c.cod}}',-1);renderCart();renderGridPage()">-</button>
        <span>${{c.qty}}</span>
        <button onclick="cambiarQty('${{c.cod}}',1);renderCart();renderGridPage()">+</button>
      </div>
    </div>
    <button class="btn-rm" onclick="quitarItem('${{c.cod}}');renderCart();renderGridPage()">✕</button>
  </div>`).join('');
  document.getElementById('cart-total').textContent=fmt(cart.reduce((s,c)=>s+c.precio*c.qty,0));
}}
function abrirCart(){{ document.getElementById('cart-panel').classList.add('open'); document.getElementById('overlay').style.display='block'; renderCart(); }}
function cerrarCart(){{ document.getElementById('cart-panel').classList.remove('open'); document.getElementById('overlay').style.display='none'; }}

// PEDIDO / CONSULTA — toda opcion de WhatsApp del sitio (nav, hero, carrito)
// pasa SIEMPRE por este formulario, para poder identificar a que WhatsApp
// del equipo redirigir segun la zona del cliente (elegirDestino, arriba).
// (la version original de este sitio ruteaba a "hermanos" con datos
// ficticios que no correspondian a nadie del equipo real; se saco esa logica
// y ahora el numero es real, aunque el mapeo por zona todavia esta pendiente)
let _pedidoTipo='';
function pedirPorWA(){{ if(!cart.length) return; _pedidoTipo='wa'; abrirFormModal(); }}
function pedirPorEmail(){{ if(!cart.length) return; _pedidoTipo='email'; abrirFormModal(); }}
function abrirConsultaWA(){{ _pedidoTipo='wa'; abrirFormModal(); }}

function abrirFormModal(){{
  document.getElementById('fmodal-ov').classList.add('open');
  ['f-nombre','f-localidad'].forEach(id=>document.getElementById(id).classList.remove('ferr'));
}}
function cerrarFormModal(){{
  document.getElementById('fmodal-ov').classList.remove('open');
}}
function confirmarPedido(){{
  const nombre=document.getElementById('f-nombre').value.trim();
  const negocio=document.getElementById('f-negocio').value.trim();
  const localidad=document.getElementById('f-localidad').value.trim();
  let ok=true;
  document.getElementById('f-nombre').classList.toggle('ferr',!nombre); if(!nombre) ok=false;
  document.getElementById('f-localidad').classList.toggle('ferr',!localidad); if(!localidad) ok=false;
  if(!ok) return;

  let msg = cart.length ? 'Hola! Nuevo pedido desde la web:\\n\\n' : 'Hola! Quiero hacer una consulta desde la web:\\n\\n';
  msg+='Cliente: '+nombre+'\\n';
  if(negocio) msg+='Negocio: '+negocio+'\\n';
  msg+='Localidad: '+localidad+'\\n';
  if(cart.length){{
    msg+='\\nProductos:\\n';
    cart.forEach(c=>{{msg+=`- ${{c.cod}} x${{c.qty}}\\n`;}});
  }}
  msg+='\\nMuchas gracias!';

  cerrarFormModal();

  if(_pedidoTipo==='wa'){{
    const destino=elegirDestino(localidad);
    window.open('https://wa.me/'+destino+'?text='+encodeURIComponent(msg),'_blank');
  }} else {{
    const subj='Nuevo pedido MONCAR RMH'+(negocio?' - '+negocio:'');
    window.location.href='mailto:'+CONTACTO_EMAIL+'?subject='+encodeURIComponent(subj)+'&body='+encodeURIComponent(msg);
  }}
}}

guardarCart();
if(location.hash) aplicarHash(); else filtrar();
</script>
</body>
</html>'''

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    productos_mc  = leer_moncar()
    productos_rmh = leer_rmh()
    todos = productos_mc + productos_rmh

    print("\n=== Distribucion por categoria ===")
    cats = Counter(p['cat'] for p in todos)
    for cat, cnt in sorted(cats.items(), key=lambda x: -x[1]):
        print(f"  {cat}: {cnt}")

    mc = Counter(p['marca'] for p in todos if p['linea']=='MONCAR' and p['marca'] not in EXCLUIR_MARCAS)
    destacadas = [(m,c) for m,c in mc.items() if m in MARCAS_CONFIG]
    print(f"\n=== Marcas destacadas para homepage ({len(destacadas)}) ===")
    for m,c in sorted(destacadas, key=lambda x: -x[1]):
        print(f"  {m}: {c}")

    print("\nGenerando HTML...")
    html = generar_html(todos)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\nSitio generado: {OUT_PATH} ({len(html)//1024} KB)")
