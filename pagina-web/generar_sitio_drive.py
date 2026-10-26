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
import html as _html
import urllib.request, urllib.parse, urllib.error
from collections import Counter
from paginas_legales import generar_paginas_legales

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
# Generador HTML — rediseño 2026-10-26 (rama rediseno-web).
# Estilo: catálogo industrial sobrio (fondo claro, líneas finas, sin sombras
# ni degradés, bandas oscuras #0f0e12) con el naranja de la marca solo en la
# acción principal. Pensado para ferreteros: letra 16px, botones de 44px+,
# búsqueda guiada por marca / tipo de máquina / medidas (no por código).
#
# CSS y JS van en strings crudos aparte (no f-strings) para no tener que
# duplicar llaves; los datos entran por reemplazo de marcadores __X__.
# El ruteo por zona del WhatsApp sigue igual (ver elegirDestino en el JS).
# ---------------------------------------------------------------------------

# Nombres para mostrar (el valor interno de categoría/marca no cambia, así los
# links #productos/cat/... y #productos/marca/... siguen funcionando).
CAT_LABEL = {
    'Electrodomesticos': 'Electrodomésticos',
    'Maq. de Coser':     'Máquinas de coser',
    'Otros':             'Otras máquinas',
}
MARCA_LABEL = {
    'Black Y Decker': 'Black & Decker',
    'De Walt':        'DeWalt',
}

_SVG_MEDIDAS = '''<svg viewBox="0 0 250 232" role="img" aria-label="Cómo medir un carbón: espesor, ancho y largo">
<path d="M88 60 C88 40 88 34 88 22" stroke="#8a908d" stroke-width="3" fill="none"/>
<rect x="80" y="10" width="16" height="12" fill="#fff" stroke="#111" stroke-width="1.5"/>
<rect x="48" y="60" width="80" height="120" fill="#3b3a40" stroke="#111" stroke-width="1.5"/>
<rect x="176" y="60" width="34" height="120" fill="#5a5960" stroke="#111" stroke-width="1.5"/>
<g stroke="#ff5a00" stroke-width="1.5">
<line x1="48" y1="198" x2="128" y2="198"/><line x1="48" y1="192" x2="48" y2="204"/><line x1="128" y1="192" x2="128" y2="204"/>
<line x1="176" y1="198" x2="210" y2="198"/><line x1="176" y1="192" x2="176" y2="204"/><line x1="210" y1="192" x2="210" y2="204"/>
<line x1="30" y1="60" x2="30" y2="180"/><line x1="24" y1="60" x2="36" y2="60"/><line x1="24" y1="180" x2="36" y2="180"/>
</g>
<g font-family="Inter,sans-serif" font-size="13" font-weight="600" fill="#111" text-anchor="middle">
<text x="88" y="222">Ancho</text><text x="193" y="222">Espesor</text>
<text x="0" y="0" transform="translate(16 120) rotate(-90)">Largo</text>
</g>
</svg>'''

_CSS = r'''
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f6f8f7;--surface:#fff;--line:#e3e6e4;--line-2:#cdd2cf;
  --ink:#111214;--ink-2:#33373a;--muted:#5c6360;
  --carbon:#0f0e12;--carbon-2:#1b1a20;--carbon-line:#2e2d35;
  --orange:#ff5a00;--orange-2:#e04f00;--orange-txt:#c2410c;
  --wa:#1d9e52;--wa-2:#178a47;--ok:#15803d;--err:#c81e1e;
  --r:4px;--head-h:64px;--head-total:64px;--wrap:1240px;
  --font:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  color-scheme:light;
}
html{scroll-behavior:smooth;-webkit-text-size-adjust:100%}
body{background:var(--bg);color:var(--ink);font-family:var(--font);font-size:16px;line-height:1.5}
body.lock{overflow:hidden}
button,input,select{font:inherit;color:inherit}
button{cursor:pointer}
[hidden]{display:none!important}
:focus-visible{outline:2px solid var(--orange);outline-offset:2px}
.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}
.wrap{max-width:var(--wrap);margin:0 auto;padding-left:20px;padding-right:20px}
section[id^="sec-"]{scroll-margin-top:calc(var(--head-total) + 8px)}

/* Botones */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:44px;padding:0 18px;border-radius:var(--r);border:1px solid transparent;background:none;font-size:15px;font-weight:600;text-decoration:none;white-space:nowrap;transition:background-color .15s,border-color .15s,color .15s}
.btn-lg{min-height:52px;padding:0 24px;font-size:16px}
.btn-block{width:100%}
.btn-primary{background:var(--orange);color:#fff}
.btn-primary:hover{background:var(--orange-2)}
.btn-line{background:var(--surface);border-color:var(--line-2);color:var(--ink)}
.btn-line:hover{border-color:var(--ink)}
.btn-wa{background:var(--wa);color:#fff}
.btn-wa:hover{background:var(--wa-2)}
.btn-ghost-dark{border-color:rgba(255,255,255,.4);color:#fff}
.btn-ghost-dark:hover{border-color:#fff}
.wa-ico{color:var(--wa);font-size:18px}

/* Campos */
.inp{width:100%;min-height:48px;border:1px solid var(--line-2);border-radius:var(--r);background:#fff;padding:0 12px;font-size:16px;color:var(--ink)}
.inp:focus{outline:none;border-color:var(--orange);box-shadow:inset 0 0 0 1px var(--orange)}
.inp::placeholder{color:#8b918e}
.inp.err{border-color:var(--err)}
select.inp{padding-right:8px}
.inp-unit{position:relative}
.inp-unit .inp{padding-right:40px}
.inp-unit::after{content:'mm';position:absolute;right:12px;top:50%;transform:translateY(-50%);color:var(--muted);font-size:14px;pointer-events:none}
.fg-err{color:var(--err);font-size:14px;margin-top:6px}

/* ===== Encabezado ===== */
.top{position:sticky;top:0;z-index:100;background:var(--surface);border-bottom:1px solid var(--line)}
.top-in{height:var(--head-h);display:flex;align-items:center;gap:28px}
.logo{background:none;border:0;padding:0;font-size:22px;font-weight:800;letter-spacing:.02em;color:var(--ink)}
.logo span{color:var(--orange)}
.nav{display:flex;height:100%}
.nav-btn{background:none;border:0;border-bottom:2px solid transparent;padding:0 14px;height:100%;font-size:15px;font-weight:500;color:var(--muted)}
.nav-btn:hover{color:var(--ink)}
.nav-btn.active{color:var(--ink);font-weight:600;border-bottom-color:var(--orange)}
.top-acts{margin-left:auto;display:flex;gap:8px}
.cart-n{display:inline-flex;align-items:center;justify-content:center;min-width:22px;height:22px;padding:0 6px;border-radius:11px;background:#fff;color:var(--orange-txt);font-size:13px;font-weight:700}
.tabs{display:none}

/* ===== Inicio ===== */
.hero{background:var(--carbon);color:#fff}
.hero-in{display:grid;grid-template-columns:1.1fr 1fr;gap:56px;align-items:center;padding-top:64px;padding-bottom:64px}
.hero-kicker{font-size:13px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#a9abb1;margin-bottom:16px}
.hero h1{font-size:clamp(32px,4.4vw,52px);font-weight:700;line-height:1.08;letter-spacing:-.015em;margin-bottom:18px}
.hero h1 em{font-style:normal;color:var(--orange)}
.hero-p{font-size:18px;color:#c9cbcf;max-width:540px;margin-bottom:30px}
.hero-btns{display:flex;gap:12px;flex-wrap:wrap}
.paths{display:grid;gap:10px}
.paths-t{font-size:15px;font-weight:600;margin-bottom:2px}
.path{display:flex;align-items:center;gap:16px;width:100%;text-align:left;background:var(--carbon-2);border:1px solid var(--carbon-line);border-radius:var(--r);padding:16px 18px;color:#fff;transition:border-color .15s,background-color .15s}
.path:hover{border-color:var(--orange);background:#222129}
.path-ico{width:46px;height:46px;flex-shrink:0;display:flex;align-items:center;justify-content:center;border:1px solid #3a3942;border-radius:var(--r);color:var(--orange);font-size:20px}
.path-tx strong{display:block;font-size:17px;font-weight:600;line-height:1.3}
.path-tx span{font-size:14px;color:#a9abb1}
.path-arrow{margin-left:auto;color:#77767f}
.hero-search{margin-top:8px;padding-top:16px;border-top:1px solid var(--carbon-line)}
.hero-search label{display:block;font-size:14px;color:#c9cbcf;margin-bottom:8px}
.hs-row{display:flex;gap:8px}
.hs-row .inp{flex:1;min-width:0;border-color:#3a3942}

.stats{background:var(--surface);border-bottom:1px solid var(--line)}
.stats-in{display:grid;grid-template-columns:repeat(4,1fr)}
.stat{display:flex;gap:14px;align-items:center;padding:22px 20px;border-left:1px solid var(--line)}
.stat:first-child{border-left:0;padding-left:0}
.stat>i{color:var(--orange);font-size:22px;width:26px;text-align:center}
.stat strong{display:block;font-size:16px}
.stat span{font-size:14px;color:var(--muted)}

.sec{padding:64px 0}
.sec+.sec{border-top:1px solid var(--line)}
.sec-head{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:26px}
.sec-num{font-size:13px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--orange-txt);margin-bottom:8px}
.sec h2{font-size:clamp(24px,3vw,32px);font-weight:700;line-height:1.2;letter-spacing:-.01em}
.sec-sub{color:var(--muted);margin-top:6px;max-width:640px}
.link-more{background:none;border:0;padding:8px 0;font-size:15px;font-weight:600;color:var(--ink);text-decoration:underline;text-underline-offset:4px;white-space:nowrap}
.link-more:hover{color:var(--orange-txt)}

/* Grillas con líneas finas (fijas: 6 marcas y 5 categorías por fila) */
.ruled{display:grid;background:var(--surface);border-top:1px solid var(--line);border-left:1px solid var(--line)}
.ruled>*{border:0;border-right:1px solid var(--line);border-bottom:1px solid var(--line);background:var(--surface);text-align:left;transition:background-color .15s}
.marca-grid{grid-template-columns:repeat(6,1fr)}
.cat-grid{grid-template-columns:repeat(5,1fr)}
.marca-card{position:relative;padding:24px 18px 20px}
.marca-card::before{content:'';position:absolute;left:0;top:0;right:0;height:3px;background:var(--mc,var(--line))}
.marca-card:hover,.cat-card:hover{background:#fafbfa}
.marca-name{display:block;font-size:17px;font-weight:700}
.marca-sub{display:block;font-size:14px;color:var(--muted);margin-top:2px}
.cat-card{display:flex;flex-direction:column;gap:12px;padding:22px 18px}
.cat-icon{font-size:30px;line-height:1;height:30px;color:var(--orange)}
.cat-icon svg{width:30px;height:30px}
.cat-name{font-size:16px;font-weight:600;line-height:1.3}
.cat-count{font-size:14px;color:var(--muted);margin-top:-8px}

.med-box{display:grid;grid-template-columns:320px 1fr;background:var(--surface);border:1px solid var(--line)}
.med-fig{display:flex;align-items:center;justify-content:center;padding:28px;border-right:1px solid var(--line);background:#fbfcfb}
.med-fig svg{width:100%;max-width:250px;height:auto}
.med-form{padding:32px}
.med-fields{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}
.field label{display:block;font-size:15px;font-weight:600;margin-bottom:6px}
.field small{font-size:13px;font-weight:400;color:var(--muted)}
.med-help{font-size:15px;color:var(--muted);margin-bottom:20px;max-width:560px}

.ayuda{background:var(--carbon);color:#fff}
.ayuda-in{display:flex;align-items:center;justify-content:space-between;gap:24px;flex-wrap:wrap;padding-top:44px;padding-bottom:44px}
.ayuda h2{font-size:24px;font-weight:700;line-height:1.25}
.ayuda p{color:#c9cbcf;margin-top:6px;max-width:620px}

/* ===== Productos ===== */
.prod-head{background:var(--surface);border-bottom:1px solid var(--line)}
.prod-head-in{padding-top:26px;padding-bottom:20px}
.prod-head h1{font-size:28px;font-weight:700;letter-spacing:-.01em}
.prod-head p{color:var(--muted)}
.search{position:relative;margin-top:16px;max-width:760px}
.search>i{position:absolute;left:15px;top:50%;transform:translateY(-50%);color:var(--muted)}
.search .inp{padding-left:44px;padding-right:48px;min-height:52px}
.search-clear{position:absolute;right:4px;top:50%;transform:translateY(-50%);width:44px;height:44px;border:0;background:none;color:var(--muted);font-size:18px}
.prod-layout{display:grid;grid-template-columns:290px 1fr;gap:28px;align-items:start;padding-top:24px;padding-bottom:64px}
.filtros{position:sticky;top:calc(var(--head-total) + 16px);max-height:calc(100vh - var(--head-total) - 32px);overflow:auto;background:var(--surface);border:1px solid var(--line)}
.f-head,.f-foot{display:none}
.f-sec{padding:18px;border-bottom:1px solid var(--line)}
.f-tit{display:flex;align-items:center;gap:8px;font-size:13px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-bottom:12px}
.f-tit i{color:var(--orange);width:16px;text-align:center}
.f-med{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.f-med label,.f-lbl{display:block;font-size:14px;font-weight:600;margin-bottom:4px}
.f-med .inp{min-height:44px;padding:0 8px}
.f-help{font-size:13px;color:var(--muted);margin:8px 0 14px}
.chk{display:flex;align-items:center;gap:10px;min-height:38px;font-size:15px;cursor:pointer}
.chk input{width:18px;height:18px;accent-color:var(--orange);flex-shrink:0;cursor:pointer}
.chk span:first-of-type{flex:1}
.chk .n{font-size:13px;color:var(--muted)}
.f-more{background:none;border:0;padding:8px 0 0;font-size:14px;font-weight:600;text-decoration:underline;text-underline-offset:3px}
.seg{display:grid;gap:6px}
.seg button{min-height:42px;border:1px solid var(--line-2);border-radius:var(--r);background:#fff;padding:0 12px;text-align:left;font-size:15px}
.seg button:hover{border-color:var(--ink)}
.seg button.active{background:var(--ink);border-color:var(--ink);color:#fff;font-weight:600}
.f-range{display:flex;align-items:center;gap:8px}
.f-range .inp{min-height:44px}
.f-range span{color:var(--muted)}
.f-ov{display:none}

.toolbar{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px}
.res-n{font-weight:600}
.orden{margin-left:auto;display:flex;align-items:center;gap:8px;font-size:14px;color:var(--muted)}
.orden .inp{min-height:42px;width:auto;font-size:15px}
.btn-filtros{display:none}
.f-count{display:inline-flex;align-items:center;justify-content:center;min-width:22px;height:22px;border-radius:11px;background:var(--orange);color:#fff;font-size:13px;font-weight:700;padding:0 6px}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:16px}
.chip{display:inline-flex;align-items:center;gap:8px;min-height:36px;padding:0 12px;border:1px solid var(--line-2);border-radius:var(--r);background:#fff;font-size:14px}
.chip i{color:var(--muted)}
.chip:hover{border-color:var(--ink)}
.chip-all{background:none;border:0;min-height:36px;padding:0 4px;font-size:14px;font-weight:600;text-decoration:underline;text-underline-offset:3px}

.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(236px,1fr));gap:16px}
.card{display:flex;flex-direction:column;background:var(--surface);border:1px solid var(--line);border-radius:var(--r);transition:border-color .15s}
.card:hover{border-color:var(--line-2)}
.card-img{height:180px;display:flex;align-items:center;justify-content:center;padding:12px;border-bottom:1px solid var(--line)}
.card-img img{max-width:100%;max-height:100%;object-fit:contain}
.card-img svg{width:56px;height:56px}
.card-body{flex:1;display:flex;flex-direction:column;gap:4px;padding:14px}
.card-tipo{font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
.card-marca{font-size:17px;font-weight:700;line-height:1.25}
.card-desc{font-size:14px;line-height:1.4;color:var(--ink-2);overflow-wrap:anywhere}
.card-dims{font-size:14px;font-weight:500;margin-top:2px}
.card-dims i{color:var(--muted);margin-right:4px}
.tag{align-self:flex-start;margin-top:2px;padding:2px 8px;border:1px solid var(--line-2);border-radius:var(--r);font-size:12px;font-weight:600;color:var(--muted)}
.tag-ok{border-color:#bfe5cb;background:#effaf2;color:var(--ok)}
.card-cod{margin-top:auto;padding-top:6px;font-size:13px;color:var(--muted)}
.card-precio{font-size:22px;font-weight:700;letter-spacing:-.01em}
.card-buy{display:flex;gap:8px;margin-top:8px}
.card-buy .btn{flex:1;padding:0 10px}
.card-encart{font-size:13px;font-weight:500;color:var(--ok);margin-top:4px}
.qty{display:flex;height:44px;border:1px solid var(--line-2);border-radius:var(--r);background:#fff;flex-shrink:0}
.qty button{width:36px;border:0;background:none;font-size:20px;line-height:1}
.qty button:hover{background:var(--bg)}
.qty input,.qty span{width:42px;border:0;border-left:1px solid var(--line);border-right:1px solid var(--line);text-align:center;font-size:16px;font-weight:600;-moz-appearance:textfield;display:flex;align-items:center;justify-content:center}
.qty input::-webkit-inner-spin-button,.qty input::-webkit-outer-spin-button{-webkit-appearance:none;margin:0}
.mas{text-align:center;margin-top:28px}
.mas p{font-size:14px;color:var(--muted);margin-bottom:10px}
.vacio{background:var(--surface);border:1px solid var(--line);padding:44px 24px;text-align:center}
.vacio h3{font-size:20px;margin-bottom:8px}
.vacio p{color:var(--muted);max-width:480px;margin:0 auto 20px}
.vacio .acts{display:flex;gap:10px;justify-content:center;flex-wrap:wrap}

/* ===== Nosotros ===== */
.nos{display:grid;grid-template-columns:1.25fr 1fr;gap:44px;align-items:start}
.nos h1{font-size:clamp(28px,3.6vw,40px);font-weight:700;line-height:1.15;letter-spacing:-.01em;margin-bottom:16px}
.nos-lead{font-size:18px;color:var(--ink-2);margin-bottom:28px}
.nos-list{grid-template-columns:1fr 1fr}
.nos-item{padding:22px}
.nos-item>i{display:block;color:var(--orange);font-size:22px;margin-bottom:12px}
.nos-item h3{font-size:16px;font-weight:700;margin-bottom:4px}
.nos-item p{font-size:14px;color:var(--muted)}
.contacto{background:var(--surface);border:1px solid var(--line)}
.contacto h2{font-size:18px;font-weight:700;padding:18px 20px;border-bottom:1px solid var(--line)}
.c-row{display:flex;align-items:center;gap:14px;width:100%;padding:16px 20px;border:0;border-bottom:1px solid var(--line);background:none;text-align:left;text-decoration:none;color:inherit}
.c-row:last-child{border-bottom:0}
.c-row:hover .c-val{color:var(--orange-txt)}
.c-ico{width:44px;height:44px;flex-shrink:0;display:flex;align-items:center;justify-content:center;border:1px solid var(--line);border-radius:var(--r);color:var(--orange);font-size:18px}
.c-lbl{display:block;font-size:13px;color:var(--muted)}
.c-val{display:block;font-weight:600;overflow-wrap:anywhere}

/* ===== Pie ===== */
footer{background:var(--carbon);color:#c9cbcf}
.f-in{display:grid;grid-template-columns:1.4fr 1fr 1fr 1fr;gap:28px;padding-top:48px;padding-bottom:28px}
.f-brand{font-size:20px;font-weight:800;color:#fff;margin-bottom:8px}
.f-brand span{color:var(--orange)}
.f-in p{font-size:14px;color:#a9abb1;max-width:300px}
footer h3{font-size:13px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#8d8f96;margin-bottom:10px}
.f-link{display:block;padding:5px 0;border:0;background:none;color:#e6e7e9;font-size:14px;text-align:left;text-decoration:none}
.f-link:hover{color:#fff;text-decoration:underline}
.f-bot{border-top:1px solid var(--carbon-line);padding-top:16px;padding-bottom:20px;font-size:13px;color:#8d8f96}
.f-bot a{color:inherit}

/* ===== Pedido (panel lateral) ===== */
.ov{position:fixed;inset:0;background:rgba(15,14,18,.5);z-index:250}
.drawer{position:fixed;top:0;right:0;z-index:260;width:min(440px,100%);height:100%;display:flex;flex-direction:column;background:#fff;border-left:1px solid var(--line);transform:translateX(100%);visibility:hidden;transition:transform .25s,visibility 0s .25s}
.drawer.open{transform:none;visibility:visible;transition:transform .25s}
.dr-head{display:flex;align-items:center;justify-content:space-between;padding:10px 12px 10px 20px;border-bottom:1px solid var(--line)}
.dr-head h2{font-size:20px;font-weight:700}
.x{width:44px;height:44px;border:0;background:none;color:var(--muted);font-size:20px}
.x:hover{color:var(--ink)}
.dr-items{flex:1;overflow:auto;padding:0 20px}
.ci{display:grid;grid-template-columns:60px 1fr auto;gap:12px;align-items:start;padding:16px 0;border-bottom:1px solid var(--line)}
.ci-img{width:60px;height:60px;display:flex;align-items:center;justify-content:center;padding:4px;border:1px solid var(--line)}
.ci-img img{max-width:100%;max-height:100%;object-fit:contain}
.ci-t{font-size:15px;font-weight:600;line-height:1.3}
.ci-cod{font-weight:400;color:var(--muted);font-size:13px}
.ci-d{font-size:13px;color:var(--muted);overflow-wrap:anywhere}
.ci-p{font-size:14px;margin-top:4px}
.ci .qty{height:40px;margin-top:8px;width:max-content}
.btn-rm{width:40px;height:40px;border:0;background:none;color:var(--muted)}
.btn-rm:hover{color:var(--err)}
.dr-vacio{padding:56px 0;text-align:center;color:var(--muted)}
.dr-vacio p{margin-bottom:16px}
.dr-foot{display:grid;gap:8px;padding:16px 20px calc(16px + env(safe-area-inset-bottom));border-top:1px solid var(--line)}
.dr-total{display:flex;justify-content:space-between;align-items:baseline}
.dr-total span{font-size:15px;color:var(--muted)}
.dr-total strong{font-size:26px;letter-spacing:-.01em}
.dr-note{font-size:13px;color:var(--muted);margin-bottom:6px}
.dr-vaciar{justify-self:center;background:none;border:0;padding:6px;font-size:14px;color:var(--muted);text-decoration:underline}

/* ===== Formulario ===== */
.modal{position:fixed;inset:0;z-index:400;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(15,14,18,.55)}
.modal-box{width:min(480px,100%);max-height:calc(100vh - 32px);overflow:auto;background:#fff;border-radius:var(--r);padding:28px}
.modal h2{font-size:22px;font-weight:700;line-height:1.25;margin-bottom:6px}
.modal-sub{font-size:15px;color:var(--muted);margin-bottom:20px}
.fg{margin-bottom:16px}
.fg label{display:block;font-size:15px;font-weight:600;margin-bottom:6px}
.fg .opt{font-weight:400;color:var(--muted)}
.aviso{font-size:13px;color:var(--muted)}
.aviso a{color:var(--orange-txt)}
.modal-acts{display:flex;gap:10px;margin-top:22px}
.modal-acts .btn:last-child{flex:1}

/* ===== Aviso "agregado" y barra de pedido (celular) ===== */
.toast{position:fixed;left:50%;bottom:20px;z-index:350;display:flex;align-items:center;gap:14px;width:max-content;max-width:calc(100% - 32px);padding:10px 10px 10px 18px;border-radius:var(--r);background:var(--carbon);color:#fff;font-size:15px;opacity:0;pointer-events:none;transform:translate(-50%,16px);transition:opacity .2s,transform .2s}
.toast.show{opacity:1;pointer-events:auto;transform:translate(-50%,0)}
.toast i{color:#4ade80}
.toast button{min-height:38px;padding:0 12px;border:1px solid rgba(255,255,255,.4);border-radius:var(--r);background:none;color:#fff;font-size:14px;font-weight:600;white-space:nowrap}
.pbar{display:none}

/* ===== Responsive ===== */
@media(max-width:1000px){
  .marca-grid{grid-template-columns:repeat(3,1fr)}
  .cat-grid{grid-template-columns:repeat(2,1fr)}
  .hero-in{grid-template-columns:1fr;gap:36px;padding-top:44px;padding-bottom:44px}
  .stats-in{grid-template-columns:1fr 1fr}
  .stat{border-left:0;border-top:1px solid var(--line);padding:16px 0}
  .stat:nth-child(-n+2){border-top:0}
  .prod-layout{grid-template-columns:1fr;padding-top:16px}
  .btn-filtros{display:inline-flex}
  .f-ov{display:block;position:fixed;inset:0;z-index:290;background:rgba(15,14,18,.5)}
  .filtros{position:fixed;top:0;left:0;bottom:0;z-index:300;width:min(400px,100%);max-height:none;border:0;display:flex;flex-direction:column;transform:translateX(-100%);visibility:hidden;transition:transform .25s,visibility 0s .25s}
  .filtros.open{transform:none;visibility:visible;transition:transform .25s}
  .f-head{display:flex;align-items:center;justify-content:space-between;padding:8px 8px 8px 18px;border-bottom:1px solid var(--line);font-size:18px}
  .f-body{flex:1;overflow:auto}
  .f-foot{display:grid;grid-template-columns:auto 1fr;gap:10px;padding:12px 18px calc(12px + env(safe-area-inset-bottom));border-top:1px solid var(--line)}
  .f-clear-d{display:none}
  .nos{grid-template-columns:1fr}
  .f-in{grid-template-columns:1fr 1fr}
  .f-in>div:first-child{grid-column:1/-1}
}
@media(max-width:760px){
  :root{--head-h:56px;--head-total:101px}
  .wrap{padding-left:16px;padding-right:16px}
  .top-in{gap:12px}
  .logo{font-size:19px}
  .nav{display:none}
  .tabs{display:flex;border-top:1px solid var(--line)}
  .tabs .nav-btn{flex:1;height:44px;padding:0}
  .top-acts .btn{padding:0 12px}
  .hide-sm{display:none}
  .sec{padding:44px 0}
  .sec-head{flex-direction:column;align-items:flex-start;gap:8px}
  .med-box{grid-template-columns:1fr}
  .med-fig{border-right:0;border-bottom:1px solid var(--line);padding:16px}
  .med-fig svg{max-width:200px}
  .med-form{padding:20px}
  .med-fields{gap:8px}
  .med-form .btn{width:100%}
  .ayuda-in .btn{width:100%}
  .orden span{display:none}
  body.con-pedido{padding-bottom:76px}
  body.con-pedido .pbar{display:flex}
  .pbar{position:fixed;left:0;right:0;bottom:0;z-index:200;align-items:center;gap:12px;padding:10px 16px calc(10px + env(safe-area-inset-bottom));background:var(--surface);border-top:1px solid var(--line)}
  .pbar-tx{flex:1;font-size:14px;line-height:1.3;color:var(--muted)}
  .pbar-tx strong{display:block;font-size:18px;color:var(--ink)}
  body.con-pedido .toast{bottom:88px}
}
@media(max-width:640px){
  .grid{grid-template-columns:1fr;gap:10px}
  .card{display:grid;grid-template-columns:104px 1fr;align-items:start}
  .card-img{height:auto;aspect-ratio:1;border-bottom:0;border-right:1px solid var(--line);padding:8px}
  .card-body{padding:12px}
  .card-marca{font-size:16px}
  .card-precio{font-size:20px}
}
@media(max-width:560px){
  .marca-grid,.cat-grid{grid-template-columns:repeat(2,1fr)}
  .nos-list{grid-template-columns:1fr}
  .hero-btns .btn{width:100%}
}
@media(prefers-reduced-motion:reduce){
  html{scroll-behavior:auto}
  *,*::before,*::after{transition:none!important}
}
'''

_JS = r'''
const PRODUCTOS = __PRODUCTOS__;
const CAT_LABEL = __CAT_LABEL__;
const MARCA_LABEL = __MARCA_LABEL__;
const CONTACTO_EMAIL = 'contacto.carbonkey@gmail.com';
// Equipo real (sacado de la lista de precios RMH). El negocio se maneja por
// zonas, cada uno tiene su zona — pero todavia no tenemos cargado el mapa de
// "que zona le corresponde a quien", asi que por ahora TODO pedido cae en
// Franco (default) hasta que se defina el resto. Cuando este el mapa, cambiar
// elegirDestino() para que devuelva el contacto correcto segun localidad.
const CONTACTOS = {
  FRANCO: '5491136267009',
  ANDI:   '5491139219269',
  DAVID:  '5491150225627',
  MAXI:   '5491144712072',
};
function elegirDestino(localidad){
  return CONTACTOS.FRANCO;
}

// "Marcas" de la hoja que en realidad son tipos de producto: no se muestran
// como marca de la máquina en la tarjeta.
const NO_MARCA = new Set(['RMH','Moncar','Electrodomesticos','Cobreados','Porta Carbones','Power Tools',
  'Lavarropas','Aspiradora','Ventilador','Batidora','Maquina De Coser','Lustra Aspiradora','Hidrolavadora',
  'Enceradoras','Afeitadoras','Bordeadora','Barre Alfombras']);

const SVG_MC=`<svg viewBox="0 0 50 50" fill="none" aria-hidden="true"><rect x="16" y="14" width="18" height="28" stroke="#9aa09d" stroke-width="1.8"/><path d="M25 14V5" stroke="#9aa09d" stroke-width="1.8"/></svg>`;
const SVG_RMH=`<svg viewBox="0 0 50 50" fill="none" aria-hidden="true"><rect x="9" y="15" width="32" height="20" rx="2" stroke="#9aa09d" stroke-width="1.8"/><rect x="18" y="21" width="14" height="8" stroke="#9aa09d" stroke-width="1.8"/></svg>`;

const PAG_SIZE = 40;
const TOL_EA = 0.3;  // tolerancia en mm para espesor y ancho
const TOL_L  = 1.5;  // el largo varía más (un carbón gastado mide menos)

let cart = leerCart();
let filtroLinea = '';
let visiblesActuales = [];
let pagMostrados = 0;
let ultimoFocus = null;
let tFiltro, tToast;

function arriba(){ window.scrollTo({top: 0, behavior: 'instant'}); }

const $ = id => document.getElementById(id);
function norm(t){ return String(t||'').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,''); }
function esc(t){ return String(t).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function jsArg(t){ return esc(JSON.stringify(t)); }
function catLabel(c){ return CAT_LABEL[c] || c; }
function marcaLabel(m){ return MARCA_LABEL[m] || m; }
function qid(cod){ return 'qty-'+cod.replace(/[^A-Za-z0-9]/g,'_'); }
function fmt(n){ return n ? '$'+Math.round(n).toLocaleString('es-AR') : 'Consultar'; }
function fmtMm(v){ return String(Math.round(v*100)/100).replace('.', ','); }
function parseMm(v){ const n = parseFloat(String(v||'').trim().replace(',', '.')); return isFinite(n) && n > 0 ? n : null; }
function parsePrecio(v){ const n = parseFloat(String(v||'').replace(/[$\s.]/g,'').replace(',', '.')); return isFinite(n) ? n : null; }
function marcaMaquina(p){ return NO_MARCA.has(p.marca) ? '' : marcaLabel(p.marca); }
function tipoProd(p){ return p.linea === 'RMH' ? 'Interruptor RMH' : 'Carbón MONCAR'; }

// Texto de búsqueda precalculado: sin acentos, con sinónimos de uso común.
PRODUCTOS.forEach(p => {
  const extra = p.linea === 'RMH' ? 'interruptor interruptores llave llaves tecla' : 'carbon carbones escobilla escobillas';
  p._txt = norm([p.cod, p.cod.replace(/\s+/g,''), p.desc.replace(/b\s*&\s*d/gi, 'b&d black decker'), p.marca, marcaLabel(p.marca),
                 p.marca.replace(/\s+/g,''), p.cat, catLabel(p.cat), extra].join(' '));
});

// ---------- Navegación (una sola página, ruteo por #hash) ----------
const PAGINAS = ['inicio','productos','nosotros'];
let hashPropio = false;
function setHash(h){ if(location.hash.slice(1) !== h){ hashPropio = true; location.hash = h; } }
function mostrarPagina(id){
  PAGINAS.forEach(n => {
    $('page-'+n).hidden = n !== id;
    document.querySelectorAll('[data-nav="'+n+'"]').forEach(b => {
      b.classList.toggle('active', n === id);
      if(n === id) b.setAttribute('aria-current','page'); else b.removeAttribute('aria-current');
    });
  });
  cerrarFiltros();
}
function irAInicio(seccion){
  mostrarPagina('inicio');
  setHash('inicio');
  const el = seccion && $(seccion);
  if(el) el.scrollIntoView(); else arriba();
}
function irANosotros(){ mostrarPagina('nosotros'); setHash('nosotros'); arriba(); }
// Sin argumentos conserva los filtros que había; con categoría, marca,
// medida o búsqueda arranca de cero con ese filtro.
function irAProductos(cat, marca, opts){
  opts = opts || {};
  mostrarPagina('productos');
  if(cat || marca || opts.medida || opts.q != null){
    resetFiltros();
    if(cat) marcarCheck('f-cats', cat);
    if(marca) marcarCheck('f-marcas', marca);
    if(opts.medida){
      $('m-esp').value  = opts.medida.esp  ? fmtMm(opts.medida.esp)  : '';
      $('m-anch').value = opts.medida.anch ? fmtMm(opts.medida.anch) : '';
      $('m-larg').value = opts.medida.larg ? fmtMm(opts.medida.larg) : '';
    }
    if(opts.q != null) $('search').value = opts.q;
  }
  filtrar();
  arriba();
  if(cat) setHash('productos/cat/'+encodeURIComponent(cat));
  else if(marca) setHash('productos/marca/'+encodeURIComponent(marca));
  else if(opts.medida) setHash('productos/medidas/'+[opts.medida.esp,opts.medida.anch,opts.medida.larg].map(v => v || '').join('x'));
  else setHash('productos');
}
function irACat(c){ irAProductos(c, null); }
function irAMarca(m){ irAProductos(null, m); }
function marcarCheck(cont, val){
  const cb = [...document.querySelectorAll('#'+cont+' input')].find(i => i.value === val);
  if(!cb) return;
  cb.checked = true;
  if(cb.closest('#f-marcas-mas')) mostrarTodasMarcas(true);
}
function aplicarHash(){
  const h = location.hash.slice(1);
  if(!h || h === 'inicio'){ mostrarPagina('inicio'); arriba(); return; }
  if(h === 'nosotros'){ irANosotros(); return; }
  if(h === 'productos'){ irAProductos(); return; }
  let m = h.match(/^productos\/cat\/(.+)$/);
  if(m){ irAProductos(decodeURIComponent(m[1]), null); return; }
  m = h.match(/^productos\/marca\/(.+)$/);
  if(m){ irAProductos(null, decodeURIComponent(m[1])); return; }
  m = h.match(/^productos\/medidas\/([\d.]*)x([\d.]*)x([\d.]*)$/);
  if(m){ irAProductos(null, null, {medida:{esp:parseMm(m[1]), anch:parseMm(m[2]), larg:parseMm(m[3])}}); return; }
  irAInicio();
}
window.addEventListener('hashchange', () => { if(hashPropio){ hashPropio = false; return; } aplicarHash(); });

// ---------- Buscadores de la portada ----------
function buscarHome(ev){ ev.preventDefault(); irAProductos(null, null, {q: $('h-q').value.trim()}); }
function buscarMedidaHome(ev){
  ev.preventDefault();
  const m = {esp: parseMm($('h-esp').value), anch: parseMm($('h-anch').value), larg: parseMm($('h-larg').value)};
  if(!m.esp && !m.anch){ $('h-med-err').hidden = false; $('h-esp').focus(); return; }
  $('h-med-err').hidden = true;
  irAProductos(null, null, {medida: m});
}

// ---------- Filtros ----------
function filtrarPronto(){ clearTimeout(tFiltro); tFiltro = setTimeout(filtrar, 200); }
function checks(cont){ return [...document.querySelectorAll('#'+cont+' input:checked')].map(c => c.value); }
function leerMedida(){
  const m = {esp: parseMm($('m-esp').value), anch: parseMm($('m-anch').value), larg: parseMm($('m-larg').value)};
  return (m.esp || m.anch || m.larg) ? m : null;
}
// Distancia entre la medida buscada y el producto (null = no coincide).
// Acepta espesor y ancho invertidos, porque es fácil confundirlos al medir.
function distanciaMedida(p, m){
  if(m.esp || m.anch){
    if(p.esp == null || p.anch == null) return null;
  }
  const dist = (pe, pa) => {
    let d = 0;
    if(m.esp){ const x = Math.abs(pe - m.esp); if(x > TOL_EA) return null; d += x; }
    if(m.anch){ const x = Math.abs(pa - m.anch); if(x > TOL_EA) return null; d += x; }
    return d;
  };
  let d = (m.esp || m.anch) ? dist(p.esp, p.anch) : 0;
  if(m.esp || m.anch){ const d2 = dist(p.anch, p.esp); if(d == null || (d2 != null && d2 < d)) d = d2; }
  if(d == null) return null;
  if(m.larg){
    if(p.larg == null) return null;
    const x = Math.abs(p.larg - m.larg);
    if(x > TOL_L) return null;
    d += x / 3;
  }
  return d;
}
function filtrar(){
  const q = norm($('search').value).trim();
  const tokens = q.split(/\s+/).filter(Boolean);
  $('search-clear').hidden = !q;
  const selCats = checks('f-cats'), selMarcas = checks('f-marcas');
  const pMin = parsePrecio($('precio-min').value) || 0;
  const pMax = parsePrecio($('precio-max').value) || Infinity;
  const term = $('fil-term').value;
  const med = leerMedida();

  const res = [];
  for(const p of PRODUCTOS){
    if(filtroLinea && p.linea !== filtroLinea) continue;
    if(selCats.length && !selCats.includes(p.cat)) continue;
    if(selMarcas.length && !selMarcas.includes(p.marca)) continue;
    if(p.precio > 0 && (p.precio < pMin || p.precio > pMax)) continue;
    if(term && String(p.terminal) !== term) continue;
    if(tokens.length && !tokens.every(t => p._txt.includes(t))) continue;
    let dist = null;
    if(med){ dist = distanciaMedida(p, med); if(dist == null) continue; }
    res.push({p, dist});
  }
  ordenar(res);

  visiblesActuales = res;
  pagMostrados = PAG_SIZE;
  const n = res.length;
  $('results-info').textContent = n + (n === 1 ? ' producto' : ' productos');
  $('btn-ver-res').textContent = 'Ver ' + n + (n === 1 ? ' producto' : ' productos');
  renderChips(selCats, selMarcas, med, term, pMin, pMax, q);
  renderVacio(n, med);
  renderGridPage();
}
function ordenar(res){
  const o = $('orden').value;
  const porMarca = (a, b) => a.p.marca.localeCompare(b.p.marca) || a.p.cod.localeCompare(b.p.cod);
  if(o === 'precio-asc') res.sort((a, b) => (a.p.precio || Infinity) - (b.p.precio || Infinity) || porMarca(a, b));
  else if(o === 'precio-desc') res.sort((a, b) => b.p.precio - a.p.precio || porMarca(a, b));
  else res.sort((a, b) => ((a.dist != null && b.dist != null) ? a.dist - b.dist : 0) || porMarca(a, b));
}
function renderChips(selCats, selMarcas, med, term, pMin, pMax, q){
  const chip = (txt, fn) => `<button type="button" class="chip" onclick="${fn}">${esc(txt)} <i class="fa-solid fa-xmark" aria-hidden="true"></i><span class="sr-only">(quitar)</span></button>`;
  const ch = [];
  if(q) ch.push(chip('Búsqueda: '+$('search').value.trim(), 'limpiarBusqueda()'));
  if(med) ch.push(chip('Medidas: '+[med.esp, med.anch, med.larg].map(v => v ? fmtMm(v) : '—').join(' × ')+' mm', 'limpiarMedida()'));
  if(term) ch.push(chip('Terminal '+term, "$('fil-term').value='';filtrar()"));
  selCats.forEach(c => ch.push(chip(catLabel(c), 'quitarCheck(\'f-cats\','+jsArg(c)+')')));
  selMarcas.forEach(m => ch.push(chip(marcaLabel(m), 'quitarCheck(\'f-marcas\','+jsArg(m)+')')));
  if(filtroLinea) ch.push(chip(filtroLinea === 'RMH' ? 'Interruptores (RMH)' : 'Carbones (MONCAR)', "setLineaVal('');filtrar()"));
  if(pMin > 0 || pMax < Infinity) ch.push(chip('Precio', 'limpiarPrecio()'));
  if(ch.length > 1) ch.push('<button type="button" class="chip-all" onclick="limpiarFiltros()">Limpiar todo</button>');
  const el = $('chips');
  el.innerHTML = ch.join('');
  el.hidden = !ch.length;
  const nf = selCats.length + selMarcas.length + (med ? 1 : 0) + (term ? 1 : 0) + (filtroLinea ? 1 : 0) + ((pMin > 0 || pMax < Infinity) ? 1 : 0);
  $('f-count').textContent = nf;
  $('f-count').hidden = !nf;
}
function renderVacio(n, med){
  const el = $('no-results');
  el.hidden = n > 0;
  $('grid').hidden = n === 0;
  if(n) return;
  const consejo = med
    ? 'Revisá la medida o probá dejando el largo vacío (si el carbón está gastado mide menos).'
    : 'Probá sacando algún filtro o buscando con otras palabras.';
  el.innerHTML = `<h3>No encontramos productos con esos filtros</h3>
    <p>${consejo} Si no lo encontrás, mandanos una foto del carbón o de la máquina por WhatsApp y te decimos cuál es.</p>
    <div class="acts"><button type="button" class="btn btn-line" onclick="limpiarFiltros()">Limpiar filtros</button>
    <button type="button" class="btn btn-wa" onclick="abrirConsultaWA()"><i class="fa-brands fa-whatsapp" aria-hidden="true"></i> Consultar por WhatsApp</button></div>`;
}
function resetFiltros(){
  document.querySelectorAll('#f-cats input,#f-marcas input').forEach(cb => cb.checked = false);
  ['precio-min','precio-max','m-esp','m-anch','m-larg','search'].forEach(id => $(id).value = '');
  $('fil-term').value = '';
  setLineaVal('');
}
function limpiarFiltros(){ resetFiltros(); filtrar(); }
function limpiarBusqueda(){ $('search').value = ''; filtrar(); $('search').focus(); }
function limpiarPrecio(){ $('precio-min').value = ''; $('precio-max').value = ''; filtrar(); }
function limpiarMedida(){ ['m-esp','m-anch','m-larg'].forEach(id => $(id).value = ''); filtrar(); }
function quitarCheck(cont, v){ const cb = [...document.querySelectorAll('#'+cont+' input')].find(i => i.value === v); if(cb) cb.checked = false; filtrar(); }
function setLinea(btn){ setLineaVal(btn.dataset.val); filtrar(); }
function setLineaVal(v){
  filtroLinea = v;
  document.querySelectorAll('#f-linea button').forEach(b => { const on = b.dataset.val === v; b.classList.toggle('active', on); b.setAttribute('aria-pressed', on); });
}
function mostrarTodasMarcas(abrir){
  const mas = $('f-marcas-mas'), btn = $('f-marcas-btn');
  if(!mas) return;
  const ver = abrir === true ? true : mas.hidden;
  mas.hidden = !ver;
  btn.textContent = ver ? 'Ver menos marcas' : btn.dataset.txt;
  btn.setAttribute('aria-expanded', ver);
}
function abrirFiltros(){ ultimoFocus = document.activeElement; $('filtros').classList.add('open'); $('f-ov').hidden = false; document.body.classList.add('lock'); setTimeout(() => $('f-close').focus(), 50); }
function cerrarFiltros(){
  const f = $('filtros');
  if(!f || !f.classList.contains('open')) return;
  f.classList.remove('open'); $('f-ov').hidden = true; document.body.classList.remove('lock');
  if(ultimoFocus) ultimoFocus.focus();
}

// ---------- Grilla ----------
function renderGridPage(){
  renderGrid(visiblesActuales.slice(0, pagMostrados));
  const wrap = $('cargar-mas-wrap');
  wrap.hidden = visiblesActuales.length <= pagMostrados;
  if(!wrap.hidden) $('cargar-mas-info').textContent = `Mostrando ${pagMostrados} de ${visiblesActuales.length}`;
}
function cargarMas(){ pagMostrados += PAG_SIZE; renderGridPage(); }
function renderGrid(lista){
  $('grid').innerHTML = lista.map(cardHTML).join('');
}
function cardHTML(r){
  const p = r.p, cod = jsArg(p.cod);
  const marca = marcaMaquina(p);
  const img = p.img
    ? `<img src="imagenes/${esc(p.img)}" alt="${esc(tipoProd(p)+' '+p.cod)}" loading="lazy">`
    : (p.linea === 'RMH' ? SVG_RMH : SVG_MC);
  const dims = p.esp != null
    ? `${fmtMm(p.esp)} × ${fmtMm(p.anch)}${p.larg != null ? ' × '+fmtMm(p.larg) : ''} mm`
    : '';
  let tagMed = '';
  if(r.dist != null) tagMed = r.dist < 0.05 ? '<span class="tag tag-ok">Medida exacta</span>' : '<span class="tag">Medida aproximada</span>';
  const enCart = cart.find(c => c.cod === p.cod);
  return `<article class="card">
    <div class="card-img">${img}</div>
    <div class="card-body">
      <div class="card-tipo">${tipoProd(p)}</div>
      ${marca ? `<div class="card-marca">${esc(marca)}</div>` : ''}
      <div class="card-desc">${esc(p.desc)}</div>
      ${dims ? `<div class="card-dims"><i class="fa-solid fa-ruler" aria-hidden="true"></i>${dims}${p.terminal != null ? ' · Terminal '+p.terminal : ''}</div>` : ''}
      ${tagMed}
      <div class="card-cod">Cód. ${esc(p.cod)}</div>
      <div class="card-precio">${fmt(p.precio)}</div>
      <div class="card-buy">
        <div class="qty"><button type="button" aria-label="Restar uno" onclick="cambiarQtyCard(${cod},-1)">−</button><input type="number" inputmode="numeric" min="1" value="1" id="${qid(p.cod)}" aria-label="Cantidad de ${esc(p.cod)}"><button type="button" aria-label="Sumar uno" onclick="cambiarQtyCard(${cod},1)">+</button></div>
        <button type="button" class="btn btn-primary" onclick="agregarAlCarrito(${cod})">Agregar</button>
      </div>
      ${enCart ? `<div class="card-encart"><i class="fa-solid fa-check" aria-hidden="true"></i> Tenés ${enCart.qty} en el pedido</div>` : ''}
    </div>
  </article>`;
}
function cambiarQtyCard(cod, delta){
  const inp = $(qid(cod));
  if(!inp) return;
  inp.value = Math.max(1, (parseInt(inp.value) || 1) + delta);
}

// ---------- Pedido ----------
function leerCart(){
  try{ const c = JSON.parse(localStorage.getItem('moncar_cart') || '[]'); return Array.isArray(c) ? c : []; }
  catch(e){ return []; }
}
function precioActual(c){ const p = PRODUCTOS.find(x => x.cod === c.cod); return p ? p.precio : c.precio; }
function guardarCart(){
  try{ localStorage.setItem('moncar_cart', JSON.stringify(cart)); }catch(e){}
  actualizarContador();
  renderCart();
}
function actualizarContador(){
  const n = cart.reduce((s, c) => s + c.qty, 0);
  const tot = cart.reduce((s, c) => s + precioActual(c) * c.qty, 0);
  document.querySelectorAll('.cart-n').forEach(el => el.textContent = n);
  $('pbar-n').textContent = n + (n === 1 ? ' unidad' : ' unidades');
  $('pbar-total').textContent = fmt(tot);
  document.body.classList.toggle('con-pedido', cart.length > 0);
}
function agregarAlCarrito(cod){
  const p = PRODUCTOS.find(x => x.cod === cod); if(!p) return;
  const inp = $(qid(cod));
  const cant = inp ? Math.max(1, parseInt(inp.value) || 1) : 1;
  const ex = cart.find(c => c.cod === cod);
  if(ex) ex.qty += cant; else cart.push({cod: p.cod, desc: p.desc, precio: p.precio, qty: cant});
  guardarCart();
  renderGridPage();
  mostrarToast(`Agregaste ${cant} × ${p.cod} al pedido`);
}
function cambiarQty(cod, delta){
  const ex = cart.find(c => c.cod === cod); if(!ex) return;
  ex.qty += delta;
  if(ex.qty <= 0) cart = cart.filter(c => c.cod !== cod);
  guardarCart(); renderGridPage();
}
function quitarItem(cod){ cart = cart.filter(c => c.cod !== cod); guardarCart(); renderGridPage(); }
function vaciarCart(){ if(confirm('¿Vaciar el pedido?')){ cart = []; guardarCart(); renderGridPage(); } }
function renderCart(){
  const el = $('cart-items');
  $('dr-foot').hidden = !cart.length;
  if(!cart.length){
    el.innerHTML = '<div class="dr-vacio"><p>Todavía no agregaste productos.</p><button type="button" class="btn btn-line" onclick="cerrarCart();irAProductos()">Ver productos</button></div>';
    return;
  }
  el.innerHTML = cart.map(c => {
    const p = PRODUCTOS.find(x => x.cod === c.cod);
    const cod = jsArg(c.cod);
    const tit = p ? (marcaMaquina(p) || tipoProd(p)) : c.cod;
    const img = p && p.img ? `<img src="imagenes/${esc(p.img)}" alt="">` : '';
    return `<div class="ci">
      <div class="ci-img">${img}</div>
      <div>
        <div class="ci-t">${esc(tit)} <span class="ci-cod">· ${esc(c.cod)}</span></div>
        <div class="ci-d">${esc(c.desc)}</div>
        <div class="ci-p">${fmt(precioActual(c))} c/u</div>
        <div class="qty"><button type="button" aria-label="Restar uno" onclick="cambiarQty(${cod},-1)">−</button><span aria-label="Cantidad">${c.qty}</span><button type="button" aria-label="Sumar uno" onclick="cambiarQty(${cod},1)">+</button></div>
      </div>
      <button type="button" class="btn-rm" aria-label="Quitar ${esc(c.cod)}" onclick="quitarItem(${cod})"><i class="fa-solid fa-trash-can" aria-hidden="true"></i></button>
    </div>`;
  }).join('');
  $('cart-total').textContent = fmt(cart.reduce((s, c) => s + precioActual(c) * c.qty, 0));
}
function abrirCart(){
  ultimoFocus = document.activeElement;
  $('toast').classList.remove('show');
  renderCart();
  $('cart-ov').hidden = false;
  $('cart').classList.add('open');
  document.body.classList.add('lock');
  setTimeout(() => $('cart-close').focus(), 50);
}
function cerrarCart(){
  if(!$('cart').classList.contains('open')) return;
  $('cart').classList.remove('open');
  $('cart-ov').hidden = true;
  document.body.classList.remove('lock');
  if(ultimoFocus) ultimoFocus.focus();
}
function mostrarToast(msg){
  $('toast-msg').textContent = msg;
  $('toast').classList.add('show');
  clearTimeout(tToast);
  tToast = setTimeout(() => $('toast').classList.remove('show'), 3500);
}

// ---------- Formulario de pedido / consulta ----------
// Toda opcion de WhatsApp del sitio (nav, portada, carrito) pasa SIEMPRE por
// este formulario, para poder identificar a que WhatsApp del equipo
// redirigir segun la zona del cliente (elegirDestino, arriba).
let _pedidoTipo = '', _esPedido = false, focusModal = null;
function pedirPorWA(){ if(!cart.length) return; _pedidoTipo = 'wa'; abrirFormModal(true); }
function pedirPorEmail(){ if(!cart.length) return; _pedidoTipo = 'email'; abrirFormModal(true); }
function abrirConsultaWA(){ _pedidoTipo = 'wa'; abrirFormModal(false); }
function abrirFormModal(esPedido){
  _esPedido = esPedido;
  focusModal = document.activeElement;
  $('fm-title').textContent = esPedido ? 'Último paso para enviar tu pedido' : 'Antes de escribirnos';
  $('fm-btn').innerHTML = _pedidoTipo === 'wa'
    ? '<i class="fa-brands fa-whatsapp" aria-hidden="true"></i> Abrir WhatsApp'
    : '<i class="fa-solid fa-envelope" aria-hidden="true"></i> Abrir email';
  $('fm-btn').className = 'btn ' + (_pedidoTipo === 'wa' ? 'btn-wa' : 'btn-primary');
  ['f-nombre','f-localidad'].forEach(id => { $(id).classList.remove('err'); $(id).removeAttribute('aria-invalid'); $(id+'-err').hidden = true; });
  $('fmodal').hidden = false;
  document.body.classList.add('lock');
  setTimeout(() => $('f-nombre').focus(), 50);
}
function cerrarFormModal(){
  if($('fmodal').hidden) return;
  $('fmodal').hidden = true;
  if(!$('cart').classList.contains('open')) document.body.classList.remove('lock');
  if(focusModal) focusModal.focus();
}
function confirmarPedido(ev){
  if(ev) ev.preventDefault();
  const nombre = $('f-nombre').value.trim();
  const negocio = $('f-negocio').value.trim();
  const localidad = $('f-localidad').value.trim();
  let primero = null;
  [['f-nombre', nombre], ['f-localidad', localidad]].forEach(([id, v]) => {
    $(id).classList.toggle('err', !v);
    if(!v){ $(id).setAttribute('aria-invalid','true'); } else { $(id).removeAttribute('aria-invalid'); }
    $(id+'-err').hidden = !!v;
    if(!v && !primero) primero = $(id);
  });
  if(primero){ primero.focus(); return; }

  const conItems = _esPedido && cart.length;
  let msg = conItems ? 'Hola! Nuevo pedido desde la web:\n\n' : 'Hola! Quiero hacer una consulta desde la web:\n\n';
  msg += 'Cliente: '+nombre+'\n';
  if(negocio) msg += 'Negocio: '+negocio+'\n';
  msg += 'Localidad: '+localidad+'\n';
  if(conItems){
    msg += '\nProductos:\n';
    cart.forEach(c => { msg += `- ${c.cod} x${c.qty}\n`; });
  }
  msg += '\nMuchas gracias!';

  cerrarFormModal();
  if(_pedidoTipo === 'wa'){
    window.open('https://wa.me/'+elegirDestino(localidad)+'?text='+encodeURIComponent(msg), '_blank');
  } else {
    const subj = 'Nuevo pedido MONCAR RMH' + (negocio ? ' - '+negocio : '');
    window.location.href = 'mailto:'+CONTACTO_EMAIL+'?subject='+encodeURIComponent(subj)+'&body='+encodeURIComponent(msg);
  }
}

document.addEventListener('keydown', e => {
  if(e.key !== 'Escape') return;
  if(!$('fmodal').hidden) cerrarFormModal();
  else if($('cart').classList.contains('open')) cerrarCart();
  else cerrarFiltros();
});

actualizarContador();
renderCart();
if(location.hash) aplicarHash(); else { mostrarPagina('inicio'); filtrar(); }
'''


def generar_html(productos):
    cat_count = Counter(p['cat'] for p in productos)
    cats_orden = [c for c, _ in sorted(cat_count.items(), key=lambda x: -x[1])]

    marca_count = Counter(
        p['marca'] for p in productos
        if p['linea'] == 'MONCAR' and p['marca'] not in EXCLUIR_MARCAS
    )
    marcas_destacadas = [(m, cnt) for m, cnt in marca_count.items() if m in MARCAS_CONFIG]
    marcas_destacadas.sort(key=lambda x: -x[1])

    terms = sorted(set(p['terminal'] for p in productos if p['terminal'] is not None))

    def e(t):
        return _html.escape(str(t), quote=True)

    def js_arg(t):
        # argumento JS seguro dentro de un atributo onclick="..."
        return e(json.dumps(t, ensure_ascii=False))

    def cat_lbl(c):
        return CAT_LABEL.get(c, c)

    def marca_lbl(m):
        return MARCA_LABEL.get(m, m)

    def cat_cards():
        cards = []
        for cat in cats_orden:
            cfg = CAT_CONFIG.get(cat, CAT_CONFIG['Otros'])
            cards.append(
                f'<button type="button" class="cat-card" onclick="irACat({js_arg(cat)})">'
                f'<span class="cat-icon" aria-hidden="true">{cfg["icon"]}</span>'
                f'<span class="cat-name">{e(cat_lbl(cat))}</span>'
                f'<span class="cat-count">{cat_count[cat]} productos</span>'
                f'</button>'
            )
        return '\n'.join(cards)

    def marca_cards():
        cards = []
        for m, cnt in marcas_destacadas:
            cards.append(
                f'<button type="button" class="marca-card" style="--mc:{MARCAS_CONFIG[m]["bg"]}" '
                f'onclick="irAMarca({js_arg(m)})">'
                f'<span class="marca-name">{e(marca_lbl(m))}</span>'
                f'<span class="marca-sub">{cnt} carbones</span>'
                f'</button>'
            )
        return '\n'.join(cards)

    def check(valor, texto, cnt):
        return (f'<label class="chk"><input type="checkbox" value="{e(valor)}" onchange="filtrar()">'
                f'<span>{e(texto)}</span><span class="n">{cnt}</span></label>')

    def filtro_cats():
        return '\n'.join(check(c, cat_lbl(c), cat_count[c]) for c in cats_orden)

    def filtro_marcas():
        orden = sorted(marca_count.items(), key=lambda x: -x[1])
        primeras = '\n'.join(check(m, marca_lbl(m), c) for m, c in orden[:10])
        resto = orden[10:]
        if not resto:
            return f'<div id="f-marcas">{primeras}</div>'
        mas = '\n'.join(check(m, marca_lbl(m), c) for m, c in resto)
        txt = f'Ver todas las marcas ({len(orden)})'
        return (f'<div id="f-marcas">{primeras}<div id="f-marcas-mas" hidden>{mas}</div></div>'
                f'<button type="button" class="f-more" id="f-marcas-btn" data-txt="{e(txt)}" '
                f'aria-expanded="false" onclick="mostrarTodasMarcas()">{e(txt)}</button>')

    def opt(vals):
        return '\n'.join(f'<option value="{v}">{v}</option>' for v in vals)

    # '</' escapado para que ningún texto de producto pueda cerrar el <script>
    productos_json = json.dumps(productos, ensure_ascii=False).replace('</', '<\\/')
    js = (_JS.replace('__PRODUCTOS__', productos_json)
             .replace('__CAT_LABEL__', json.dumps(CAT_LABEL, ensure_ascii=False))
             .replace('__MARCA_LABEL__', json.dumps(MARCA_LABEL, ensure_ascii=False)))

    n_prod = len(productos)
    n_marcas = len(marca_count)
    wa_ico = '<i class="fa-brands fa-whatsapp" aria-hidden="true"></i>'

    return f'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MONCAR RMH — Carbones e interruptores para herramientas</title>
<meta name="description" content="Carbones eléctricos MONCAR e interruptores RMH para herramientas y electrodomésticos. Buscá por marca, tipo de máquina o medidas. Envíos a todo el país.">
<meta name="theme-color" content="#0f0e12">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<style>{_CSS}</style>
</head>
<body>

<header class="top">
  <div class="wrap top-in">
    <button type="button" class="logo" onclick="irAInicio()" aria-label="MONCAR RMH, ir al inicio">MONCAR <span>RMH</span></button>
    <nav class="nav" aria-label="Principal">
      <button type="button" class="nav-btn active" data-nav="inicio" onclick="irAInicio()">Inicio</button>
      <button type="button" class="nav-btn" data-nav="productos" onclick="irAProductos()">Productos</button>
      <button type="button" class="nav-btn" data-nav="nosotros" onclick="irANosotros()">Nosotros</button>
    </nav>
    <div class="top-acts">
      <button type="button" class="btn btn-line" onclick="abrirConsultaWA()" aria-label="Consultar por WhatsApp"><i class="fa-brands fa-whatsapp wa-ico" aria-hidden="true"></i><span class="hide-sm">WhatsApp</span></button>
      <button type="button" class="btn btn-primary" onclick="abrirCart()" aria-label="Ver mi pedido"><i class="fa-solid fa-cart-shopping" aria-hidden="true"></i><span class="hide-sm">Mi pedido</span><span class="cart-n">0</span></button>
    </div>
  </div>
  <nav class="tabs" aria-label="Secciones">
    <button type="button" class="nav-btn active" data-nav="inicio" onclick="irAInicio()">Inicio</button>
    <button type="button" class="nav-btn" data-nav="productos" onclick="irAProductos()">Productos</button>
    <button type="button" class="nav-btn" data-nav="nosotros" onclick="irANosotros()">Nosotros</button>
  </nav>
</header>

<main>
<!-- ===== INICIO ===== -->
<div id="page-inicio">
  <section class="hero">
    <div class="wrap hero-in">
      <div>
        <p class="hero-kicker">Carbones MONCAR · Interruptores RMH</p>
        <h1>Repuestos eléctricos para <em>herramientas</em></h1>
        <p class="hero-p">Carbones e interruptores para las principales marcas de herramientas y electrodomésticos. Para ferreterías, casas de repuestos y talleres. Envíos a todo el país.</p>
        <div class="hero-btns">
          <button type="button" class="btn btn-primary btn-lg" onclick="irAProductos()">Ver catálogo completo</button>
          <button type="button" class="btn btn-ghost-dark btn-lg" onclick="abrirConsultaWA()">{wa_ico} Consultar por WhatsApp</button>
        </div>
      </div>
      <div class="paths">
        <p class="paths-t">¿Cómo querés buscar tu repuesto?</p>
        <button type="button" class="path" onclick="irAInicio('sec-marcas')"><span class="path-ico"><i class="fa-solid fa-tag" aria-hidden="true"></i></span><span class="path-tx"><strong>Por marca de la máquina</strong><span>Bosch, Makita, DeWalt, Black &amp; Decker…</span></span><i class="fa-solid fa-chevron-right path-arrow" aria-hidden="true"></i></button>
        <button type="button" class="path" onclick="irAInicio('sec-cats')"><span class="path-ico"><i class="fa-solid fa-screwdriver-wrench" aria-hidden="true"></i></span><span class="path-tx"><strong>Por tipo de máquina</strong><span>Amoladora, taladro, sierra, lavarropas…</span></span><i class="fa-solid fa-chevron-right path-arrow" aria-hidden="true"></i></button>
        <button type="button" class="path" onclick="irAInicio('sec-medidas')"><span class="path-ico"><i class="fa-solid fa-ruler-combined" aria-hidden="true"></i></span><span class="path-tx"><strong>Por medidas del carbón</strong><span>Espesor × ancho × largo, en milímetros</span></span><i class="fa-solid fa-chevron-right path-arrow" aria-hidden="true"></i></button>
        <form class="hero-search" role="search" onsubmit="buscarHome(event)">
          <label for="h-q">¿Sabés el modelo de la máquina?</label>
          <div class="hs-row"><input id="h-q" class="inp" type="search" placeholder="Ej: amoladora Bosch GWS 7" autocomplete="off"><button type="submit" class="btn btn-primary">Buscar</button></div>
        </form>
      </div>
    </div>
  </section>

  <section class="stats" aria-label="Por qué elegirnos">
    <div class="wrap stats-in">
      <div class="stat"><i class="fa-solid fa-boxes-stacked" aria-hidden="true"></i><div><strong>{n_prod} productos</strong><span>Carbones e interruptores</span></div></div>
      <div class="stat"><i class="fa-solid fa-tags" aria-hidden="true"></i><div><strong>{n_marcas} marcas</strong><span>Herramientas y electrodomésticos</span></div></div>
      <div class="stat"><i class="fa-solid fa-truck-fast" aria-hidden="true"></i><div><strong>Envíos a todo el país</strong><span>Despachamos a cualquier provincia</span></div></div>
      <div class="stat"><i class="fa-solid fa-headset" aria-hidden="true"></i><div><strong>Atención personalizada</strong><span>Te ayudamos a encontrar el repuesto</span></div></div>
    </div>
  </section>

  <section class="sec" id="sec-marcas">
    <div class="wrap">
      <div class="sec-head">
        <div><p class="sec-num">01 — Por marca</p><h2>Buscá por marca de la máquina</h2><p class="sec-sub">Elegí la marca y te mostramos todos los carbones compatibles.</p></div>
        <button type="button" class="link-more" onclick="irAProductos()">Ver todas las marcas</button>
      </div>
      <div class="ruled marca-grid">{marca_cards()}</div>
    </div>
  </section>

  <section class="sec" id="sec-cats">
    <div class="wrap">
      <div class="sec-head">
        <div><p class="sec-num">02 — Por máquina</p><h2>Buscá por tipo de máquina</h2><p class="sec-sub">Carbones e interruptores separados por el tipo de herramienta o electrodoméstico.</p></div>
      </div>
      <div class="ruled cat-grid">{cat_cards()}</div>
    </div>
  </section>

  <section class="sec" id="sec-medidas">
    <div class="wrap">
      <div class="sec-head">
        <div><p class="sec-num">03 — Por medidas</p><h2>Buscá por medidas del carbón</h2><p class="sec-sub">¿Tenés el carbón viejo en la mano? Medilo y encontrá el equivalente.</p></div>
      </div>
      <div class="med-box">
        <div class="med-fig">{_SVG_MEDIDAS}</div>
        <form class="med-form" onsubmit="buscarMedidaHome(event)" novalidate>
          <div class="med-fields">
            <div class="field"><label for="h-esp">Espesor</label><div class="inp-unit"><input id="h-esp" class="inp" inputmode="decimal" placeholder="6,3" autocomplete="off"></div></div>
            <div class="field"><label for="h-anch">Ancho</label><div class="inp-unit"><input id="h-anch" class="inp" inputmode="decimal" placeholder="10" autocomplete="off"></div></div>
            <div class="field"><label for="h-larg">Largo <small>(opcional)</small></label><div class="inp-unit"><input id="h-larg" class="inp" inputmode="decimal" placeholder="17" autocomplete="off"></div></div>
          </div>
          <p class="med-help">Medí en milímetros, con calibre si tenés. Te mostramos los que coinciden y los que se acercan. Si el carbón está gastado, dejá el largo vacío.</p>
          <p class="fg-err" id="h-med-err" role="alert" hidden>Completá al menos el espesor o el ancho.</p>
          <button type="submit" class="btn btn-primary btn-lg">Buscar por medidas</button>
        </form>
      </div>
    </div>
  </section>

  <section class="ayuda">
    <div class="wrap ayuda-in">
      <div><h2>¿No sabés cuál es el repuesto?</h2><p>Mandanos una foto del carbón o de la chapa de la máquina por WhatsApp y te decimos cuál te sirve.</p></div>
      <button type="button" class="btn btn-wa btn-lg" onclick="abrirConsultaWA()">{wa_ico} Consultar por WhatsApp</button>
    </div>
  </section>
</div>

<!-- ===== PRODUCTOS ===== -->
<div id="page-productos" hidden>
  <div class="prod-head">
    <div class="wrap prod-head-in">
      <h1>Catálogo de productos</h1>
      <p>Filtrá por tipo de máquina, marca o medidas. Precios de lista.</p>
      <div class="search" role="search">
        <i class="fa-solid fa-magnifying-glass" aria-hidden="true"></i>
        <label for="search" class="sr-only">Buscar productos</label>
        <input id="search" class="inp" type="search" placeholder="Buscá por máquina, marca o modelo (ej: taladro Bosch)" oninput="filtrarPronto()" autocomplete="off">
        <button type="button" class="search-clear" id="search-clear" onclick="limpiarBusqueda()" aria-label="Borrar búsqueda" hidden><i class="fa-solid fa-xmark" aria-hidden="true"></i></button>
      </div>
    </div>
  </div>

  <div class="wrap prod-layout">
    <div class="f-ov" id="f-ov" onclick="cerrarFiltros()" hidden></div>
    <aside class="filtros" id="filtros" aria-label="Filtros">
      <div class="f-head"><strong>Filtros</strong><button type="button" class="x" id="f-close" onclick="cerrarFiltros()" aria-label="Cerrar filtros"><i class="fa-solid fa-xmark" aria-hidden="true"></i></button></div>
      <div class="f-body">
        <div class="f-sec">
          <h2 class="f-tit"><i class="fa-solid fa-screwdriver-wrench" aria-hidden="true"></i> Tipo de máquina</h2>
          <div id="f-cats">{filtro_cats()}</div>
        </div>
        <div class="f-sec">
          <h2 class="f-tit"><i class="fa-solid fa-tag" aria-hidden="true"></i> Marca</h2>
          {filtro_marcas()}
        </div>
        <div class="f-sec">
          <h2 class="f-tit"><i class="fa-solid fa-ruler-combined" aria-hidden="true"></i> Medidas (mm)</h2>
          <div class="f-med">
            <div><label for="m-esp">Espesor</label><input id="m-esp" class="inp" inputmode="decimal" autocomplete="off" oninput="filtrarPronto()"></div>
            <div><label for="m-anch">Ancho</label><input id="m-anch" class="inp" inputmode="decimal" autocomplete="off" oninput="filtrarPronto()"></div>
            <div><label for="m-larg">Largo</label><input id="m-larg" class="inp" inputmode="decimal" autocomplete="off" oninput="filtrarPronto()"></div>
          </div>
          <p class="f-help">Sirven medidas aproximadas. El largo es opcional.</p>
          <label for="fil-term" class="f-lbl">Terminal</label>
          <select id="fil-term" class="inp" onchange="filtrar()"><option value="">Todos</option>{opt(terms)}</select>
        </div>
        <div class="f-sec">
          <h2 class="f-tit"><i class="fa-solid fa-box" aria-hidden="true"></i> Producto</h2>
          <div class="seg" id="f-linea">
            <button type="button" class="active" data-val="" aria-pressed="true" onclick="setLinea(this)">Todos</button>
            <button type="button" data-val="MONCAR" aria-pressed="false" onclick="setLinea(this)">Carbones (MONCAR)</button>
            <button type="button" data-val="RMH" aria-pressed="false" onclick="setLinea(this)">Interruptores (RMH)</button>
          </div>
        </div>
        <div class="f-sec">
          <h2 class="f-tit"><i class="fa-solid fa-dollar-sign" aria-hidden="true"></i> Precio</h2>
          <div class="f-range">
            <input id="precio-min" class="inp" inputmode="numeric" placeholder="Mínimo" aria-label="Precio mínimo" oninput="filtrarPronto()">
            <span>a</span>
            <input id="precio-max" class="inp" inputmode="numeric" placeholder="Máximo" aria-label="Precio máximo" oninput="filtrarPronto()">
          </div>
        </div>
        <div class="f-sec f-clear-d"><button type="button" class="btn btn-line btn-block" onclick="limpiarFiltros()">Limpiar filtros</button></div>
      </div>
      <div class="f-foot">
        <button type="button" class="btn btn-line" onclick="limpiarFiltros()">Limpiar</button>
        <button type="button" class="btn btn-primary" id="btn-ver-res" onclick="cerrarFiltros()">Ver productos</button>
      </div>
    </aside>

    <div class="prod-main">
      <div class="toolbar">
        <button type="button" class="btn btn-line btn-filtros" onclick="abrirFiltros()"><i class="fa-solid fa-sliders" aria-hidden="true"></i> Filtros <span class="f-count" id="f-count" hidden></span></button>
        <p class="res-n" id="results-info" aria-live="polite"></p>
        <label class="orden"><span>Ordenar:</span>
          <select id="orden" class="inp" onchange="filtrar()" aria-label="Ordenar productos">
            <option value="rec">Recomendado</option>
            <option value="precio-asc">Menor precio</option>
            <option value="precio-desc">Mayor precio</option>
          </select>
        </label>
      </div>
      <div class="chips" id="chips" hidden></div>
      <div class="grid" id="grid"></div>
      <div class="vacio" id="no-results" hidden></div>
      <div class="mas" id="cargar-mas-wrap" hidden>
        <p id="cargar-mas-info"></p>
        <button type="button" class="btn btn-line" onclick="cargarMas()">Mostrar más productos</button>
      </div>
    </div>
  </div>
</div>

<!-- ===== NOSOTROS ===== -->
<div id="page-nosotros" hidden>
  <section class="sec">
    <div class="wrap nos">
      <div>
        <p class="sec-num">Nosotros</p>
        <h1>Repuestos eléctricos para herramientas, desde una empresa familiar</h1>
        <p class="nos-lead">Carbones eléctricos MONCAR e interruptores RMH, compatibles con las principales marcas de herramientas y electrodomésticos del mercado.</p>
        <div class="ruled nos-list">
          <div class="nos-item"><i class="fa-solid fa-bolt" aria-hidden="true"></i><h3>Experiencia en el rubro</h3><p>Años de trayectoria en repuestos eléctricos para herramientas profesionales y de uso doméstico.</p></div>
          <div class="nos-item"><i class="fa-solid fa-tags" aria-hidden="true"></i><h3>Multimarca</h3><p>Bosch, DeWalt, Makita, Black &amp; Decker, Metabo, Hitachi, Stanley, Gamma, Milwaukee y muchas más.</p></div>
          <div class="nos-item"><i class="fa-solid fa-truck-fast" aria-hidden="true"></i><h3>Envíos a todo el país</h3><p>Despachamos a cualquier punto del país. Consultá por tu zona y los tiempos de entrega.</p></div>
          <div class="nos-item"><i class="fa-solid fa-headset" aria-hidden="true"></i><h3>Atención personalizada</h3><p>Te ayudamos a encontrar el repuesto exacto para tu herramienta. Consultá sin compromiso.</p></div>
        </div>
      </div>
      <div class="contacto">
        <h2>Contacto</h2>
        <a class="c-row" href="https://wa.me/5491136267009" target="_blank" rel="noopener"><span class="c-ico">{wa_ico}</span><span><span class="c-lbl">WhatsApp</span><span class="c-val">+54 9 11 3626-7009</span></span></a>
        <a class="c-row" href="mailto:contacto.carbonkey@gmail.com"><span class="c-ico"><i class="fa-solid fa-envelope" aria-hidden="true"></i></span><span><span class="c-lbl">Email</span><span class="c-val">contacto.carbonkey@gmail.com</span></span></a>
        <button type="button" class="c-row" onclick="irAProductos()"><span class="c-ico"><i class="fa-solid fa-boxes-stacked" aria-hidden="true"></i></span><span><span class="c-lbl">Catálogo</span><span class="c-val">{n_prod} productos disponibles</span></span></button>
      </div>
    </div>
  </section>
</div>
</main>

<footer>
  <div class="wrap f-in">
    <div>
      <div class="f-brand">MONCAR <span>RMH</span></div>
      <p>Carbones eléctricos e interruptores para herramientas. Envíos a todo el país.</p>
    </div>
    <div><h3>Navegación</h3><button type="button" class="f-link" onclick="irAInicio()">Inicio</button><button type="button" class="f-link" onclick="irAProductos()">Productos</button><button type="button" class="f-link" onclick="irANosotros()">Nosotros</button></div>
    <div><h3>Contacto</h3><a class="f-link" href="https://wa.me/5491136267009" target="_blank" rel="noopener">WhatsApp: +54 9 11 3626-7009</a><a class="f-link" href="mailto:contacto.carbonkey@gmail.com">contacto.carbonkey@gmail.com</a></div>
    <div><h3>Legal</h3><a class="f-link" href="privacidad.html">Política de privacidad</a><a class="f-link" href="terminos.html">Términos y condiciones</a></div>
  </div>
  <div class="wrap f-bot">© 2026 MONCAR RMH — Todos los derechos reservados. Ícono de máquina de coser: Caro Asercion, <a href="https://game-icons.net" target="_blank" rel="noopener">game-icons.net</a> (<a href="https://creativecommons.org/licenses/by/3.0/" target="_blank" rel="noopener">CC BY 3.0</a>).</div>
</footer>

<!-- PEDIDO (panel lateral) -->
<div class="ov" id="cart-ov" onclick="cerrarCart()" hidden></div>
<aside class="drawer" id="cart" aria-label="Mi pedido">
  <div class="dr-head"><h2>Mi pedido</h2><button type="button" class="x" id="cart-close" onclick="cerrarCart()" aria-label="Cerrar pedido"><i class="fa-solid fa-xmark" aria-hidden="true"></i></button></div>
  <div class="dr-items" id="cart-items"></div>
  <div class="dr-foot" id="dr-foot" hidden>
    <div class="dr-total"><span>Total estimado</span><strong id="cart-total">$0</strong></div>
    <p class="dr-note">Precios de lista. Te confirmamos el pedido y el envío por WhatsApp o email.</p>
    <button type="button" class="btn btn-wa btn-lg" onclick="pedirPorWA()">{wa_ico} Enviar pedido por WhatsApp</button>
    <button type="button" class="btn btn-line" onclick="pedirPorEmail()"><i class="fa-solid fa-envelope" aria-hidden="true"></i> Enviar por email</button>
    <button type="button" class="dr-vaciar" onclick="vaciarCart()">Vaciar pedido</button>
  </div>
</aside>

<!-- FORMULARIO PEDIDO / CONSULTA -->
<div class="modal" id="fmodal" role="dialog" aria-modal="true" aria-labelledby="fm-title" hidden onclick="if(event.target===this)cerrarFormModal()">
  <form class="modal-box" onsubmit="confirmarPedido(event)" novalidate>
    <h2 id="fm-title">Antes de escribirnos</h2>
    <p class="modal-sub">Atendemos por zonas: con estos datos te derivamos directo con quien te puede atender.</p>
    <div class="fg">
      <label for="f-nombre">Nombre</label>
      <input type="text" id="f-nombre" class="inp" autocomplete="name" aria-describedby="f-nombre-err">
      <p class="fg-err" id="f-nombre-err" hidden>Escribí tu nombre.</p>
    </div>
    <div class="fg">
      <label for="f-negocio">Nombre del negocio <span class="opt">(opcional)</span></label>
      <input type="text" id="f-negocio" class="inp" autocomplete="organization" placeholder="Ferretería, taller, bulonera…">
    </div>
    <div class="fg">
      <label for="f-localidad">Localidad</label>
      <input type="text" id="f-localidad" class="inp" autocomplete="address-level2" placeholder="Ej: Escobar, Mar del Plata…" aria-describedby="f-localidad-err">
      <p class="fg-err" id="f-localidad-err" hidden>Escribí tu localidad, así te deriva quien atiende tu zona.</p>
    </div>
    <p class="aviso">Usamos estos datos solo para responder tu consulta y derivarte con quien atiende tu zona. No los guardamos en este sitio: viajan en el mensaje que vos enviás. <a href="privacidad.html" target="_blank">Política de privacidad</a></p>
    <div class="modal-acts">
      <button type="button" class="btn btn-line" onclick="cerrarFormModal()">Cancelar</button>
      <button type="submit" class="btn btn-wa" id="fm-btn">{wa_ico} Abrir WhatsApp</button>
    </div>
  </form>
</div>

<!-- Aviso "agregado" + barra de pedido en el celular -->
<div class="toast" id="toast" role="status" aria-live="polite"><i class="fa-solid fa-circle-check" aria-hidden="true"></i><span id="toast-msg"></span><button type="button" onclick="abrirCart()">Ver pedido</button></div>
<div class="pbar"><div class="pbar-tx"><strong id="pbar-total">$0</strong><span id="pbar-n">0 unidades</span></div><button type="button" class="btn btn-primary" onclick="abrirCart()"><i class="fa-solid fa-cart-shopping" aria-hidden="true"></i> Ver pedido</button></div>

<script>{js}</script>
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
    generar_paginas_legales(OUT_DIR)
