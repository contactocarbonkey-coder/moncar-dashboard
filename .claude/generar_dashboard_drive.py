# -*- coding: utf-8 -*-
# Version DRIVE: lee los datos desde Google Sheets (CSV publico) en vez de Excel local.
# Copia de generar_dashboard.py — mismo TEMPLATE, distinta fuente de datos.
import sys, os, json, re, webbrowser, csv, io, time
import urllib.request, urllib.parse, urllib.error
from datetime import datetime

# Fix Windows stdout encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

DIR      = r'C:\Users\franc\OneDrive\Documentos\KeyCarbon\claude MONCAR RMH'
DEV_MODE = '--dev' in sys.argv
OUT_PATH = DIR + (r'\dashboard_drive_dev.html' if DEV_MODE else r'\dashboard_drive.html')

# IDs de los Google Sheets nativos (conversión de los .xlsx que estaban en Drive).
# Requieren acceso "Cualquiera con el enlace - Lector" para poder leerse sin login.
VENTAS_SHEET_ID    = '1K8wcDYl5E6hf7B-Go-R-3BPnCISNFKibnEV2C4AmfOE'
STOCK_MC_SHEET_ID  = '1u9xG2uyeT3V8osBi9Qz0xofDCPmFmhOXlu85c1OIEQ8'
STOCK_RMH_SHEET_ID = '1Msl0CC5abCoIaM0MpeCtYOoHVHI0mjPlfElc5x8ytJE'

# ─── helpers ────────────────────────────────────────────────────────────────
def fetch_sheet_csv(sheet_id, sheet_name):
    """Descarga una hoja puntual de un Google Sheet como filas CSV, con reintentos."""
    url = ('https://docs.google.com/spreadsheets/d/%s/gviz/tq?tqx=out:csv&sheet=%s'
           % (sheet_id, urllib.parse.quote(sheet_name)))
    for intento in range(4):
        try:
            with urllib.request.urlopen(url, timeout=20) as resp:
                text = resp.read().decode('utf-8')
            return list(csv.reader(io.StringIO(text)))
        except urllib.error.HTTPError as e:
            if e.code == 400:
                return []  # la hoja no existe con ese nombre
            if intento < 3:
                print(f"  Error de red leyendo '{sheet_name}', reintentando... ({e})")
                time.sleep(2)
            else:
                raise RuntimeError(f"No se pudo descargar la hoja '{sheet_name}': {e}")
        except Exception as e:
            if intento < 3:
                print(f"  Error de red leyendo '{sheet_name}', reintentando... ({e})")
                time.sleep(2)
            else:
                raise RuntimeError(f"No se pudo descargar la hoja '{sheet_name}': {e}")

def s(val):
    return str(val).strip() if val is not None else ""

def norm_tablero(val):
    """N° de tablero de Ventas.TABLERO ("9", "9.0", "009") -> "009". Vacio si no hay."""
    t = s(val)
    if t.endswith('.0'):
        t = t[:-2]
    return t.zfill(3) if t.isdigit() else t.upper()

def n(val):
    """Convierte texto de celda exportado por Sheets (formato AR: punto=miles, coma=decimal) a float."""
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

def parse_fecha(txt):
    """Sheets exporta la fecha como texto dd/mm/yyyy; la pasamos a yyyy-mm-dd."""
    txt = s(txt)
    try:
        return datetime.strptime(txt, "%d/%m/%Y").strftime("%Y-%m-%d")
    except ValueError:
        return txt

# ─── read VENTAS ─────────────────────────────────────────────────────────────
def read_ventas():
    print("Leyendo VENTAS desde Google Sheets...")

    # ── LP (precios) — hoja legacy, puede no existir mas (se separó en LP MONCAR / LP RMH)
    precios = {}
    lp_rows = fetch_sheet_csv(VENTAS_SHEET_ID, 'LP')
    for r in lp_rows[3:]:
        cod = s(r[0]) if len(r) > 0 else ""
        if cod:
            precios[cod] = n(r[4]) if len(r) > 4 else 0.0

    # ── Ventas
    # OJO: gviz a veces devuelve una respuesta "exitosa" pero con menos filas de
    # las que la hoja realmente tiene (visto en vivo: 183 en vez de 804, sin
    # ningun error de red — probablemente por un recalculo/edicion concurrente
    # en el momento del fetch). Como esta es la hoja mas critica de todas, la
    # pedimos dos veces y nos quedamos con la lectura mas grande.
    ventas_rows = fetch_sheet_csv(VENTAS_SHEET_ID, 'Ventas')
    time.sleep(1)
    ventas_rows_2 = fetch_sheet_csv(VENTAS_SHEET_ID, 'Ventas')
    if len(ventas_rows_2) != len(ventas_rows):
        print(f"  AVISO: 'Ventas' devolvió {len(ventas_rows)} filas en el primer pedido "
              f"y {len(ventas_rows_2)} en el segundo — usando la lectura más grande.")
    if len(ventas_rows_2) > len(ventas_rows):
        ventas_rows = ventas_rows_2
    if not ventas_rows:
        raise RuntimeError("No se pudo leer la hoja 'Ventas' del Google Sheet de VENTAS")

    # Mapeo por nombre de columna — robusto ante reordenamientos futuros
    raw_headers = [s(h).replace('\n', ' ').upper() for h in ventas_rows[0]]
    def ci(name):
        for i, h in enumerate(raw_headers):
            if h == name.upper():
                return i
        return None
    iREMITO    = ci('N° REMITO')   or 0
    iVENDEDOR  = ci('VENDEDOR')    or 1
    iFECHA     = ci('FECHA')       or 2
    iLINEA     = ci('LINEA') or ci('LÍNEA') or 3
    iCOD       = ci('COD')         or 4
    iCANTIDAD  = ci('CANTIDAD')    or 5
    iCLIENTE   = ci('CLIENTE')     or 6
    iCOMIS     = ci('COMIS%')      or 7
    iPRECIO    = ci('PRECIO UNIT') or 8
    iSUBTOTAL  = ci('SUBTOTAL')    or 9
    iMES       = ci('AÑO-MES')     or 10
    iCOMENT    = ci('COMENTARIOS') or 11
    iCMV       = ci('CMV')         or 12
    iMARGEN    = ci('MARGEN')      or 13
    iMARCA     = ci('MARCA')       or 14
    iTIPO      = ci('TIPO CLIENTE') or 15
    iZONA      = ci('ZONA')        or 16
    iLOCALIDAD = ci('LOCALIDAD')   or 17
    iPROVINCIA = ci('PROVINCIA')   or 18
    iDESC      = ci('DESCRIPCION') or 19
    # Columnas nuevas (2026-09) — sin posicion de respaldo fija porque no
    # existian antes; si no estan (Sheet viejo) quedan en None y se ignoran.
    iCOBRADO   = ci('COBRADO_EN_MES')
    iMESCOBRO  = ci('MES_COBRO')
    # 2026-10: componentes de un tablero organizador vendido (remitos-app).
    iTABLERO   = ci('TABLERO')

    raw = []
    for row in ventas_rows[1:]:
        if len(row) <= iSUBTOTAL:
            continue
        REMITO    = s(row[iREMITO])    if len(row) > iREMITO    else ""
        VENDEDOR  = s(row[iVENDEDOR])  if len(row) > iVENDEDOR  else ""
        FECHA     = parse_fecha(row[iFECHA]) if len(row) > iFECHA else ""
        LINEA_raw = s(row[iLINEA])     if len(row) > iLINEA     else ""
        COD       = s(row[iCOD])       if len(row) > iCOD       else ""
        CANTIDAD  = n(row[iCANTIDAD])  if len(row) > iCANTIDAD  else 0.0
        CLIENTE   = s(row[iCLIENTE])   if len(row) > iCLIENTE   else ""
        com_raw   = n(row[iCOMIS])     if len(row) > iCOMIS     else 0.0
        COMISION  = com_raw if 0 < com_raw <= 1 else 0.30
        PRECIO    = n(row[iPRECIO])    if len(row) > iPRECIO    else 0.0
        SUBTOTAL  = n(row[iSUBTOTAL])
        MES       = s(row[iMES])       if len(row) > iMES       else ""
        COMENTARIOS = s(row[iCOMENT]) if len(row) > iCOMENT else ""
        CMV       = n(row[iCMV])      if len(row) > iCMV    else 0.0
        MARGEN    = n(row[iMARGEN])   if len(row) > iMARGEN else 0.0
        MARCA     = s(row[iMARCA])    if len(row) > iMARCA  else ""
        TIPO_CLI  = s(row[iTIPO])     if len(row) > iTIPO   else ""
        ZONA      = s(row[iZONA])     if len(row) > iZONA   else ""
        LOCALIDAD = s(row[iLOCALIDAD]) if len(row) > iLOCALIDAD else ""
        PROVINCIA = s(row[iPROVINCIA]) if len(row) > iPROVINCIA else ""
        _desc_raw = s(row[iDESC])     if len(row) > iDESC   else ""
        DESC      = "" if _desc_raw.startswith('DESCRIPCION') else _desc_raw

        # Vista financiera (rendicion): mes en que esa linea realmente se
        # cobra/liquida, separado del mes en que se vendio (MES, arriba).
        # Si no esta tildado "cobrado en el mes" y hay un mes cargado, se usa
        # ese; si no, se asume cobrado el mismo mes de la venta — cubre tanto
        # el caso normal (tildado) como remitos viejos sin este dato.
        _cobrado_txt  = s(row[iCOBRADO])  if iCOBRADO  is not None and len(row) > iCOBRADO  else ""
        _mescobro_txt = s(row[iMESCOBRO]) if iMESCOBRO is not None and len(row) > iMESCOBRO else ""
        MES_COBRO = _mescobro_txt if _cobrado_txt.upper() == 'NO' and re.match(r'^\d{4}-\d{2}$', _mescobro_txt) else MES

        TABLERO = norm_tablero(row[iTABLERO]) if iTABLERO is not None and len(row) > iTABLERO else ""

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
            'PRECIO': PRECIO, 'SUBTOTAL': SUBTOTAL, 'MES': MES, 'MES_COBRO': MES_COBRO,
            'COMENTARIOS': COMENTARIOS, 'CMV': CMV, 'MARGEN': MARGEN,
            'MARCA': MARCA, 'TIPO_CLI': TIPO_CLI, 'ZONA': ZONA,
            'LOCALIDAD': LOCALIDAD, 'PROVINCIA': PROVINCIA, 'DESC': DESC,
            'TABLERO': TABLERO,
        })

    months     = sorted(set(r['MES'] for r in raw))
    vendedores = sorted(set(r['VENDEDOR'] for r in raw if r['VENDEDOR']))
    print(f"  -> {len(raw)} filas válidas, {len(months)} meses")
    return raw, months, vendedores, precios

# ─── ventas manuales sin remito (STOCK_RMH, historico pre-app) ───────────────
def leer_ventas_manuales_sin_remito():
    """
    Ventas de RMH que pasaron ANTES de que existiera remitos-app (mayo a
    julio 2026) y quedaron registradas solo como movimientos de salida en
    STOCK_RMH!Movimientos (comentario "vta mayo"/"vta junio excel maxi"/etc,
    SIN "Remito N°") — nunca se cargaron como fila en Ventas. Pedido
    explicito de Franco (2026-09-07): sumarlas a las estadisticas del
    dashboard ("tambien son ventas") sin tocar el Sheet de Ventas para nada.

    Se arman como filas sinteticas con la MISMA forma que devuelve
    read_ventas(), para que fluyan solas por el resto del dashboard
    (Overview, Ventas Detalle, Productos, Marcas, P&L) sin tocar ni una
    linea mas de codigo — se distinguen por REMITO vacio y por el texto en
    COMENTARIOS ("Venta manual sin remito").

    Estas filas NO tienen precio de venta real guardado (solo cantidad), asi
    que el precio/subtotal se ESTIMA con LP RMH columna "Vigencia Abr-26"
    (la mejor referencia historica disponible para mayo-julio) — es una
    aproximacion, no lo que Maxi realmente cobro en su momento. El CMV se
    calcula igual que en el resto del dashboard (Costos!COSTO PESOS x
    cantidad); si un producto no tiene costo cargado en Costos (ej. RM 029),
    CMV queda en 0 para esa linea, mismo comportamiento que ya tiene la
    formula de Ventas cuando el VLOOKUP no encuentra el producto.

    Filtra cualquier "Salida Venta X" cuyo comentario matchee "Remito N° X"
    — esas ya estan contadas en Ventas via remitos-app, incluirlas de nuevo
    las duplicaria.
    """
    print("Leyendo ventas manuales sin remito (STOCK_RMH, historico Maxi)...")
    mov_rows = fetch_sheet_csv(STOCK_RMH_SHEET_ID, 'Movimientos')
    if not mov_rows or len(mov_rows) < 2:
        return []

    # La fila de headers no esta siempre en la misma posicion (el export CSV
    # de gviz puede saltear filas de instrucciones vacias de forma distinta a
    # como las ve la API de Sheets que usa remitos-app) — se busca por
    # contenido en vez de asumir un indice fijo, mismo criterio que
    # encontrarFilaHeader() en remitos-app/src/lib/sheets.ts.
    def normalizar_headers(fila):
        return [s(h).replace('\n', ' ').upper() for h in fila]

    idx_header = None
    raw_headers = None
    for i, fila in enumerate(mov_rows[:6]):
        norm = normalizar_headers(fila)
        if 'COMENTARIOS' in norm:
            idx_header = i
            raw_headers = norm
            break
    if idx_header is None:
        print("  Aviso: no se encontro la fila de headers en Movimientos, se omite.")
        return []

    def ci(name):
        for i, h in enumerate(raw_headers):
            if h == name.upper():
                return i
        return None
    iFECHA    = ci('FECHA')
    iTIPO     = ci('TIPO MOVIMIENTO')
    iCOD      = ci('COD PRODUCTO')
    iCANTIDAD = ci('CANTIDAD')
    iCOMENT   = ci('COMENTARIOS')
    if None in (iFECHA, iTIPO, iCOD, iCANTIDAD, iCOMENT):
        print("  Aviso: no se encontraron todas las columnas esperadas en Movimientos, se omite.")
        return []

    re_remito = re.compile(r'^Remito N[°º]\s*\d+$', re.IGNORECASE)

    # Precio de referencia (LP RMH) y costo (Costos) por COD. OJO: el export
    # CSV de gviz colapsa/desalinea filas de instrucciones de forma distinta
    # segun la hoja (confirmado empiricamente: en LP RMH junta 3 filas en 1,
    # en Costos deja celdas de header en blanco pero SIN correr filas) — no
    # se puede asumir un indice fijo, se busca por contenido.
    def indice_fila_header(rows, texto_ancla, max_filas=6):
        ancla = texto_ancla.upper()
        for i, fila in enumerate(rows[:max_filas]):
            norm = [s(c).replace('\n', ' ').upper() for c in fila]
            if ancla in norm:
                return i
        return None

    lp_rmh = fetch_sheet_csv(VENTAS_SHEET_ID, 'LP RMH')
    precio_ref, marca_ref, desc_ref = {}, {}, {}
    idx_lp = indice_fila_header(lp_rmh, 'DESCRIPCION')
    for r in (lp_rmh[idx_lp + 1:] if idx_lp is not None else []):
        if len(r) < 6 or not s(r[0]):
            continue
        cod = s(r[0])
        marca_ref[cod] = s(r[2])
        desc_ref[cod]  = s(r[3])
        precio_ref[cod] = n(r[5])
    if idx_lp is None:
        print("  Aviso: no se encontro la fila de headers en LP RMH, sin precios de referencia.")

    costos_rows = fetch_sheet_csv(VENTAS_SHEET_ID, 'Costos')
    costo_pesos = {}
    idx_costos = indice_fila_header(costos_rows, 'COD')
    for r in (costos_rows[idx_costos + 1:] if idx_costos is not None else []):
        if len(r) < 5 or not s(r[0]):
            continue
        costo_pesos[s(r[0])] = n(r[4])

    filas = []
    omitidas_sin_precio = 0
    for row in mov_rows[idx_header + 1:]:
        if len(row) <= max(iFECHA, iTIPO, iCOD, iCANTIDAD):
            continue
        tipo = s(row[iTIPO])
        if not tipo.startswith('Salida Venta'):
            continue
        comentario = s(row[iCOMENT]) if len(row) > iCOMENT else ""
        if re_remito.match(comentario):
            continue  # ya contada en Ventas via remitos-app

        cod = s(row[iCOD])
        cantidad = n(row[iCANTIDAD])
        if not cod or cantidad <= 0:
            continue
        fecha = parse_fecha(row[iFECHA])
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', fecha):
            continue
        mes = fecha[:7]

        precio = precio_ref.get(cod, 0.0)
        if precio <= 0:
            omitidas_sin_precio += 1
            continue  # sin precio de referencia no se puede estimar $ - se omite en vez de mostrar $0 enganoso

        subtotal = precio * cantidad
        cmv = costo_pesos.get(cod, 0.0) * cantidad
        vendedor = tipo.replace('Salida Venta', '').strip()

        filas.append({
            'REMITO': '', 'VENDEDOR': vendedor, 'FECHA': fecha,
            'LINEA': 'RMH', 'COD': cod,
            'CANTIDAD': cantidad, 'CLIENTE': 'Sin dato (venta manual)', 'COMISION': 0.30,
            'PRECIO': precio, 'SUBTOTAL': subtotal, 'MES': mes, 'MES_COBRO': mes,
            'COMENTARIOS': f"Venta manual sin remito ({comentario})" if comentario else "Venta manual sin remito",
            'CMV': cmv, 'MARGEN': subtotal - cmv,
            'MARCA': marca_ref.get(cod, ''), 'TIPO_CLI': 'Sin dato', 'ZONA': 'Sin dato',
            'LOCALIDAD': '', 'PROVINCIA': '', 'DESC': desc_ref.get(cod, ''),
        })

    if omitidas_sin_precio:
        print(f"  Aviso: {omitidas_sin_precio} lineas omitidas por no tener precio de referencia en LP RMH")
    print(f"  -> {len(filas)} ventas manuales sin remito reconstruidas (RMH)")
    return filas

# ─── read STOCK ──────────────────────────────────────────────────────────────
def read_stock_mc():
    print("Leyendo STOCK MONCAR desde Google Sheets...")

    # Valores base de Productos, por si 'Stock' tuviera alguna celda sin calcular
    base = {}
    for row in fetch_sheet_csv(STOCK_MC_SHEET_ID, 'Productos')[2:]:
        if row and s(row[0]):
            cod = s(row[0])
            base[cod] = {
                'NO_EMB':  n(row[3]) if len(row) > 3 else 0.0,
                'EMB_DEP': n(row[4]) if len(row) > 4 else 0.0,
                'PREM_DEP':n(row[5]) if len(row) > 5 else 0.0,
                'CABA':    n(row[6]) if len(row) > 6 else 0.0,
            }

    rows = []
    sin_cache = 0
    for row in fetch_sheet_csv(STOCK_MC_SHEET_ID, 'Stock')[2:]:
        if not row: continue
        COD = s(row[0]) if len(row) > 0 else ""
        if not COD or not (COD.startswith('CC') or COD.startswith('CF') or COD.startswith('CE')):
            continue
        no_emb_txt  = row[2] if len(row) > 2 else ""
        emb_dep_txt = row[3] if len(row) > 3 else ""
        if not s(no_emb_txt) and not s(emb_dep_txt) and COD in base:
            sin_cache += 1
            b = base[COD]
            no_emb_v, emb_dep_v = b['NO_EMB'], b['EMB_DEP']
            prem_dep_v, caba_v  = b['PREM_DEP'], b['CABA']
            vendible_v = emb_dep_v + prem_dep_v + caba_v
            total_v    = no_emb_v + vendible_v
        else:
            no_emb_v   = n(no_emb_txt);  emb_dep_v  = n(emb_dep_txt)
            prem_dep_v = n(row[4]) if len(row) > 4 else 0.0
            caba_v     = n(row[5]) if len(row) > 5 else 0.0
            # OJO col.6 (0-based) es 'STOCK DAVID' desde 2026-09 (la app de
            # remitos la agrego al lado de STOCK CABA) — este dashboard no
            # muestra el stock de David, solo hay que saltearla para no leer
            # VENDIBLE/TOTAL/PREST/DEBO en la posicion vieja.
            vendible_v = n(row[7]) if len(row) > 7 else 0.0
            total_v    = n(row[8]) if len(row) > 8 else 0.0
        rows.append({
            'COD': COD,
            'MARCA':    s(row[1]) if len(row) > 1 else "",
            'NO_EMB':   no_emb_v,
            'EMB_DEP':  emb_dep_v,
            'PREM_DEP': prem_dep_v,
            'CABA':     caba_v,
            'VENDIBLE': vendible_v,
            'TOTAL':    total_v,
            'PREST':    n(row[9]) if len(row) > 9 else 0.0,
            'DEBO':     n(row[10]) if len(row) > 10 else 0.0,
        })
    if sin_cache:
        print(f"  -> AVISO: {sin_cache} productos sin datos en 'Stock' — usando valores base de 'Productos'")
    print(f"  -> {len(rows)} productos MONCAR")
    return rows

def read_stock_rmh():
    print("Leyendo STOCK RMH desde Google Sheets...")
    rows = []
    for row in fetch_sheet_csv(STOCK_RMH_SHEET_ID, 'Stock')[2:]:
        if not row: continue
        COD = s(row[0]) if len(row) > 0 else ""
        if not COD or not COD.startswith('RM'):
            continue
        # OJO col.5 (0-based) es 'STOCK DAVID' desde 2026-09 (la app de
        # remitos la agrego al lado de STOCK CABA) — este dashboard no
        # muestra el stock de David, solo hay que saltearla.
        rows.append({
            'COD':     COD,
            'DESC':    s(row[1]) if len(row) > 1 else "",
            'NO_EMB':  n(row[2]) if len(row) > 2 else 0.0,
            'EMB_DEP': n(row[3]) if len(row) > 3 else 0.0,
            'CABA':    n(row[4]) if len(row) > 4 else 0.0,
            'MAXI':    n(row[6]) if len(row) > 6 else 0.0,
            'VENDIBLE':n(row[7]) if len(row) > 7 else 0.0,
            'TOTAL':   n(row[8]) if len(row) > 8 else 0.0,
            'PREST':   n(row[9]) if len(row) > 9 else 0.0,
            'DEBO':    n(row[10]) if len(row) > 10 else 0.0,
        })
    print(f"  -> {len(rows)} productos RMH")
    return rows

def read_lita():
    """Movimientos "Salida Lita" de STOCK_MONCAR (ventas MONCAR de Maxi, que
    salen del stock de Lita y Franco le repone). Pendiente = todavia no se
    separo/entrego (sigue contando en EMBOLSADO DEP.). Entregado = ya se resto.
    El boton "Entregar" del Overview escribe via entregas_lita.py.
    Filas con CONFIRMACIÓN vacia = cargas manuales viejas, se ignoran."""
    print("Leyendo entregas a Lita (Movimientos MONCAR)...")
    rows = fetch_sheet_csv(STOCK_MC_SHEET_ID, 'Movimientos')
    norm = lambda t: ' '.join(s(t).replace('\n', ' ').split()).upper()
    # Header buscado por contenido (gviz desalinea las filas de instrucciones)
    ih = next((i for i, r in enumerate(rows[:8]) if any('CONFIRMACI' in norm(c) for c in r)), None)
    if ih is None:
        print("  Aviso: Movimientos MONCAR sin columna CONFIRMACIÓN, se omite Lita.")
        return {'pend': [], 'ent': []}
    header = [norm(c) for c in rows[ih]]
    def col(pred):
        return next((i for i, h in enumerate(header) if pred(h)), None)
    c_fecha = col(lambda h: h == 'FECHA')
    c_tipo  = col(lambda h: h.startswith('TIPO'))
    c_est   = col(lambda h: h == 'ESTADO')
    c_cod   = col(lambda h: h.startswith('COD'))
    c_cant  = col(lambda h: h == 'CANTIDAD')
    c_com   = col(lambda h: h.startswith('COMENTARIO'))
    c_am    = col(lambda h: h.startswith('AÑO-MES'))
    c_conf  = col(lambda h: h.startswith('CONFIRMACI'))
    c_fent  = col(lambda h: h.startswith('FECHA ENTREGA'))  # puede no existir todavia
    cel = lambda r, c: s(r[c]) if c is not None and c < len(r) else ''

    pend, ent = [], []
    for r in rows[ih + 1:]:
        if cel(r, c_tipo) != 'Salida Lita':
            continue
        conf = cel(r, c_conf)
        cod  = ' '.join(cel(r, c_cod).split()).upper()
        qty  = n(cel(r, c_cant))
        if not cod or qty <= 0:
            continue
        fecha = parse_fecha(cel(r, c_fecha))
        if conf == 'Pendiente' and cel(r, c_est) == 'Embolsado':
            pend.append({'cod': cod, 'qty': qty, 'fecha': fecha, 'remito': cel(r, c_com)})
        elif conf == 'Entregado':
            fent = parse_fecha(cel(r, c_fent))
            con_fecha = bool(re.match(r'\d{4}-\d{2}-\d{2}$', fent))
            # Entregas previas a la columna FECHA ENTREGA: se agrupan por mes de venta
            mes = fent[:7] if con_fecha else (cel(r, c_am)[:7] or fecha[:7])
            ent.append({'cod': cod, 'qty': qty, 'fent': fent if con_fecha else '',
                        'mes': mes, 'estado': cel(r, c_est)})
    print(f"  -> {len(pend)} filas pendientes, {len(ent)} entregadas")
    return {'pend': pend, 'ent': ent}

# ─── main ────────────────────────────────────────────────────────────────────
def main():
    raw, months, vendedores, precios = read_ventas()

    try:
        raw_manual = leer_ventas_manuales_sin_remito()
        if raw_manual:
            raw = raw + raw_manual
            months = sorted(set(months) | set(r['MES'] for r in raw_manual))
            vendedores = sorted(set(vendedores) | set(r['VENDEDOR'] for r in raw_manual if r['VENDEDOR']))
    except Exception as e:
        print(f"  ADVERTENCIA ventas manuales sin remito: {e}")

    stock_mc_rows = []
    try:
        stock_mc_rows = read_stock_mc()
    except Exception as e:
        print(f"  ADVERTENCIA stock MONCAR: {e}")

    stock_rmh_rows = []
    try:
        stock_rmh_rows = read_stock_rmh()
    except Exception as e:
        print(f"  ADVERTENCIA stock RMH: {e}")

    lita = {'pend': [], 'ent': []}
    try:
        lita = read_lita()
    except Exception as e:
        print(f"  ADVERTENCIA entregas Lita: {e}")

    data = {
        'ventas': {
            'raw': raw,
            'months': months,
            'vendedores': vendedores,
        },
        'stock_mc':  {'rows': stock_mc_rows},
        'stock_rmh': {'rows': stock_rmh_rows},
        'lita':      lita,
        'precios':   precios,
        'generated_at': datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

    json_str = json.dumps(data, ensure_ascii=False)
    html = TEMPLATE.replace('/*INJECT_DATA*/', f'window.DATA = {json_str};')

    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    modo = "[DEV] " if DEV_MODE else ""
    print(f"✓ {modo}Dashboard generado (desde Google Drive): {OUT_PATH}")
    if '--no-browser' not in sys.argv:
        print(f"✓ Abriendo en el navegador...")
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
body.light .topbar,body.light .sidebar{box-shadow:0 1px 4px rgba(0,0,0,.06)}
body.light select{color:var(--text)}
body{background:var(--bg);color:var(--text);font-family:'Space Grotesk',sans-serif;min-height:100vh}
/* LAYOUT: menu lateral + contenido */
.app{display:flex;min-height:100vh}
.main{flex:1;min-width:0}
/* SIDEBAR (plegable: body.sb-collapsed) */
.sidebar{width:232px;flex-shrink:0;position:sticky;top:0;height:100vh;z-index:110;
  background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column;transition:width .18s}
.sb-head{display:flex;align-items:center;justify-content:space-between;height:56px;padding:0 10px 0 18px;
  border-bottom:1px solid var(--border);flex-shrink:0}
.logo{font-weight:700;font-size:18px;color:var(--accent);white-space:nowrap;letter-spacing:-0.5px;overflow:hidden}
.logo span{color:var(--accent2)}
.sb-nav{flex:1;overflow-y:auto;overflow-x:hidden;padding:6px 10px 16px}
.sb-sec{font-size:10px;font-weight:600;letter-spacing:1.1px;text-transform:uppercase;color:var(--muted);
  padding:16px 10px 6px;white-space:nowrap}
.sb-item{display:flex;align-items:center;gap:12px;width:100%;padding:8px 10px;margin-bottom:2px;border-radius:7px;
  border:1px solid transparent;background:none;color:var(--muted);font-family:inherit;font-size:13.5px;font-weight:500;
  cursor:pointer;text-align:left;white-space:nowrap;transition:background .15s,color .15s}
.sb-item svg{width:18px;height:18px;flex-shrink:0}
.sb-item:hover{background:var(--surface2);color:var(--text)}
.sb-item.active{background:rgba(0,212,170,.1);border-color:rgba(0,212,170,.4);color:var(--accent)}
body.sb-collapsed .sidebar{width:64px}
body.sb-collapsed .sb-head{padding:0;justify-content:center}
body.sb-collapsed .logo,body.sb-collapsed .sb-label,body.sb-collapsed .sb-sec span{display:none}
body.sb-collapsed .sb-sec{padding:8px 0;margin:6px 8px;border-top:1px solid var(--border)}
body.sb-collapsed .sb-item{justify-content:center;padding:9px 0}
/* TOPBAR */
.topbar{position:sticky;top:0;z-index:100;background:var(--surface);border-bottom:1px solid var(--border);
  display:flex;align-items:center;padding:0 24px;height:56px;gap:12px}
.crumb{font-size:13px;color:var(--muted);white-space:nowrap}
.crumb strong{color:var(--text);font-weight:600}
.nav-right{display:flex;align-items:center;gap:10px;margin-left:auto}
.icon-btn{display:inline-flex;align-items:center;justify-content:center;width:32px;height:32px;padding:0;
  background:none;border:1px solid var(--border2);border-radius:6px;color:var(--muted);cursor:pointer;transition:all .2s}
.icon-btn:hover{color:var(--text);border-color:var(--accent)}
.icon-btn svg{width:16px;height:16px}
.icon-btn.flat{border-color:transparent}
.icon-btn.flat:hover{background:var(--surface2);border-color:transparent}
.icon-btn.err{color:var(--red);border-color:var(--red)}
.icon-btn.spin svg{animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
@media print{.sidebar,.topbar{display:none!important}}
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

<div class="app">

<!-- MENÚ LATERAL (plegable con el botón de arriba; se recuerda en localStorage) -->
<aside class="sidebar">
  <div class="sb-head">
    <div class="logo">MONCAR<span>+RMH</span></div>
    <button class="icon-btn flat" onclick="toggleSidebar()" title="Plegar / desplegar menú"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M9 3v18"/></svg></button>
  </div>
  <nav class="sb-nav">
    <div class="sb-sec"><span>General</span></div>
    <button class="sb-item active" data-page="overview" data-sec="General" title="Overview" onclick="showPage('overview')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg><span class="sb-label">Overview</span></button>
    <div class="sb-sec"><span>Ventas</span></div>
    <button class="sb-item" data-page="ventas" data-sec="Ventas" title="Ventas" onclick="showPage('ventas')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/></svg><span class="sb-label">Ventas</span></button>
    <button class="sb-item" data-page="productos" data-sec="Ventas" title="Productos" onclick="showPage('productos')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 21.73a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73z"/><path d="M12 22V12"/><path d="m3.3 7 8.7 5 8.7-5"/></svg><span class="sb-label">Productos</span></button>
    <button class="sb-item" data-page="marcas" data-sec="Ventas" title="Marcas" onclick="showPage('marcas')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r="1"/></svg><span class="sb-label">Marcas</span></button>
    <button class="sb-item" data-page="clientes" data-sec="Ventas" title="Clientes" onclick="showPage('clientes')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg><span class="sb-label">Clientes</span></button>
    <div class="sb-sec"><span>Producto y stock</span></div>
    <button class="sb-item" data-page="stock" data-sec="Producto y stock" title="Stock" onclick="showPage('stock')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="20" height="5" x="2" y="3" rx="1"/><path d="M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"/><path d="M10 12h4"/></svg><span class="sb-label">Stock</span></button>
    <div class="sb-sec"><span>Plata</span></div>
    <button class="sb-item" data-page="pl" data-sec="Plata" title="P&L" onclick="showPage('pl')"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg><span class="sb-label">P&amp;L</span></button>
  </nav>
</aside>

<div class="main">
<!-- BARRA SUPERIOR -->
<header class="topbar">
  <div class="crumb" id="crumb"><span>General</span> › <strong>Overview</strong></div>
  <div class="nav-right">
    <div style="position:relative">
      <button class="stk-btn active" id="vend-btn" onclick="toggleVendDD(event)" style="font-size:12px;padding:6px 12px">Todos los vendedores ▾</button>
      <div id="vend-dd" onmousedown="event.preventDefault()" style="display:none;position:absolute;top:34px;right:0;background:var(--surface);border:1px solid var(--border);border-radius:6px;min-width:170px;max-height:260px;overflow-y:auto;z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.3);padding:6px"></div>
    </div>
    <div class="badge" id="badge-info">Cargando...</div>
    <button class="icon-btn" id="refresh-btn" onclick="refreshData()" title="Actualizar datos"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg></button>
    <button class="theme-btn" id="theme-toggle" onclick="toggleTheme()" title="Cambiar tema">☀</button>
  </div>
</header>

<!-- ── Controles compartidos: período + línea (visible en Overview, Productos, Marcas, Clientes) ── -->
<div id="analytics-controls">
  <div class="period-bar">
    <div class="period-tabs">
      <button class="period-tab active" data-mode="mes">Mes</button>
      <button class="period-tab" data-mode="3m">3 Meses</button>
      <button class="period-tab" data-mode="6m">6 Meses</button>
      <button class="period-tab" data-mode="mat" title="Moving Annual Total: últimos 12 meses móviles">MAT</button>
      <button class="period-tab" data-mode="año">YTD</button>
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

  <!-- Tableros vendidos (mismos filtros de período/línea/vendedor que el resto del Overview) -->
  <div class="card" id="ov-tab" style="margin-bottom:16px"></div>

  <!-- Mercadería Lita (no depende de los filtros de período/línea/vendedor) -->
  <div class="card" id="ov-lita" style="margin-bottom:16px"></div>

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
    <div class="kpi-row cols4" id="prod-totales" style="margin-bottom:14px"></div>
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
    <div class="kpi-row cols4" id="marc-totales" style="margin-bottom:14px"></div>
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
    <div class="kpi-row cols4" id="cli-totales" style="margin-bottom:14px"></div>
    <div class="tbl-wrap"><table id="tbl-cli"><thead><tr>
      <th>#</th><th>CLIENTE</th><th class="num">FACTURACIÓN</th><th class="num">UDS</th><th class="num">% TOTAL</th><th>PROVINCIA</th>
    </tr></thead><tbody></tbody></table></div>
  </div>
</div>

<div id="page-ventas" class="page">
  <h2>Ventas — Detalle</h2>
  <p style="font-size:12px;color:var(--muted);margin:-8px 0 14px">
    El per&iacute;odo y la l&iacute;nea se filtran con los controles de arriba de todo. Ac&aacute; abajo pod&eacute;s
    cruzar por vendedor, marca, cliente, zona o tipo de cliente — como armar una tabla din&aacute;mica a mano, pero solo.
  </p>
  <div class="card" style="margin-bottom:14px">
    <div id="vdet-filters" style="display:flex;align-items:flex-start;gap:18px;flex-wrap:wrap"></div>
  </div>
  <div class="card">
    <div class="kpi-row cols4" id="vdet-totales" style="margin-bottom:14px"></div>
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-wrap:wrap">
      <input id="vdet-search" type="text" placeholder="Buscar por cliente, c&oacute;digo, marca o descripci&oacute;n..."
        oninput="renderVentasDetalle()"
        style="flex:1;min-width:220px;background:var(--surface2);border:1px solid var(--border2);color:var(--text);
          font-family:inherit;font-size:12px;padding:7px 12px;border-radius:6px;outline:none">
      <span id="vdet-count" style="font-size:12px;color:var(--muted);white-space:nowrap"></span>
      <button class="stk-btn active" style="padding:6px 14px;font-size:12px" onclick="exportVentasCSV()">Descargar CSV</button>
    </div>
    <div class="tbl-wrap" style="max-height:65vh;overflow-y:auto">
      <table id="vdet-table">
        <thead><tr id="vdet-head"></tr></thead>
        <tbody id="vdet-body"></tbody>
      </table>
    </div>
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
  <div class="stock-toggle">
    <button class="stk-btn active" id="pl-btn-eco" onclick="togglePLModo('economica')">Económica</button>
    <button class="stk-btn" id="pl-btn-fin" onclick="togglePLModo('financiera')">Financiera (Rendición)</button>
  </div>

  <div id="pl-economica-body">
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

  <div id="pl-financiera-body" style="display:none">
    <p style="font-size:12px;color:var(--muted);margin:-4px 0 14px">
      Agrupado por el mes en que cada venta se cobra de verdad (no por cuando se vendió) —
      lo que hay que cobrar y liquidar a cada vendedor este mes.
    </p>
    <div class="pills" id="pills-pl-fin"></div>
    <div class="card" id="pl-fin-table-wrap">
      <h3>Rendición del mes</h3>
      <div id="pl-fin-table"></div>
    </div>
  </div>
</div>

</div><!-- /.main -->
</div><!-- /.app -->

<script>
/*INJECT_DATA*/

// ── Data aliases ─────────────────────────────────────────────────────────────
const V       = window.DATA.ventas;
const MC_ROWS = window.DATA.stock_mc.rows;
const RMH_ROWS= window.DATA.stock_rmh.rows;
const LP      = window.DATA.precios;

// ── State ────────────────────────────────────────────────────────────────────
let curPage = 'overview';
// F.vendedores: Set de nombres seleccionados. Vacío = todos (sin filtrar).
const F = { vendedores: new Set(), linea:'ALL', pl:'ALL', stock:'mc' };
function applyVendedorFilter(rows) {
  return F.vendedores.size === 0 ? rows : rows.filter(r => F.vendedores.has(r.VENDEDOR));
}

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
  let rows = applyVendedorFilter(V.raw);
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

function toggleSidebar() {
  const c = document.body.classList.toggle('sb-collapsed');
  try { localStorage.setItem('moncar_sidebar', c ? 'collapsed' : 'open'); } catch(e) {}
  // Los gráficos de Chart.js se reacomodan al nuevo ancho
  setTimeout(()=>Object.values(charts).forEach(ch=>ch.resize()), 200);
}

// Botón refresh (arriba a la derecha): vuelve a leer los Google Sheets vía el servidor
function refreshData() {
  const btn = document.getElementById('refresh-btn');
  if(btn.classList.contains('spin')) return;
  btn.classList.remove('err'); btn.classList.add('spin'); btn.title = 'Actualizando...';
  fetch('/refresh', {cache:'no-store'})
    .then(r => r.json())
    .then(d => {
      if(d.ok) { location.reload(); return; }
      btn.classList.remove('spin'); btn.classList.add('err'); btn.title = 'Error al generar — reintentar';
    })
    .catch(() => { btn.classList.remove('spin'); btn.classList.add('err'); btn.title = 'Sin conexión con el servidor — reintentar'; });
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
    return [...new Set(applyVendedorFilter(V.raw).map(r=>r.MES))].sort();
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
  if(mode==='mat') { const ms=lastNMonths(cur,12); return `${periodLabel('mes',ms[0])} – ${periodLabel('mes',ms[11])}`; }
  if(mode==='custom') return PERIOD.cStart&&PERIOD.cEnd ? `${PERIOD.cStart} → ${PERIOD.cEnd}` : 'Personalizado';
  return cur;
}

function prevPer(mode,cur) {
  if(mode==='mes'||mode==='3m'||mode==='6m'||mode==='mat') return prevMes(cur);
  if(mode==='año')  return String(+cur-1);
  if(mode==='semana') { const[y,w]=cur.split('-W').map(Number); return w===1?`${y-1}-W52`:`${y}-W${String(w-1).padStart(2,'0')}`; }
  return cur;
}
function nextPer(mode,cur) {
  if(mode==='mes'||mode==='3m'||mode==='6m'||mode==='mat') return nextMes(cur);
  if(mode==='año')  return String(+cur+1);
  if(mode==='semana') { const[y,w]=cur.split('-W').map(Number); return w>=52?`${y+1}-W01`:`${y}-W${String(w+1).padStart(2,'0')}`; }
  return cur;
}

function rowsForPeriod(mode, cur, cStart, cEnd) {
  const base = applyVendedorFilter(V.raw);
  if(!cur && mode!=='custom') return base;
  if(mode==='mes')    return base.filter(r=>r.MES===cur);
  if(mode==='año')    return base.filter(r=>r.MES&&r.MES.startsWith(cur));
  if(mode==='3m')     { const ms=lastNMonths(cur,3); return base.filter(r=>ms.includes(r.MES)); }
  if(mode==='6m')     { const ms=lastNMonths(cur,6); return base.filter(r=>ms.includes(r.MES)); }
  if(mode==='mat')    { const ms=lastNMonths(cur,12); return base.filter(r=>ms.includes(r.MES)); }
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
  // Para 3m/6m/mat: desplaza el fin del período N meses atrás
  const n = PERIOD.mode==='3m'?3 : PERIOD.mode==='6m'?6 : PERIOD.mode==='mat'?12 : 1;
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
    if(PERIOD.mode==='3m'||PERIOD.mode==='6m'||PERIOD.mode==='mat') {
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
    else if(PERIOD.mode==='3m'||PERIOD.mode==='6m'||PERIOD.mode==='mat') PERIOD.current = last;
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
  const allRows = applyLineaFilter(applyVendedorFilter(V.raw));
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

  renderTableros(rows, cmp);

  renderLita();
}

// ─────────────────────────────────────────────────────────────────────────────
// MERCADERÍA LITA — lo que Maxi vendió del stock de Lita y hay que reponerle.
// Pendiente = filas "Salida Lita" con CONFIRMACIÓN=Pendiente (todavía cuentan
// en EMBOLSADO DEP.). "Entregar" → POST /lita/entregar (servidor_drive.py).
// Lita solo recibe embolsado ESTÁNDAR: el premium no cubre lo pendiente.
// ─────────────────────────────────────────────────────────────────────────────
const LITA = window.DATA.lita || {pend:[], ent:[]};
const LITA_PEND = {};   // cod → {qty, desde, remitos:Set}
LITA.pend.forEach(p => {
  const o = LITA_PEND[p.cod] || (LITA_PEND[p.cod] = {qty:0, desde:p.fecha, remitos:new Set()});
  o.qty += p.qty;
  if(p.fecha && p.fecha < o.desde) o.desde = p.fecha;
  if(p.remito) o.remitos.add(p.remito.replace('Remito N° ',''));
});
const MC_BY_COD = {};
MC_ROWS.forEach(r => { MC_BY_COD[r.COD.split(' ').join('').toUpperCase()] = r; });
function mcRow(cod) { return MC_BY_COD[cod.split(' ').join('').toUpperCase()]; }
function litaPendQty(cod) { const o = LITA_PEND[cod.split(' ').filter(Boolean).join(' ').toUpperCase()]; return o ? o.qty : 0; }

let LITA_TAB = 'pend';
let LITA_MES = null;
function fmtFecha(iso) { const p = (iso||'').split('-'); return p.length===3 ? p[2]+'/'+p[1]+'/'+p[0] : (iso||'—'); }

function renderLita() {
  const el = document.getElementById('ov-lita');
  if(!el) return;
  const tabs = '<div class="stock-toggle" style="margin-bottom:14px">'
    + '<button class="stk-btn '+(LITA_TAB==='pend'?'active':'')+'" onclick="LITA_TAB=\\'pend\\';renderLita()">Pendiente</button>'
    + '<button class="stk-btn '+(LITA_TAB==='hist'?'active':'')+'" onclick="LITA_TAB=\\'hist\\';renderLita()">Historial de entregas</button>'
    + '</div>';
  el.innerHTML = '<div class="card-title"><span class="dot"></span>Mercadería Lita · a devolver (solo MONCAR)</div>'
    + tabs + (LITA_TAB==='pend' ? litaPendHTML() : litaHistHTML());
}

function litaPendHTML() {
  const items = Object.entries(LITA_PEND).map(([cod,o]) => {
    const r = mcRow(cod) || {};
    const emb = r.EMB_DEP||0, noEmb = r.NO_EMB||0;
    const falta = Math.max(0, o.qty - Math.max(0, emb));  // stock negativo = 0, no "debe" más
    return {cod, marca:r.MARCA||'—', qty:o.qty, desde:o.desde, remitos:[...o.remitos], emb, noEmb, falta};
  }).filter(i => i.qty > 0)
    .sort((a,b) => (b.falta>0)-(a.falta>0) || b.qty-a.qty || a.cod.localeCompare(b.cod));

  if(!items.length) return '<div class="empty" style="padding:30px">No hay mercadería pendiente para Lita.</div>';

  const totUds  = items.reduce((a,i)=>a+i.qty,0);
  const aEmb    = items.filter(i=>i.falta>0);
  const totFalta= aEmb.reduce((a,i)=>a+i.falta,0);
  const kpis = '<div class="kpi-row cols4" style="margin-bottom:14px">'
    + kpi('Uds pendientes', fmtN(totUds), items.length+' códigos', 'amber')
    + kpi('Listos para separar', fmtN(items.length-aEmb.length), 'códigos con embolsado suficiente', 'green')
    + kpi('A embolsar', fmtN(aEmb.length), fmtN(totFalta)+' uds faltan embolsar', aEmb.length?'red':'green')
    + kpi('Pendiente desde', fmtFecha(items.reduce((m,i)=>i.desde<m?i.desde:m, items[0].desde)), 'venta más vieja sin entregar', 'blue')
    + '</div>';

  const rows = items.map(i => {
    let emb;
    if(i.falta<=0) emb = '<span style="color:var(--green);font-weight:600">OK</span>';
    else if(i.noEmb<=0) emb = '<span style="color:var(--red);font-weight:600">Faltan '+fmtN(i.falta)+' — sin stock sin embolsar</span>';
    else if(i.noEmb<i.falta) emb = '<span style="color:var(--red);font-weight:600">Embolsar '+fmtN(i.noEmb)+'</span> <span style="color:var(--muted);font-size:11px">(faltan '+fmtN(i.falta)+', no alcanza el sin embolsar)</span>';
    else emb = '<span style="color:var(--red);font-weight:600">Embolsar '+fmtN(i.falta)+'</span>';
    return '<tr class="'+(i.falta>0?'row-danger':'')+'">'
      + '<td class="cod">'+i.cod+'</td>'
      + '<td style="font-size:12px">'+i.marca+'</td>'
      + '<td class="num" style="font-weight:700">'+fmtN(i.qty)+'</td>'
      + '<td class="num" style="color:var(--muted)">'+fmtFecha(i.desde)+'</td>'
      + '<td class="num">'+fmtN(i.emb)+'</td>'
      + '<td class="num" style="color:var(--muted)">'+fmtN(i.noEmb)+'</td>'
      + '<td style="font-size:12px">'+emb+'</td>'
      + '<td style="white-space:nowrap">'
      +   '<input type="number" min="1" max="'+i.qty+'" value="'+i.qty+'" id="lita-q-'+i.cod.split(' ').join('_')+'" class="date-inp" style="width:70px;padding:4px 6px">'
      +   ' <button class="stk-btn" style="padding:4px 12px;font-size:12px" onclick="litaEntregar(\\''+i.cod+'\\', this)">Entregar</button>'
      + '</td>'
      + '</tr>';
  }).join('');

  // Texto para WhatsApp (Lita lo chequea contra lo suyo): por código, negrita con *...*
  const hoy = new Date().toLocaleDateString('es-AR', {day:'2-digit', month:'2-digit', year:'numeric'});
  LITA_COPY_TXT = '*Pendiente Lita — '+hoy+'*\\n'
    + [...items].sort((a,b)=>a.cod.localeCompare(b.cod)).map(i => i.cod+' x '+fmtN(i.qty)).join('\\n')
    + '\\n*Total: '+fmtN(totUds)+' uds ('+items.length+' códigos)*';
  const copiar = '<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">'
    + '<button class="stk-btn active" style="padding:5px 14px;font-size:12px" onclick="litaCopiar(this)">Copiar lista</button>'
    + '<span style="font-size:12px;color:var(--muted)">Códigos y cantidades pendientes, listo para pegar en WhatsApp</span>'
    + '</div>';

  return kpis + copiar
    + '<div class="tbl-wrap"><table><thead><tr>'
    + '<th>COD</th><th>MARCA</th><th class="num">Pendiente</th><th class="num">Desde</th>'
    + '<th class="num">Emb. estándar dep.</th><th class="num">Sin emb.</th><th>¿Embolsar?</th><th>Marcar entregado</th>'
    + '</tr></thead><tbody>'+rows+'</tbody></table></div>'
    + '<div class="note">Se acumula por código con cada venta MONCAR de Maxi (remitos-app). '
    + '<strong>Emb. estándar dep.</strong>: embolsado estándar en depósito — incluye lo pendiente de Lita, que recién se resta al marcarlo entregado. El premium no cuenta (Lita solo recibe estándar). '
    + '<strong>Entregar</strong>: podés poner una cantidad parcial; se descuenta primero de las ventas más viejas y se resta del stock.</div>';
}

function litaHistHTML() {
  const meses = [...new Set(LITA.ent.map(e=>e.mes).filter(Boolean))].sort().reverse();
  if(!meses.length) return '<div class="empty" style="padding:30px">Todavía no hay entregas registradas.</div>';
  if(!LITA_MES || (LITA_MES!=='ALL' && !meses.includes(LITA_MES))) LITA_MES = meses[0];
  const sel = LITA.ent.filter(e => LITA_MES==='ALL' || e.mes===LITA_MES);
  const aprox = sel.some(e=>!e.fent);

  const map = {};
  sel.forEach(e => {
    const o = map[e.cod] || (map[e.cod] = {cod:e.cod, qty:0, fechas:new Set()});
    o.qty += e.qty;
    if(e.fent) o.fechas.add(e.fent);
  });
  const items = Object.values(map).sort((a,b)=>b.qty-a.qty || a.cod.localeCompare(b.cod));
  const total = items.reduce((a,i)=>a+i.qty,0);
  const soloAprox = m => LITA.ent.filter(e=>e.mes===m).every(e=>!e.fent);

  const selector = '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;font-size:12px;color:var(--muted)">'
    + 'Mes de entrega: <select onchange="LITA_MES=this.value;renderLita()">'
    + meses.map(m=>'<option value="'+m+'" '+(m===LITA_MES?'selected':'')+'>'+shortMes(m)+(soloAprox(m)?' (aprox.)':'')+'</option>').join('')
    + '<option value="ALL" '+(LITA_MES==='ALL'?'selected':'')+'>Todos</option>'
    + '</select>'
    + '<span style="margin-left:auto">'+fmtN(total)+' uds · '+items.length+' códigos</span></div>';

  const rows = items.map(i => '<tr>'
    + '<td class="cod">'+i.cod+'</td>'
    + '<td style="font-size:12px">'+((mcRow(i.cod)||{}).MARCA||'—')+'</td>'
    + '<td class="num" style="font-weight:700">'+fmtN(i.qty)+'</td>'
    + '<td style="font-size:12px;color:var(--muted)">'+([...i.fechas].sort().map(fmtFecha).join(', ')||'sin fecha')+'</td>'
    + '</tr>').join('');

  return selector
    + '<div class="tbl-wrap"><table><thead><tr><th>COD</th><th>MARCA</th><th class="num">Uds entregadas</th><th>Fecha(s) de entrega</th></tr></thead>'
    + '<tbody>'+rows+'<tr style="font-weight:700"><td>TOTAL</td><td></td><td class="num">'+fmtN(total)+'</td><td></td></tr></tbody></table></div>'
    + (aprox ? '<div class="note">Entregas marcadas antes de que existiera la columna FECHA ENTREGA (25/09/2026): se agrupan por mes de la venta, no de la entrega.</div>' : '');
}

let LITA_COPY_TXT = '';
function litaCopiar(btn) {
  const ok = () => { btn.textContent = '¡Copiado!'; setTimeout(()=>{ btn.textContent = 'Copiar lista'; }, 1800); };
  const fallback = () => {  // por si el navegador bloquea navigator.clipboard (ej. HTML abierto suelto)
    const ta = document.createElement('textarea');
    ta.value = LITA_COPY_TXT; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); ok(); } catch(e) { alert('No se pudo copiar'); }
    ta.remove();
  };
  if(navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(LITA_COPY_TXT).then(ok, fallback);
  else fallback();
}

function litaEntregar(cod, btn) {
  const inp = document.getElementById('lita-q-'+cod.split(' ').join('_'));
  const q = parseInt(inp && inp.value, 10);
  const max = litaPendQty(cod);
  if(!(q>0) || q>max) { alert('Cantidad inválida: tiene que ser entre 1 y '+max+'.'); return; }
  if(!confirm('¿Marcar '+q+' uds de '+cod+' como entregadas a Lita?\\n\\nSe restan del embolsado en depósito (se escribe en STOCK_MONCAR).')) return;
  btn.disabled = true; btn.textContent = 'Guardando...';
  fetch('/lita/entregar', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({cod, cantidad:q})})
    .then(r => r.json())
    .then(d => {
      if(d.ok) { location.reload(); return; }
      alert('No se pudo guardar: '+(d.error||'error desconocido'));
      btn.disabled = false; btn.textContent = 'Entregar';
    })
    .catch(() => {
      alert('No se pudo conectar con el servidor. ¿Abriste el dashboard con abrir_dashboard_drive.bat?');
      btn.disabled = false; btn.textContent = 'Entregar';
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// TABLEROS VENDIDOS — un tablero organizador se vende como un item en
// remitos-app y queda en Ventas como una fila por codigo con TABLERO="009".
// Un tablero vendido = par (REMITO, TABLERO). Solo lectura sobre V.raw.
// ─────────────────────────────────────────────────────────────────────────────
function agruparTableros(rows) {
  const m = {};
  rows.forEach(r => {
    if(!r.TABLERO) return;
    const k = r.REMITO + '|' + r.TABLERO;
    const o = m[k] || (m[k] = {num:r.TABLERO, remito:r.REMITO, fecha:r.FECHA, cliente:r.CLIENTE, vend:r.VENDEDOR, uds:0, rev:0, cmv:0, cods:new Set()});
    o.uds += r.CANTIDAD; o.rev += r.SUBTOTAL; o.cmv += (r.CMV||0); o.cods.add(r.COD);
  });
  return Object.values(m);
}

function renderTableros(rows, cmp) {
  const el = document.getElementById('ov-tab');
  if(!el) return;
  const tabs = agruparTableros(rows).sort((a,b) => (b.fecha||'').localeCompare(a.fecha||'') || b.num.localeCompare(a.num));
  const prevN = cmp ? agruparTableros(cmp).length : null;
  const n = tabs.length;
  const rev = tabs.reduce((a,t)=>a+t.rev,0);
  const uds = tabs.reduce((a,t)=>a+t.uds,0);
  const cmv = tabs.reduce((a,t)=>a+t.cmv,0);
  const revTot = rows.reduce((a,r)=>a+r.SUBTOTAL,0);
  const kpis = '<div class="kpi-row cols4">' + [
    kpi('Tableros vendidos', fmtN(n), n ? fmtN(uds)+' uds en tableros' : 'ninguno en el período', 'green', changeBadge(n, prevN)),
    kpi('Facturación tableros', fmt(rev), revTot>0 ? pct(rev/revTot)+' de la facturación' : '—', 'amber', ''),
    kpi('Promedio por tablero', n ? fmt(rev/n) : '—', n ? fmtN(uds/n)+' uds por tablero' : '—', 'blue', ''),
    kpi('Margen bruto tableros', rev>0 ? pct((rev-cmv)/rev) : '—', 'sin costo del organizador', 'muted', ''),
  ].join('') + '</div>';
  const tabla = n ? '<div class="tbl-wrap"><table><thead><tr>'
      + '<th>N° TABLERO</th><th>FECHA</th><th>CLIENTE</th><th>VENDEDOR</th><th class="num">REMITO</th><th class="num">CÓDIGOS</th><th class="num">UDS</th><th class="num">FACTURACIÓN</th>'
      + '</tr></thead><tbody>'
      + tabs.map(t => '<tr><td class="cod">'+t.num+'</td><td>'+fmtFecha(t.fecha)+'</td><td>'+(t.cliente||'—')+'</td><td>'+(t.vend||'—')+'</td>'
        + '<td class="num">'+(t.remito||'—')+'</td><td class="num">'+t.cods.size+'</td><td class="num">'+fmtN(t.uds)+'</td><td class="num">'+fmt(t.rev)+'</td></tr>').join('')
      + '</tbody></table></div>'
    : '<div class="empty" style="padding:20px">Sin tableros vendidos en el período seleccionado.</div>';
  el.innerHTML = '<div class="card-title"><span class="dot amber"></span>Tableros vendidos · período seleccionado</div>' + kpis + tabla;
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

  // Totales del filtro actual (reacciona a período/línea/vendedor y al buscador de producto)
  const totalesEl = document.getElementById('prod-totales');
  if(totalesEl) {
    const totalUds = sorted.reduce((a,p)=>a+p.uds, 0);
    const prom = totalUds>0 ? totalRev/totalUds : 0;
    totalesEl.innerHTML = [
      kpi('Facturación filtrada', fmt(totalRev), `${sorted.length} producto${sorted.length===1?'':'s'}`, 'green'),
      kpi('Unidades', fmtN(totalUds), 'en el filtro actual', 'blue'),
      kpi('Precio Promedio', prom>0?fmt(prom):'—', 'ponderado por unidades', 'accent'),
      kpi('Productos distintos', sorted.length, 'con ventas en el período', 'amber'),
    ].join('');
  }

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

  // Totales del filtro actual
  const totalesEl = document.getElementById('marc-totales');
  if(totalesEl) {
    const totalUds = sorted.reduce((a,m)=>a+m.uds,0);
    const totalCmv = sorted.reduce((a,m)=>a+m.cmv,0);
    totalesEl.innerHTML = [
      kpi('Facturación filtrada', fmt(totalRev), `${sorted.length} marca${sorted.length===1?'':'s'}`, 'green'),
      kpi('Unidades', fmtN(totalUds), 'en el filtro actual', 'blue'),
      kpi('Margen Bruto', fmt(totalRev-totalCmv), pct(totalRev>0?(totalRev-totalCmv)/totalRev:0), 'accent'),
      kpi('Marcas distintas', sorted.length, 'con ventas en el período', 'amber'),
    ].join('');
  }

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

  // Totales del filtro actual
  const totalesEl = document.getElementById('cli-totales');
  if(totalesEl) {
    const totalUds = sorted.reduce((a,c)=>a+c.uds,0);
    const prom = totalUds>0 ? totalRev/totalUds : 0;
    totalesEl.innerHTML = [
      kpi('Facturación filtrada', fmt(totalRev), `${sorted.length} cliente${sorted.length===1?'':'s'}`, 'green'),
      kpi('Unidades', fmtN(totalUds), 'en el filtro actual', 'blue'),
      kpi('Precio Promedio', prom>0?fmt(prom):'—', 'ponderado por unidades', 'accent'),
      kpi('Clientes distintos', sorted.length, 'con compras en el período', 'amber'),
    ].join('');
  }

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
    // Lo pendiente de Lita ya está vendido (sigue en EMB_DEP hasta marcarlo
    // entregado): se reserva del embolsado ESTÁNDAR — el premium no la cubre.
    const litaPend  = isMC ? litaPendQty(r.COD) : 0;
    const faltaLita = Math.max(0, litaPend - Math.max(0, r.EMB_DEP));
    const embDep = !isMC ? r.EMB_DEP
      : litaPend>0 ? Math.max(0, r.EMB_DEP-litaPend)+r.PREM_DEP
      : r.EMB_DEP+r.PREM_DEP;  // sin Lita pendiente: cálculo original, sin tocar
    const noEmb  = r.NO_EMB;
    const totalS = isMC
      ? noEmb+r.EMB_DEP+r.PREM_DEP+r.CABA
      : noEmb+r.EMB_DEP+r.MAXI;
    const wEmb  = avgW>0 ? embDep/avgW : Infinity;
    const wMaxi = (!isMC&&avgW>0) ? (r.MAXI||0)/avgW : null;
    const mTot  = avgM>0 ? totalS/avgM : Infinity;
    const toEmb = Math.min(faltaLita + (avgM>0 ? Math.max(0, Math.ceil(avgM*1.5-embDep)) : 0), noEmb);
    const buy   = isFinite(mTot)&&mTot<2;
    const buyQty= buy ? Math.max(0, Math.ceil(3*avgM-totalS)) : 0;
    // Prioridad forzada: sin ventas registradas pero Maxi se quedó sin stock y hay
    // sin-embolsar disponible en depósito — hay que embolsar igual, aunque no
    // haya historial de ventas que lo justifique por velocidad.
    const urgentSinVenta = !isMC && avgM<=0 && noEmb>0 && (r.MAXI||0)<=0;
    // Prioridad forzada: no hay embolsado estándar suficiente para lo que se le debe a Lita.
    const urgentLita = faltaLita>0;
    return {
      cod:r.COD, lbl:isMC?r.MARCA:r.DESC,
      embDep, noEmb, maxi:isMC?null:r.MAXI,
      avgM, wEmb, wMaxi, mTot, toEmb, buy, buyQty, urgentSinVenta,
      litaPend, faltaLita, urgentLita
    };
  }).filter(Boolean).sort((a,b)=> (b.urgentSinVenta-a.urgentSinVenta) || (b.urgentLita-a.urgentLita) || (a.wEmb-b.wEmb) || (b.noEmb-a.noEmb));
}

function renderEmbolsadoCard(topEmb, isMC, topN) {
  if(!topEmb.length) return '<div class="card" style="margin-bottom:16px"><div class="empty">Sin datos de venta en los últimos 6 meses para calcular prioridad de embolsado.</div></div>';
  const mFmt = m => isFinite(m)&&m<=36 ? m.toFixed(1)+'m' : '∞';
  const rows = topEmb.map(r => {
    const maxiTd  = !isMC ? '<td class="num">'+fmtN(r.maxi||0)+'</td>' : '';
    const maxiSem = !isMC ? '<td class="num" style="color:'+urgColor(r.wMaxi)+'">'+urgLabel(r.wMaxi)+'</td>' : '';
    const rc = (r.urgentSinVenta || r.urgentLita || r.wEmb<1) ? 'row-danger' : r.wEmb<3 ? 'row-warn' : '';
    const tag = (r.urgentSinVenta ? ' <span style="color:var(--red);font-weight:600">· Maxi sin stock</span>' : '')
      + (r.urgentLita ? ' <span style="color:var(--red);font-weight:600">· Lita: faltan '+fmtN(r.faltaLita)+'</span>' : '');
    const litaTd = isMC ? '<td class="num" style="color:var(--muted)">'+(r.litaPend>0?fmtN(r.litaPend):'—')+'</td>' : '';
    return '<tr class="'+rc+'">'
      +'<td class="cod">'+r.cod+'</td>'
      +'<td style="font-size:12px">'+(r.lbl||'—')+tag+'</td>'
      +'<td class="num">'+fmtN(r.embDep)+'</td>'
      +'<td class="num">'+fmtN(r.noEmb)+'</td>'
      +litaTd
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
    +(isMC ? '<th class="num">Lita pend.</th>' : '')
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
    +(!isMC ? ' &nbsp;|&nbsp; <strong>Maxi sin stock</strong>: se prioriza igual aunque no haya ventas registradas, si hay sin embolsar disponible en depósito.' : '')
    +(isMC ? ' &nbsp;|&nbsp; <strong>Lita pend.</strong>: uds vendidas por Maxi todavía sin entregar a Lita — se reservan del embolsado estándar, así que <strong>Emb. Dep.</strong> ya las descuenta. <strong>Lita: faltan X</strong>: no alcanza el estándar para Lita, se prioriza arriba de todo.' : '')
    +'</div>'
    +'</div>';
}

// ─────────────────────────────────────────────────────────────────────────────
// VENTAS — DETALLE (tabla filtrable estilo tabla dinámica)
// ─────────────────────────────────────────────────────────────────────────────
const VDET_COLS = [
  ['FECHA','Fecha'], ['VENDEDOR','Vendedor'], ['LINEA','Línea'], ['COD','Código'],
  ['DESC','Descripción'], ['CLIENTE','Cliente'], ['CANTIDAD','Cant.'],
  ['PRECIO','Precio'], ['SUBTOTAL','Subtotal'], ['MARCA','Marca'], ['ZONA','Zona'], ['MES','Mes'],
];
let VDET_SORT = { col:'FECHA', dir:-1 };

function setVdetSort(col) {
  if(VDET_SORT.col===col) VDET_SORT.dir *= -1;
  else VDET_SORT = { col, dir:1 };
  renderVentasDetalle();
}

// ── Filtros cruzados de esta página: vendedor (compartido con F.vendedores) +
//    marca / cliente / zona / tipo de cliente (propios de esta página) ────────
const VDET_FILTERS = { MARCA:new Set(), CLIENTE:new Set(), ZONA:new Set(), TIPO_CLI:new Set() };
const VDET_FILTER_DEFS = [
  { key:'VENDEDOR', label:'Vendedor',     mode:'pills' },
  { key:'MARCA',    label:'Marca',        mode:'dropdown' },
  { key:'CLIENTE',  label:'Cliente',      mode:'dropdown' },
  { key:'ZONA',     label:'Zona',         mode:'dropdown' },
  { key:'TIPO_CLI', label:'Tipo cliente', mode:'dropdown' },
];
let VDET_DOM = {};

function distinctValues(key) {
  return [...new Set(V.raw.map(r=>r[key]).filter(Boolean))].sort();
}
function vdetActiveSet(key) {
  return key==='VENDEDOR' ? F.vendedores : VDET_FILTERS[key];
}

function toggleVdetFilter(key, val) {
  if(key==='VENDEDOR') { onVendCheck(val, !F.vendedores.has(val)); return; }
  const set = VDET_FILTERS[key];
  if(set.has(val)) set.delete(val); else set.add(val);
  refreshVdetFilterUI();
  renderPage(curPage);
}

function clearVdetFilters() {
  Object.values(VDET_FILTERS).forEach(s=>s.clear());
  F.vendedores.clear();
  updateVendBtnLabel();
  renderVendDD();
  const search = document.getElementById('vdet-search');
  if(search) search.value = '';
  refreshVdetFilterUI();
  renderPage(curPage);
}

function toggleVdetDD(key, e) {
  e.stopPropagation();
  const panel = VDET_DOM[key].panelEl;
  const isOpen = panel.style.display==='block';
  Object.values(VDET_DOM).forEach(d => { if(d.panelEl) d.panelEl.style.display='none'; });
  panel.style.display = isOpen ? 'none' : 'block';
}

function buildVdetFilters() {
  const wrap = document.getElementById('vdet-filters');
  if(!wrap) return;
  wrap.innerHTML = '';
  VDET_DOM = {};

  VDET_FILTER_DEFS.forEach(def => {
    const opts = distinctValues(def.key);
    const box = document.createElement('div');
    const title = document.createElement('div');
    title.textContent = def.label;
    title.style.cssText = 'font-size:10px;font-weight:700;letter-spacing:.06em;color:var(--muted);text-transform:uppercase;margin-bottom:6px';
    box.appendChild(title);

    const entry = { pills: [] };

    if(def.mode==='pills') {
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;gap:6px;flex-wrap:wrap;max-width:280px';
      opts.forEach(v => {
        const p = document.createElement('span');
        p.className = 'pill';
        p.style.cursor = 'pointer';
        p.textContent = v;
        p.addEventListener('click', () => toggleVdetFilter(def.key, v));
        row.appendChild(p);
        entry.pills.push({ val:v, el:p });
      });
      box.appendChild(row);
    } else {
      const posWrap = document.createElement('div');
      posWrap.style.position = 'relative';
      const btn = document.createElement('button');
      btn.className = 'pill';
      btn.style.cursor = 'pointer';
      btn.textContent = def.label + ' ▾';
      btn.addEventListener('click', (e) => toggleVdetDD(def.key, e));
      posWrap.appendChild(btn);

      const panel = document.createElement('div');
      panel.style.cssText = 'display:none;position:absolute;top:32px;left:0;background:var(--surface);'
        + 'border:1px solid var(--border);border-radius:6px;width:220px;max-height:260px;overflow-y:auto;'
        + 'z-index:200;box-shadow:0 4px 16px rgba(0,0,0,.3);padding:6px';

      const search = document.createElement('input');
      search.type = 'text'; search.placeholder = 'Buscar...';
      search.style.cssText = 'width:100%;box-sizing:border-box;background:var(--surface2);border:1px solid var(--border2);'
        + 'color:var(--text);font-family:inherit;font-size:12px;padding:5px 8px;border-radius:5px;outline:none;margin-bottom:6px';
      const list = document.createElement('div');

      opts.forEach(v => {
        const lbl = document.createElement('label');
        lbl.style.cssText = 'display:flex;align-items:center;gap:8px;padding:5px 6px;font-size:12px;cursor:pointer;border-radius:4px';
        lbl.addEventListener('mouseover', () => { lbl.style.background='var(--surface2)'; });
        lbl.addEventListener('mouseout',  () => { lbl.style.background=''; });
        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.addEventListener('change', () => toggleVdetFilter(def.key, v));
        lbl.appendChild(chk);
        lbl.appendChild(document.createTextNode(' '+v));
        list.appendChild(lbl);
        entry.pills.push({ val:v, el:lbl, chk });
      });
      search.addEventListener('input', () => {
        const q = search.value.trim().toLowerCase();
        entry.pills.forEach(({val,el}) => { el.style.display = (!q || val.toLowerCase().includes(q)) ? 'flex' : 'none'; });
      });

      panel.appendChild(search);
      panel.appendChild(list);
      posWrap.appendChild(panel);

      const chips = document.createElement('div');
      chips.style.cssText = 'display:flex;gap:4px;flex-wrap:wrap;margin-top:6px;max-width:220px';
      posWrap.appendChild(chips);

      box.appendChild(posWrap);
      entry.panelEl = panel;
      entry.chipsEl = chips;
    }

    wrap.appendChild(box);
    VDET_DOM[def.key] = entry;
  });

  const clearPill = document.createElement('span');
  clearPill.className = 'pill';
  clearPill.style.cssText = 'cursor:pointer;border-style:dashed;align-self:flex-end';
  clearPill.textContent = '✕ Limpiar filtros';
  clearPill.addEventListener('click', clearVdetFilters);
  wrap.appendChild(clearPill);

  document.addEventListener('click', () => {
    Object.values(VDET_DOM).forEach(d => { if(d.panelEl) d.panelEl.style.display='none'; });
  });

  refreshVdetFilterUI();
}

function refreshVdetFilterUI() {
  VDET_FILTER_DEFS.forEach(def => {
    const set = vdetActiveSet(def.key);
    const entry = VDET_DOM[def.key];
    if(!entry) return;
    if(def.mode==='pills') {
      entry.pills.forEach(({val,el}) => el.classList.toggle('active', set.has(val)));
    } else {
      entry.pills.forEach(({val,chk}) => { chk.checked = set.has(val); });
      entry.chipsEl.innerHTML = '';
      set.forEach(val => {
        const chip = document.createElement('span');
        chip.className = 'pill active';
        chip.style.cssText = 'font-size:11px;padding:3px 8px;cursor:pointer';
        chip.textContent = val + ' ✕';
        chip.addEventListener('click', () => toggleVdetFilter(def.key, val));
        entry.chipsEl.appendChild(chip);
      });
    }
  });
}

function vdetFilteredRows() {
  const q = (document.getElementById('vdet-search')?.value||'').trim().toLowerCase();
  let rows = rowsFiltered();
  ['MARCA','CLIENTE','ZONA','TIPO_CLI'].forEach(key => {
    const set = VDET_FILTERS[key];
    if(set.size>0) rows = rows.filter(r => set.has(r[key]));
  });
  if(q) {
    rows = rows.filter(r =>
      (r.CLIENTE||'').toLowerCase().includes(q) ||
      (r.COD||'').toLowerCase().includes(q) ||
      (r.MARCA||'').toLowerCase().includes(q) ||
      (r.DESC||'').toLowerCase().includes(q)
    );
  }
  const {col,dir} = VDET_SORT;
  const numeric = col==='CANTIDAD'||col==='PRECIO'||col==='SUBTOTAL';
  return rows.slice().sort((a,b) => {
    const av=a[col], bv=b[col];
    if(numeric) return (av-bv)*dir;
    return String(av||'').localeCompare(String(bv||''))*dir;
  });
}

function renderVentasDetalle() {
  if(VDET_DOM.VENDEDOR) refreshVdetFilterUI();
  const rows = vdetFilteredRows();

  // Totales del filtro actual (período/línea/vendedor globales + los propios de esta página)
  const totalesEl = document.getElementById('vdet-totales');
  if(totalesEl) {
    const totalFact = rows.reduce((a,r)=>a+r.SUBTOTAL, 0);
    const totalUds  = rows.reduce((a,r)=>a+r.CANTIDAD, 0);
    const prom = totalUds>0 ? totalFact/totalUds : 0;
    totalesEl.innerHTML = [
      kpi('Facturación filtrada', fmt(totalFact), `${rows.length} fila${rows.length===1?'':'s'}`, 'green'),
      kpi('Unidades', fmtN(totalUds), 'en el filtro actual', 'blue'),
      kpi('Precio Promedio', prom>0?fmt(prom):'—', 'ponderado por unidades', 'accent'),
      kpi('Filas', fmtN(rows.length), 'líneas de venta', 'amber'),
    ].join('');
  }

  const head = VDET_COLS.map(([key,label]) => {
    const arrow = VDET_SORT.col===key ? (VDET_SORT.dir===1?' ▲':' ▼') : '';
    return `<th style="cursor:pointer;user-select:none" onclick="setVdetSort('${key}')">${label}${arrow}</th>`;
  }).join('');
  document.getElementById('vdet-head').innerHTML = head;

  const body = rows.map(r => `<tr>`
    + `<td>${r.FECHA||''}</td>`
    + `<td>${r.VENDEDOR||''}</td>`
    + `<td>${r.LINEA||''}</td>`
    + `<td class="cod">${r.COD||''}</td>`
    + `<td style="font-size:12px">${truncate(r.DESC,40)}</td>`
    + `<td style="font-size:12px">${r.CLIENTE||''}</td>`
    + `<td class="num">${fmtN(r.CANTIDAD)}</td>`
    + `<td class="num">${fmt(r.PRECIO)}</td>`
    + `<td class="num">${fmt(r.SUBTOTAL)}</td>`
    + `<td style="font-size:12px">${r.MARCA||''}</td>`
    + `<td style="font-size:12px">${r.ZONA||''}</td>`
    + `<td>${r.MES||''}</td>`
    + `</tr>`).join('');
  document.getElementById('vdet-body').innerHTML = body || '<tr><td colspan="12" class="empty">Sin resultados con estos filtros.</td></tr>';

  document.getElementById('vdet-count').textContent = `${rows.length} fila${rows.length===1?'':'s'}`;
}

function exportVentasCSV() {
  const rows = vdetFilteredRows();
  const bom = String.fromCharCode(0xFEFF);
  const header = VDET_COLS.map(([,label])=>label);
  const lines = [header.join(';')];
  rows.forEach(r => {
    lines.push(VDET_COLS.map(([key]) => {
      let v = r[key];
      if(key==='CANTIDAD') v = Math.round(v);
      if(key==='PRECIO'||key==='SUBTOTAL') v = Math.round(v);
      v = String(v==null?'':v).replace(/;/g,',');
      return v;
    }).join(';'));
  });
  const csv = bom + lines.join('\\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `ventas_detalle_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
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
  const col4    = isMC ? sRows.reduce((a,r)=>a+r.CABA,0)     : sRows.reduce((a,r)=>a+r.MAXI,0);
  const col3lbl = isMC ? 'Prem. Dep' : 'CABA';
  const col4lbl = isMC ? 'CABA'      : 'Maxi';

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
              const col4val = isMC ? r.CABA     : r.MAXI;
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
      ? r.NO_EMB + r.EMB_DEP + r.PREM_DEP + r.CABA
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
// P&L — VISTA FINANCIERA (RENDICIÓN)
// Agrupa por MES_COBRO (cuando la plata entra de verdad) en vez de por MES
// (cuando se vendió) — separa lo economico de lo financiero. Enfocada solo
// en cobros y comisiones a liquidar, sin CMV/Margen Bruto/Resultado Neto
// (esos numeros solo tienen sentido en la vista economica).
// ─────────────────────────────────────────────────────────────────────────────
let PL_MODO = 'economica';
let PL_MES_COBRO = null;

function togglePLModo(modo) {
  PL_MODO = modo;
  document.getElementById('pl-btn-eco').classList.toggle('active', modo==='economica');
  document.getElementById('pl-btn-fin').classList.toggle('active', modo==='financiera');
  document.getElementById('pl-economica-body').style.display = modo==='economica' ? '' : 'none';
  document.getElementById('pl-financiera-body').style.display = modo==='financiera' ? '' : 'none';
  if(modo==='economica') renderPL(); else renderPLFinanciera();
}

function setPillFin(mes) {
  PL_MES_COBRO = mes;
  renderPLFinanciera();
}

function renderPLFinanciera() {
  const rowsBase = applyVendedorFilter(V.raw);
  const meses = [...new Set(rowsBase.map(r=>r.MES_COBRO).filter(Boolean))].sort();
  if(!PL_MES_COBRO || !meses.includes(PL_MES_COBRO)) PL_MES_COBRO = meses[meses.length-1] || null;

  const pillsEl = document.getElementById('pills-pl-fin');
  if(pillsEl) {
    pillsEl.innerHTML = meses.map(m =>
      `<span class="pill${PL_MES_COBRO===m?' active':''}" onclick="setPillFin('${m}')">${m}</span>`
    ).join('');
  }

  const rows = rowsBase.filter(r => r.MES_COBRO === PL_MES_COBRO);
  const totalCobrar = rows.reduce((a,r)=>a+r.SUBTOTAL,0);

  const comPorVend = {};
  rows.forEach(r => { if(r.VENDEDOR) comPorVend[r.VENDEDOR] = (comPorVend[r.VENDEDOR]||0) + r.SUBTOTAL*r.COMISION; });
  const comVendRows = Object.entries(comPorVend).sort((a,b)=>b[1]-a[1]);
  const totalCom = comVendRows.reduce((a,[,c])=>a+c,0);
  const comAdmin = totalCobrar * 0.20;
  const resto = totalCobrar - totalCom - comAdmin;

  const rp = (v) => totalCobrar>0 ? '<span class="pl-pct">'+pct(Math.abs(v)/totalCobrar)+'</span>' : '';

  const filas = [
    {label:'Total a cobrar este mes', val:totalCobrar, style:'color:var(--green);font-weight:700', badge:'<span class="pl-pct">100%</span>'},
  ];

  const vendHtml = comVendRows.map(([v,c]) => `
    <tr>
      <td style="padding:7px 12px 7px 28px;color:var(--accent3);font-size:12px;border-bottom:1px solid var(--border);opacity:.85">↳ ${v}</td>
      <td class="num" style="padding:7px 12px;color:var(--accent3);font-size:12px;border-bottom:1px solid var(--border);opacity:.85">${fmt(-c)}${rp(c)}</td>
    </tr>`).join('');

  document.getElementById('pl-fin-table').innerHTML = !PL_MES_COBRO
    ? '<div class="empty">No hay ventas con mes de cobro cargado todavía.</div>'
    : `
    <table style="width:100%">
      ${filas.map(row=>`<tr>
        <td style="padding:10px 12px;${row.style||''};border-bottom:1px solid var(--border)">${row.label}</td>
        <td class="num" style="padding:10px 12px;${row.style||''};border-bottom:1px solid var(--border)">${fmt(row.val||0)}${row.badge||''}</td>
      </tr>`).join('')}
      <tr>
        <td style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">A liquidar a vendedores</td>
        <td class="num" style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">${fmt(-totalCom)}${rp(totalCom)}</td>
      </tr>
      ${vendHtml}
      <tr>
        <td style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">Comisión administrativa (Franco 20%)</td>
        <td class="num" style="padding:10px 12px;color:var(--accent3);border-bottom:1px solid var(--border)">${fmt(-comAdmin)}${rp(comAdmin)}</td>
      </tr>
      <tr>
        <td style="padding:10px 12px;color:${resto>=0?'var(--green)':'var(--red)'};font-weight:700;font-size:16px;border-top:1px solid var(--border2)">Te queda a vos</td>
        <td class="num" style="padding:10px 12px;color:${resto>=0?'var(--green)':'var(--red)'};font-weight:700;font-size:16px;border-top:1px solid var(--border2)">${fmt(resto)}${rp(resto)}</td>
      </tr>
    </table>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTER
// ─────────────────────────────────────────────────────────────────────────────
function renderPage(name) {
  if(name==='overview')  renderOverview();
  if(name==='ventas')    renderVentasDetalle();
  if(name==='productos') renderProductos();
  if(name==='marcas')    renderMarcas();
  if(name==='clientes')  renderClientes();
  if(name==='stock')     renderStock();
  if(name==='pl')        { if(PL_MODO==='economica') renderPL(); else renderPLFinanciera(); }
}

function showPage(name) {
  curPage = name;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  document.querySelectorAll('.sb-item[data-page]').forEach(t=>{
    const on = t.dataset.page===name;
    t.classList.toggle('active', on);
    if(on) document.getElementById('crumb').innerHTML = '<span>'+t.dataset.sec+'</span> › <strong>'+t.title+'</strong>';
  });
  window.scrollTo(0,0);
  const ac = document.getElementById('analytics-controls');
  if(ac) ac.style.display = ['overview','ventas','productos','marcas','clientes'].indexOf(name) >= 0 ? '' : 'none';
  renderPage(name);
}

// ── Selector de vendedor (multi-selección, se arma solo con V.vendedores) ─────
function toggleVendDD(e) {
  if(e) e.stopPropagation();
  const dd = document.getElementById('vend-dd');
  dd.style.display = dd.style.display==='none' ? 'block' : 'none';
}
function renderVendDD() {
  const dd = document.getElementById('vend-dd');
  dd.innerHTML = V.vendedores.map(v => {
    const checked = F.vendedores.has(v) ? 'checked' : '';
    const id = 'vend-chk-'+v.replace(/[^a-zA-Z0-9]/g,'');
    return `<label for="${id}" style="display:flex;align-items:center;gap:8px;padding:5px 8px;font-size:12px;cursor:pointer;border-radius:4px" onmouseover="this.style.background='var(--surface2)'" onmouseout="this.style.background=''">`
      + `<input type="checkbox" id="${id}" ${checked} onchange="onVendCheck('${v}',this.checked)" style="cursor:pointer">${v}</label>`;
  }).join('');
}
function onVendCheck(v, checked) {
  if(checked) F.vendedores.add(v); else F.vendedores.delete(v);
  updateVendBtnLabel();
  renderVendDD();
  if(VDET_DOM.VENDEDOR) refreshVdetFilterUI();
  renderPage(curPage);
}
function updateVendBtnLabel() {
  const btn = document.getElementById('vend-btn');
  const n = F.vendedores.size;
  btn.textContent = (n===0 ? 'Todos los vendedores' : n===1 ? [...F.vendedores][0] : `${n} vendedores`) + ' ▾';
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

  // Menú lateral: estado guardado; si no hay, plegado en pantallas angostas
  let sb = null;
  try { sb = localStorage.getItem('moncar_sidebar'); } catch(e) {}
  if(sb === 'collapsed' || (sb === null && window.innerWidth < 900)) document.body.classList.add('sb-collapsed');
  // Sin servidor (HTML abierto como archivo) no hay /refresh
  if(location.protocol === 'file:') document.getElementById('refresh-btn').style.display = 'none';

  // Selector de vendedor: se arma solo con los nombres que aparecen en Ventas
  renderVendDD();

  // Filtros cruzados de la página Ventas - Detalle (vendedor/marca/cliente/zona/tipo cliente)
  buildVdetFilters();

  // Badge
  const totalReg = V.raw.length;
  document.getElementById('badge-info').textContent = `${window.DATA.generated_at} · ${totalReg} reg.`;

  // Cerrar dropdowns al hacer click fuera
  document.addEventListener('click', function(e){
    const checks = [['prod-dd','prod-search'],['marc-dd','marc-search'],['vend-dd','vend-btn']];
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
