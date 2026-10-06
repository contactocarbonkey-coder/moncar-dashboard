# -*- coding: utf-8 -*-
# Marca mercaderia de Lita como entregada en STOCK_MONCAR!Movimientos.
#
# Contexto: cada venta MONCAR de Maxi escribe (desde remitos-app) una fila
# "Salida Lita" con CONFIRMACIÓN=Pendiente. La formula de Stock recien la
# resta de EMBOLSADO DEP. cuando CONFIRMACIÓN pasa a "Entregado". Este modulo
# hace ese cambio desde el dashboard (servidor_drive.py → POST /lita/entregar)
# y guarda la fecha en FECHA ENTREGA para el historial mensual.
#
# Es la UNICA parte del dashboard que escribe en Sheets. Usa la cuenta de
# servicio de remitos-app (lee su .env.local, no duplica la clave).
import os, json, threading, urllib.request
from datetime import datetime

DIR      = r'C:\Users\franc\OneDrive\Documentos\KeyCarbon\claude MONCAR RMH'
ENV_PATH = os.path.join(DIR, 'remitos-app', '.env.local')
STOCK_MC_SHEET_ID = '1u9xG2uyeT3V8osBi9Qz0xofDCPmFmhOXlu85c1OIEQ8'
HOJA = 'Movimientos'
API  = 'https://sheets.googleapis.com/v4/spreadsheets/' + STOCK_MC_SHEET_ID

# Dos clicks seguidos (o dos pestañas) no deben partir la misma fila dos veces.
_lock = threading.Lock()
_creds = None


def _leer_env():
    env = {}
    with open(ENV_PATH, encoding='utf-8') as f:
        for linea in f:
            linea = linea.strip()
            if not linea or linea.startswith('#') or '=' not in linea:
                continue
            k, v = linea.split('=', 1)
            v = v.strip()
            if len(v) >= 2 and v[0] == v[-1] and v[0] in '"\'':
                v = v[1:-1]
            env[k.strip()] = v
    return env


def _token():
    global _creds
    from google.oauth2 import service_account
    from google.auth.transport.requests import Request
    if _creds is None:
        env = _leer_env()
        _creds = service_account.Credentials.from_service_account_info(
            {
                'type': 'service_account',
                'client_email': env['GOOGLE_SERVICE_ACCOUNT_EMAIL'],
                'private_key': env['GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY'].replace('\\n', '\n'),
                'token_uri': 'https://oauth2.googleapis.com/token',
            },
            scopes=['https://www.googleapis.com/auth/spreadsheets'],
        )
    if not _creds.valid:
        _creds.refresh(Request())
    return _creds.token


def _api(method, path, body=None):
    req = urllib.request.Request(
        API + path, method=method,
        data=json.dumps(body).encode('utf-8') if body is not None else None,
        headers={'Authorization': 'Bearer ' + _token(), 'Content-Type': 'application/json'},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode('utf-8'))


def _col(i):
    s = ''
    i += 1
    while i > 0:
        i, rem = divmod(i - 1, 26)
        s = chr(65 + rem) + s
    return s


def _norm(txt):
    return ' '.join(str(txt).replace('\n', ' ').split()).upper()


def _num(txt):
    t = str(txt).strip().replace('.', '').replace(',', '.')
    try:
        return float(t)
    except ValueError:
        return 0.0


def _fecha_key(txt):
    try:
        return datetime.strptime(str(txt).strip(), '%d/%m/%Y')
    except ValueError:
        return datetime.max


def entregar(cod, cantidad):
    """Atajo para un solo codigo (ver entregar_varios)."""
    res = entregar_varios([{'cod': cod, 'cantidad': cantidad}])
    return {'cod': res['items'][0]['cod'], 'entregado': res['uds'],
            'filas': res['filas'], 'partida': res['items'][0]['partida']}


def entregar_varios(items):
    """Pasa a Entregado las uds pendientes de varios codigos de una vez:
    items = [{'cod': 'CF 123', 'remito': 'Remito N° 4280', 'cantidad': 10}, ...].
    Con `remito` (lo que manda el dashboard) solo se tocan las filas de ese
    remito (columna COMENTARIOS), asi una entrega no se imputa a otra venta.
    Sin `remito` (compatibilidad) se toma FIFO entre todos los remitos del codigo.
    De la fila mas vieja a la mas nueva; si la ultima no se entrega entera, la
    parte en dos: la original queda Entregado con lo entregado y se agrega una
    fila nueva Pendiente con el resto (mismo remito).
    Todo-o-nada: se valida la lista entera antes de escribir y se manda en un
    solo batchUpdate, asi un codigo invalido no deja entregas a medias."""
    pedido = {}
    for it in items or []:
        cod = _norm(it.get('cod', ''))
        # remito ausente = FIFO entre todos; remito '' = filas sin comentario
        rem = _norm(it['remito']) if it.get('remito') is not None else None
        cantidad = int(it.get('cantidad', 0))
        if not cod:
            raise ValueError('Falta el codigo')
        if cantidad <= 0:
            raise ValueError('%s: la cantidad tiene que ser mayor a 0' % cod)
        pedido[(cod, rem)] = pedido.get((cod, rem), 0) + cantidad
    if not pedido:
        raise ValueError('No hay codigos seleccionados')
    # Un codigo "sin remito" agarraria filas de los remitos pedidos aparte.
    for cod, rem in pedido:
        if rem is None and any(c == cod and r is not None for c, r in pedido):
            raise ValueError('%s: no se puede mezclar entrega con y sin remito' % cod)

    with _lock:
        # FORMATTED_VALUE: fechas/año-mes vuelven como se ven ("25/09/2026",
        # "2026-09") y se reescriben igual que las escribe remitos-app.
        filas = _api('GET', '/values/' + HOJA + '?valueRenderOption=FORMATTED_VALUE').get('values', [])
        ih = next((i for i, f in enumerate(filas[:8]) if any('CONFIRMACI' in _norm(c) for c in f)), None)
        if ih is None:
            raise RuntimeError('No encontre la columna CONFIRMACIÓN en Movimientos')
        header = [_norm(c) for c in filas[ih]]

        def idx(pred, nombre):
            for i, h in enumerate(header):
                if pred(h):
                    return i
            raise RuntimeError('No encontre la columna %s en Movimientos' % nombre)

        c_fecha = idx(lambda h: h == 'FECHA', 'FECHA')
        c_tipo  = idx(lambda h: h.startswith('TIPO'), 'TIPO MOVIMIENTO')
        c_est   = idx(lambda h: h == 'ESTADO', 'ESTADO')
        c_cod   = idx(lambda h: h.startswith('COD'), 'COD PRODUCTO')
        c_cant  = idx(lambda h: h == 'CANTIDAD', 'CANTIDAD')
        c_com   = idx(lambda h: h.startswith('COMENTARIO'), 'COMENTARIOS')
        c_am    = idx(lambda h: h.startswith('AÑO-MES'), 'AÑO-MES')
        c_conf  = idx(lambda h: h.startswith('CONFIRMACI'), 'CONFIRMACIÓN')
        c_fent  = idx(lambda h: h.startswith('FECHA ENTREGA'), 'FECHA ENTREGA (correr scripts/ampliar-rango-movimientos.ts)')

        def cel(f, c):
            return str(f[c]).strip() if c < len(f) else ''

        pend = {k: [] for k in pedido}
        for i, f in enumerate(filas[ih + 1:], start=ih + 1):
            if not (cel(f, c_tipo) == 'Salida Lita' and cel(f, c_est) == 'Embolsado'
                    and cel(f, c_conf) == 'Pendiente'):
                continue
            cod, rem = _norm(cel(f, c_cod)), _norm(cel(f, c_com))
            if (cod, rem) in pend:
                pend[(cod, rem)].append((i, f))
            elif (cod, None) in pend:
                pend[(cod, None)].append((i, f))

        # Validar todo antes de escribir nada.
        errores = []
        for (cod, rem), cantidad in pedido.items():
            total = sum(_num(cel(f, c_cant)) for _, f in pend[(cod, rem)])
            if cantidad > total:
                errores.append('%s%s tiene %d uds pendientes, no se pueden entregar %d'
                               % (cod, ' (' + rem.title() + ')' if rem else '', total, cantidad))
        if errores:
            raise ValueError('No se guardo nada. ' + '; '.join(errores))

        hoy = datetime.now().strftime('%d/%m/%Y')
        updates, nuevas, resumen, tocadas_tot = [], [], [], 0
        for (cod, rem), cantidad in pedido.items():
            lista = sorted(pend[(cod, rem)], key=lambda t: (_fecha_key(cel(t[1], c_fecha)), t[0]))
            resto, tocadas, partida = cantidad, 0, False
            for i, f in lista:
                if resto <= 0:
                    break
                tocadas += 1
                fila_hoja = i + 1
                qty = int(_num(cel(f, c_cant)))
                if qty > resto:
                    updates.append({'range': '%s!%s%d' % (HOJA, _col(c_cant), fila_hoja), 'values': [[resto]]})
                    nuevas.append({c_fecha: cel(f, c_fecha), c_tipo: 'Salida Lita', c_est: 'Embolsado',
                                   c_cod: cel(f, c_cod), c_cant: qty - resto, c_com: cel(f, c_com),
                                   c_am: cel(f, c_am), c_conf: 'Pendiente'})
                    partida = True
                updates.append({'range': '%s!%s%d' % (HOJA, _col(c_conf), fila_hoja), 'values': [['Entregado']]})
                updates.append({'range': '%s!%s%d' % (HOJA, _col(c_fent), fila_hoja), 'values': [[hoy]]})
                resto -= min(qty, resto)
            tocadas_tot += tocadas
            resumen.append({'cod': cod, 'remito': rem or '', 'entregado': cantidad, 'partida': partida})

        # Misma logica que appendRows() de remitos-app: filas exactas al final.
        # La columna A (Nº) es ARRAYFORMULA — se deja vacia a proposito.
        for k, nueva in enumerate(nuevas):
            fila_nueva = len(filas) + 1 + k
            desde = min(nueva)
            valores = [nueva.get(c, '') for c in range(desde, max(nueva) + 1)]
            updates.append({'range': '%s!%s%d' % (HOJA, _col(desde), fila_nueva), 'values': [valores]})

        _api('POST', '/values:batchUpdate', {'valueInputOption': 'USER_ENTERED', 'data': updates})
        return {'items': resumen, 'codigos': len({c for c, _ in pedido}), 'uds': sum(pedido.values()),
                'filas': tocadas_tot, 'partidas': len(nuevas)}
