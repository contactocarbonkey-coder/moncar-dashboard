# -*- coding: utf-8 -*-
# Registra prestamos de mercaderia entre zonas en Movimientos (STOCK_MONCAR y
# STOCK_RMH). Lo usa servidor_drive.py → POST /prestamos/registrar, desde la
# pagina "Prestamos" del dashboard.
#
# Modelo (2026-10-06): una fila por codigo, con los TIPO MOVIMIENTO que ya
# existian y que las formulas de Stock ya suman/restan segun ESTADO:
#   recibido   → "Préstamo recibido"             (entra stock; le debemos a PERSONA)
#   devolucion → "Devolución préstamo recibido"  (sale stock; cancela deuda)
#   dado       → "Préstamo dado"                 (sale stock; PERSONA nos debe)
#   recupero   → "Recupero préstamo dado"        (entra stock; cancela lo que nos deben)
# Columnas nuevas: PERSONA, SALDA COD, SALDA CANT (en una devolucion/recupero:
# que deuda cancela — se puede devolver con otro codigo y otra cantidad).
# Las formulas de Stock NO se tocaron. Columna A (Nº) y las automaticas no se
# escriben: se escribe campo por campo, igual que remitos-app.
import json, threading, urllib.request
from datetime import datetime
import entregas_lita  # reusa la credencial (_token) — no se modifica ese modulo

SHEETS = {
    'MONCAR': '1u9xG2uyeT3V8osBi9Qz0xofDCPmFmhOXlu85c1OIEQ8',
    'RMH':    '1Msl0CC5abCoIaM0MpeCtYOoHVHI0mjPlfElc5x8ytJE',
}
ESTADOS = {'MONCAR': ['No embolsado', 'Embolsado', 'Premium'], 'RMH': ['No embolsado', 'Embolsado']}
PERSONAS = ['Luvi', 'Lita', 'Naza', 'Mapi', 'Fede']
OPS = {
    'recibido':   'Préstamo recibido',
    'devolucion': 'Devolución préstamo recibido',
    'dado':       'Préstamo dado',
    'recupero':   'Recupero préstamo dado',
}
# devolucion cancela lo de recibido; recupero cancela lo de dado
CANCELA = {'devolucion': 'recibido', 'recupero': 'dado'}
HOJA = 'Movimientos'
_lock = threading.Lock()


def _api(linea, method, path, body=None):
    req = urllib.request.Request(
        'https://sheets.googleapis.com/v4/spreadsheets/' + SHEETS[linea] + path, method=method,
        data=json.dumps(body).encode('utf-8') if body is not None else None,
        headers={'Authorization': 'Bearer ' + entregas_lita._token(), 'Content-Type': 'application/json'},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode('utf-8'))


_norm = entregas_lita._norm
_col  = entregas_lita._col


def _cod(txt):
    return ' '.join(str(txt).split()).upper()


def _leer(linea):
    """Filas de Movimientos + indices de columna (por nombre) + codigos validos de Stock."""
    filas = _api(linea, 'GET', '/values/' + HOJA + '?valueRenderOption=FORMATTED_VALUE').get('values', [])
    ih = next((i for i, f in enumerate(filas[:8]) if 'COMENTARIOS' in [_norm(c) for c in f]), None)
    if ih is None:
        raise RuntimeError('No encontre el header de Movimientos en ' + linea)
    header = [_norm(c) for c in filas[ih]]
    cols = {}
    for n in ('FECHA', 'TIPO MOVIMIENTO', 'ESTADO', 'COD PRODUCTO', 'CANTIDAD', 'COMENTARIOS',
              'PERSONA', 'SALDA COD', 'SALDA CANT'):
        if n not in header:
            raise RuntimeError('Falta la columna %s en Movimientos de %s (correr remitos-app/scripts/preparar-prestamos.ts)' % (n, linea))
        cols[n] = header.index(n)
    cols['AÑO-MES'] = next(i for i, h in enumerate(header) if h.startswith('AÑO-MES'))
    stock = _api(linea, 'GET', '/values/Stock!A:A').get('values', [])
    codigos = {_cod(f[0]) for f in stock if f and _cod(f[0])[:2] in ('CF', 'CC', 'CE', 'RM')}
    return filas, ih, cols, codigos


def _saldos(datos, persona, op_origen):
    """Saldo abierto por codigo (en las 2 lineas) de `persona` para recibido o dado."""
    tipo_origen, tipo_cancela = OPS[op_origen], OPS[[k for k, v in CANCELA.items() if v == op_origen][0]]
    saldo = {}
    for filas, ih, c, _ in datos.values():
        for f in filas[ih + 1:]:
            g = lambda n: str(f[c[n]]).strip() if c[n] < len(f) else ''
            if g('PERSONA') != persona:
                continue
            if g('TIPO MOVIMIENTO') == tipo_origen:
                k = _cod(g('COD PRODUCTO'))
                saldo[k] = saldo.get(k, 0) + entregas_lita._num(g('CANTIDAD'))
            elif g('TIPO MOVIMIENTO') == tipo_cancela:
                k = _cod(g('SALDA COD') or g('COD PRODUCTO'))
                saldo[k] = saldo.get(k, 0) - entregas_lita._num(g('SALDA CANT') or g('CANTIDAD'))
    return saldo


def registrar(op, persona, fecha, items, comentario=''):
    """items: [{linea, cod, cantidad, estado, salda_cod?, salda_cant?}] — una fila por item.
    fecha: 'yyyy-mm-dd'. Valida todo antes de escribir; si algo falla no escribe nada."""
    if op not in OPS:
        raise ValueError('Operacion invalida')
    if persona not in PERSONAS:
        raise ValueError('Persona invalida: elegi una de ' + ', '.join(PERSONAS))
    try:
        dt = datetime.strptime(fecha, '%Y-%m-%d')
    except (TypeError, ValueError):
        raise ValueError('Fecha invalida')
    if dt.date() > datetime.now().date():
        raise ValueError('La fecha no puede ser futura')
    if not items:
        raise ValueError('Agrega al menos un codigo')
    comentario = str(comentario or '').strip()[:200]

    with _lock:
        lineas = sorted({str(it.get('linea')) for it in items} | ({'MONCAR', 'RMH'} if op in CANCELA else set()))
        if any(l not in SHEETS for l in lineas):
            raise ValueError('Linea invalida')
        datos = {l: _leer(l) for l in lineas}
        saldo = _saldos(datos, persona, CANCELA[op]) if op in CANCELA else None

        limpios = []
        for it in items:
            linea = it['linea']
            cod = _cod(it.get('cod', ''))
            cant = int(it.get('cantidad') or 0)
            estado = it.get('estado')
            if cod not in datos[linea][3]:
                raise ValueError('El codigo %s no existe en %s' % (cod or '(vacio)', linea))
            if cant <= 0:
                raise ValueError('Cantidad invalida en ' + cod)
            if estado not in ESTADOS[linea]:
                raise ValueError('Estado invalido en %s (para %s: %s)' % (cod, linea, ', '.join(ESTADOS[linea])))
            fila = {'COD PRODUCTO': cod, 'CANTIDAD': cant, 'ESTADO': estado}
            if op in CANCELA:
                s_cod = _cod(it.get('salda_cod') or cod)
                s_cant = int(it.get('salda_cant') or cant)
                abierto = saldo.get(s_cod, 0)
                if s_cant <= 0 or s_cant > abierto:
                    raise ValueError('%s: %s tiene %d uds abiertas de %s, no se pueden saldar %d'
                                     % (OPS[op], persona, abierto, s_cod, s_cant))
                saldo[s_cod] = abierto - s_cant
                # Solo se completa SALDA si difiere de lo entregado (si no, vacio = mismo)
                if s_cod != cod or s_cant != cant:
                    fila['SALDA COD'], fila['SALDA CANT'] = s_cod, s_cant
            limpios.append((linea, fila))

        fecha_txt, anio_mes = dt.strftime('%d/%m/%Y'), dt.strftime('%Y-%m')
        por_linea = {}
        for linea, fila in limpios:
            por_linea.setdefault(linea, []).append(fila)
        for linea, filas_nuevas in por_linea.items():
            filas, ih, c, _ = datos[linea]
            proxima = len(filas) + 1
            data = []
            for k, fila in enumerate(filas_nuevas):
                valores = dict(fila, **{'FECHA': fecha_txt, 'TIPO MOVIMIENTO': OPS[op], 'PERSONA': persona,
                                        'COMENTARIOS': comentario, 'AÑO-MES': anio_mes})
                for nombre, v in valores.items():
                    if v == '':
                        continue
                    data.append({'range': '%s!%s%d' % (HOJA, _col(c[nombre]), proxima + k), 'values': [[v]]})
            _api(linea, 'POST', '/values:batchUpdate', {'valueInputOption': 'USER_ENTERED', 'data': data})
        return {'op': op, 'persona': persona, 'filas': len(limpios)}
