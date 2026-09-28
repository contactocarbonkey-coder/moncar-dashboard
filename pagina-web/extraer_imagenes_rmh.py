# -*- coding: utf-8 -*-
"""
extraer_imagenes_rmh.py — extrae las fotos de interruptores del PDF
"LISTA DE PRECIO RMH AGOSTO 2026.pdf" y las guarda en sitio/imagenes/,
nombradas por codigo (ej. "RM 022" -> imagenes/RM_022.jpeg).

Layout distinto al catalogo de MONCAR (ver extraer_imagenes_catalogo.py):
ac ac es una tabla con columnas CODIGO | FOTO | APTO PARA | MODELOS | PRECIO,
una fila por producto. El codigo queda a la izquierda y su foto al lado
(misma fila, no debajo). Algunos productos tienen 2 fotos en la fila (dos
angulos); se usa la de mas a la izquierda como foto principal.

Requiere: pip install pymupdf

Correr de nuevo solo si Franco actualiza este PDF con productos RMH nuevos.
"""
import fitz, re, os

DIR      = os.path.dirname(os.path.abspath(__file__))
PDF_PATH = os.path.join(DIR, 'LISTA DE PRECIO RMH  AGOSTO 2026.pdf')
OUT_DIR  = os.path.join(DIR, 'sitio', 'imagenes')


def cod_anchors(words):
    anchors = []
    for i, w in enumerate(words):
        if w[4] == 'RM' and i + 1 < len(words):
            nxt = words[i + 1]
            if re.match(r'^\d+[A-Za-z]?$', nxt[4]) and abs(nxt[1] - w[1]) < 3:
                anchors.append({'cod': f'RM {nxt[4]}', 'y0': w[1]})
    return anchors


def main():
    if not os.path.exists(PDF_PATH):
        print(f"ERROR: no se encontro '{PDF_PATH}'")
        return

    os.makedirs(OUT_DIR, exist_ok=True)
    doc = fitz.open(PDF_PATH)

    resultado = {}
    sin_match = []

    for pno in range(len(doc)):
        p = doc[pno]
        words = p.get_text('words')
        anchors = sorted(cod_anchors(words), key=lambda a: a['y0'])
        if not anchors:
            continue

        imgs = [im for im in p.get_image_info(xrefs=True)
                if im['bbox'][2] - im['bbox'][0] >= 50]  # descarta el logo chico del encabezado

        foto_y1 = 0
        for w in words:
            if w[4].upper() == 'FOTO':
                foto_y1 = max(foto_y1, w[3])
        imgs = [im for im in imgs if im['bbox'][1] >= foto_y1 - 2]

        for idx, a in enumerate(anchors):
            y_ini = a['y0'] - 15
            y_fin = anchors[idx + 1]['y0'] - 10 if idx + 1 < len(anchors) else 99999
            candidatos = [im for im in imgs if y_ini <= (im['bbox'][1] + im['bbox'][3]) / 2 < y_fin]
            if not candidatos:
                sin_match.append((pno + 1, a['cod']))
                continue
            candidatos.sort(key=lambda im: im['bbox'][0])  # la mas a la izquierda = foto principal
            resultado[a['cod']] = candidatos[0]['xref']

    guardadas = 0
    for cod, xref in resultado.items():
        base = doc.extract_image(xref)
        fname = cod.replace(' ', '_') + '.' + base['ext']
        with open(os.path.join(OUT_DIR, fname), 'wb') as f:
            f.write(base['image'])
        guardadas += 1

    print(f"Imagenes RMH guardadas: {guardadas} en {OUT_DIR}")
    if sin_match:
        print(f"\nAVISO: {len(sin_match)} codigos sin foto asociada:")
        for pagina, cod in sin_match:
            print(f"   pagina {pagina}, codigo {cod}")


if __name__ == '__main__':
    main()
