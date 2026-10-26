# -*- coding: utf-8 -*-
"""
extraer_imagenes_catalogo.py — extrae las fotos de productos del PDF
"CATALOGO CARBONES MONCAR (1) (2).pdf" y las guarda en sitio/imagenes/,
nombradas por codigo (ej. "CF 590" -> imagenes/CF_590.jpeg).

Como funciona: el PDF tiene un layout de grilla donde cada celda es
COD (texto) + foto (imagen embebida) + MEDIDA/DESCRIPCION (texto, en una
columna aparte). La foto en si no tiene texto superpuesto, asi que
extraer el objeto de imagen embebido tal cual (sin recortar) ya da
exactamente lo que se necesita. Para saber a que codigo corresponde cada
imagen, se busca el texto "CF"/"CC"/"CE" + numero mas cercano arriba de
la imagen, con solape horizontal (mismo layout que ya vimos: el codigo
queda siempre justo arriba de su foto, en la misma columna).

Requiere: pip install pymupdf

Correr de nuevo solo si Franco actualiza el PDF del catalogo con
productos nuevos — no hace falta correrlo en cada generacion del sitio,
generar_sitio_drive.py simplemente lee lo que ya haya en sitio/imagenes/.
"""
import fitz, re, os

DIR      = os.path.dirname(os.path.abspath(__file__))
PDF_PATH = os.path.join(DIR, 'CATALOGO CARBONES MONCAR (1) (2).pdf')
OUT_DIR  = os.path.join(DIR, 'sitio', 'imagenes')

COD_RE = re.compile(r'^(CF|CC|CE|RM)$', re.IGNORECASE)


def get_cod_anchors(words):
    """Combina palabras tipo 'CF' + '990' en un mismo renglon en anchors
    con su posicion, para despues asociarlas a la imagen que tienen debajo."""
    anchors = []
    for i, w in enumerate(words):
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        if COD_RE.match(text):
            for w2 in words[i + 1:i + 3]:
                x0b, y0b, x1b, y1b, text2 = w2[0], w2[1], w2[2], w2[3], w2[4]
                if abs(y0b - y0) < 3 and re.match(r'^\d+[A-Za-z]?$', text2):
                    cod = f"{text.upper()} {text2}"
                    anchors.append({'cod': cod, 'x0': x0, 'x1': x1b, 'y0': y0, 'y1': max(y1, y1b)})
                    break
    return anchors


def main():
    if not os.path.exists(PDF_PATH):
        print(f"ERROR: no se encontro '{PDF_PATH}'")
        return

    os.makedirs(OUT_DIR, exist_ok=True)
    doc = fitz.open(PDF_PATH)

    resultados = {}       # cod -> {xref}
    sin_matchear = []
    ya_asignados_dup = []

    for pno in range(len(doc)):
        p = doc[pno]
        words = p.get_text('words')
        anchors = get_cod_anchors(words)
        for im in p.get_image_info(xrefs=True):
            bx0, by0, bx1, by1 = im['bbox']
            w, h = bx1 - bx0, by1 - by0
            if w > 300 and h < 60:
                continue  # banner/logo de encabezado, no es foto de producto

            best, best_gap = None, None
            for a in anchors:
                if a['y1'] <= by0 + 2:
                    ox0 = max(a['x0'], bx0)
                    ox1 = min(a['x1'] + 40, bx1 + 40)
                    if ox0 <= ox1:
                        gap = by0 - a['y1']
                        if gap < 60 and (best_gap is None or gap < best_gap):
                            best, best_gap = a, gap

            if not best:
                sin_matchear.append({'pagina': pno + 1, 'bbox': im['bbox']})
                continue
            cod = best['cod']
            if cod in resultados:
                ya_asignados_dup.append({'pagina': pno + 1, 'cod': cod})
                continue
            resultados[cod] = im['xref']

    guardadas = 0
    for cod, xref in resultados.items():
        base = doc.extract_image(xref)
        fname = cod.replace(' ', '_') + '.' + base['ext']
        with open(os.path.join(OUT_DIR, fname), 'wb') as f:
            f.write(base['image'])
        guardadas += 1

    print(f"Imagenes guardadas: {guardadas} en {OUT_DIR}")
    if sin_matchear:
        print(f"\nAVISO: {len(sin_matchear)} imagenes sin poder asociar a un codigo (revisar manualmente):")
        for u in sin_matchear:
            print(f"   pagina {u['pagina']}, bbox {[round(v,1) for v in u['bbox']]}")
    if ya_asignados_dup:
        print(f"\nAVISO: {len(ya_asignados_dup)} imagenes con codigo repetido (posible variante sin distinguir, ej. 'CF 635' vs 'CF 635 T11'):")
        for d in ya_asignados_dup:
            print(f"   pagina {d['pagina']}, codigo {d['cod']}")


if __name__ == '__main__':
    main()
