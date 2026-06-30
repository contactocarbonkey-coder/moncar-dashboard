"""
Reemplaza el conditional formatting en namespace x14 (ignorado por Excel)
por el namespace estándar en el XML interno del STOCK_MONCAR.xlsx.
Los DXFs ya existen en styles.xml: dxfId 0=rojo, 1=naranja, 2=verde.
"""
import zipfile, shutil, re, os

PATH  = r'C:\Users\franc\OneDrive\Documentos\KeyCarbon\claude MONCAR RMH\STOCK_MONCAR.xlsx'
TEMP  = PATH + '.tmp'

# ── Construcción de fórmulas ──────────────────────────────────────────────────
def vlk(col):
    return f'IFERROR(VLOOKUP(E4,Stock!$A:$J,{col},0),0)'

NO_EMB     = vlk(3)
EMB_DEP    = vlk(4)
PREM_DEP   = vlk(5)
JUANI      = vlk(6)
STOCK_PACK = f'IF(D4="Premium",{PREM_DEP},{EMB_DEP})'

def orc4(*tipos):
    return 'OR(' + ','.join(f'C4="{t}"' for t in tipos) + ')'

IS_PACK  = orc4("Salida Venta Franco", "Salida Lita", "Entrega a Juani")
IS_JUANI = 'C4="Salida Venta Juani"'
IS_ALL   = orc4("Salida Venta Franco", "Salida Lita", "Entrega a Juani", "Salida Venta Juani")
HAS      = 'F4&lt;&gt;"",E4&lt;&gt;""'    # XML-encoded <>

# Fórmulas con operadores XML-encoded (< → &lt;  >= → &gt;=  etc.)
# El AND/OR anida IF: la lógica está en el cálculo del stock post-movimiento.
# Stock ya incluye la fila actual, así que: stock<0 → insuficiente.

def encode(s):
    """Codifica < > para XML, pero deja comillas literales como están."""
    return (s
        .replace('>=', '__GTE__')
        .replace('<=', '__LTE__')
        .replace('<>', '__NEQ__')
        .replace('>', '&gt;')
        .replace('<', '&lt;')
        .replace('__GTE__', '&gt;=')
        .replace('__LTE__', '&lt;=')
        .replace('__NEQ__', '&lt;&gt;')
    )

f_green  = encode(
    f'AND({IS_ALL},F4<>"",E4<>"",IF({IS_JUANI},{JUANI},{STOCK_PACK})>=0)'
)
f_orange = encode(
    f'AND({IS_PACK},F4<>"",E4<>"",{STOCK_PACK}<0,{NO_EMB}>=(-({STOCK_PACK})))'
)
f_red    = encode(
    f'AND({IS_ALL},F4<>"",E4<>"",OR('
    f'AND({IS_JUANI},{JUANI}<0),'
    f'AND({IS_PACK},{STOCK_PACK}<0,{NO_EMB}<(-({STOCK_PACK})))'
    f'))'
)

# ── XML del conditional formatting estándar ───────────────────────────────────
# dxfId 0=rojo, 1=naranja, 2=verde  (ya están en styles.xml)
CF_XML = (
    '<conditionalFormatting sqref="F4:F1003">'
    f'<cfRule type="expression" priority="1" dxfId="2"><formula>{f_green}</formula></cfRule>'
    f'<cfRule type="expression" priority="2" dxfId="1"><formula>{f_orange}</formula></cfRule>'
    f'<cfRule type="expression" priority="3" dxfId="0"><formula>{f_red}</formula></cfRule>'
    '</conditionalFormatting>'
)

# ── Procesar el xlsx como ZIP ────────────────────────────────────────────────
shutil.copy2(PATH, PATH + '.bak2')

with zipfile.ZipFile(PATH, 'r') as zin:
    infos = zin.infolist()
    files = {item.filename: (zin.read(item.filename), item.compress_type) for item in infos}

# Modificar sheet2.xml (Movimientos)
raw, ctype = files['xl/worksheets/sheet2.xml']
sheet_xml = raw.decode('utf-8')

# 1. Eliminar el bloque x14 del extLst
sheet_xml = re.sub(
    r'<ext\b[^>]*>\s*<x14:conditionalFormattings\b.*?</x14:conditionalFormattings>\s*</ext>',
    '', sheet_xml, flags=re.DOTALL
)

# 2. Eliminar cualquier CF estándar previo para col F (por si hay duplicados)
sheet_xml = re.sub(
    r'<conditionalFormatting sqref="F[^"]*">.*?</conditionalFormatting>',
    '', sheet_xml, flags=re.DOTALL
)

# 3. Insertar el CF estándar antes de <extLst> o antes de </worksheet>
if '<extLst>' in sheet_xml:
    sheet_xml = sheet_xml.replace('<extLst>', CF_XML + '<extLst>', 1)
else:
    sheet_xml = sheet_xml.replace('</worksheet>', CF_XML + '</worksheet>', 1)

files['xl/worksheets/sheet2.xml'] = (sheet_xml.encode('utf-8'), ctype)

# ── Escribir el xlsx final ────────────────────────────────────────────────────
with zipfile.ZipFile(TEMP, 'w') as zout:
    for name, (data, compress) in files.items():
        zout.writestr(name, data, compress_type=compress)

os.replace(TEMP, PATH)
print("OK: conditional formatting reescrito en namespace estandar.")
print()

# ── Verificar en el XML resultante ───────────────────────────────────────────
with zipfile.ZipFile(PATH, 'r') as z:
    xml = z.read('xl/worksheets/sheet2.xml').decode('utf-8')

has_standard = '<conditionalFormatting sqref="F4:F1003">' in xml
has_x14      = 'x14:conditionalFormattings' in xml
print(f"CF estandar presente: {has_standard}")
print(f"CF x14 (ext) presente: {has_x14}")
if has_standard:
    idx = xml.find('<conditionalFormatting sqref="F4:F1003">')
    print()
    print("Fragmento CF estandar:")
    print(xml[idx:idx+300])
