"""
paginas_legales.py — genera sitio/privacidad.html y sitio/terminos.html.

Las llama generar_sitio_drive.py al final. Son páginas estáticas con el mismo
look del catálogo (Oswald + Inter, negro/naranja). Si cambia algo legal
(responsable, datos que se piden, herramientas como Analytics o el pixel de
Meta), actualizar el texto acá y la fecha ACTUALIZADO.
"""
import os

ACTUALIZADO = '21 de octubre de 2026'
EMAIL = 'contacto.carbonkey@gmail.com'
WHATSAPP = '+54 9 11 3626-7009'

_HEAD = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{titulo} — MONCAR RMH</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{--bg:#f0f0f0;--white:#fff;--black:#000;--orange:#FF5A00;--text:#111;--muted:#666;--border:#e0e0e0;
  --font-head:'Oswald',sans-serif;--font-body:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif}}
body{{background:var(--bg);color:var(--text);font-family:var(--font-body);line-height:1.65}}
header{{background:var(--black);padding:22px 16px;border-bottom:3px solid var(--orange)}}
header a{{font-family:var(--font-head);font-size:24px;font-weight:700;color:#fff;text-decoration:none;text-transform:uppercase}}
header a span{{color:var(--orange)}}
main{{max-width:780px;margin:28px auto;background:var(--white);border-radius:12px;padding:32px 24px}}
h1{{font-family:var(--font-head);font-size:30px;text-transform:uppercase;margin-bottom:4px}}
.upd{{font-size:13px;color:var(--muted);margin-bottom:24px}}
h2{{font-family:var(--font-head);font-size:19px;text-transform:uppercase;margin:26px 0 8px}}
p,li{{font-size:15px;margin-bottom:10px}}
ul{{padding-left:20px}}
a{{color:var(--orange)}}
.volver{{display:inline-block;margin-top:24px;font-weight:600}}
footer{{text-align:center;font-size:12px;color:var(--muted);padding:0 16px 28px}}
@media(max-width:600px){{main{{margin:16px 12px;padding:24px 18px}}h1{{font-size:25px}}}}
</style>
</head>
<body>
<header><a href="./">MONCAR<span> RMH</span></a></header>
<main>
<h1>{titulo}</h1>
<div class="upd">Última actualización: {actualizado}</div>
"""

_FOOT = """
<a class="volver" href="./">&#8592; Volver al catálogo</a>
</main>
<footer>© 2026 MONCAR RMH · <a href="privacidad.html">Política de privacidad</a> · <a href="terminos.html">Términos y condiciones</a></footer>
</body>
</html>
"""

PRIVACIDAD = f"""
<p>En MONCAR RMH respetamos tus datos personales. Esta política explica qué datos
recibimos a través de este sitio, para qué los usamos y cómo podés ejercer tus
derechos, de acuerdo con la Ley 25.326 de Protección de Datos Personales.</p>

<h2>Quién es el responsable</h2>
<p>MONCAR RMH, empresa familiar de venta de carbones eléctricos e interruptores para
herramientas. Contacto: <a href="mailto:{EMAIL}">{EMAIL}</a> · WhatsApp {WHATSAPP}.</p>

<h2>Qué datos recibimos</h2>
<p>Cuando querés hacernos una consulta o un pedido, te pedimos tu <strong>nombre</strong>,
tu <strong>localidad</strong> y, si querés, el <strong>nombre de tu negocio</strong>.
Este sitio <strong>no guarda esos datos en ningún servidor</strong>: con ellos se arma
un mensaje que se abre en tu WhatsApp o en tu correo, y nos llega solo si vos lo
enviás. A partir de ahí, quedan en la conversación con nuestro equipo.</p>
<p>No pedimos contraseñas, documentos ni datos de pago, y no tenemos cuentas de
usuario en este sitio.</p>

<h2>Para qué los usamos</h2>
<ul>
<li>Responder tu consulta o tu pedido.</li>
<li>Derivarte con la persona del equipo que atiende tu zona.</li>
</ul>
<p>No vendemos ni cedemos tus datos a terceros, y no te vamos a mandar publicidad sin
que nos lo pidas.</p>

<h2>Carrito y almacenamiento en tu navegador</h2>
<p>Los productos que agregás al carrito se guardan en tu propio navegador
(almacenamiento local) para que no se pierdan si recargás la página. Esa información
no nos llega; podés borrarla vaciando el carrito o los datos del sitio en tu
navegador.</p>

<h2>Cookies y seguimiento</h2>
<p>Este sitio <strong>no usa cookies de publicidad ni de análisis</strong> (no usamos
Google Analytics ni el pixel de Meta). Si en el futuro los incorporamos, te lo vamos a
pedir antes y vamos a actualizar esta política.</p>

<h2>Servicios de terceros</h2>
<p>Para funcionar, el sitio usa servicios externos que pueden recibir datos técnicos
de tu conexión (como la dirección IP), según sus propias políticas:</p>
<ul>
<li><strong>Vercel</strong>: aloja el sitio.</li>
<li><strong>Google Fonts</strong> y <strong>Cloudflare (cdnjs)</strong>: sirven las
tipografías y los íconos.</li>
<li><strong>WhatsApp (Meta)</strong> o tu proveedor de correo: por donde nos mandás el
mensaje, si elegís hacerlo.</li>
</ul>

<h2>Tus derechos</h2>
<p>Podés pedirnos en cualquier momento acceder a tus datos, corregirlos o que los
borremos (por ejemplo, la conversación con tu pedido), escribiendo a
<a href="mailto:{EMAIL}">{EMAIL}</a>. Respondemos dentro de los plazos de la ley.</p>
<p>El titular de los datos personales tiene la facultad de ejercer el derecho de
acceso a los mismos en forma gratuita a intervalos no inferiores a seis meses, salvo
que se acredite un interés legítimo al efecto conforme lo establecido en el artículo
14, inciso 3 de la Ley Nº 25.326.</p>
<p>La AGENCIA DE ACCESO A LA INFORMACIÓN PÚBLICA, en su carácter de Órgano de Control
de la Ley Nº 25.326, tiene la atribución de atender las denuncias y reclamos que
interpongan quienes resulten afectados en sus derechos por incumplimiento de las
normas vigentes en materia de protección de datos personales.</p>

<h2>Cambios</h2>
<p>Si cambiamos esta política, publicamos la versión nueva en esta página con su fecha
de actualización.</p>
"""

TERMINOS = f"""
<p>Al usar este sitio aceptás estos términos. Si tenés dudas, escribinos a
<a href="mailto:{EMAIL}">{EMAIL}</a>.</p>

<h2>Qué es este sitio</h2>
<p>Es un <strong>catálogo informativo</strong> de los productos MONCAR (carbones
eléctricos) y RMH (interruptores). No se vende ni se cobra online: el carrito sirve
para armar una consulta que nos mandás por WhatsApp o por correo. El pedido queda
confirmado recién cuando alguien de nuestro equipo te responde y acuerdan las
condiciones.</p>

<h2>Precios</h2>
<p>Los precios publicados son de lista, en pesos argentinos, y pueden cambiar sin
previo aviso. El precio final, los impuestos que correspondan, los descuentos y la
forma de pago se informan al confirmar el pedido.</p>

<h2>Disponibilidad y fotos</h2>
<p>Los productos están sujetos a disponibilidad de stock al momento de confirmar el
pedido. Las fotos son ilustrativas y pueden no coincidir exactamente con el producto
entregado; las medidas publicadas son orientativas.</p>

<h2>Marcas y compatibilidad</h2>
<p>Las marcas de herramientas y máquinas que aparecen en el sitio (por ejemplo Bosch,
Makita, DeWalt o Black &amp; Decker) pertenecen a sus respectivos dueños y se mencionan
solo para indicar compatibilidad. MONCAR RMH no está asociada ni patrocinada por esas
marcas. Ante la duda sobre qué repuesto corresponde a tu herramienta, consultanos antes
de comprar.</p>

<h2>Envíos</h2>
<p>Hacemos envíos a todo el país. El costo, la forma y el plazo de entrega se coordinan
al confirmar el pedido.</p>

<h2>Garantía y reclamos</h2>
<p>Si un producto llega fallado o no es el que pediste, contactanos y lo resolvemos.
Cuando la compra la hace un consumidor final, rigen además los derechos de la Ley
24.240 de Defensa del Consumidor.</p>

<h2>Contenido del sitio</h2>
<p>Los textos, fotos y diseño del sitio son de MONCAR RMH, salvo los elementos de
terceros que usamos bajo sus licencias: las tipografías Oswald e Inter (SIL Open Font
License), los íconos de Font Awesome Free, Material Symbols y Material Design Icons, y el ícono de
máquina de coser de Caro Asercion
(<a href="https://game-icons.net" target="_blank" rel="noopener">game-icons.net</a>),
bajo licencia
<a href="https://creativecommons.org/licenses/by/3.0/" target="_blank" rel="noopener">CC BY 3.0</a>.</p>

<h2>Datos personales</h2>
<p>Cómo tratamos los datos que nos mandás está explicado en la
<a href="privacidad.html">Política de privacidad</a>.</p>

<h2>Ley aplicable</h2>
<p>Estos términos se rigen por las leyes de la República Argentina.</p>
"""


def generar_paginas_legales(out_dir):
    paginas = {
        'privacidad.html': ('Política de privacidad', PRIVACIDAD),
        'terminos.html': ('Términos y condiciones', TERMINOS),
    }
    for archivo, (titulo, cuerpo) in paginas.items():
        html = _HEAD.format(titulo=titulo, actualizado=ACTUALIZADO) + cuerpo + _FOOT
        with open(os.path.join(out_dir, archivo), 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"Pagina legal generada: {archivo}")
