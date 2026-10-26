# Página web / catálogo MONCAR RMH

**Reactivado el 2026-09-22** (había quedado pausado el 2026-09-10, ver `../../HISTORIAL.md`).
**En vivo desde el 2026-09-28** en `https://moncar-rmh.com.ar`.

**Ubicación (desde 2026-10-26):** `proyectos/moncar/moncar-web/pagina-web/`. Al lado está
`../backend-web/` (el panel, repo `moncar-panel`), que va a ser el backend de esta web:
el objetivo es que el generador lea productos, categorías y textos de la base del panel
(todavía no está conectado). El generador se versiona en el repo raíz (`moncar-dashboard`).

## Qué es

Catálogo público de productos MONCAR (carbones) y RMH (interruptores), pensado como
**carta de presentación** para mandar a clientes y potenciales clientes — no es una
tienda online con cobro integrado. Muestra precios de lista con fotos reales de cada
producto, permite armar un pedido tipo carrito (con selector de cantidad) y lo manda
por WhatsApp o Email para cerrar la venta ahí. Diseño inspirado en la web de Dowen
Paggio (tipografía Oswald, paleta negro/gris/naranja).

## Cómo se genera

```
python generar_sitio_drive.py   →   sitio/index.html
```

Lee precios/stock en vivo de los mismos 3 Google Sheets que usa el Dashboard (ver
`../../CLAUDE.md` para los IDs — son los mismos). Las dimensiones de los carbones MONCAR
(espesor/ancho/largo/terminal) todavía salen de `../../LP/LP 10 CARBONES 01.04.2026.xlsx`
porque ese detalle no está migrado a Sheets (pendiente general del proyecto); si ese
archivo no está o cambia de formato, el sitio se genera igual, solo sin esos filtros.

`generar_sitio.py` (sin sufijo `_drive`) es la versión vieja, lee de Excel locales que
ya no existen en las rutas que espera — **no usar**, quedó solo de referencia histórica.

### Fotos de producto

Extraídas de los catálogos en PDF (no vienen de los Sheets):

- `extraer_imagenes_catalogo.py` → lee `CATALOGO CARBONES MONCAR (1) (2).pdf` → 322/340
  fotos MONCAR matcheadas por posición (código arriba, foto debajo, en la misma columna).
- `extraer_imagenes_rmh.py` → lee `LISTA DE PRECIO RMH AGOSTO 2026.pdf` → 39/39 fotos
  RMH matcheadas (layout de tabla: código a la izquierda, foto al lado en la misma fila).

Ambos guardan en `sitio/imagenes/<COD_con_guion_bajo>.<ext>`. `generar_sitio_drive.py`
lee ese directorio en tiempo de generación (`leer_imagenes()`) — no hace falta volver a
correr los extractores salvo que Franco actualice alguno de los dos PDF con productos
nuevos. Productos sin foto (los ~18 MONCAR que no matchearon — ver aviso al correr el
script) usan el ícono SVG genérico de respaldo (`SVG_MC`/`SVG_RMH` en el JS).

## Deploy

`sitio/` es un **repo git propio** (rama `main`), separado del repo de este proyecto —
ver `.gitignore` del proyecto raíz (`moncar-web/pagina-web/sitio/` está excluido a propósito,
mismo criterio que `remitos-app/`). Conectado a
`github.com/contactocarbonkey-coder/moncar-rmh-web` (cuenta de `contacto.carbonkey@gmail.com`
= Andi, no la de Franco) → Vercel detecta cada push a `main` y redespliega solo, sin
pasos manuales.

**Para publicar un cambio**: correr `generar_sitio_drive.py`, después desde `sitio/`:
```
git add -A
git commit -m "..."
git push
```
(`desplegar.bat` automatiza esto, pero asumía rutas viejas — revisar antes de confiar
en él a ciegas, mejor los comandos de arriba a mano por ahora).

`sitio/.gitignore` excluye `dashboard_live.html` (un archivo del dashboard interno que
había quedado pegado ahí por error en junio, con datos de ventas/stock — nunca debe
subirse a este repo público).

### Dominio propio — moncar-rmh.com.ar (en vivo desde 2026-09-28)

Cadena completa: **nic.ar** (registro, a nombre de Maria Julieta Montero / Juli, CUIT
27275034237) → delegado a nameservers de **Cloudflare** (`melina.ns.cloudflare.com`,
`thomas.ns.cloudflare.com`) → Cloudflare tiene los registros DNS (**A** `@` →
`216.198.79.1`, **CNAME** `www` → `aa21ec42206176ad.vercel-dns-017.com`, ambos en modo
**"DNS only"**, sin proxy — el proxy de Cloudflare rompe el SSL automático de Vercel) →
**Vercel** sirve el sitio real. Cloudflare acá **solo hace de DNS**, no aloja nada — se
usó porque el panel de nic.ar no tiene forma de cargar registros A/CNAME directamente,
solo delegación de nameservers.

**Gotcha importante de nic.ar**: después de agregar los nameservers en
"Delegaciones", hay un botón separado **"EJECUTAR CAMBIOS"** que hay que apretar para
que se aplique de verdad al registro. La primera vez (24/09) no se apretó — quedó
"guardado" en la UI pero el WHOIS oficial (`whois.nic.ar` puerto 43) no tenía ningún
`nserver` cargado, así que estuvo **3 días sin funcionar** sin ningún error visible. Se
detectó consultando el WHOIS directo (no confiar solo en lo que muestra el panel), y se
resolvió re-cargando los nameservers y esta vez sí ejecutando el cambio.

**Corrección sobre el costo**: **el `.com.ar` NO fue gratis** — a diferencia de lo que
se asumió al principio, el registro en el circuito oficial de nic.ar
(tramitesadistancia.gob.ar) pidió pago igual (monto menor, no confirmado exacto). Ojo
con esto para el próximo dominio: no asumir que es gratis solo por ser `.com.ar`.

### Plan de Vercel — Hobby por ahora, no Pro

Decisión de Franco (2026-09-24): usar el plan **Hobby (gratis)** porque por ahora es
"uso personal, para probar la funcionalidad" — no lo está mandando todavía como
material de venta real a clientes. **Ojo**: en cuanto esto se use para mostrarle el
catálogo a clientes reales, técnicamente pasa a ser uso comercial según los términos de
Vercel (aunque no cobre online — "advertising the sale of a product or service" ya
cuenta) y en ese momento sí hay que pasar a **Pro (USD 20/mes)**, o Vercel puede pausar
el sitio sin aviso. No pasar esto por alto cuando Franco decida empezar a usarlo en
serio.

## Diseño

**v3 en producción desde 2026-10-26** (plantilla en `generar_html()` de `generar_sitio_drive.py`,
CSS y JS en los strings crudos `_CSS` / `_JS`, datos por marcadores `__PRODUCTOS__` etc.). Es el
diseño de la v1 (Oswald, negro/naranja, logo) con la usabilidad de la v2:
- Banner: logo + buscador guiado (por marca / por tipo de máquina / por medidas / por modelo).
- "Nuestras líneas": MONCAR = carbones eléctricos, RMH = interruptores eléctricos. Marca, categoría y
  medidas de la portada son **de carbones** (fijan la línea MONCAR).
- Menú "Productos" desplegable (hover en compu, toque en celular) con acceso directo a cada línea;
  selector de línea grande arriba del catálogo; con RMH se ocultan medidas/marca/terminal.
- Búsqueda por medidas con tolerancia (±0,3 mm espesor/ancho, ±1,5 mm largo, acepta E/A invertidos);
  imagen `sitio/medidas-carbon.webp`. El código NO es la búsqueda principal (decisión de Franco: el
  ferretero no lo sabe).
- Grilla: 5 por fila en compu (4 en pantallas medianas), **2 por fila en celular con la medida visible**.
- Al agregar no se abre el carrito: aviso + barra de pedido fija en el celular.
- Rutas por hash: `#productos/linea/MONCAR|RMH`, `/cat/<cat>`, `/marca/<marca>`, `/medidas/6.3x10x`.

Detalle histórico de la v1 (sigue valiendo para íconos y grillas fijas):

- Tipografía **Oswald** (Google Fonts) para títulos/nav/botones/categorías, **Inter**
  para texto de cuerpo.
- Paleta: negro (`--black`/`--black2`), grises, blanco, naranja (`--orange: #FF5A00`).
- Logo (`imagenes/logo moncar rmh.png` en la raíz del proyecto es el original pesado —
  no se sube a `sitio/`; se generaron 3 versiones optimizadas dentro de `sitio/`:
  `logo-nav.png` 160px sin uso actual, `logo-hero.png` 640px para el banner principal,
  `logo-footer.png` 260px para el footer). El nav actual es **texto** "MONCAR RMH" (sin
  logo, a pedido de Franco — se probó con logo y se volvió atrás).
- Íconos de categoría: la mayoría son **SVG reales** (Material Symbols / MDI /
  Game-Icons vía Iconify, `fill="currentColor"` embebidos como constantes
  `_SVG_*` en `generar_sitio_drive.py` — no se descargan en runtime) — taladro, sierra
  circular, máquina de coser, aspiradora, lavarropas, bomba de agua (proxy de
  hidrolavadora, no existe ícono especifico en ninguna libreria libre revisada).
  Amoladoras/Electrodomesticos/Otros usan Font Awesome (`fa-gear`/`fa-plug`/`fa-toolbox`)
  porque **no existe ningún ícono de amoladora/angle-grinder en ninguna librería libre**
  (se buscó explícitamente: "angle grinder", "grinding wheel", "disc sander" → 0
  resultados en Iconify). Los íconos de "Nosotros" (contacto y las 4 tarjetitas de
  "experiencia/multimarca/envíos/atención") usan el mismo estilo Font Awesome naranja.
- Grilla de "Compra por marca" (6 columnas, 12 marcas) y "Compra por categoría" (5
  columnas, 10 categorías) son **grillas fijas**, no responsive-inteligentes — se probó
  una versión con flexbox que se auto-ajustaba sin dejar huecos, pero a Franco no le
  gustó el resultado visual y pidió volver a columnas fijas. **Si el día de mañana
  cambia la cantidad de marcas destacadas o categorías, hay que revisar/ajustar las
  columnas a mano** (`repeat(6,1fr)` / `repeat(5,1fr)` en el CSS).
- Categoría "Equipos de Pintura" se eliminó (solo tenía 2 productos) — esos productos
  ahora caen en "Otros" vía `detectar_categoria()`.

## Funcionalidad

- **Stock: pausado/oculto a pedido de Franco (2026-09-24)**. El dato se sigue leyendo
  de Sheets y guardando en el JSON de cada producto (`stock: true/false`), pero no se
  muestra en ningún lado del sitio ni se usa para deshabilitar el botón "Agregar" — se
  puede pedir cualquier producto esté o no en stock. Reactivar es cuestión de volver a
  usar ese campo en el JS (`renderGrid`, filtro sidebar) si Franco lo pide.
- **Orden por defecto**: alfabético por marca, después por código (estable). No hay
  ordenamiento por stock (se sacó junto con el punto anterior).
- **Selector de cantidad** (+/-) en cada card de producto antes de "Agregar".
- **Paginación "Cargar más"**: la grilla muestra 40 productos por vez (`PAG_SIZE`),
  con botón para traer más — antes renderizaba los ~380 productos de una, lento en
  celulares con las fotos reales ya cargadas.
- **Ruteo liviano por URL** (`aplicarHash()`/`setHash()` en el JS): `#inicio`,
  `#productos`, `#productos/cat/<cat>`, `#productos/marca/<marca>`, `#nosotros`. Permite
  compartir un link directo a una categoría/marca y que el botón atrás del navegador
  funcione. **No es SSR real** — sigue siendo una sola página, Google no puede indexar
  productos individuales por esta via. Si en algún momento se necesita SEO real por
  producto, hay que rearmar el sitio con un framework con server-side rendering (Next.js
  o similar) — evaluado y descartado por ahora, demasiado esfuerzo para una carta de
  presentación.

## WhatsApp con identificación por zona (pendiente de completar)

**Decisión actualizada (2026-09-25)**: se revirtió la decisión original de "sin ruteo
por zona" del 2026-09-22. El negocio sí se maneja por zonas geográficas con distintos
vendedores, así que **todo botón de WhatsApp que inicia una conversación** (el del nav,
"Consultar por WhatsApp" del hero, "Pedir por WhatsApp" del carrito) pasa siempre por un
formulario (nombre, negocio opcional, localidad) antes de abrir WhatsApp —
`abrirConsultaWA()` / `pedirPorWA()` → `abrirFormModal()` → `confirmarPedido()` →
`elegirDestino(localidad)`.

Se encontraron los **números reales del equipo** en el encabezado de
`LISTA DE PRECIO RMH AGOSTO 2026.pdf` (cargados en la constante `CONTACTOS` del JS):

| Persona | WhatsApp |
|---|---|
| Franco Ahmad | 5491136267009 |
| Andrés Castagna (Andi) | 5491139219269 |
| David Wu | 5491150225627 |
| Maximiliano Ojeda (Maxi) | 5491144712072 |

**Falta el mapa de qué zona/localidad le corresponde a cada uno** — `elegirDestino()`
hoy devuelve siempre `CONTACTOS.FRANCO` sin importar la localidad que cargue el
cliente. Cuando Franco pase la lista de zonas por persona, hay que reescribir
`elegirDestino()` para que rutee de verdad (mismo criterio de "no inventar datos" que
llevó a sacar el ruteo ficticio de la versión vieja del sitio — ver más abajo).

Los datos de contacto que se muestran como **texto informativo** en "Nosotros" (número
de WhatsApp, email) y en el footer quedaron con link directo, sin pasar por el
formulario — son solo para mostrar el contacto públicamente, no inician un pedido.

### Historia: por qué la versión vieja no ruteaba por zona

La versión vieja (`generar_sitio.py`, Excel) tenía una función que intentaba mandar el
pedido a "hermanos" (Mapi, Luvi, Fede) según la localidad, con teléfonos y emails
**ficticios** que no correspondían a nadie del equipo real — nunca hubiera funcionado.
Se sacó esa lógica al reactivar el sitio (2026-09-22) y todo iba a un único número real.
El 2026-09-25 Franco confirmó que el ruteo por zona es una necesidad real del negocio
(no un capricho de la versión vieja), así que se reintrodujo — pero con datos reales
desde el principio, no inventados.

## Pendientes abiertos

- [ ] **Mapa de zona → persona** para `elegirDestino()` (ver sección de arriba) — sin
      esto, todo pedido de WhatsApp cae en Franco sin importar la zona real del cliente.
- [ ] Decidir cuándo pasar de plan Hobby a **Pro** en Vercel (antes de usar el catálogo
      con clientes reales).
- [ ] Revisar/afinar el texto de la sección "Nosotros" (valores de la empresa) — Franco
      lo va a escribir cuando se le ocurra, por ahora se deja el genérico.
- [ ] ~18 productos MONCAR sin foto (3 códigos con foto ambigua en el PDF: CF 635,
      CF 668, CF 670 — probablemente variantes tipo "CF 635 T11" que el PDF no distingue
      por texto; 1 imagen en pág. 15 sin código cercano). Usan el ícono genérico.
- [ ] `desplegar.bat` tiene rutas asumidas de antes de la reorganización — revisar antes
      de confiar en él, o usar los comandos git a mano (ver sección Deploy).
