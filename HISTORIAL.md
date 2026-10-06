# MONCAR RMH Dashboard — Historial

Cambios pasados y pendientes resueltos. El `CLAUDE.md` tiene solo el estado actual; esto
es la memoria de "por qué está así". Fechas absolutas.

---

## Entregas a Lita imputadas por remito — 2026-10-06 (solo versión Drive)

**Problema** (detectado por Franco el mismo día de la entrega en lote): con cantidad
suelta por código, el backend repartía FIFO entre remitos y, con fechas iguales,
desempataba por orden de fila. Franco entregó remito 4280 (AYL MAQUINAS) completo y el
sistema imputó CF 674 y CF 020 en parte a los remitos 4279/4275 y partió filas del 4280.
**Fix**: la unidad de entrega pasa a ser (código, remito) — desplegable por código con
remito + cliente, cantidad por remito (parcial solo dentro del mismo remito). Verificado
con simulación contra la hoja real (escritura interceptada) y Edge headless sobre el HTML
de prueba.

**Corrección de datos** (el modo automático de Claude no permitió escribir en el Sheet;
la hace Franco a mano en `STOCK_MONCAR!Movimientos`): F738=40, F742=30; filas 664, 730 y
731 → L=Pendiente y M vacío; borrar filas 855 y 856 (partes Pendiente del 4280).
**Hecha y verificada el mismo día** (relectura de la hoja: valores correctos, hoja en
855 filas, sin filas sueltas del 4280).

## Entrega en lote de Mercadería Lita — 2026-10-06 (solo versión Drive)

**Pedido de Franco**: dejar de marcar entregas de a un código. Casillas por fila +
"Entregar seleccionados"; el botón individual se sacó (decisión de Franco: una sola forma
de hacerlo). La casilla del encabezado tilda solo los OK; las casillas arrancan vacías.
Backend: `entregas_lita.entregar_varios()` — una lectura, un `batchUpdate`, todo-o-nada
(si un código pide más de lo pendiente no se escribe nada). Las filas partidas por
entregas parciales se agregan consecutivas al final. Verificado con simulación contra la
hoja real interceptando la escritura (3 códigos, 1 completo + 2 parciales) y `node --check`
del JS generado.

---

## Indicador "Tableros vendidos" en el Overview — 2026-10-05 (los dos generadores)

**Pedido de Franco**, con la condición explícita de que nada de lo existente cambie.
Tarjeta nueva en el Overview (debajo de los Top 10, arriba de Lita) — detalle en
`REFERENCIA.md` → "Tableros vendidos". Mismo parche en los dos generadores (script con
anclas que deben aparecer exactamente una vez; respeta CRLF del generador Excel).

**Regresión verificada**: (1) el diff contra la copia previa de cada generador da **0
líneas borradas/modificadas**, solo agregadas; (2) `window.DATA` nuevo = viejo + campo
`TABLERO` (idéntico en todo lo demás, Drive y Excel); (3) Edge headless sobre el HTML
viejo y el nuevo, recorriendo 7 páginas × 5 períodos × 3 filtros de línea + Stock RMH +
P&L financiera = 107 vistas por versión: **texto idéntico** (excluyendo la tarjeta
nueva) y **0 errores JS**; (4) captura de la tarjeta con datos de prueba (copia en
scratchpad marcando el remito 4290 como tablero 009).

**Hallazgo**: el margen bruto de MONCAR da ~95% en todo sept-2026 porque varios códigos no
tienen costo en `Costos` (CMV 0). No es del indicador — anotado en Pendientes.

---

## Tableros organizadores + bug de N° REMITO en Ventas — 2026-10-05

**Tableros** (implementado en `remitos-app`, detalle en su `HISTORIAL.md`): la primera
venta a un cliente se hace con un tablero organizador (regalo) lleno de productos MONCAR.
Ahora el vendedor carga "Tablero 009" como un solo ítem y la app escribe una fila por
código en `Ventas` (con `TABLERO=009`) y un movimiento `"Salida Tablero"` por código en
STOCK_MONCAR. **El dashboard no necesitó cambios**: ve filas por código como siempre.
Cambios en Sheets: hoja nueva `VENTAS!TABLEROS` (con el tablero 009), columna
`Ventas.TABLERO`, tipo "Salida Tablero" en el desplegable y en las fórmulas de EMBOLSADO
DEP./PREMIUM DEP. (`remitos-app/scripts/setup-tableros.ts`).

**Bug encontrado de paso**: A1 de `Ventas` estaba renombrada a "remito" (no "N° REMITO")
desde ~2026-10-02. La app escribía filas sin número de remito, y al eliminar un remito
no encontraba sus filas para borrarlas. Afectó a 2 remitos: **4290** (Juli, 60 filas,
se completó el número — verificado contra `Remitos`, mismo total $704.870) y **4289**
(prueba de Franco ya eliminada, quedó 1 fila huérfana de Casa Franchi → borrada). A1
vuelve a ser `N° REMITO` y `headerIndex()` de la app acepta "REMITO" como alias. El
dashboard leía la col 0 por fallback, no se rompió.

---

## Página Préstamos — 2026-10-06 (solo Drive)

**Pedido de Franco**: registrar y gestionar desde el dashboard los préstamos de
mercadería con las otras zonas (Luvi, Lita, Naza, Mapi, Fede), MONCAR y RMH, en las dos
direcciones, con registro claro de cuánto se debe y desde cuándo. Primer paso de "hacer
TODO desde el dashboard". Detalle en `REFERENCIA.md` → Préstamos.

- **Encontrado al relevar**: quién prestó estaba solo en COMENTARIOS (texto libre,
  inconsistente) y 3 devoluciones (CF 179 x20, CE 045 x14, CF 033 x5) se habían cargado
  sin ESTADO → la deuda bajaba pero el stock no (39 uds de más). Franco confirmó que se
  devolvieron embolsadas. Personas de los préstamos viejos: las dio Franco (CF 971,
  CF 179, CF 033 → Lita; RM 717 → Naza).
- **Decisiones de Franco**: incluir RMH y "préstamo dado"; las devoluciones a veces son
  con otro código/cantidad ("otra cosa que el otro quiera") → columnas SALDA COD/CANT.
- **Sin tocar**: fórmulas de `Stock`, `remitos-app`, `entregas_lita.py`, resto de
  páginas. Verificado comparando la versión commiteada vs la nueva con los mismos datos.
- Solo versión Drive (depende de escribir en Sheets), igual que Lita.

---

## Menú lateral plegable + refresh como ícono — 2026-10-02 (los dos generadores)

**Pedido de Franco** (con captura de referencia de otra app): pasar las pestañas de
arriba a un menú lateral izquierdo plegable, agrupado por secciones, porque va a sumar
muchas páginas (costos, contabilidad, rendición, compras, embolsado…). Sacar los botones
"Actualizar datos" y "PDF"; el refresh queda como ícono chico arriba a la derecha.

- Aplicado con el mismo patch a `generar_dashboard_drive.py` y `generar_dashboard.py`
  (bloques idénticos; la versión Excel usa finales CRLF). Servidores: `servidor_drive.py`
  conserva solo el auto-refresh de 4 min; `servidor.py` ya no inyecta nada.
  (2026-10-06: Franco pidió sacar también el auto-refresh — `servidor_drive.py` ya no
  inyecta nada; los datos se actualizan solo con el ícono ↻.)
- Verificado: 33 vistas (todas las páginas × 5 períodos) sin errores JS; JS de los dos
  generadores pasa `node --check`; capturas con menú abierto, plegado, tema claro y
  servido por HTTP (ícono refresh visible).
- Mismo día: botón **Copiar lista** en la tarjeta de Lita (texto para WhatsApp).

---

## Mercadería Lita en el Overview + rangos abiertos en Stock — 2026-09-25 (solo Drive)

**Pedido de Franco**: ver en el dashboard qué le debe devolver a Lita (lo que Maxi vende
en MONCAR sale del stock físico de Lita y Franco se lo repone), marcar entregas
(parciales incluidas) y que eso alimente la prioridad de embolsado.

- **Datos**: ya existían — cada venta MONCAR de Maxi escribe desde `remitos-app` una fila
  `Salida Lita` con `CONFIRMACIÓN=Pendiente`; la fórmula de `Stock` resta de `EMBOLSADO
  DEP.` solo las `Entregado`. Se agregó la columna **M `FECHA ENTREGA`** en
  `STOCK_MONCAR!Movimientos` para el historial mensual.
- **Primera escritura del dashboard en Sheets**: `.claude/entregas_lita.py` +
  `POST /lita/entregar` en `servidor_drive.py`. Usa la cuenta de servicio de
  `remitos-app` (lee `remitos-app/.env.local`). Requiere `pip install google-auth`.
  Entrega FIFO por fecha; si una fila no se entrega entera se parte (fila original →
  Entregado con lo entregado; fila nueva Pendiente con el resto, mismo remito).
  Endpoint protegido por chequeo de `Origin` (solo el propio dashboard).
- **Decisiones de Franco**: Lita solo recibe embolsado **estándar** (el premium no la
  cubre); filas `Salida Lita` con CONFIRMACIÓN vacía (cargas manuales viejas) se
  ignoran; Franco espera ≥2 semanas tras el remito antes de marcar, así que el edge case
  "editar remito ya entregado" no le preocupa.
- **Prioridad de embolsado MONCAR**: lo pendiente de Lita se reserva del `EMB_DEP`
  estándar antes de calcular semanas; si no alcanza → sube arriba, rojo, "· Lita:
  faltan X", y se suma a "Embolsar ahora". Nueva columna "Lita pend.".
- Solo en la versión Drive (depende de Movimientos vía Sheets), a propósito.
- **Bug latente encontrado de paso**: las fórmulas de `Stock` leían
  `Movimientos!$X$4:$X$922` (MONCAR) y `$X$451` (RMH) — lo cargado debajo quedaba fuera
  del cálculo sin error; además la grilla MONCAR terminaba en 922 filas y `appendRows()`
  de remitos-app falla si la fila no existe. MONCAR iba por la fila ~764, RMH por la ~370.
  Script `remitos-app/scripts/ampliar-rango-movimientos.ts` (dry-run por default,
  `--apply`): rangos abiertos `$X$4:$X` + grilla a 5000 filas + header FECHA ENTREGA.
  Franco lo corrió el 2026-09-25; verificado (0 celdas pendientes, grillas en 5000).

---

## Bug resuelto: Stock MONCAR roto en la versión Excel local (2026-07-28)

`generar_dashboard.py` → `read_stock_mc()` tenía una variable `caba_v` que en la rama
normal (cuando la hoja `Stock` ya tiene fórmulas calculadas, el caso común) nunca se
asignaba — quedó de un rename viejo de la columna "Juani" a "CABA". El script tiraba
`cannot access local variable 'caba_v'` y la página Stock-MONCAR quedaba vacía (0
productos). Franco lo detectó al abrir `abrir_dashboard.bat` y ver la página en blanco.

Arreglado en los dos generadores (`generar_dashboard.py` y `generar_dashboard_drive.py`
usan `row[5]` para `caba_v`, la columna correcta según el layout real de la hoja
`Stock`). Verificado: ambos leen 341/340 productos MONCAR sin error (la diferencia de a
uno es normal, por desfasaje entre el Excel local y el Sheet).

---

## Ventas manuales sin remito (RMH, histórico) — 2026-09-07

**Motivo** (pedido explícito de Franco): antes de `remitos-app` (mayo a julio 2026), las
ventas de Maxi se registraban solo como movimientos de salida en `STOCK_RMH!Movimientos`
(`TIPO MOVIMIENTO = "Salida Venta Maxi"`, comentario tipo `"vta mayo"` / `"vta junio
excel maxi"`), **sin** ninguna fila en `Ventas`. Esas unidades no aparecían en ninguna
estadística del dashboard (ejemplo real: `RM 029` en MAT mostraba 2 unidades cargadas por
remito e ignoraba 4 más vendidas en mayo/junio por el camino viejo).

**Solución, solo en `generar_dashboard_drive.py`** (Franco decidió no portarlo a la
versión Excel local): función `leer_ventas_manuales_sin_remito()`, llamada en `main()`
justo después de `read_ventas()`, que **extiende el mismo `raw`** (lista de dicts que
consume todo el dashboard) con filas sintéticas del mismo shape. Cero cambios en el
TEMPLATE — al tener la misma forma, fluyen solas por Overview, Ventas Detalle, Productos,
Marcas y P&L.

**Cómo distingue "sin remito"**: filtra `Movimientos` de `STOCK_RMH` por `TIPO
MOVIMIENTO` que empiece con `"Salida Venta"` y descarta las que tengan comentario que
matchee `"Remito N° <numero>"` (esas ya las contó `remitos-app` en `Ventas`). De 73
encontradas, 17 tenían `CANTIDAD=0` (error de carga) → quedan 56, reconstruidas con:

- `PRECIO` / `SUBTOTAL`: **estimado** con `LP RMH` columna "Vigencia Abr-26" (mejor
  referencia histórica disponible, no es lo que Maxi cobró realmente). Si el producto no
  está ahí, la línea se descarta en vez de mostrar `$0` engañoso.
- `CMV` / `MARGEN`: igual que cualquier venta real, contra la hoja `Costos` ("COSTO
  PESOS" × cantidad). Si el producto no está en `Costos` (pasa con `RM 029`), `CMV=0`
  para esa línea — mismo comportamiento que la fórmula real cuando el VLOOKUP falla.
- `CLIENTE` / `TIPO_CLI` / `ZONA` = `"Sin dato"`. `REMITO` vacío a propósito (inventar un
  número falso sería peor). Efecto secundario deliberado: el contador de "cantidad de
  ventas/remitos" del TEMPLATE (`new Set(rows.map(r=>r.REMITO))...`) NO cuenta estas
  líneas, aunque sí entran en unidades, facturación y margen. `COMENTARIOS` = `"Venta
  manual sin remito (<comentario original>)"`.

---

## Historial reciente

### 2026-09-22 a 2026-09-28 — Página web: fotos reales, rediseño y dominio propio en vivo

Después de la reactivación del 22/09 (ver entrada de abajo), una semana de trabajo
sobre `pagina-web/` que la dejó publicada en `moncar-rmh.com.ar`. Detalle completo y
actualizado en `pagina-web/CLAUDE.md` — acá solo el resumen de qué pasó y por qué:

1. **Fotos reales de producto**: se encontraron y procesaron dos PDF que Franco tenía
   (`CATALOGO CARBONES MONCAR (1) (2).pdf` y `LISTA DE PRECIO RMH AGOSTO 2026.pdf`).
   Scripts nuevos `extraer_imagenes_catalogo.py` / `extraer_imagenes_rmh.py` matchean
   cada foto embebida en el PDF con su código de producto por posición (layouts
   distintos entre los dos PDF). Resultado: 322/340 fotos MONCAR + 39/39 fotos RMH.
2. **Stock oculto**: a pedido de Franco, se pausó por completo (badge, filtro "solo
   stock", botón deshabilitado) — el catálogo ahora muestra todo como pedible sin
   importar disponibilidad real. El dato se sigue leyendo, solo no se usa en el JS.
3. **Rediseño completo** inspirado en la web de Dowen Paggio (que Franco mandó de
   referencia): tipografía Oswald, paleta negro/gris/naranja, íconos de categoría reales
   (no emoji) sacados de Material Symbols/MDI/Game-Icons vía Iconify (embebidos como SVG
   fijo, sin dependencia externa en runtime), selector de cantidad en cada producto,
   paginación "Cargar más" (por performance, ya que las fotos reales pesan bastante más
   que los íconos SVG viejos), y ruteo liviano por URL (`#productos/cat/...`) para que
   el botón atrás y los links compartidos funcionen.
4. **WhatsApp con identificación de zona**: Franco aclaró que el negocio sí rutea
   pedidos por zona a distintos vendedores (esto no era el capricho ficticio de la
   versión vieja que se había sacado el 22/09 — es una necesidad real). Se encontraron
   los números reales del equipo en el PDF de RMH (Franco, Andrés Castagna/Andi, David
   Wu, Maximiliano Ojeda/Maxi) y se armó el mecanismo (formulario nombre/zona antes de
   abrir WhatsApp), pero **falta que Franco pase el mapa de qué zona corresponde a
   quién** — hoy todo cae en Franco por default.
5. **Dominio propio, `moncar-rmh.com.ar`, publicado 2026-09-28**: registrado en nic.ar
   (a nombre de Juli/Maria Julieta Montero) → delegado a nameservers de Cloudflare (por
   falta de editor de registros A/CNAME en el panel de nic.ar) → Cloudflare apunta a
   Vercel. **Contratiempo de 3 días**: la primera carga de nameservers en nic.ar (24/09)
   no se confirmó con el botón "EJECUTAR CAMBIOS" — quedó guardado en la UI pero el
   WHOIS real nunca tuvo los nameservers, sin ningún error visible. Se detectó
   consultando `whois.nic.ar` directo (puerto 43) en vez de confiar en el panel, y se
   resolvió recargando y confirmando. **Corrección importante**: el `.com.ar` **no
   resultó gratis** como se asumía al principio — el circuito oficial de nic.ar pidió
   pago igual.
6. **Vercel se quedó en plan Hobby** (no Pro) — decisión de Franco: por ahora es "uso
   personal, para probar la funcionalidad", no lo está mandando a clientes reales
   todavía. Pendiente pasar a Pro antes de usarlo en serio (ver nota de uso comercial en
   `pagina-web/CLAUDE.md`).

### 2026-09-22 — Reactivación de la página web / catálogo

Franco decidió retomar el catálogo público (`_legacy/pagina-web/` → `pagina-web/`,
salió de `_legacy`) como carta de presentación para clientes y potenciales clientes —
no como tienda con cobro online. Cambios:

1. **Nuevo generador `generar_sitio_drive.py`** (el viejo `generar_sitio.py` leía de
   3 Excel locales cuyas rutas ya no existen desde la reorganización, y Franco dejó de
   mantenerlos desde que migró a Sheets en julio). El nuevo lee precios/stock en vivo
   de los mismos 3 Google Sheets del Dashboard (`LP MONCAR`, `LP RMH`, `STOCK_MONCAR`/
   `STOCK_RMH` hoja `Productos`). Las dimensiones de carbones MONCAR (esp/anch/larg/
   terminal) siguen viniendo de `LP/LP 10 CARBONES 01.04.2026.xlsx` porque ese detalle
   no está en Sheets todavía — no rompe la generación si el archivo falta, solo se
   pierden esos filtros.
2. **Se sacó el ruteo por zona a "hermanos" ficticios** (Mapi, Luvi, Fede — teléfonos
   y emails inventados que no correspondían a nadie del equipo real). Ahora todo pedido
   del carrito va siempre al WhatsApp/email real de la empresa. Detalle en
   `pagina-web/CLAUDE.md`.
3. **Vercel plan Pro, no Hobby**: aunque el sitio no cobra online, mostrar/publicitar
   productos de un negocio real cuenta como "uso comercial" en los términos de Vercel
   → Hobby (gratis) no está permitido para este caso, hace falta Pro (USD 20/mes) antes
   de conectar un dominio propio.
4. Pendiente para Franco: crear el repo GitHub + proyecto Vercel (`sitio/` hoy no tiene
   remoto configurado) y comprar el dominio.

### 2026-09-10 — Reorganización de la documentación

Se movieron ~10 archivos sin uso confirmado a `_legacy/` (catálogo web + `generar_sitio.py`
+ `desplegar.bat`, `generar_equivalencias.py` + `EQUIVALENCIAS_MONCAR.xlsx`,
`CONTEXTO_PROYECTO.md`, `RENDICION.xlsx`, `VENTAS_TEST.xlsx`). Se borraron `ca.txt` y los
locks `~$*.xlsx`. El `CLAUDE.md` de 588 líneas se partió en tres: `CLAUDE.md` (mapa +
router, se carga siempre), `HISTORIAL.md` (esto) y `REFERENCIA.md` (detalle estable).
`.gitignore`: se agregó `*_dev.html` y `ca.txt`.

### 2026-09-08 — Fix de lectura truncada + KPIs de totales

Franco reportó que el dashboard Drive "no sumaba todo". `generar_dashboard_drive.py` había
leído la hoja `Ventas` con 183 filas / 3 meses en vez de las 800 reales (12 meses). No fue
bug de código: `fetch_sheet_csv()` (gviz) devolvió una respuesta "exitosa" pero truncada,
sin error — probablemente por una edición/recálculo concurrente en el momento del fetch
(justo después del backfill de 150 movimientos de Juli del día anterior). Corriendo el
script de nuevo sin cambios leyó los 800 bien.

**Fix**: `read_ventas()` ahora pide la hoja `Ventas` dos veces (1s de por medio) y se
queda con la lectura más grande, avisando por consola si difieren. Solo para esa hoja, la
más crítica. Solo en `generar_dashboard_drive.py` (la versión Excel local usa openpyxl, no
puede sufrir esto). Verificado: `dashboard_drive.html` queda con 800 (Ventas) + 56 (ventas
manuales sin remito) = 856 filas en modo MAT.

De paso se agregaron KPIs de totales/subtotales arriba de las tablas de **Ventas
(Detalle)**, **Productos**, **Marcas** y **Clientes** (reaccionan a período/línea/vendedor
y, en Ventas Detalle, a los filtros propios de esa página). Aplicado a ambos generadores,
probado con jsdom antes de promover.

### 2026-09-01 — Columna STOCK DAVID + P&L Económica/Financiera

1. La columna `STOCK DAVID` (agregada en `remitos-app` al lado de `STOCK CABA` en la hoja
   `Stock`) corrió de posición a `TOTAL VENDIBLE` / `TOTAL GENERAL` / `PRÉSTAMOS` / `DEBO`.
   El dashboard las lee por posición fija → se actualizaron los índices en
   `read_stock_mc()` / `read_stock_rmh()` de los dos generadores. Verificado comparando
   Total Vendible / Total General / Stock Maxi antes y después (idénticos).
2. P&L con toggle **Económica** (agrupa por fecha de venta, la de siempre) vs **Financiera
   (Rendición)** (agrupa por el mes en que la venta se cobra de verdad — para saber qué
   liquidarle a cada vendedor sin mezclar con lo vendido no cobrado). Sale de dos columnas
   nuevas en `Ventas` / `Remitos`: `COBRADO_EN_MES` (SI/NO) y `MES_COBRO`, cargadas desde
   `remitos-app` (casilla "Cobrado en el mes"). La vista Financiera es a propósito más
   simple (sin CMV / Margen Bruto / Resultado Neto — esos solo tienen sentido por fecha de
   venta). Implementado en los dos generadores.

### 2026-08-25 — MAT, filtro multi-vendedor, página Ventas (Detalle)

Tres mejoras al TEMPLATE, aplicadas a los dos generadores:
1. Período **MAT** (moving annual total, 12 meses móviles) nuevo, al lado de 3/6 Meses. El
   botón "Año" se renombró a **YTD** (la lógica ya era year-to-date, solo cambió el texto).
2. Filtro de vendedor pasó de single-select a **multi-selección** (`F.vendedores` es un
   `Set`; antes era el string `F.vendedor`). La lista sale sola de los datos reales, no
   está hardcodeada — por eso "Andi" apareció solo al empezar a vender, sin tocar código.
3. Página nueva **"Ventas" (Detalle)**: tabla tipo tabla dinámica con filtros cruzados
   (Vendedor como pills, Marca/Cliente/Zona/Tipo Cliente como desplegable con buscador),
   buscador de texto libre, orden por columna, botón "Descargar CSV". Filtros locales a la
   página (`VDET_FILTERS`); Vendedor comparte estado con `F.vendedores` (sincronizado en
   ambas direcciones).

Se probó antes en `dashboard_drive_dev.html` (`python generar_dashboard_drive.py --dev`,
flag calcado del que ya tenía la versión Excel local) con test jsdom + Node que simula
clicks reales. Franco confirmó visualmente antes de promover.

Se confirmó que "Andi" es el nombre real de la vendedora de CABA (no "Juani", no solo
"CABA"). También se mencionó a "David", vendedor nuevo sin ventas cargadas todavía.

### 2026-08-06 — remitos-app escribe en los Sheets reales

`remitos-app/` pasó de escribir en copias de prueba a escribir directo en los Sheets
reales de este proyecto (VENTAS y STOCK_RMH). Verificado con un remito de prueba creado y
borrado antes de confirmar. Desde acá, toda venta cargada desde esa app es real y afecta
este dashboard al actualizar datos.

### 2026-08-02 a 2026-08-04 — Nace remitos-app

Se armó `remitos-app/` (repo separado) para que Franco y Maxi carguen remitos desde el
celular, con PDF, envío por mail/WhatsApp y escritura automática en el Sheet VENTAS (más
STOCK_RMH para descontar stock de Maxi). Detalle completo en `remitos-app/CLAUDE.md`.

### 2026-07-22 — Migración a Google Sheets

Franco dejó de usar Office y pasó a cargar datos en Google Sheets. Se armó el sistema
Drive en paralelo (`generar_dashboard_drive.py` + `servidor_drive.py` +
`abrir_dashboard_drive.bat`) sin tocar el sistema original. Se detectó y arregló (primero
solo en Drive) el bug de `caba_v`. Se agregó la regla de prioridad de embolsado para
productos sin ventas registradas pero con Maxi en 0 stock (RMH).

---

## Pendientes resueltos

- [x] Bug de `caba_v` en `generar_dashboard.py` — 2026-07-28, arreglado en ambos generadores.
- [x] Comisión admin de Franco (20%) en el P&L — ya estaba implementada; la nota que
      decía "pendiente" estaba desactualizada.
- [x] Rol / comisión de David — dado de alta en `remitos-app` 2026-09, mismo esquema que
      Andi (stock propio, 30% default).
- [x] Vista económica vs financiera / rendición en el P&L — implementado 2026-09.
- [x] `remitos-app/` apuntaba a copias de PRUEBA — desde 2026-08-06 escribe en los Sheets reales.
- [x] Juli (y Franco vendedor) vendían sin generar movimiento de stock — corregido
      2026-09-07 (ver `remitos-app/CLAUDE.md`, "Pools JULI/FRANCO"), con backfill de 150
      movimientos históricos de Juli. De paso se detectó que `CF 635 T11` (real, existe en
      `LP MONCAR`) no tenía fila en `Productos` / `Stock` de STOCK_MONCAR — se renombró la
      fila ambigua `"CF 635"` a `"CF 635 T11"` en las 7 filas donde aparecía. Si aparece
      un `"CF 635 T11"` que antes no existía o un `"CF 635"` que desapareció, es este
      cambio, no un bug. Quedaron productos en stock negativo tras el backfill (lista en
      `remitos-app/CLAUDE.md`) — pendiente de revisión física de Franco.
- [x] No había forma de devolver stock al depósito desde el pool de un vendedor —
      agregado 2026-09-07, tipo de movimiento "Retiro X" (CABA/David/Maxi) en
      STOCK_MONCAR/STOCK_RMH. Carga manual de Franco, no lo dispara `remitos-app`.
- [x] Ventas RMH pre-`remitos-app` (mayo-julio 2026) que quedaron solo como "Salida Venta
      Maxi" en `STOCK_RMH!Movimientos` — agregado 2026-09-07 solo en
      `generar_dashboard_drive.py` (ver "Ventas manuales sin remito" arriba).
- [x] `CONTEXTO_PROYECTO.md` redundante — movido a `_legacy/` el 2026-09-10.
- [x] Raíz del proyecto con ~10 archivos sin explorar mezclados con los activos — movidos
      a `_legacy/` el 2026-09-10; `CLAUDE.md` partido en CLAUDE.md + HISTORIAL.md +
      REFERENCIA.md.
