# MONCAR RMH Dashboard — Referencia técnica y de negocio

Detalle estable que no cambia seguido. El `CLAUDE.md` remite acá cuando hace falta este
nivel de detalle; no es necesario leerlo para tareas chicas.

---

## VENTAS — hojas clave

(Misma estructura en el `.xlsx` local y en el Google Sheet convertido.)

- **Ventas**: ~800 filas válidas / 12 meses (medido 2026-09-08). Columnas: N° REMITO,
  VENDEDOR, FECHA, LÍNEA, COD, CANTIDAD, CLIENTE, COMIS%, PRECIO UNIT (fórmula), SUBTOTAL,
  AÑO-MES, COMENTARIOS, CMV, MARGEN, MARCA, TIPO CLIENTE, ZONA, LOCALIDAD, PROVINCIA,
  DESCRIPCION. Desde 2026-09 además: `COBRADO_EN_MES` (SI/NO), `MES_COBRO`, `DESCUENTO`
  ($ por línea), `DESCUENTO_PCT` (%). `SUBTOTAL` ya viene neto de descuento. Desde
  2026-10-05: `TABLERO` (col AC, texto, ej "009") — marca las filas que son componentes
  de un tablero vendido; vacía en ventas sueltas. **A1 tiene que decir `N° REMITO`**: el
  2026-10 apareció renombrada a "remito" y la app escribió remitos sin número (ver
  `HISTORIAL.md`); la app ya acepta el alias, pero no renombrarla.
- **TABLEROS** (desde 2026-10-05): registro de tableros organizadores armados. Una fila
  por código: `N° TABLERO` (texto, "009") | `FECHA ARMADO` | `COD` (desplegable de LP
  MONCAR) | `CANTIDAD`. Sin precio (sale de LP MONCAR al vender) y sin estado (vendido =
  su N° aparece en `Ventas.TABLERO`). Desde 2026-10-06 Franco la carga subiendo el PDF
  de su planilla en remitos-app → `/admin` → Tableros (también se puede a mano).
  Detalle en `remitos-app/CLAUDE.md` → "Tableros".
- **LP MONCAR**: lista de precios solo productos MONCAR (CF, CC, CE). Fila 3 = headers,
  datos desde fila 4. Columna E = precio vigente.
- **LP RMH**: lista de precios productos RM. Cols: A=COD, B=EMPRESA, C=MARCA,
  D=DESCRIPCION, E=Precio anterior, F=Vigencia Abr-26 (precios nuevos desde abril 2026).
- **Costos**, **Clientes**: datos auxiliares.

**Fórmula PRECIO UNIT** (col I de Ventas):
```
=IF(TRIM(E2)="","",IF(LEFT(TRIM(E2),2)="RM",
  IFERROR(IF(C2>=DATE(2026,4,1),INDEX('LP RMH'!$F:$F,MATCH(TRIM(E2),'LP RMH'!$A:$A,0)),
                                 INDEX('LP RMH'!$E:$E,MATCH(TRIM(E2),'LP RMH'!$A:$A,0))),0),
  IFERROR(INDEX('LP MONCAR'!$E:$E,MATCH(TRIM(E2),'LP MONCAR'!$A:$A,0)),0)))
```

---

## Arquitectura de STOCK_MONCAR / STOCK_RMH (3 hojas, misma lógica en ambos)

**Productos** (stock base inicial — fila 1 título, fila 2 TOTALES, fila 3 headers, fila
4+ datos): COD, MARCA/DESCRIPCION, PRECIO UNIT, BASE NO EMB, BASE EMBOLSADO, BASE PREMIUM
(solo MONCAR), BASE CABA. Solo se toca para el stock inicial; después todo pasa por
Movimientos.

**Movimientos** (diario de cambios): Nº (auto), FECHA, TIPO MOVIMIENTO (desplegable),
ESTADO (desplegable), COD PRODUCTO, CANTIDAD, QUIÉN EMBOLSÓ (solo MONCAR), COMENTARIOS,
AÑO-MES (auto). RMH no tiene "quién embolsó" y tiene menos filas disponibles que MONCAR.

Tipos de movimiento en uso: Ingreso, Embolsado, Embolsado Premium (solo MONCAR), Salida
Venta Franco, Salida Lita, Entrega a Juani / Entrega Maxi, Salida Venta Juani / Maxi /
CABA, Devolución simple, Préstamo dado/recibido, Devolución préstamo. Desde 2026-09-07:
Salida Venta Juli/Franco (automáticas desde `remitos-app`, ver "Pools JULI/FRANCO" en su
CLAUDE.md) y **Retiro CABA/David/Maxi** (carga manual de Franco — hace el camino inverso a
"Entrega a X": resta del pool del vendedor y devuelve a EMBOLSADO DEP. / PREMIUM DEP. del
depósito). Desde 2026-10-05: **Salida Tablero** (solo MONCAR, automática desde
`remitos-app`): componentes de un tablero vendido, restan de EMBOLSADO DEP. (o PREMIUM
DEP. según ESTADO) venda quien venda.

**Stock** (estado actual, fórmulas VLOOKUP+SUMIFS sobre Movimientos): COD,
MARCA/DESCRIPCION, NO EMBOLSADO, EMBOLSADO DEP., PREMIUM DEP. (solo MONCAR), STOCK CABA,
STOCK DAVID, STOCK MAXI (solo RMH), TOTAL VENDIBLE, TOTAL GENERAL, PRÉSTAMOS (dado), DEBO
(recibido).

**Índices de columna (0-based) que usa el script** — se leen por posición, no por nombre:
- MONCAR: `[COD, MARCA, NO_EMB, EMB_DEP, PREM_DEP, CABA, VENDIBLE, TOTAL, PREST, DEBO]`
- RMH: `[COD, DESC, NO_EMB, EMB_DEP, CABA, MAXI, VENDIBLE, TOTAL, PREST, DEBO]`

Actualizados el 2026-09-01 cuando entró STOCK DAVID (ver `HISTORIAL.md`). Si vuelve a
moverse una columna en la hoja `Stock`, actualizar `read_stock_mc()` / `read_stock_rmh()`
en los dos generadores.

---

## Lectura de Google Sheets vía gviz — cuidados (solo versión Drive)

`fetch_sheet_csv()` usa `tqx=out:csv`, que **no preserva la posición de fila igual que la
API de Sheets** cuando hay filas de instrucciones/títulos arriba del header real, y lo
hace **distinto según la hoja**:

- `STOCK_RMH!Movimientos`: la API cuenta 2 filas de instrucciones antes del header
  (`remitos-app` escribe desde la fila 4), pero el CSV de gviz las colapsa y el header
  aparece en la fila 0 del array.
- `LP RMH`: 3 filas (título, nota de vigencia, header) que gviz junta en la fila 0, con el
  texto de las dos primeras pegado dentro de la celda de columna A del header.
- `Costos`: **no** colapsa filas (header queda en su fila 2), pero las celdas de header de
  columnas D en adelante (`COSTO USD`, `COSTO PESOS`, …) llegan **vacías** por CSV aunque
  tengan texto vía API — hay que ubicarlas por posición fija después de encontrar la fila
  de headers por `"COD"` en columna A.

**Regla práctica**: nunca asumir un índice de fila fijo para saltear instrucciones en una
lectura CSV nueva. Buscar la fila de headers por contenido (como hace `read_ventas()` para
la hoja `Ventas`, que no tiene este problema) y validar con una lectura de prueba.

Además: gviz puede devolver una respuesta **truncada** sin error (ver 2026-09-08 en
`HISTORIAL.md`). Por eso `read_ventas()` lee la hoja `Ventas` dos veces y se queda con la
más grande.

---

## Dashboard — páginas disponibles

**Navegación (desde 2026-10-02, los dos generadores)**: menú lateral izquierdo plegable
(`<aside class="sidebar">`, `toggleSidebar()`, estado en `localStorage.moncar_sidebar`;
plegado = solo íconos, 64px) agrupado por secciones: General / Ventas / Producto y
stock / Plata. Barra superior (`.topbar`): breadcrumb `#crumb` (sección › página),
selector de vendedores, fecha de datos, ícono refresh (`refreshData()` → `/refresh`,
oculto si el HTML se abre como archivo) y tema. Los botones grandes "Actualizar datos" y
"PDF" que inyectaba el servidor se sacaron; el servidor Drive solo inyecta el auto-refresh
de 4 min.

**Sumar una página nueva**: (1) un `<button class="sb-item" data-page="X" data-sec="Sección"
title="Nombre" onclick="showPage('X')">` con su ícono SVG dentro del `<nav class="sb-nav">`;
(2) un `<div id="page-X" class="page">`; (3) su `renderX()` en `renderPage()`; (4) si usa
los filtros de período, sumarla a la lista de `showPage()` que muestra
`#analytics-controls`. En los dos generadores.

| Página | Qué muestra |
|--------|-------------|
| Overview | KPIs generales, evolución, top vendedores, desglose por línea + tarjeta **Tableros vendidos** (los dos generadores) + tarjeta **Mercadería Lita** (abajo de todo, solo Drive) |
| Ventas (Detalle) | Tabla completa tipo tabla dinámica — filtros cruzados por vendedor (pills), marca/cliente/zona/tipo cliente (desplegable con buscador), orden por columna, buscador de texto libre, "Descargar CSV". KPIs de totales arriba. |
| Productos | Buscador + detalle por producto (KPIs, evolución dual-axis, top clientes) + tabla resumen |
| Marcas | Buscador + detalle por marca (KPIs, evolución, top productos) + tabla resumen |
| Clientes | Análisis por cliente |
| Stock - MONCAR | Tabla de stock de carbones + alertas de reposición + Prioridad de Embolsado |
| Stock - RMH | Tabla de stock de interruptores + alertas de reposición + Prioridad de Embolsado |
| P&L | Estado de resultados, toggle Económica / Financiera (Rendición) |

**Filtros globales**: período (Mes / 3 Meses / 6 Meses / **MAT** = 12 meses móviles /
**YTD** = year to date / Personalizado), línea (MONCAR / RMH / Ambas), vendedor
(multi-selección vía botón "▾" con checkboxes; la lista sale de
`window.DATA.ventas.vendedores`, no hardcodeada).

### Prioridad de Embolsado (Stock RMH y MONCAR)

`calcEmbolsado()` / `renderEmbolsadoCard()`. Ordena qué embolsar primero por urgencia =
semanas de stock embolsado en depósito ÷ velocidad de venta (promedio últimos 6 meses).

**Regla agregada 2026-07-22 (solo RMH)**: si un producto no tuvo ventas en los últimos 6
meses pero hay stock sin embolsar en depósito y Maxi está en 0 unidades, se prioriza
igual — sube al tope, se resalta en rojo, etiqueta "· Maxi sin stock". Antes esos
productos caían al fondo (urgencia infinita por falta de historial). Aplicado en los dos
generadores.

**Regla agregada 2026-09-25 (solo MONCAR, solo Drive)**: lo pendiente de Lita
(`litaPendQty()`) se reserva del `EMB_DEP` estándar (el premium no la cubre): `Emb. Dep.`
mostrado = `max(0, EMB_DEP − pendLita) + PREM_DEP`. Si el estándar no alcanza →
`urgentLita`: sube al tope (después de "Maxi sin stock"), rojo, "· Lita: faltan X", y el
faltante se suma a "Embolsar ahora". Columna "Lita pend.".

### Mercadería Lita (Overview, solo versión Drive)

Qué le debe Franco a Lita: ventas MONCAR de Maxi (salen del stock físico de Lita).
Fuente: filas `Salida Lita` de `STOCK_MONCAR!Movimientos` (`read_lita()` → `DATA.lita =
{pend, ent}`). Pendiente = `CONFIRMACIÓN=Pendiente` y `ESTADO=Embolsado`; Entregado va
al historial agrupado por `FECHA ENTREGA` (col M), o por mes de venta si no tiene fecha
(entregas previas al 2026-09-25, marcadas "aprox."). CONFIRMACIÓN vacía = se ignora.

- **Pestaña Pendiente**: por código — pendiente, desde, embolsado estándar en depósito,
  sin embolsar, "¿Embolsar?" (faltan = pendiente − max(0, EMB_DEP)), y botón
  **Entregar** con cantidad editable (parcial permitido). Botón **Copiar lista**
  (`litaCopiar()`, 2026-10-02): copia "COD x cant" ordenado por código + total, con
  negrita de WhatsApp, para mandarle a Lita.
- **Pestaña Historial**: selector de mes de entrega (o Todos), uds por código y fechas.
- **Escritura**: `litaEntregar()` → `POST /lita/entregar` (`servidor_drive.py`) →
  `.claude/entregas_lita.py`: lee la hoja por la API (no gviz), marca FIFO por fecha,
  parte la última fila si hace falta (fila nueva Pendiente al final, col A vacía porque
  es ARRAYFORMULA), escribe `Entregado` + fecha, espera 2s y regenera. Solo funciona con
  el dashboard servido por `abrir_dashboard_drive.bat` (no abriendo el HTML suelto).
- Credencial: cuenta de servicio de `remitos-app` (lee `remitos-app/.env.local`).
  Dependencia Python: `google-auth`.

---

### Tableros vendidos (Overview, los dos generadores, desde 2026-10-05)

`read_ventas()` agrega el campo `TABLERO` a cada fila (de `Ventas.TABLERO`, normalizado a
3 dígitos con `norm_tablero()`; vacío si la columna no existe, ej. el Excel local).
`renderTableros(rows, cmp)` (llamada al final de `renderOverview()`) agrupa por par
(`REMITO`, `TABLERO`) sobre las mismas filas filtradas del Overview (período, línea,
vendedor): KPIs de cantidad (con variación vs "Comparar con"), facturación y % del total,
promedio por tablero, margen bruto (sin costo del organizador) + tabla N° / fecha /
cliente / vendedor / remito / códigos / uds / facturación. Solo lectura, no toca ningún
otro KPI (las filas de tablero ya estaban sumadas por código en todo el dashboard).

## JS — estado global y funciones clave

```javascript
window.DATA          // JSON inyectado por Python: {ventas, stock_mc, stock_rmh, lita (solo Drive), precios}
LITA_PEND            // cod → {qty, desde, remitos} pendiente de Lita (solo Drive)
const MC_ROWS        // window.DATA.stock_mc.rows
const RMH_ROWS       // window.DATA.stock_rmh.rows
const F              // filtros globales {vendedores: Set, linea, pl, stock} — vendedores es multi-selección
let PROD_SEL         // producto seleccionado en página Productos (null = todos)
let MARCA_SEL        // marca seleccionada en página Marcas (null = todas)
VDET_FILTERS         // filtros locales de la página Ventas (Detalle)
```

Funciones: `renderOverview()`, `renderVentasDetalle()`, `renderProductos()`,
`renderMarcas()`, `renderClientes()`, `renderStock(isMC)`, `renderPL()`,
`calcEmbolsado()`, `renderEmbolsadoCard()`.

El `TEMPLATE` es casi idéntico entre `generar_dashboard.py` y `generar_dashboard_drive.py`
— solo difiere la parte Python que lee los datos. Editar el TEMPLATE = aplicar el mismo
cambio en los dos (hasta que se decida cuál queda).

---

## Negocio — equipo y comisiones (detalle)

- **MONCAR**: carbones eléctricos. Códigos CF/CC/CE. ~340 SKUs.
- **RMH**: interruptores. Códigos RM. ~35 SKUs.
- **Proceso del producto (MONCAR)**: ingresa materia prima (NO EMBOLSADO) → se embolsa
  (standard o premium) → sale por venta Franco, Salida Lita, o entrega a un vendedor con
  stock propio.

**Equipo**:
- **Juli**: hermana, jefa. Había dejado sus zonas a Franco para enfocarse en
  administración, pero desde el 24/08/2026 volvió a vender activamente (alta como
  vendedora en `remitos-app`, MONCAR + RMH). Vende directo del depósito central igual que
  Franco, sin stock propio — ver `remitos-app/CLAUDE.md`, "Pools JULI/FRANCO".
- **Franco**: admin del negocio + vendedor zona Escobar. Líneas MONCAR + RMH.
- **Andi**: vendedora zona CABA, stock físico propio. Comisión 10% + fijo $400.000/mes.
  Cuenta Google `contacto.carbonkey@gmail.com` (owner de los Sheets nuevos). Su nombre
  real es Andi (confirmado 2026-08-25); "CABA" y "Juani" eran nombres de columna/rol
  viejos. En `remitos-app`: `STOCK_ACCESO_MONCAR` / `STOCK_ACCESO_RMH` = `CABA`. En la
  hoja Ventas aparece como VENDEDOR="Andi" desde agosto 2026.
- **David**: vendedor nuevo (alta en `remitos-app` 2026-09). Mismo esquema que Andi: stock
  propio en las dos líneas (`STOCK_ACCESO_*` = `DAVID`), comisión 30% default. Sin ventas
  reales cargadas todavía al momento de esta nota.
- **Maxi**: vendedor, stock propio físico, comisión 30%, líneas RMH + gestiona ventas de
  Lita en MONCAR.
- **La madre**: integrante familiar del negocio.
- **Lita**: clienta especial MONCAR. Sin ventas propias registradas; Maxi las gestiona con
  stock del depósito central ("Salida Lita" en Movimientos MONCAR).
- **Sofi, Bauti**: colaboradores que embolsan (columna "Quién embolsó" en Movimientos MONCAR).

**Comisiones**:
- Andi (CABA): 10% sobre ventas + fijo $400.000/mes.
- Maxi: 30% sobre ventas propias. David: 30% default.
- Franco vendedor: 15% o 30% según el cliente (definido en VENTAS por cliente).
- Franco admin: 20% adicional sobre facturación total del mes.

**Estado en el P&L**: la comisión admin de Franco (20%) SÍ está implementada
(`renderPL()`, `com_admin = total_rev * 0.20`). El fijo mensual de Andi ($400.000) sigue
sin implementar.

### P&L — vista Económica vs Financiera/Rendición (2026-09)

Toggle arriba de la página P&L. **Económica**: agrupa por fecha de venta (Facturación,
CMV, Margen Bruto, comisiones, Resultado Neto). **Financiera (Rendición)**: agrupa por el
mes en que cada venta se cobra de verdad — para saber qué liquidarle a cada vendedor y qué
le queda a Franco cada mes, sin mezclar con lo vendido pero no cobrado. Sale de
`COBRADO_EN_MES` (SI/NO) y `MES_COBRO` en `Ventas` / `Remitos`, cargadas desde
`remitos-app` (casilla "Cobrado en el mes", tildada por default; si se destilda pide el
mes de cobro). La vista Financiera es a propósito más simple (sin CMV / Margen Bruto /
Resultado Neto — solo tienen sentido por fecha de venta). Implementado en los dos
generadores.

### Google Sheets del vendedor CABA

Creados 2026-07-15, carpeta compartida por `contacto.carbonkey@gmail.com`: "VENTAS CABA"
(Andi carga sus ventas, Franco las importa mensualmente a VENTAS) y "STOCK CABA"
(movimientos de su inventario propio). Relacionado con `setup_stock_caba.gs` en la raíz
del repo (sin explorar en detalle — confirmar si sigue vivo).

---

## Archivos Drive descartados (no confundir)

Durante la migración aparecieron dos Google Sheets con nombres parecidos a los oficiales
pero estructura distinta (hojas PRECIOS/SALIDAS/EMBOLSADO) — confirmado con Franco que son
intentos viejos abandonados, **no son la fuente real**:
- `STOCK_SISTEMA_MONCAR_RMH` (creado y nunca completado, 2026-05-12)
- `STOCK MONCAR 2026` (tracking manual del vendedor CABA, última edición 2026-05-21)
