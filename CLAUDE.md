# MONCAR RMH — Índice de proyectos

**Última actualización: 2026-09-28.**

Franco tiene 3 proyectos MONCAR RMH. Este archivo es el mapa del **Dashboard** y el
router hacia los otros dos.

## Router — a dónde ir según la tarea

| Proyecto | Estado | Carpeta / repo | Cómo se corre | Doc detallada |
|---|---|---|---|---|
| **Dashboard** (este) | Activo | raíz + `.claude/` | `abrir_dashboard_drive.bat` | este archivo → `HISTORIAL.md`, `REFERENCIA.md` |
| **App remitos** | Activo | `remitos-app/` (repo git propio) | Vercel / `npm run dev` | `remitos-app/CLAUDE.md` |
| **Página web / catálogo** | Activo, en vivo en `moncar-rmh.com.ar` desde 2026-09-28 | `pagina-web/` | ver `pagina-web/CLAUDE.md` | `pagina-web/CLAUDE.md` |
| **Web multi-zona** (Next.js + Supabase) | En desarrollo desde 2026-10-05 — solo panel admin, local | `moncar-web/` (git propio) | `npm run dev` | `moncar-web/CLAUDE.md` |

Si la tarea es de **remitos**, leer `remitos-app/CLAUDE.md` primero — nada de eso vive
acá. Si la tarea es de la **página web / catálogo** actual, leer `pagina-web/CLAUDE.md`
primero — nada de eso vive acá tampoco. La **web multi-zona** es la propuesta nueva para
las 6 zonas de la empresa (no solo la de Juli): su doc está en `moncar-web/CLAUDE.md`.

---

## Dashboard — qué es

Dashboard de una sola página (HTML + JS) para análisis de ventas y control de stock de
MONCAR RMH (empresa familiar argentina: carbones eléctricos "MONCAR", interruptores
"RMH"). Es la fuente de ingresos principal de Franco. Un script Python lee los datos y
los inyecta como JSON (`window.DATA`) dentro de un `TEMPLATE` HTML.

## Estado actual — dos versiones del mismo dashboard

Mismo `TEMPLATE`, distinta fuente de datos:

| | **Drive (Google Sheets) — la que se usa** | Excel local — en evaluación, sin uso activo |
|---|---|---|
| Lanzador | `abrir_dashboard_drive.bat` | `abrir_dashboard.bat` |
| Servidor | `.claude/servidor_drive.py` (puerto 8051) | `.claude/servidor.py` (puerto 8050) |
| Generador | `.claude/generar_dashboard_drive.py` | `.claude/generar_dashboard.py` |
| Fuente | 3 Google Sheets vía CSV público (gviz) | `ARCHIVOS BASE/*.xlsx` (openpyxl) |
| Output | `dashboard_drive.html` | `dashboard_live.html` |

Franco ya no tiene Office; carga todo en Google Sheets, así que en la práctica solo corre
la versión Drive. **La decisión formal de cuál queda como definitiva sigue abierta** — no
asumir que la Excel local se puede borrar.

**Regla al tocar el `TEMPLATE`** (HTML/JS, lógica de negocio en JS): aplicar el mismo
cambio en los dos generadores (`generar_dashboard_drive.py` y `generar_dashboard.py`) o
se desincronizan. Algunas features nuevas (ventas manuales sin remito, fix de lectura
doble) están **solo en la versión Drive** a propósito — ver `HISTORIAL.md`.

## Estructura de archivos

```
claude MONCAR RMH/
├── CLAUDE.md               ← este archivo (mapa + router)
├── HISTORIAL.md            ← cambios pasados y pendientes resueltos
├── REFERENCIA.md           ← detalle estable: hojas, fórmulas, columnas, páginas, JS
├── .claude/
│   ├── generar_dashboard_drive.py   ← ACTIVO: lee Sheets → dashboard_drive.html
│   ├── servidor_drive.py            ← servidor 8051, expone /refresh (ícono ↻ arriba a la derecha) y POST /lita/entregar
│   ├── entregas_lita.py             ← escribe en Sheets: marca entregas a Lita
│   ├── prestamos.py                 ← escribe en Sheets: préstamos entre zonas (página Préstamos)
│   ├── generar_dashboard.py         ← versión Excel local (sin uso activo)
│   ├── servidor.py                  ← servidor 8050
│   └── settings.local.json
├── abrir_dashboard_drive.bat        ← ACTIVO
├── abrir_dashboard.bat              ← versión Excel local
├── dashboard_drive.html             ← output ACTIVO
├── dashboard_live.html              ← output Excel local (gitignored)
├── *_dev.html                       ← outputs de prueba (flag --dev), gitignored
├── ARCHIVOS BASE/*.xlsx             ← fuente Excel local (gitignored)
├── BACKUPS/                         ← backups de datos (gitignored)
├── LP/LP 10 CARBONES 01.04.2026.xlsx  ← lista de precios, NO integrada al dashboard (pendiente)
├── setup_stock_caba.gs              ← Apps Script hoja STOCK CABA de Andi — CONFIRMAR si sigue vivo
├── remitos-app/                     ← proyecto hermano, repo git propio (ver su CLAUDE.md)
└── _legacy/                         ← archivos sin uso confirmado (ver _legacy/LEEME.txt)
```

## Cómo funciona (versión Drive)

```
abrir_dashboard_drive.bat → servidor_drive.py (8051)
   ↓ al cargar o al clickear "Actualizar datos"
generar_dashboard_drive.py
   ↓ descarga cada hoja como CSV público (fetch_sheet_csv, retry x3)
3 Google Sheets  →  mismo procesamiento y mismo TEMPLATE que el original
   ↓
dashboard_drive.html  (window.DATA = {...})
```

`fetch_sheet_csv(id, hoja)` arma
`https://docs.google.com/spreadsheets/d/{ID}/gviz/tq?tqx=out:csv&sheet={HOJA}`, la baja
con `urllib` y la parsea con `csv`. Sin OAuth, sin service account. Números en formato
argentino (`"10.298,15"` → `10298.15`, función `n()`); fechas `dd/mm/yyyy` → `yyyy-mm-dd`
(`parse_fecha()`). Google calcula las fórmulas server-side, así que el CSV siempre trae
valores ya resueltos.

**IDs de los Google Sheets** (owner `contacto.carbonkey@gmail.com` = cuenta de Andi, no la de Franco):

| Archivo | ID | Hojas |
|---|---|---|
| VENTAS | `1K8wcDYl5E6hf7B-Go-R-3BPnCISNFKibnEV2C4AmfOE` | Ventas, LP MONCAR, LP RMH, Costos, Clientes |
| STOCK_MONCAR | `1u9xG2uyeT3V8osBi9Qz0xofDCPmFmhOXlu85c1OIEQ8` | Productos, Movimientos, Stock |
| STOCK_RMH | `1Msl0CC5abCoIaM0MpeCtYOoHVHI0mjPlfElc5x8ytJE` | Productos, Movimientos, Stock |

**Requisito crítico**: los 3 Sheets tienen que quedar en "Cualquiera con el enlace -
Lector". Si Franco los pasa a privado, el script Drive deja de leerlos y falla (es CSV
público, no hay login de por medio).

**Excepción de escritura (2026-09-25)**: el botón "Entregar" de la tarjeta **Mercadería
Lita** (Overview) escribe en `STOCK_MONCAR!Movimientos` (CONFIRMACIÓN → Entregado +
FECHA ENTREGA) vía `entregas_lita.py`, con la cuenta de servicio de `remitos-app` (lee
`remitos-app/.env.local`; requiere `pip install google-auth`). Desde 2026-10-06 también
la página **Préstamos** (`prestamos.py`, `POST /prestamos/registrar`) escribe en
`Movimientos` de STOCK_MONCAR y STOCK_RMH (columnas nuevas PERSONA / SALDA COD / SALDA
CANT). Detalle de los dos en `REFERENCIA.md`.

## Gotchas vigentes

- **`servidor*.py` carga el generador como módulo una sola vez al arrancar.** Si editás
  el generador con el servidor corriendo, reiniciá el servidor.
- **Tocar el `TEMPLATE` = editar los dos generadores** (`_drive` y el original).
- **El dashboard lee las hojas Stock por posición de columna (`row[N]`), no por nombre.**
  Si en `remitos-app` se agrega una columna a la hoja `Stock` de STOCK_MONCAR/STOCK_RMH,
  hay que actualizar los índices en `read_stock_mc()` / `read_stock_rmh()` de los dos
  generadores o se lee el número equivocado en silencio. Índices actuales en `REFERENCIA.md`.
- **Lectura CSV de Sheets (gviz) desalinea las filas de instrucciones distinto según la
  hoja.** Nunca asumir índice de fila fijo para el header: buscarlo por contenido. Casos
  confirmados (`Movimientos`, `LP RMH`, `Costos`) en `REFERENCIA.md`.
- **`read_ventas()` (solo versión Drive) pide la hoja `Ventas` dos veces y se queda con
  la lectura más grande** — gviz puede devolver una respuesta truncada sin ningún error.
  No quitar ese doble fetch.
- **Las fórmulas de `Stock` deben leer `Movimientos` con rango abierto** (`$F$4:$F`, no
  `$F$4:$F$922`): con rango cerrado, lo cargado debajo queda fuera del stock sin error.
  Si se agregan columnas/fórmulas nuevas en `Stock`, usar rango abierto. Script de
  corrección: `remitos-app/scripts/ampliar-rango-movimientos.ts`.
- **`load_wb()` (solo versión Excel local)**: reintenta 3× con 2s de delay para saltar
  locks de OneDrive. Si Excel tiene el archivo con lock exclusivo, hay que cerrarlo.

## Negocio — referencia rápida

- **MONCAR** = carbones, códigos CF/CC/CE, ~340 SKUs. **RMH** = interruptores, códigos RM, ~35 SKUs.
- **Equipo**: Juli (hermana, jefa; volvió a vender el 24/08/2026), Franco (admin +
  vendedor Escobar), Andi (vendedora CABA, stock propio), David (vendedor nuevo, stock
  propio), Maxi (RMH + ventas de Lita en MONCAR, stock propio), Lita (clienta especial,
  sus ventas las carga Maxi). "Juani" / "CABA" eran nombres viejos de Andi.
- **Comisiones**: Andi 10% + fijo $400.000/mes (el fijo NO está en el P&L todavía); Maxi
  y David 30%; Franco vendedor 15% o 30% según cliente; Franco admin 20% sobre
  facturación del mes (SÍ está en el P&L).
- Detalle completo de equipo, hojas, fórmulas y arquitectura de stock: **`REFERENCIA.md`**.

## Pendientes abiertos

- [ ] **Correr una vez** (desde `remitos-app/`) `npx tsx scripts/preparar-prestamos.ts
      --apply`: columnas PERSONA/SALDA en Movimientos + datos de los préstamos viejos.
      Sin esto la página Préstamos muestra todo como "Sin asignar" y no deja guardar.

- [ ] Decidir cuál dashboard queda como definitivo (Drive vs Excel local) o si conviven.
- [ ] Fijo mensual de Andi ($400.000) no implementado en el P&L.
- [ ] **Análisis de descuentos** (pedido de Franco): la hoja `Ventas` ya tiene
      `DESCUENTO` ($) y `DESCUENTO_PCT` (%) cargados desde `remitos-app`. `SUBTOTAL` ya
      viene neto, así que el P&L ya refleja el impacto; falta una **vista/tabla dedicada**
      a mostrar el detalle del descuento por cliente/producto/vendedor. Es solo-lectura
      sobre columnas que ya existen.
- [ ] Formato condicional (verde/naranja/rojo) en la columna CANTIDAD de `Movimientos`
      para señalar stock negativo. Las fórmulas viejas eran para Excel (separador `;`);
      hay que reescribirlas para Sheets (`,`) si se retoma.
- [ ] `LP 10 CARBONES 01.04.2026.xlsx` sigue sin integrarse al dashboard.
- [ ] Confirmar si `setup_stock_caba.gs` sigue en uso (hoja STOCK CABA de Andi).
- [ ] **Tableros** (2026-10-05, feature en `remitos-app`): sumar el costo del organizador
      regalado (fila extra en `Ventas` con SUBTOTAL 0 y CMV = costo) cuando Franco pase
      la hoja de costos. (El indicador "Tableros vendidos" del Overview ya está hecho.)
- [ ] Códigos MONCAR sin costo en `Costos` → CMV 0 en `Ventas` (ej. CF 712, CF 970, CF
      971, CF 972, CE 309; 44 de 365 filas MONCAR de sept-2026). Infla el margen; revisar
      junto con la hoja de costos nueva.
- [ ] Productos en stock negativo tras el backfill de Juli (2026-09-07) — lista en
      `remitos-app/CLAUDE.md`; Franco tiene que revisar conteo físico si le interesa.

Pendientes ya resueltos y su detalle: **`HISTORIAL.md`**.
