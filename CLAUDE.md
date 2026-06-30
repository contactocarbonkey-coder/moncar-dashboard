# MONCAR RMH — Briefing del proyecto

Proyecto de Franco: dashboard de análisis de ventas y control de stock para MONCAR RMH
(empresa familiar de repuestos eléctricos para herramientas: carbones e interruptores).

---

## Estructura de archivos

```
claude MONCAR RMH/
├── CLAUDE.md                   ← este archivo
├── dashboard_live.html         ← dashboard generado (output)
├── ARCHIVOS BASE/
│   ├── VENTAS.xlsx             ← fuente principal de ventas y listas de precios
│   ├── STOCK_MONCAR.xlsx       ← stock de carbones (341 productos)
│   └── STOCK_RMH.xlsx          ← stock de interruptores (36 productos)
└── .claude/
    ├── generar_dashboard.py    ← script principal: lee Excel, genera HTML
    └── servidor.py             ← servidor HTTP en puerto 8050; expone /refresh
```

---

## Cómo funciona el sistema

```
abrir_dashboard.bat → servidor.py (puerto 8050)
                         ↓ al cargar o al clickear "Actualizar datos"
                    generar_dashboard.py
                         ↓ lee con openpyxl (load_wb con retry x3 para locks)
                    VENTAS.xlsx + STOCK_MONCAR.xlsx + STOCK_RMH.xlsx
                         ↓ inyecta JSON en variable TEMPLATE
                    dashboard_live.html  (window.DATA = {...})
```

`generar_dashboard.py` tiene todo el HTML como string `TEMPLATE`. Para modificar el
dashboard, siempre editar `TEMPLATE` dentro de ese script y ejecutar para regenerar.

---

## VENTAS.xlsx — hojas clave

- **Ventas**: filas 2–761. Cols: REMITO, VENDEDOR, FECHA, LINEA, COD, CANTIDAD, CLIENTE,
  COMISION, PRECIO UNIT (fórmula), SUBTOTAL (manual), COMENTARIOS, CMV, MARGEN,
  MARCA, TIPO_CLI, ZONA, LOCALIDAD, PROVINCIA, DESC
- **LP MONCAR**: lista de precios solo para productos MONCAR (CF, CC, CE). Fila 3 = headers,
  datos desde fila 4. Columna E = precio vigente.
- **LP RMH**: lista de precios para productos RM. Cols: A=COD, B=EMPRESA, C=MARCA,
  D=DESCRIPCION, E=Precio anterior, F=Vigencia Abr-26 (precios nuevos desde abril 2026).
- **Costos** y **Clientes**: datos auxiliares.

**Fórmula PRECIO UNIT** (col I de Ventas, filas 2–761):
```
=IF(TRIM(E2)="","",IF(LEFT(TRIM(E2),2)="RM",
  IFERROR(IF(C2>=DATE(2026,4,1),INDEX('LP RMH'!$F:$F,MATCH(TRIM(E2),'LP RMH'!$A:$A,0)),
                                 INDEX('LP RMH'!$E:$E,MATCH(TRIM(E2),'LP RMH'!$A:$A,0))),0),
  IFERROR(INDEX('LP MONCAR'!$E:$E,MATCH(TRIM(E2),'LP MONCAR'!$A:$A,0)),0)))
```

---

## Dashboard — páginas disponibles

| Página | Qué muestra |
|--------|-------------|
| Ventas | KPIs generales, evolución, top vendedores, desglose por línea |
| Productos | Buscador + detalle por producto (KPIs, evolución dual-axis, top clientes) + tabla resumen |
| Marcas | Buscador + detalle por marca (KPIs, evolución, top productos) + tabla resumen |
| Clientes | Análisis por cliente |
| Stock - MONCAR | Tabla de stock de carbones con alertas de reposición |
| Stock - RMH | Tabla de stock de interruptores |
| P&L | Estado de resultados (en desarrollo) |

Filtros globales: fecha (desde/hasta mes), línea (MONCAR/RMH/todas), vendedor.

---

## JS — estado global y funciones clave

```javascript
window.DATA          // JSON inyectado por Python: {ventas, stock_mc, stock_rmh, precios}
const MC_ROWS        // window.DATA.stock_mc.rows
const RMH_ROWS       // window.DATA.stock_rmh.rows
let F                // filtros globales {desde, hasta, linea, vendedor}
let PROD_SEL         // producto seleccionado en página Productos (null = todos)
let MARCA_SEL        // marca seleccionada en página Marcas (null = todas)
```

Funciones principales: `renderVentas()`, `renderProductos()`, `renderMarcas()`,
`renderClientes()`, `renderStock(isMC)`, `renderPL()`.

---

## Decisiones técnicas importantes

- **load_wb()**: wrapper sobre openpyxl que lee en modo binario (BytesIO) con 3 reintentos
  y delay de 2s entre intentos. Permite leer archivos aunque OneDrive los esté sincronizando.
  Si el archivo está abierto en Excel con lock exclusivo, falla igual → cerrar Excel primero.

- **LP split**: LP MONCAR y LP RMH son hojas separadas desde abril 2026 porque RMH tuvo
  actualización de precios propia. La detección es `LEFT(TRIM(COD),2)="RM"`.

- **servidor.py**: usa `importlib.util` para importar `generar_dashboard` como módulo.
  Si se edita `generar_dashboard.py` con el servidor corriendo, hay que reiniciar el servidor
  para que tome los cambios (el módulo se carga una sola vez).

---

## Negocio — contexto rápido

- **MONCAR**: carbones eléctricos. Códigos CF/CC/CE. 341 SKUs.
- **RMH**: interruptores. Códigos RM. 36 SKUs.
- **Vendedores**: Franco (admin + vendedor zona Escobar), Juani (30% comisión, stock propio),
  Maxi (30% comisión, stock propio + gestiona ventas de Lita).
- **Comisión Franco admin**: 20% sobre facturación total — aún no implementada en P&L.

---

## Dirección futura discutida

- **Supabase** como reemplazo de Excel: en evaluación. Permitiría editar datos desde el HTML
  directamente (browser → Supabase API, sin servidor Python de por medio para los datos).
  Pendiente decisión de Franco sobre si migrar y por dónde empezar.
