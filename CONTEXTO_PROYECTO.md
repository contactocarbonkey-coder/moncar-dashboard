# CONTEXTO DEL PROYECTO — MONCAR + RMH

## Última actualización: 2026-06-02

---

## El negocio

**MONCAR RMH** es una empresa familiar argentina de venta de repuestos eléctricos para herramientas:

- **MONCAR**: venta de carbones eléctricos para herramientas
- **RMH**: venta de interruptores eléctricos para herramientas

El producto se importa en bruto (sin embolsar), se procesa (embolsado standard y premium) y se distribuye a ferreterías y reparadores.

### La familia / equipo
- **Juli**: la jefa (hermana). Dejó a Franco a cargo de sus zonas para trabajar en la parte administrativa y operativa.
- **Franco**: rol doble — admin del negocio + vendedor en zona Escobar.
- **Juani**: hermano vendedor. Stock propio físico. Líneas MONCAR+RMH. Comisión 30%.
- **Maxi**: hermano vendedor. Stock propio físico. Líneas RMH + gestiona ventas de Lita. Comisión 30%.
- **La madre**: integrante de la empresa.
- **Lita**: clienta especial MONCAR. No registra ventas propias. Maxi gestiona sus ventas y usa stock del depósito central. Cada movimiento de venta de Lita se registra como "Salida Lita" en el archivo de stock MONCAR, lo que descuenta directo del depósito.
- **Sofi, Bauti**: colaboradores que realizan el embolsado. Se registran en la columna "Quién embolsó" de Movimientos MONCAR.

### Comisiones
- **Juani**: 30% sobre ventas propias
- **Maxi**: 30% sobre ventas propias
- **Franco como vendedor**: 15% o 30% según el cliente (configurado por cliente en el Excel)
- **Franco como admin**: 20% adicional sobre la facturación total del mes (independiente de lo que venda) — **NO está implementado en el P&L del dashboard todavía**

---

## Archivos principales (todos en la misma carpeta)

```
C:\Users\franc\OneDrive\Documentos\KeyCarbon\claude MONCAR RMH\
```

### Archivos de uso diario
| Archivo | Propósito |
|---|---|
| `STOCK_MONCAR.xlsx` | Control de stock de carbones eléctricos |
| `STOCK_RMH.xlsx` | Control de stock de interruptores |
| `VENTAS.xlsx` | Registro de ventas (fuente del dashboard) |
| `dashboard_live.html` | Dashboard visual generado, se abre en el navegador |

### Archivos de soporte del sistema
| Archivo | Propósito |
|---|---|
| `generar_dashboard.py` | Script Python que genera el dashboard_live.html |
| `abrir_dashboard.bat` | Lanzador de doble clic para generar y abrir el dashboard |

### Archivos no integrados al sistema todavía
| Archivo | Estado |
|---|---|
| `RENDICION.xlsx` | Pendiente de explorar e integrar |
| `LP 10 CARBONES 01.04.2026.xlsx` | Lista de precios de carbones (abril 2026), no usada en dashboard |

### Scripts de mantenimiento (ya ejecutados, no son de uso diario)
| Archivo | Qué hizo |
|---|---|
| `formatear_stock.py` | Aplicó formato visual y validaciones al STOCK_MONCAR.xlsx |
| `formatear_rmh.py` | Ídem para STOCK_RMH.xlsx |
| `colorear_stock_moncar.py` | Intento fallido de CF por Python (no borrar, referencia histórica) |
| `fix_cf_moncar.py` | Intento fallido de corrección de CF (no borrar, referencia histórica) |

### Archivos legacy
| Archivo | Estado |
|---|---|
| `dashboard_ventas_v2.html` | Versión vieja, no se usa |
| `dashboard_ventas_v3.html` | Versión vieja, no se usa |

---

## Arquitectura de STOCK_MONCAR.xlsx

El archivo tiene **3 hojas** con una lógica encadenada:

### Hoja "Productos" — stock base inicial
- Fila 1: título
- Fila 2: SUBTOTALES automáticos
- Fila 3: cabeceras
- Fila 4 en adelante: un producto por fila

| Col | Campo | Notas |
|-----|-------|-------|
| A | COD | Prefijos: CC, CF, CE |
| B | MARCA | |
| C | PRECIO UNIT | |
| D | BASE NO EMB | Stock sin embolsar al inicio |
| E | BASE EMBOLSADO | Stock embolsado standard al inicio |
| F | BASE PREMIUM | Stock embolsado premium al inicio |
| G | BASE JUANI | Stock en manos de Juani al inicio |

**Esta hoja solo se toca para el stock inicial. Después todo pasa por Movimientos.**
Total: 341 productos (CODs CC, CF, CE).

### Hoja "Movimientos" — diario de cambios
- Fila 1: título
- Fila 2: instrucciones
- Fila 3: cabeceras
- Filas 4–1003: 1000 filas disponibles para movimientos

| Col | Campo | Tipo |
|-----|-------|------|
| A | Nº | **Automático** — fórmula `=IF(F<>"",ROW()-3,"")` |
| B | FECHA | Manual — formato DD/MM/YYYY |
| C | TIPO MOVIMIENTO | Manual — **desplegable** (ver lista abajo) |
| D | ESTADO | Manual — **desplegable**: No embolsado / Embolsado / Premium |
| E | COD PRODUCTO | Manual |
| F | CANTIDAD | Manual — número entero |
| G | QUIÉN EMBOLSÓ | Manual — **desplegable**: Franco / Sofi / Bauti |
| H | COMENTARIOS | Manual — texto libre |
| I | AÑO-MES | **Automático** — fórmula `=IF(B<>"",TEXT(B,"YYYY-MM"),"")` |

**Tipos de movimiento MONCAR y qué columna del Stock afectan:**

| TIPO | ESTADO | Efecto |
|------|--------|--------|
| Ingreso | No embolsado | + NO_EMB |
| Ingreso | Embolsado | + EMB_DEP |
| Ingreso | Premium | + PREM_DEP |
| Embolsado | — | - NO_EMB, + EMB_DEP |
| Embolsado Premium | — | - NO_EMB, + PREM_DEP |
| Salida Venta Franco | Embolsado / Premium | - EMB_DEP o PREM_DEP |
| Salida Lita | Embolsado / Premium | - EMB_DEP o PREM_DEP (Maxi vende para Lita desde depósito) |
| Entrega a Juani | Embolsado / Premium | - EMB_DEP o PREM_DEP, + JUANI |
| Salida Venta Juani | — | - JUANI |
| Devolución simple | No embolsado / Embolsado / Premium | + columna correspondiente |
| Préstamo dado | No embolsado / Embolsado / Premium | - columna, + PREST |
| Préstamo recibido | No embolsado / Embolsado / Premium | + columna, + DEBO |
| Recupero préstamo dado | No embolsado / Embolsado / Premium | + columna, - PREST |
| Devolución préstamo recibido | No embolsado / Embolsado / Premium | - columna, - DEBO |

### Hoja "Stock" — estado actual calculado
- Fila 1: título
- Fila 2: SUBTOTALES automáticos
- Fila 3: cabeceras
- Filas 4–344: fórmulas por producto

| Col | Campo | Fórmula base |
|-----|-------|-------------|
| A | COD | Valor directo |
| B | MARCA | Valor directo |
| C | NO EMBOLSADO | VLOOKUP(Productos BASE) + SUMIFS(Movimientos) |
| D | EMBOLSADO DEP. | VLOOKUP(Productos BASE) + SUMIFS(Movimientos) |
| E | PREMIUM DEP. | VLOOKUP(Productos BASE) + SUMIFS(Movimientos) |
| F | STOCK JUANI | VLOOKUP(Productos BASE) + SUMIFS(Movimientos) |
| G | TOTAL VENDIBLE | =D+E+F |
| H | TOTAL GENERAL | =C+D+E+F |
| I | PRÉSTAMOS (dado) | SUMIFS neto |
| J | DEBO (recibido) | SUMIFS neto |

**El rango de Movimientos que leen las fórmulas es F4:F1003 (1000 filas).**

**IMPORTANTE:** La primera vez que se abre el archivo en Excel después de crearlo con Python, hay que presionar **Ctrl+Shift+F9** y guardar con **Ctrl+S** para que Excel calcule todas las fórmulas y guarde los valores en caché. Después de eso, el archivo funciona normalmente.

---

## Arquitectura de STOCK_RMH.xlsx

Misma lógica de 3 hojas. Diferencias con MONCAR:

- **Productos**: col B = DESCRIPCION (no MARCA). 36 productos (CODs RM).
- **Movimientos**: 8 columnas (no 9 — no tiene "Quién embolsó"). Auto-fecha en col H (no I). 500 filas (F4:F503).
- **Stock**: col 5 = STOCK CABA (no PREM_DEP), col 6 = STOCK MAXI (no JUANI). Solo 2 estados: No embolsado / Embolsado (sin Premium).

**Tipos de movimiento RMH:**
Ingreso, Embolsado, Devolución simple, Recupero préstamo dado, Préstamo recibido, Préstamo dado, Devolución préstamo recibido, Entrega CABA, Salida Venta CABA, Entrega Maxi, Salida Venta Maxi, Salida Venta Franco.

---

## Formato visual de los archivos de stock

Ambos archivos tienen el mismo esquema de diseño aplicado con Python (scripts formatear_*.py):

**STOCK_MONCAR** — paleta azul:
- Cabeceras: `#1F3864` (azul oscuro) con texto blanco
- Totales: `#2E75B6` (azul medio) con texto blanco
- Filas alternas: `#DEEAF1` (azul muy claro) y blanco
- Cols VENDIBLE/TOTAL: verde claro `#E2EFDA`
- Cols PREST/DEBO: dorado claro `#FFF2CC`
- Campos automáticos (Nº y Año-Mes): gris `#EFEFEF`

**STOCK_RMH** — paleta bordeaux:
- Cabeceras: `#5C1A26` (vino oscuro) con texto blanco
- Totales: `#8B3A47` (vino medio) con texto blanco
- Filas alternas: `#F5EBF0` (blush muy claro) y blanco
- Cols VENDIBLE/TOTAL y PREST/DEBO: mismo verde y dorado que MONCAR

Ambos tienen:
- Sin grilla visible
- Freeze panes: Movimientos en B4 (Nº visible al scrollear), Stock en C4 (COD+MARCA visibles)
- Validaciones de datos con desplegables (sin mensaje de ayuda al hacer clic)

---

## Alerta de colores en STOCK_MONCAR — Movimientos (pendiente de aplicar)

Se intentó aplicar via Python pero falló por incompatibilidad de namespace en openpyxl 3.1.5. Las fórmulas están listas para aplicar **manualmente en Excel** (Inicio → Formato condicional → Nueva regla → Usar fórmula).

**Rango:** F4:F1003 (columna CANTIDAD). Solo activa para tipos Salida.

Lógica: las fórmulas del Stock ya incluyen el efecto de la fila actual. Si el stock queda negativo, hay problema.

**Separador:** Excel en español argentino usa `;` (punto y coma).

**VERDE** — hay stock embolsado suficiente:
```
=AND(OR(C4="Salida Venta Franco";C4="Salida Lita";C4="Entrega a Juani";C4="Salida Venta Juani");F4<>"";E4<>"";IF(C4="Salida Venta Juani";IFERROR(VLOOKUP(E4;Stock!$A:$J;6;0);0);IF(D4="Premium";IFERROR(VLOOKUP(E4;Stock!$A:$J;5;0);0);IFERROR(VLOOKUP(E4;Stock!$A:$J;4;0);0)))>=0)
```

**NARANJA** — falta embolsado pero hay materia prima sin embolsar para compensar:
```
=AND(OR(C4="Salida Venta Franco";C4="Salida Lita";C4="Entrega a Juani");F4<>"";E4<>"";IF(D4="Premium";IFERROR(VLOOKUP(E4;Stock!$A:$J;5;0);0);IFERROR(VLOOKUP(E4;Stock!$A:$J;4;0);0))<0;IFERROR(VLOOKUP(E4;Stock!$A:$J;3;0);0)>=(-(IF(D4="Premium";IFERROR(VLOOKUP(E4;Stock!$A:$J;5;0);0);IFERROR(VLOOKUP(E4;Stock!$A:$J;4;0);0)))))
```

**ROJA** — stock negativo y sin materia prima para cubrir (o Juani sin stock):
```
=AND(OR(C4="Salida Venta Franco";C4="Salida Lita";C4="Entrega a Juani";C4="Salida Venta Juani");F4<>"";E4<>"";OR(AND(C4="Salida Venta Juani";IFERROR(VLOOKUP(E4;Stock!$A:$J;6;0);0)<0);AND(OR(C4="Salida Venta Franco";C4="Salida Lita";C4="Entrega a Juani");IF(D4="Premium";IFERROR(VLOOKUP(E4;Stock!$A:$J;5;0);0);IFERROR(VLOOKUP(E4;Stock!$A:$J;4;0);0))<0;IFERROR(VLOOKUP(E4;Stock!$A:$J;3;0);0)<(-(IF(D4="Premium";IFERROR(VLOOKUP(E4;Stock!$A:$J;5;0);0);IFERROR(VLOOKUP(E4;Stock!$A:$J;4;0);0)))))))
```

Orden de prioridad en "Administrar reglas": verde arriba, naranja en el medio, rojo abajo.
Colores sugeridos: verde `D5EBC2`, naranja `FAD7B4`, rojo `F2B8B8`.

---

## Sistema técnico — Dashboard

### Flujo de uso
1. Franco hace doble clic en `abrir_dashboard.bat`
2. El .bat ejecuta `python generar_dashboard.py`
3. El script lee los tres Excel (VENTAS, STOCK_MONCAR, STOCK_RMH) con openpyxl
4. Genera `dashboard_live.html` inyectando todos los datos como JSON
5. Abre el HTML en el navegador automáticamente

### generar_dashboard.py — detalles importantes
- Lee VENTAS.xlsx con openpyxl `data_only=True`
- Lee STOCK_MONCAR.xlsx: hoja "Stock" desde fila 4, CODs CC/CF/CE
  - **Fallback**: si los valores del Stock son None (fórmulas no calculadas por Excel todavía), lee los valores BASE de la hoja "Productos" automáticamente y muestra un aviso en consola.
- Lee STOCK_RMH.xlsx: hoja "Stock" desde fila 4, CODs RM
- Dependencia Python: solo `openpyxl` (`pip install openpyxl`)

### VENTAS.xlsx — estructura hoja "Ventas"
Columnas (0-based):
| Idx | Campo | Notas |
|-----|-------|-------|
| 0 | REMITO | |
| 1 | VENDEDOR | |
| 2 | FECHA | |
| 3 | LINEA | "MONCAR" o "RMH" |
| 4 | COD | |
| 5 | CANTIDAD | |
| 6 | CLIENTE | |
| 7 | COMISION | Decimal ej. 0.30 |
| 8 | PRECIO | |
| 9 | SUBTOTAL | |
| 10 | MES | String "YYYY-MM" |
| 11 | COMENTARIOS | |
| 12 | CMV | Costo de Mercadería Vendida |
| 13 | MARGEN | |
| 14 | MARCA | |
| 15 | TIPO_CLI | |
| 16 | ZONA | |
| 17 | LOCALIDAD | |
| 18 | PROVINCIA | |
| 19 | DESC | |

Filas excluidas si: SUBTOTAL <= 0, MES no es "YYYY-MM", COD vacío o empieza con "COD".

### Dashboard — tabs y funcionalidades
1. **Overview** — KPIs del período (facturación, resultado neto, COGS, margen, órdenes, ticket promedio). Gráficos: facturación mensual MONCAR vs RMH, unidades, top 10 productos y clientes. Selector de período: Semana / Mes / 3M / 6M / Año / Personalizado con flechas. Comparación opcional vs período anterior.
2. **Productos** — Top 15 por facturación y unidades, evolución top 8.
3. **Marcas** — Facturación y unidades por marca, evolución top 6, tabla.
4. **Clientes** — Top 15 por facturación y unidades, tabla completa con provincia.
5. **Stock** — Toggle MONCAR/RMH. KPIs, gráfico top 15, donut distribución, tabla con "meses de stock".
6. **P&L** — Facturación, CMV, margen bruto, comisiones por vendedor (expandible), resultado neto. Waterfall, barras mensuales, línea de margen %.

### Lógica de negocio en el dashboard
- Línea MONCAR/RMH: del campo LINEA si válido, sino por COD (empieza "RM" → RMH, resto → MONCAR)
- CMV: viene del Excel directamente
- Comisión = SUBTOTAL × COMISION (tasa por fila)
- Resultado Neto = Facturación − CMV − Comisión
- **Comisión admin de Franco (20% sobre facturación total) NO implementada en P&L**

---

## Pendientes conocidos

- [ ] Aplicar formato condicional de colores (verde/naranja/rojo) en col F de Movimientos MONCAR — fórmulas listas en este documento, hay que hacerlo manualmente en Excel
- [ ] Comisión admin de Franco (20%) no está reflejada en el P&L del dashboard
- [ ] Explorar e integrar RENDICION.xlsx
- [ ] LP 10 CARBONES 01.04.2026.xlsx no está integrada al sistema

---

## Dependencias Python

```
pip install openpyxl
```
Python estándar + openpyxl. Sin otros paquetes externos.
