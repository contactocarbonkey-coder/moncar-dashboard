# RELEVO — MONCAR RMH (raíz: dashboard + router de proyectos)
_Actualizado: 2026-11-20 · Rama: b2b-login (repo raíz) · remitos-app en rama visitas-por-vendedor (otra sesión)_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
P1 nuevo (2026-11-20): en **remitos-app**, mostrar el stock disponible al lado de cada código al cargar un remito,
solo para el usuario admin (Franco). Pendientes de antes: PEDIDO LLAVES 2027 (transferir y compartir) y apagar `moncar-visitas`.

## Objetivo de la última sesión
Carpeta `logos/` + ficha completa del producto en el ERP (detalle en `moncar-erp/RELEVO.md`) — cumplido, en producción.

## Estado del repo
- Repo raíz (rama b2b-login), sin commitear: `logos/` (esta sesión: logos e imágenes de marca renombrados + LEEME.txt),
  `.claude/launch.json` (esta sesión sumó "erp-vista-local" en :3199), y de sesiones anteriores `.gitignore`, `CLAUDE.md`,
  `moncar-web/propuesta/`, `moncar-web/pagina-web/extraer_imagenes_rmh.py` (no lo tocó esta sesión).
- remitos-app está en la rama `visitas-por-vendedor` (otra sesión, falta pasar a main): crear la rama nueva desde `main`.

## Decisiones pendientes (Franco)
- Transferir la propiedad de PEDIDO LLAVES 2027 a contacto.carbonkey (Compartir → Transferir propiedad) y aceptarla.
- Si se quiere que cada familiar edite solo su hoja: pasar los 5 mails.
- Sumar la cuenta de servicio de remitos-app como editora de las protecciones de Rutas CABA (pendiente P2).
- Qué hacer con `moncar-web/propuesta/` y `logos/` (versionar o ignorar).

## Comandos útiles
- Vista previa: `.claude/launch.json` → "remitos-app" en :3011.
- Google Sheets nuevo con formato y fotos: armar .xlsx con openpyxl y subirlo con Drive `create_file` (referencia sha256,
  archivo en el scratchpad, máx 10 MB). Detalle en Cerebro `google-sheets-conector-claude`.

## Trampas / no hacer
- El conector de Google está logueado con la cuenta **personal**: lo que se crea no queda en MONCAR. La cuenta de MONCAR es
  `contacto.carbonkey@gmail.com` (Franco a veces escribe "keycarbon").
- Las fotos `=IMAGE()` de PEDIDO LLAVES 2027 dependen de `www.moncar-rmh.com.ar/imagenes/`: no renombrar ni borrar esas imágenes.
- No guardar visitas de prueba en Rutas CABA desde la app: escribe en la planilla real.
- No leer ni mostrar PINes/claves de `.env.local`.

## Historial
- 2026-11-20: carpeta logos/ y ficha completa del producto en el ERP (0048, producción); decidido el stock en remitos-app solo admin.
- 2026-11-19: planilla Google Sheets PEDIDO LLAVES 2027 (hoja por familiar + Pedido total bloqueada, fotos de la web).
- 2026-11-17: Visitas de Andi: inicio con resumen, zona con sugerencias, ficha con contacto y notas, navegación (producción).
- 2026-11-06: visitas de Andi a Rutas CABA → app moncar-visitas y después sección Visitas en remitos-app (producción).
- 2026-11-03: planilla Rutas CABA de Andi rearmada (Locales + Dashboard + bloqueos), en Google Sheets.
