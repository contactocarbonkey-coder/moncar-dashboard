# RELEVO — MONCAR RMH (raíz: dashboard + router de proyectos)
_Actualizado: 2026-11-19 · Rama: b2b-login (repo raíz) · remitos-app en main 693c801_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
Preguntarle a Franco si transfirió la planilla PEDIDO LLAVES 2027 a contacto.carbonkey, tocó "Permitir acceso" de las fotos
y la compartió con la familia. Después seguir con Visitas: cómo le fue a Andi en su celu y apagar `moncar-visitas`.

## Objetivo de la última sesión
Google Sheets del pedido de llaves (interruptores RM) 2027 desde `Descargas/NUEVO PEDIDO LLAVES  2027.xlsx`: una hoja por
familiar (Fede, Luvi, Juli, Mapi, Naza) y Pedido total bloqueada que suma — cumplido.
Planilla: `1ZC_Ykcks7lWkpD21di3cVAFbkHO5R4u1sDhbVzCqrFk` (hoy dueña la cuenta personal; contacto.carbonkey es editora).

## Estado del repo
- Repo raíz, cambios sin commitear (de sesiones anteriores, no de esta): `.gitignore`, `CLAUDE.md`, `.claude/launch.json`,
  `logos/`, `moncar-web/propuesta/`. Esta sesión no tocó código.

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
- 2026-11-19: planilla Google Sheets PEDIDO LLAVES 2027 (hoja por familiar + Pedido total bloqueada, fotos de la web).
- 2026-11-17: Visitas de Andi: inicio con resumen, zona con sugerencias, ficha con contacto y notas, navegación (producción).
- 2026-11-06: visitas de Andi a Rutas CABA → app moncar-visitas y después sección Visitas en remitos-app (producción).
- 2026-11-03: planilla Rutas CABA de Andi rearmada (Locales + Dashboard + bloqueos), en Google Sheets.
- 2026-10-28: propuesta web para los jefes de zona ($1.650.000) + machete de preguntas, en moncar-web/propuesta/.
