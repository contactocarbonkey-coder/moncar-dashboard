# RELEVO — MONCAR RMH (raíz: dashboard + router de proyectos)
_Actualizado: 2026-11-19 · Rama: b2b-login (repo raíz) · remitos-app en main 693c801_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
Preguntarle a Franco cómo le fue a Andi con Visitas en su celu (entra directo a Visitas, resumen, botón atrás de Android,
Llamar/WhatsApp). Si anda: apagar `moncar-visitas` (proyecto Vercel + archivar el repo) y borrar cargas de prueba.

## Objetivo de la última sesión
Mejorar Visitas para Andi en remitos-app: inicio en Visitas con resumen y sugerencias, zona con autocompletar o nueva,
volver al resumen, ficha del local con contacto y notas, lista por prioridad — cumplido, en producción (remitos main 693c801).

## Estado del repo
- Repo raíz, cambios sin commitear (de sesiones anteriores, no de esta): `.gitignore`, `CLAUDE.md`, `.claude/launch.json`,
  `logos/`, `moncar-web/propuesta/`.
- remitos-app: limpio en `main`. Ramas `visitas-inicio-resumen` y `visitas-navegacion-ficha` ya mergeadas; Franco pidió dejarlas.

## Decisiones pendientes (Franco)
- Sumar la cuenta de servicio de remitos-app como editora de las protecciones de Rutas CABA (si no, una zona nueva no
  entra al desplegable ni al Dashboard). Detalle en el pendiente P2.
- Qué hacer con `moncar-web/propuesta/` y `logos/` (versionar o ignorar).

## Comandos útiles
- Vista previa: `.claude/launch.json` → "remitos-app" en :3011 (panel de vista previa).
- Probar remitos-app como un vendedor sin su PIN: firmar un JWT de prueba con `SESSION_SECRET` (jose, `setExpirationTime("1h")`,
  datos `{vendedor, esAdmin, stockAccesoMoncar, stockAccesoRmh}`) y ponerlo en la cookie `session` de localhost; borrarla al terminar.
- `npm run build` de remitos-app necesita red (Google Fonts): en el sandbox falla sin ella.
- Chequeo contra la planilla: `remitos-app/scripts/_check-tmp.ts` (dotenv `.env.local`), `npx tsx`, borrarlo.

## Trampas / no hacer
- No guardar visitas de prueba en Rutas CABA desde la app: escribe en la planilla real (Locales + Visitas).
- La cuenta de servicio NO puede `updateTable` ni escribir en Dashboard de Rutas CABA (protecciones); sí formatos condicionales.
- Visitas nunca escribe N:O de Locales (fórmulas MAP). La tabla llega hasta la fila 600.
- No leer ni mostrar PINes/claves de `.env.local`.

## Historial
- 2026-11-17: Visitas de Andi: inicio con resumen, zona con sugerencias, ficha con contacto y notas, navegación (producción).
- 2026-11-06: visitas de Andi a Rutas CABA → app moncar-visitas y después sección Visitas en remitos-app (producción).
- 2026-11-03: planilla Rutas CABA de Andi rearmada (Locales + Dashboard + bloqueos), en Google Sheets.
- 2026-10-28: propuesta web para los jefes de zona ($1.650.000) + machete de preguntas, en moncar-web/propuesta/.
- 2026-10-27: web B2B en rama b2b-login (cuentas email + Google, pantalla Clientes, migraciones 0003/0004); MAT12 analizada.
