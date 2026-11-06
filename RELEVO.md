# RELEVO — MONCAR RMH (raíz: dashboard + router de proyectos)
_Actualizado: 2026-11-06 · Rama: b2b-login (repo raíz)_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
Preguntarle a Franco si revisó la sección **Visitas** de remitos-app (moncar-rmh-remitos.vercel.app → menú, solo Andi + admin).
Si anda: apagar la app separada `moncar-visitas` (borrar el proyecto en Vercel, archivar el repo
`contactocarbonkey-coder/moncar-visitas`, la carpeta `moncar-visitas/` puede quedar) y borrar de la planilla lo que haya cargado de prueba.

## Objetivo de la última sesión
Herramienta para que Andi cargue visitas a ferreterías desde el celular a la planilla Rutas CABA — cumplido
(primero app aparte, después sección de remitos-app; en producción, remitos main d4c32db).

## Estado del repo
- Cambios sin commitear: `.gitignore` (ignora `moncar-visitas/`), `CLAUDE.md` (fila de la app de visitas en el router),
  `.claude/launch.json` (nuevo: levantar moncar-visitas en :3010 y remitos-app en :3011 desde el panel de vista previa),
  `moncar-web/propuesta/` (de sesiones anteriores, sin decidir).
- Último commit: eea9df8 CLAUDE.md: contacto.carbonkey es la cuenta de trabajo de Franco, no de Andi
- remitos-app (repo propio): `main` = d4c32db con Visitas; rama `visitas-andi` ya mergeada.

## Decisiones pendientes (Franco)
- Revisar Visitas y dar el OK para apagar moncar-visitas.
- Qué hacer con `moncar-web/propuesta/` (versionar o ignorar; el machete es privado).

## Comandos útiles
- Planilla Rutas CABA: `1HpGsvXlZDHtEMS9iMyKFAYI9GN8pDbTeFs5vq0IXUG4` (detalle en Cerebro `moncar-rutas-caba-sheet`).
- Probar la lectura de Visitas desde remitos sin la app: script descartable en `remitos-app/scripts/_check-tmp.ts`
  que importe `leerLocales` de `src/lib/visitas.ts` (con dotenv `.env.local`), `npx tsx`, borrarlo.
- `git push` colgado: `git -c credential.helper= -c "credential.helper=!gh auth git-credential" push ...`

## Trampas / no hacer
- Visitas nunca escribe las columnas N:O de Locales (fórmulas MAP). La tabla llega hasta la fila 600: la cuenta de
  servicio no puede agrandarla (choca con las protecciones); si se llena, agrandarla desde la cuenta de Franco.
- Repos nuevos con Vercel: `git config user.email contacto.carbonkey@gmail.com` ANTES del primer commit (si no, deploy "Blocked").
- No leer ni mostrar PINes/claves de `.env.local` (el permiso automático lo bloquea, y está bien).
- Sheets por API: locale es_ES (fórmulas con `;`); las tablas no aceptan `setDataValidation`.

## Historial
- 2026-11-06: visitas de Andi a Rutas CABA → app moncar-visitas y después sección Visitas en remitos-app (producción).
- 2026-11-03: planilla Rutas CABA de Andi rearmada (Locales + Dashboard + bloqueos), en Google Sheets.
- 2026-10-28: propuesta web para los jefes de zona ($1.650.000) + machete de preguntas, en moncar-web/propuesta/.
- 2026-10-27: web B2B en rama b2b-login (cuentas email + Google, pantalla Clientes, migraciones 0003/0004); MAT12 analizada.
- 2026-10-26: carpetas de la web agrupadas en moncar-web/ (backend-web + pagina-web); panel definido como backend.
