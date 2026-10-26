# RELEVO — MONCAR RMH (página web / catálogo)
_Actualizado: 2026-10-26 17:33 · Rama: dev_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
Carpetas de la web ya agrupadas (2026-10-26): `moncar-web/backend-web/` (panel, repo moncar-panel) y
`moncar-web/pagina-web/` (generador; `sitio/` = repo moncar-rmh-web). Franco definió que el panel sea el
backend de la web. Lo próximo: conectar el generador con la base del panel (productos, categorías, textos).

## Objetivo de la última sesión
Rediseñar UX/UI de moncar-rmh.com.ar en una rama, con preview en Vercel, y pasarlo a main — cumplido (v3).

## Estado del repo
- Cambios sin commitear: ninguno
- Último commit: 1412026 Web v3: 2 productos por fila en el celular con medidas visibles
- Sitio publicado (`moncar-web/pagina-web/sitio`, repo moncar-rmh-web): main 6246933

## Decisiones pendientes (Franco)
- Resuelto: el panel es el backend de la web (falta conectarlo). De dónde salen los precios (ERP) vs el catálogo (panel).

## Comandos útiles
- Generar: `python moncar-web/pagina-web/generar_sitio_drive.py` → `moncar-web/pagina-web/sitio/index.html` (lee Sheets en vivo).
- Publicar: desde `moncar-web/pagina-web/sitio`: `git add -A && git commit && git push` (main = producción). Otra rama = preview (pide login de Vercel).
- Probar local: `python -m http.server 8090 --directory moncar-web/pagina-web/sitio`.

## Trampas / no hacer
- La plantilla de la web vive en el generador (`generar_html`, `_CSS`, `_JS`), no editar `sitio/index.html` a mano: se pisa al regenerar.
- `sitio/` es repo aparte (moncar-rmh-web); el generador está en el repo raíz (moncar-dashboard, rama dev). Cambiar diseño = commit en los dos.

## Historial
- 2026-10-26: carpetas de la web agrupadas en moncar-web/ (backend-web + pagina-web); panel definido como backend.
- 2026-10-26: web v3 (diseño v1 + usabilidad v2) en producción; ramas rediseno / rediseno-web mergeadas.
