# RELEVO — MONCAR RMH (página web / catálogo)
_Actualizado: 2026-10-26 17:33 · Rama: dev_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
La web v3 está en vivo en moncar-rmh.com.ar. Lo próximo que pidió Franco: mover `moncar-web/` dentro de
`pagina-web/` como `backend-web/` (solo carpetas, cada uno con su repo, Vercel y base; NO compartir base).
Hacerlo desde una sesión abierta en `proyectos/moncar` (no desde moncar-web) y con VS Code cerrado.

## Objetivo de la última sesión
Rediseñar UX/UI de moncar-rmh.com.ar en una rama, con preview en Vercel, y pasarlo a main — cumplido (v3).

## Estado del repo
- Cambios sin commitear: ninguno
- Último commit: 1412026 Web v3: 2 productos por fila en el celular con medidas visibles
- Sitio publicado (`pagina-web/sitio`, repo moncar-rmh-web): main 6246933

## Decisiones pendientes (Franco)
- Para qué queda el panel `moncar-web` si la web pública no lee su base (lo cargado ahí no aparece en la web).

## Comandos útiles
- Generar: `python pagina-web/generar_sitio_drive.py` → `pagina-web/sitio/index.html` (lee Sheets en vivo).
- Publicar: desde `pagina-web/sitio`: `git add -A && git commit && git push` (main = producción). Otra rama = preview (pide login de Vercel).
- Probar local: `python -m http.server 8090 --directory pagina-web/sitio`.

## Trampas / no hacer
- La plantilla de la web vive en el generador (`generar_html`, `_CSS`, `_JS`), no editar `sitio/index.html` a mano: se pisa al regenerar.
- `sitio/` es repo aparte (moncar-rmh-web); el generador está en el repo raíz (moncar-dashboard, rama dev). Cambiar diseño = commit en los dos.

## Historial
- 2026-10-26: web v3 (diseño v1 + usabilidad v2) en producción; ramas rediseno / rediseno-web mergeadas.
