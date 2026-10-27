# RELEVO — MONCAR RMH (web pública + panel)
_Actualizado: 2026-10-27 · Rama: b2b-login (repo raíz, sitio y panel)_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
La web B2B (catálogo sin precios + cuentas de clientes con email o Google) está lista en la rama `b2b-login`
de los 3 repos, con preview en Vercel. Antes de publicar, Franco tiene que habilitar el registro y Google en
Supabase (P1 en pendientes). Después: probar en el preview y pasar las 3 ramas a `main` juntas.

## Objetivo de la última sesión
Agrupar la web en `moncar-web/`, comparar con la propuesta de MAT12 y pasar la web a B2B con login — cumplido
(falta la configuración de Supabase/Google que hace Franco y publicar).

## Estado del repo
- Cambios sin commitear: `moncar-web/propuesta/propuesta-web-moncar.html` (sin trackear, no es de esta sesión).
  En `moncar-erp/`: `.claude/` sin trackear (venía de antes).
- Último commit (raíz, b2b-login): a8bf3ab Web B2B: login con Google (PKCE) + formulario para completar la ficha
- `sitio` (moncar-rmh-web) b2b-login: ff606d0 · preview https://moncar-rmh-z0zz8oqxb-moncar-rmh.vercel.app
- `backend-web` (moncar-panel) b2b-login: fb08b55 · migraciones 0003 y 0004 YA aplicadas en Supabase
- Producción (main de los 3) sin cambios: moncar-rmh.com.ar sigue mostrando precios públicos.

## Decisiones pendientes (Franco)
- CAPTCHA (Turnstile) antes de abrir el registro con email, o abrir sin CAPTCHA.
- De dónde sale el catálogo (panel) vs los precios (ERP) cuando se conecte la web con el panel.

## Comandos útiles
- Generar y subir precios: `python moncar-web/pagina-web/generar_sitio_drive.py` (lee Sheets, sube a `precios_web`
  con la clave de `backend-web/.env.local`; si falla la subida NO genera el HTML). Sin tocar Supabase: `--sin-subir-precios`.
- Migraciones del panel: desde `moncar-web/backend-web`: `npm run migrar`.
- Prueba de permisos de clientes (19 casos, crea y borra usuarios de prueba): `node scripts/probar-clientes-web.mjs`.
- Probar local: `python -m http.server 8090 --directory moncar-web/pagina-web/sitio` (recargar con `?v=N` por la caché).
- Publicar: commit + push en `sitio` (main = producción); otra rama = preview (pide login de Vercel).

## Trampas / no hacer
- Publicar en este orden: registro habilitado en Supabase → generar con subida de precios → recién ahí las 3 ramas a main.
  Si se publica antes, los clientes con cuenta no ven precios.
- Los precios NO van en el HTML: solo en `precios_web` (RLS: cliente activo o admin). No volver a meterlos en PRODUCTOS.
- supabase-js con `flowType: 'pkce'`: la vuelta de Google llega como `?code=`; el `#` lo usa la web para navegar.
- Renombrar/mover carpetas falla si una sesión de Claude tiene esa carpeta como directorio de trabajo: archivarla antes.
- La plantilla de la web vive en el generador (`generar_html`, `_CSS`, `_JS`): no editar `sitio/index.html` a mano.

## Historial
- 2026-10-27: web B2B en rama b2b-login (cuentas email + Google, pantalla Clientes, migraciones 0003/0004); MAT12 analizada.
- 2026-10-26: carpetas de la web agrupadas en moncar-web/ (backend-web + pagina-web); panel definido como backend.
- 2026-10-26: web v3 (diseño v1 + usabilidad v2) en producción; ramas rediseno / rediseno-web mergeadas.
