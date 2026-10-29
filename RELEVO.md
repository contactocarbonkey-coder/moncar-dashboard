# RELEVO — MONCAR RMH (web pública + panel)
_Actualizado: 2026-10-28 · Rama: b2b-login (repo raíz)_
Backlog completo: C:/Users/franc/Cerebro/wiki/proyectos/moncar-rmh/moncar-rmh-pendientes.md

## Para arrancar la próxima sesión
Preguntarle a Franco cómo salió la reunión del 29/10 con los jefes de zona. Si aceptaron: crear las cuentas de
la empresa (GitHub/Vercel/Supabase/dominio) y transferir. Las respuestas sobre marcas por zona cambian la etapa 3.
Sigue en pie lo de la sesión anterior: habilitar registro y Google en Supabase y pasar `b2b-login` a main (P1).

## Objetivo de la última sesión
Armar la propuesta de web + panel para los 6 jefes de zona (alternativa a MAT12) y un machete de preguntas — cumplido.

## Estado del repo
- Cambios sin commitear: `moncar-web/propuesta/` (sin trackear: propuesta, descargable y machete). El machete
  es privado de Franco: decidir si se versiona o se agrega al .gitignore.
- Último commit: b0a612c Relevo de la sesión 2026-10-27 (web B2B con login)
- Links: propuesta https://claude.ai/artifact/H1hQ3Fxr5sTwhc9gLmNuts · machete https://claude.ai/artifact/4GgdU4AxJ8jhxtRhVCwNdu

## Decisiones pendientes (Franco)
- Lo que salga de la reunión: marcas que ve cada zona, mapa zona-localidades, listas por zona, quién aprueba y quién paga.
- CAPTCHA (Turnstile) antes de abrir el registro con email.
- De dónde sale el catálogo (panel) vs los precios (ERP) cuando se conecte la web con el panel.

## Comandos útiles
- Regenerar el descargable de la propuesta: envolver `propuesta-web-moncar.html` con doctype/head/body (el fuente del
  artifact no lo tiene); o editar los dos archivos con el mismo cambio.
- Generar y subir precios: `python moncar-web/pagina-web/generar_sitio_drive.py` (`--sin-subir-precios` para no tocar Supabase).
- Migraciones del panel: desde `moncar-web/backend-web`: `npm run migrar`.
- Prueba de permisos de clientes: `node scripts/probar-clientes-web.mjs`.

## Trampas / no hacer
- Propuesta: no mencionar que Juli no paga; precio solo en totales ($1.050.000 + $600.000 = $1.650.000).
- Al republicar un artifact después de compactar, pasar `url` del original o se crea un link nuevo.
- Publicar la web B2B en orden: registro habilitado en Supabase → generar con subida de precios → 3 ramas a main.
- Los precios NO van en el HTML: solo en `precios_web` (RLS).
- La plantilla de la web vive en el generador: no editar `sitio/index.html` a mano.

## Historial
- 2026-10-28: propuesta web para los jefes de zona ($1.650.000) + machete de preguntas, en moncar-web/propuesta/.
- 2026-10-27: web B2B en rama b2b-login (cuentas email + Google, pantalla Clientes, migraciones 0003/0004); MAT12 analizada.
- 2026-10-26: carpetas de la web agrupadas en moncar-web/ (backend-web + pagina-web); panel definido como backend.
- 2026-10-26: web v3 (diseño v1 + usabilidad v2) en producción; ramas rediseno / rediseno-web mergeadas.
