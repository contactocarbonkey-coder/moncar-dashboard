@echo off
echo.
echo === MONCAR RMH - Actualizar sitio web ===
echo.

cd /d "%~dp0"

echo [1/3] Generando sitio desde Google Sheets...
python generar_sitio_drive.py
if errorlevel 1 (
    echo ERROR al generar el sitio.
    pause
    exit /b 1
)

echo.
echo [2/3] Subiendo cambios a GitHub...
cd sitio
if not exist ".git" (
    echo ERROR: la carpeta "sitio" todavia no es un repo git conectado a GitHub/Vercel.
    echo        Hay que configurarlo una vez ^(ver moncar-web\pagina-web\CLAUDE.md o pedirselo a Claude^).
    pause
    exit /b 1
)
git add index.html
git commit -m "Actualizar catalogo"
git push

echo.
echo [3/3] Listo! Vercel detecta el cambio y publica en ~30 segundos.
echo       Tu sitio: https://moncar-rmh.vercel.app
echo.
pause
