# -*- coding: utf-8 -*-
import sys, os, json, importlib.util
from http.server import HTTPServer, BaseHTTPRequestHandler

PORT        = 8050
DIR         = r'C:\Users\franc\OneDrive\Documentos\KeyCarbon\claude MONCAR RMH'
HTML_PATH   = os.path.join(DIR, 'dashboard_live.html')
SCRIPT_PATH = os.path.join(DIR, '.claude', 'generar_dashboard.py')

# Cargar el modulo de generacion directamente (evita lentitud de subprocess en Windows)
def _load_gen_mod():
    spec = importlib.util.spec_from_file_location("generar_dashboard", SCRIPT_PATH)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

_gen_mod = _load_gen_mod()


def regenerar():
    old_argv = sys.argv[:]
    sys.argv  = [SCRIPT_PATH, '--no-browser']
    try:
        _gen_mod.main()
        return True
    except Exception as e:
        print("Error al generar dashboard:", e)
        import traceback; traceback.print_exc()
        return False
    finally:
        sys.argv = old_argv


# Boton que se inyecta al servir el HTML (solo ASCII en bytes literal)
BOTON = (
    '<style>@media print{#_refresh_wrap{display:none!important}}</style>'
    '<div id="_refresh_wrap" style="position:fixed;top:8px;left:50%;transform:translateX(-50%);z-index:9999;display:flex;align-items:center;gap:10px">'
    '<span id="_refresh_st" style="font-size:12px;color:#aaa"></span>'
    '<button id="_refresh_btn" onclick="_doRefresh()" style="'
    'background:#2563EB;color:#fff;border:none;padding:8px 18px;'
    'border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;'
    'box-shadow:0 2px 8px rgba(0,0,0,.35)">'
    'Actualizar datos'
    '</button>'
    '<button onclick="window.print()" style="'
    'background:#16A34A;color:#fff;border:none;padding:8px 18px;'
    'border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;'
    'box-shadow:0 2px 8px rgba(0,0,0,.35)">'
    'PDF'
    '</button>'
    '</div>'
    '<script>'
    'function _doRefresh(){'
    'var btn=document.getElementById("_refresh_btn");'
    'var st=document.getElementById("_refresh_st");'
    'btn.disabled=true;btn.textContent="Actualizando...";st.textContent="";'
    'fetch("/refresh")'
    '.then(function(r){return r.json();})'
    '.then(function(d){'
    'if(d.ok){st.textContent="Listo";setTimeout(function(){location.reload();},400);}'
    'else{btn.disabled=false;btn.textContent="Actualizar datos";st.textContent="Error al generar";}'
    '})'
    '.catch(function(){'
    'btn.disabled=false;btn.textContent="Actualizar datos";st.textContent="Error de conexion";'
    '});}'
    '</script>'
    '</body>'
).encode('utf-8')


def abrir_chrome():
    import subprocess
    subprocess.Popen(
        f'start chrome http://localhost:{PORT}',
        shell=True
    )


class Handler(BaseHTTPRequestHandler):

    def do_GET(self):
        if self.path == '/refresh':
            ok = regenerar()
            if ok:
                print("Refresh ejecutado correctamente.")
            body = json.dumps({'ok': ok}).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif self.path in ('/', '/dashboard'):
            try:
                with open(HTML_PATH, 'rb') as f:
                    content = f.read()
            except FileNotFoundError:
                content = b'<html><body>Dashboard no encontrado.</body></html>'
            content = content.replace(b'</body>', BOTON, 1)
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)

        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, fmt, *args):
        pass


if __name__ == '__main__':
    os.chdir(DIR)
    print("Generando dashboard inicial...")
    regenerar()
    print(f"Servidor activo en http://localhost:{PORT}")
    print("Deja esta ventana abierta mientras uses el dashboard.")
    print("Cerrala para detener el servidor.")
    abrir_chrome()
    try:
        HTTPServer(('localhost', PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
