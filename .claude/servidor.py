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


# Antes se inyectaban aca los botones "Actualizar datos" y "PDF". Desde 2026-10-02 el
# refresh es un icono de la barra superior del TEMPLATE (refreshData()) y el PDF se saco.
BOTON = b'</body>'


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
