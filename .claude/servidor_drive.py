# -*- coding: utf-8 -*-
# Version DRIVE del servidor: sirve dashboard_drive.html, generado desde Google Sheets.
# Copia de servidor.py en otro puerto para no chocar con el dashboard original.
import sys, os, json, importlib.util, threading, time
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

PORT        = 8051
DIR         = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # raiz del proyecto (carpeta padre de .claude)
HTML_PATH   = os.path.join(DIR, 'dashboard_drive.html')
SCRIPT_PATH = os.path.join(DIR, '.claude', 'generar_dashboard_drive.py')

# Cargar el modulo de generacion directamente (evita lentitud de subprocess en Windows)
def _load_gen_mod():
    spec = importlib.util.spec_from_file_location("generar_dashboard_drive", SCRIPT_PATH)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

_gen_mod = _load_gen_mod()

# Modulos que ESCRIBEN en Sheets: entregas de mercaderia a Lita y prestamos entre zonas.
sys.path.insert(0, os.path.join(DIR, '.claude'))
import entregas_lita
import prestamos

# ThreadingHTTPServer atiende cada request en su propio hilo. regenerar() toca
# sys.argv (global de proceso) y escribe el mismo archivo de salida, asi que si
# dos requests (ej. auto-refresh + click manual) piden refrescar al mismo tiempo
# hay que serializarlas para no pisarse.
_regen_lock = threading.Lock()

def regenerar():
    with _regen_lock:
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


# Antes se inyectaban aca los botones "Actualizar datos" y "PDF" (sacados 2026-10-02)
# y un auto-refresh cada 4 minutos (sacado 2026-10-06, pedido de Franco). Los datos
# se actualizan solo con el icono de refresh de la barra superior (refreshData()).
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
            self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
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
            self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
            self.send_header('Pragma', 'no-cache')
            self.end_headers()
            self.wfile.write(content)

        else:
            self.send_response(404)
            self.end_headers()

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path not in ('/lita/entregar', '/prestamos/registrar'):
            self.send_response(404); self.end_headers(); return
        # Solo desde el propio dashboard: otra pagina abierta en el navegador
        # no puede mandar un POST con JSON a localhost sin este Origin.
        origin = self.headers.get('Origin', '')
        if origin not in ('http://localhost:%d' % PORT, 'http://127.0.0.1:%d' % PORT) \
                or 'application/json' not in self.headers.get('Content-Type', ''):
            self._json(403, {'ok': False, 'error': 'Origen no permitido'}); return
        try:
            largo = int(self.headers.get('Content-Length', 0))
            datos = json.loads(self.rfile.read(largo).decode('utf-8'))
            if self.path == '/lita/entregar':
                # Acepta una lista {items:[{cod,cantidad},...]} o un solo {cod,cantidad}.
                items = datos.get('items') or [{'cod': datos.get('cod'), 'cantidad': datos.get('cantidad')}]
                res = entregas_lita.entregar_varios(items)
                print("Lita: entregados %d codigos, %d uds (%s)" % (
                    res['codigos'], res['uds'], ', '.join('%s x %d %s' % (i['cod'], i['entregado'], i['remito']) for i in res['items'])))
            else:
                res = prestamos.registrar(datos.get('op'), datos.get('persona'), datos.get('fecha'),
                                          datos.get('items') or [], datos.get('comentario', ''))
                print("Prestamos: %s %s (%d filas)" % (res['op'], res['persona'], res['filas']))
        except ValueError as e:
            self._json(400, {'ok': False, 'error': str(e)}); return
        except Exception as e:
            import traceback; traceback.print_exc()
            self._json(500, {'ok': False, 'error': str(e)}); return
        # La lectura CSV (gviz) puede tardar un instante en reflejar la escritura.
        time.sleep(2)
        regenerar()
        self._json(200, {'ok': True, **res})

    def log_message(self, fmt, *args):
        pass


if __name__ == '__main__':
    os.chdir(DIR)
    print("Generando dashboard inicial (desde Google Drive)...")
    regenerar()
    print(f"Servidor activo en http://localhost:{PORT}")
    print("Deja esta ventana abierta mientras uses el dashboard.")
    print("Cerrala para detener el servidor.")
    abrir_chrome()
    try:
        srv = ThreadingHTTPServer(('localhost', PORT), Handler)
        srv.daemon_threads = True
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
