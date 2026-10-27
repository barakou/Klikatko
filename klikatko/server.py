import argparse
import base64
import json
import mimetypes
import os
import secrets
import sys
import threading
import webbrowser
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from uuid import uuid4

from .export import generate, write_version
from .model import ROOT, Store, catalog, safe_name


class Application:
    def __init__(self, directory):
        self.store = Store(directory)
        self.token = secrets.token_urlsafe(32)
        self.downloads = {}
        (self.store.root / 'reports').mkdir(exist_ok=True)
        (self.store.root / 'images').mkdir(exist_ok=True)

    def state(self):
        return {'catalog': catalog(), 'projects': self.store.list('projects'),
                'presets': self.store.list('presets'), 'output_dir': str(self.store.root / 'reports'),
                'today': date.today().isoformat()}


def make_handler(app):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            pass

        def send(self, status, body, content_type='application/json; charset=utf-8', extra=None):
            if isinstance(body, (dict, list)):
                body = json.dumps(body, ensure_ascii=False).encode('utf-8')
            elif isinstance(body, str):
                body = body.encode('utf-8')
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Referrer-Policy', 'no-referrer')
            self.send_header('Content-Security-Policy', "default-src 'self'; img-src 'self' blob: data:; style-src 'self'; script-src 'self'; object-src 'none'; frame-ancestors 'self'")
            for key, value in (extra or {}).items():
                self.send_header(key, value)
            self.end_headers()
            self.wfile.write(body)

        def authorized(self):
            return secrets.compare_digest(self.headers.get('X-Klikatko-Token', ''), app.token)

        def do_GET(self):
            try:
                url = urlparse(self.path)
                if url.path in ('/', '/index.html'):
                    html = (ROOT / 'klikatko/web/index.html').read_text(encoding='utf-8').replace('__TOKEN__', app.token)
                    return self.send(200, html, 'text/html; charset=utf-8')
                if url.path in ('/app.js', '/style.css'):
                    path = ROOT / 'klikatko/web' / url.path[1:]
                    return self.send(200, path.read_bytes(), 'text/javascript; charset=utf-8' if path.suffix == '.js' else 'text/css; charset=utf-8')
                if url.path == '/favicon.ico':
                    return self.send(204, b'')
                if not self.authorized():
                    return self.send(403, {'error': 'Obnovte stránku aplikace.'})
                query = parse_qs(url.query)
                if url.path == '/api/state':
                    return self.send(200, app.state())
                if url.path == '/api/folders':
                    path = Path(query.get('path', [str(Path.home())])[0]).expanduser().resolve()
                    if not path.is_dir():
                        raise ValueError('Složka neexistuje.')
                    children = []
                    for item in path.iterdir():
                        try:
                            if item.is_dir() and not item.name.startswith('.'):
                                children.append({'name': item.name, 'path': str(item)})
                        except OSError:
                            continue
                    drives = [f'{letter}:\\' for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' if Path(f'{letter}:\\').exists()] if os.name == 'nt' else ['/']
                    return self.send(200, {'path': str(path), 'parent': str(path.parent), 'children': sorted(children, key=lambda x: x['name'].casefold()), 'drives': drives})
                if url.path == '/api/download':
                    item = app.downloads.get(query.get('id', [''])[0])
                    if not item:
                        return self.send(404, {'error': 'Soubor už není dostupný. Vygenerujte jej znovu.'})
                    from urllib.parse import quote
                    return self.send(200, item.read_bytes(), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', {'Content-Disposition': "attachment; filename*=UTF-8''" + quote(item.name)})
                return self.send(404, {'error': 'Nenalezeno.'})
            except (ValueError, OSError) as exc:
                self.send(400, {'error': str(exc)})

        def do_POST(self):
            if not self.authorized():
                return self.send(403, {'error': 'Obnovte stránku aplikace.'})
            try:
                size = int(self.headers.get('Content-Length', 0))
                if size > 16_000_000 or size < 1:
                    raise ValueError('Požadavek je prázdný nebo příliš velký (maximum 12 MB pro obrázek).')
                data = json.loads(self.rfile.read(size))
                route = urlparse(self.path).path
                if route in ('/api/projects', '/api/presets'):
                    return self.send(200, app.store.save(route.rsplit('/', 1)[1], data))
                if route == '/api/delete':
                    app.store.delete(data['kind'], data['id'])
                    return self.send(200, {'ok': True})
                if route == '/api/image':
                    from PIL import Image
                    import io
                    content = base64.b64decode(data['content'], validate=True)
                    with Image.open(io.BytesIO(content)) as image:
                        if image.width * image.height > 40_000_000:
                            raise ValueError('Obrázek je příliš velký (maximum 40 megapixelů).')
                        image.verify()
                    path = app.store.root / 'images' / (uuid4().hex + '.img')
                    path.write_bytes(content)
                    return self.send(200, {'path': str(path)})
                if route == '/api/generate':
                    content = generate(data)
                    path = write_version(content, data.get('output_dir') or app.store.root / 'reports', data.get('filename') or safe_name(data.get('name', '')))
                    ident = uuid4().hex
                    app.downloads[ident] = path
                    return self.send(200, {'id': ident, 'name': path.name, 'path': str(path), 'size': len(content)})
                return self.send(404, {'error': 'Nenalezeno.'})
            except (ValueError, OSError, KeyError, TypeError) as exc:
                self.send(400, {'error': str(exc)})
            except Exception:
                import traceback
                traceback.print_exc()
                self.send(500, {'error': 'Dokument se nepodařilo vytvořit. Zkontrolujte předlohu a oprávnění ke složce.'})
    return Handler


def main():
    parser = argparse.ArgumentParser(description='Klikátko – příprava Wordové zprávy')
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--data-dir', default=os.environ.get('KLIKATKO_DATA_DIR', str(Path.home() / 'Klikatko-data')))
    args = parser.parse_args()
    server = ThreadingHTTPServer(('127.0.0.1', args.port), make_handler(Application(args.data_dir)))
    url = f'http://127.0.0.1:{server.server_port}'
    print(f'Klikátko: {url}\nPro ukončení stiskněte Ctrl+C.', flush=True)
    if not args.no_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
