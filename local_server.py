"""Launch the local typesetting UI and serialize native Word export jobs."""
import hashlib
from collections import OrderedDict
import base64
import io
import zipfile
import json
import secrets
import subprocess
import sys
import tempfile
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent
TOKEN = secrets.token_urlsafe(32)
LOCK = threading.Lock()
EXPORT_CACHE = OrderedDict()
CACHE_LIMIT = 64 * 1024 * 1024
MAX_UPLOAD = 300 * 1024 * 1024

def safe_path(root, name):
    path = (root / name.replace('\\', '/')).resolve()
    if not path.is_relative_to(root.resolve()) or path == root.resolve():
        raise ValueError('文件路径超出导入目录')
    return path

class Handler(BaseHTTPRequestHandler):
    def reply(self, status, data, content_type):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = urlparse(self.path).path
        if self.headers.get('Host') != f'127.0.0.1:{self.server.server_port}':
            return self.reply(403, b'Forbidden', 'text/plain')
        if path == '/api/session':
            return self.reply(200, json.dumps({'token': TOKEN}).encode(), 'application/json')
        allowed = {'/': '排版工具.html', '/排版工具.html': '排版工具.html', '/word-export.js': 'word-export.js', '/word-import.js': 'word-import.js'}
        if path.startswith('/vendor/') and Path(path).name in {'html-to-image-1.11.13.js', 'pdf-lib-1.17.1.min.js'}:
            name = path.lstrip('/')
        else:
            name = allowed.get(path)
        if not name:
            return self.reply(404, b'Not found', 'text/plain')
        self.reply(200, (ROOT / name).read_bytes(), 'text/html; charset=utf-8' if name.endswith('.html') else 'application/javascript; charset=utf-8')

    def convert_word(self, data):
        name = Path(data['name'].replace('\\', '/')).name
        if not name.lower().endswith('.docx'):
            raise ValueError('请选择 .docx 文件；旧版 .doc 请先另存为 .docx')
        with tempfile.TemporaryDirectory(prefix='word_to_md_') as temp:
            root = Path(temp)
            source = root / 'input.docx'
            source.write_bytes(base64.b64decode(data['data'], validate=True))
            output = root / (Path(name).stem + '.md')
            command = [sys.executable, str(ROOT / 'word2md.py'), str(source), str(output)]
            if data.get('bracketSubtitle'):
                command.append('--bracket-subtitle')
            result = subprocess.run(command, cwd=ROOT, capture_output=True, timeout=300)
            if result.returncode:
                raise ValueError((result.stderr or result.stdout).decode('utf-8', 'replace')[-4000:])
            report = result.stdout.decode('utf-8', 'replace')
            # Hide temporary paths in the user-facing conversion report.
            report = report.replace(str(root), '转换结果')
            files = [output] + sorted((root / '图').glob('*'))
            archive = io.BytesIO()
            payload = []
            with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as bundle:
                for file in files:
                    relative = file.relative_to(root).as_posix()
                    content = file.read_bytes()
                    bundle.writestr(relative, content)
                    payload.append({'path': relative, 'data': base64.b64encode(content).decode('ascii')})
                bundle.writestr('转换报告.txt', report)
            self.reply(200, json.dumps({'files': payload, 'report': report, 'archive': base64.b64encode(archive.getvalue()).decode('ascii')}, ensure_ascii=False).encode('utf-8'), 'application/json; charset=utf-8')

    def do_POST(self):
        origin = f'http://127.0.0.1:{self.server.server_port}'
        if self.path not in {'/api/export-word', '/api/convert-word'} or self.headers.get('Origin') != origin or self.headers.get('X-Export-Token') != TOKEN:
            return self.reply(403, b'Forbidden', 'text/plain')
        if not LOCK.acquire(blocking=False):
            return self.reply(409, '已有 Word 导出任务，请稍后再试'.encode(), 'text/plain; charset=utf-8')
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= MAX_UPLOAD:
                raise ValueError('导入内容过大（上限 300 MB）')
            data = json.loads(self.rfile.read(length))
            if self.path == '/api/convert-word':
                return self.convert_word(data)
            # Cache only successful complete exports; key includes files, styles and artwork.
            cache_key = hashlib.sha256(json.dumps(data, sort_keys=True, separators=(',', ':')).encode()).digest()
            if cache_key in EXPORT_CACHE:
                EXPORT_CACHE.move_to_end(cache_key)
                return self.reply(200, EXPORT_CACHE[cache_key], 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            with tempfile.TemporaryDirectory(prefix='book_word_') as temp:
                root = Path(temp)
                for file in data['files']:
                    target = safe_path(root, file['path'])
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(base64.b64decode(file['data'], validate=True))
                source = safe_path(root, data['source'])
                if source.suffix.lower() != '.md' or not source.is_file():
                    raise ValueError('未找到主书稿')
                artwork = []
                for index, encoded in enumerate(data.get('chapterArtwork', [])):
                    image = root / f'_chapter_{index}.png'
                    image.write_bytes(base64.b64decode(encoded, validate=True))
                    artwork.append(str(image))
                config = root / '_export_config.json'
                output = root / '_export_result.docx'
                config.write_text(json.dumps(dict(source=str(source), output=str(output), root=str(root), name=data['name'], params=data['params'], chapterArtwork=artwork), ensure_ascii=False), encoding='utf-8')
                result = subprocess.run([sys.executable, str(ROOT / 'word_worker.py'), str(config)], cwd=ROOT, capture_output=True, timeout=1800)
                if result.returncode:
                    raise RuntimeError((result.stderr or result.stdout).decode('utf-8', 'replace')[-4000:])
                content = output.read_bytes()
                if len(content) <= CACHE_LIMIT:
                    EXPORT_CACHE[cache_key] = content
                    while len(EXPORT_CACHE) > 2 or sum(map(len, EXPORT_CACHE.values())) > CACHE_LIMIT:
                        EXPORT_CACHE.popitem(last=False)
                self.reply(200, content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        except Exception as error:
            self.reply(400, str(error).encode('utf-8'), 'text/plain; charset=utf-8')
        finally:
            LOCK.release()

if __name__ == '__main__':
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    url = f'http://127.0.0.1:{server.server_port}/'
    print('排版工具：' + url + '\n请保留此窗口；关闭窗口将停止 Word 导出。', flush=True)
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()
