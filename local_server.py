"""Launch the local typesetting UI and serialize native Word export jobs."""
import hashlib
from collections import OrderedDict
import base64
import io
import shutil
import zipfile
import json
import re
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
CONVERSION_RESULTS = OrderedDict()
CONVERSION_RESULTS_LOCK = threading.Lock()
EXPORT_UPLOADS = OrderedDict()
THEME_HISTORY_FILE = ROOT / '.theme-colors.json'
THEME_HISTORY_LIMIT = 24
THEME_HISTORY_LOCK = threading.Lock()

def normalize_theme_colors(colors):
    result = []
    if not isinstance(colors, list):
        return result
    for color in colors:
        if isinstance(color, str) and re.fullmatch(r'#[0-9a-fA-F]{6}', color.strip()):
            color = color.strip().upper()
            if color not in result:
                result.append(color)
                if len(result) == THEME_HISTORY_LIMIT:
                    break
    return result

def read_theme_colors():
    try:
        return normalize_theme_colors(json.loads(THEME_HISTORY_FILE.read_text(encoding='utf-8')))
    except (OSError, ValueError, TypeError):
        return []

def remember_theme_color(color):
    valid = normalize_theme_colors([color])
    if not valid:
        raise ValueError('无效的主题色')
    with THEME_HISTORY_LOCK:
        colors = normalize_theme_colors(valid + read_theme_colors())
        temporary = THEME_HISTORY_FILE.with_suffix('.json.tmp')
        temporary.write_text(json.dumps(colors), encoding='utf-8')
        temporary.replace(THEME_HISTORY_FILE)
    return colors

def safe_path(root, name):
    path = (root / name.replace('\\', '/')).resolve()
    try:
        relative = path.relative_to(root.resolve())
    except ValueError:
        raise ValueError('文件路径超出导入目录') from None
    if not relative.parts:
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
        if path == '/api/theme-colors':
            return self.reply(200, json.dumps({'colors': read_theme_colors()}).encode(), 'application/json')
        if path == '/api/session':
            return self.reply(200, json.dumps({'token': TOKEN, 'wordBinaryUpload': True, 'wordFileResults': True, 'wordBinaryExport': True}).encode(), 'application/json')
        if path.startswith('/api/word-result/'):
            parts = path.split('/')
            with CONVERSION_RESULTS_LOCK:
                result = CONVERSION_RESULTS.get(parts[3]) if len(parts) == 5 else None
                file = result[1].get(parts[4]) if result else None
                if file is None:
                    return self.reply(404, b'Conversion result expired', 'text/plain')
                # Open before releasing the cache lock so cleanup cannot race retrieval.
                stream = file.open('rb')
            with stream:
                self.send_response(200)
                self.send_header('Content-Type', 'application/zip' if parts[4] == 'archive' else 'application/octet-stream')
                self.send_header('Content-Length', str(file.stat().st_size))
                self.send_header('Cache-Control', 'no-store')
                if parts[4] == 'archive':
                    self.send_header('Content-Disposition', 'attachment; filename="Markdown.zip"')
                self.end_headers()
                shutil.copyfileobj(stream, self.wfile, 1024 * 1024)
            return
        allowed = {'/': '排版工具.html', '/排版工具.html': '排版工具.html', '/word-export.js': 'word-export.js', '/word-import.js': 'word-import.js', '/project-format.js': 'project-format.js', '/project-settings.js': 'project-settings.js'}
        if path.startswith('/vendor/') and Path(path).name in {'html-to-image-1.11.13.js', 'pdf-lib-1.17.1.min.js'}:
            name = path.lstrip('/')
        else:
            name = allowed.get(path)
        if not name:
            return self.reply(404, b'Not found', 'text/plain')
        self.reply(200, (ROOT / name).read_bytes(), 'text/html; charset=utf-8' if name.endswith('.html') else 'application/javascript; charset=utf-8')

    def convert_word(self, data, upload_length=None):
        name = Path(data['name'].replace('\\', '/')).name
        if not name.lower().endswith('.docx'):
            raise ValueError('请选择 .docx 文件；旧版 .doc 请先另存为 .docx')
        with tempfile.TemporaryDirectory(prefix='word_to_md_') as temp:
            root = Path(temp)
            source = root / 'input.docx'
            if upload_length is None:
                source.write_bytes(base64.b64decode(data['data'], validate=True))
            else:
                remaining = upload_length
                with source.open('wb') as target:
                    while remaining:
                        chunk = self.rfile.read(min(1024 * 1024, remaining))
                        if not chunk:
                            raise ValueError('Word 文件上传不完整，请重新选择文件后重试')
                        target.write(chunk)
                        remaining -= len(chunk)
            output = root / (Path(name).stem + '.md')
            command = [sys.executable, str(ROOT / 'word2md.py'), str(source), str(output)]
            if data.get('bracketSubtitle'):
                command.append('--bracket-subtitle')
            result = subprocess.run(command, cwd=ROOT, capture_output=True, timeout=900)
            if result.returncode:
                raise ValueError((result.stderr or result.stdout).decode('utf-8', 'replace')[-4000:])
            report = result.stdout.decode('utf-8', 'replace')
            # Hide temporary paths in the user-facing conversion report.
            report = report.replace(str(root), '转换结果')
            files = [output] + sorted((root / '图').glob('*'))
            if data.get('fileResults'):
                saved = tempfile.TemporaryDirectory(prefix='word_result_')
                try:
                    result_root = Path(saved.name)
                    key = secrets.token_urlsafe(24)
                    base = '/api/word-result/' + key + '/'
                    paths, manifest = {}, []
                    for index, file in enumerate(files):
                        relative = file.relative_to(root).as_posix()
                        destination = result_root / str(index)
                        shutil.move(str(file), destination)
                        paths[str(index)] = destination
                        manifest.append({'path': relative, 'url': base + str(index)})
                    archive_path = result_root / 'Markdown.zip'
                    with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as bundle:
                        for index, entry in enumerate(manifest):
                            bundle.write(paths[str(index)], entry['path'])
                        bundle.writestr('转换报告.txt', report)
                    paths['archive'] = archive_path
                    with CONVERSION_RESULTS_LOCK:
                        CONVERSION_RESULTS[key] = (saved, paths)
                        while len(CONVERSION_RESULTS) > 2:
                            _, expired = CONVERSION_RESULTS.popitem(last=False)
                            expired[0].cleanup()
                    self.reply(200, json.dumps({'files': manifest, 'report': report, 'archiveUrl': base + 'archive'}, ensure_ascii=False).encode('utf-8'), 'application/json; charset=utf-8')
                except Exception:
                    with CONVERSION_RESULTS_LOCK:
                        CONVERSION_RESULTS.pop(key, None)
                    saved.cleanup()
                    raise
                return
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
        if self.path not in {'/api/export-word', '/api/export-file', '/api/convert-word', '/api/theme-colors'} or self.headers.get('Origin') != origin or self.headers.get('X-Export-Token') != TOKEN:
            return self.reply(403, b'Forbidden', 'text/plain')
        if self.path == '/api/theme-colors':
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 < length <= 1024:
                    raise ValueError('无效的主题色请求')
                data = json.loads(self.rfile.read(length))
                colors = remember_theme_color(data.get('color'))
                return self.reply(200, json.dumps({'colors': colors}).encode(), 'application/json')
            except Exception as error:
                return self.reply(400, str(error).encode('utf-8'), 'text/plain; charset=utf-8')
        if not LOCK.acquire(blocking=False):
            return self.reply(409, '已有 Word 导出任务，请稍后再试'.encode(), 'text/plain; charset=utf-8')
        export_upload_id = None
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if self.path == '/api/export-file':
                upload_id = self.headers.get('X-Upload-Id', '')
                index = self.headers.get('X-File-Index', '')
                if not re.fullmatch(r'[a-zA-Z0-9-]{16,80}', upload_id) or not re.fullmatch(r'\d{1,6}', index):
                    raise ValueError('无效的上传标识')
                if not 0 <= length <= 1024 * 1024 * 1024:
                    raise ValueError('单个文件不能超过 1 GB')
                if upload_id not in EXPORT_UPLOADS:
                    EXPORT_UPLOADS[upload_id] = tempfile.TemporaryDirectory(prefix='word_upload_')
                    while len(EXPORT_UPLOADS) > 2:
                        _, old = EXPORT_UPLOADS.popitem(last=False)
                        old.cleanup()
                target = Path(EXPORT_UPLOADS[upload_id].name) / index
                remaining = length
                with target.open('wb') as output:
                    while remaining:
                        chunk = self.rfile.read(min(1024 * 1024, remaining))
                        if not chunk:
                            raise ValueError('文件上传中断，请重试')
                        output.write(chunk)
                        remaining -= len(chunk)
                return self.reply(200, b'{}', 'application/json')
            if self.path == '/api/convert-word' and self.headers.get('Content-Type', '').split(';')[0] == 'application/octet-stream':
                from urllib.parse import unquote
                if not 0 < length <= 1024 * 1024 * 1024:
                    raise ValueError('Word 文件过大（上限 1 GB）')
                return self.convert_word({'name': unquote(self.headers.get('X-Word-Name', '')), 'bracketSubtitle': self.headers.get('X-Bracket-Subtitle') == 'true', 'fileResults': self.headers.get('X-Word-Results') == 'files'}, upload_length=length)
            if not 0 < length <= MAX_UPLOAD:
                raise ValueError('导入内容过大（上限 300 MB）')
            data = json.loads(self.rfile.read(length))
            if self.path == '/api/convert-word':
                return self.convert_word(data)
            export_upload_id = data.get('uploadId')
            if export_upload_id and export_upload_id not in EXPORT_UPLOADS:
                raise ValueError('上传文件已失效，请重新导出')
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
                    if export_upload_id:
                        index = str(file.get('index', ''))
                        if not re.fullmatch(r'\d{1,6}', index):
                            raise ValueError('无效的文件索引')
                        shutil.copyfile(Path(EXPORT_UPLOADS[export_upload_id].name) / index, target)
                    else:
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
                    details = (result.stderr or result.stdout).decode('utf-8', 'replace')
                    # Show the actual cause before the stack so browser dialogs keep it visible.
                    cause = details.rsplit('RuntimeError: ', 1)[-1].strip()
                    raise RuntimeError(cause[-4000:])
                if output.stat().st_size <= CACHE_LIMIT:
                    content = output.read_bytes()
                    EXPORT_CACHE[cache_key] = content
                    while len(EXPORT_CACHE) > 2 or sum(map(len, EXPORT_CACHE.values())) > CACHE_LIMIT:
                        EXPORT_CACHE.popitem(last=False)
                    self.reply(200, content, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                else:
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                    self.send_header('Content-Length', str(output.stat().st_size))
                    self.send_header('Cache-Control', 'no-store')
                    self.end_headers()
                    with output.open('rb') as stream:
                        shutil.copyfileobj(stream, self.wfile, 1024 * 1024)
        except Exception as error:
            self.reply(400, str(error).encode('utf-8'), 'text/plain; charset=utf-8')
        finally:
            if export_upload_id in EXPORT_UPLOADS:
                EXPORT_UPLOADS.pop(export_upload_id).cleanup()
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
