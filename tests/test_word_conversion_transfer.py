import io
import json
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import local_server as server


class ConversionTransferTests(unittest.TestCase):
    def tearDown(self):
        for saved, _ in server.CONVERSION_RESULTS.values():
            saved.cleanup()
        server.CONVERSION_RESULTS.clear()
        for saved in server.EXPORT_UPLOADS.values():
            saved.cleanup()
        server.EXPORT_UPLOADS.clear()

    def test_word_export_uses_binary_assets_without_base64(self):
        upload_id = 'test-upload-123456789'
        handler = object.__new__(server.Handler)
        handler.server = SimpleNamespace(server_port=12345)
        replies = []
        handler.reply = lambda *args: replies.append(args)

        def post(path, content, extra=None):
            handler.path = path
            handler.rfile = io.BytesIO(content)
            handler.headers = {'Origin': 'http://127.0.0.1:12345', 'X-Export-Token': server.TOKEN,
                               'Content-Length': str(len(content)), **(extra or {})}
            handler.do_POST()
            self.assertEqual(replies[-1][0], 200, replies[-1])

        for index, content in enumerate([b'# manuscript', b'image content']):
            post('/api/export-file', content, {'X-Upload-Id': upload_id, 'X-File-Index': str(index)})

        def export(command, **kwargs):
            config = json.loads(Path(command[-1]).read_text(encoding='utf-8'))
            self.assertEqual(Path(config['source']).read_bytes(), b'# manuscript')
            self.assertEqual((Path(config['root']) / 'book/images/1.jpg').read_bytes(), b'image content')
            Path(config['output']).write_bytes(b'word result')
            return SimpleNamespace(returncode=0, stdout=b'', stderr=b'')

        data = {'source': 'book/book.md', 'name': 'book', 'params': {}, 'uploadId': upload_id,
                'files': [{'path': 'book/book.md', 'index': 0}, {'path': 'book/images/1.jpg', 'index': 1}]}
        with patch.object(server.subprocess, 'run', side_effect=export):
            post('/api/export-word', json.dumps(data).encode())
        self.assertEqual(replies[-1][1], b'word result')
        self.assertNotIn(upload_id, server.EXPORT_UPLOADS)

    def test_binary_upload_returns_files_and_disk_archive(self):
        handler = object.__new__(server.Handler)
        handler.rfile = io.BytesIO(b'uploaded docx')
        replies = []
        handler.reply = lambda *args: replies.append(args)

        def convert(command, **kwargs):
            self.assertEqual(Path(command[2]).read_bytes(), b'uploaded docx')
            Path(command[3]).write_text('# 书稿', encoding='utf-8')
            images = Path(command[3]).parent / '图'
            images.mkdir()
            (images / 'img1.jpg').write_bytes(b'image bytes')
            return SimpleNamespace(returncode=0, stdout=b'converted', stderr=b'')

        with patch.object(server.subprocess, 'run', side_effect=convert):
            handler.convert_word({'name': '书稿.docx', 'fileResults': True}, upload_length=13)
        status, raw, _ = replies[0]
        self.assertEqual(status, 200)
        manifest = json.loads(raw)
        self.assertNotIn('archive', manifest)
        self.assertTrue(all('data' not in entry for entry in manifest['files']))
        self.assertLess(len(raw), 2048)
        key = manifest['archiveUrl'].split('/')[-2]
        _, paths = server.CONVERSION_RESULTS[key]
        self.assertEqual(paths['1'].read_bytes(), b'image bytes')
        with zipfile.ZipFile(paths['archive']) as archive:
            self.assertEqual(archive.read('图/img1.jpg'), b'image bytes')
            self.assertEqual(archive.read('书稿.md').decode(), '# 书稿')


if __name__ == '__main__':
    unittest.main()
