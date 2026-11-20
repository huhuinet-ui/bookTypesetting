import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from docx import Document
from PIL import Image, features
import book_layout as book


class WordImageCompatibilityTests(unittest.TestCase):
    def setUp(self):
        root = Path(__file__).resolve().parents[1] / 'tmp' / 'word-image-tests'
        root.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=str(root))
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    @unittest.skipUnless(features.check('webp'), 'Pillow has no WebP codec')
    def test_webp_with_jpg_extension_exports_as_png(self):
        path = self.root / '图片.jpg'
        Image.new('RGBA', (40, 30), (200, 10, 0, 120)).save(str(path), format='WEBP')
        before = path.read_bytes()
        doc = Document()
        book.add_image(doc, str(path), '图注')
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        loaded = Document(stream)
        self.assertEqual(len(loaded.inline_shapes), 1)
        self.assertEqual(path.read_bytes(), before)
        parts = [r.target_part for r in loaded.part.rels.values() if r.reltype.endswith('/image')]
        self.assertEqual(parts[0].content_type, 'image/png')
        with Image.open(io.BytesIO(parts[0].blob)) as im:
            self.assertEqual(im.size, (40, 30))
            self.assertEqual(im.mode, 'RGBA')

    def test_native_png_is_kept(self):
        path = self.root / 'native.png'
        Image.new('RGB', (20, 30)).save(str(path))
        self.assertEqual(book.word_image_source(str(path)), str(path))

    def test_jpeg_without_jfif_stays_jpeg_and_is_reused_across_passes(self):
        path = self.root / 'photo.jpg'
        image = io.BytesIO()
        Image.new('RGB', (120, 80), 'red').save(image, format='JPEG')
        raw = image.getvalue()
        self.assertEqual(raw[2:4], b'\xff\xe0')
        # Camera JPEGs without JFIF/Exif headers are readable by Pillow but not docx.
        length = int.from_bytes(raw[4:6], 'big')
        path.write_bytes(raw[:2] + raw[4+length:])
        before = path.read_bytes()
        with patch.object(book, '_word_image_source', wraps=book._word_image_source) as normalize:
            with book.image_source_cache():
                source = book.word_image_source(str(path))
                self.assertEqual(book.word_image_source(str(path)), source)
                self.assertEqual(normalize.call_count, 1)
                with Image.open(source) as result:
                    self.assertEqual(result.format, 'JPEG')
                    self.assertEqual(result.size, (120, 80))
                doc = Document()
                doc.add_paragraph().add_run().add_picture(source)
                self.assertEqual(len(doc.inline_shapes), 1)
            self.assertFalse(Path(source).exists())
        self.assertEqual(path.read_bytes(), before)

    def test_corrupt_image_error_names_file(self):
        path = self.root / 'broken.jpg'
        path.write_bytes(b'broken image')
        with self.assertRaisesRegex(RuntimeError, 'broken.jpg'):
            book.word_image_source(str(path))
