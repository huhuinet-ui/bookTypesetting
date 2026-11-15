import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from PIL import Image
from word2md import convert


class ImageOrientationTests(unittest.TestCase):
    def fixture(self):
        image = Image.new('RGB', (60, 40), 'red')
        for x in range(30, 60):
            for y in range(40):
                image.putpixel((x, y), (0, 0, 255))
        stream = io.BytesIO()
        image.save(stream, format='PNG')
        return stream.getvalue()

    def extract(self, blob, rotations=(0,), flip=False):
        root = Path(__file__).resolve().parents[1] / 'tmp' / 'orientation-tests'
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as temp:
            doc = Document()
            for angle in rotations:
                shape = doc.add_paragraph().add_run().add_picture(io.BytesIO(blob))
                transform = shape._inline.find('.//' + qn('a:xfrm'))
                transform.set('rot', str(angle * 60000))
                if flip:
                    transform.set('flipH', '1')
            source = Path(temp) / 'input.docx'
            output = Path(temp) / 'output.md'
            doc.save(str(source))
            with contextlib.redirect_stdout(io.StringIO()):
                convert(str(source), str(output))
            return [(p.name, p.read_bytes()) for p in sorted((Path(temp) / '图').iterdir())]

    def test_unrotated_image_is_unchanged(self):
        blob = self.fixture()
        self.assertEqual(self.extract(blob)[0][1], blob)

    def test_same_source_keeps_each_word_rotation(self):
        files = self.extract(self.fixture(), (90, 270))
        for (_, blob), top_color in zip(files, [(255, 0, 0), (0, 0, 255)]):
            with Image.open(io.BytesIO(blob)) as im:
                self.assertEqual(im.size, (40, 60))
                self.assertEqual(im.getpixel((20, 5))[:3], top_color)

    def test_horizontal_flip_is_preserved(self):
        blob = self.extract(self.fixture(), flip=True)[0][1]
        with Image.open(io.BytesIO(blob)) as im:
            self.assertEqual(im.getpixel((5, 20))[:3], (0, 0, 255))

    def test_exif_orientation_is_baked_without_double_rotation(self):
        with Image.open(io.BytesIO(self.fixture())) as im:
            exif = im.getexif()
            exif[274] = 6
            stream = io.BytesIO()
            im.save(stream, format='JPEG', exif=exif.tobytes())
        name, blob = self.extract(stream.getvalue())[0]
        self.assertTrue(name.endswith('.png'))
        with Image.open(io.BytesIO(blob)) as im:
            self.assertEqual(im.size, (40, 60))
            self.assertIn(im.getexif().get(274, 1), (1, None))
            self.assertGreater(im.getpixel((20, 5))[0], 240)
