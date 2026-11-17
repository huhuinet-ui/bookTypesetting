import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from docx import Document
from PIL import Image
import book_layout as book


class FrontImageTests(unittest.TestCase):
    def test_front_matter_images_are_embedded_with_caption_and_text_order(self):
        root=Path(__file__).resolve().parents[1]/'tmp'/'front-image-tests';root.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as directory:
            path=Path(directory);(path/'图').mkdir()
            Image.new('RGB',(80,60),'red').save(str(path/'图'/'照片.png'))
            for text in ('前面的文字\n\n![照片说明](图/照片.png)\n\n后面的文字',
                         '前面的文字\n\n图片：图/照片.png\n图片名：照片说明\n\n后面的文字'):
                doc=Document()
                with patch.object(book,'SRC',str(path/'source.md')):
                    book.add_front_page(doc,'前言',text)
                out=io.BytesIO();doc.save(out);out.seek(0);loaded=Document(out)
                self.assertEqual(len(loaded.inline_shapes),1)
                paragraphs=[p.text for p in loaded.paragraphs if p.text]
                self.assertEqual(paragraphs[-3:],['前面的文字','照片说明','后面的文字'])
                self.assertNotIn('![', ''.join(paragraphs))

    def test_linebreaks_remain_distinct_from_paragraphs(self):
        doc=Document();book.add_front_page(doc,'前言','甲\n乙\n\n丙')
        self.assertEqual([p.text for p in doc.paragraphs if p.text],['前言','甲\n乙','丙'])
