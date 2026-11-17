import io
import unittest
from unittest.mock import patch
from docx import Document
from docx.shared import RGBColor
import book_layout as book


class TitlePageColorTests(unittest.TestCase):
    def check_color(self,custom,theme,expected):
        doc=Document()
        with patch.object(book,'TITLE_PAGE_COLOR',custom), patch.object(book,'BOOK_COLOR',theme):
            book.add_title_page(doc,'书名',author='作者',text='扉页正文\n第二行')
        stream=io.BytesIO();doc.save(stream);stream.seek(0)
        runs=[r for p in Document(stream).paragraphs for r in p.runs if r.text]
        self.assertEqual([r.text for r in runs],['书名','作者','扉页正文','第二行'])
        self.assertTrue(all(str(r.font.color.rgb)==expected for r in runs))

    def test_theme_color_follows_current_theme(self):
        for color in ('123456','ABCDEF'):
            self.check_color(None,RGBColor.from_string(color),color)

    def test_custom_color_overrides_theme_including_black(self):
        for color in ('C00000','000000'):
            self.check_color(RGBColor.from_string(color),RGBColor.from_string('ABCDEF'),color)
