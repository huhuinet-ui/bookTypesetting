import io
import unittest
from unittest.mock import patch
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
import book_layout as book


class SubtitleColorTests(unittest.TestCase):
    def test_default_subtitle_is_black_without_background(self):
        p = Document().add_paragraph()
        r = book.add_subtitle(p, '小标题')
        self.assertEqual(str(r.font.color.rgb), '000000')
        self.assertIsNone(r._r.rPr.find(qn('w:shd')))

    def test_saved_word_preserves_subtitle_colors(self):
        doc = Document()
        with patch.object(book, 'SUBTITLE_COLOR', RGBColor.from_string('C00000')), \
             patch.object(book, 'SUBTITLE_BG_ON', True), \
             patch.object(book, 'SUBTITLE_BG_COLOR', 'FFF2CC'):
            book.add_subtitle(doc.add_paragraph(), '彩色小标题')
            book.add_body_para(doc.add_paragraph(), '普通正文')
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        loaded = Document(output)
        run = loaded.paragraphs[0].runs[0]
        self.assertEqual(str(run.font.color.rgb), 'C00000')
        self.assertEqual(run._r.rPr.find(qn('w:shd')).get(qn('w:fill')), 'FFF2CC')
        self.assertIsNone(loaded.paragraphs[1].runs[0]._r.rPr.find(qn('w:shd')))


if __name__ == '__main__':
    unittest.main()
