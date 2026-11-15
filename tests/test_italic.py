import io
import unittest
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn

import book_layout as book


class ItalicTests(unittest.TestCase):
    def paragraph(self, text, mode='bold'):
        paragraph = Document().add_paragraph()
        with patch.object(book, 'MARK_MODE', mode):
            book.add_body_para(paragraph, text)
        return paragraph

    def test_short_italic_body_is_not_a_heading(self):
        self.assertEqual(book.classify(['*斜体短句*', ''], base_dir='.'), [('p', '*斜体短句*')])

    def test_mixed_italic_and_plain_text(self):
        paragraph = self.paragraph('前文*斜体*后文')
        self.assertEqual(paragraph.text, '前文斜体后文')
        self.assertEqual([(r.text, r.italic) for r in paragraph.runs],
                         [('前文', False), ('斜体', True), ('后文', False)])

    def test_nested_bold_and_italic(self):
        paragraph = self.paragraph('**加粗 *同时斜体* 加粗**')
        self.assertEqual(paragraph.text, '加粗 同时斜体 加粗')
        self.assertTrue(all(r.bold for r in paragraph.runs))
        self.assertEqual([r.italic for r in paragraph.runs], [False, True, False])
        paragraph = self.paragraph('***两种效果***')
        self.assertEqual(paragraph.text, '两种效果')
        self.assertTrue(paragraph.runs[0].bold)
        self.assertTrue(paragraph.runs[0].italic)

    def test_alt_font_and_color_preserve_italic(self):
        paragraph = self.paragraph('{{***叠加字体***}}', mode='color')
        run = paragraph.runs[0]
        self.assertEqual(paragraph.text, '叠加字体')
        self.assertTrue(run.italic)
        self.assertFalse(run.bold)
        self.assertEqual(run._r.rPr.rFonts.get(qn('w:eastAsia')), book.F_BODY_ALT)

    def test_unmatched_spaced_and_escaped_stars_remain_literal(self):
        for source, expected in ((r'\*原样星号\*', '*原样星号*'), ('* *', '* *'),
                                 ('2 * 3', '2 * 3'), ('未闭合*号', '未闭合*号'),
                                 ('* 有空格*', '* 有空格*')):
            with self.subTest(source=source):
                paragraph = self.paragraph(source)
                self.assertEqual(paragraph.text, expected)
                self.assertFalse(any(r.italic for r in paragraph.runs))

    def test_italic_survives_saved_word_document(self):
        doc = Document()
        book.add_body_para(doc.add_paragraph(), '*保存的斜体*普通文字')
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        saved = Document(buffer)
        self.assertTrue(saved.paragraphs[0].runs[0].italic)
        self.assertFalse(saved.paragraphs[0].runs[1].italic)

    def test_poetry_body_uses_italic_renderer(self):
        doc = Document()
        book._render_poems(doc, [('poem', '*诗句*')])
        self.assertEqual(doc.paragraphs[0].text, '诗句')
        self.assertTrue(doc.paragraphs[0].runs[0].italic)


if __name__ == '__main__':
    unittest.main()
