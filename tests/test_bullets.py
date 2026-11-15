import io
import unittest
from docx import Document
import book_layout as book


class BulletTests(unittest.TestCase):
    def test_adjacent_bullets_stay_separate(self):
        self.assertEqual(book.classify(['正文', '- 第一项', '- **第二项**', '后文']),
                         [('p', '正文'), ('li', '第一项'), ('li', '**第二项**'), ('p', '后文')])

    def test_hyphen_without_space_is_plain_text(self):
        self.assertEqual(book.classify(['-不是列表']), [('p', '-不是列表')])

    def test_word_native_bullet_and_hanging_indent(self):
        doc = Document()
        book.add_bullet_para(doc.add_paragraph(), '列表内容')
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        p = Document(stream).paragraphs[0]
        self.assertEqual(p.style.name, 'List Bullet')
        self.assertIsNotNone(p.style.element.pPr.numPr)
        self.assertEqual(p.text, '列表内容')
        self.assertEqual(p.paragraph_format.left_indent.pt, book.BODY_SIZE * 1.5)
        self.assertEqual(p.paragraph_format.first_line_indent.pt, -book.BODY_SIZE * 1.5)
