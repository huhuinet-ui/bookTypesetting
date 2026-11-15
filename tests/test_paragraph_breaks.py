import io
import unittest
from docx import Document
from docx.oxml.ns import qn
import book_layout as book


class ParagraphBreakTests(unittest.TestCase):
    def test_line_break_vs_blank_line(self):
        self.assertEqual(book.classify(['第一行', '第二行', '', '下一段']),
                         [('p', '第一行\n第二行'), ('p', '下一段')])

    def test_structural_boundaries_do_not_merge(self):
        self.assertEqual(book.classify(['正文', '##### 小标题', '后文', '1', '末段']),
                         [('p', '正文'), ('h', '小标题'), ('p', '后文'), ('num', '1'), ('p', '末段')])
        self.assertEqual(book.classify(['{{', '甲', '乙', '}}', '丙']),
                         [('palt', '甲\n乙'), ('p', '丙')])

    def test_saved_word_has_one_paragraph_with_line_break(self):
        doc = Document()
        for _, text in book.classify(['甲', '乙', '', '丙']):
            book.add_body_para(doc.add_paragraph(), text)
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        loaded = Document(stream)
        self.assertEqual([p.text for p in loaded.paragraphs], ['甲\n乙', '丙'])
        self.assertEqual(len(list(loaded.paragraphs[0]._p.iter(qn('w:br')))), 1)

    def test_poetry_keeps_individual_lines(self):
        self.assertEqual(book.classify(['甲', '乙', '', '丙'], poem=True),
                         [('poem', '甲'), ('poem', '乙'), ('stanza', None), ('poem', '丙')])
