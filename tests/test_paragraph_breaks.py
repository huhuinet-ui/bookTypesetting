import io
import unittest
from unittest.mock import patch
from docx import Document
from docx.oxml.ns import qn
import book_layout as book


class ParagraphBreakTests(unittest.TestCase):
    def test_legacy_each_line_mode_retains_explicit_manual_breaks(self):
        with patch.object(book, 'BODY_EACH_LINE', True):
            self.assertEqual(book.classify(['第一段', '第二段', '', '上行\\', '下行']),
                             [('p', '第一段'), ('p', '第二段'), ('p', '上行\n下行')])

    def test_word_manual_breaks_do_not_expand_spacing(self):
        doc = Document()
        para = doc.add_paragraph()
        book.set_para(para, book.JUSTIFY)
        book.add_body_para(para, '上行\n下行')
        self.assertEqual(para.alignment, book.JUSTIFY)
        flags = list(doc.settings.element.iter(qn('w:doNotExpandShiftReturn')))
        self.assertEqual(len(flags), 1)
        self.assertEqual(flags[0].get(qn('w:val')), '1')
        self.assertEqual(para.text, '上行\n下行')

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
