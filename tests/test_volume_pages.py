import importlib
import importlib.util
import io
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn

import book_layout as book
import test_volumes as fixtures
from word_chapters import install_chapters


class VolumeTypographyTests(unittest.TestCase):
    def test_default_number_and_title_are_22pt_with_multiple_spacing(self):
        self.assertEqual(book.VOLUME_NUMBER_SIZE, 22)
        self.assertEqual(book.VOLUME_TITLE_SIZE, 22)
        self.check_saved_typography(22, 22)

    def test_custom_sizes_are_preserved_without_fixed_height_clipping(self):
        self.check_saved_typography(18, 36)

    def check_saved_typography(self, number_size, title_size):
        volume = {'number': '卷一', 'title': '那些值得被记住的时光与远方'}
        with patch.object(book, 'VOLUME_NUMBER_SIZE', number_size), \
                patch.object(book, 'VOLUME_TITLE_SIZE', title_size):
            doc = Document()
            book.add_volume_page(doc, volume)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        saved = Document(buffer)
        for paragraph, text, size in zip(saved.paragraphs, volume.values(), (number_size, title_size)):
            self.assertEqual(paragraph.text, text)
            self.assertEqual(paragraph.runs[0].font.size.pt, size)
            self.assertEqual(paragraph.paragraph_format.line_spacing_rule, WD_LINE_SPACING.ONE_POINT_FIVE)
            spacing = paragraph._p.get_or_add_pPr().find(qn('w:spacing'))
            self.assertEqual(spacing.get(qn('w:lineRule')), 'auto')
            self.assertEqual(paragraph.paragraph_format.line_spacing, 1.5)
            self.assertTrue(paragraph.paragraph_format.keep_together)
        self.assertTrue(saved.paragraphs[0].paragraph_format.keep_with_next)


@unittest.skipUnless(importlib.util.find_spec('fitz'), 'Native export module requires PyMuPDF')
class VolumePageTests(unittest.TestCase):
    setUp = fixtures.VolumeTests.setUp
    parse = fixtures.VolumeTests.parse

    def test_each_volume_gets_one_page_before_its_first_chapter(self):
        full = importlib.import_module('book_full')
        data = self.parse(fixtures.volume_manuscript())
        doc, section = book.new_doc()
        anchors = []
        full.build_body(doc, data['chapters'], section, anchors=anchors)
        texts = [p.text for p in doc.paragraphs]
        for number, name, chapter in [('卷一', '回望来路', '第1章 章节1'),
                                       ('第二卷', '向远方去', '第4章 章节4')]:
            self.assertEqual(sum(t.startswith(number) for t in texts), 1)
            self.assertEqual(texts.count(name), 1)
            self.assertLess(texts.index(name), next(i for i, t in enumerate(texts) if t.startswith(chapter)))
        self.assertEqual(len(doc.sections), 14)  # Two volume sections + six chapter/article pairs.
        self.assertEqual(len(anchors), 14)
        for index in (1, 8):
            self.assertEqual(doc.sections[index].start_type, WD_SECTION.NEW_PAGE)
        self.assertEqual(doc.sections[7].start_type, WD_SECTION.ODD_PAGE)
        starts = [s._sectPr.find(qn('w:pgNumType')) for s in doc.sections]
        self.assertEqual(starts[0].get(qn('w:start')), '1')
        self.assertTrue(all(s is None for s in starts[1:]))

    def test_unvolumed_book_does_not_gain_a_divider(self):
        full = importlib.import_module('book_full')
        data = self.parse('### 第一章\n\n#### 文章\n\n正文。')
        doc, section = book.new_doc()
        full.build_body(doc, data['chapters'], section)
        self.assertEqual(len(doc.sections), 2)

    def test_volume_page_stays_theme_colored_with_white_chapter_style(self):
        full = importlib.import_module('book_full')
        adapter = SimpleNamespace(book=book, build_body=full.build_body)
        data = self.parse(fixtures.volume_manuscript())
        with patch.object(book, 'add_page_background', book.add_page_background), \
                patch.object(book, 'add_volume_background', book.add_volume_background), \
                patch.object(book, 'add_chapter_page', book.add_chapter_page):
            install_chapters(adapter, {'chapterStyle': '版式二'}, [])
            doc, section = book.new_doc()
            adapter.build_body(doc, data['chapters'], section)
            self.assertIn('fillcolor="#' + str(book.BOOK_COLOR) + '"', doc.sections[0].header._element.xml)
            self.assertIn('fillcolor="#FFFFFF"', doc.sections[1].header._element.xml)


if __name__ == '__main__':
    unittest.main()
