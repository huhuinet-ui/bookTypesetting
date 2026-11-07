from types import SimpleNamespace
import unittest

from docx import Document
from docx.enum.section import WD_SECTION

import book_layout as book
import test_volumes as fixtures
from word_styles import install_styles


class IllustratedTocTests(unittest.TestCase):
    setUp = fixtures.VolumeTests.setUp
    parse = fixtures.VolumeTests.parse

    def adapter(self):
        full = SimpleNamespace(book=book, TOC_TITLE='目录', CHAPTER_TOC_BOLD=True,
                               build_body=lambda *args, **kwargs: 'body')
        install_styles(full, {'tocStyle': '版式三：年轮枝叶+右侧层级目录', 'chapterStyle': '版式一'}, [])
        return full

    def test_volumes_chapters_and_articles_have_measured_page_numbers(self):
        data = self.parse(fixtures.volume_manuscript())
        pages = {'卷一 回望来路': 1, '第二卷 向远方去': 15}
        for number in range(1, 7):
            pages['第{}章 章节{}'.format(number, number)] = number * 3
            pages['文章{}'.format(number)] = number * 3 + 1
        doc = Document()
        self.adapter().build_toc(doc, data['chapters'], pages)
        texts = [p.text for p in doc.paragraphs]
        self.assertIn('卷一\n回望来路\t001', texts)
        self.assertIn('第二卷\n向远方去\t015', texts)
        self.assertIn('第4章 章节4\t012', texts)
        self.assertIn('文章4\t013', texts)
        self.assertEqual(len(texts), 14)  # Two volumes, six chapters, six articles.
        self.assertEqual(len(doc.tables), 0)
        self.assertTrue(all(p.paragraph_format.keep_together for p in doc.paragraphs))
        self.assertTrue(doc.paragraphs[0].paragraph_format.keep_with_next)
        self.assertFalse(doc.paragraphs[2].paragraph_format.keep_with_next)

    def test_without_volumes_and_without_chapters(self):
        full = self.adapter()
        data = self.parse('#### 一篇文章\n\n作者：张三\n\n正文。')
        doc = Document()
        full.build_toc(doc, data['chapters'], {'一篇文章': 1})
        self.assertEqual([p.text for p in doc.paragraphs], ['一篇文章\t001'])
        self.assertEqual(doc.paragraphs[0].paragraph_format.left_indent.pt, 0)

    def test_page_art_and_vertical_title_repeat_in_both_headers(self):
        full = self.adapter()
        doc = Document()
        full.build_toc(doc, [], {})
        section = doc.sections[0]
        for header in (section.header, section.even_page_header):
            ellipses = list(header._element.iter('{urn:schemas-microsoft-com:vml}oval'))
            self.assertEqual(len(ellipses), 17)
            self.assertIn('layout-flow:vertical-ideographic', header._element.xml)
            self.assertIn('目录', header._element.xml)
            self.assertIn('fillcolor="#faf8f2"', header._element.xml)
        width = book.PAGE_W + book.BLEED_MM * 2
        self.assertAlmostEqual(section.left_margin.mm, width * .49, places=1)

    def test_body_does_not_inherit_toc_art_or_column_margins(self):
        full = self.adapter()
        doc = Document()
        full.build_toc(doc, [], {})
        body = doc.add_section(WD_SECTION.ODD_PAGE)
        self.assertEqual(full.build_body(doc, [], body), 'body')
        self.assertAlmostEqual(body.left_margin.mm, book.M_LEFT + book.BLEED_MM, places=1)
        for header in (body.header, body.even_page_header):
            self.assertNotIn('toc3_', header._element.xml)
        self.assertIn('toc3_', doc.sections[0].header._element.xml)


if __name__ == '__main__':
    unittest.main()
