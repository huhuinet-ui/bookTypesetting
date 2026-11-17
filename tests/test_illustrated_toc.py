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
        texts = [row.cells[0].text + '\t' + row.cells[1].text for row in doc.tables[0].rows]
        self.assertIn('卷一\n回望来路\t001', texts)
        self.assertIn('第二卷\n向远方去\t015', texts)
        self.assertIn('第4章 章节4\t012', texts)
        self.assertIn('文章4\t013', texts)
        self.assertEqual(len(texts), 14)  # Two volumes, six chapters, six articles.
        self.assertEqual(len(doc.tables), 1)
        self.assertTrue(all(cell.paragraphs[0].paragraph_format.keep_together for row in doc.tables[0].rows for cell in row.cells))
        self.assertTrue(doc.tables[0].rows[0].cells[0].paragraphs[0].paragraph_format.keep_with_next)
        self.assertFalse(doc.tables[0].rows[2].cells[0].paragraphs[0].paragraph_format.keep_with_next)

    def test_default_font_sizes_and_line_spacing(self):
        doc = Document()
        data = self.parse(fixtures.volume_manuscript())
        self.adapter().build_toc(doc, data['chapters'], {})
        rows = doc.tables[0].rows
        volume_runs = [r for r in rows[0].cells[0].paragraphs[0].runs if r.text.strip()]
        self.assertEqual([r.font.size.pt for r in volume_runs], [13, 13])
        self.assertEqual(rows[1].cells[0].paragraphs[0].runs[0].font.size.pt, 10.5)
        self.assertEqual(rows[2].cells[0].paragraphs[0].runs[0].font.size.pt, 9.5)
        for row in rows:
            for cell in row.cells:
                self.assertAlmostEqual(cell.paragraphs[0].paragraph_format.line_spacing, 1.2)
            self.assertEqual(row.cells[1].paragraphs[0].runs[0].font.size.pt, 10)

    def test_without_volumes_and_without_chapters(self):
        full = self.adapter()
        data = self.parse('#### 一篇文章\n\n作者：张三\n\n正文。')
        doc = Document()
        full.build_toc(doc, data['chapters'], {'一篇文章': 1})
        self.assertEqual([row.cells[0].text + '\t' + row.cells[1].text for row in doc.tables[0].rows], ['一篇文章\t001'])
        self.assertEqual(doc.tables[0].rows[0].cells[0].paragraphs[0].paragraph_format.left_indent.pt, 0)

    def test_art_alternates_while_background_and_title_repeat(self):
        full = self.adapter()
        doc = Document()
        full.build_toc(doc, [], {})
        section = doc.sections[0]
        for index, header in enumerate((section.header, section.even_page_header)):
            ellipses = list(header._element.iter('{urn:schemas-microsoft-com:vml}oval'))
            self.assertEqual(len(ellipses), 17 if index == 0 else 0)
            self.assertIn('layout-flow:vertical-ideographic', header._element.xml)
            self.assertIn('目录', header._element.xml)
            self.assertIn('fillcolor="#faf8f2"', header._element.xml)
        from docx.oxml.ns import qn
        self.assertEqual(section._sectPr.find(qn('w:pgNumType')).get(qn('w:start')), '1')
        self.assertTrue(doc.settings.odd_and_even_pages_header_footer)
        width = book.PAGE_W + book.page_inset() * 2
        self.assertAlmostEqual(section.left_margin.mm, book.page_inset() + max(10, book.M_RIGHT), places=1)
        self.assertAlmostEqual(section.right_margin.mm, section.left_margin.mm, places=1)
        self.assertIn('tblInd', doc.tables[0]._tbl.xml)

    def test_trim_size_and_safe_area_match_body_for_custom_paper(self):
        from unittest.mock import patch
        from docx.oxml.ns import qn
        for width, height, bleed in ((140, 200, 3), (166, 226, 3), (140, 210, 5), (180, 240, 0)):
            with patch.object(book, 'PAGE_W', width), patch.object(book, 'PAGE_H', height), patch.object(book, 'BLEED_MM', bleed):
                doc = Document()
                self.adapter().build_toc(doc, [], {})
                section = doc.sections[0]
                self.assertAlmostEqual(section.page_width.mm, width+2*(bleed+book.PRINT_MARGIN_MM), places=1)
                self.assertAlmostEqual(section.page_height.mm, height+2*(bleed+book.PRINT_MARGIN_MM), places=1)
                table = doc.tables[0]
                indent_mm = int(table._tbl.tblPr.find(qn('w:tblInd')).get(qn('w:w'))) / 20 * 25.4 / 72
                table_width_mm = sum(col.width.mm for col in table.columns)
                right_edge = section.left_margin.mm + indent_mm + table_width_mm
                self.assertAlmostEqual(right_edge, book.page_inset()+width-max(10,book.M_RIGHT), places=1)
                self.assertGreaterEqual(section.bottom_margin.mm, book.page_inset()+book.M_BOT-.1)

    def test_long_titles_keep_page_numbers_in_separate_fixed_columns(self):
        import io
        from docx.oxml.ns import qn
        title = '很长的文章标题用于验证换行时不会挤乱右侧页码' * 4
        chapters = [{'title': None, 'articles': [{'title': title} for _ in range(45)]}]
        doc = Document()
        self.adapter().build_toc(doc, chapters, {title: 1234})
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        loaded = Document(stream)
        table = loaded.tables[0]
        self.assertFalse(table.autofit)
        self.assertEqual(len(table.rows), 45)
        for row in table.rows:
            self.assertEqual(row.cells[0].text, title)
            self.assertEqual(row.cells[1].text, '1234')
            self.assertIsNotNone(row._tr.trPr.find(qn('w:cantSplit')))
            self.assertNotIn('\t', row.cells[0].text)

    def test_body_does_not_inherit_toc_art_or_column_margins(self):
        full = self.adapter()
        doc = Document()
        full.build_toc(doc, [], {})
        body = doc.add_section(WD_SECTION.ODD_PAGE)
        self.assertEqual(full.build_body(doc, [], body), 'body')
        self.assertAlmostEqual(body.left_margin.mm, book.M_LEFT + book.page_inset(), places=1)
        for header in (body.header, body.even_page_header):
            self.assertNotIn('toc3_', header._element.xml)
        self.assertIn('toc3_', doc.sections[0].header._element.xml)


if __name__ == '__main__':
    unittest.main()
