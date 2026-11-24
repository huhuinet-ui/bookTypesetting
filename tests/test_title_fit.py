import unittest
from types import SimpleNamespace
from unittest.mock import patch
from docx.oxml.ns import qn
import book_layout as book
from word_chapters import install_chapters


class TitleFitTests(unittest.TestCase):
    def test_style_two_uses_normalized_label_measured_size_and_safe_height(self):
        label='第一章 我的童年'
        for separator in [' | ', ':', '：', '｜：']:
            with self.subTest(separator=separator), patch.object(book,'TITLE_FIT_PLAN',{('chapter2',label):12}), patch.object(book,'add_chapter_page'), patch.object(book,'add_page_background'), patch.object(book,'add_volume_background',create=True):
                full=SimpleNamespace(book=book,build_body=lambda *args:None)
                install_chapters(full,{'chapterStyle':'版式二','chapter2FrameOn':False,'chapter2Top':40},[])
                doc,_=book.new_doc()
                book.add_chapter_page(doc,'第一章'+separator+'我的童年')
                texts=list(doc.element.iter(qn('w:t')))
                self.assertEqual(texts[-1].text,label.replace(' ','\u00a0'))
                self.assertEqual(texts[-1].getparent().find(qn('w:rPr')).find(qn('w:sz')).get(qn('w:val')),'24')
                shapes=list(doc.element.iter('{urn:schemas-microsoft-com:vml}rect'))
                self.assertEqual(len(shapes),1)  # No border when switched off.
                style=shapes[0].get('style')
                top=(book.page_inset()+book.PAGE_H*.4)*72/25.4
                height=(book.PAGE_H*.6-8)*72/25.4-8
                self.assertIn(f'margin-top:{top}pt',style)
                self.assertIn(f'height:{height}pt',style)

    def test_native_article_uses_measured_size(self):
        with patch.object(book, 'TITLE_FIT_PLAN', {('article', '超长文章标题'): 10}), patch.object(book, 'TITLE_SIZE', 14):
            doc, _ = book.new_doc()
            book.build_article(doc, {'title': '超长文章标题', 'author': '', 'items': [], 'poem': False}, True)
            title = next(p for p in doc.paragraphs if p.text == '超长文章标题')
            self.assertEqual(title.runs[0].font.size.pt, 10)

    def test_all_three_editable_chapter_styles_share_the_plan(self):
        for style in ['版式三：玫瑰', '版式四：纸飞机', '版式五：指纹树']:
            with self.subTest(style=style), patch.object(book, 'TITLE_FIT_PLAN', {('chapter', '章节名称'): 12}), patch.object(book, 'add_chapter_page'), patch.object(book, 'add_page_background'), patch.object(book, 'add_volume_background', create=True):
                full = SimpleNamespace(book=book, build_body=lambda *args: None)
                install_chapters(full, {'chapterStyle': style}, [])
                doc, _ = book.new_doc()
                book.add_chapter_page(doc, '第一章 章节名称')
                text = next(t for t in doc.element.iter(qn('w:t')) if t.text == '章节名称')
                self.assertEqual(text.getparent().find(qn('w:rPr')).find(qn('w:sz')).get(qn('w:val')), '24')

    def test_missing_or_invalid_measurement_keeps_configured_size(self):
        for value in [100, -1, True, '12', float('nan')]:
            with patch.object(book, 'TITLE_FIT_PLAN', {('article', '标题'): value}):
                self.assertEqual(book.fitted_title_size('article', '标题', 14), 14)
