import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from docx import Document
from docx.oxml.ns import qn
import book_layout as book
import book_full as full


class ArticleYearTests(unittest.TestCase):
    def parse(self,text):
        with tempfile.TemporaryDirectory() as directory:
            source=Path(directory)/'book.md'
            source.write_text(text,encoding='utf-8')
            return book.parse_book(str(source))

    def test_year_and_author_either_order_and_omitted_metadata(self):
        for metadata in ['年份：2020年\n作者：张三','作者：张三\n\n年份: 2020年']:
            art=self.parse('#### 文章\n'+metadata+'\n\n正文。')['chapters'][0]['articles'][0]
            self.assertEqual((art['author'],art['year']),('张三','2020年'))
            self.assertEqual(art['items'],[('p','正文。')])
        art=self.parse('#### 文章\n年份：2019年\n正文。')['chapters'][0]['articles'][0]
        self.assertEqual(art['author'],'');self.assertEqual(art['year'],'2019年')
        self.assertEqual(self.parse('#### 文章\n正文。')['chapters'][0]['articles'][0]['year'],'')

    def test_body_year_line_is_not_removed_after_body_begins(self):
        art=self.parse('#### 文章\n正文。\n\n年份：这是正文内容')['chapters'][0]['articles'][0]
        self.assertEqual(art['year'],'')
        self.assertIn(('p','年份：这是正文内容'),art['items'])

    def test_appendix_metadata_and_word_title_block_without_toc_year(self):
        article=book.parse_appendix('附录',['#### 回望','年份：2021年','作者：张三','','正文。'])['articles'][0]
        self.assertEqual(article['year'],'2021年')
        doc,_=book.new_doc()
        book.build_article(doc,article,True)
        texts=[p.text for p in doc.paragraphs]
        self.assertEqual(texts[texts.index('回望')+2],'2021年')
        year=next(p for p in doc.paragraphs if p.text=='2021年')
        self.assertEqual(year.runs[0].font.size.pt,book.YEAR_SIZE)
        self.assertEqual(year.runs[0]._element.rPr.rFonts.get(qn('w:eastAsia')),book.F_YEAR)
        self.assertEqual(year.runs[0].font.color.rgb,book.BOOK_COLOR)
        self.assertEqual(year.runs[0].font.name,book.F_YEAR)
        rules=[p for p in doc.paragraphs if p._p.find('.//'+qn('w:pBdr')) is not None]
        self.assertEqual(len(rules),1)  # Author does not add a second line.
        self.assertEqual(rules[0]._p.find('.//'+qn('w:bottom')).get(qn('w:sz')),'2')
        self.assertAlmostEqual(year.paragraph_format.space_before.pt,book.YEAR_RULE_AFTER,places=1)
        toc=Document()
        full.build_toc(toc,[{'title':None,'articles':[article]}],{'回望':1})
        self.assertNotIn('2021年',''.join(t.text or '' for t in toc.element.iter(qn('w:t'))))

    def test_year_alone_adds_rule_and_theme_color_on_both_page_sides(self):
        for recto in [True,False]:
            with self.subTest(recto=recto):
                doc,_=book.new_doc()
                book.build_article(doc,{'title':'文章','year':'1970年/12岁','author':'','items':[]},True,recto=recto)
                rule=next(p for p in doc.paragraphs if p._p.find('.//'+qn('w:pBdr')) is not None)
                self.assertTrue(rule.paragraph_format.keep_with_next)
                indent=rule.paragraph_format.left_indent if recto else rule.paragraph_format.right_indent
                self.assertAlmostEqual(indent.cm,(book.PAGE_W-book.M_LEFT-book.M_RIGHT)/10-book.YEAR_RULE_W,places=2)
                self.assertEqual(sum(bool(p._p.find('.//'+qn('w:pBdr')) is not None) for p in doc.paragraphs),1)

    def test_browser_metadata_matches_native(self):
        source=(Path(__file__).resolve().parents[1]/'排版工具.html').read_text(encoding='utf-8')
        function=source[source.index('function readArticleMetadata('):source.index('function parseAppendix(')]
        cases=[['年份：2020年','作者：张三','','正文'],['作者:张三','年份: 2020年','正文'],['正文','年份：不处理'],['**年份：2021年**','正文']]
        script=function+';console.log(JSON.stringify('+json.dumps(cases,ensure_ascii=False)+'.map(lines=>readArticleMetadata(lines))));'
        result=subprocess.run(['node','-e',script],check=True,capture_output=True)
        actual=json.loads(result.stdout.decode('utf-8'))
        for lines,parsed in zip(cases,actual):
            author,year,next_line=book.read_article_metadata(lines)
            self.assertEqual(parsed,{'author':author,'year':year,'next':next_line})
