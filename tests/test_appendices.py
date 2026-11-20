from pathlib import Path
import tempfile
import unittest
from docx.oxml.ns import qn
import book_layout as book
import book_full as full


class AppendixTests(unittest.TestCase):
    def test_legacy_article_level_afterword_ends_appendix(self):
        root=Path(__file__).resolve().parents[1]/'tmp'/'appendix-tests'
        root.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as directory:
            source=Path(directory)/'book.md'
            source.write_text('#### 文章\n正文。\n# 附录\n#### 资料\n附录正文。\n#### 后记\n后记正文。',encoding='utf-8')
            data=book.parse_book(str(source))
        self.assertEqual(data['afterword'],'后记正文。')
        self.assertEqual(data['afterword_title'],'后记')
        self.assertEqual(len(data['appendices'][0]['articles']),1)
        self.assertEqual(len(data['chapters'][0]['articles']),1)

    def test_appendices_are_separate_from_chapters_with_content_and_images(self):
        root=Path(__file__).resolve().parents[1]/'tmp'/'appendix-tests'
        root.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as directory:
            source=Path(directory)/'book.md'
            source.write_text('# 书名\n### 第一章\n#### 文章\n正文。\n# 后记\n结语。\n# 附录一：年表\n**加粗正文**\n\n#### 小标题\n![照片](图/a.png)\n### 附录二：资料\n第二份资料。',encoding='utf-8')
            data=book.parse_book(str(source))
        self.assertEqual(len(data['chapters']),1)
        self.assertEqual([item['title'] for item in data['appendices']],['附录一：年表','附录二：资料'])
        self.assertIn('##### 小标题',data['appendices'][0]['text'])
        self.assertEqual(book.classify(data['appendices'][0]['text'].splitlines())[-1][0],'img')
        self.assertEqual(data['afterword'],'结语。')

    def test_nested_appendix_headings_do_not_escape_into_previous_chapter(self):
        root=Path(__file__).resolve().parents[1]/'tmp'/'appendix-tests'
        root.mkdir(parents=True,exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as directory:
            source=Path(directory)/'book.md'
            source.write_text('### 章节\n#### 文章\n正文。\n# 附录\n#### 附录内文章\n**内容一**\n\n内容二。\n# 后记\n结语。',encoding='utf-8')
            data=book.parse_book(str(source))
        self.assertEqual(len(data['chapters'][0]['articles']),1)
        self.assertIn('##### 附录内文章',data['appendices'][0]['text'])
        self.assertIn('内容二。',data['appendices'][0]['text'])
        self.assertEqual(data['afterword'],'结语。')
        article=data['appendices'][0]['articles'][0]
        self.assertEqual(article['title'],'附录内文章')
        self.assertEqual(article['items'],[('p','**内容一**'),('p','内容二。')])
        doc,anchors=full.build_document(data,{},3)
        title=next(p for p in doc.paragraphs if p.text=='附录内文章')
        self.assertEqual(title.runs[0].font.size.pt,book.TITLE_SIZE)
        self.assertIn('v:rect',title._p.xml)
        self.assertIn('__book_appendix_0_article_0__',[key for _,key in anchors])

    def test_appendix_pages_have_bookmarks_and_no_chapter_decoration(self):
        data={'chapters':[{'title':None,'articles':[{'title':'文章','author':'','items':[('p','正文。')]}]}],
              'preface':None,'afterword':None,
              'appendices':[{'title':'附录一：年表','text':'资料正文。'},{'title':'附录二：索引','text':'索引正文。'}]}
        doc,anchors=full.build_document(data,{'__book_appendix_0__':3,'__book_appendix_1__':4},3)
        self.assertEqual([key for _,key in anchors[-2:]],['__book_appendix_0__','__book_appendix_1__'])
        for index,item in enumerate(data['appendices']):
            title=next(p for p in doc.paragraphs if p.text==item['title'])
            self.assertIsNotNone(title._p.find(qn('w:bookmarkStart')))
            self.assertNotIn('v:rect',title._p.xml)
            self.assertIn(item['title']+'\t00'+str(index+3),[p.text for p in doc.paragraphs])

    def test_original_appendix_then_afterword_order_is_preserved(self):
        root=Path(__file__).resolve().parents[1]/'tmp'/'appendix-tests'
        root.mkdir(parents=True,exist_ok=True)
        for headings in [('附录','后记'),('后记','附录')]:
            with tempfile.TemporaryDirectory(dir=str(root)) as directory:
                source=Path(directory)/'book.md'
                source.write_text('#### 文章\n正文。\n# '+headings[0]+'\n内容一。\n# '+headings[1]+'\n内容二。',encoding='utf-8')
                data=book.parse_book(str(source))
            entries=book.endmatter_entries(data)
            self.assertEqual([item['title'] for item in entries],list(headings))
            doc,anchors=full.build_document(data,{},3)
            self.assertEqual([key for _,key in anchors[-2:]],[item['key'] for item in entries])
            titles=[p.text for p in doc.paragraphs if p.text in headings]
            self.assertEqual(titles,list(headings))
