import io
import unittest
from types import SimpleNamespace
from docx import Document
from docx.enum.section import WD_SECTION
import book_layout as book
from word_styles import install_styles
from word_toc_vertical import split_chapter_title,article_chunks

class VerticalTocTests(unittest.TestCase):
    def adapter(self):
        full=SimpleNamespace(book=book,build_body=lambda *a,**k:'body')
        install_styles(full,{'tocStyle':'版式四：竖排章节+标题页码','chapterStyle':'版式一'},[])
        return full

    def test_title_and_sequence_split(self):
        self.assertEqual(split_chapter_title('第一章 家的刻印与远行'),('第一章','家的刻印与远行'))
        self.assertEqual(split_chapter_title('没有序号'),('','没有序号'))

    def test_editable_titles_and_measured_page_numbers_with_volume(self):
        doc=Document();full=self.adapter()
        volume={'label':'卷一 回望','number':'卷一','title':'回望'}
        chapters=[{'title':'第一章 家的刻印与远行','volume':volume,'articles':[{'title':'妈妈的节约哲学'},{'title':'初高中住校六年'}]}]
        full.build_toc(doc,chapters,{'卷一 回望':1,'妈妈的节约哲学':7,'初高中住校六年':13})
        stream=io.BytesIO();doc.save(stream);stream.seek(0);loaded=Document(stream)
        self.assertIn('卷一 回望 / 001',[p.text for p in loaded.paragraphs])
        row=loaded.tables[0].rows[0]
        self.assertEqual(row.cells[0].text.replace('\n',''),'家的刻印与远行')
        self.assertEqual(row.cells[1].text.replace('\n',''),'─第一章─')
        self.assertEqual(row.cells[3].text,'妈妈的节约哲学 / 007\n初高中住校六年 / 013')
        self.assertAlmostEqual(loaded.sections[0].page_width.mm,166,places=1)
        body=doc.add_section(WD_SECTION.ODD_PAGE);self.assertEqual(full.build_body(doc,[],body),'body')
        self.assertAlmostEqual(body.left_margin.mm,book.M_LEFT+book.page_inset(),places=1)

    def test_pairs_page_break_shift_and_unsplit_labels(self):
        from docx.oxml.ns import qn
        doc=Document()
        chapters=[{'title':'第%d章 章节名称完整不跨页'%i,'articles':[{'title':'文章%d'%i}]} for i in range(1,5)]
        self.adapter().build_toc(doc,chapters,{})
        breaks=[p for p in doc.paragraphs if p.paragraph_format.page_break_before]
        self.assertEqual(len(breaks),1)
        for index,table in enumerate(doc.tables):
            self.assertIsNotNone(table.rows[0]._tr.trPr.find(qn('w:cantSplit')))
            indent=table._tbl.tblPr.find(qn('w:tblInd'))
            if index%2:
                self.assertAlmostEqual(int(indent.get(qn('w:w')))/20*25.4/72,6,places=1)
            else:
                self.assertIsNone(indent)
            label=table.cell(0,0).paragraphs[0]
            self.assertTrue(label.paragraph_format.keep_together)
            self.assertEqual(label.text.replace('\n',''),'章节名称完整不跨页')

    def test_long_chapters_chunk_without_missing_entries(self):
        articles=[{'title':'这是一篇长文章标题'+str(i)*10} for i in range(60)]
        chunks=article_chunks(articles,140,200,9.5)
        self.assertGreater(len(chunks),1)
        self.assertEqual([a for chunk in chunks for a in chunk],articles)
        doc=Document();self.adapter().build_toc(doc,[{'title':None,'articles':articles}],{})
        self.assertEqual(len(doc.tables),len(chunks))
