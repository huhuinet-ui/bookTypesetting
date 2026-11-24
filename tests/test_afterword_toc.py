from types import SimpleNamespace
import unittest
from docx.oxml.ns import qn
import book_full as full
import book_layout as book
from word_styles import install_styles


class AfterwordTocTests(unittest.TestCase):
    def test_afterword_header_uses_custom_title_and_continuous_page_field(self):
        doc,_=book.new_doc()
        section=full.add_front_matter(doc,'写在最后','后记正文',show_header=True,pad=3)
        self.assertIn('写在最后',''.join(p.text for p in section.header.paragraphs))
        self.assertIn(book.BOOK_NAME,''.join(p.text for p in section.even_page_header.paragraphs))
        self.assertIn('PAGE',''.join(t.text or '' for t in section.header._element.iter(qn('w:instrText'))))
        self.assertIsNone(section._sectPr.find(qn('w:pgNumType')))
        self.assertFalse(section.different_first_page_header_footer)
        preface=full.add_front_matter(doc,'前言','前言正文')
        self.assertEqual(''.join(p.text for p in preface.header.paragraphs),'')

    def test_all_toc_styles_include_custom_afterword_and_measured_page(self):
        chapters=[{'title':'第一章 起点','articles':[{'title':'一篇文章','author':''}]}]
        for style in ('版式一','版式二','版式三','版式四'):
            with self.subTest(style=style):
                adapter=SimpleNamespace(book=book,toc=full.toc,build_toc=full.build_toc,
                    build_body=full.build_body,add_toc_line=full.add_toc_line,
                    TOC_TITLE='目录',CHAPTER_TOC_SIZE=10.5,CHAPTER_TOC_BOLD=True,ARTICLE_INDENT_PT=24)
                install_styles(adapter,{'tocStyle':style,'chapterStyle':'版式一'},[])
                doc,_=book.new_doc()
                pages={'第一章 起点':1,'一篇文章':2,'__book_afterword__':27,'__book_appendix_0__':31,'appendix-article':32}
                adapter.build_toc(doc,chapters,pages,pad=3,afterword=('写在最后','__book_afterword__'),appendices=[('附录一：年表','__book_appendix_0__')],
                    endmatter_articles={'__book_appendix_0__':[{'title':'附录内的文章','author':'','toc_key':'appendix-article'}]})
                text=''.join(doc.part.element.itertext())
                self.assertIn('写在最后',text)
                self.assertIn('027',text)
                self.assertIn('附录一：年表',text)
                self.assertIn('031',text)
                self.assertIn('附录内的文章',text)
                self.assertIn('032',text)
                if style=='版式二':
                    chapter=next(p for p in doc.paragraphs if p.text=='第一章 起点')
                    closing=next(p for p in doc.paragraphs if p.text.startswith('写在最后'))
                    for attr in ('space_before','space_after','line_spacing'):
                        self.assertEqual(getattr(chapter.paragraph_format,attr),getattr(closing.paragraph_format,attr))
                    self.assertEqual(chapter.runs[0].bold,closing.runs[0].bold)
                    self.assertEqual(closing.runs[0].font.size.pt,10)
                    self.assertTrue(closing.runs[0].bold)
                    self.assertEqual(closing.runs[0]._element.rPr.rFonts.get(qn('w:eastAsia')),book.F_BODY)
                if style=='版式四':
                    labels=[table.cell(0,0).text for table in doc.tables]
                    self.assertIn('写\n在\n最\n后',labels)
                    self.assertIn('附\n录\n一\n：\n年\n表',labels)
                    closing_table=next(table for table in doc.tables if table.cell(0,0).text=='写\n在\n最\n后')
                    run=closing_table.cell(0,0).paragraphs[0].runs[0]
                    self.assertEqual(run.font.size.pt,8)
                    self.assertTrue(run.bold)
                    self.assertEqual(run._element.rPr.rFonts.get(qn('w:eastAsia')),book.F_BODY)
                    self.assertEqual(doc.tables[-1].cell(0,0).paragraphs[0].runs[0].font.size,
                                     doc.tables[0].cell(0,0).paragraphs[0].runs[0].font.size)
                absent,_=book.new_doc()
                adapter.build_toc(absent,chapters,pages,pad=3)
                self.assertNotIn('写在最后',''.join(absent.part.element.itertext()))
                ordered,_=book.new_doc()
                adapter.build_toc(ordered,chapters,pages,pad=3,endmatter_order=[('附录一：年表','__book_appendix_0__'),('写在最后','__book_afterword__')])
                ordered_text=''.join(node.text or '' for node in ordered.part.element.iter(qn('w:t'))).replace('\n','')
                self.assertLess(ordered_text.index('附录一：年表'),ordered_text.index('写在最后'))

    def test_final_book_has_nonprinting_afterword_anchor_and_toc_entry(self):
        data={'chapters':[{'title':None,'articles':[{'title':'文章','author':'','items':[('p','正文。')]}]}],
              'preface':None,'afterword':'后记正文。','afterword_title':'写在最后'}
        doc,anchors=full.build_document(data,{'__book_afterword__':7},3)
        self.assertEqual(anchors[-1][1],'__book_afterword__')
        token=anchors[-1][0]
        paragraph=next(p for p in doc.paragraphs if p.text=='写在最后')
        self.assertEqual(paragraph._p.find(qn('w:bookmarkStart')).get(qn('w:name')),token)
        self.assertIn('写在最后\t007',[p.text for p in doc.paragraphs])
