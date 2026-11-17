import io
import unittest
from types import SimpleNamespace
from docx import Document
from docx.shared import Pt
import book_layout as book
from word_styles import install_styles


class ReferenceTocFontTests(unittest.TestCase):
    def test_custom_sizes_survive_word_save_and_line_height_grows(self):
        def run(p,text,size,bold=False):
            r=p.add_run(text);r.font.size=Pt(size);r.bold=bold;return r
        for chapter_size,entry_size in ((12,10.5),(24,20)):
            full=SimpleNamespace(book=book,TOC_TITLE='目录',CHAPTER_TOC_BOLD=True,
                build_body=lambda *a,**k:None,toc=SimpleNamespace(R=run))
            params={'tocStyle':'版式二','chapterStyle':'版式一',
                    'chapterTocSize':chapter_size,'tocEntrySize':entry_size,'tocTitleSize':chapter_size+6}
            install_styles(full,params,[])
            doc=Document()
            full.build_toc(doc,[{'title':'章节','articles':[{'title':'文章','author':'作者'}]}],{'文章':12})
            data=io.BytesIO();doc.save(data);data.seek(0);loaded=Document(data)
            heading=next(p for p in loaded.paragraphs if p.text=='目录')
            self.assertEqual(heading.runs[0].font.size.pt,chapter_size+6)
            self.assertGreaterEqual(heading.paragraph_format.line_spacing.pt,(chapter_size+6)*1.2-.1)
            chapter=next(p for p in loaded.paragraphs if p.text=='章节')
            self.assertEqual(chapter.runs[0].font.size.pt,chapter_size)
            self.assertGreaterEqual(chapter.paragraph_format.line_spacing.pt,chapter_size*1.3-.1)
            titlep=loaded.tables[0].cell(0,0).paragraphs[0]
            self.assertEqual(titlep.runs[0].font.size.pt,entry_size)
            self.assertGreaterEqual(titlep.paragraph_format.line_spacing.pt,entry_size*1.3-.1)
