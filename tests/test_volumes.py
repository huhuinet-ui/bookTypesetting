import contextlib
import io
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest

from docx import Document
from docx.oxml.ns import qn

import book_layout
from word2md import convert
from word_styles import install_styles


def volume_manuscript():
    lines = []
    for volume, name, start in (("卷一", "回望来路", 1), ("第二卷", "向远方去", 4)):
        lines.append("## {} {}".format(volume, name))
        for number in range(start, start + 3):
            lines.extend(["### 第{}章 章节{}".format(number, number),
                          "简介：本章引言。", "#### 文章{}".format(number),
                          "作者：作者{}".format(number), "正文内容。"])
    return "\n\n".join(lines)


class VolumeTests(unittest.TestCase):
    def setUp(self):
        work = Path(__file__).resolve().parents[1] / "tmp" / "volume-tests"
        work.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=str(work))
        self.addCleanup(self.temp.cleanup)

    def parse(self, text):
        source = Path(self.temp.name) / "source.md"
        source.write_text(text, encoding="utf-8")
        return book_layout.parse_book(str(source))

    def test_two_volumes_six_chapters_keep_numbering_and_authors(self):
        data = self.parse(volume_manuscript())
        self.assertEqual([v['number'] for v in data['volumes']], ['卷一', '第二卷'])
        self.assertEqual([v['title'] for v in data['volumes']], ['回望来路', '向远方去'])
        self.assertEqual(len(data['chapters']), 6)
        self.assertEqual([c['volume']['id'] for c in data['chapters']], [0, 0, 0, 1, 1, 1])
        for number, chapter in enumerate(data['chapters'], 1):
            self.assertEqual(chapter['title'], '第{}章 章节{}'.format(number, number))
            self.assertEqual(chapter['intro'], '本章引言。')
            self.assertEqual(chapter['articles'][0]['author'], '作者{}'.format(number))

    def test_existing_manuscript_without_volumes(self):
        data = self.parse('### 第一章\n\n#### 文章\n\n正文。')
        self.assertEqual(data['volumes'], [])
        self.assertIsNone(data['chapters'][0]['volume'])
        self.assertEqual(len(data['chapters'][0]['articles']), 1)

    def test_new_volume_does_not_append_articles_to_previous_chapter(self):
        data = self.parse('## 卷一 旧卷\n### 第一章\n#### 旧文\n旧文正文。\n## 卷二 新卷\n#### 新文\n新文正文。')
        self.assertEqual(len(data['chapters']), 2)
        self.assertEqual(data['chapters'][0]['articles'][0]['title'], '旧文')
        self.assertEqual(data['chapters'][1]['articles'][0]['title'], '新文')
        self.assertEqual(data['chapters'][1]['volume']['number'], '卷二')

    def test_word_heading_two_becomes_volume_heading(self):
        doc = Document()
        doc.add_heading('卷一 回望来路', level=2)
        doc.add_heading('第一章 童年', level=3)
        doc.add_heading('往事', level=4)
        doc.add_paragraph('正文。')
        source = Path(self.temp.name) / 'source.docx'
        doc.save(str(source))
        output = source.with_suffix('.md')
        with contextlib.redirect_stdout(io.StringIO()):
            convert(str(source), str(output))
        self.assertIn('## 卷一 回望来路', output.read_text(encoding='utf-8'))
        self.assertEqual(book_layout.parse_book(str(output))['volumes'][0]['title'], '回望来路')

    def test_native_reference_toc_preserves_volume_groups_and_page_numbers(self):
        data = self.parse(volume_manuscript())
        full = SimpleNamespace(book=book_layout, TOC_TITLE='目录', CHAPTER_TOC_BOLD=True,
                               build_body=lambda *args, **kwargs: None,
                               toc=SimpleNamespace(R=lambda p, text, size, bold=False: p.add_run(text)))
        install_styles(full, {'tocStyle': '版式二', 'chapterStyle': '版式一',
                              'bookColor': '#336699', 'tocShowSequence': True}, [])
        doc = Document()
        pages = {'文章{}'.format(n): n * 5 for n in range(1, 7)}
        full.build_toc(doc, data['chapters'], pages)
        headings = [p.text for p in doc.paragraphs]
        self.assertEqual(sum('回望来路' in s for s in headings), 1)
        self.assertEqual(sum('向远方去' in s for s in headings), 1)
        self.assertLess(headings.index('第二卷  向远方去'), headings.index('第4章 章节4'))
        self.assertEqual(len(doc.tables), 6)
        self.assertEqual(doc.tables[3].cell(0, 0).text, '01')
        self.assertEqual(doc.tables[3].cell(0, 2).text, '作者4 / 020')
        self.assertTrue(doc.tables[0]._tbl.tblPr.find(qn('w:tblInd')) is not None)
        self.assertEqual(sum(p._p.get_or_add_pPr().find(qn('w:pBdr')) is not None
                             for p in doc.paragraphs), 2)


if __name__ == '__main__':
    unittest.main()
