import json
from pathlib import Path
import shutil
import subprocess
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from docx import Document
from docx.oxml.ns import qn
import book_layout as book

from word_chapters import split_chapter_heading, install_chapters


CASES = [
    ('第一辑：那时我还不知道', ('第一辑', '那时我还不知道')),
    ('第五辑：原来你们是这样看我的', ('第五辑', '原来你们是这样看我的')),
    ('第2辑 标题', ('第2辑', '标题')),
    ('辑三：标题', ('辑三', '标题')),
    ('上辑 标题', ('上辑', '标题')),
    ('第一章 标题', ('第一章', '标题')),
    ('第一幕：初见', ('第一幕', '初见')),
    ('第十二幕 归途', ('第十二幕', '归途')),
    ('第2幕 标题', ('第2幕', '标题')),
    ('第３幕：新的开始', ('第３幕', '新的开始')),
    ('第一幕', ('第一幕', '')),
    ('一幕 标题', ('一幕', '标题')),
    ('幕三：标题', ('幕三', '标题')),
    ('终章：原来这就是成长', ('', '终章：原来这就是成长')),
    ('没有序号的标题', ('', '没有序号的标题')),
]


class ChapterSequenceTests(unittest.TestCase):
    def test_act_sequence_is_displayed_in_all_four_word_styles(self):
        for style in ('版式二','版式三：玫瑰','版式四：纸飞机','版式五：指纹树'):
            full=SimpleNamespace(book=book,build_body=lambda *a,**k:None)
            with patch.object(book,'add_chapter_page',book.add_chapter_page),patch.object(book,'add_page_background',book.add_page_background),patch.object(book,'add_volume_background',book.add_volume_background):
                install_chapters(full,{'chapterStyle':style},[])
                doc=Document();book.add_chapter_page(doc,'第一幕：新的起点','章节引言')
                texts=[node.text for node in doc.part.element.iter(qn('w:t'))]
                if style=='版式二':
                    self.assertIn('第一幕：新的起点',texts)
                else:
                    self.assertIn('第一幕',texts)
                    self.assertIn('新的起点',texts)

    def test_word_sequence(self):
        for heading, expected in CASES:
            with self.subTest(heading=heading):
                self.assertEqual(split_chapter_heading(heading), expected)

    @unittest.skipUnless(shutil.which('node'), 'Node.js is required')
    def test_preview_sequence_matches_word(self):
        source = (Path(__file__).resolve().parents[1] / '排版工具.html').read_text(encoding='utf-8')
        function = source.split('function splitChapterHeading(value){', 1)[1].split('function renderBook(){', 1)[0]
        script = 'function splitChapterHeading(value){' + function
        script += '\nconsole.log(JSON.stringify(' + json.dumps([item[0] for item in CASES]) + '.map(splitChapterHeading)));'
        result = subprocess.run(['node', '-e', script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        actual = json.loads(result.stdout.decode('utf-8'))
        self.assertEqual([(item['number'], item['title']) for item in actual], [item[1] for item in CASES])
