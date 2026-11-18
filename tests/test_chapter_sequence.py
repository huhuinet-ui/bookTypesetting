import json
from pathlib import Path
import shutil
import subprocess
import unittest

from word_chapters import split_chapter_heading


CASES = [
    ('第一辑：那时我还不知道', ('第一辑', '那时我还不知道')),
    ('第五辑：原来你们是这样看我的', ('第五辑', '原来你们是这样看我的')),
    ('第2辑 标题', ('第2辑', '标题')),
    ('辑三：标题', ('辑三', '标题')),
    ('上辑 标题', ('上辑', '标题')),
    ('第一章 标题', ('第一章', '标题')),
    ('终章：原来这就是成长', ('', '终章：原来这就是成长')),
    ('没有序号的标题', ('', '没有序号的标题')),
]


class ChapterSequenceTests(unittest.TestCase):
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
