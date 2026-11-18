import io
import json
from pathlib import Path
import shutil
import subprocess
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import book_layout as book


class SignatureTests(unittest.TestCase):
    def test_both_orders_with_and_without_blank_lines(self):
        for pair in [('谢林婉', '2026年6月30日'), ('2026年6月30日', '谢林婉')]:
            for gap in [[], [''], ['', '']]:
                lines = ['正文。', '', pair[0]] + gap + [pair[1], '']
                self.assertEqual(book.classify(lines), [('p', '正文。'), ('signature', '\n'.join(pair))])

    def test_date_does_not_absorb_following_body_or_heading(self):
        self.assertEqual(book.classify(['2026年6月30日', '', '这是正文。']),
                         [('signature', '2026年6月30日'), ('p', '这是正文。')])
        self.assertEqual(book.classify(['2026年6月30日', '', '##### 小标题']),
                         [('signature', '2026年6月30日'), ('h', '小标题')])

    def test_word_alignment_and_page_binding_survive_save(self):
        doc = Document()
        book.add_front_page(doc, '后记', '正文。\n\n2026年6月30日\n\n谢林婉')
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        paragraphs = Document(out).paragraphs
        signature = paragraphs[-1]
        self.assertEqual(signature.text, '2026年6月30日\n谢林婉')
        self.assertEqual(signature.alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertTrue(signature.paragraph_format.keep_together)
        self.assertTrue(paragraphs[-2].paragraph_format.keep_with_next)

    @unittest.skipUnless(shutil.which('node'), 'Node.js is required')
    def test_preview_classification_matches_word(self):
        source = (Path(__file__).resolve().parents[1] / '排版工具.html').read_text(encoding='utf-8')
        script = source[source.index('const NOTES ='):source.index('function normalizeBookPath(')]
        script += source[source.index('function normalizeBodyBold('):source.index('function classifyPoem(')]
        cases = [['正文。', '', '2026年6月30日', '', '谢林婉'],
                 ['正文。', '', '谢林婉', '', '2026年6月30日'],
                 ['2026年6月30日', '', '##### 小标题']]
        script += '\nconsole.log(JSON.stringify(' + json.dumps(cases) + '.map(lines=>classify(lines,false))));'
        result = subprocess.run(['node', '-e', script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        actual = json.loads(result.stdout.decode('utf-8'))
        expected = json.loads(json.dumps([book.classify(lines) for lines in cases]))
        self.assertEqual(actual, expected)
