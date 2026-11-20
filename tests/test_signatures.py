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
    def test_common_date_signatures_and_body_are_distinguished(self):
        dates=['2026年11月','2026年11月于北京','2026年11月20日 写于上海',
               '二〇二六年十一月','2026.11.20','2026-11','2026/11/20',
               '**2026年11月于北京**','２０２６年１１月','2026年11月\\']
        for date in dates:
            with self.subTest(date=date):
                self.assertEqual(book.classify(['正文。','',date]),[('p','正文。'),('signature',date.rstrip('\\'))])
        for text in ['2026年11月开始记录孩子的成长。','这篇文章写于2026年11月。','2026年11月发生了很多事','![2026年11月](图/a.jpg)']:
            self.assertFalse(book.is_signature_date(text))

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
    def test_preview_keeps_styled_body_with_signature(self):
        source=(Path(__file__).resolve().parents[1]/'排版工具.html').read_text(encoding='utf-8')
        script=source[source.index('function renderItems('):source.index('function renderPoems(')]
        script+='\nfunction renderPara(text){return \'<div class="para body-alt" style="text-align-last:left">\'+text+\'</div>\';}'
        script+='\nfunction renderInline(text){return text;} function groupPictureRuns(items){return items.join("");}'
        script+='\nconsole.log(renderItems([["palt","Body"],["signature","2026年11月于北京"]]));'
        result=subprocess.run(['node','-e',script],stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=True)
        html=result.stdout.decode('utf-8')
        self.assertIn('class="para signature-leading body-alt" style="text-align-last:left"',html)
        self.assertIn('class="signature"',html)

    @unittest.skipUnless(shutil.which('node'), 'Node.js is required')
    def test_preview_classification_matches_word(self):
        source = (Path(__file__).resolve().parents[1] / '排版工具.html').read_text(encoding='utf-8')
        script = source[source.index('const NOTES ='):source.index('function normalizeBookPath(')]
        script += source[source.index('function normalizeBodyBold('):source.index('function classifyPoem(')]
        cases = [['正文。', '', '2026年6月30日', '', '谢林婉'],
                 ['正文。', '', '谢林婉', '', '2026年6月30日'],
                 ['2026年6月30日', '', '##### 小标题']]
        cases += [['正文。','','2026年11月于北京'],['正文。','','二〇二六年十一月'],
                  ['正文。','','**2026年11月于北京**'],['正文。','','2026/11/20'],['正文。','','2026年11月\\']]
        script += '\nconsole.log(JSON.stringify(' + json.dumps(cases) + '.map(lines=>classify(lines,false))));'
        result = subprocess.run(['node', '-e', script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        actual = json.loads(result.stdout.decode('utf-8'))
        expected = json.loads(json.dumps([book.classify(lines) for lines in cases]))
        self.assertEqual(actual, expected)
