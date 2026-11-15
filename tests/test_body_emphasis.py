import unittest
from unittest.mock import patch

from docx import Document

import book_layout as book


class BodyEmphasisTests(unittest.TestCase):
    def test_short_emphasized_text_is_body_not_guessed_heading(self):
        for text in ('**短句。**', '**01**', '**重点：**继续说明', '**前段****后段**'):
            self.assertEqual(book.classify([text, ''], base_dir='.'), [('p', text)])

    def test_multiline_emphasis_is_balanced_per_paragraph(self):
        lines = ['**第一行', '', '第二行', '', '最后一行**', '', '普通正文。']
        self.assertEqual(book.classify(lines, base_dir='.'), [
            ('p', '**第一行**'), ('p', '**第二行**'),
            ('p', '**最后一行**'), ('p', '普通正文。')])

    def test_only_explicit_headings_are_headings(self):
        self.assertEqual(book.classify(['##### 正式标题', '', '普通短标题', ''], base_dir='.'),
                         [('h', '正式标题'), ('p', '普通短标题')])
        self.assertEqual(book.normalize_body_bold(['**开始', '##### 新标题', '普通正文。']),
                         ['**开始**', '##### 新标题', '普通正文。'])

    def test_word_retains_bold_when_requested(self):
        doc = Document()
        with patch.object(book, 'MARK_MODE', 'bold'):
            for _, text in book.classify(['**短句。**', '', '**第一行', '最后一行**'], base_dir='.'):
                paragraph = doc.add_paragraph()
                book.add_body_para(paragraph, text)
        self.assertEqual([p.text for p in doc.paragraphs], ['短句。', '第一行\n最后一行'])
        self.assertTrue(all(run.bold for p in doc.paragraphs for run in p.runs if run.text.strip()))

    def test_color_mode_remains_an_explicit_choice(self):
        paragraph = Document().add_paragraph()
        with patch.object(book, 'MARK_MODE', 'color'):
            book.add_body_para(paragraph, '**短句。**')
        self.assertEqual(paragraph.text, '短句。')
        self.assertFalse(paragraph.runs[0].bold)


if __name__ == '__main__':
    unittest.main()
