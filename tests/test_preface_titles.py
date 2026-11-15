import contextlib
import io
from pathlib import Path
import unittest
from docx import Document
import book_layout as book
from word2md import convert
import test_volumes as fixtures


class PrefaceTitleTests(unittest.TestCase):
    setUp = fixtures.VolumeTests.setUp
    parse = fixtures.VolumeTests.parse

    def test_bare_keywords_are_preserved(self):
        for key in book.PREFACE_KEYS:
            with self.subTest(keyword=key):
                data = self.parse('# ' + key + '\n\n前言正文。\n\n#### 文章\n正文。')
                self.assertEqual(data['preface_title'], '前言' if key == '前沿' else key)
                self.assertEqual(data['preface'], '前言正文。')

    def test_custom_titles_drop_keyword_and_separator(self):
        for key in book.PREFACE_KEYS:
            for separator in ('', ' ', '：', ': ', '——', ' — '):
                with self.subTest(keyword=key, separator=separator):
                    data = self.parse('# ' + key + separator + '写在出发之前\n\n正文。')
                    self.assertEqual(data['preface_title'], '写在出发之前')
                    self.assertIsNone(data['title_page'])

    def test_word_import_keeps_keyword_for_preface_roundtrip(self):
        doc = Document()
        doc.add_heading('前沿：写在出发之前', level=4)
        doc.add_paragraph('前言正文。')
        source = Path(self.temp.name) / 'input.docx'
        output = Path(self.temp.name) / 'output.md'
        doc.save(str(source))
        with contextlib.redirect_stdout(io.StringIO()):
            convert(str(source), str(output))
        markdown = output.read_text(encoding='utf-8')
        self.assertIn('# 前言：写在出发之前', markdown)
        self.assertEqual(book.parse_book(str(output))['preface_title'], '写在出发之前')
