import io
from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.oxml.ns import qn
import book_layout as book
import book_full as full
from word_pagination import measured_pages, save_paginated


class FinalPaginationTests(unittest.TestCase):
    def test_restarting_pages_replaces_inherited_reset(self):
        doc = Document()
        section = doc.sections[0]
        full.set_page_number_start(section._sectPr, 5)
        full.set_page_number_start(section._sectPr, 1)
        starts = section._sectPr.findall(qn('w:pgNumType'))
        self.assertEqual(len(starts), 1)
        self.assertEqual(starts[0].get(qn('w:start')), '1')
        full.remove_page_number_start(section._sectPr)
        self.assertFalse(section._sectPr.findall(qn('w:pgNumType')))

    def test_bookmark_does_not_add_visible_text_or_runs(self):
        doc = Document()
        paragraph = doc.add_paragraph('A title exactly filling a line')
        before = (paragraph.text, len(paragraph.runs))
        book.add_anchor(paragraph, 'zq0000wz')
        book.add_anchor(doc.add_paragraph('Next'), 'zq0001wz')
        out = io.BytesIO(); doc.save(out); out.seek(0)
        loaded = Document(out)
        self.assertEqual((loaded.paragraphs[0].text, len(loaded.paragraphs[0].runs)), before)
        starts = loaded.part.element.findall('.//' + qn('w:bookmarkStart'))
        self.assertEqual([s.get(qn('w:name')) for s in starts], ['zq0000wz', 'zq0001wz'])
        self.assertEqual(len({s.get(qn('w:id')) for s in starts}), 2)

    def test_uses_printed_page_numbers_instead_of_pdf_page_index(self):
        result = {'pages': {'a': 2}, 'physical': {'a': 16}}
        self.assertEqual(measured_pages(result, [('a', 'Title')]), {'Title': 2})
        with self.assertRaises(RuntimeError):
            measured_pages(result, [('missing', 'Title')])

    def test_rebuilds_whole_book_until_pages_and_parity_stabilize(self):
        builds, engines = [], []
        def build(pages, pad):
            builds.append((dict(pages), pad))
            doc = Document()
            doc.add_paragraph(str(pages))
            return doc, [('a', 'Title')]
        def inspect(path, bookmarks, engine=None):
            self.assertTrue(Path(path).is_file())
            self.assertEqual(bookmarks, ['a'])
            engines.append(engine)
            return {'pages': {'a': 2}, 'physical': {'a': 16}, 'total': 31, 'engine': 'kwps.Application'}
        root = Path(__file__).resolve().parents[1] / 'tmp' / 'pagination-tests'
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as directory:
            output = Path(directory) / 'result.docx'
            result = save_paginated(build, output, inspect=inspect)
            self.assertTrue(output.is_file())
        self.assertEqual(builds, [({}, 3), ({'Title': 2}, 2)])
        self.assertEqual(engines, [None, 'kwps.Application'])
        self.assertEqual(result[:3], ({'Title': 2}, 17, 2))

    def test_unstable_pagination_does_not_replace_existing_output(self):
        calls = []
        def build(pages, pad):
            return Document(), [('a', 'Title')]
        def inspect(path, bookmarks, engine=None):
            page = 2 if len(calls) % 2 == 0 else 3
            calls.append(page)
            return {'pages': {'a': page}, 'physical': {'a': page}, 'total': 10, 'engine': 'Word.Application'}
        root = Path(__file__).resolve().parents[1] / 'tmp' / 'pagination-tests'
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=str(root)) as directory:
            output = Path(directory) / 'existing.docx'
            output.write_bytes(b'original')
            with self.assertRaises(RuntimeError):
                save_paginated(build, output, inspect=inspect)
            self.assertEqual(output.read_bytes(), b'original')
        self.assertEqual(len(calls), 3)
