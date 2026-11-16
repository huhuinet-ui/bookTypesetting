import io
import unittest
from unittest.mock import patch
from docx import Document
from docx.oxml.ns import qn
import book_layout as book


class PrintMarginTests(unittest.TestCase):
    def test_paper_expands_without_changing_text_area(self):
        for margin in (0,10):
            with patch.object(book,'PRINT_MARGIN_MM',margin):
                doc, sec = book.new_doc()
                book.add_print_margin(sec)
                book.add_crop_marks(sec)
                stream = io.BytesIO(); doc.save(stream); stream.seek(0)
                sec = Document(stream).sections[0]
                self.assertAlmostEqual(sec.page_width.mm,146+2*margin,places=1)
                self.assertAlmostEqual(sec.page_height.mm,206+2*margin,places=1)
                self.assertAlmostEqual(sec.page_width.mm-sec.left_margin.mm-sec.right_margin.mm,94,places=1)
                self.assertAlmostEqual(sec.page_height.mm-sec.top_margin.mm-sec.bottom_margin.mm,148,places=1)
                self.assertAlmostEqual(sec.left_margin.mm,27+3+margin,places=1)
                rects = list(sec.header._element.iter('{urn:schemas-microsoft-com:vml}rect'))
                white_margin = [r for r in rects if 'margin' in r.get('id','')]
                self.assertEqual(len(white_margin),4 if margin else 0)

    def test_crop_marks_are_outside_bleed_and_align_with_trim(self):
        with patch.object(book,'PRINT_MARGIN_MM',10):
            doc,sec=book.new_doc()
            with patch.object(book,'_crop_rect') as draw:
                book.add_crop_marks(sec)
            for call in draw.call_args_list:
                _,x,y,w,h,_,_=call[0]
                x,y,w,h=[v*25.4/72 for v in (x,y,w,h)]
                self.assertTrue(x+w<=10.3 or x>=155.7 or y+h<=10.3 or y>=215.7)
