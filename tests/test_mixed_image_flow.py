from pathlib import Path
import subprocess
import shutil
import unittest
from unittest.mock import patch
import book_layout as book


class MixedImageFlowTests(unittest.TestCase):
    def test_text_and_image_groups_do_not_force_new_word_sections(self):
        items=[('p','一句说明。'),('img',('a.jpg','图一')),('img',('b.jpg','图二')),
               ('p','另一句说明。'),('img',('c.jpg',None)),('img',('d.jpg',None))]
        doc,_=book.new_doc()
        with patch.object(book,'IMG_ONLY_PAGE_LIFT_LINES',2):
            self.assertEqual(list(book.picture_page_items(doc,items)),items)
        self.assertEqual(len(doc.sections),1)

    def test_all_image_articles_keep_expanded_picture_page(self):
        items=[('img',('a.jpg',None)),('img',('b.jpg',None))]
        doc,section=book.new_doc();top=section.top_margin
        with patch.object(book,'IMG_ONLY_PAGE_LIFT_LINES',2):
            self.assertEqual(list(book.picture_page_items(doc,items)),items)
        self.assertEqual(len(doc.sections),2)
        self.assertLess(doc.sections[-1].top_margin,top)

    @unittest.skipUnless(shutil.which('node'),'Node.js required')
    def test_mixed_browser_flow_has_no_named_page_breaks(self):
        source=(Path(__file__).resolve().parents[1]/'排版工具.html').read_text(encoding='utf-8')
        script='const P={imgOnlyPageLiftLines:2};const assert=require("assert");\n'
        script+=source[source.index('function groupPictureRuns('):source.index('function renderItems(')]
        script+='\nconst images=[\'<figure class="img-unit">one</figure>\',\'<figure class="img-unit">two</figure>\'];'
        script+='\nassert(!groupPictureRuns([\'<div class="para">Text</div>\',...images]).includes("picture-pagination"));'
        script+='\nassert(groupPictureRuns(images).includes("picture-pagination"));'
        subprocess.run(['node','-e',script],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
