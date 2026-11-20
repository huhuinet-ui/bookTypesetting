import base64
import contextlib
import io
from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from word2md import convert, paragraph_markdown


class Word2MarkdownBoldTests(unittest.TestCase):
    def test_repeated_manual_breaks_become_indented_paragraphs(self):
        doc=Document()
        p=doc.add_paragraph('第一段正文。\n\n第二段正文。\n段内换行。\n \n')
        p.add_run('第三段加粗。').bold=True
        text,_=self.convert_document(doc)
        self.assertEqual(text,'第一段正文。\n\n第二段正文。\\\n段内换行。\n\n**第三段加粗。**\n')
        import book_layout as book
        items=book.classify(text.splitlines())
        self.assertEqual(items,[('p','第一段正文。'),('p','第二段正文。\n段内换行。'),('p','**第三段加粗。**')])
        output=Document()
        for _,value in items:
            paragraph=output.add_paragraph()
            book.set_para(paragraph,indent_chars=book.BODY_INDENT)
            book.add_body_para(paragraph,value)
        for paragraph in output.paragraphs:
            self.assertEqual(paragraph.paragraph_format.first_line_indent.pt,book.BODY_SIZE*book.BODY_INDENT)

    def test_afterword_heading_maps_to_front_matter(self):
        doc=Document()
        doc.add_heading('后记',level=4)
        doc.add_paragraph('后记正文。')
        text,_=self.convert_document(doc)
        self.assertEqual(text,'# 后记\n\n后记正文。\n')

    def test_appendix_heading_is_independent_from_chapters(self):
        doc=Document()
        doc.add_heading('附录一：年表',level=3)
        doc.add_paragraph('资料正文。')
        doc.add_heading('附录二：参考资料',level=4)
        doc.add_paragraph('参考内容。')
        text,_=self.convert_document(doc)
        self.assertIn('# 附录一：年表\n',text)
        self.assertIn('# 附录二：参考资料\n',text)
        self.assertNotIn('### 附录',text)

    def convert_document(self, doc, **kwargs):
        # Save and reopen through the real converter, rather than only testing runs.
        work = Path(__file__).resolve().parents[1] / "tmp" / "word2md-tests"
        work.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=work) as directory:
            source = Path(directory) / "input.docx"
            output = Path(directory) / "output.md"
            doc.save(source)
            with contextlib.redirect_stdout(io.StringIO()):
                _, images = convert(str(source), str(output), **kwargs)
            return output.read_text(encoding="utf-8"), images

    def test_mixed_body_and_adjacent_bold_runs(self):
        doc = Document()
        doc.add_heading("文章标题", level=4)
        para = doc.add_paragraph("普通正文，")
        para.add_run("重点").bold = True
        para.add_run()  # An empty run must not split the emphasis.
        para.add_run("内容").bold = True
        para.add_run("，继续正文。")
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "#### 文章标题\n\n普通正文，**重点内容**，继续正文。\n")

    def test_whitespace_and_soft_line_breaks(self):
        doc = Document()
        para = doc.add_paragraph("前文")
        para.add_run("  加粗一\n加粗二 \t").bold = True
        para.add_run("后文")
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "前文  **加粗一**\\\n**加粗二** \t后文\n")

    def test_enter_and_shift_enter_remain_distinct(self):
        doc = Document()
        doc.add_paragraph('第一段')
        doc.add_paragraph('第二段上行\n第二段下行')
        text, _ = self.convert_document(doc)
        self.assertEqual(text, '第一段\n\n第二段上行\\\n第二段下行\n')

    def test_inherited_character_style_and_explicit_off(self):
        doc = Document()
        base = doc.styles.add_style("BoldBase", WD_STYLE_TYPE.CHARACTER)
        base.font.bold = True
        derived = doc.styles.add_style("BoldChild", WD_STYLE_TYPE.CHARACTER)
        derived.base_style = base
        para = doc.add_paragraph("正文")
        para.add_run("样式加粗").style = derived
        run = para.add_run("明确取消")
        run.style = derived
        run.bold = False
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "正文**样式加粗**明确取消\n")

    def test_inherited_paragraph_style(self):
        doc = Document()
        base = doc.styles.add_style("BoldBody", WD_STYLE_TYPE.PARAGRAPH)
        base.font.bold = True
        child = doc.styles.add_style("ChildBody", WD_STYLE_TYPE.PARAGRAPH)
        child.base_style = base
        para = doc.add_paragraph("这段正文通过段落样式继承加粗。", style=child)
        para.add_run("这部分取消加粗。").bold = False
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "**这段正文通过段落样式继承加粗。**这部分取消加粗。\n")

    def test_style_toggle_does_not_false_positive(self):
        doc = Document()
        base = doc.styles.add_style("FirstToggle", WD_STYLE_TYPE.CHARACTER)
        base.font.bold = True
        child = doc.styles.add_style("SecondToggle", WD_STYLE_TYPE.CHARACTER)
        child.base_style = base
        child.font.bold = True
        para = doc.add_paragraph("正文")
        para.add_run("两次样式切换后不加粗").style = child
        self.assertEqual(paragraph_markdown(para), "正文两次样式切换后不加粗")

    def test_document_default_bold_and_direct_override(self):
        doc = Document()
        props = doc.styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr")[0]
        props.append(OxmlElement("w:b"))
        para = doc.add_paragraph("继承默认加粗")
        para.add_run("直接取消").bold = False
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "**继承默认加粗**直接取消\n")

    def test_full_bold_long_body_and_plain_body(self):
        doc = Document()
        doc.add_paragraph("这一段没有加粗。")
        body = "这是一段完整加粗的较长正文，用于确认整段文字的加粗能够保留，同时不会被短标题猜测规则当成文章标题。"
        doc.add_paragraph().add_run(body).bold = True
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "这一段没有加粗。\n\n**" + body + "**\n")

    def test_headings_and_captions_remain_plain(self):
        doc = Document()
        doc.add_heading("篇章", level=3).runs[0].bold = True
        doc.add_heading("标题：文章", level=4).runs[0].bold = True
        doc.add_heading("【小标题】", level=5).runs[0].bold = True
        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC"
            "AAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII="
        )
        para = doc.add_paragraph("图旁的普通文字，")
        para.add_run("保留重点").bold = True
        para.add_run().add_picture(io.BytesIO(png))
        doc.add_paragraph()  # Blank paragraphs before a caption are supported.
        doc.add_paragraph().add_run("图1-1：加粗图注").bold = True
        doc.add_paragraph("后续正文。")
        text, images = self.convert_document(doc, bracket_subtitle=True)
        self.assertEqual(images, 1)
        self.assertIn("### 篇章\n\n#### 文章\n\n##### 小标题", text)
        self.assertIn("图旁的普通文字，**保留重点**", text)
        self.assertIn("![图1-1：加粗图注](图/img1.png)", text)
        self.assertNotIn("**加粗图注**", text)
        self.assertIn("后续正文。", text)

    def test_isolated_short_bold_body_is_not_a_heading(self):
        doc = Document()
        doc.add_paragraph().add_run("推测标题").bold = True
        doc.add_paragraph()
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "**推测标题**\n")

    def test_growth_diary_bold_paragraphs_keep_body_semantics(self):
        samples = [
            '用打针来吓唬小朋友是有代价的，因为那一刻真的会来临。',
            '我们脆弱的父子关系似乎只能靠巧克力饼干来维持了。',
            '第一个原则是从来不骗他。',
            '第二个原则是尽可能不要限制他的自由。',
            '第三个原则：对他保持尊重。',
            '（2）姥爷眼里的皮皮——2026.9.15 皮姥爷记述',
        ]
        doc = Document()
        doc.add_heading('真正的文章标题', level=4)
        for sample in samples:
            doc.add_paragraph()
            doc.add_paragraph().add_run(sample).bold = True
            doc.add_paragraph()
        text, _ = self.convert_document(doc)
        self.assertEqual(text.count('#### '), 1)
        for sample in samples:
            self.assertIn('**' + sample + '**', text)

    def test_hyperlinks_keep_text_order_and_bold(self):
        doc = Document()
        para = doc.add_paragraph("前文")
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), "bookmark")
        run = OxmlElement("w:r")
        props = OxmlElement("w:rPr")
        props.append(OxmlElement("w:b"))
        run.append(props)
        text = OxmlElement("w:t")
        text.text = "链接文字"
        run.append(text)
        hyperlink.append(run)
        para._p.append(hyperlink)
        para.add_run("后文")
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "前文**链接文字**后文\n")

    def test_heading_character_style_recognized_after_body_without_blank(self):
        doc = Document()
        style = doc.styles.add_style('Heading 4 Char_imported', WD_STYLE_TYPE.CHARACTER)
        style.font.bold = True
        doc.add_paragraph('上一篇的最后一段正文。')
        para = doc.add_paragraph('\u00a0\u00a0')
        para.add_run('寒冷除夕夜，最暖母女情').style = style
        doc.add_paragraph()
        doc.add_paragraph('新文章的正文。')
        text, _ = self.convert_document(doc)
        self.assertIn('#### 寒冷除夕夜，最暖母女情', text)
        self.assertNotIn('**寒冷除夕夜，最暖母女情**', text)

    def test_mixed_heading_character_style_is_still_body(self):
        doc = Document()
        style = doc.styles.add_style('Heading 4 Char_imported', WD_STYLE_TYPE.CHARACTER)
        style.font.bold = True
        para = doc.add_paragraph('普通正文里')
        para.add_run('局部强调').style = style
        para.add_run('继续说明。')
        text, _ = self.convert_document(doc)
        self.assertEqual(text, '普通正文里**局部强调**继续说明。\n')

    def test_inherited_heading_character_style_is_recognized(self):
        doc = Document()
        parent = doc.styles.add_style('Heading 4 Char_imported', WD_STYLE_TYPE.CHARACTER)
        child = doc.styles.add_style('ImportedArticleTitle', WD_STYLE_TYPE.CHARACTER)
        child.base_style = parent
        doc.add_paragraph('前文。')
        doc.add_paragraph().add_run('继承字符样式的文章标题').style = child
        text, _ = self.convert_document(doc)
        self.assertIn('#### 继承字符样式的文章标题', text)

    def test_missing_default_paragraph_style_keeps_body_and_bold(self):
        doc = Document()
        para = doc.add_paragraph('普通正文，')
        para.add_run('保留加粗').bold = True
        para.add_run('，继续正文。')
        doc.styles.element.remove(doc.styles['Normal'].element)
        self.assertIsNone(para.style)
        text, _ = self.convert_document(doc)
        self.assertEqual(text, '普通正文，**保留加粗**，继续正文。\n')

    def test_undefined_paragraph_style_does_not_abort_conversion(self):
        doc = Document()
        para = doc.add_paragraph('未定义样式的正文。')
        para._p.style = 'MissingBodyStyle'
        doc.styles.element.remove(doc.styles['Normal'].element)
        self.assertIsNone(para.style)
        text, _ = self.convert_document(doc)
        self.assertEqual(text, '未定义样式的正文。\n')

    def test_image_and_caption_work_without_default_paragraph_style(self):
        doc = Document()
        png = base64.b64decode(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC'
            'AAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII='
        )
        doc.add_paragraph().add_run().add_picture(io.BytesIO(png))
        doc.add_paragraph()
        caption = doc.add_paragraph('图注：样式缺失的图注')
        doc.styles.element.remove(doc.styles['Normal'].element)
        self.assertIsNone(caption.style)
        text, images = self.convert_document(doc)
        self.assertEqual(images, 1)
        self.assertEqual(text, '![样式缺失的图注](图/img1.png)\n')


if __name__ == "__main__":
    unittest.main()
