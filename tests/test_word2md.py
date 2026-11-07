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
        self.assertEqual(text, "前文  **加粗一**\n**加粗二** \t后文\n")

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

    def test_legacy_short_bold_heading_guess_stays_plain(self):
        doc = Document()
        doc.add_paragraph().add_run("推测标题").bold = True
        doc.add_paragraph()
        text, _ = self.convert_document(doc)
        self.assertEqual(text, "#### 推测标题\n")

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


if __name__ == "__main__":
    unittest.main()
