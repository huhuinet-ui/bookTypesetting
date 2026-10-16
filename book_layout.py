# -*- coding: utf-8 -*-
"""把 .md 纯文本自动排成一本 Word 书稿（参考书风格）。

用法（在命令行里运行）：
    python book_layout.py "要排版的书.md" ["书名"]
    例如：python book_layout.py "深度写作营.md" "我与我·深度写作营文集"

运行后，会在源文件同一目录下生成一个名字带"_排版"的 Word 文件。
想改样式：只改下面"样式参数区"里的数值，改完重新运行本脚本即可。

--------------------- 源文件的约定写法 ---------------------
  # 自序 / # 前言 / # 序言 / # 代序 / # 写在前面
                         （可选）前言页；排在目录之前，下面是前言正文（可多段）
  # 后记 / # 跋          （可选）后记页；排在正文之后，下面是后记正文（可多段）

  ### 第一章            （可选）章节名，章节页会单独占一页
  简介：……             （可选）紧跟章节名写一行，作为章节页文本框里的引言

  #### 标题：文章标题    （必填）一篇文章的开头
  作者：xxx             （可选）作者；没有作者就不画标题下划线和作者名
  诗歌：                （可选）整篇按诗歌排版：每个 ##### 是一首诗（一首一页、诗句居中）
  （正文段落，空行分段）

  ##### 小标题：xxx      （可选）文章内的小节标题，加粗左对齐

  图片：C:\\图\\a.jpg   （可选）插图，插在正文当前位置（自定义写法）
  图片名：这是图片名称  （可选）图注，紧跟图片下方、居中
  也支持标准 Markdown 写法：![图注](图/a.jpg)   （Pandoc 等工具从 Word 转出的图）
  图片路径写相对路径时，会按「本 .md 文件所在目录」去找（换电脑不用改路径）
-----------------------------------------------------------
"""
import io, os, re, sys, math, html
if (getattr(sys.stdout, "encoding", "") or "").lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

# ============================================================
#  样式参数区 —— 想改排版，只改这里（改完重新运行本脚本）
#  单位：mm=毫米  cm=厘米  pt=磅(1磅≈0.35毫米)
#  常用字号对照：小五=9磅、五号=10.5磅、小四=12磅、四号=14磅、三号=16磅、小二=18磅
# ============================================================

# ---- 输入与输出 ----
SRC = sys.argv[1] if len(sys.argv) > 1 else r"D:\纸质文集\我与我·深度写作营文集(第一期).md"
# SRC：要排版的源文件（.md 纯文本）。一般不用改，运行命令时把文件名传进来即可。
BOOK_NAME = sys.argv[2] if len(sys.argv) > 2 else "我与我·深度写作营文集"
# BOOK_NAME：显示在每一页页眉上的书名。
LIMIT = None
# LIMIT：只排前几篇。None=排全部；填 2 就只排前 2 篇（改样式想快速看效果时用）。
NAME = os.path.splitext(os.path.basename(SRC))[0] + "_排版"
# NAME：生成的 Word 文件名（默认 = 源文件名 + "_排版"），一般不用改。
OUT_DIR = os.path.dirname(os.path.abspath(SRC))
# OUT_DIR：生成的 Word 文件保存到哪个文件夹（默认和源文件放一起），一般不用改。

# ---- 纸张与页边距 ----
PAGE_W, PAGE_H = 166, 226
# 纸张大小（毫米）：宽 166 × 高 226，约等于 16 开大小的书。这是裁切后的「成品尺寸」。
BLEED_MM = 3
# 出血量（毫米）：0=不加出血（屏幕看/普通打印）；送印厂付印时改成 3。
# 设成 3 后页面自动从 166×226 变 172×232、边距同步外扩，正文位置不变，全书底色/插图能铺满到边、裁切后不留白边。
M_TOP, M_BOT = 45, 33
# 页边距（毫米）：上边距 45、下边距 33，即页面上下边缘到正文的距离（参考《秀茹的一生》实测）。
M_LEFT, M_RIGHT = 40, 32
# 页边距（毫米）：内侧 40（装订/书脊一侧）、外侧 32。镜像页边距已开启：奇数页内侧在左、偶数页内侧在右。
HEADER_DIST = 24
# 页眉位置：页眉文字距离页面顶部 24 毫米（参考《秀茹的一生》实测）。

# ---- 字体（需要电脑里装了这些字体才生效，否则 Word 会用宋体替代显示）----
F_TITLE = "方正清刻本悦宋简体"    # 文章标题 / 章节名 / 前言后记标题用什么字体
F_AUTHOR = "方正楷体简体"        # 作者名字用什么字体
F_BODY = "方正书宋_GBK"          # 正文用什么字体
F_HEAD = "方正清刻本悦宋简体"    # 页眉上的书名用什么字体
F_NUM = "方正楷体简体"           # 页码用什么字体

# ---- 文章标题 ----
TITLE_SIZE = 14
# 标题字号（磅）：14 磅，就是 Word 里的"四号"。
TITLE_TOP_LINES = 3
# 标题上方空几行：每篇开头，标题从页面顶部往下空 3 行才开始（参考《秀茹的一生》：标题距上边距约 20mm）。
RULE_W, RULE_COLOR, RULE_PT = 2.54, "000000", 0.5
# 标题下面的那条横线：宽 2.54 厘米、黑色(#000000)、粗细 0.5 磅。
TITLE_BODY_GAP_LINES = 3
# 标题（含横线、作者名）到正文之间空几行：0=紧挨着，几就是空几行。
DECO_W, DECO_H = 34, 73
# 装饰矩形（标题左侧的竖条色块）：宽 34 磅 × 高 73 磅。填充色=书籍主色调(BOOK_COLOR)。
DECO_X = 28.35
# 装饰矩形左边 距成品左边缘 10mm（28.35 磅，参考《秀茹的一生》实测，位于装订侧页边距内）。
# 镜像：奇数页(右页)装订边在左→色块在左；偶数页(左页)装订边在右→色块翻到右。由 book_full.py 按真实页码判定。
DECO_TOP = 177.17
# 装饰矩形顶边 距成品顶边缘 62.5mm（177.17 磅），比标题顶高约 2.7mm（参考《秀茹的一生》实测）。

# ---- 书籍主色调（统一控制：作者名、章节页底色、标题装饰矩形）----
BOOK_COLOR = RGBColor(0xDE, 0xB8, 0x87)
# 书籍主色调：棕褐色 #DEB887。同时用于①作者名颜色 ②章节页整页底色 ③每篇标题处的装饰矩形。
# 换主色只改这一行即可（改成黑色就写 RGBColor(0, 0, 0)）。

# ---- 作者名 ----
AUTHOR_SIZE = 13
# 作者名字号（磅）：13 磅。

# ---- 正文 ----
BODY_SIZE = 9
# 正文字号（磅）：9 磅，就是 Word 里的"小五"。
BODY_INDENT = 1.5
# 首行缩进：每个自然段的第一行往里缩进 1.5 个汉字。
BODY_AFTER = 19.5
# 段后间距（磅）：每段结束后空 19.5 磅，让段与段之间有空隙。
BODY_LINE = 19.5
# 行距（磅）：行与行之间固定 19.5 磅，约等于字号的 2 倍，行距较宽松。

# ---- 正文标记着色（源文件里用 **文字** 包起来的部分）----
MARK_COLOR_ON = True
# 开关：**文字** 标记的部分要不要上色（只上色、不加粗）。True=上色；False=不上色、保持黑色。
MARK_COLOR = None
# 标记文字的颜色。None=自动用书籍主色调 BOOK_COLOR；想自定义就写 RGBColor(红, 绿, 蓝)，例如黑色 RGBColor(0, 0, 0)。

# ---- 分页控制（解决「一页只剩一两行字」的孤行）----
WIDOW_CONTROL = True
# 孤行控制：段落跨页时，Word 每页至少保留 2 行，消除「下一页只剩 1 行」的孤行。
# 这是 Word 排版里最常用、零副作用的做法，建议一直开着。
KEEP_TOGETHER = False
# 段中不分页：整个段落不在页中间断开，放不下就整段挪到下一页。
# 能彻底消除「一页只剩 1~2 行」，但段落偏长时页面底部会留白。想更彻底就改成 True。

# ---- 小标题（文章内的小节标题，如「目标感治好了我的矫情」）----
SUBTITLE_BOLD = True
# 小标题是否加粗。True=加粗左对齐（和正文同字号 9 磅）。

# ---- 章节页（篇章页）----
CHAPTER_SIZE = 18
# 章节名字号（磅）：14 磅，和文章标题一样大。
CHAPTER_TOP_LINES = 8
# 章节名从页顶往下空几行（用空段落占位，1 行≈正文字行高）。
CHAPTER_RULE_W = 4.36
# 章节名下面的居中下划线宽度（厘米）。
CHAPTER_RULE_PT = 1.0
# 章节名下面的下划线粗细（磅）。

# ---- 章节页文本框（章节名下方，可写一小段引言）----
INTRO_SIZE = 12
# 文本框内文字字号（磅）：12 磅。
INTRO_BOX_W = 8.34
# 文本框宽度（厘米），页面里居中摆放；框内文字左对齐。

# ---- 前言 / 后记页 ----
FRONT_SIZE = 20
# 前言/后记大标题字号（磅）：20 磅加粗居中，和目录大标题一样。
FRONT_TOP_LINES = 3
# 前言/后记标题从页顶往下空几行。

# ---- 插图 ----
CAPTION_SIZE = 9
# 图注字号（磅）：9 磅，居中加粗，放在图片正下方。
IMAGE_MAX_W = None
# 插图最大宽度（厘米）。None=自动用版心宽（=纸张宽-左右边距），超过会自动缩到该宽度。
IMG_MAX_W_RATIO = 0.8
# 图片最大宽度占版心宽的比例（0~1）。改小图片会等比变小，一页能多放几张；1.0=占满版心宽。

# ---- 插图装饰（图片底下衬色块 + 图注引线，参考《秀茹的一生》）----
IMG_BLOCK_ON = True
# 开关：图片底下要不要衬一块色块（像照片贴在卡纸上）。True=有；False=没有。
IMG_BLOCK_COLOR = None
# 色块颜色。None=自动用书籍主色调 BOOK_COLOR；想自定义就写 RGBColor(红,绿,蓝)。
IMG_BLOCK_SIZE = 0.7
# 色块的尺寸 = 图片长宽的该比例。0.5=色块长宽各为图片的一半。
IMG_BLOCK_EXPOSE = 6
# 色块往左下角露出的量（磅）：左、下各露出这么多，其余部分被图片盖住。
IMG_LINE_ON = True
# 开关：图注要不要用折线连到图片（引线）。True=有；False=没有。
IMG_LINE_COLOR = None
# 引线颜色。None=黑色；想自定义就写 RGBColor(红,绿,蓝)。
IMG_LINE_GAP = 14
# 图片和图注之间的空隙（磅）：引线竖线就画在这段空隙里，图注会跟着往下移。

# ---- 诗歌排版（源文件里用「诗歌：」标记整篇为诗歌）----
POEM_TITLE_SIZE = 12
# 诗题字号（磅）。
POEM_TOP_LINES = 4
# 每首诗（从第二首起）从页顶往下空几行。
POEM_TITLE_AFTER = 12
# 诗题到副题/诗句的空隙（磅）。
POEM_META_SIZE = 9
# 副题（——xxx）/日期（xxxx年xx月）字号（磅）。
POEM_META_AFTER = 10
# 副题/日期到诗句的空隙（磅）。
POEM_LINE_SIZE = 10.5
# 诗句字号（磅）。
POEM_LINE = 26
# 诗句行距（磅，固定）。
POEM_STANZA_GAP = 14
# 诗节（空行）之间的额外空隙（磅）。

# ---- 页眉与页码 ----
HEAD_SIZE = 6
# 页眉第一行字号（磅）：6 磅小字（奇数页=章节名、偶数页=书名）。
NUM_SIZE = 7
# 页码字号（磅）：7 磅，带下划线。
# 页眉规则：奇数页「章节名+页码」靠右（右对齐）；偶数页「书名+页码」靠左（左对齐）。
# 原因：书翻开时，奇数页（右页）的外侧在右、偶数页（左页）的外侧在左，页眉统一放外侧。
# 页码按全书页数补零（87 页→01、123 页→001），位数由 book_full.py 自动算。

# ---- 四角裁切标记（角线）----
CROP_MARKS = True
# 是否在每页四角加裁切角线（印刷裁切定位用）。True=加；屏幕预览想干净就设 False。
CROP_BLACK_W = 0.1
# 角线黑线粗细（毫米）。
CROP_WHITE_W = 0.5
# 角线白线（挖空）粗细（毫米）。白线比黑线粗、垫在黑线下面，让黑线在有色底色上也看得清。
# ============================================================

LEFT = WD_ALIGN_PARAGRAPH.LEFT
RIGHT = WD_ALIGN_PARAGRAPH.RIGHT
CENTER = WD_ALIGN_PARAGRAPH.CENTER
JUSTIFY = WD_ALIGN_PARAGRAPH.JUSTIFY
BLACK = RGBColor(0, 0, 0)


def R(p, text, cn=F_BODY, size=BODY_SIZE, color=BLACK, bold=False,
      underline=False, af="Times New Roman"):
    r = p.add_run(text)
    r.font.name = af
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.underline = underline
    r.font.color.rgb = color
    r._element.rPr.rFonts.set(qn("w:eastAsia"), cn)
    return r


def add_body_para(p, text, cn=F_BODY, size=BODY_SIZE):
    """渲染正文段落，支持 **着色** 标记（**文字** 只上色、不加粗）。着色部分按 MARK_COLOR_ON/MARK_COLOR 控制。"""
    mark_color = (MARK_COLOR if MARK_COLOR is not None else BOOK_COLOR) if MARK_COLOR_ON else BLACK
    for i, seg in enumerate(text.split("**")):
        if not seg:
            continue
        colored = (i % 2 == 1)
        R(p, seg, cn=cn, size=size,
          color=(mark_color if colored else BLACK))


def add_page_field(p, size, cn, color=BLACK, underline=False, pad=None):
    # pad：页码补零到几位（pad=3 → 001、002……）；None=不补零
    instr = ' PAGE \\# "%s" ' % ("0" * pad) if pad else " PAGE "
    r1 = p.add_run(); f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin"); r1._r.append(f1)
    r2 = p.add_run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr; r2._r.append(it)
    r3 = p.add_run(); f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate"); r3._r.append(f2)
    r4 = p.add_run(str(1).zfill(pad) if pad else "1")
    r5 = p.add_run(); f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end"); r5._r.append(f3)
    for r in (r1, r2, r3, r4, r5):
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.underline = underline
        rPr = r._element.get_or_add_rPr()
        rf = rPr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); rPr.append(rf)
        rf.set(qn("w:eastAsia"), cn)
    return r4


def set_para(p, align=LEFT, before=0, after=0, line=None, indent_chars=None):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = Pt(line) if line != 1.0 else 1.0
    pf.widow_control = WIDOW_CONTROL
    if KEEP_TOGETHER:
        pf.keep_together = True
    if indent_chars:
        pPr = p._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(indent_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(indent_chars * BODY_SIZE * 20)))
    return p


def set_border_bottom(p, color_hex, width_pt=0.5, space=0):
    pPr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(int(width_pt * 8)))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color_hex)
    borders.append(bottom)
    pPr.append(borders)


def add_rule(doc, recto=True):
    """标题下的横线：零高度细线，宽 RULE_W、黑色。奇数页右对齐、偶数页左对齐（镜像）。"""
    text_w_cm = (PAGE_W - M_LEFT - M_RIGHT) / 10.0   # 版心宽 cm
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(1)
    if recto:
        pf.left_indent = Cm(text_w_cm - RULE_W)    # 右对齐：左缩进限宽
    else:
        pf.right_indent = Cm(text_w_cm - RULE_W)   # 左对齐：右缩进限宽
    set_border_bottom(p, RULE_COLOR, RULE_PT)
    R(p, "", size=1)
    return p


def add_chapter_rule(doc):
    """章节名下面的居中下划线：宽 CHAPTER_RULE_W、左右等距缩进使其居中。"""
    text_w_cm = (PAGE_W - M_LEFT - M_RIGHT) / 10.0
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(1)
    indent = Cm((text_w_cm - CHAPTER_RULE_W) / 2)
    pf.left_indent = indent
    pf.right_indent = indent
    set_border_bottom(p, RULE_COLOR, CHAPTER_RULE_PT)
    R(p, "", size=1)
    return p


def deco_x(recto):
    """装饰矩形的水平坐标（镜像、贴外侧）：奇数页(recto=右页)外侧在右→色块在右；偶数页(verso=左页)外侧在左→色块在左。"""
    page_w_pt = PAGE_W * 72 / 25.4
    if recto:
        return page_w_pt - DECO_X - DECO_W   # 右侧：右边缘距纸边 DECO_X
    return DECO_X                             # 左侧：左边缘距纸边 DECO_X


def add_deco_rect(p, x_pt, y_pt, w_pt, h_pt, fill_hex):
    """在页面绝对位置画一个填充矩形（锚定在 p 段落，相对页面定位，置于文字下方）"""
    offset = BLEED_MM * 72 / 25.4   # 出血时坐标整体偏移一个出血量，保持和文字的相对位置
    x_pt += offset
    y_pt += offset
    vml = ('<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" '
           'xmlns:o="urn:schemas-microsoft-com:office:office">'
           '<v:rect id="deco" o:allowincell="f" '
           f'style="position:absolute;margin-left:{x_pt:.1f}pt;margin-top:{y_pt:.1f}pt;'
           f'width:{w_pt}pt;height:{h_pt}pt;'
           'z-index:-251658240;'
           'mso-position-horizontal-relative:page;mso-position-vertical-relative:page" '
           f'fillcolor="#{fill_hex}" stroked="f"/>'
           '</w:pict>')
    run = p.add_run()
    run._r.append(parse_xml(vml))


def add_page_background(sec, color_hex):
    """给这一节的所有页面加整页底色：在页眉段落里锚定 VML 矩形，相对页面定位、置于文字下方。
    奇数页、偶数页页眉各放一份（开奇偶页眉后篇章页可能落在偶数页）。"""
    w_pt = (PAGE_W + 2 * BLEED_MM) * 72 / 25.4 + 16
    h_pt = (PAGE_H + 2 * BLEED_MM) * 72 / 25.4 + 16
    for hdr in (sec.header, sec.even_page_header):
        p = hdr.paragraphs[0]
        vml = (
            '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office">'
            '<v:rect id="pagebg" o:allowincell="f" '
            f'style="position:absolute;margin-left:-8pt;margin-top:-8pt;'
            f'width:{w_pt:.1f}pt;height:{h_pt:.1f}pt;'
            'z-index:-251658240;'
            'mso-position-horizontal-relative:page;mso-position-vertical-relative:page" '
            f'fillcolor="#{color_hex}" stroked="f"/>'
            '</w:pict>'
        )
        run = p.add_run()
        run._r.append(parse_xml(vml))


def _crop_rect(p, x_pt, y_pt, w_pt, h_pt, color_hex, idx):
    """在页面绝对位置画一个细长矩形（用作角线的一段）。"""
    vml = ('<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" '
           'xmlns:o="urn:schemas-microsoft-com:office:office">'
           f'<v:rect id="crop{idx}" o:allowincell="f" '
           f'style="position:absolute;margin-left:{x_pt:.2f}pt;margin-top:{y_pt:.2f}pt;'
           f'width:{w_pt:.2f}pt;height:{h_pt:.2f}pt;'
           'z-index:251658240;'
           'mso-position-horizontal-relative:page;mso-position-vertical-relative:page" '
           f'fillcolor="#{color_hex}" stroked="f"/>'
           '</w:pict>')
    run = p.add_run()
    run._r.append(parse_xml(vml))


def add_crop_marks(sec):
    """每页四角加裁切角线（供印刷裁切定位）：每条线=细黑线压在粗白线(挖空)上。
    坐标按参考书《秀茹的一生》角线实测：距成品边 10/13mm、黑 0.1mm 白 0.5mm。"""
    if not CROP_MARKS:
        return
    W, H = PAGE_W, PAGE_H
    lines = [
        (10, 13, 0, 13), (13, 10, 13, 0), (13, 10, 3, 10), (10, 13, 10, 3),                      # 左上
        (W - 10, 13, W, 13), (W - 13, 10, W - 13, 0), (W - 13, 10, W - 3, 10), (W - 10, 13, W - 10, 3),  # 右上
        (10, H - 13, 0, H - 13), (13, H - 10, 13, H), (13, H - 10, 3, H - 10), (10, H - 13, 10, H - 3),  # 左下
        (W - 10, H - 13, W, H - 13), (W - 13, H - 10, W - 13, H), (W - 13, H - 10, W - 3, H - 10), (W - 10, H - 13, W - 10, H - 3),  # 右下
    ]
    M = 72 / 25.4   # 毫米 → 磅
    idx = 0
    for hdr in (sec.header, sec.even_page_header):
        p = hdr.paragraphs[0]
        for (x1, y1, x2, y2) in lines:
            x1 += BLEED_MM; y1 += BLEED_MM; x2 += BLEED_MM; y2 += BLEED_MM
            if abs(y2 - y1) < 1e-6:   # 横线
                x0, y0, L = min(x1, x2), y1, abs(x2 - x1)
                idx += 1; _crop_rect(p, x0 * M, (y0 - CROP_WHITE_W / 2) * M, L * M, CROP_WHITE_W * M, "FFFFFF", idx)
                idx += 1; _crop_rect(p, x0 * M, (y0 - CROP_BLACK_W / 2) * M, L * M, CROP_BLACK_W * M, "000000", idx)
            else:                      # 竖线
                x0, y0, L = x1, min(y1, y2), abs(y2 - y1)
                idx += 1; _crop_rect(p, (x0 - CROP_WHITE_W / 2) * M, y0 * M, CROP_WHITE_W * M, L * M, "FFFFFF", idx)
                idx += 1; _crop_rect(p, (x0 - CROP_BLACK_W / 2) * M, y0 * M, CROP_BLACK_W * M, L * M, "000000", idx)


def add_textbox(doc, text, width_cm=INTRO_BOX_W, size=INTRO_SIZE, line=1.0):
    """插入一个居中的无边框文本框，框内文字左对齐（用于章节页引言等）。"""
    if not text:
        return
    if width_cm is None:
        width_cm = (PAGE_W - M_LEFT - M_RIGHT) / 10.0
    w_emu = int(width_cm * 360000)          # 1 cm = 360000 EMU
    # 每行能放几个汉字：按 1.4 倍字号宽保守估算（含框内边距），避免框高不够把文字截断
    cpl = max(1, int(width_cm / (size * 0.05)))
    lines = max(1, math.ceil(len(text) / cpl))
    # 高度：每行按 1.8 倍字号高，再多留 1 行余量（框无边框，多余高度不可见）
    h_emu = int((lines + 1) * size * 1.8 * line * 12700)   # 1 pt = 12700 EMU
    h_emu = max(h_emu, int(2.0 * 360000))
    sz = str(int(size * 2))
    p = doc.add_paragraph()
    set_para(p, CENTER, 0, 0, 1.0)
    r = p.add_run()
    xml = (
        '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{w_emu}" cy="{h_emu}"/>'
        '<wp:docPr id="1" name="文本框"/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w_emu}" cy="{h_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
        '</wps:spPr><wps:txbx><w:txbxContent>'
        f'<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="240" w:lineRule="atLeast"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{F_BODY}"/>'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'<w:t xml:space="preserve">{html.escape(text)}</w:t></w:r></w:p>'
        '</w:txbxContent></wps:txbx>'
        '<wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>'
        '</wps:wsp></a:graphicData></a:graphic>'
        '</wp:inline></w:drawing>'
    )
    r._r.append(parse_xml(xml))


def _img_block(p, w_pt, h_pt, left_pt, fill_hex):
    """在图片段落里锚定一个色块矩形：衬在图片底下、随段落流动（相对版心/段落定位）。"""
    blk_w = w_pt * IMG_BLOCK_SIZE
    blk_h = h_pt * IMG_BLOCK_SIZE
    blk_left = left_pt - IMG_BLOCK_EXPOSE
    blk_top = h_pt * (1 - IMG_BLOCK_SIZE) + IMG_BLOCK_EXPOSE
    vml = ('<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" '
           'xmlns:o="urn:schemas-microsoft-com:office:office">'
           '<v:rect id="imgblk" o:allowincell="f" '
           f'style="position:absolute;margin-left:{blk_left:.1f}pt;margin-top:{blk_top:.1f}pt;'
           f'width:{blk_w:.1f}pt;height:{blk_h:.1f}pt;'
           'z-index:-251658240;'
           'mso-position-horizontal-relative:margin;mso-position-vertical-relative:paragraph" '
           f'fillcolor="#{fill_hex}" stroked="f"/>'
           '</w:pict>')
    run = p.add_run()
    run._r.append(parse_xml(vml))


def _img_line(p, left_pt, y0, y1, x_end, color_hex):
    """图注引线：竖线（left_pt, y0→y1）+ 横线（left_pt→x_end, 高 y1），L 形折线连到图注。"""
    vml = ('<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" '
           'xmlns:o="urn:schemas-microsoft-com:office:office">'
           f'<v:line id="imgline" from="{left_pt:.1f}pt,{y0:.1f}pt" to="{left_pt:.1f}pt,{y1:.1f}pt" '
           f'style="position:absolute;z-index:251658240;'
           'mso-position-horizontal-relative:margin;mso-position-vertical-relative:paragraph" '
           f'strokecolor="#{color_hex}" strokeweight="0.75pt"/>'
           f'<v:line id="imgline2" from="{left_pt:.1f}pt,{y1:.1f}pt" to="{x_end:.1f}pt,{y1:.1f}pt" '
           f'style="position:absolute;z-index:251658240;'
           'mso-position-horizontal-relative:margin;mso-position-vertical-relative:paragraph" '
           f'strokecolor="#{color_hex}" strokeweight="0.75pt"/>'
           '</w:pict>')
    run = p.add_run()
    run._r.append(parse_xml(vml))


def image_layout_size(path, max_width_cm):
    """返回图片按 max_width_cm 缩放后的 (宽度cm, 宽度pt, 高度pt)。"""
    native_cm = None
    px_w = px_h = None
    try:
        from PIL import Image
        with Image.open(path) as im:
            px_w, px_h = im.size
            dpi = im.info.get("dpi", (96, 96)) or (96, 96)
            dpi_x = dpi[0] or 96
        native_cm = px_w / dpi_x * 2.54
    except Exception:
        native_cm = None
    w_cm = native_cm if (native_cm is not None and native_cm <= max_width_cm) else max_width_cm
    w_pt = w_cm * 28.3465
    h_pt = w_pt * (px_h / px_w) if (px_w and px_h) else w_pt * 0.75
    return w_cm, w_pt, h_pt


def image_total_height(path, caption, max_width_cm):
    """一张图（图 + 图注 + 段距）在正文里占的总高度（磅），用于判断一页能放几张。"""
    _, _, h_pt = image_layout_size(path, max_width_cm)
    extra = BODY_AFTER
    if caption:
        extra += CAPTION_SIZE * 1.2
        if IMG_LINE_ON:
            extra += IMG_LINE_GAP
    return h_pt + extra


def add_image(doc, path, caption=None, max_width_cm=None, page_break=False):
    """插图：图片居中；图片底下可衬色块（IMG_BLOCK_*）；图注可加折线引线（IMG_LINE_*）。"""
    if max_width_cm is None:
        base = IMAGE_MAX_W or (PAGE_W - M_LEFT - M_RIGHT) / 10.0
        max_width_cm = base * IMG_MAX_W_RATIO
    text_w_pt = (PAGE_W - M_LEFT - M_RIGHT) * 72 / 25.4
    w_cm, w_pt, h_pt = image_layout_size(path, max_width_cm)
    p = doc.add_paragraph()
    set_para(p, CENTER, 0, 0, 1.0)
    if page_break:
        p.paragraph_format.page_break_before = True
    r = p.add_run()
    r.add_picture(path, width=Cm(w_cm))
    left_pt = (text_w_pt - w_pt) / 2

    if IMG_BLOCK_ON:
        _img_block(p, w_pt, h_pt, left_pt,
                   str(IMG_BLOCK_COLOR) if IMG_BLOCK_COLOR else str(BOOK_COLOR))

    if caption:
        p.paragraph_format.keep_with_next = True   # 图片与图注同页，不分开
        if IMG_LINE_ON:
            cap_w = len(caption) * CAPTION_SIZE
            cap_left = (text_w_pt - cap_w) / 2
            line_end = max(cap_left - 4, left_pt + 12)
            cap_y = h_pt + IMG_LINE_GAP
            _img_line(p, left_pt, h_pt, cap_y, line_end,
                      str(IMG_LINE_COLOR) if IMG_LINE_COLOR else "000000")
            pc = doc.add_paragraph()
            set_para(pc, CENTER, 0, BODY_AFTER, 1.0)
            pc.paragraph_format.space_before = Pt(IMG_LINE_GAP - CAPTION_SIZE * 0.6)   # 图注上移，让文字中线和引线横线对齐
            R(pc, caption, cn=F_BODY, size=CAPTION_SIZE)
        else:
            pc = doc.add_paragraph()
            set_para(pc, CENTER, 0, BODY_AFTER, 1.0)
            R(pc, caption, cn=F_BODY, size=CAPTION_SIZE)
    else:
        p.paragraph_format.space_after = Pt(BODY_AFTER)
    return p


NOTES = ["（这中间是不是接得有点突兀）",
         "（这里感觉缺了什么，但是又不知道怎么去连接得更好）"]
END_PUNCT = "。！？!?；;：:…，,、～~”」』）)"
PREFACE_KEYS = ("自序", "前言", "序言", "代序", "写在前面")   # 单 # 开头 → 前言页（目录之前）
AFTERWORD_KEYS = ("后记", "跋")                              # 单 # 开头 → 后记页（正文之后）
IMG_RE = re.compile(r"^图片\s*[:：]\s*(.+)$")
CAP_RE = re.compile(r"^图片名\s*[:：]\s*(.+)$")
MD_IMG_RE = re.compile(r"^!\s*\[([^\]]*)\]\s*\(([^)]+)\)")   # ![图注](路径) 标准 Markdown 图片
SUB_RE = re.compile(r"^#####\s*(?:小标题\s*[:：]?\s*)?(.+)$")


def clean_line(s):
    s = re.sub(r"~~.*?~~", "", s)
    for note in NOTES:
        s = s.replace(note, "")
    return s.strip()   # 保留 **，供正文「加粗+主题色」拆分使用


def _resolve_img(path, base_dir):
    """相对路径的图片按 .md 文件所在目录解析成绝对路径；绝对路径原样返回。"""
    if not path or base_dir is None or os.path.isabs(path):
        return path
    return os.path.normpath(os.path.join(base_dir, path))


def classify(lines, base_dir=None, poem=False):
    """把一篇文章的正文行分类成 items 列表：
    ('p', text) 正文段  ('h', text) 小标题  ('num', text) 序号  ('img', (path, caption)) 插图
    poem=True 时改用诗歌规则：('h' 诗题) ('meta' 副题/日期) ('poem' 诗句) ('stanza' 诗节分隔)"""
    if poem:
        return _classify_poem(lines, base_dir)
    ls = [clean_line(s) for s in lines]
    items, m, k = [], len(ls), 0
    while k < m:
        s = ls[k]
        if not s:
            k += 1
            continue
        sm = SUB_RE.match(s)
        if sm:
            items.append(("h", sm.group(1).strip()))
            k += 1
            continue
        im = IMG_RE.match(s)
        if im:
            path = _resolve_img(im.group(1).strip(), base_dir)
            caption = None
            if k + 1 < m:
                cm = CAP_RE.match(ls[k + 1])
                if cm:
                    caption = cm.group(1).strip()
                    k += 1
            items.append(("img", (path, caption)))
            k += 1
            continue
        mdim = MD_IMG_RE.match(s)
        if mdim:
            path = _resolve_img(mdim.group(2).strip(), base_dir)
            caption = mdim.group(1).strip() or None
            items.append(("img", (path, caption)))
            k += 1
            continue
        pb = (k == 0 or not ls[k - 1])
        nb = (k == m - 1 or not ls[k + 1])
        if re.fullmatch(r"\d{1,2}", s):
            items.append(("num", s))
        elif pb and nb and 2 <= len(s) <= 16 and s[-1] not in END_PUNCT and s[0] not in "“\"‘「（":
            items.append(("h", s))
        else:
            items.append(("p", s))
        k += 1
    return items


def _classify_poem(lines, base_dir=None):
    """诗歌分类：##### → 诗题(h)；「——xxx」副题、「xxxx年xx月」日期 → (meta)；
    普通行 → 诗句(poem)；空行 → 诗节分隔(stanza)。"""
    ls = [clean_line(s) for s in lines]
    items, m, k = [], len(ls), 0
    while k < m:
        s = ls[k]
        if not s:
            k += 1
            while k < m and not ls[k]:
                k += 1
            # 空行只在「前面是诗句、后面还有非标题内容」时算诗节分隔
            if items and k < m and items[-1][0] == "poem" and not SUB_RE.match(ls[k]):
                items.append(("stanza", None))
            continue
        sm = SUB_RE.match(s)
        if sm:
            items.append(("h", sm.group(1).strip()))
            k += 1
            continue
        im = IMG_RE.match(s)
        if im:
            path = _resolve_img(im.group(1).strip(), base_dir)
            caption = None
            if k + 1 < m:
                cm = CAP_RE.match(ls[k + 1])
                if cm:
                    caption = cm.group(1).strip()
                    k += 1
            items.append(("img", (path, caption)))
            k += 1
            continue
        mdim = MD_IMG_RE.match(s)
        if mdim:
            path = _resolve_img(mdim.group(2).strip(), base_dir)
            caption = mdim.group(1).strip() or None
            items.append(("img", (path, caption)))
            k += 1
            continue
        if re.fullmatch(r"\d{4}\s*年\s*\d{1,2}\s*月(\s*\d{1,2}\s*[日号])?", s):
            items.append(("meta", s))
        elif s.startswith("——") or s.startswith("--"):
            items.append(("meta", "——" + s.lstrip("-— ")))
        else:
            items.append(("poem", s))
        k += 1
    return items


def parse_book(src=None):
    """读取源文件，解析出前言/后记/章节/文章。
    返回 dict：{"preface": str|None, "afterword": str|None,
                "chapters": [{"title": str|None, "intro": str|None, "articles": [...]}]}"""
    src = src or SRC
    src_dir = os.path.dirname(os.path.abspath(src))
    raw = open(src, encoding="utf-8").read().split("\n")
    n = len(raw)
    preface, afterword = [], []
    preface_title = afterword_title = None
    chapters = []
    i = 0
    while i < n:
        s = raw[i].strip()
        if not s:
            i += 1
            continue
        # 前言/自序/序言/代序/写在前面 与 后记/跋（单 # 开头，且不是 ## 或 ### 或 ####）
        if s.startswith("# ") and not s.startswith("##"):
            kind = s.lstrip("#").strip()
            if kind in PREFACE_KEYS:
                target, preface_title = preface, kind
            elif kind in AFTERWORD_KEYS:
                target, afterword_title = afterword, kind
            else:
                target = None
            i += 1
            if target is None:
                continue
            while i < n and not raw[i].strip().startswith("#"):
                if raw[i].strip():
                    target.append(raw[i].strip())
                i += 1
            continue
        # 章节（### 且不是 ####）
        if s.startswith("###") and not s.startswith("####"):
            title = s.lstrip("#").strip()
            intro = None
            j = i + 1
            while j < n and not raw[j].strip():
                j += 1
            if j < n and raw[j].strip().startswith("简介"):
                intro = re.sub(r"^简介\s*[:：]?\s*", "", raw[j].strip()).strip()
                j += 1
            chapters.append({"title": title, "intro": intro, "articles": []})
            i = j
            continue
        # 文章（####）
        if s.startswith("####"):
            title = re.sub(r"^标题\s*[:：]\s*", "", s.lstrip("#").strip()).strip()
            author, j = "", i + 1
            while j < n and j <= i + 5:
                l = raw[j].strip().strip("*").strip()
                if not l:
                    j += 1
                    continue
                if l.startswith("作者"):
                    author = re.sub(r"^作者\s*[:：]?\s*", "", l).strip()
                    j += 1
                break
            body = []
            while j < n:
                sj = raw[j].strip()
                # 小标题(#####)属于文章内容；遇到 # / ## / ### / #### 才算文章结束
                if sj.startswith("#") and not sj.startswith("#####"):
                    break
                body.append(sj)
                j += 1
            # 「诗歌：」标记：整篇按诗歌排版（每首诗一页、诗句居中）
            poem = False
            body2 = []
            for ln in body:
                if re.fullmatch(r"诗歌\s*[:：]?\s*", ln):
                    poem = True
                else:
                    body2.append(ln)
            art = {"title": title, "author": author, "poem": poem,
                   "items": classify(body2, base_dir=src_dir, poem=poem)}
            if chapters:
                chapters[-1]["articles"].append(art)
            else:
                chapters.append({"title": None, "intro": None, "articles": [art]})
            i = j
            continue
        i += 1
    return {
        "preface": "\n".join(preface) if preface else None,
        "preface_title": preface_title,
        "afterword": "\n".join(afterword) if afterword else None,
        "afterword_title": afterword_title,
        "chapters": chapters,
    }


# ---------------- 组装文档 ----------------
def new_doc():
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(BODY_SIZE)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), F_BODY)
    settings = doc.settings.element
    mm = OxmlElement("w:mirrorMargins")                    # 镜像页边距：左/右边距变内/外侧
    settings.insert_element_before(mm, "w:proofState", "w:defaultTabStop", "w:evenAndOddHeaders", "w:updateFields")
    doc.settings.odd_and_even_pages_header_footer = True   # 奇偶页用不同页眉
    uf = OxmlElement("w:updateFields"); uf.set(qn("w:val"), "true"); settings.append(uf)
    sec = doc.sections[0]
    # 页面 = 成品尺寸 + 四周出血；边距也同步外扩一个出血量，正文相对裁切线位置不变
    sec.page_width = Mm(PAGE_W + 2 * BLEED_MM); sec.page_height = Mm(PAGE_H + 2 * BLEED_MM)
    sec.top_margin = Mm(M_TOP + BLEED_MM); sec.bottom_margin = Mm(M_BOT + BLEED_MM)
    sec.left_margin = Mm(M_LEFT + BLEED_MM); sec.right_margin = Mm(M_RIGHT + BLEED_MM)
    sec.header_distance = Mm(HEADER_DIST + BLEED_MM)
    sec.footer_distance = Mm(15 + BLEED_MM)
    for g in sec._sectPr.findall(qn("w:docGrid")):
        sec._sectPr.remove(g)
    return doc, sec


def setup_header(sec, chapter=None, pad=None):
    """页眉：奇数页=章节名+页码（靠右，右侧是奇数页的外侧）；偶数页=书名+页码（靠左，左侧是偶数页的外侧）。
    chapter：章节标题（如「第二章 大家眼中的你」），奇数页第一行显示它；无章节时退回书名。"""
    # 奇数页页眉（默认页眉）：放在页面右侧，显示章节名
    odd = sec.header
    odd.is_linked_to_previous = False
    hp1 = odd.paragraphs[0]
    set_para(hp1, RIGHT, 0, 0, 1.0)
    R(hp1, chapter if chapter else BOOK_NAME, cn=F_HEAD, size=HEAD_SIZE)
    hp2 = odd.add_paragraph()
    set_para(hp2, RIGHT, 0, 0, 1.0)
    add_page_field(hp2, NUM_SIZE, F_NUM, underline=True, pad=pad)
    # 偶数页页眉：放在页面左侧，显示书名
    even = sec.even_page_header
    even.is_linked_to_previous = False
    ep1 = even.paragraphs[0]
    set_para(ep1, LEFT, 0, 0, 1.0)
    R(ep1, BOOK_NAME, cn=F_HEAD, size=HEAD_SIZE)
    ep2 = even.add_paragraph()
    set_para(ep2, LEFT, 0, 0, 1.0)
    add_page_field(ep2, NUM_SIZE, F_NUM, underline=True, pad=pad)


def clear_header(sec):
    """清空这一节的页眉（奇数页+偶数页都置空），用于暗码页（不显示页码）。"""
    sec.header.is_linked_to_previous = False
    sec.even_page_header.is_linked_to_previous = False


def build_article(doc, art, first, recto=True):
    align = RIGHT if recto else LEFT   # 标题块镜像：奇数页靠右、偶数页靠左（贴外侧）
    for k in range(TITLE_TOP_LINES):
        sp = doc.add_paragraph()
        set_para(sp, LEFT, 0, 0, BODY_LINE)
        if not first and k == 0:
            sp.paragraph_format.page_break_before = True
        R(sp, "", size=BODY_SIZE)
    pt = doc.add_paragraph()
    set_para(pt, align, 0, 0, 1.0)
    R(pt, art["title"], cn=F_TITLE, size=TITLE_SIZE)

    add_deco_rect(pt, deco_x(recto), DECO_TOP, DECO_W, DECO_H, str(BOOK_COLOR))

    # 有作者才画标题下划线和作者名；没有作者就只留文章标题
    if art["author"]:
        add_rule(doc, recto)
        pa = doc.add_paragraph()
        set_para(pa, align, 0, 0, 1.0)
        R(pa, art["author"], cn=F_AUTHOR, size=AUTHOR_SIZE, color=BOOK_COLOR)

    # 标题（含横线、作者名）到正文之间空行
    for _k in range(TITLE_BODY_GAP_LINES):
        gp = doc.add_paragraph()
        set_para(gp, LEFT, 0, 0, BODY_LINE)
        R(gp, "", size=BODY_SIZE)

    if art.get("poem"):
        _render_poems(doc, art["items"])
        return

    page_h_pt = (PAGE_H - M_TOP - M_BOT) * 72 / 25.4
    img_max_w = (IMAGE_MAX_W or (PAGE_W - M_LEFT - M_RIGHT) / 10.0) * IMG_MAX_W_RATIO
    group_h = 0.0
    for kind, text in art["items"]:
        if kind == "img":
            path, caption = text
            h = image_total_height(path, caption, img_max_w)
            if group_h > 0 and group_h + h > page_h_pt:
                add_image(doc, path, caption, img_max_w, page_break=True)
                group_h = h
            else:
                add_image(doc, path, caption, img_max_w)
                group_h += h
            continue
        group_h = 0.0
        p = doc.add_paragraph()
        if kind == "h":
            set_para(p, LEFT, 0, BODY_AFTER, BODY_LINE)
            p.paragraph_format.keep_with_next = True   # 小标题与下一段正文同页，不单独落在页尾
            R(p, text, cn=F_BODY, size=BODY_SIZE, bold=SUBTITLE_BOLD)
        elif kind == "p":
            set_para(p, JUSTIFY, 0, BODY_AFTER, BODY_LINE, indent_chars=BODY_INDENT)
            add_body_para(p, text)
        else:  # num
            set_para(p, LEFT, 0, BODY_AFTER, BODY_LINE)
            R(p, text, cn=F_BODY, size=BODY_SIZE)


def _render_poems(doc, items):
    """渲染诗歌：每个诗题(h)从第二首起另起一页；诗句居中、行距放宽；空行=诗节留白。"""
    first_poem = True
    stanza = False
    for kind, text in items:
        if kind == "h":
            if not first_poem:
                for k in range(POEM_TOP_LINES):
                    sp = doc.add_paragraph()
                    set_para(sp, LEFT, 0, 0, BODY_LINE)
                    if k == 0:
                        sp.paragraph_format.page_break_before = True
                    R(sp, "", size=BODY_SIZE)
            p = doc.add_paragraph()
            set_para(p, CENTER, 0, POEM_TITLE_AFTER, 1.0)
            R(p, text, cn=F_TITLE, size=POEM_TITLE_SIZE)
            first_poem = False
            stanza = False
        elif kind == "meta":
            p = doc.add_paragraph()
            set_para(p, CENTER, 0, POEM_META_AFTER, 1.0)
            R(p, text, cn=F_BODY, size=POEM_META_SIZE)
            stanza = False
        elif kind == "poem":
            p = doc.add_paragraph()
            set_para(p, CENTER, 0, 0, POEM_LINE)
            if stanza:
                p.paragraph_format.space_before = Pt(POEM_STANZA_GAP)
            R(p, text, cn=F_BODY, size=POEM_LINE_SIZE)
            stanza = False
        elif kind == "stanza":
            stanza = True
        elif kind == "img":
            path, caption = text
            add_image(doc, path, caption)
            stanza = False


def add_chapter_page(doc, title, intro=None):
    """篇章页：章节名居中 + 居中下划线 +（可选）下方居中文本框。占一整页。
    换页由「节」的分页负责，这里不再加 page_break。"""
    for k in range(CHAPTER_TOP_LINES):
        sp = doc.add_paragraph()
        set_para(sp, LEFT, 0, 0, BODY_LINE)
        R(sp, "", size=BODY_SIZE)
    p = doc.add_paragraph()
    set_para(p, CENTER, 0, 0, 1.0)
    R(p, title, cn=F_TITLE, size=CHAPTER_SIZE)
    add_chapter_rule(doc)
    if intro:
        sp = doc.add_paragraph()
        set_para(sp, LEFT, 0, 0, BODY_LINE)
        R(sp, "", size=BODY_SIZE)
        add_textbox(doc, intro)


def add_front_page(doc, title, text):
    """前言/后记页：大标题居中（FRONT_SIZE 加粗），下面是正文段落。"""
    for k in range(FRONT_TOP_LINES):
        sp = doc.add_paragraph()
        set_para(sp, LEFT, 0, 0, BODY_LINE)
        R(sp, "", size=BODY_SIZE)
    p = doc.add_paragraph()
    set_para(p, CENTER, 0, 18, 1.0)
    R(p, title, cn=F_TITLE, size=FRONT_SIZE, bold=True)
    for para in (text or "").split("\n"):
        if not para.strip():
            continue
        bp = doc.add_paragraph()
        set_para(bp, JUSTIFY, 0, BODY_AFTER, BODY_LINE, indent_chars=BODY_INDENT)
        R(bp, para, cn=F_BODY, size=BODY_SIZE)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    doc, sec = new_doc()
    setup_header(sec)
    add_crop_marks(sec)
    data = parse_book()
    articles = [a for ch in data["chapters"] for a in ch["articles"]]
    arts = articles[:LIMIT] if LIMIT else articles
    for idx, a in enumerate(arts):
        build_article(doc, a, idx == 0)
    out = os.path.join(OUT_DIR, f"{NAME}.docx")
    doc.save(out)
    print(f"已生成 {out}")
    print(f"排入 {len(arts)} 篇（共 {len(articles)} 篇）")
    for idx, a in enumerate(arts):
        print(f"  {idx+1}. {a['title']} / {a['author'] or '（无作者）'}  正文 {len(a['items'])} 段")


if __name__ == "__main__":
    main()
