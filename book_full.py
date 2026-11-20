# -*- coding: utf-8 -*-
"""把「前言 + 目录 + 正文 + 后记」合成一本书（一个 Word 文档）。

用法：python book_full.py ["要合成的书.md" ["书名"]]
不带参数时用下方 SRC/BOOK_NAME 默认值；输出的 docx 自动生成在源文件同目录、同名 .docx。

源文件约定写法见 book_layout.py 顶部注释：
  # 自序 / # 前言 / # 序言 / # 代序 / # 写在前面   —— 前言页（排在目录之前）
  # 后记 / # 跋         —— 后记页（排在正文之后）
  ### 章节 + 简介：      —— 章节页（章节名 + 下划线 + 文本框引言）
  #### 标题：+ 作者：    —— 文章（无作者则省略标题下划线和作者名）
  图片：+ 图片名：       —— 插图 + 图注

成书顺序：扉页 → 前言页 → 目录页 → 正文 → 后记页（前言/目录不显示页码，正文从第 1 页开始；
目录里「前言/自序」只列标题、不写页码）。
目录页码自动计算：完整排版后读取书签的实际页码，复核目录与标题左右位置，分页稳定后才导出。
改样式：分别去改 book_layout.py 和 book_toc_horizontal.py 顶部参数区。
"""
import io, os, sys, tempfile
if (getattr(sys.stdout, "encoding", "") or "").lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

import book_layout as book
import book_toc_horizontal as toc

# ==================== 输入与输出（命令行传参，带默认值） ====================
SRC = sys.argv[1] if len(sys.argv) > 1 else r"D:\纸质文集\你总是有选择的.md"   # 正文源文件
BOOK_NAME = sys.argv[2] if len(sys.argv) > 2 else "你总是有选择的"              # 页眉上显示的书名
OUT = os.path.join(os.path.dirname(os.path.abspath(SRC)),
                   os.path.splitext(os.path.basename(SRC))[0] + ".docx")        # 输出到源文件同目录、同名
TITLE_PAGE = True                         # 在序言/目录前增加独立扉页
TITLE_PAGE_FROM_SOURCE = True             # 使用开篇一级标题及正文，或显式“扉页：标题”
TITLE_PAGE_TEXT = ""                      # 自定义扉页正文，留空时按开关读取书稿
TITLE_PAGE_TITLE = ""                     # 留空使用书名
TITLE_PAGE_AUTHOR = ""                    # 扉页作者，可留空
TITLE_PAGE_SIZE = 24                      # 扉页书名字号（磅）
TOC_TITLE = "目录"                          # 目录页大标题

# ---- 目录里章节/文章行的样式 ----
CHAPTER_TOC_SIZE = 10.5        # 目录里章节行的字号（磅）
CHAPTER_TOC_BOLD = True      # 目录里章节行是否加粗
ARTICLE_INDENT_PT = 24       # 目录里文章行相对章节行缩进多少（磅）
# ==========================================================


def add_toc_line(doc, text, page, size, bold=False, indent_pt=0, pad=None, keep_next=False):
    """目录里加一行：标题左对齐（可缩进/加粗），页码右对齐（按 pad 补零）。
    右侧为页码预留一段宽度：标题过长时在该宽度内折行，保证标题与页码之间至少空一格；
    keep_next=True 时本行与下一行同页，避免目录末页只剩一行（寡行）。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = toc.ENTRY_SPACING
    pf.widow_control = True          # 段内不跨页留孤行
    if keep_next:
        pf.keep_with_next = True     # 与下一行绑定同页，防目录末页只有一行
    num = "" if not page else (str(page).zfill(pad) if pad else str(page))
    text_w_mm = book.PAGE_W - book.M_LEFT - book.M_RIGHT
    digits = len(num) if num else (pad or 1)
    pf.right_indent = Mm(digits * 2.4 + 3)   # 为页码预留的右侧宽度（含一个空格）
    if indent_pt:
        pf.left_indent = Pt(indent_pt)
    pf.tab_stops.add_tab_stop(Mm(text_w_mm), WD_TAB_ALIGNMENT.RIGHT, toc.ENTRY_LEADER)
    toc.R(p, text, size, bold=bold)
    p.add_run("\t")
    toc.R(p, num, size)


def build_toc(doc, chapters, page_map, pad=None, preface_title=None, afterword=None, appendices=None, endmatter_order=None, endmatter_articles=None):
    """在文档里排目录页（横排）。前言/序言在目录之前独立展示，不加入目录。"""
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before = Pt(toc.TOP_LINES * toc.ENTRY_SIZE)
    tp.paragraph_format.space_after = Pt(toc.TITLE_AFTER)
    tp.paragraph_format.line_spacing = 1.0
    toc.R(tp, TOC_TITLE, toc.TITLE_SIZE, bold=True)

    entries = []
    last_volume = None
    for ch in chapters:
        volume = ch.get("volume")
        if volume and volume is not last_volume:
            entries.append((volume["label"], "", CHAPTER_TOC_SIZE + 2, True, 0))
        last_volume = volume
        if ch["title"] is not None:
            entries.append((ch["title"], page_map.get(ch["title"], ""),
                            CHAPTER_TOC_SIZE, CHAPTER_TOC_BOLD, 0))
        for a in ch["articles"]:
            entries.append((a["title"], page_map.get(a["title"], ""),
                            toc.ENTRY_SIZE, False, ARTICLE_INDENT_PT))
    for afterword in (endmatter_order if endmatter_order is not None else ([afterword] if afterword else []) + (appendices or [])):
        entries.append((afterword[0], page_map.get(afterword[1], ''), CHAPTER_TOC_SIZE, CHAPTER_TOC_BOLD, 0))
        for article in (endmatter_articles or {}).get(afterword[1], []):
            entries.append((article['title'],page_map.get(article['toc_key'],''),toc.ENTRY_SIZE,False,ARTICLE_INDENT_PT))
    last = len(entries) - 1
    for i, (text, page, size, bold, indent_pt) in enumerate(entries):
        # 只把「倒数第二行」与最后一行绑定同页：防止目录末页只剩一行（寡行）。
        # 不能给所有行都加 keep_with_next，否则 Word 会把整块目录推到下一页，首页反而只剩标题。
        add_toc_line(doc, text, page, size, bold=bold, indent_pt=indent_pt,
                     pad=pad, keep_next=(i == last - 1))


def build_body(doc, chapters, first_sec, pad=None, page_map=None, anchors=None):
    """在 doc 里排正文（第一个正文节用 first_sec，正文页码从 1 开始）。
    篇章页单独成节、置空页眉（暗码：不显示页码）；章内文章合为一节、带奇偶页眉。
    page_map：每篇标题 → 正文起始页（用于判断该页是奇数页/偶数页，好让标题色块镜像）。
    anchors：传入一个空列表时，给卷、章节和文章挂非打印书签，用于读取完整成书的最终页码。"""
    set_page_number_start(first_sec._sectPr, 1)   # 只有正文第一节重置页码，其余节连续编号

    def _tok(title):
        # 生成唯一锚点 token 并登记；anchors 为 None（正式文档）时返回 None，不挂锚点
        if anchors is None:
            return None
        tok = f"zq{len(anchors):04d}wz"
        anchors.append((tok, title))
        return tok

    def _recto(title):
        # 正文页码 1 起：奇数页=右页(recto)→色块在右/标题右对齐；偶数页=左页(verso)→色块在左/标题左对齐
        if not page_map:
            return True
        pg = page_map.get(title, 1)
        return pg % 2 == 1

    last_volume = None
    for ci, ch in enumerate(chapters):
        volume = ch.get("volume")
        starts_volume = volume and volume is not last_volume
        if starts_volume:
            volume_sec = first_sec if ci == 0 else book.add_book_section(doc, WD_SECTION.ODD_PAGE)
            if ci != 0:
                remove_page_number_start(volume_sec._sectPr)
            book.clear_header(volume_sec)
            book.add_volume_background(volume_sec)
            book.add_volume_page(doc, volume, anchor=_tok(volume["label"]))
        last_volume = volume
        if ch["title"] is not None:
            if ci == 0 and not starts_volume:
                page_sec = first_sec
            else:
                page_sec = book.add_book_section(doc, WD_SECTION.NEW_PAGE if starts_volume else WD_SECTION.ODD_PAGE)   # 首章紧接卷页，其余章节右页开始
                remove_page_number_start(page_sec._sectPr)
            book.clear_header(page_sec)                      # 暗码：篇章页不显示页码
            book.add_page_background(page_sec, str(book.BOOK_COLOR))   # 篇章页整页底色（用书籍主色调）
            book.add_chapter_page(doc, ch["title"], ch.get("intro"), anchor=_tok(ch["title"]))
            art_sec = book.add_book_section(doc, WD_SECTION.NEW_PAGE)
            remove_page_number_start(art_sec._sectPr)
            book.setup_header(art_sec, chapter=ch["title"], pad=pad)   # 文章节奇偶页眉
            for i, a in enumerate(ch["articles"]):
                book.build_article(doc, a, i == 0, recto=_recto(a["title"]), anchor=_tok(a["title"]))
        else:
            art_sec = book.add_book_section(doc, WD_SECTION.NEW_PAGE) if starts_volume else doc.sections[-1]
            if starts_volume:
                remove_page_number_start(art_sec._sectPr)
            book.setup_header(art_sec, pad=pad)
            for i, a in enumerate(ch["articles"]):
                book.build_article(doc, a, i == 0, recto=_recto(a["title"]), anchor=_tok(a["title"]))


def add_front_matter(doc, title, text, anchor=None):
    """前言/后记页：单独成节、空页眉（暗码，不显示页码）。"""
    sec = book.add_book_section(doc, WD_SECTION.NEW_PAGE)
    remove_page_number_start(sec._sectPr)
    book.clear_header(sec)
    book.add_front_page(doc, title, text, anchor=anchor)
    return sec


def remove_page_number_start(sectPr):
    """移除节里的页码重置（w:pgNumType），让它接着上一节连续编号。"""
    for pg in list(sectPr.findall(qn("w:pgNumType"))):
        sectPr.remove(pg)


def set_page_number_start(sectPr, start=1):
    """让这一节的页码从 start 开始编号（放在 w:cols 之前，符合 Word 元素顺序）。"""
    remove_page_number_start(sectPr)
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), str(start))
    cols = sectPr.find(qn("w:cols"))
    if cols is not None:
        cols.addprevious(pg)
    else:
        sectPr.append(pg)


def build_document(data, page_map, pad):
    """Build the exact final book, including front matter and print decorations."""
    chapters = data["chapters"]
    anchors = []
    doc, sec1 = book.new_doc()      # 第一节 = 前言页（无前言时就是目录页）

    if TITLE_PAGE:
        book.clear_header(sec1)
        remove_page_number_start(sec1._sectPr)
        book.add_title_page(doc,
                            TITLE_PAGE_TITLE or (data.get("title_page_title") if TITLE_PAGE_FROM_SOURCE else None) or BOOK_NAME,
                            TITLE_PAGE_AUTHOR, TITLE_PAGE_SIZE,
                            TITLE_PAGE_TEXT or ((data.get("title_page") or "") if TITLE_PAGE_FROM_SOURCE else ""))
        sec1 = book.add_book_section(doc, WD_SECTION.NEW_PAGE)
        remove_page_number_start(sec1._sectPr)
        book.clear_header(sec1)

    # 序言接在扉页之后，目录之前，不显示页码。
    if data["preface"] is not None:
        book.clear_header(sec1)
        book.add_front_page(doc, data["preface_title"], data["preface"])
        toc_sec = book.add_book_section(doc, WD_SECTION.NEW_PAGE)   # 目录页另起一页
        remove_page_number_start(toc_sec._sectPr)
        book.clear_header(toc_sec)
    else:
        toc_sec = sec1

    endmatter = book.endmatter_entries(data)
    build_toc(doc, chapters, page_map, pad=pad, endmatter_order=[(item['title'], item['key']) for item in endmatter],
              endmatter_articles={item['key']:item.get('articles',[]) for item in endmatter})

    sec2 = book.add_book_section(doc, WD_SECTION.ODD_PAGE)   # 正文第一节：从奇数页(右页)开始，保证正文页码奇偶=物理页奇偶，色块/标题镜像才和页眉一致
    build_body(doc, chapters, sec2, pad=pad, page_map=page_map, anchors=anchors)

    for item in endmatter:
        token = 'zq{:04d}wz'.format(len(anchors))
        anchors.append((token, item['key']))
        articles = item.get('articles', [])
        if not articles:
            add_front_matter(doc, item['title'], item['text'], anchor=token)
            continue
        section = book.add_book_section(doc, WD_SECTION.NEW_PAGE)
        remove_page_number_start(section._sectPr)
        book.setup_header(section, chapter=item['title'], pad=pad)
        if item.get('intro'):
            book.add_front_page(doc,item['title'],item['intro'],anchor=token)
        else:
            book.add_anchor(doc.add_paragraph(),token)
        for index, article in enumerate(articles):
            article_token='zq{:04d}wz'.format(len(anchors))
            anchors.append((article_token,article['toc_key']))
            book.build_article(doc,article,index==0 and not item.get('intro'),
                               recto=page_map.get(article['toc_key'],1)%2==1,anchor=article_token)

    for sec in doc.sections:
        book.add_print_margin(sec)

    if book.CROP_MARKS:                # 每页四角加裁切角线
        for sec in doc.sections:
            book.add_crop_marks(sec)

    return doc, anchors


def main():
    from word_pagination import save_paginated
    book.SRC = SRC
    book.BOOK_NAME = BOOK_NAME
    data = book.parse_book()
    chapters = data["chapters"]
    with book.image_source_cache():
        page_map, total_pages, pad, engine = save_paginated(
            lambda pages, digits: build_document(data, pages, digits), OUT)
    n_ch = sum(1 for c in chapters if c["title"] is not None)
    n_art = sum(len(c["articles"]) for c in chapters)
    print(f"已生成 {OUT}")
    print(f"章节 {n_ch} 个；文章 {n_art} 篇；前言={'有' if data['preface'] else '无'}；"
          f"后记={'有' if data['afterword'] else '无'}；正文共 {total_pages} 页；"
          f"页码补零 {pad} 位；目录页码已自动计算")


if __name__ == "__main__":
    main()
