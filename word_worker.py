"""Isolated per-export adapter for the existing Word typesetter."""
import json, sys
from pathlib import Path

def main():
    config = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    import book_full as full
    from docx.shared import RGBColor
    from docx.enum.text import WD_TAB_LEADER
    p = config["params"]
    mapping = {'pageW': 'PAGE_W', 'pageH': 'PAGE_H', 'bleed': 'BLEED_MM', 'mTop': 'M_TOP', 'mBot': 'M_BOT', 'mLeft': 'M_LEFT', 'mRight': 'M_RIGHT', 'headerDist': 'HEADER_DIST', 'fTitle': 'F_TITLE', 'fAuthor': 'F_AUTHOR', 'fBody': 'F_BODY', 'fHead': 'F_HEAD', 'fNum': 'F_NUM', 'titleSize': 'TITLE_SIZE', 'titleTopLines': 'TITLE_TOP_LINES', 'titleBodyGapLines': 'TITLE_BODY_GAP_LINES', 'ruleW': 'RULE_W', 'rulePt': 'RULE_PT', 'decoW': 'DECO_W', 'decoH': 'DECO_H', 'decoX': 'DECO_X', 'decoTop': 'DECO_TOP', 'authorSize': 'AUTHOR_SIZE', 'bodySize': 'BODY_SIZE', 'bodyIndent': 'BODY_INDENT', 'bodyAfter': 'BODY_AFTER', 'bodyLine': 'BODY_LINE', 'markColorOn': 'MARK_COLOR_ON', 'widowControl': 'WIDOW_CONTROL', 'keepTogether': 'KEEP_TOGETHER', 'subtitleBold': 'SUBTITLE_BOLD', 'chapterSize': 'CHAPTER_SIZE', 'chapterTopLines': 'CHAPTER_TOP_LINES', 'chapterRuleW': 'CHAPTER_RULE_W', 'chapterRulePt': 'CHAPTER_RULE_PT', 'introSize': 'INTRO_SIZE', 'introBoxW': 'INTRO_BOX_W', 'frontSize': 'FRONT_SIZE', 'frontTopLines': 'FRONT_TOP_LINES', 'captionSize': 'CAPTION_SIZE', 'imgMaxWRatio': 'IMG_MAX_W_RATIO', 'imgBlockOn': 'IMG_BLOCK_ON', 'imgBlockSize': 'IMG_BLOCK_SIZE', 'imgBlockExpose': 'IMG_BLOCK_EXPOSE', 'imgLineOn': 'IMG_LINE_ON', 'imgLineGap': 'IMG_LINE_GAP', 'poemTitleSize': 'POEM_TITLE_SIZE', 'poemTopLines': 'POEM_TOP_LINES', 'poemTitleAfter': 'POEM_TITLE_AFTER', 'poemMetaSize': 'POEM_META_SIZE', 'poemMetaAfter': 'POEM_META_AFTER', 'poemLineSize': 'POEM_LINE_SIZE', 'poemLine': 'POEM_LINE', 'poemStanzaGap': 'POEM_STANZA_GAP', 'headSize': 'HEAD_SIZE', 'numSize': 'NUM_SIZE', 'cropMarks': 'CROP_MARKS', 'cropBlackW': 'CROP_BLACK_W', 'cropWhiteW': 'CROP_WHITE_W'}
    for key, attr in mapping.items():
        if key in p:
            value = p[key]
            default = getattr(full.book, attr)
            if isinstance(default, (int, float)) and not isinstance(default, bool):
                if not isinstance(value, (int, float)) or not 0 <= value <= 1000:
                    raise ValueError("无效参数：" + key)
            setattr(full.book, attr, value)
    for key, attr in {"fChapterNumber":"F_CHAPTER_NUMBER", "fChapterTitle":"F_CHAPTER_TITLE", "fChapterIntro":"F_CHAPTER_INTRO"}.items():
        if key in p: setattr(full.book, attr, str(p[key]))
    full.book.BOOK_COLOR = RGBColor.from_string(p["bookColor"].lstrip("#"))
    full.book.RULE_COLOR = p["ruleColor"].lstrip("#")
    if p.get("imgStyle", "").startswith("简洁居中"):
        full.book.IMG_BLOCK_ON = full.book.IMG_LINE_ON = False
    for key, attr in {"tocTitle":"TOC_TITLE", "chapterTocSize":"CHAPTER_TOC_SIZE", "chapterTocBold":"CHAPTER_TOC_BOLD", "articleIndentPt":"ARTICLE_INDENT_PT"}.items():
        if key in p: setattr(full, attr, p[key])
    for key, attr in {"tocTitleSize":"TITLE_SIZE", "tocEntrySize":"ENTRY_SIZE", "entrySpacing":"ENTRY_SPACING", "fTitle":"F_TOC"}.items():
        if key in p: setattr(full.toc, attr, p[key])
    full.toc.ENTRY_LEADER = {"空白":WD_TAB_LEADER.SPACES,"点线":WD_TAB_LEADER.DOTS,"短横线":WD_TAB_LEADER.DASHES}[p.get("tocLeader","空白")]
    full.SRC, full.BOOK_NAME, full.OUT = config["source"], config["name"], config["output"]
    full.TITLE_PAGE = bool(p.get("titlePageOn", True))
    full.TITLE_PAGE_FROM_SOURCE = bool(p.get("titlePageFromSource", True))
    full.TITLE_PAGE_TEXT = str(p.get("titlePageText", ""))
    full.TITLE_PAGE_TITLE = str(p.get("titlePageTitle", ""))
    full.TITLE_PAGE_AUTHOR = str(p.get("titlePageAuthor", ""))
    full.TITLE_PAGE_SIZE = max(10, min(72, float(p.get("titlePageSize", 28))))
    full.book.SRC = full.SRC
    # Uploaded manuscripts may reference only uploaded files, never local disk paths.
    root = Path(config["root"]).resolve()
    data = full.book.parse_book()
    for ch in data["chapters"]:
        for article in ch["articles"]:
            for kind, val in article["items"]:
                if kind == "img":
                    target = Path(val[0]).resolve()
                    if not target.is_relative_to(root):
                        raise ValueError("图片必须位于导入文件夹内：" + str(target))
                    if not target.is_file():
                        raise ValueError("找不到图片：" + str(target))
    from word_styles import install_styles
    style_params = dict(p)
    if not any(ch["title"] is not None for ch in data["chapters"]):
        style_params["chapterStyle"] = "版式一"
    install_styles(full, style_params, config.get("chapterArtwork", []))
    full.tempfile.gettempdir = lambda: config["root"]
    full.main()

if __name__ == "__main__": main()
