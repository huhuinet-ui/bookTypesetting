# -*- coding: utf-8 -*-
"""把 Word(.docx) 转成 .md（供 book_layout.py / book_full.py 排版用）。

用法：python word2md.py "学员.docx" [输出.md]

转换规则：
- 按 Word 标题样式识别结构：标题3→### 章节、标题4→#### 文章、标题5→##### 小节
- 「自序/前言/序」这类标题4 → # 前言
- 图片自动提取到「图/」子文件夹，生成 ![图注](图/imgN.png)
- 图注：支持跳过图片后的空段落、同段说明及“图片说明：/图注：/图1：”等写法；下一段若为「图：xxx」「图1-1：xxx」或 Word 题注样式，自动取为图注
  （带编号的「图1-1：xxx」整行保留，含编号）
- 兜底（样式丢失时）：加粗独立短段→文章标题；「数字.」开头独立短段→小节标题；
  这些靠猜的行会列在最后的「需人工核对」清单里
- 可选开关（默认关闭）：整段被【】包起来的行→去掉【】。既处理普通正文段（转成 ##### 小标题），
  也处理带标题样式的段（如「标题5」的【xxx】→ ##### xxx，只去括号、层级不变）。
  打开方式：把 BRACKET_AS_SUBTITLE 改成 True，或命令行加 --bracket-subtitle
"""
import io, os, re, sys
if (getattr(sys.stdout, "encoding", "") or "").lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from docx import Document
from docx.oxml.ns import qn

PREFACE_KEYS = ("自序", "前言", "序言", "代序", "写在前面")
IMG_EXT = {"jpeg": "jpg", "jpe": "jpg"}

# 开关：把「整段被【】包起来」的行识别成正文小标题（输出 #####，去掉【】）。默认关闭。
# 打开方式：改成 True，或运行时加命令行参数 --bracket-subtitle。
BRACKET_AS_SUBTITLE = False
BRACKET_SUB_RE = re.compile(r"^【\s*(.+?)\s*】$")   # 整行形如【xxx】才认；xxx 作为小标题文字

# 图注识别（图片的下一段）：
#   带编号：图1-1：xxx / 图 1-1 xxx / 图2-6：xxx …… → 整行作为图注（保留编号）
#   旧写法：图：xxx → 去掉「图：」前缀，取后面的说明
#   另外，Word「题注」样式(Caption)的段落也直接当图注
CAP_NUM_RE = re.compile(r"^图\s*\d+(?:\s*[-－—–.]\s*\d+)*(?:\s*[:：、.．)）]\s*|\s+|(?=[^\d\s])).+$")
CAP_PLAIN_RE = re.compile(r"^(?:图片说明|图注说明|图片名|图片|图注|图说|照片说明|说明|图)\s*[:：]\s*(.+)$")
CAP_STYLE_RE = re.compile(r"题注|caption", re.I)


def caption_text(para):
    # Text in Word textboxes may be absent from Paragraph.text.
    text = para.text.strip()
    if not text:
        text = "".join(node.text or "" for node in para._p.iter(qn("w:t"))).strip()
    return text.replace("\u2060", "").replace("\ufeff", "")


def find_caption(paras, index):
    """Skip empty paragraphs, never cross an intervening table, heading or image."""
    previous = paras[index]
    for j in range(index+1, len(paras)):
        candidate = paras[j]
        if previous._p.getnext() is not candidate._p:
            return None, index
        if candidate._p.find(".//"+qn("a:blip")) is not None:
            return None, index
        if heading_level(candidate.style.name or ""):
            return None, index
        text = caption_text(candidate)
        if text:
            return match_caption(candidate), j
        previous = candidate
    return None, index


def match_caption(para):
    """这一段是不是图注？是则返回图注文字，否则返回 None。"""
    t = caption_text(para)
    if not t:
        return None
    if CAP_STYLE_RE.search(para.style.name or ""):
        return t
    if CAP_NUM_RE.match(t):          # 图1-1：xxx —— 整行都要（编号是图注的一部分）
        return t
    m = CAP_PLAIN_RE.match(t)        # 图：xxx —— 只取说明部分（保持旧行为）
    if m:
        return m.group(1).strip()
    return None


def heading_level(style_name):
    """按样式名返回标题层级 1~6，非标题（正文）返回 0。"""
    m = re.search(r"(?:heading|标题)\s*([1-6])", style_name, re.I)
    return int(m.group(1)) if m else 0


def clean_title(s):
    """去掉学员加的「标题1:」「标题：」这类占位前缀。"""
    return re.sub(r"^标题[0-9\s]*[:：]\s*", "", s).strip()


def _empty(paras, j):
    return j < 0 or j >= len(paras) or not paras[j].text.strip()


def convert(src, out=None, bracket_subtitle=None):
    if bracket_subtitle is None:
        bracket_subtitle = BRACKET_AS_SUBTITLE
    src = os.path.abspath(src)
    doc = Document(src)
    if out is None:
        out = os.path.splitext(src)[0] + ".md"
    img_dir = os.path.join(os.path.dirname(src), "图")
    os.makedirs(img_dir, exist_ok=True)

    paras = doc.paragraphs
    n = len(paras)
    lines, suspicious = [], []
    img_no = 0
    i = 0
    while i < n:
        p = paras[i]
        text = p.text.strip()
        style = p.style.name

        # 图片段落：提取到 图/imgN.png，生成 ![](图/imgN.png)；下一段是图注（图：xxx / 图1-1：xxx / 题注样式）
        blip = p._p.find(".//" + qn("a:blip"))
        if blip is not None:
            rid = blip.get(qn("r:embed"))
            part = doc.part.related_parts.get(rid)
            if part is not None:
                ext = part.content_type.split("/")[-1].lower()
                ext = IMG_EXT.get(ext, ext)
                img_no += 1
                fn = f"img{img_no}.{ext}"
                with open(os.path.join(img_dir, fn), "wb") as f:
                    f.write(part.blob)
                # Prefer explicit same-paragraph caption; otherwise scan past blanks.
                caption = match_caption(p) if text else None
                if caption is None:
                    caption, caption_index = find_caption(paras, i)
                    if caption is not None:
                        i = caption_index
                if text and caption is None:
                    # Do not silently discard text that shares an image paragraph.
                    lines.append(text)
                lines.append(f"![{caption or ''}](图/{fn})")
            i += 1
            continue

        if not text:
            i += 1
            continue

        # 有标题样式的：直接映射
        level = heading_level(style)
        if level:
            t = clean_title(text)
            if bracket_subtitle:            # 开关打开：标题若整段被【】包起来，去掉【】（如「标题5」样式的小节）
                bm = BRACKET_SUB_RE.match(t)
                if bm:
                    t = bm.group(1).strip()
            if level == 4 and any(t.startswith(key) for key in PREFACE_KEYS):
                lines.append(f"# {t}")
            else:
                lines.append(f"{'#' * level} {t}")
            i += 1
            continue

        # 开关打开时：整段被【】包起来 → 正文小标题（#####），去掉【】
        if bracket_subtitle:
            bm = BRACKET_SUB_RE.match(text)
            if bm:
                lines.append(f"##### {bm.group(1).strip()}")
                i += 1
                continue

        # 兜底：样式丢失，靠「加粗独立短段」猜文章标题、「数字.开头独立短段」猜小节标题
        bold_all = all(r.bold for r in p.runs if r.text.strip())
        if bold_all and len(text) <= 30 and _empty(paras, i - 1) and _empty(paras, i + 1):
            lines.append(f"#### {clean_title(text)}")
            suspicious.append(("文章标题(猜)", clean_title(text)))
            i += 1
            continue
        if re.match(r"^\d{1,2}\.\s", text) and len(text) <= 30 and _empty(paras, i - 1):
            lines.append(f"##### {text}")
            suspicious.append(("小节标题(猜)", text))
            i += 1
            continue

        lines.append(text)
        i += 1

    with open(out, "w", encoding="utf-8") as f:
        f.write("\n\n".join(lines) + "\n")

    print(f"已生成 {out}")
    print(f"提取图片 {img_no} 张 → {img_dir}")
    if suspicious:
        print(f"\n以下 {len(suspicious)} 处是「样式丢失后靠猜」的标题，请人工核对：")
        for kind, t in suspicious:
            print(f"  [{kind}] {t}")
    return out, img_no


if __name__ == "__main__":
    args = sys.argv[1:]
    bracket = False
    if "--bracket-subtitle" in args:
        bracket = True
        args.remove("--bracket-subtitle")
    if not args:
        print("用法：python word2md.py \"学员.docx\" [输出.md] [--bracket-subtitle]")
        print("  --bracket-subtitle  把整段被【】包起来的行当正文小标题（#####，去掉【】）；默认关闭")
        sys.exit(1)
    convert(args[0], args[1] if len(args) > 1 else None, bracket_subtitle=bracket or None)
