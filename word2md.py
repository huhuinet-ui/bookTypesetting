# -*- coding: utf-8 -*-
"""把 Word(.docx) 转成 .md（供 book_layout.py / book_full.py 排版用）。

用法：python word2md.py "学员.docx" [输出.md]

转换规则：
- 按 Word 标题样式识别结构：标题3→### 章节、标题4→#### 文章、标题5→##### 小节
- 「自序/前言/序」这类标题4 → # 前言
- 图片自动提取到「图/」子文件夹，生成 ![图注](图/imgN.png)
- 图注：图片下一行「图：xxx」自动变成图注
- 兜底（样式丢失时）：加粗独立短段→文章标题；「数字.」开头独立短段→小节标题；
  这些靠猜的行会列在最后的「需人工核对」清单里
"""
import io, os, re, sys
if (getattr(sys.stdout, "encoding", "") or "").lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from docx import Document
from docx.oxml.ns import qn

PREFACE_KEYS = ("自序", "前言", "序言", "代序", "写在前面")
IMG_EXT = {"jpeg": "jpg", "jpe": "jpg"}


def heading_level(style_name):
    """按样式名返回标题层级 1~6，非标题（正文）返回 0。"""
    m = re.search(r"(?:heading|标题)\s*([1-6])", style_name, re.I)
    return int(m.group(1)) if m else 0


def clean_title(s):
    """去掉学员加的「标题1:」「标题：」这类占位前缀。"""
    return re.sub(r"^标题[0-9\s]*[:：]\s*", "", s).strip()


def _empty(paras, j):
    return j < 0 or j >= len(paras) or not paras[j].text.strip()


def convert(src, out=None):
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

        # 图片段落：提取到 图/imgN.png，生成 ![](图/imgN.png)；下一行「图：xxx」=图注
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
                caption = None
                if i + 1 < n:
                    cm = re.match(r"^图\s*[:：]\s*(.+)$", paras[i + 1].text.strip())
                    if cm:
                        caption = cm.group(1).strip()
                        i += 1
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
            if level == 4 and t.rstrip("。：:") in PREFACE_KEYS:
                lines.append(f"# {t}")
            else:
                lines.append(f"{'#' * level} {t}")
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
    if len(sys.argv) < 2:
        print("用法：python word2md.py \"学员.docx\" [输出.md]")
        sys.exit(1)
    convert(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
