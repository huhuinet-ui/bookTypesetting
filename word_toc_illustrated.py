"""Native, editable right-column TOC with repeated tree-ring/leaf line art."""
import html

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor


def add_illustrated_header(section, book, title, title_size):
    width = (book.PAGE_W + 2 * book.page_inset()) * 72 / 25.4
    height = (book.PAGE_H + 2 * book.page_inset()) * 72 / 25.4
    bleed = book.page_inset() * 72 / 25.4
    trim_w = book.PAGE_W * 72 / 25.4
    trim_h = book.PAGE_H * 72 / 25.4
    safe_right = max(10, book.M_RIGHT) * 72 / 25.4
    rings = []
    for index in range(17):
        radius = 22 + index * 17
        rings.append('<v:oval style="position:absolute;left:{left};top:{top};width:{w};height:{h}" '
                     'filled="f" strokecolor="#c6bda6" strokeweight=".5pt"/>'.format(
                         left=5-radius, top=570-radius*1.18, w=radius*2, h=radius*2.36))
    branch = ('m72,888 c48,800,132,734,104,650 c84,590,40,542,61,478 c77,432,56,380,72,280 '
              'm104,650 c110,530,170,453,177,338 c185,236,237,178,275,94 '
              'm104,733 c181,706,205,736,247,699 c280,670,243,618,309,581 '
              'm177,338 c220,310,251,296,310,299 m203,237 c252,208,278,223,337,201 '
              'm142,455 c99,422,99,388,87,357 e')
    leaves = [
        'm275,94 c275,68,287,58,300,48 c297,78,286,88,275,94 x e',
        'm203,237 c181,216,181,201,185,189 c208,207,207,225,203,237 x e',
        'm260,216 c272,189,283,187,295,187 c284,211,272,216,260,216 x e',
        'm284,299 c307,276,320,276,333,280 c311,301,296,301,284,299 x e',
        'm177,338 c157,309,159,297,163,285 c183,312,182,328,177,338 x e',
        'm87,357 c64,336,65,322,69,309 c90,329,91,344,87,357 x e',
        'm309,581 c311,550,325,543,337,537 c333,566,321,575,309,581 x e',
        'm135,720 c144,690,156,686,168,682 c162,710,147,717,135,720 x e',
    ]
    position = 'position:absolute;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;'
    for index, header in enumerate((section.header, section.even_page_header)):
        header.is_linked_to_previous = False
        ident = 'toc3_' + str(index)
        leaf_xml = ''.join('<v:shape style="position:absolute;width:420;height:900" coordsize="420,900" '
                           'path="{}" fillcolor="#eef0e5" strokecolor="#89966f" strokeweight=".7pt"/>'.format(path)
                           for path in leaves)
        xml = '''<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
        <v:rect id="{ident}_background" o:allowincell="f" style="{position}margin-left:0pt;margin-top:0pt;width:{width}pt;height:{height}pt;z-index:-251658240" fillcolor="#faf8f2" stroked="f"/>
        <v:group id="{ident}_art" o:allowincell="f" coordsize="420,900" style="{position}margin-left:{bleed}pt;margin-top:{top}pt;width:{artw}pt;height:{arth}pt;z-index:-10">
        {rings}<v:shape style="position:absolute;width:420;height:900" coordsize="420,900" path="{branch}" filled="f" strokecolor="#89966f" strokeweight="1pt"/>{leaves}
        </v:group>
        <v:rect id="{ident}_title" style="{position}margin-left:{titlex}pt;margin-top:{titley}pt;width:{titlew}pt;height:{titleh}pt;z-index:1" filled="f" stroked="f">
        <v:textbox inset="0,0,0,0" style="layout-flow:vertical-ideographic"><w:txbxContent><w:p><w:pPr><w:textDirection w:val="tbRl"/><w:spacing w:before="0" w:after="0" w:line="312" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="{font}"/><w:color w:val="59604F"/><w:sz w:val="{size}"/><w:spacing w:val="120"/></w:rPr><w:t>{title}</w:t></w:r></w:p></w:txbxContent></v:textbox>
        </v:rect></w:pict>'''.format(
            ident=ident, position=position, width=width, height=height,
            bleed=bleed, top=bleed+trim_h*.14, artw=trim_w*.48, arth=trim_h*.78,
            rings=''.join(rings), branch=branch, leaves=leaf_xml,
            titlex=bleed+trim_w-safe_right-trim_w*.08, titley=bleed+trim_h*.08, titlew=trim_w*.08, titleh=trim_h*.18,
            font=html.escape(book.F_TITLE, quote=True), size=round(title_size*2), title=html.escape(title))
        pict = parse_xml(xml)
        if index == 1:
            # Even pages within the TOC retain paper color and title, without art.
            for group in list(pict.iter('{urn:schemas-microsoft-com:vml}group')):
                group.getparent().remove(group)
        header.paragraphs[0].add_run()._r.append(pict)


def install_illustrated_toc(full, params):
    book = full.book

    def build_toc(doc, chapters, page_map, pad=None, preface_title=None):
        section = doc.sections[-1]
        doc.settings.odd_and_even_pages_header_footer = True
        section.different_first_page_header_footer = False
        # Count from the first TOC page, independently of cover/preface length.
        page_number = section._sectPr.find(qn('w:pgNumType'))
        if page_number is None:
            page_number = OxmlElement('w:pgNumType')
            section._sectPr.append(page_number)
        page_number.set(qn('w:start'), '1')
        width_mm = book.PAGE_W + 2 * book.page_inset()
        height_mm = book.PAGE_H + 2 * book.page_inset()
        section.page_width = Mm(width_mm)
        section.page_height = Mm(height_mm)
        safe_right = max(10, book.M_RIGHT)
        shift_left = max(0, min(20, float(params.get("toc3ShiftLeft", 5))))
        extra_width = max(0, min(20, float(params.get("toc3ExtraWidth", 5)), safe_right+shift_left-10))
        column_margin = safe_right + shift_left - extra_width
        section.top_margin = Mm(book.page_inset() + max(book.PAGE_H*.26, book.M_TOP))
        section.bottom_margin = Mm(book.page_inset() + max(book.PAGE_H*.10, book.M_BOT))
        section.left_margin = Mm(book.page_inset() + column_margin)
        section.right_margin = Mm(book.page_inset() + column_margin)
        add_illustrated_header(section, book, full.TOC_TITLE, params.get('tocTitleSize', 20))
        text_width = book.PAGE_W*.51 - safe_right + extra_width
        chapter_size = float(params.get('chapterTocSize', 10.5))
        article_size = float(params.get('tocEntrySize', 9.5))
        digits = max(3, int(params.get('pad', 3)))
        # Equal margins neutralize the document-wide mirrored margins on verso pages.
        # The table itself reserves the illustration area on every TOC page.
        table = doc.add_table(rows=0, cols=2)
        table.autofit = False
        number_width = max(18, digits * 5) + 8
        content_width = text_width * 72 / 25.4
        widths = (content_width - number_width, number_width)
        for column, width in zip(table.columns, widths):
            column.width = Pt(width)
        props = table._tbl.tblPr
        table_width = props.find(qn('w:tblW'))
        table_width.set(qn('w:type'), 'dxa')
        table_width.set(qn('w:w'), str(round(content_width*20)))
        indent = OxmlElement('w:tblInd')
        indent.set(qn('w:w'), str(round((book.PAGE_W*.49-column_margin-shift_left)*72/25.4*20)))
        indent.set(qn('w:type'), 'dxa')
        props.append(indent)
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            element = OxmlElement('w:' + edge)
            element.set(qn('w:val'), 'nil')
            borders.append(element)
        props.append(borders)
        margins = OxmlElement('w:tblCellMar')
        for edge in ('top', 'left', 'bottom', 'right'):
            element = OxmlElement('w:' + edge)
            element.set(qn('w:w'), '0')
            element.set(qn('w:type'), 'dxa')
            margins.append(element)
        props.append(margins)


        def entry(text, key, kind, indent=0, number=None):
            row = table.add_row()
            cant_split = OxmlElement('w:cantSplit')
            row._tr.get_or_add_trPr().append(cant_split)
            for cell, width in zip(row.cells, widths):
                cell.width = Pt(width)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM if kind == 'volume' else WD_CELL_VERTICAL_ALIGNMENT.TOP
            p = row.cells[0].paragraphs[0]
            page_p = row.cells[1].paragraphs[0]
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.space_before = Pt(18 if kind == 'volume' else 10 if kind == 'chapter' else 0)
            pf.space_after = Pt(12 if kind == 'volume' else 5)
            pf.line_spacing = float(params.get('entrySpacing', 1.2))
            pf.keep_together = True
            pf.keep_with_next = kind != 'article'
            pf.left_indent = Pt(indent)
            pf.right_indent = Pt(8)
            for attr in ('space_before', 'space_after', 'line_spacing', 'keep_together', 'keep_with_next'):
                setattr(page_p.paragraph_format, attr, getattr(pf, attr))
            page_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            snap = OxmlElement('w:snapToGrid'); snap.set(qn('w:val'), '0')
            p._p.get_or_add_pPr().append(snap)
            size = 13 if kind == 'volume' else chapter_size if kind == 'chapter' else article_size
            font = book.F_CHAPTER_TITLE if kind != 'article' else book.F_BODY
            if number:
                book.R(p, number, cn=font, size=size, color=RGBColor.from_string('454D42'))
                if text:
                    p.add_run('\n')
            book.R(p, text, cn=font, size=size, bold=kind == 'chapter' and full.CHAPTER_TOC_BOLD,
                   color=RGBColor.from_string('454D42'))
            page = page_map.get(key)
            book.R(page_p, str(page).zfill(digits) if page else '', size=10, color=RGBColor.from_string('454D42'))

        last_volume = None
        for chapter in chapters:
            volume = chapter.get('volume')
            if volume and volume is not last_volume:
                entry(volume['title'], volume['label'], 'volume', number=volume['number'])
            last_volume = volume
            if chapter['title'] is not None:
                entry(chapter['title'], chapter['title'], 'chapter', 6 if volume else 0)
            for article in chapter['articles']:
                entry(article['title'], article['title'], 'article', 14 if chapter['title'] is not None else 6 if volume else 0)

    full.build_toc = build_toc
    original_body = full.build_body

    def body_with_margins(doc, chapters, first_sec, *args, **kwargs):
        first_sec.top_margin = Mm(book.M_TOP+book.page_inset())
        first_sec.bottom_margin = Mm(book.M_BOT+book.page_inset())
        first_sec.left_margin = Mm(book.M_LEFT+book.page_inset())
        first_sec.right_margin = Mm(book.M_RIGHT+book.page_inset())
        # TOC artwork belongs only to its section, including when a body section
        # initially inherits the TOC's header references.
        for header in (first_sec.header, first_sec.even_page_header):
            header.is_linked_to_previous = False
            for paragraph in header.paragraphs:
                paragraph.clear()
        return original_body(doc, chapters, first_sec, *args, **kwargs)

    full.build_body = body_with_margins
