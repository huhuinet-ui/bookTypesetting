"""Additional browser-selected Word styles, installed per worker process."""
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def install_styles(full, params, artwork):
    book = full.book
    if not params.get('chapterStyle', '版式一').startswith('版式一'):
        if not artwork:
            raise ValueError('缺少章节页图像，请更新并刷新排版工具后重试。')
        index = [0]
        # Chapter images already contain background, ornament and trim marks.
        book.add_page_background = lambda *args, **kwargs: None
        def chapter(doc, title, intro=None, anchor=None):
            if index[0] >= len(artwork):
                raise ValueError('章节页图像数量不足')
            path = artwork[index[0]]
            index[0] += 1
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = Pt(1)
            book.add_anchor(paragraph, anchor)
            run = paragraph.add_run()
            run.add_picture(path, width=Mm(book.PAGE_W+2*book.BLEED_MM), height=Mm(book.PAGE_H+2*book.BLEED_MM))
            inline = run._r.find('.//' + qn('wp:inline'))
            # Anchor to physical page, so margins do not shrink or move the artwork.
            inline.tag = qn('wp:anchor')
            for key, val in dict(simplePos='0', relativeHeight='0', behindDoc='1', locked='0', layoutInCell='1', allowOverlap='1').items():
                inline.set(key, val)
            for tag in ('simplePos', 'positionH', 'positionV'):
                element = OxmlElement('wp:'+tag)
                if tag == 'simplePos':
                    element.set('x','0');element.set('y','0')
                else:
                    element.set('relativeFrom','page')
                    offset=OxmlElement('wp:posOffset');offset.text='0';element.append(offset)
                inline.insert({'simplePos':0,'positionH':1,'positionV':2}[tag],element)
            wrap=OxmlElement('wp:wrapNone')
            inline.insert(list(inline).index(inline.find(qn('wp:docPr'))),wrap)
        # Measurement and final export traverse chapters separately.
        original_build = full.build_body
        def build_body(*args, **kwargs):
            index[0] = 0
            return original_build(*args, **kwargs)
        full.build_body = build_body
        book.add_chapter_page = chapter

    if params.get('tocStyle', '').startswith('版式二'):
        def build_toc(doc, chapters, page_map, pad=None, preface_title=None):
            # Match the browser's dedicated reference-TOC page margins.
            section=doc.sections[-1]
            section.top_margin=Mm((book.PAGE_H+2*book.BLEED_MM)*.185)
            section.bottom_margin=Mm((book.PAGE_H+2*book.BLEED_MM)*.15)
            section.left_margin=Mm((book.PAGE_W+2*book.BLEED_MM)*107.7/484.724)
            section.right_margin=Mm((book.PAGE_W+2*book.BLEED_MM)*(484.724-389.7)/484.724)
            width=(book.PAGE_W+2*book.BLEED_MM)*(389.7-107.7)/484.724
            scale=(book.PAGE_W+2*book.BLEED_MM)/(484.724*25.4/72)
            row_size=10*scale
            chapter_size=13.45*scale
            row_pitch=20*scale
            title=doc.add_paragraph();title.alignment=WD_ALIGN_PARAGRAPH.RIGHT
            pf=title.paragraph_format;pf.space_before=Pt(0);pf.space_after=Pt(42*scale)
            pf.line_spacing=Pt(13*scale);pf.keep_with_next=True
            full.toc.R(title, full.TOC_TITLE, 10.57*scale, bold=False)
            title.add_run().add_break()
            english=title.add_run('Content');english.font.name='Times New Roman'
            english.font.size=Pt(9.13*scale)
            if preface_title:
                p=doc.add_paragraph();p.paragraph_format.space_after=Pt(12)
                full.toc.R(p,preface_title,full.CHAPTER_TOC_SIZE)
            for chapter in chapters:
                if chapter['title'] is not None:
                    p=doc.add_paragraph();pf=p.paragraph_format
                    pf.space_before=Pt(20*scale);pf.space_after=Pt(17*scale);pf.line_spacing=Pt(19*scale)
                    pf.keep_with_next=True
                    full.toc.R(p, chapter['title'], chapter_size, bold=full.CHAPTER_TOC_BOLD)
                if not chapter['articles']: continue
                table=doc.add_table(rows=0,cols=3);table.autofit=False
                sequence_mm=23*scale*25.4/72
                widths=[sequence_mm,width-sequence_mm-width*.34,width*.34]
                for column,w in zip(table.columns,widths):column.width=Mm(w)
                props=table._tbl.tblPr
                borders=OxmlElement('w:tblBorders')
                for edge in ('top','left','bottom','right','insideH','insideV'):
                    element=OxmlElement('w:'+edge);element.set(qn('w:val'),'nil');borders.append(element)
                props.append(borders)
                margins=OxmlElement('w:tblCellMar')
                for edge,value in [('top',0),('bottom',0),('left',0),('right',0)]:
                    element=OxmlElement('w:'+edge);element.set(qn('w:w'),str(value));element.set(qn('w:type'),'dxa');margins.append(element)
                props.append(margins)
                for index, article in enumerate(chapter['articles'],1):
                    row=table.add_row()
                    cant=OxmlElement('w:cantSplit');row._tr.get_or_add_trPr().append(cant)
                    for cell,w in zip(row.cells,widths):
                        cell.width=Mm(w)
                        paragraph=cell.paragraphs[0];pf=paragraph.paragraph_format
                        pf.space_before=Pt(0);pf.space_after=Pt(0);pf.line_spacing=Pt(row_pitch)
                        snap=OxmlElement('w:snapToGrid');snap.set(qn('w:val'),'0');paragraph._p.get_or_add_pPr().append(snap)
                    seq,titlep,meta=[cell.paragraphs[0] for cell in row.cells]
                    run=seq.add_run(f'{index:02d}');run.font.name='Times New Roman'
                    run.font.size=Pt(9*scale);run.bold=True;run.italic=True
                    full.toc.R(titlep,article['title'],row_size)
                    # Leave the same 8 pt breathing room as the preview's flex gap.
                    titlep.paragraph_format.right_indent=Pt(8)
                    meta.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                    page=page_map.get(article['title'])
                    number=str(page).zfill(max(3,int(params.get('pad',3)))) if page else ''
                    author=article.get('author') or ''
                    if author:
                        run=meta.add_run(author+' / ');run.font.size=Pt(row_size)
                        run.font.name=book.F_AUTHOR
                        run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),book.F_AUTHOR)
                    run=meta.add_run(number);run.font.name='Times New Roman';run.font.size=Pt(8.77*scale)
        full.build_toc=build_toc
        original_body=full.build_body
        def body_with_margins(doc, chapters, first_sec, *args, **kwargs):
            # add_section inherits TOC margins; restore normal body geometry.
            first_sec.top_margin=Mm(book.M_TOP+book.BLEED_MM)
            first_sec.bottom_margin=Mm(book.M_BOT+book.BLEED_MM)
            first_sec.left_margin=Mm(book.M_LEFT+book.BLEED_MM)
            first_sec.right_margin=Mm(book.M_RIGHT+book.BLEED_MM)
            return original_body(doc,chapters,first_sec,*args,**kwargs)
        full.build_body=body_with_margins
