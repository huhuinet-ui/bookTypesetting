"""Additional browser-selected Word styles, installed per worker process."""
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def install_styles(full, params, artwork):
    book = full.book
    if not params.get('chapterStyle', '版式一').startswith('版式一'):
        from word_chapters import install_chapters
        install_chapters(full, params, artwork)

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
            # Two explicit compact paragraphs avoid Word expanding a soft line break
            # with inherited paragraph-mark formatting or document-grid settings.
            title=doc.add_paragraph()
            english_line=doc.add_paragraph()
            for paragraph, size, after in ((title,10.57*scale,0),(english_line,9.13*scale,42*scale)):
                paragraph.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                pf=paragraph.paragraph_format
                pf.space_before=Pt(0);pf.space_after=Pt(after)
                pf.line_spacing=Pt(13*scale);pf.keep_with_next=True
                ppr=paragraph._p.get_or_add_pPr()
                snap=OxmlElement('w:snapToGrid');snap.set(qn('w:val'),'0');ppr.append(snap)
                # Format the paragraph mark too, so no larger Normal font sets its height.
                rpr=OxmlElement('w:rPr')
                sz=OxmlElement('w:sz');sz.set(qn('w:val'),str(round(size*2)));rpr.append(sz)
                ppr.append(rpr)
            full.toc.R(title, full.TOC_TITLE, 10.57*scale, bold=False)
            english=english_line.add_run('Content');english.font.name='Times New Roman'
            english.font.size=Pt(9.13*scale)
            last_volume = None
            for chapter in chapters:
                volume = chapter.get('volume')
                if volume and volume is not last_volume:
                    p=doc.add_paragraph();pf=p.paragraph_format
                    pf.space_before=Pt(24*scale);pf.space_after=Pt(14*scale)
                    pf.line_spacing=Pt(23*scale);pf.keep_with_next=True
                    full.toc.R(p,volume['number'],12.4*scale,bold=True)
                    if volume['title']: full.toc.R(p,'  '+volume['title'],15.5*scale,bold=True)
                    border=OxmlElement('w:pBdr');bottom=OxmlElement('w:bottom')
                    for key,value in [('val','single'),('sz','5'),('space','6'),('color',str(params.get('bookColor','#DEB887')).lstrip('#'))]:
                        bottom.set(qn('w:'+key),value)
                    border.append(bottom);p._p.get_or_add_pPr().append(border)
                last_volume=volume
                indent_pt=10*scale if volume else 0
                if chapter['title'] is not None:
                    p=doc.add_paragraph();pf=p.paragraph_format
                    pf.space_before=Pt(20*scale);pf.space_after=Pt(17*scale);pf.line_spacing=Pt(19*scale)
                    pf.keep_with_next=True
                    pf.left_indent=Pt(indent_pt)
                    full.toc.R(p, chapter['title'], chapter_size, bold=full.CHAPTER_TOC_BOLD)
                if not chapter['articles']: continue
                show_sequence=bool(params.get("tocShowSequence",False))
                table=doc.add_table(rows=0,cols=3 if show_sequence else 2);table.autofit=False
                sequence_mm=23*scale*25.4/72
                row_width=width-indent_pt*25.4/72
                widths=[sequence_mm,row_width-sequence_mm-width*.34,width*.34] if show_sequence else [row_width-width*.34,width*.34]
                for column,w in zip(table.columns,widths):column.width=Mm(w)
                props=table._tbl.tblPr
                if indent_pt:
                    indent=OxmlElement('w:tblInd');indent.set(qn('w:w'),str(round(indent_pt*20)));indent.set(qn('w:type'),'dxa');props.append(indent)
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
                    paragraphs=[cell.paragraphs[0] for cell in row.cells]
                    if show_sequence:
                        seq,titlep,meta=paragraphs
                        run=seq.add_run(f'{index:02d}');run.font.name='Times New Roman'
                        run.font.size=Pt(9*scale);run.bold=True;run.italic=True
                    else:
                        titlep,meta=paragraphs
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
