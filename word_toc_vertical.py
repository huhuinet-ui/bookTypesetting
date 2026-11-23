"""Reference contents: vertical chapter labels and inline title / page entries."""
import math
import re
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def split_chapter_title(title):
    match = re.match(r'^(第[一二三四五六七八九十百零〇两\d]+章|[一二三四五六七八九十百零〇两\d]+[、.．])\s*(.*)$', title or '')
    return (match.group(1), match.group(2)) if match else ('', title or '')


def article_chunks(articles, width_mm, height_mm, size):
    capacity = max(6, math.floor((height_mm-85)/(size*2.6*25.4/72)))
    chars = max(8, math.floor((width_mm-56)/(size*25.4/72)))
    chunks, chunk, used = [], [], 0
    for article in articles:
        cost = max(1, math.ceil((len(article['title'])+7)/chars))
        if chunk and used+cost > capacity:
            chunks.append(chunk); chunk, used = [], 0
        chunk.append(article); used += cost
    return chunks+[chunk] if chunk or not chunks else chunks


def install_vertical_toc(full, params):
    book = full.book

    def build_toc(doc, chapters, page_map, pad=None, preface_title=None, afterword=None, appendices=None, endmatter_order=None, endmatter_articles=None):
        sec = doc.sections[-1]
        offset = book.page_inset()
        sec.page_width=Mm(book.PAGE_W+2*offset);sec.page_height=Mm(book.PAGE_H+2*offset)
        sec.top_margin=Mm(offset+40);sec.bottom_margin=Mm(offset+20)
        sec.left_margin=sec.right_margin=Mm(offset+16)
        size=float(params.get('tocEntrySize',8)); chapter_size=float(params.get('chapterTocSize',9.5))
        digits=max(1,int(pad if pad is not None else params.get('pad',3)))
        last_volume=None
        endmatter=[{'title':title,'articles':(endmatter_articles or {}).get(key,[]),'toc_key':key} for title,key in (endmatter_order if endmatter_order is not None else ([afterword] if afterword else [])+(appendices or []))]
        for chapter_index, chapter in enumerate(list(chapters)+endmatter):
            shift = 6 if chapter_index % 2 else 0
            if chapter_index and chapter_index % 2 == 0:
                start = doc.add_paragraph()
                start.paragraph_format.page_break_before = True
                start.paragraph_format.keep_with_next = True
                start.paragraph_format.space_before = Pt(0)
                start.paragraph_format.space_after = Pt(0)
                start.paragraph_format.line_spacing = Pt(1)
            volume=chapter.get('volume')
            if volume and volume is not last_volume:
                p=doc.add_paragraph();p.paragraph_format.space_after=Pt(18);p.paragraph_format.keep_with_next=True
                number=page_map.get(volume['label'])
                label=volume['label']+(' / '+str(number).zfill(digits) if number else '')
                book.R(p,label,cn=book.F_CHAPTER_TITLE,size=chapter_size+2)
            last_volume=volume
            sequence,name=split_chapter_title(chapter.get('title'))
            for chunk in article_chunks(chapter['articles'],book.PAGE_W-shift,book.PAGE_H,size):
                table=doc.add_table(rows=1,cols=4);table.autofit=False
                widths=(7,7,10,book.PAGE_W-56-shift)
                props=table._tbl.tblPr
                total=props.find(qn('w:tblW'));total.set(qn('w:type'),'dxa');total.set(qn('w:w'),str(round((book.PAGE_W-32-shift)*72/25.4*20)))
                if shift:
                    indent=OxmlElement('w:tblInd');indent.set(qn('w:type'),'dxa');indent.set(qn('w:w'),str(round(shift*72/25.4*20)));props.append(indent)
                for col,w in zip(table.columns,widths):col.width=Mm(w)
                margins=OxmlElement('w:tblCellMar')
                for side in ('top','left','bottom','right'):
                    item=OxmlElement('w:'+side);item.set(qn('w:w'),'0');item.set(qn('w:type'),'dxa');margins.append(item)
                props.append(margins)
                borders=OxmlElement('w:tblBorders')
                for side in ('top','left','bottom','right','insideH','insideV'):
                    item=OxmlElement('w:'+side);item.set(qn('w:val'),'nil');borders.append(item)
                props.append(borders)
                row=table.rows[0]
                row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
                for cell,w in zip(row.cells,widths):cell.width=Mm(w)
                for cell,text,gray in ((row.cells[0],name,False),(row.cells[1],sequence,True)):
                    p=cell.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before=Mm(0 if gray else 5);p.paragraph_format.space_after=Pt(0)
                    p.paragraph_format.line_spacing=Pt(chapter_size*1.4)
                    p.paragraph_format.keep_together=True
                    label=('─\n'+'\n'.join(text)+'\n─') if gray and text else '\n'.join(text)
                    book.R(p,label,cn=book.F_CHAPTER_TITLE,size=chapter_size,color=RGBColor.from_string('777080') if gray else book.BLACK)
                cell=row.cells[3]
                if 'toc_key' in chapter:
                    p=cell.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                    p.paragraph_format.space_before=Mm(5)
                    p.paragraph_format.space_after=Pt(0)
                    p.paragraph_format.line_spacing=Pt(size*1.8)
                    number=page_map.get(chapter['toc_key'])
                    book.R(p,str(number).zfill(digits) if number else '',size=size,af='Times New Roman')
                for index,article in enumerate(chunk):
                    p=cell.paragraphs[0] if index==0 and 'toc_key' not in chapter else cell.add_paragraph()
                    p.paragraph_format.space_before=Mm(10 if index==0 else 0);p.paragraph_format.space_after=Pt(size*.8)
                    p.paragraph_format.line_spacing=Pt(size*1.8);p.paragraph_format.keep_together=True
                    book.R(p,article['title'],cn=book.F_BODY,size=size)
                    number=page_map.get(article.get('toc_key',article['title']))
                    book.R(p,' / '+(str(number).zfill(digits) if number else ''),size=size,af='Times New Roman')
                gap=doc.add_paragraph();gap.paragraph_format.space_after=Pt(18);gap.paragraph_format.line_spacing=Pt(1)
    full.build_toc=build_toc
    original=full.build_body
    def body(doc,chapters,first_sec,*args,**kwargs):
        first_sec.top_margin=Mm(book.M_TOP+book.page_inset());first_sec.bottom_margin=Mm(book.M_BOT+book.page_inset())
        first_sec.left_margin=Mm(book.M_LEFT+book.page_inset());first_sec.right_margin=Mm(book.M_RIGHT+book.page_inset())
        return original(doc,chapters,first_sec,*args,**kwargs)
    full.build_body=body
