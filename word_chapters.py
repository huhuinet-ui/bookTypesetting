"""Editable page-positioned chapter text and independent decoration assets."""
import html
import re
from docx.oxml import parse_xml
from docx.shared import Pt

def wrap_vertical_intro(text, capacity):
    """Plan vertical columns without single-character tails or stranded punctuation."""
    closing=set("，。！？；：、,.!?;:）)】〕］〉》」』”’…")
    opening=set("（(【〔［〈《「『“‘")
    result=[]
    for paragraph in str(text).split("\n"):
        remaining=paragraph.strip()
        if not remaining:
            result.append("");continue
        while len(remaining)>capacity:
            end=capacity
            while end>2 and (remaining[end] in closing or remaining[end-1] in opening):end-=1
            tail=remaining[end:]
            if len(tail.rstrip("".join(closing)))<2:
                end=max(2,end-(2-len(tail.rstrip("".join(closing)))))
                while end>1 and (remaining[end] in closing or remaining[end-1] in opening):end-=1
            result.append(remaining[:end]);remaining=remaining[end:]
        result.append(remaining)
    return "\n".join(result)


def split_chapter_heading(title):
    """Preserve the manuscript's sequence label; never invent a chapter number."""
    text=str(title).strip()
    numbers=r"[一二三四五六七八九十百千零〇两壹贰叁肆伍陆柒捌玖拾佰仟０-９0-9]+"
    separators=r"[\s　|｜¦:：/／_·•、.．—–-]"
    patterns=[
        rf"^((?:十一|十二|十|[一二三四五六七八九])月)(?:{separators}+(.*)|$)",
        rf"^((?:第)?{numbers}(?:部分|篇章|章节|单元|章|篇|部|卷|集|辑|节)|(?:部分|篇章|章节|单元|章|篇|部|卷|集|辑|节){numbers}|[上中下][章篇部卷集辑]){separators}*(.*)$",
        rf"^({numbers}[、.．）)]){separators}*(.*)$",
        rf"^({numbers}){separators}+(.*)$",
    ]
    for pattern in patterns:
        match=re.match(pattern,text)
        if match:return match[1],(match[2] or '').strip()
    return "",text


def tree_art_box(width,height,params,title_size,has_intro=True):
    scale=max(.5,min(3,float(params.get('chapter6ArtScale',180))/100))
    cx,cy=width*.765,height*.475
    base_w,base_h=width*.13,height*.09
    top=cy-base_h*scale/2
    if params.get('chapter6ArtAutoFit',True):
        safe=(params.get('bleed',3)+params.get('printMargin',10)+2)*72/25.4
        lower=height*.29+title_size*1.6+12+8
        upper=height*.63-8 if has_intro else height-safe
        scale=max(0,min(scale,(upper-lower)/base_h,2*min(cx-safe,width-safe-cx)/base_w))
        top=max(lower,min(cy-base_h*scale/2,upper-base_h*scale))
    return cx-base_w*scale/2,top,base_w*scale,base_h*scale


def install_chapters(full, params, artwork):
    book=full.book
    style=str(params.get('chapterStyle',''))
    second=style.startswith('版式二')
    plane='纸飞机' in style
    tree='指纹树' in style or style.startswith('版式五')
    prefix='chapter5' if plane or tree else 'chapter4'
    chapter_color=params.get('chapterTextColor',params.get(prefix+'TitleColor','#20251c'))
    W=(book.PAGE_W+2*book.page_inset())*72/25.4
    H=(book.PAGE_H+2*book.page_inset())*72/25.4
    original_background=book.add_page_background
    background = (params.get('bookColor', '#DEB887') if params.get('chapterBgFollowTheme', True)
                  else params.get('chapter4Bg', '#DFE9C9'))
    book.add_page_background=lambda sec,color:original_background(sec,'FFFFFF' if second else background.lstrip('#'))
    book.add_volume_background=lambda sec:original_background(sec,str(book.BOOK_COLOR))
    ordinal=[0]
    original_body=full.build_body
    def body(*args,**kwargs):
        ordinal[0]=0
        return original_body(*args,**kwargs)
    full.build_body=body
    def chapter(doc,title,intro=None,anchor=None):
        ordinal[0]+=1
        paragraph=doc.add_paragraph()
        paragraph.paragraph_format.space_after=Pt(0)
        paragraph.paragraph_format.line_spacing=Pt(1)
        book.add_anchor(paragraph,anchor)
        def shape(x,y,w,h,text='',font=None,size=12,color='#20251c',vertical=False,border=None,image=None,fill=None,underline=False,preserve_spaces=False,letter_spacing=0):
            ident=f'chapter_{ordinal[0]}_{len(paragraph.runs)}'
            attr=f'fillcolor="{html.escape(fill)}"' if fill else 'filled="f"'
            attr+=f' strokecolor="{html.escape(border)}" strokeweight="{params.get("chapterRulePt",1)}pt"' if border else ' stroked="f"'
            xml=f'<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v:rect id="{ident}" style="position:absolute;margin-left:{x}pt;margin-top:{y}pt;width:{w}pt;height:{h}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:1" {attr}>'
            if image:
                rid,_=doc.part.get_or_add_image(image)
                xml+=f'<v:imagedata r:id="{rid}"/>'
            else:
                direction=""
                textbox_style=' style="layout-flow:vertical-ideographic"' if vertical else ""
                alignment="left" if vertical else "right"
                paragraph_border=""
                text_underline=""
                if underline and float(params.get('chapterRulePt',book.CHAPTER_RULE_PT)) > 0:
                    rule_color=str(chapter_color).lstrip('#')
                    text_underline=f'<w:u w:val="single" w:color="{rule_color}"/>'
                    textbox_style=' style="mso-fit-shape-to-text:t"'
                lines=[]
                for line in str(text).split('\n'):
                    # Word's vertical VML layout can collapse ordinary spaces.
                    # Explicit nonbreaking spaces retain the source separator.
                    if preserve_spaces:
                        line=line.replace(" ", "\u00a0").replace("\t", "\u00a0")
                    character_spacing=f'<w:spacing w:val="{round(letter_spacing*20)}"/>'
                    lines.append(f'<w:p><w:pPr><w:jc w:val="{alignment}"/><w:spacing w:before="0" w:after="0" w:line="{round(size*1.6*20)}" w:lineRule="atLeast"/><w:snapToGrid w:val="0"/><w:kinsoku w:val="1"/><w:widowControl w:val="1"/>{paragraph_border}{direction}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{html.escape(font or book.F_CHAPTER_TITLE,quote=True)}"/><w:color w:val="{color.lstrip("#")}"/><w:sz w:val="{round(size*2)}"/>{text_underline}{character_spacing}</w:rPr><w:t xml:space="preserve">{html.escape(line)}</w:t></w:r></w:p>')
                xml+=f'<v:textbox{textbox_style} inset="0,0,0,0"><w:txbxContent>{"".join(lines)}</w:txbxContent></v:textbox>'
            xml+='</v:rect></w:pict>'
            paragraph.add_run()._r.append(parse_xml(xml))
        if second:
            inset=(book.page_inset()+params.get('chapter2Inset',13))*72/25.4
            if params.get('chapter2FrameOn',True):
                shape(inset,inset,W-2*inset,H-2*inset,border=params.get('bookColor','#DEB887'))
            if artwork:shape(0,0,W,H,image=artwork[0])
            size=params.get('chapter2TitleSize',18)
            number,name=split_chapter_heading(title)
            label=' '.join(part for part in (number,name) if part) if number else title.strip()
            top=book.page_inset()+book.PAGE_H*params.get('chapter2Top',40)/100
            bottom=max(8,params.get('chapter2Inset',13) if params.get('chapter2FrameOn',True) else 8)
            height=max(1,(book.PAGE_H*(1-params.get('chapter2Top',40)/100)-bottom)*72/25.4-8)
            if params.get('chapter2SingleColumn',True):
                size=book.fitted_title_size('chapter2',label,size)
            shape(W*.46,top*72/25.4,size*1.7,height,label,size=size,color=params.get('bookColor','#DEB887'),vertical=True,fill='#FFFFFF',preserve_spaces=True,letter_spacing=3)
        else:
            number,name=split_chapter_heading(title)
            if number:
                shape(W*.36,H*.20,W*.46,H*.08,number,font=book.F_CHAPTER_NUMBER,size=params.get(prefix+'NumberSize',21),color=chapter_color)
            title_size=params.get(prefix+'TitleSize',21)
            title_size=book.fitted_title_size('chapter',name,title_size)
            shape(W*.24,H*.29,W*.58,title_size*1.6+12,name,font=book.F_CHAPTER_TITLE,size=title_size,color=chapter_color,underline=True)
            if artwork:
                box=tree_art_box(W,H,params,title_size,bool(intro)) if tree else (W*.70,H*.43,W*.13,H*.09)
                shape(*box,image=artwork[0])
            if intro:shape(W*.44,H*.63,W*.38,H*.24,wrap_vertical_intro(intro,max(3,int(H*.24/(params.get(prefix+'IntroSize',10)+1))-1)),font=book.F_CHAPTER_INTRO,size=params.get(prefix+'IntroSize',10),color=chapter_color,vertical=True)
    book.add_chapter_page=chapter
