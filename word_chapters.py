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


def install_chapters(full, params, artwork):
    book=full.book
    style=str(params.get('chapterStyle',''))
    second=style.startswith('版式二')
    plane='纸飞机' in style or style.startswith('版式五')
    prefix='chapter5' if plane else 'chapter4'
    W=(book.PAGE_W+2*book.BLEED_MM)*72/25.4
    H=(book.PAGE_H+2*book.BLEED_MM)*72/25.4
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
                if underline and float(params.get('chapterRulePt',book.CHAPTER_RULE_PT)) > 0:
                    weight=max(2,round(float(params.get('chapterRulePt',book.CHAPTER_RULE_PT))*8))
                    rule_color=str(params.get('ruleColor','#000000')).lstrip('#')
                    paragraph_border=f'<w:pBdr><w:bottom w:val="single" w:sz="{weight}" w:space="5" w:color="{rule_color}"/></w:pBdr>'
                    textbox_style=' style="mso-fit-shape-to-text:t"'
                lines=[]
                for line in str(text).split('\n'):
                    # Word's vertical VML layout can collapse ordinary spaces.
                    # Explicit nonbreaking spaces retain the source separator.
                    if preserve_spaces:
                        line=line.replace(" ", "\u00a0").replace("\t", "\u00a0")
                    character_spacing=f'<w:spacing w:val="{round(letter_spacing*20)}"/>'
                    lines.append(f'<w:p><w:pPr><w:jc w:val="{alignment}"/><w:spacing w:before="0" w:after="0" w:line="{round(size*1.6*20)}" w:lineRule="atLeast"/><w:snapToGrid w:val="0"/><w:kinsoku w:val="1"/><w:widowControl w:val="1"/>{paragraph_border}{direction}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{html.escape(font or book.F_CHAPTER_TITLE,quote=True)}"/><w:color w:val="{color.lstrip("#")}"/><w:sz w:val="{round(size*2)}"/>{character_spacing}</w:rPr><w:t xml:space="preserve">{html.escape(line)}</w:t></w:r></w:p>')
                xml+=f'<v:textbox{textbox_style} inset="0,0,0,0"><w:txbxContent>{"".join(lines)}</w:txbxContent></v:textbox>'
            xml+='</v:rect></w:pict>'
            paragraph.add_run()._r.append(parse_xml(xml))
        if second:
            inset=(book.BLEED_MM+params.get('chapter2Inset',13))*72/25.4
            shape(inset,inset,W-2*inset,H-2*inset,border=params.get('bookColor','#DEB887'))
            if artwork:shape(0,0,W,H,image=artwork[0])
            size=params.get('chapter2TitleSize',18)
            shape(W*.46,H*params.get('chapter2Top',40)/100,size*1.7,H*.52,title,size=size,color=params.get('bookColor','#DEB887'),vertical=True,fill='#FFFFFF',preserve_spaces=True,letter_spacing=3)
        else:
            match=re.match(r'^(第[一二三四五六七八九十百零〇两\d]+[章篇部卷])[\s　:：、.．—-]*(.*)$',title)
            number=match[1] if match else f'第{ordinal[0]}章'
            name=match[2] if match and match[2] else title
            shape(W*.36,H*.20,W*.46,H*.08,number,font=book.F_CHAPTER_NUMBER,size=params.get(prefix+'NumberSize',21),color=params.get(prefix+'NumberColor','#20251c'))
            title_size=params.get(prefix+'TitleSize',21)
            shape(W*.24,H*.29,W*.58,title_size*1.6+12,name,font=book.F_CHAPTER_TITLE,size=title_size,color=params.get(prefix+'TitleColor','#20251c'),underline=True)
            if artwork:shape(W*.70,H*.43,W*.13,H*.09,image=artwork[0])
            if intro:shape(W*.44,H*.63,W*.38,H*.24,wrap_vertical_intro(intro,max(3,int(H*.24/(params.get(prefix+'IntroSize',10)+1))-1)),font=book.F_CHAPTER_INTRO,size=params.get(prefix+'IntroSize',10),color=params.get(prefix+'IntroColor','#20251c'),vertical=True)
    book.add_chapter_page=chapter
