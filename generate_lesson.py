#!/usr/bin/env python3
"""Offline premium lesson generator with adaptive vector infographics. No image APIs or search.

python -m pip install -r requirements-lesson.txt
python generate_lesson.py --fonts-dir /path/to/fonts
python generate_lesson.py --fonts-dir /path/to/fonts --push
Fonts: georgia.ttf, georgiai.ttf, arial.ttf. Narration is preserved in slide Notes.
"""
from __future__ import annotations
import argparse
from dataclasses import dataclass
import json
from pathlib import Path
import re
import subprocess

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.xmlchemy import OxmlElement
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

NAME='Ваш Кристалл незабываемости'
BG=(26,26,26); GOLD=(212,175,55); WHITE=(255,255,255); CARD=(40,40,40)
PAGE_W=13.333*72; PAGE_H=7.5*72
LEFT=4.5*72; COLUMN=8.3*72
TEXT_TOP=130; TEXT_BOTTOM=488
TERMS=re.compile(r'(\b(?:следы|след|лёгкость|память|уважение|вдохновение|комбинацией|незабываемая|незабываемые)\b)',re.I)
FONT_FILES={'Georgia':'georgia.ttf', 'GeorgiaItalic':'georgiai.ttf', 'Arial':'arial.ttf'}


def normalized(text):return ' '.join(text.split())
def color(value):return RGBColor(*value)
def pdf_color(value):return tuple(v/255 for v in value)


def load_content(path):
    if path.suffix.lower()=='.json':
        content=json.loads(path.read_text(encoding='utf-8'))
        if isinstance(content,dict):content=content['lessons']
        assert all(d.get('title') and d.get('notes','').strip() for d in content)
        return content
    # Markdown headings define topics, never fixed slide boundaries.
    result=[]; title=NAME; paragraphs=[]
    for raw in path.read_text(encoding='utf-8').splitlines():
        line=re.sub(r'\*\*|__','',raw).strip().lstrip('> ')
        if re.match(r'^#{1,6}\s',line):
            if paragraphs:result.append(dict(title=title,notes='\n'.join(paragraphs)))
            title=re.sub(r'^#{1,6}\s*(?:Слайд\s*\d+[.:]?\s*)?','',line)
            paragraphs=[]
        elif line and line!='---':paragraphs.append(line)
    if paragraphs:result.append(dict(title=title,notes='\n'.join(paragraphs)))
    return result


def wrap(value,font,size,width):
    lines=[]
    for part in value.split('\n'):
        current=''
        for word in part.split():
            candidate=(current+' '+word).strip()
            if current and pdfmetrics.stringWidth(candidate,font,size)>width:
                lines.append(current);current=word
            else:current=candidate
        lines.append(current)
    if any(pdfmetrics.stringWidth(row,font,size)>width for row in lines):
        raise ValueError('An unbroken word exceeds its text column; edit that word or formula.')
    return lines


def caption(paragraph):
    # Keep the author's words; choose a complete sentence or clause as its thesis.
    if len(paragraph)<=175:return paragraph
    sentences=re.split(r'(?<=[.!?])\s+',paragraph)
    if len(sentences[0])<=175:return sentences[0]
    clauses=re.split(r'(?<=[,;:])\s+',paragraph)
    text=clauses[0]
    for clause in clauses[1:]:
        if len(text)+len(clause)+1>160:break
        text+=' '+clause
    if len(text)>175:
        words=text.split();text=''
        for word in words:
            if len(text)+len(word)+1>175:break
            text=(text+' '+word).strip()
    return text


def narration_fragments(paragraph):
    fragments=[]
    for sentence in re.split(r'(?<=[.!?])\s+',paragraph):
        while len(sentence)>420:
            # Prefer a clause boundary; fall back to a word boundary, never drop speech.
            boundaries=[m.end() for m in re.finditer(r'[,;:]\s+',sentence[:420])]
            cut=boundaries[-1] if boundaries else sentence.rfind(' ',0,420)
            if cut<=0:raise ValueError('An unbroken narration token exceeds 420 characters.')
            fragments.append(sentence[:cut].strip());sentence=sentence[cut:].strip()
        if sentence:fragments.append(sentence)
    return fragments


@dataclass
class Unit:
    thesis:str
    notes:str
    quote:bool=False


@dataclass
class TeachingSlide:
    title:str
    units:list[Unit]
    @property
    def notes(self):return '\n'.join(u.notes for u in self.units)


def paginate(content,max_slides):
    slides=[]
    for topic in content:
        units=[]; prefix=[]
        for paragraph in filter(None,(p.strip() for p in topic['notes'].splitlines())):
            if normalized(paragraph).lower()==normalized(topic['title']).lower() or paragraph in ('Вводный блок','«Ваш Кристалл незабываемости»','Почему одних женщин забывают, а других — нет?','Например'):
                prefix.append(paragraph);continue
            # Long narration paragraphs split at complete sentences, preserving all speech.
            fragments=narration_fragments(paragraph) if len(paragraph)>420 else [paragraph]
            for fragment in fragments:
                notes='\n'.join(prefix+[fragment]);prefix=[]
                quote=any(word in fragment for word in ('А есть незабываемые.','«хорошая женщина»','Чем больше раскрытых граней','Вот такую женщину','Он запоминает себя рядом с вами.','Она создаётся комбинацией.','Но это один след.'))
                units.append(Unit(caption(fragment),notes,quote))
        if prefix:
            if units:units[-1].notes+='\n'+'\n'.join(prefix)
            else:units=[Unit(caption(' '.join(prefix)),'\n'.join(prefix))]
        topic_pages=[];batch=[]; height=0
        for unit in units:
            font='GeorgiaItalic' if unit.quote else 'Arial'
            rows=wrap(unit.thesis,font,18,COLUMN-28)
            required=len(rows)*22.32+26
            if required>TEXT_BOTTOM-TEXT_TOP:
                raise ValueError('A thesis is too long: split its sentence into shorter clauses.')
            if batch and (len(batch)==3 or height+required>TEXT_BOTTOM-TEXT_TOP):
                topic_pages.append(TeachingSlide(topic['title'],batch));batch=[];height=0
            batch.append(unit);height+=required+10
        if batch:topic_pages.append(TeachingSlide(topic['title'],batch))
        # Rebalance a 3+1 split into 2+2 when the final card would be orphaned.
        if len(topic_pages)>1 and len(topic_pages[-1].units)==1 and len(topic_pages[-2].units)>=3:
            candidate=[topic_pages[-2].units[-1]]+topic_pages[-1].units
            needed=sum(len(wrap(u.thesis,'GeorgiaItalic' if u.quote else 'Arial',18,COLUMN-28))*22.32+36 for u in candidate)
            if needed-10<=TEXT_BOTTOM-TEXT_TOP:
                topic_pages[-1].units.insert(0,topic_pages[-2].units.pop())
        slides.extend(topic_pages)
    expected=normalized('\n'.join(d['notes'] for d in content))
    actual=normalized('\n'.join(s.notes for s in slides))
    assert expected==actual,'Narration must remain complete, ordered and free of duplication.'
    if len(slides)>max_slides:raise ValueError(f'{len(slides)} slides exceed --max-slides={max_slides}; no truncation performed.')
    return slides


def infer_layout(lesson):
    speech=lesson.notes
    # Meaning determines the scheme; slide numbers and page positions are irrelevant.
    if 'ГРАНЬ → СОПРИКОСНОВЕНИЕ' in speech:
        return 'chain',['ГРАНЬ','СОПРИКОСНОВЕНИЕ','СЛЕД','ПАМЯТЬ','НЕЗАБЫВАЕМОСТЬ']
    pairs=[]
    for paragraph in speech.splitlines():
        if '→' in paragraph:
            pairs.extend([(a.strip(),b.strip()) for a,b in re.findall(r'([^,→]+?)\s*→\s*([^,→]+)',paragraph)])
    if pairs:return 'relations',pairs[:4]
    if 'Проходит несколько дней.' in speech and 'Так появляется память о вас.' in speech:
        return 'chain',['Мужчина вспоминает вас','возвращается это ощущение','появляется память о вас']
    contrasts=re.findall(r'(?:она\s+)?([^.!?,:]+?),\s*но\s+([^.!?]+)',speech,re.I)
    if len(contrasts)>=2:
        return 'comparison',(['\n'.join(a.strip() for a,b in contrasts),'\n'.join(b.strip() for a,b in contrasts)])
    if 'И наоборот:' in speech:
        return 'comparison',['потрясающая женщина —\nкрасивая, умная, успешная','сначала интерес,\nпотом привыкание,\nпотом охлаждение.']
    if 'Он может забыть' in speech and 'как ему было рядом с вами' in speech:
        return 'comparison',['Он может забыть ресторан.\nОн может даже забыть какие-то ваши слова.','как ему было рядом с вами.']
    if 'Есть красивые женщины.' in speech and 'А есть незабываемые.' in speech:
        return 'comparison',['Есть красивые женщины.\nЕсть умные.\nЕсть сексуальные.\nЕсть интересные.','А есть незабываемые.']
    if all(word in speech for word in ('сначала интерес','потом привыкание','потом охлаждение')):
        return 'chain',['интерес','привыкание','охлаждение']
    if len(lesson.units)==3 and not any(u.quote for u in lesson.units):
        return 'tiles',[u.thesis for u in lesson.units]
    return 'list',None


class Deck:
    def __init__(self,output):
        self.output=output;self.prs=Presentation();self.layouts=[]
        self.prs.slide_width=Inches(13.333);self.prs.slide_height=Inches(7.5)
        self.prs.core_properties.title=NAME;self.prs.core_properties.language='ru-RU'
        # Remove slide-number fields from every master and layout as well.
        templates=list(self.prs.slide_masters)+[layout for master in self.prs.slide_masters for layout in master.slide_layouts]
        for template in templates:
            for shape in list(template.shapes):
                if shape._element.xpath('.//p:ph[@type="sldNum"]') or shape._element.xpath('.//a:fld[@type="slidenum"]'):
                    shape._element.getparent().remove(shape._element)
        self.pdf=canvas.Canvas(str(output/(NAME+'.pdf')),pagesize=(PAGE_W,PAGE_H))
        self.pdf.setTitle(NAME)
    def box(self,x,y,w,h,fill,border=None,name='Panel'):
        shape=self.slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Pt(x),Pt(y),Pt(w),Pt(h));shape.name=name
        shape.fill.solid();shape.fill.fore_color.rgb=color(fill)
        if border:shape.line.color.rgb=color(border);shape.line.width=Pt(.5)
        else:shape.line.fill.background()
        self.pdf.setFillColorRGB(*pdf_color(fill))
        if border:self.pdf.setStrokeColorRGB(*pdf_color(border));self.pdf.setLineWidth(.5)
        self.pdf.rect(x,PAGE_H-y-h,w,h,fill=1,stroke=int(bool(border)))
        return shape
    def text(self,lines,x,y,w,size=18,font='Arial',fill=WHITE,padding=0):
        shape=self.slide.shapes.add_textbox(Pt(x),Pt(y),Pt(w),Pt((len(lines)-1)*size*1.24+size+6))
        frame=shape.text_frame;frame.word_wrap=False;frame.margin_left=frame.margin_right=Pt(padding);frame.margin_top=frame.margin_bottom=Pt(0)
        asc=pdfmetrics.getAscent(font)*size/1000
        for i,line in enumerate(lines):
            p=frame.paragraphs[0] if i==0 else frame.add_paragraph()
            p.font.name='Georgia' if font.startswith('Georgia') else 'Arial';p.font.size=Pt(size);p.font.italic=font.endswith('Italic');p.font.color.rgb=color(fill)
            p.space_before=p.space_after=Pt(0);p.line_spacing=Pt(size*1.24)
            parts=TERMS.split(line) if font=='Arial' and fill==WHITE else [line];xx=x+padding
            for part in filter(None,parts):
                shade=GOLD if TERMS.fullmatch(part) else fill
                run=p.add_run();run.text=part;run.font.color.rgb=color(shade)
                self.pdf.setFont(font,size);self.pdf.setFillColorRGB(*pdf_color(shade))
                self.pdf.drawString(xx,PAGE_H-y-asc-i*size*1.24,part);xx+=pdfmetrics.stringWidth(part,font,size)
    def connector(self,x1,y1,x2,y2,arrow=True):
        shape=self.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Pt(x1),Pt(y1),Pt(x2),Pt(y2))
        shape.name='Logical connector';shape.line.color.rgb=color(GOLD);shape.line.width=Pt(.8)
        if arrow:
            end=OxmlElement('a:tailEnd');end.set('type','triangle');shape.line._get_or_add_ln().append(end)
        self.pdf.setStrokeColorRGB(*pdf_color(GOLD));self.pdf.setLineWidth(.8)
        self.pdf.line(x1,PAGE_H-y1,x2,PAGE_H-y2)
        if arrow:
            import math
            angle=math.atan2(y2-y1,x2-x1);length=5;spread=2.1
            ex=x2-length*math.cos(angle);ey=y2-length*math.sin(angle)
            path=self.pdf.beginPath();path.moveTo(x2,PAGE_H-y2)
            path.lineTo(ex-spread*math.sin(angle),PAGE_H-ey-spread*math.cos(angle))
            path.lineTo(ex+spread*math.sin(angle),PAGE_H-ey+spread*math.cos(angle));path.close()
            self.pdf.setFillColorRGB(*pdf_color(GOLD));self.pdf.drawPath(path,fill=1,stroke=0)
    def diamond(self,x,y):
        shape=self.slide.shapes.add_shape(MSO_SHAPE.DIAMOND,Pt(x),Pt(y),Pt(5),Pt(5));shape.name='Gold diamond marker'
        shape.fill.solid();shape.fill.fore_color.rgb=color(GOLD);shape.line.fill.background()
        path=self.pdf.beginPath();path.moveTo(x+2.5,PAGE_H-y);path.lineTo(x+5,PAGE_H-y-2.5)
        path.lineTo(x+2.5,PAGE_H-y-5);path.lineTo(x,PAGE_H-y-2.5);path.close()
        self.pdf.setFillColorRGB(*pdf_color(GOLD));self.pdf.drawPath(path,fill=1,stroke=0)
    def panel(self,x,y,w,h,name='Diagram card'):
        shape=self.box(x,y,w,h,CARD,name=name)
        self.box(x+.5,y+.5,w-1,h-1,CARD,GOLD,'Inset gold outline')
        self.blocks.append((x,y,w,h,name))
        return shape
    def quotation(self,unit,y):
        rows=wrap(unit.thesis,'GeorgiaItalic',18,COLUMN-50);height=len(rows)*22.32+28
        self.box(LEFT+.5,y,COLUMN-1,height,BG,GOLD,'Quote')
        self.blocks.append((LEFT+.5,y,COLUMN-1,height,'Quote'))
        self.connector(LEFT+12,y+12,LEFT+12,y+height-12,False)
        self.text(rows,LEFT,y+12,COLUMN,18,'GeorgiaItalic',GOLD,26)
        return height
    def page(self,lesson):
        self.slide=self.prs.slides.add_slide(self.prs.slide_layouts[6]);self.slide.background.fill.solid();self.slide.background.fill.fore_color.rgb=color(BG)
        self.slide._element.set('showMasterSp','0');self.blocks=[]
        self.pdf.setFillColorRGB(*pdf_color(BG));self.pdf.rect(0,0,PAGE_W,PAGE_H,fill=1,stroke=0)
        title=wrap(lesson.title,'Georgia',32,COLUMN-4)
        if len(title)>2:raise ValueError('Shorten this topic heading: '+lesson.title)
        self.text(title,LEFT,24,COLUMN,32,'Georgia',GOLD,2)
        mode,data=infer_layout(lesson);self.layouts.append(mode)
        inline_quote=next((u for u in lesson.units if u.quote and mode=='comparison' and u.thesis==data[1]),None)
        quotes=[u for u in lesson.units if u.quote and u is not inline_quote]
        cards=[u for u in lesson.units if not u.quote]
        quote_heights=[len(wrap(u.thesis,'GeorgiaItalic',18,COLUMN-50))*22.32+28 for u in quotes]
        quote_space=sum(quote_heights)+max(0,len(quotes)-1)*10
        quote_top=TEXT_BOTTOM-quote_space
        content_bottom=quote_top-18 if quotes else TEXT_BOTTOM
        if mode=='comparison':
            gap=38;w=(COLUMN-gap)/2;y=145
            texts=data
            # Duplicate wording remains only in the lower quote, not both columns and quote.
            h=min(250,content_bottom-y)
            for index,value in enumerate(texts):
                x=LEFT+index*(w+gap)
                face='GeorgiaItalic' if index==1 and inline_quote else 'Arial'
                rows=wrap(value,face,18,w-40)
                required=len(rows)*22.32+36
                assert required<=h,('Comparison requires an additional slide',required,h)
                self.panel(x,y,w,h,'Comparison card')
                self.text(rows,x,y+24,w,18,face,GOLD if face=='GeorgiaItalic' else WHITE,padding=20)
                if face=='GeorgiaItalic':self.connector(x+10,y+20,x+10,y+h-20,False)
            self.connector(LEFT+w+5,y+h/2,LEFT+w+gap-5,y+h/2)
        elif mode=='chain':
            # Responsive horizontal chain: long five-step formulas wrap to a second row.
            groups=[data] if len(data)<=3 else [data[:3],data[3:]]
            previous=None
            for row_index,group in enumerate(groups):
                widths=[max(90,pdfmetrics.stringWidth(label,'Arial',16)+24) for label in group] if all(label==label.upper() for label in group) else [(COLUMN-20*(len(group)-1))/len(group)]*len(group)
                assert sum(widths)+20*(len(group)-1)<=COLUMN
                x=LEFT+(COLUMN-sum(widths)-20*(len(group)-1))/2;y=155+row_index*133
                positions=[]
                for label,w in zip(group,widths):
                    rows=wrap(label,'Arial',16,w-24)
                    assert len(rows)<=3
                    self.panel(x,y,w,76,'Chain node');self.text(rows,x,y+(76-(len(rows)*19.84))/2,w,16,fill=GOLD,padding=12)
                    positions.append((x,w));x+=w+20
                for (x1,w1),(x2,w2) in zip(positions,positions[1:]):self.connector(x1+w1+3,y+38,x2-3,y+38)
                if previous:
                    px,pw,py=previous;endx=positions[0][0]+positions[0][1]/2;mid=py+103
                    self.connector(px+pw/2,py+76,px+pw/2,mid,False)
                    self.connector(px+pw/2,mid,endx,mid,False);self.connector(endx,mid,endx,y)
                previous=(positions[-1][0],positions[-1][1],y)
        elif mode=='tiles':
            gap=16;w=(COLUMN-2*gap)/3;y=145
            wrapped=[wrap(value,'Arial',18,w-40) for value in data]
            h=max(len(rows)*22.32+44 for rows in wrapped)
            assert y+h<=content_bottom
            for index,rows in enumerate(wrapped):
                x=LEFT+index*(w+gap);self.panel(x,y,w,h,'Thesis tile')
                for row_index in range(len(rows)):self.diamond(x+12,y+25+row_index*22.32)
                self.text(rows,x,y+18,w,18,padding=26)
        elif mode=='relations':
            y=145
            for left,right in data:
                self.panel(LEFT,y,COLUMN,57,'Relation row')
                self.diamond(LEFT+14,y+25)
                self.text(wrap(left,'Arial',18,210),LEFT+30,y+17,230,18,fill=GOLD)
                self.connector(LEFT+238,y+29,LEFT+282,y+29)
                self.text(wrap(right,'Arial',18,COLUMN-320),LEFT+305,y+17,COLUMN-315,18)
                y+=69
            assert y-12<=content_bottom
        else:
            y=TEXT_TOP
            for unit in cards:
                rows=wrap(unit.thesis,'Arial',18,COLUMN-54);height=len(rows)*22.32+28
                assert y+height<=content_bottom
                panel=self.panel(LEFT,y,COLUMN,height,'Teaching card');panel.left=Inches(4.5);panel.width=Inches(8.3)
                for row_index in range(len(rows)):self.diamond(LEFT+15,y+18+row_index*22.32)
                self.text(rows,LEFT,y+12,COLUMN,18,padding=34)
                y+=height+12
        y=quote_top
        for unit,height in zip(quotes,quote_heights):self.quotation(unit,y);y+=height+10
        # Validate all semantic block rectangles, including comparison columns and nodes.
        for i,(x,y,w,h,name) in enumerate(self.blocks):
            assert x>=LEFT and x+w<=LEFT+COLUMN+.01 and y+h<=TEXT_BOTTOM+.01
            for xx,yy,ww,hh,other in self.blocks[i+1:]:
                overlap=min(x+w,xx+ww)-max(x,xx)>.01 and min(y+h,yy+hh)-max(y,yy)>.01
                assert not overlap,('Blocks overlap',name,other)
        # Explicit full-width, borderless final overlay. No page numbers are generated.
        mask=self.box(0,7*72,13.333*72,.5*72,BG,name='Bottom page-number mask')
        mask.left=Inches(0);mask.top=Inches(7);mask.width=Inches(13.333);mask.height=Inches(.5)
        tree=self.slide.shapes._spTree;tree.remove(mask._element);tree.insert_element_before(mask._element,'p:extLst')
        self.slide.notes_slide.notes_text_frame.text=lesson.notes;self.pdf.showPage()
    def finish(self):self.prs.save(self.output/(NAME+'.pptx'));self.pdf.save()


def validate(output,slides,source):
    import fitz
    deck=Presentation(output/(NAME+'.pptx'))
    assert len(deck.slides)==len(slides)
    assert normalized('\n'.join(s.notes_slide.notes_text_frame.text for s in deck.slides))==normalized('\n'.join(r['notes'] for r in source))
    for page in deck.slides:
        assert not any(s.shape_type==13 for s in page.shapes), 'Text-only slides must contain no pictures.'
        for shape in list(page.shapes)[:-1]:assert shape.left>=Inches(4.5)
        cards=[s for s in page.shapes if s.name=='Teaching card']
        quotes=[s for s in page.shapes if s.name=='Quote']
        for card in cards:
            assert (card.left,card.width)==(Inches(4.5),Inches(8.3))
        for card in cards:
            for quote in quotes:assert card.top+card.height<quote.top
        for shape in page.shapes:
            if shape.has_text_frame:assert not re.fullmatch(r'\d{1,3}\s*/\s*\d{1,3}',shape.text.strip())
        mask=page.shapes[-1];assert mask.name=='Bottom page-number mask'
        assert mask.fill.fore_color.rgb==color(BG)
        assert mask.line._get_or_add_ln().find('{http://schemas.openxmlformats.org/drawingml/2006/main}noFill') is not None
        assert (mask.left,mask.top,mask.width,mask.height)==(Inches(0),Inches(7),Inches(13.333),Inches(.5))
    document=fitz.open(output/(NAME+'.pdf'));assert len(document)==len(slides)
    for page in document:
        assert not page.get_images(), 'Text-only PDF must contain no images.'
        for bounds in (fitz.Rect(0,0,324,540),fitz.Rect(0,504,PAGE_W,540)):
            pix=page.get_pixmap(clip=bounds,alpha=False)
            assert pix.samples==bytes(BG)*pix.width*pix.height
    print(f'Validated {len(slides)} vector-infographic slides: no images; no overlaps; blank video zone; final masks verified.',flush=True)


def publish(home,output):
    if home!=output.resolve():raise ValueError('--push requires output in the repository containing this script.')
    names=[NAME+'.pptx',NAME+'.pdf','generate_lesson.py','lesson_content.json','requirements-lesson.txt','.gitignore','README.md']
    subprocess.run(['git','add','--',*names],cwd=home,check=True)
    state=subprocess.run(['git','diff','--cached','--quiet'],cwd=home).returncode
    if state==1:subprocess.run(['git','commit','-m','Add adaptive vector infographics, comparison columns and logical chains'],cwd=home,check=True)
    elif state:raise RuntimeError('Cannot inspect the staged files.')
    subprocess.run(['git','push','origin','HEAD:main'],cwd=home,check=True)
    local=subprocess.check_output(['git','rev-parse','HEAD'],cwd=home,text=True).strip()
    remote=subprocess.check_output(['git','ls-remote','origin','refs/heads/main'],cwd=home,text=True).split()[0]
    assert local==remote;print('Push verified:',local,flush=True)


def main():
    home=Path(__file__).resolve().parent
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input',type=Path,default=home/'lesson_content.json')
    parser.add_argument('--output-dir',type=Path,default=home)
    parser.add_argument('--fonts-dir',type=Path,default=home/'work'/'fonts')
    parser.add_argument('--max-slides',type=int,default=650)
    parser.add_argument('--push',action='store_true')
    args=parser.parse_args();args.output_dir.mkdir(parents=True,exist_ok=True);args.fonts_dir.mkdir(parents=True,exist_ok=True)
    # Font files are local. This generator makes no HTTP requests.
    for family,filename in FONT_FILES.items():
        candidates=[args.fonts_dir/filename,Path('/usr/share/fonts/truetype/msttcorefonts')/filename,
                    Path('C:/Windows/Fonts')/filename,Path('/Library/Fonts')/filename]
        path=next((candidate for candidate in candidates if candidate.is_file()),None)
        if path is None:raise FileNotFoundError(f'Local font {filename} not found; supply --fonts-dir.')
        pdfmetrics.registerFont(TTFont(family,str(path)))
    source=load_content(args.input);slides=paginate(source,args.max_slides)
    print(f'Automatically planned {len(slides)} slides; capacity limit {args.max_slides}; no duplicate padding.',flush=True)
    deck=Deck(args.output_dir)
    for index,slide in enumerate(slides,1):
        deck.page(slide)
        print(f'Rendered text slide {index}/{len(slides)}',flush=True)
    deck.finish();validate(args.output_dir,slides,source)
    from collections import Counter
    print('Automatic layouts:',dict(Counter(deck.layouts)),flush=True)
    if args.push:publish(home,args.output_dir)

if __name__=='__main__':main()
