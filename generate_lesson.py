#!/usr/bin/env python3
"""Clean slide generator: JSON narration -> editable PPTX + matching PDF.

python -m pip install -r requirements-lesson.txt
python generate_lesson.py --fonts-dir /path/to/fonts
python generate_lesson.py --fonts-dir /path/to/fonts --push
The push uses an explicit lease and never replaces unseen remote commits.
"""
from __future__ import annotations
import argparse
from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import re
import subprocess
from urllib.request import urlopen

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

TITLE = 'Ваш Кристалл незабываемости'
BACKGROUND = (26, 26, 26)
PANEL = (40, 40, 40)
GOLD = (212, 175, 55)
WHITE = (255, 255, 255)
PAGE_WIDTH, PAGE_HEIGHT = 13.333 * 72, 7.5 * 72
CONTENT_LEFT, CONTENT_WIDTH = 4.5 * 72, 8.0 * 72
SOURCE_HASH = '618f688756018c7010acb9c11e7ecc09340c80c0537e3d2ecc10cba238f96ef4'
KEYWORDS = re.compile(r'(\b(?:следы|след|лёгкость|память|уважение|вдохновение|комбинацией|незабываемая|незабываемые)\b)', re.I)

@dataclass
class Element:
    kind: str
    x: float
    y: float
    width: float
    height: float
    name: str
    text: str = ''
    font: str = 'Arial'
    size: int = 18
    color: tuple = WHITE
    fill: tuple = BACKGROUND
    stroke: tuple | None = None
    weight: float = .5
    opacity: float = 1.0
    inset: float = 0.0


def prepare_fonts(directory: Path):
    directory.mkdir(parents=True, exist_ok=True)
    files = {'Georgia': 'georgia.ttf', 'GeorgiaItalic': 'georgiai.ttf', 'Arial': 'arial.ttf'}
    remote = {'Georgia': 'Georgia.ttf', 'GeorgiaItalic': 'Georgia%20Italic.ttf', 'Arial': 'Arial.ttf'}
    for family, filename in files.items():
        local = directory / filename
        if not local.exists():
            url = 'https://raw.githubusercontent.com/justrajdeep/fonts/master/' + remote[family]
            with urlopen(url, timeout=40) as response:
                local.write_bytes(response.read())
        pdfmetrics.registerFont(TTFont(family, str(local)))


def linebreaks(text, font, size, available):
    result = []
    for paragraph in text.split('\n'):
        current = ''
        for word in paragraph.split():
            candidate = (current + ' ' + word).strip()
            if current and pdfmetrics.stringWidth(candidate, font, size) > available:
                result.append(current)
                current = word
            else:
                current = candidate
        result.append(current)
    assert all(pdfmetrics.stringWidth(t, font, size) <= available for t in result)
    return result


def compose(item):
    elements = []
    def put(kind, x, y, w, h, name, **style):
        elements.append(Element(kind, x, y, w, h, name, **style))
    def label(value, y, size=18, font='Arial', color=WHITE, inset=2):
        rows = linebreaks(value, font, size, CONTENT_WIDTH - 2 * inset)
        put('text', CONTENT_LEFT, y, CONTENT_WIDTH, len(rows)*size*1.24+4,
            'Text', text='\n'.join(rows), font=font, size=size, color=color)
        # The text box remains exactly eight inches wide; inset is an internal margin.
        elements[-1].inset = inset
        return len(rows)*size*1.24
    # Texture remains on the right; the divider at 4.3 inches is an explicit exception.
    points = [(833,428), (900,397), (946,447), (928,493), (866,499), (821,479)]
    for a,b in list(zip(points,points[1:]+points[:1])) + [(points[0],points[2]),(points[1],points[3]),(points[2],points[5])]:
        put('line', a[0],a[1],b[0]-a[0],b[1]-a[1], 'Crystal texture', color=GOLD, weight=.3, opacity=.10)
    put('line', 4.3*72, 1.0*72, 0, 5.5*72, 'Video divider', color=GOLD, weight=.65)
    heading_height = label(item['title'], 36, 34, 'Georgia', GOLD)
    assert heading_height <= 34*1.24*2
    y = max(145, 36+heading_height+24)
    for number, text in enumerate(item.get('cards', []), 1):
        rows = linebreaks(text,'Arial',18,CONTENT_WIDTH-48)
        assert len(rows) <= 3, 'Split narration into another slide; never reduce its font.'
        height = len(rows)*18*1.24+24
        put('rect',CONTENT_LEFT,y,CONTENT_WIDTH,height,f'Card {number}',fill=PANEL,stroke=GOLD,weight=.5)
        put('diamond',CONTENT_LEFT+12,y+17,4,4,'Gold bullet',fill=GOLD)
        label('\n'.join(rows),y+11,color=GOLD if item.get('accent_all') else WHITE,inset=24)
        y += height+10
    for number, (left,right) in enumerate(item.get('table', []), 1):
        put('rect',CONTENT_LEFT,y,CONTENT_WIDTH,57,f'Card table {number}',fill=PANEL,stroke=GOLD,weight=.5)
        label(left,y+17,color=GOLD,inset=14)
        # Independent internal text cell; the enclosing content block is eight inches wide.
        put('text',CONTENT_LEFT+226,y+17,CONTENT_WIDTH-240,28,'Table value',text='→ '+right)
        elements[-1].inset = 0
        put('line',CONTENT_LEFT+200,y+12,0,33,'Table separator',color=GOLD,opacity=.45)
        y += 67
    if item.get('quote'):
        y += 8
        rows = linebreaks(item['quote'],'GeorgiaItalic',18,CONTENT_WIDTH-28)
        height = len(rows)*18*1.24+24
        put('rect',CONTENT_LEFT+.5,y,CONTENT_WIDTH-1,height,'Author quote',stroke=GOLD,weight=.75)
        label('\n'.join(rows),y+11,18,'GeorgiaItalic',GOLD,14)
        y += height
    if item.get('photo_prompt'):
        y = max(352,y+16)
        rows = linebreaks(item['photo_prompt'],'Arial',16,CONTENT_WIDTH-28)
        height = 58+len(rows)*16*1.24
        assert y+height <= 501, 'Photo instructions must stay above the bottom mask.'
        put('rect',CONTENT_LEFT+.5,y,CONTENT_WIDTH-1,height,'Photo prompt container',stroke=GOLD,weight=.65)
        label('ПРОМПТ ДЛЯ MIDJOURNEY',y+12,16,color=GOLD,inset=14)
        label('\n'.join(rows),y+43,16,inset=14)
    assert len(item.get('cards',item.get('table',[]))) <= 4
    assert all(e.y+e.height <= 504 for e in elements if e.kind != 'line')
    # LAST ELEMENT: masks the full bottom 0.5 inches, including any inherited footer.
    put('rect',0,7.0*72,13.333*72,.5*72,'Bottom page-number mask',fill=BACKGROUND)
    return elements


def draw_slide(prs, document, item, elements):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(*BACKGROUND)
    document.setFillColorRGB(*(v/255 for v in BACKGROUND))
    document.rect(0,0,PAGE_WIDTH,PAGE_HEIGHT,fill=1,stroke=0)
    for e in elements:
        document.saveState()
        document.setFillAlpha(e.opacity);document.setStrokeAlpha(e.opacity)
        if e.kind == 'text':
            inset = getattr(e,'inset',0)
            box = slide.shapes.add_textbox(Pt(e.x),Pt(e.y),Pt(e.width),Pt(e.height))
            frame=box.text_frame;frame.word_wrap=False
            frame.margin_left=frame.margin_right=Pt(inset)
            frame.margin_top=frame.margin_bottom=Pt(0)
            for i,row in enumerate(e.text.split('\n')):
                paragraph=frame.paragraphs[0] if i==0 else frame.add_paragraph()
                paragraph.font.name='Georgia' if e.font.startswith('Georgia') else 'Arial'
                paragraph.font.size=Pt(e.size);paragraph.font.italic=e.font.endswith('Italic')
                paragraph.space_before=paragraph.space_after=Pt(0)
                paragraph.line_spacing=Pt(e.size*1.24)
                cursor=e.x+inset
                parts=KEYWORDS.split(row) if e.font=='Arial' and e.color==WHITE else [row]
                for part in filter(None,parts):
                    shade=GOLD if KEYWORDS.fullmatch(part) else e.color
                    run=paragraph.add_run();run.text=part;run.font.color.rgb=RGBColor(*shade)
                    document.setFont(e.font,e.size);document.setFillColorRGB(*(v/255 for v in shade))
                    baseline=PAGE_HEIGHT-e.y-pdfmetrics.getAscent(e.font)*e.size/1000-i*e.size*1.24
                    document.drawString(cursor,baseline,part)
                    cursor+=pdfmetrics.stringWidth(part,e.font,e.size)
        elif e.kind == 'line':
            box=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Pt(e.x),Pt(e.y),Pt(e.x+e.width),Pt(e.y+e.height))
            if e.name=='Video divider':
                box.left=Inches(4.3);box.top=Inches(1.0);box.width=Inches(0);box.height=Inches(5.5)
            box.line.color.rgb=RGBColor(*e.color);box.line.width=Pt(e.weight)
            if e.opacity<1:
                solid=box.line._get_or_add_ln().find('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
                alpha=OxmlElement('a:alpha');alpha.set('val',str(int(e.opacity*100000)));solid[0].append(alpha)
            document.setStrokeColorRGB(*(v/255 for v in e.color));document.setLineWidth(e.weight)
            document.line(e.x,PAGE_HEIGHT-e.y,e.x+e.width,PAGE_HEIGHT-e.y-e.height)
        else:
            kind=MSO_SHAPE.DIAMOND if e.kind=='diamond' else MSO_SHAPE.RECTANGLE
            box=slide.shapes.add_shape(kind,Pt(e.x),Pt(e.y),Pt(e.width),Pt(e.height))
            box.fill.solid();box.fill.fore_color.rgb=RGBColor(*e.fill)
            if e.stroke:
                box.line.color.rgb=RGBColor(*e.stroke);box.line.width=Pt(e.weight)
            else:box.line.fill.background()
            document.setFillColorRGB(*(v/255 for v in e.fill))
            if e.stroke:
                document.setStrokeColorRGB(*(v/255 for v in e.stroke));document.setLineWidth(e.weight)
            if e.kind=='diamond':
                path=document.beginPath();cx=e.x+e.width/2;cy=PAGE_HEIGHT-e.y-e.height/2
                path.moveTo(cx,cy+2);path.lineTo(cx+2,cy);path.lineTo(cx,cy-2);path.lineTo(cx-2,cy);path.close()
                document.drawPath(path,fill=1,stroke=0)
            else:document.rect(e.x,PAGE_HEIGHT-e.y-e.height,e.width,e.height,fill=1,stroke=int(bool(e.stroke)))
        box.name=e.name
        document.restoreState()
    slide.notes_slide.notes_text_frame.text=item['notes']
    document.showPage()


def verify(output, content):
    import fitz
    deck=Presentation(output/(TITLE+'.pptx'))
    assert len(deck.slides)==len(content)==16
    merged=' '.join(' '.join(d['notes'].split()) for d in content)
    assert hashlib.sha256(merged.encode()).hexdigest()==SOURCE_HASH
    for slide,item in zip(deck.slides,content):
        assert slide.notes_slide.notes_text_frame.text==item['notes']
        mask=slide.shapes[-1]
        assert mask.name=='Bottom page-number mask'
        assert (mask.left,mask.top,mask.width,mask.height)==(Inches(0),Inches(7),Inches(13.333),Inches(.5))
        assert mask.fill.fore_color.rgb==RGBColor(*BACKGROUND)
        divider=next(s for s in slide.shapes if s.name=='Video divider')
        assert (divider.left,divider.top,divider.height)==(Inches(4.3),Inches(1),Inches(5.5))
        for shape in slide.shapes:
            if shape.name.startswith('Card'):
                assert shape.fill.fore_color.rgb==RGBColor(*PANEL)
                assert shape.line.color.rgb==RGBColor(*GOLD)
            if shape.has_text_frame and shape.text:
                assert shape.left>=Inches(4.5)
                assert not re.fullmatch(r'\d{2} / \d{2}',shape.text.strip())
    pdf=fitz.open(output/(TITLE+'.pdf'))
    assert len(pdf)==16
    for page in pdf:
        for area in (fitz.Rect(0,504,PAGE_WIDTH,540),fitz.Rect(0,0,308,540)):
            pix=page.get_pixmap(clip=area,alpha=False)
            assert pix.samples==bytes(BACKGROUND)*pix.width*pix.height
    print('Verified: 16 slides; last-layer bottom masks; exact 4.3-inch dividers; framed cards; unchanged Notes.')


def publish(output):
    tracked=[TITLE+'.pptx',TITLE+'.pdf','generate_lesson.py','lesson_content.json','requirements-lesson.txt']
    expected=subprocess.check_output(['git','rev-parse','refs/remotes/origin/main'],cwd=output,text=True).strip()
    subprocess.run(['git','add','--',*tracked],cwd=output,check=True)
    if subprocess.run(['git','diff','--cached','--quiet'],cwd=output).returncode:
        subprocess.run(['git','commit','-m','Replace generator from scratch; add footer masks, divider and framed cards'],cwd=output,check=True)
    subprocess.run(['git','push','--force-with-lease=refs/heads/main:'+expected,'origin','HEAD:main'],cwd=output,check=True)
    remote=subprocess.check_output(['git','ls-remote','origin','refs/heads/main'],cwd=output,text=True).split()[0]
    local=subprocess.check_output(['git','rev-parse','HEAD'],cwd=output,text=True).strip()
    assert local==remote
    print('Force-with-lease push verified:',local)


def main():
    home=Path(__file__).resolve().parent
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir',type=Path,default=home)
    parser.add_argument('--fonts-dir',type=Path,default=home/'work'/'fonts')
    parser.add_argument('--push',action='store_true')
    args=parser.parse_args()
    content=json.loads((home/'lesson_content.json').read_text(encoding='utf-8'))
    args.output_dir.mkdir(parents=True,exist_ok=True)
    prepare_fonts(args.fonts_dir)
    prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5)
    prs.core_properties.title=TITLE;prs.core_properties.language='ru-RU'
    document=canvas.Canvas(str(args.output_dir/(TITLE+'.pdf')),pagesize=(PAGE_WIDTH,PAGE_HEIGHT))
    document.setTitle(TITLE)
    for item in content:draw_slide(prs,document,item,compose(item))
    prs.save(args.output_dir/(TITLE+'.pptx'));document.save()
    verify(args.output_dir,content)
    if args.push:publish(args.output_dir)

if __name__=='__main__':main()
