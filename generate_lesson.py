#!/usr/bin/env python3
"""Build a premium illustrated lesson, from a short lesson to a 650-slide program.

Install: python -m pip install -r requirements-lesson.txt
Build:   python generate_lesson.py --fonts-dir /path/to/fonts
Publish: python generate_lesson.py --fonts-dir /path/to/fonts --push
Input:   --input lesson_content.json or --input lecture.md
Photos:  --provider catalog (no account), or --provider unsplash with a secret
         UNSPLASH_ACCESS_KEY and permitted api.unsplash.com/images.unsplash.com.
No slides are duplicated or padded to reach 650. Narration is divided by meaning
and available space; cached photos and search results keep large builds practical.
"""
from __future__ import annotations
import argparse
from collections import Counter
from dataclasses import dataclass
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import re
import subprocess

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

NAME='Ваш Кристалл незабываемости'
BG=(26,26,26); GOLD=(212,175,55); WHITE=(255,255,255); CARD=(40,40,40)
PAGE_W=13.333*72; PAGE_H=7.5*72
LEFT=4.5*72; COLUMN=4.2*72; IMAGE_LEFT=9.0*72
IMAGE_TOP=1.5*72; IMAGE_W=3.5*72; IMAGE_H=5.0*72
TEXT_TOP=130; TEXT_BOTTOM=488
TERMS=re.compile(r'(\b(?:следы|след|лёгкость|память|уважение|вдохновение|комбинацией|незабываемая|незабываемые)\b)',re.I)
CATEGORIES={
 'crystal':(r'гран|кристалл|комбинац|многогран','luxury black crystal faceted abstract'),
 'elegance':(r'красот|красив|эстетик|восхищ|женствен','elegant minimal gold abstract architecture'),
 'lightness':(r'лёгк|легк|тепло|нежн|вдохнов','soft light mist forest minimal photography'),
 'memory':(r'памят|след|вспомин|незабыва|впечатл','quiet water reflection golden light'),
 'intellect':(r'интеллект|умн|глубок|разговар','minimal elegant study desk still life'),
 'individuality':(r'индивидуал|уникальн|своим миром|любопыт','quiet luxury travel camera personal world'),
 'contrast':(r'контраст|неожидан|новизн|характер','minimal monochrome abstract contrast'),
 'dignity':(r'достоин|уважен|сильн|спокойн','timeless gold architecture quiet luxury'),
}
FONT_FILES={'Georgia':('georgia.ttf','Georgia.ttf'), 'GeorgiaItalic':('georgiai.ttf','Georgia%20Italic.ttf'), 'Arial':('arial.ttf','Arial.ttf')}


def digest(data):return hashlib.sha256(data).hexdigest()
def normalized(text):return ' '.join(text.split())
def color(value):return RGBColor(*value)
def pdf_color(value):return tuple(v/255 for v in value)


def http_session():
    session=requests.Session()
    session.headers['User-Agent']='PremiumLessonBuilder/1.0 (photo export; cached downloads)'
    session.verify=os.environ.get('REQUESTS_CA_BUNDLE') or os.environ.get('SSL_CERT_FILE') or True
    retry=Retry(total=3,backoff_factor=.6,status_forcelist=(429,500,502,503,504),allowed_methods=('GET',),respect_retry_after_header=True)
    session.mount('https://',HTTPAdapter(max_retries=retry))
    return session


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
                quote=any(word in fragment for word in ('А есть незабываемые.','«хорошая женщина»','Чем больше раскрытых граней','Вот такую женщину'))
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


def category(slide):
    text=(slide.title+' '+slide.notes).lower()
    weights={name:len(re.findall(pattern,text)) for name,(pattern,query) in CATEGORIES.items()}
    return max(weights,key=weights.get) if max(weights.values()) else 'elegance'


class StockPhotos:
    def __init__(self,home,cache,provider):
        self.session=http_session();self.cache=cache;cache.mkdir(parents=True,exist_ok=True)
        self.catalog=json.loads((home/'photo_catalog.json').read_text(encoding='utf-8'))
        self.key=os.environ.get('UNSPLASH_ACCESS_KEY','')
        self.provider=('unsplash' if self.key else 'catalog') if provider=='auto' else provider
        if self.provider=='unsplash' and not self.key:raise ValueError('Set UNSPLASH_ACCESS_KEY securely; never commit it.')
        self.count=Counter();self.searches={};self.previous=None
    def search(self,tag):
        query=CATEGORIES[tag][1]
        if tag in self.searches:return self.searches[tag]
        if self.provider=='catalog':
            matches=[p for p in self.catalog if tag in p['tags']]
            results=matches or [p for p in self.catalog if 'elegance' in p['tags']]
        else:
            cached=self.cache/('search-'+digest(query.encode())+'.json')
            if cached.exists():payload=json.loads(cached.read_text())
            else:
                response=self.session.get('https://api.unsplash.com/search/photos',params={'query':query,'orientation':'portrait','per_page':30},headers={'Authorization':'Client-ID '+self.key},timeout=(10,40))
                response.raise_for_status();payload=response.json();cached.write_text(json.dumps(payload),encoding='utf-8')
            results=[dict(id=p['id'],url=p['urls']['regular'],credit=p['user']['name']+' / Unsplash',source=p['links']['html'],license='https://unsplash.com/license',tracking=p['links']['download_location'],focal=(.5,.5)) for p in payload['results']]
            if not results:raise RuntimeError('No stock photo matches '+query)
        self.searches[tag]=results;return results
    def obtain(self,tag):
        choices=self.search(tag);index=self.count[tag]%len(choices)
        if len(choices)>1 and choices[index]['id']==self.previous:index=(index+1)%len(choices)
        self.count[tag]=index+1;photo=dict(choices[index]);self.previous=photo['id']
        destination=self.cache/(digest(photo['url'].encode())+'.image')
        if not destination.exists():
            if photo.get('tracking'):
                tracked=self.session.get(photo['tracking'],headers={'Authorization':'Client-ID '+self.key},timeout=(10,40));tracked.raise_for_status()
            response=self.session.get(photo['url'],timeout=(10,40),stream=True);response.raise_for_status()
            stream=BytesIO()
            for block in response.iter_content(65536):
                stream.write(block)
                if stream.tell()>20*1024*1024:raise ValueError('Stock image exceeds the 20 MiB download limit.')
            image=Image.open(BytesIO(stream.getvalue()));image.load()
            if min(image.size)<450:raise ValueError('Stock image resolution is insufficient.')
            temporary=destination.with_suffix('.part');temporary.write_bytes(stream.getvalue());temporary.replace(destination)
        raw=destination.read_bytes();image=Image.open(BytesIO(raw));image=ImageOps.exif_transpose(image).convert('RGB')
        photo['original_size']=image.size;photo['original_sha256']=digest(raw)
        # Only a centered/focal-point crop for the 3.5 x 5 inch frame; no retouching.
        width,height=image.size;ratio=IMAGE_W/IMAGE_H
        crop_w=min(width,height*ratio);crop_h=crop_w/ratio
        fx,fy=photo.get('focal',(.5,.5));x=max(0,min(width-crop_w,width*fx-crop_w/2));y=max(0,min(height-crop_h,height*fy-crop_h/2))
        image=image.crop((round(x),round(y),round(x+crop_w),round(y+crop_h)))
        if image.height>1400:image.thumbnail((980,1400),Image.Resampling.LANCZOS)
        output=BytesIO();image.save(output,format='JPEG',quality=91,optimize=True)
        photo['query']=CATEGORIES[tag][1];photo['category']=tag
        photo['provider']=self.provider;photo['embedded_sha256']=digest(output.getvalue())
        return output.getvalue(),photo


class Deck:
    def __init__(self,output):
        self.output=output;self.prs=Presentation()
        self.prs.slide_width=Inches(13.333);self.prs.slide_height=Inches(7.5)
        self.prs.core_properties.title=NAME;self.prs.core_properties.language='ru-RU'
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
        shape=self.slide.shapes.add_textbox(Pt(x),Pt(y),Pt(w),Pt(len(lines)*size*1.24+5))
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
    def page(self,lesson,image):
        self.slide=self.prs.slides.add_slide(self.prs.slide_layouts[6]);self.slide.background.fill.solid();self.slide.background.fill.fore_color.rgb=color(BG)
        self.pdf.setFillColorRGB(*pdf_color(BG));self.pdf.rect(0,0,PAGE_W,PAGE_H,fill=1,stroke=0)
        title=wrap(lesson.title,'Georgia',32,8*72-4)
        if len(title)>2:raise ValueError('Shorten this topic heading: '+lesson.title)
        self.text(title,LEFT,28,8*72,32,'Georgia',GOLD,2)
        self.slide.shapes.add_picture(BytesIO(image),Inches(9),Inches(1.5),width=Inches(3.5),height=Inches(5)).name='Embedded stock photo'
        self.pdf.drawImage(ImageReader(BytesIO(image)),IMAGE_LEFT,PAGE_H-IMAGE_TOP-IMAGE_H,IMAGE_W,IMAGE_H)
        y=TEXT_TOP
        for unit in lesson.units:
            font='GeorgiaItalic' if unit.quote else 'Arial';rows=wrap(unit.thesis,font,18,COLUMN-28);height=len(rows)*22.32+26
            self.box(LEFT+.5,y,COLUMN-1,height,BG if unit.quote else CARD,GOLD,'Quote' if unit.quote else 'Teaching card')
            self.text(rows,LEFT,y+11,COLUMN,18,font,GOLD if unit.quote else WHITE,14);y+=height+10
        assert y-10<=TEXT_BOTTOM
        # Explicit full-width, borderless final overlay. No page numbers are generated.
        mask=self.box(0,7*72,13.333*72,.5*72,BG,name='Bottom page-number mask')
        mask.left=Inches(0);mask.top=Inches(7);mask.width=Inches(13.333);mask.height=Inches(.5)
        self.slide.notes_slide.notes_text_frame.text=lesson.notes;self.pdf.showPage()
    def finish(self):self.prs.save(self.output/(NAME+'.pptx'));self.pdf.save()


def validate(output,slides,source):
    import fitz
    deck=Presentation(output/(NAME+'.pptx'))
    assert len(deck.slides)==len(slides)
    assert normalized('\n'.join(s.notes_slide.notes_text_frame.text for s in deck.slides))==normalized('\n'.join(r['notes'] for r in source))
    for page in deck.slides:
        photos=[s for s in page.shapes if s.shape_type==13]
        assert len(photos)==1 and photos[0].image.blob
        assert (photos[0].left,photos[0].top,photos[0].width,photos[0].height)==(Inches(9),Inches(1.5),Inches(3.5),Inches(5))
        for shape in list(page.shapes)[:-1]:assert shape.left>=Inches(4.5)
        mask=page.shapes[-1];assert mask.name=='Bottom page-number mask'
        assert (mask.left,mask.top,mask.width,mask.height)==(Inches(0),Inches(7),Inches(13.333),Inches(.5))
    document=fitz.open(output/(NAME+'.pdf'));assert len(document)==len(slides)
    for page in document:
        assert page.get_images(),'The PDF must contain an actual embedded image on every slide.'
        for bounds in (fitz.Rect(0,0,324,540),fitz.Rect(0,504,PAGE_W,540)):
            pix=page.get_pixmap(clip=bounds,alpha=False)
            assert pix.samples==bytes(BG)*pix.width*pix.height
    print(f'Validated {len(slides)} illustrated slides: original narration preserved; video zone blank; bottom masks last.',flush=True)


def publish(home,output):
    if home!=output.resolve():raise ValueError('--push requires output in the repository containing this script.')
    names=[NAME+'.pptx',NAME+'.pdf','generate_lesson.py','lesson_content.json','photo_catalog.json','image_sources.json','PHOTO_CREDITS.md','requirements-lesson.txt','.gitignore','README.md']
    subprocess.run(['git','add','--',*names],cwd=home,check=True)
    state=subprocess.run(['git','diff','--cached','--quiet'],cwd=home).returncode
    if state==1:subprocess.run(['git','commit','-m','Replace generator with automatic stock photos, caching and synchronized pagination'],cwd=home,check=True)
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
    parser.add_argument('--cache-dir',type=Path,default=home/'work'/'photo-cache')
    parser.add_argument('--provider',choices=('auto','catalog','unsplash'),default='auto')
    parser.add_argument('--max-slides',type=int,default=650)
    parser.add_argument('--push',action='store_true')
    args=parser.parse_args();args.output_dir.mkdir(parents=True,exist_ok=True);args.fonts_dir.mkdir(parents=True,exist_ok=True)
    session=http_session()
    for family,(filename,remote) in FONT_FILES.items():
        path=args.fonts_dir/filename
        if not path.exists():
            response=session.get('https://raw.githubusercontent.com/justrajdeep/fonts/master/'+remote,timeout=(10,40));response.raise_for_status();path.write_bytes(response.content)
        pdfmetrics.registerFont(TTFont(family,str(path)))
    source=load_content(args.input);slides=paginate(source,args.max_slides)
    print(f'Automatically planned {len(slides)} slides; capacity limit {args.max_slides}; no duplicate padding.',flush=True)
    photos=StockPhotos(home,args.cache_dir,args.provider);deck=Deck(args.output_dir);manifest=[]
    for index,slide in enumerate(slides,1):
        raw,photo=photos.obtain(category(slide));deck.page(slide,raw)
        manifest.append(dict(slide=index,title=slide.title,notes_sha256=digest(slide.notes.encode()),**photo))
        print(f'Rendered {index}/{len(slides)}: {photo["id"]} ({photo["provider"]})',flush=True)
    deck.finish();validate(args.output_dir,slides,source)
    (args.output_dir/'image_sources.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    credits=['# Photo sources','', 'Original photographs are embedded in the PPTX/PDF. Only a framing crop is applied.', 'The catalog uses publicly available photo assets from the following open-source template collections; source and license links are retained below.', 'For Unsplash API results, the individual photographer, photo page and Unsplash license are recorded in image_sources.json.','']
    unique={p['id']:p for p in manifest}
    for p in unique.values():credits.extend([f'## {p["id"]}',f'- Credit: {p["credit"]}',f'- Source: {p.get("source",p["url"])}',f'- License: {p["license"]}',''])
    (args.output_dir/'PHOTO_CREDITS.md').write_text('\n'.join(credits),encoding='utf-8')
    print(f'Embedded {len(manifest)} photos, {len(unique)} distinct sources. Originals cached for reuse.',flush=True)
    if args.push:publish(home,args.output_dir)

if __name__=='__main__':main()
