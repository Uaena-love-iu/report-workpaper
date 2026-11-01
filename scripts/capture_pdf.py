"""Capture genuine PDF pixels with exact-text red boxes; no text substitution."""
from pathlib import Path
from datetime import date
import argparse,json,re,hashlib

def capture(file,page_number,phrases,out,context=32,scale=2.5):
    path=Path(file)
    with path.open('rb')as f:
        if not f.read(5).startswith(b'%PDF-'):raise ValueError('Not a PDF; check for an HTML error page')
    import pdfplumber,pypdfium2 as pdfium
    from PIL import ImageDraw
    with pdfplumber.open(path)as doc:
        if not 1<=page_number<=len(doc.pages):raise ValueError('Physical page outside PDF')
        page=doc.pages[page_number-1];chars=[c for c in page.chars for t in re.sub(r'\s+','',c['text'])];text=''.join(re.sub(r'\s+','',c['text'])for c in page.chars);groups=[]
        for phrase in phrases:
            needle=re.sub(r'\s+','',phrase);start=text.find(needle)
            if not needle or start<0:raise ValueError('Exact PDF phrase not found: '+phrase)
            groups.append(chars[start:start+len(needle)])
        selected=[c for group in groups for c in group];left=max(0,min(c['x0']for c in selected)-context);right=min(page.width,max(c['x1']for c in selected)+context);top=max(0,min(c['top']for c in selected)-context);bottom=min(page.height,max(c['bottom']for c in selected)+context)
        if right-left<420:center=(left+right)/2;left=max(0,min(center-210,page.width-420));right=min(page.width,left+420)
        if (right-left)*(bottom-top)*scale**2>35_000_000:raise ValueError('Capture too large; split non-adjacent phrases')
    pdf=pdfium.PdfDocument(path)
    try:
        pdfpage=pdf[page_number-1]
        try:bitmap=pdfpage.render(scale=scale);image=bitmap.to_pil().convert('RGB')
        finally:pdfpage.close()
        image=image.crop(tuple(round(v*scale)for v in [left,top,right,bottom]));draw=ImageDraw.Draw(image)
        for group in groups:
            lines={}
            for c in group:lines.setdefault(round(c['top']/4)*4,[]).append(c)
            for line in lines.values():draw.rectangle([(min(c['x0']for c in line)-left-1)*scale,(min(c['top']for c in line)-top-1)*scale,(max(c['x1']for c in line)-left+1)*scale,(max(c['bottom']for c in line)-top+1)*scale],outline='#D00000',width=3)
        out=Path(out);out.parent.mkdir(parents=True,exist_ok=True);image.save(out)
    finally:pdf.close()
    return {'ok':True,'kind':'pdf_crop','pdf_page':page_number,'matched':phrases,'image':str(out),'screenshot_sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'original_pdf_sha256':hashlib.sha256(path.read_bytes()).hexdigest()}

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('pdf');p.add_argument('--page',type=int,required=True);p.add_argument('--printed-page',required=True,help='Actual printed page, or none');p.add_argument('--highlight',action='append',required=True);p.add_argument('--out',required=True);p.add_argument('--url',required=True);p.add_argument('--title',required=True);p.add_argument('--context',type=int,default=32);a=p.parse_args()
    try:
        result=capture(a.pdf,a.page,a.highlight,a.out,a.context);result.update(printed_page=None if a.printed_page=='none'else a.printed_page,url=a.url,title=a.title,collected_at=date.today().isoformat())
        Path(a.out).with_suffix('.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8');print(json.dumps(result,ensure_ascii=False))
    except (ValueError,OSError)as e:p.exit(1,str(e)+'\n')
if __name__=='__main__':main()
