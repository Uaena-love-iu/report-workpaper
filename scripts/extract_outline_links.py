"""Extract DOCX paragraph links and sparse chart points without executing content."""
from zipfile import ZipFile
from xml.etree import ElementTree as E
from pathlib import Path
import argparse,json,re
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';C='http://schemas.openxmlformats.org/drawingml/2006/chart';N={'w':W,'c':C}
def sparse(cache):
    if cache is None:return []
    count=cache.find('c:ptCount',N);points=cache.findall('c:pt',N);size=max([int(count.get('val'))if count is not None else 0]+[int(p.get('idx'))+1 for p in points]);values=[None]*size
    for pt in points:
        value=pt.find('c:v',N);values[int(pt.get('idx'))]=value.text if value is not None else None
    return values
def extract(file,include_charts=False):
    result={'paragraphs':[],'charts':[]}
    with ZipFile(file)as z:
        refs={r.get('Id'):r.get('Target')for r in E.fromstring(z.read('word/_rels/document.xml.rels'))}if'word/_rels/document.xml.rels'in z.namelist()else{}
        for i,p in enumerate(E.fromstring(z.read('word/document.xml')).findall('.//w:p',N),1):
            text=''.join(t.text or''for t in p.findall('.//w:t',N));links=[]
            for h in p.findall('.//w:hyperlink',N):
                url=refs.get(h.get('{'+R+'}id'))
                if url:links.append({'title':''.join(t.text or''for t in h.findall('.//w:t',N)),'url':url})
            for field in p.findall('.//w:fldSimple',N):
                match=re.search(r'HYPERLINK\s+"([^"]+)"',field.get('{'+W+'}instr',''),re.I)
                if match:links.append({'title':''.join(t.text or''for t in field.findall('.//w:t',N)),'url':match.group(1)})
            if text or links:result['paragraphs'].append({'paragraph':i,'text':text,'links':links})
        if include_charts:
            for name in sorted(z.namelist()):
                if not name.startswith('word/charts/chart')or not name.endswith('.xml'):continue
                chart=[]
                for series in E.fromstring(z.read(name)).findall('.//c:ser',N):
                    title=series.find('c:tx/c:strRef/c:strCache',N);title_values=sparse(title);literal=series.find('c:tx/c:v',N)
                    def cache(parent):
                        for kind in ['c:numRef/c:numCache','c:strRef/c:strCache','c:numLit','c:strLit']:
                            found=series.find(parent+'/'+kind,N)
                            if found is not None:return found
                        return None
                    chart.append({'name':next((x for x in title_values if x is not None),literal.text if literal is not None else''),'categories':sparse(cache('c:cat')),'values':sparse(cache('c:val'))})
                result['charts'].append({'part':name,'series':chart})
    return result
def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('docx');p.add_argument('--charts',action='store_true');p.add_argument('--out',required=True);a=p.parse_args();Path(a.out).write_text(json.dumps(extract(a.docx,a.charts),ensure_ascii=False,indent=2),encoding='utf8')
if __name__=='__main__':main()
