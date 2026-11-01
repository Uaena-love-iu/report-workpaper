"""Inspect OPC/XLSX integrity without treating it as evidence-truth validation."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from PIL import Image
from io import BytesIO
import argparse,hashlib,json,posixpath,copy,os,time

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';P='http://schemas.openxmlformats.org/package/2006/relationships';N={'s':S}
def parse(raw):return E.fromstring(raw,E.XMLParser(resolve_entities=False,no_network=True))
def xml(root):return E.tostring(root,encoding='utf-8',xml_declaration=True,standalone=True)
def relationships_xml(root):
    clean=E.Element('{'+P+'}Relationships',nsmap={None:P})
    for item in root:E.SubElement(clean,item.tag,attrib=dict(item.attrib))
    return xml(clean)
def sha(raw):return hashlib.sha256(raw).hexdigest()
def resolve(part,target):return target.lstrip('/')if target.startswith('/')else posixpath.normpath(posixpath.join(posixpath.dirname(part),target))
def relpath(part):return posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
def origin_of(rel):return rel.replace('/_rels/','/').removesuffix('.rels')if rel!='_rels/.rels'else''
def sheet_parts(parts):
    refs={r.get('Id'):r for r in parse(parts['xl/_rels/workbook.xml.rels'])}
    return {s.get('name'):resolve('xl/workbook.xml',refs[s.get('{'+R+'}id')].get('Target'))for s in parse(parts['xl/workbook.xml']).findall('s:sheets/s:sheet',N)}
def read_parts(path):
    with ZipFile(path)as z:
        if z.testzip():raise ValueError('ZIP CRC failure')
        return {name:z.read(name)for name in z.namelist()}
def write_parts(path,parts):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True);temporary=path.with_suffix(path.suffix+'.tmp')
    try:
        with ZipFile(temporary,'w',ZIP_DEFLATED)as z:
            for name,raw in parts.items():z.writestr(name,raw)
        for attempt,delay in enumerate([.1,.25,.5,0]):
            try:os.replace(temporary,path);break
            except PermissionError:
                if attempt==3:raise
                # Windows can briefly scan a newly written XLSX. Never close
                # apps or change permissions; a persistent lock is an error.
                time.sleep(delay)
    finally:
        if temporary.exists():temporary.unlink()

def deduplicate_media(path):
    parts=read_parts(path);canonical={};remap={}
    for name,raw in parts.items():
        if name.startswith('xl/media/'):
            digest=sha(raw)
            if digest in canonical:remap[name]=canonical[digest]
            else:canonical[digest]=name
    if not remap:return 0
    for name in list(parts):
        if name.endswith('.rels'):
            root=parse(parts[name]);origin=origin_of(name);changed=False
            for item in root:
                if item.get('TargetMode')=='External':continue
                target=resolve(origin,item.get('Target'))
                if target in remap:item.set('Target',posixpath.relpath(remap[target],posixpath.dirname(origin)));changed=True
            if changed:parts[name]=relationships_xml(root)
    ct=parse(parts['[Content_Types].xml'])
    for item in list(ct):
        if item.get('PartName','').lstrip('/')in remap:ct.remove(item)
    parts['[Content_Types].xml']=xml(ct)
    for name in remap:del parts[name]
    write_parts(path,parts);return len(remap)

def style_signature(node):
    return (node.tag,tuple(sorted(node.attrib.items())),node.text or'',tuple(style_signature(c)for c in node))

def verify(path,baseline=None,allowed_sheet=None):
    parts=read_parts(path);sheets=sheet_parts(parts);styles=parse(parts['xl/styles.xml']);style_count=len(styles.find('s:cellXfs',N));links=0;media=0
    for name,raw in parts.items():
        if name.endswith('.rels'):
            root=parse(raw);ids=[r.get('Id')for r in root]
            if len(ids)!=len(set(ids)):raise ValueError('Duplicate relationship IDs: '+name)
            for rel in root:
                if rel.get('TargetMode')=='External':continue
                target=resolve(origin_of(name),rel.get('Target'))
                if target not in parts:raise ValueError(f'Missing relationship target: {name} -> {target}')
        elif name.startswith('xl/media/'):
            try:
                with Image.open(BytesIO(raw))as im:im.verify()
            except OSError:
                # Preserve Windows EMF/WMF, which Pillow may not decode.
                if Path(name).suffix.lower()not in {'.emf','.wmf'}:raise
            media+=1
    for name,part in sheets.items():
        root=parse(parts[part]);rp=relpath(part);refs={r.get('Id')for r in parse(parts[rp])}if rp in parts else set()
        for c in root.findall('s:sheetData/s:row/s:c',N):
            if int(c.get('s','0'))>=style_count:raise ValueError('Invalid style index')
        hs=root.find('s:hyperlinks',N)
        if hs is not None and not len(hs):raise ValueError('Empty hyperlinks element')
        for h in root.findall('s:hyperlinks/s:hyperlink',N):
            rid=h.get('{'+R+'}id')
            if rid and rid not in refs:raise ValueError('Broken hyperlink relationship')
            links+=1
    result={'valid_package':True,'sheet_count':len(sheets),'hyperlinks':links,'media_files':media,'sha256':sha(Path(path).read_bytes()),'native_excel_open':False}
    if baseline:
        if not allowed_sheet:raise ValueError('--allowed-sheet required with --baseline')
        before=read_parts(baseline);old_sheets=sheet_parts(before)
        if list(old_sheets)!=list(sheets):raise ValueError('Sheet names/order changed')
        for name,part in old_sheets.items():
            if name!=allowed_sheet and before[part]!=parts[part]:raise ValueError('Unrelated worksheet changed: '+name)
        oldstyles=parse(before['xl/styles.xml'])
        for oldgroup in oldstyles:
            group=styles.find('s:'+E.QName(oldgroup).localname,N)
            if group is None or [style_signature(c)for c in group][:len(oldgroup)]!=[style_signature(c)for c in oldgroup]:raise ValueError('Existing styles changed')
        allowed={old_sheets[allowed_sheet],relpath(old_sheets[allowed_sheet]),'xl/styles.xml','[Content_Types].xml'}
        sr=relpath(old_sheets[allowed_sheet])
        if sr in before:
            for r in parse(before[sr]):
                if r.get('Type').endswith('/drawing'):
                    d=resolve(old_sheets[allowed_sheet],r.get('Target'));allowed|={d,relpath(d)}
        changed=[p for p,raw in before.items()if p not in parts or parts[p]!=raw]
        if set(changed)-allowed:raise ValueError('Unexpected modified parts: '+str(sorted(set(changed)-allowed)))
        result.update(other_worksheets_unchanged=True,changed_existing_parts=changed)
    return result

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('workbook');p.add_argument('--baseline');p.add_argument('--allowed-sheet');a=p.parse_args()
    try:print(json.dumps(verify(a.workbook,a.baseline,a.allowed_sheet),ensure_ascii=False))
    except (ValueError,OSError)as e:p.exit(1,str(e)+'\n')
if __name__=='__main__':main()
