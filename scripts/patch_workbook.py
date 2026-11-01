"""Transplant one edited worksheet while preserving all unrelated package parts.

Fails closed for unsupported target-sheet features. Does not resave the original
with a high-level workbook library and does not modify its global string table.
"""
from pathlib import Path
from lxml import etree as E
import argparse,json,copy,hashlib,posixpath
from verify_workbook import S,R,P,N,parse,xml,resolve,relpath,sha,read_parts,write_parts,sheet_parts,verify,relationships_xml,style_signature

def guard_sheet(parts,part):
    root=parse(parts[part])
    bad={'tableParts','legacyDrawing','legacyDrawingHF','oleObjects','controls','conditionalFormatting','extLst'}
    if any(E.QName(child).localname in bad for child in root):raise ValueError('Target has unsupported native objects; use host-native editing')
    refs=list(parse(parts[relpath(part)]))if relpath(part)in parts else[]
    for rel in refs:
        if not any(rel.get('Type').endswith('/'+t)for t in ['drawing','hyperlink']):raise ValueError('Unsupported target relationship: '+rel.get('Type'))
    drawings=[resolve(part,r.get('Target'))for r in refs if r.get('Type').endswith('/drawing')]
    if len(drawings)>1:raise ValueError('Multiple target drawings not supported')
    for drawing in drawings:
        rp=relpath(drawing)
        for rel in parse(parts[rp])if rp in parts else[]:
            if not rel.get('Type').endswith('/image')or rel.get('TargetMode')=='External':raise ValueError('Target drawing contains charts/objects or external images')
        for name,data in parts.items():
            if name.startswith('xl/worksheets/_rels/')and name!=relpath(part):
                origin=name.replace('/_rels/','/').removesuffix('.rels')
                if any(r.get('TargetMode')!='External'and resolve(origin,r.get('Target'))==drawing for r in parse(data)):raise ValueError('Drawing shared with another sheet')
    return root,refs,drawings

def append_styles(parts,replacement,root):
    old=parse(parts['xl/styles.xml']);new=parse(replacement['xl/styles.xml']);maps={};fmt=old.find('s:numFmts',N)
    if fmt is None:fmt=E.Element('{'+S+'}numFmts');old.insert(0,fmt)
    codes={e.get('formatCode'):int(e.get('numFmtId'))for e in fmt};nextid=max([163]+list(codes.values()))+1;fm={}
    for e in new.findall('s:numFmts/s:numFmt',N):
        code=e.get('formatCode')
        if code not in codes:c=copy.deepcopy(e);c.set('numFmtId',str(nextid));fmt.append(c);codes[code]=nextid;nextid+=1
        fm[int(e.get('numFmtId'))]=codes[code]
    fmt.set('count',str(len(fmt)))
    order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
    for tag in ['fonts','fills','borders','cellStyleXfs','cellXfs']:
        src=new.find('s:'+tag,N);dst=old.find('s:'+tag,N)
        if src is None or dst is None:raise ValueError('Missing required style group '+tag)
        mapping={}
        for i,item in enumerate(src):
            c=copy.deepcopy(item)
            if tag in ['cellStyleXfs','cellXfs']:
                for attr,group in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
                    if c.get(attr)is not None:c.set(attr,str(maps[group][int(c.get(attr))]))
                if c.get('numFmtId')is not None:c.set('numFmtId',str(fm.get(int(c.get('numFmtId')),int(c.get('numFmtId')))))
                if tag=='cellXfs'and c.get('xfId')is not None:c.set('xfId',str(maps['cellStyleXfs'][int(c.get('xfId'))]))
            mapping[i]=len(dst);dst.append(c)
        dst.set('count',str(len(dst)));maps[tag]=mapping
    for c in root.findall('s:sheetData/s:row/s:c',N):
        c.set('s',str(maps['cellXfs'][int(c.get('s','0'))]))
    parts['xl/styles.xml']=xml(old)

def patch(original,edited,sheet,out,expected_sha256=None):
    original=Path(original);out=Path(out)
    if original.resolve()==out.resolve():raise ValueError('Write a candidate first; in-place replacement is a separate authorized step')
    original_sha=sha(original.read_bytes())
    if expected_sha256 and original_sha!=expected_sha256:raise ValueError('Original workbook changed; refusing to overwrite newer edits')
    parts=read_parts(original);replacement=read_parts(edited);before=dict(parts);oldmap=sheet_parts(parts);newmap=sheet_parts(replacement)
    if sheet not in oldmap or sheet not in newmap:raise ValueError('Requested sheet not present in both workbooks')
    part=oldmap[sheet];np=newmap[sheet];oldroot,oldrefs,olddrawings=guard_sheet(parts,part);root,refs,drawings=guard_sheet(replacement,np)
    append_styles(parts,replacement,root)
    shared=list(parse(replacement['xl/sharedStrings.xml']))if'xl/sharedStrings.xml'in replacement else[]
    for c in root.findall('s:sheetData/s:row/s:c',N):
        if c.get('t')=='s':
            value=c.find('s:v',N);item=shared[int(value.text)];c.remove(value);c.set('t','inlineStr');inline=E.SubElement(c,'{'+S+'}is')
            for child in item:inline.append(copy.deepcopy(child))
        if c.get('t')=='d':raise ValueError('Normalize date cells to Excel serials in the editing backend before patching')
    newrefs=E.Element('{'+P+'}Relationships',nsmap={None:P});media={sha(v):p for p,v in parts.items()if p.startswith('xl/media/')};additions=[];content_type=parse(parts['[Content_Types].xml']);ctns=E.QName(content_type).namespace
    defaults={e.get('Extension'):e.get('ContentType')for e in content_type if e.get('Extension')}
    replacement_defaults={e.get('Extension'):e.get('ContentType')for e in parse(replacement['[Content_Types].xml'])if e.get('Extension')}
    for ref in refs:
        c=copy.deepcopy(ref)
        if c.get('Type').endswith('/drawing'):
            source_drawing=resolve(np,c.get('Target'));destination=olddrawings[0]if olddrawings else'xl/drawings/workpaper_'+original_sha[:12]+'.xml'
            if not olddrawings and destination in parts:raise ValueError('New drawing name collision')
            dr=parse(replacement[relpath(source_drawing)])
            for image in dr:
                source_image=resolve(source_drawing,image.get('Target'));raw=replacement[source_image];digest=sha(raw)
                if digest not in media:
                    target='xl/media/workpaper_'+digest+Path(source_image).suffix
                    if target in parts and parts[target]!=raw:raise ValueError('Image collision')
                    parts[target]=raw;media[digest]=target;additions.append(target)
                    extension=Path(target).suffix[1:]
                    if extension not in defaults:
                        if extension not in replacement_defaults:raise ValueError('Missing image MIME type')
                        E.SubElement(content_type,'{'+ctns+'}Default',Extension=extension,ContentType=replacement_defaults[extension]);defaults[extension]=replacement_defaults[extension]
                image.set('Target',posixpath.relpath(media[digest],posixpath.dirname(destination)))
            parts[destination]=replacement[source_drawing];parts[relpath(destination)]=relationships_xml(dr)
            if not any(e.get('PartName')=='/'+destination for e in content_type):E.SubElement(content_type,'{'+ctns+'}Override',PartName='/'+destination,ContentType='application/vnd.openxmlformats-officedocument.drawing+xml')
            c.set('Target',posixpath.relpath(destination,posixpath.dirname(part)))
        elif c.get('TargetMode')!='External':raise ValueError('Unexpected worksheet relationship')
        newrefs.append(c)
    # Keep original view, protection and print settings, even if absent in the edit.
    # ECMA-376 worksheet sequence: inserting at the old numeric index is unsafe
    # when the independently generated sheet contains a different set of nodes.
    order='sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPrs conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems tableParts extLst'.split()
    tags=['sheetPr','sheetViews','pageMargins','pageSetup','printOptions','headerFooter','sheetProtection']
    for tag in tags:
        previous=oldroot.find('s:'+tag,N);current=root.find('s:'+tag,N)
        if previous is not None:
            if current is not None:root.remove(current)
            rank=order.index(tag)
            index=next((i for i,e in enumerate(root)if E.QName(e).localname in order and order.index(E.QName(e).localname)>rank),len(root))
            root.insert(index,copy.deepcopy(previous))
    parts[part]=xml(root)
    if len(newrefs):parts[relpath(part)]=relationships_xml(newrefs)
    elif relpath(part)in parts:del parts[relpath(part)]
    if len(content_type)!=len(parse(before['[Content_Types].xml'])):parts['[Content_Types].xml']=xml(content_type)
    if sha(original.read_bytes())!=original_sha:raise ValueError('Original changed during patch preparation')
    write_parts(out,parts)
    try:report=verify(out,original,sheet)
    except Exception:
        # Candidate is retained for diagnosis. Original is never written here.
        raise
    report.update(original_sha256=original_sha,added_media=len(additions),original_unmodified=sha(original.read_bytes())==original_sha)
    return report

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('original');p.add_argument('edited');p.add_argument('--sheet',required=True);p.add_argument('--out',required=True);p.add_argument('--expected-sha256');p.add_argument('--report');a=p.parse_args()
    try:
        result=patch(a.original,a.edited,a.sheet,a.out,a.expected_sha256)
        if a.report:Path(a.report).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
        print(json.dumps(result,ensure_ascii=False))
    except (ValueError,OSError)as e:p.exit(1,str(e)+'\n')
if __name__=='__main__':main()
