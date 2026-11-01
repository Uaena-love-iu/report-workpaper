"""Manifest validation and shared native-cell layout (no network access)."""
from pathlib import Path
from datetime import date
from urllib.parse import urlsplit
import json,math,re,hashlib
from PIL import Image

ROOT=Path(__file__).resolve().parent.parent
STATUSES={'support':'已支撑','partial':'部分支撑','correction':'需更正','analysis':'分析判断','unsupported':'待补证','deferred_ifind':'iFinD待补截图'}
CLASSES={'company_filing','official','government','industry_body','institutional_media','user_database','archive','synthetic_demo'}
KINDS={'pdf_crop','web_screenshot','database_screenshot','archive_excerpt','synthetic_demo'}

def load_manifest(path):
    p=Path(path).resolve();m=json.loads(p.read_text('utf-8-sig'));m['_base']=str(p.parent)
    validate_manifest(m);return m

def local_path(m,value):
    p=Path(value);return p if p.is_absolute()else Path(m['_base'])/p

def valid_date(value):
    try:date.fromisoformat(value);return True
    except (ValueError,TypeError):return False

def validate_manifest(m):
    errors=[];names=set();ids=set();demo=m.get('demo',False)
    if not m.get('sheets'):errors.append('sheets must not be empty')
    if m.get('as_of')and not valid_date(m['as_of']):errors.append('as_of must be YYYY-MM-DD')
    for sheet in m.get('sheets',[]):
        label=sheet.get('label','')
        if not label or len(label)>31 or re.search(r'[\\/*?:\[\]]',label):errors.append(f'invalid sheet label: {label}')
        if label.casefold()in names:errors.append(f'duplicate sheet label: {label}')
        names.add(label.casefold())
        for row in sheet.get('rows',[]):
            rid=row.get('id');status=row.get('status')
            if not rid or rid in ids:errors.append(f'missing/duplicate row id: {rid}')
            ids.add(rid)
            if not isinstance(row.get('text'),str)or not row['text'].strip():errors.append(f'{rid}: editable text required; use legacy builder for screenshot-left mode')
            if status not in STATUSES:errors.append(f'{rid}: invalid status')
            if status in {'partial','analysis','unsupported','deferred_ifind'}and not row.get('note'):errors.append(f'{rid}: status requires a specific note')
            if status=='correction'and(not row.get('note')or not row.get('suggestion')):errors.append(f'{rid}: correction requires note and replacement suggestion')
            if status=='deferred_ifind'and not row.get('user_deferred'):errors.append(f'{rid}: database deferral needs explicit user_deferred=true')
            if status=='unsupported'and not row.get('search_log'):errors.append(f'{rid}: unsupported needs search_log')
            claims=row.get('sentence_ids')or[rid];covered=set();strong=0
            for src in row.get('sources',[]):
                cls=src.get('source_class');kind=src.get('kind');url=src.get('url','')
                if cls not in CLASSES:errors.append(f'{rid}: source class not permitted: {cls}')
                if kind not in KINDS:errors.append(f'{rid}: source kind not permitted: {kind}')
                if (kind=='synthetic_demo'or cls=='synthetic_demo')and not demo:errors.append(f'{rid}: synthetic evidence forbidden outside demo')
                if not src.get('title'):errors.append(f'{rid}: source title required')
                parsed=urlsplit(url)
                if not url and cls!='user_database':errors.append(f'{rid}: source URL required')
                if url and(parsed.scheme not in {'http','https'}or not parsed.hostname or parsed.username or parsed.password):errors.append(f'{rid}: unsafe or invalid source URL')
                if not valid_date(src.get('collected_at')):errors.append(f'{rid}: collected_at must be YYYY-MM-DD')
                if src.get('published_at')and not valid_date(src['published_at']):errors.append(f'{rid}: invalid publication date')
                if m.get('as_of')and src.get('published_at')and src['published_at']>m['as_of']:errors.append(f'{rid}: source publication is after research cutoff')
                if kind=='pdf_crop':
                    if not isinstance(src.get('pdf_page'),int)or src['pdf_page']<1:errors.append(f'{rid}: PDF physical page required')
                    if 'printed_page'not in src:errors.append(f'{rid}: printed_page required; null means absent')
                elif not src.get('location'):errors.append(f'{rid}: webpage/database location required')
                image=src.get('image')
                if not image:errors.append(f'{rid}: evidence screenshot required')
                else:
                    try:
                        path=local_path(m,image)
                        with Image.open(path)as im:im.verify()
                        if src.get('screenshot_sha256')and hashlib.sha256(path.read_bytes()).hexdigest()!=src['screenshot_sha256']:errors.append(f'{rid}: screenshot hash mismatch')
                    except (OSError,ValueError):errors.append(f'{rid}: missing or invalid screenshot: {image}')
                mapped=src.get('claim_ids')or(claims if len(claims)==1 else[])
                if len(claims)>1 and not src.get('claim_ids'):errors.append(f'{rid}: multi-sentence block requires source claim_ids')
                if set(mapped)-set(claims):errors.append(f'{rid}: unknown claim_ids')
                if src.get('verified')is True and (kind not in {'archive_excerpt','synthetic_demo'}or(demo and kind=='synthetic_demo')):
                    strong+=1;covered.update(mapped)
            if status=='support'and(not strong or set(claims)-covered):errors.append(f'{rid}: support needs verified screenshots covering each sentence')
        for table in sheet.get('tables',[]):
            columns=table.get('columns',[])
            if not columns:errors.append(f'{label}: native table needs columns')
            for record in table.get('rows',[]):
                if len(record)!=len(columns):errors.append(f'{label}: native table width mismatch')
                for value in record:
                    if isinstance(value,float)and not math.isfinite(value):errors.append(f'{label}: nonfinite table value')
                    if isinstance(value,dict)and 'date'in value and not valid_date(value['date']):errors.append(f'{label}: invalid typed table date')
                    if isinstance(value,dict)and 'formula'in value and not str(value['formula']).startswith('='):errors.append(f'{label}: explicit formula must start with =')
    if errors:raise ValueError('\n'.join(errors))
    return {'sheets':len(names),'rows':len(ids),'demo':demo}

def text_units(text):return sum(1 if ord(c)>255 else .55 for c in str(text))

def wrapped_lines(text,width_px,font_pt=11):
    capacity=max(1,width_px/(font_pt*96/72))
    return sum(max(1,math.ceil(text_units(line)/capacity))for line in str(text).split('\n'))

def load_profile(path=None):
    return json.loads(Path(path or ROOT/'assets/analyst-format.json').read_text('utf8'))

def source_metadata(src):
    if src['kind']=='pdf_crop':
        physical=f"PDF第{src['pdf_page']}页";printed=f"印刷第{src['printed_page']}页"if src['printed_page']is not None else'无印刷页码'
        location=physical+'；'+printed
    else:location=src.get('location','网页正文')
    if src['kind']=='archive_excerpt':location='存档摘录，非当日网页截图；'+location
    return location+'；采集 '+src['collected_at']

def layout_sheet(sheet,m,profile):
    rp=profile['report'];ep=profile['evidence'];grid=profile['grid'];cp=grid['column_px'];rh=grid['row_px'];left=rp['columns'];lane=ep['columns'];row=3 if m.get('demo')else 1;blocks=[]
    for item in sheet.get('rows',[]):
        sources=item.get('sources',[]);lanes=max(1,len(sources));note=item.get('note','')
        if item.get('suggestion'):note+='\n替换建议：'+item['suggestion']
        note=(STATUSES[item['status']]+'：'+note)if note else''
        title_rows=max([1]+[wrapped_lines(s['title'],lane*cp,ep['title_font_pt'])for s in sources])
        meta_rows=max([1]+[wrapped_lines(source_metadata(s),lane*cp,ep['metadata_font_pt'])for s in sources])
        note_rows=math.ceil(wrapped_lines(note,lanes*lane*cp,ep['note_font_pt'])*17/rh)if note else 0
        image_row=row+title_rows+meta_rows+note_rows
        image_height=0;images=[]
        for i,src in enumerate(sources):
            with Image.open(local_path(m,src['image']))as im:w,h=im.size
            height=h*ep['image_width_px']/w;image_height=max(image_height,height)
            images.append({'source':src,'file':str(local_path(m,src['image'])),'row':image_row,'col':left+i*lane+1,'width':ep['image_width_px'],'height':height})
        text_rows=math.ceil((wrapped_lines(item['text'],left*cp-20,rp['font_pt'])*20+24)/rh)
        height=max(rp['minimum_rows'],text_rows,title_rows+meta_rows+note_rows+math.ceil((image_height+ep['image_gap_px'])/rh))
        blocks.append({'item':item,'row':row,'end':row+height-1,'title_rows':title_rows,'metadata_rows':meta_rows,'note_rows':note_rows,'note':note,'image_row':image_row,'images':images,'lanes':lanes})
        row+=height
    return {'blocks':blocks,'end_row':row,'columns':left+lane*max([1]+[b['lanes']for b in blocks])}
