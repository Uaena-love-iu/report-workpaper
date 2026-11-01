"""Native editable report cells, evidence images, hyperlinks and typed tables.

Standalone Python backend. Host-specific spreadsheet tool constraints take
precedence; the same profile and manifest can drive a host-native implementation.
"""
from pathlib import Path
from datetime import date
import argparse,json,math,os,hashlib
from PIL import Image,ImageDraw,ImageFont
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment,Border,Side,PatternFill
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor,AnchorMarker
from openpyxl.utils import get_column_letter as letter
from openpyxl.utils.units import pixels_to_EMU
from workpaper_model import load_manifest,load_profile,layout_sheet,source_metadata,STATUSES,local_path,wrapped_lines

def literal(cell,value):
    cell.value=value
    if isinstance(value,str):cell.data_type='s'

def merged(ws,r,c,nr,nc,value,font,align='left'):
    ws.merge_cells(start_row=r,start_column=c,end_row=r+nr-1,end_column=c+nc-1)
    cell=ws.cell(r,c);literal(cell,value);cell.font=font;cell.alignment=Alignment(horizontal=align,vertical='center',wrap_text=True)
    return cell

def marker(x,y,cp,rp):
    col=int(x//cp);row=int(y//rp)
    return AnchorMarker(col=col,row=row,colOff=pixels_to_EMU(x-col*cp),rowOff=pixels_to_EMU(y-row*rp))

def add_image(ws,file,row,col,width,height,cp,rp):
    im=XLImage(file);x=(col-1)*cp;y=(row-1)*rp
    im.anchor=TwoCellAnchor(editAs='oneCell',_from=marker(x,y,cp,rp),to=marker(x+width,y+height,cp,rp));ws.add_image(im)

def build(manifest,out,profile=None,preview_dir=None):
    m=load_manifest(manifest);p=load_profile(profile);wb=Workbook();wb.remove(wb.active);layouts=[];assets={}
    cp=p['grid']['column_px'];rh=p['grid']['row_px'];rp=p['report'];ep=p['evidence'];colors=p['colors']
    text_font=Font(name=rp['font'],size=rp['font_pt'],color=colors['text']);thin=Side(style='thin',color=colors['separator'])
    for definition in m['sheets']:
        ws=wb.create_sheet(definition['label']);ws.sheet_view.showGridLines=p['grid']['show_grid_lines'];ws.sheet_view.zoomScale=85
        layout=layout_sheet(definition,m,p);layouts.append((definition,layout));maxcol=layout['columns']
        if m.get('demo'):merged(ws,1,1,2,maxcol,'格式演示：数据与截图均为虚构，不可用于研究结论',Font(name=rp['font'],size=11,bold=True,color='9C0006'))
        for b in layout['blocks']:
            row=b['row'];n=b['end']-row+1;item=b['item'];merged(ws,row,1,n,rp['columns'],item['text'],text_font,'center')
            for col in range(1,rp['columns']+1):ws.cell(row,col).border=Border(top=thin)
            for rr in range(row,b['end']+1):ws.cell(rr,rp['columns']).border=Border(right=thin,top=thin if rr==row else Side())
            if b['note_rows']:
                color=colors['correction']if item['status']=='correction'else colors['note']
                merged(ws,row+b['title_rows']+b['metadata_rows'],rp['columns']+1,b['note_rows'],ep['columns']*b['lanes'],b['note'],Font(name=rp['font'],size=ep['note_font_pt'],color=color))
            if not item.get('sources'):
                label='iFinD（按约定待补截图）'if item['status']=='deferred_ifind'else'作者分析（见核查说明）'if item['status']=='analysis'else'来源待补'
                merged(ws,row,rp['columns']+1,b['title_rows'],ep['columns'],label,Font(name=rp['font'],size=11,color=colors['pending']))
            for placement in b['images']:
                src=placement['source'];col=placement['col'];cell=merged(ws,row,col,b['title_rows'],ep['columns'],src['title'],Font(name=rp['font'],size=ep['title_font_pt'],color=colors['link'],underline='single'))
                if src.get('url'):cell.hyperlink=src['url']+(f"#page={src['pdf_page']}"if src['kind']=='pdf_crop'else'')
                merged(ws,row+b['title_rows'],col,b['metadata_rows'],ep['columns'],source_metadata(src),Font(name=rp['font'],size=ep['metadata_font_pt'],color=colors['note']))
                add_image(ws,placement['file'],placement['row'],col,placement['width'],placement['height'],cp,rh)
                assets[placement['file']]=hashlib.sha256(Path(placement['file']).read_bytes()).hexdigest()
        cursor=layout['end_row']+1;layout['table_cells']=[];layout['table_images']=[]
        for table in definition.get('tables',[]):
            columns=table['columns'];fontname=p['table']['product_font']if table.get('kind')=='product'else p['table']['font'];spans=[int(c.get('span',3))for c in columns];maxcol=max(maxcol,sum(spans))
            if table.get('title'):
                merged(ws,cursor,1,2,sum(spans),table['title'],Font(name=fontname,size=11,bold=True));layout['table_cells'].append({'row':cursor,'col':1,'rows':2,'cols':sum(spans),'text':table['title']});cursor+=2
            col=1
            for spec,span in zip(columns,spans):
                merged(ws,cursor,col,2,span,spec['title'],Font(name=fontname,size=11,bold=True,color=colors['table_header_text']),'center')
                layout['table_cells'].append({'row':cursor,'col':col,'rows':2,'cols':span,'text':spec['title'],'fill':colors['table_header'],'color':colors['table_header_text']})
                for rr in range(cursor,cursor+2):
                    for cc in range(col,col+span):ws.cell(rr,cc).fill=PatternFill('solid',fgColor=colors['table_header'])
                col+=span
            cursor+=2
            for record in table['rows']:
                height=max([4]+[wrapped_lines(v,span*cp,11)+1 for v,span in zip(record,spans)if not isinstance(v,dict)])
                col=1
                for spec,span,value in zip(columns,spans,record):
                    cell=merged(ws,cursor,col,height,span,None,Font(name=fontname,size=11),'center')
                    if isinstance(value,dict)and'image'in value:
                        path=local_path(m,value['image'])
                        with Image.open(path)as im:iw,ih=im.size
                        width=min(span*cp-10,(height*rh-10)*iw/ih);add_image(ws,str(path),cursor,col,width,width*ih/iw,cp,rh);layout['table_images'].append({'file':str(path),'row':cursor,'col':col,'width':width,'height':width*ih/iw})
                    elif isinstance(value,dict)and'formula'in value:cell.value=value['formula']
                    elif isinstance(value,dict)and'date'in value:cell.value=date.fromisoformat(value['date']);cell.number_format='yyyy-mm-dd'
                    else:literal(cell,value)
                    if spec.get('format'):cell.number_format=spec['format']
                    if spec.get('role')=='series':
                        for rr in range(cursor,cursor+height):
                            for cc in range(col,col+span):ws.cell(rr,cc).fill=PatternFill('solid',fgColor=colors['product_series'])
                    display=''if isinstance(value,dict)and'image'in value else value.get('date',value.get('formula',''))if isinstance(value,dict)else''if value is None else str(value)
                    layout['table_cells'].append({'row':cursor,'col':col,'rows':height,'cols':span,'text':display,'fill':colors['product_series']if spec.get('role')=='series'else None})
                    col+=span
                cursor+=height
            cursor+=2
        layout['end_row']=max(cursor,layout['end_row']);layout['columns']=maxcol
        for col in range(1,maxcol+1):ws.column_dimensions[letter(col)].width=(cp-5)/7
        for rr in range(1,max(cursor,layout['end_row'])+1):ws.row_dimensions[rr].height=rh*.75
    out=Path(out);out.parent.mkdir(parents=True,exist_ok=True);wb.save(out)
    # OPC-only image deduplication; all placements and bytes are retained.
    from verify_workbook import deduplicate_media,verify
    deduplicate_media(out);report=verify(out);report.update({'coverage':{s:sum(r['status']==s for sh in m['sheets']for r in sh.get('rows',[]))for s in STATUSES},'source_image_hashes':assets,'evidence_truth_review':'not_automatically_validated','native_excel_open':False})
    out.with_suffix('.validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    if preview_dir:
        for definition,layout in layouts:render_preview(definition,layout,m,p,Path(preview_dir)/(definition['label']+'.png'))
    return report

def font(size):
    candidates=[os.environ.get('REPORT_WORKPAPER_FONT',''),'C:/Windows/Fonts/msyh.ttc','C:/Windows/Fonts/Deng.ttf','/System/Library/Fonts/PingFang.ttc','/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc','/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']
    for candidate in candidates:
        if candidate and Path(candidate).is_file():return ImageFont.truetype(candidate,size)
    return ImageFont.load_default()

def lines(text,draw,f,width):
    result=[]
    for raw in str(text).split('\n'):
        current=''
        for char in raw:
            if current and draw.textlength(current+char,font=f)>width:result.append(current);current=''
            current+=char
        result.append(current)
    return result

def render_preview(definition,layout,m,p,out):
    cp=p['grid']['column_px'];rh=p['grid']['row_px'];width=layout['columns']*cp;height=layout['end_row']*rh
    # Bound preview memory; very long sheets are shown as one PNG per strip.
    if width*height>35_000_000:
        if len(layout['blocks'])<=1:raise ValueError('Preview strip is too large; crop source images or preview a bounded range')
        for index,block in enumerate(layout['blocks']):
            small={**layout,'blocks':[{**block,'row':1,'end':block['end']-block['row']+1,'image_row':block['image_row']-block['row']+1,'images':[{**im,'row':im['row']-block['row']+1}for im in block['images']]}],'end_row':block['end']-block['row']+2,'table_cells':[],'table_images':[]}
            render_preview(definition,small,{**m,'demo':False},p,out.with_name(out.stem+f'_{index+1}'+out.suffix))
        return
    canvas=Image.new('RGB',(width,height),'white');draw=ImageDraw.Draw(canvas);f=font(15);sm=font(13);left=p['report']['columns']*cp;lane=p['evidence']['columns']*cp
    if m.get('demo'):draw.text((8,4),'DEMO ONLY - synthetic data and images',fill='#9C0006',font=f)
    for b in layout['blocks']:
        y=(b['row']-1)*rh;bottom=b['end']*rh;draw.line((0,y,left,y),fill='black');draw.line((left,y,left,bottom),fill='black')
        ll=lines(b['item']['text'],draw,f,left-24);ty=y+(bottom-y-len(ll)*20)/2
        for line in ll:draw.text(((left-draw.textlength(line,font=f))/2,ty),line,font=f,fill='black');ty+=20
        if b['note']:
            ny=y+(b['title_rows']+b['metadata_rows'])*rh
            for line in lines(b['note'],draw,sm,lane*b['lanes']-12):draw.text((left+5,ny),line,font=sm,fill='#9C0006'if b['item']['status']=='correction'else'#555555');ny+=17
        if not b['images']:
            label='iFinD（按约定待补截图）'if b['item']['status']=='deferred_ifind'else'作者分析（见核查说明）'if b['item']['status']=='analysis'else'来源待补'
            draw.text((left+4,y),label,font=f,fill='#777777')
        for placement in b['images']:
            x=(placement['col']-1)*cp;src=placement['source'];yy=y
            for line in lines(src['title'],draw,f,lane-12):draw.text((x+4,yy),line,font=f,fill='#0563C1');yy+=20
            yy=y+b['title_rows']*rh
            for line in lines(source_metadata(src),draw,sm,lane-12):draw.text((x+4,yy),line,font=sm,fill='#666666');yy+=17
            im=Image.open(placement['file']).convert('RGB');im=im.resize((round(placement['width']),round(placement['height'])));canvas.paste(im,(x,(placement['row']-1)*rh))
    for cell in layout.get('table_cells',[]):
        x=(cell['col']-1)*cp;y=(cell['row']-1)*rh;w=cell['cols']*cp;h=cell['rows']*rh
        if cell.get('fill'):draw.rectangle((x,y,x+w,y+h),fill='#'+cell['fill'])
        content=lines(cell['text'],draw,sm,w-10);ty=y+(h-len(content)*17)/2
        for line in content:draw.text((x+(w-draw.textlength(line,font=sm))/2,ty),line,font=sm,fill='#'+cell.get('color','111111'));ty+=17
        draw.line((x,y+h,x+w,y+h),fill='#BBBBBB')
    for placement in layout.get('table_images',[]):
        im=Image.open(placement['file']).convert('RGB');im=im.resize((round(placement['width']),round(placement['height'])));canvas.paste(im,((placement['col']-1)*cp,(placement['row']-1)*rh))
    out.parent.mkdir(parents=True,exist_ok=True);canvas.save(out)

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('manifest');parser.add_argument('--out',required=True);parser.add_argument('--profile');parser.add_argument('--preview-dir');args=parser.parse_args()
    try:report=build(args.manifest,args.out,args.profile,args.preview_dir);print(json.dumps({k:v for k,v in report.items()if k!='source_image_hashes'},ensure_ascii=False))
    except (ValueError,OSError)as exc:parser.exit(1,str(exc)+'\n')
if __name__=='__main__':main()
