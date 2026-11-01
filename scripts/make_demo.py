"""Generate only synthetic, clearly labeled format examples. No research files."""
from pathlib import Path
from PIL import Image,ImageDraw
import argparse,json
from build_workpaper import build,font
ROOT=Path(__file__).resolve().parent.parent
def make_demo(outdir):
    outdir=Path(outdir);outdir.mkdir(parents=True,exist_ok=True)
    source=Image.new('RGB',(1100,240),'white');draw=ImageDraw.Draw(source);f=font(28)
    draw.text((30,20),'SYNTHETIC DEMO - NOT A REAL RESEARCH SOURCE',fill='#9C0006',font=f)
    draw.text((30,83),'项目甲：5万元      项目乙：3万元      合计：8万元',fill='black',font=f)
    draw.rectangle((25,78,850,128),outline='#D00000',width=3)
    draw.text((30,151),'示例公司持续投入品牌建设。',fill='black',font=f);draw.rectangle((25,143,650,198),outline='#D00000',width=3);source.save(outdir/'source.png')
    product=Image.new('RGB',(180,180),'#F0F4FA');draw=ImageDraw.Draw(product);draw.rectangle((35,30,145,130),outline='#2F5597',width=4);draw.text((35,145),'DEMO ONLY',font=font(17),fill='#9C0006');product.save(outdir/'product.png')
    data=json.loads((ROOT/'examples/minimal.json').read_text('utf8'));(outdir/'manifest.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
    report=build(outdir/'manifest.json',outdir/'demo.xlsx',preview_dir=outdir/'previews');return report
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--outdir',default='outputs/demo');a=p.parse_args();print(json.dumps(make_demo(a.outdir),ensure_ascii=False))
