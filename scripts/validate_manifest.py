import argparse,json
from workpaper_model import load_manifest,validate_manifest
def main():
    p=argparse.ArgumentParser(description='Validate workpaper coverage and evidence metadata');p.add_argument('manifest');a=p.parse_args()
    try:print(json.dumps({'valid':True,**validate_manifest(load_manifest(a.manifest))},ensure_ascii=False))
    except (ValueError,OSError)as e:p.exit(1,str(e)+'\n')
if __name__=='__main__':main()
