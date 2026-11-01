"""Read-only dependency discovery; never installs software or changes browsers."""
import importlib.util,json,os,sys,shutil
from pathlib import Path
def diagnose():
    paths=[os.environ.get('REPORT_WORKPAPER_BROWSER',''),'C:/Program Files/Google/Chrome/Application/chrome.exe','C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',shutil.which('google-chrome')or'',shutil.which('chromium')or'']
    modules={name:importlib.util.find_spec(name)is not None for name in ['openpyxl','PIL','lxml','playwright','pdfplumber','pypdfium2']}
    return {'python':sys.executable,'version':sys.version.split()[0],'modules':modules,'browser':next((p for p in paths if p and Path(p).is_file()),None),'core_ready':all(modules[n]for n in ['openpyxl','PIL','lxml']),'verification_boundary':'Dependency availability is not end-to-end workpaper validation.'}
if __name__=='__main__':print(json.dumps(diagnose(),ensure_ascii=False,indent=2))
