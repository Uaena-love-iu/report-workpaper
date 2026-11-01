#!/usr/bin/env python3
"""
capture_source.py — Open a source URL, locate the exact supporting phrase(s),
draw a red box around them, and save a TIGHT, READABLE screenshot for the RIGHT
column of the 底稿. Cross-node text matching + whitespace-insensitive search make
it robust to markup. If a phrase is not found, it is reported (box=null) so the
caller can fall back / pick a better phrase.

WHY THE CROP IS TIGHT (don't loosen casually):
  The 底稿 shows every source image at a FIXED display width (build_workpaper
  SW≈540–600px). On-screen text height ≈ CSS_font * SW / crop_width_css. So a
  wide full-page crop (1180px) shrinks 17px text to ~8px (unreadable). We crop
  HORIZONTALLY to the highlighted text's own column/paragraph block (≈600–800px)
  and VERTICALLY to the highlighted line(s) + a couple lines of context. Result:
  the boxed sentence fills the frame and reads clearly. Never screenshot the whole
  page just to box one sentence — capture the sentence + a little context only.

If highlights are far apart (would exceed --max-h), only the TOP cluster is
captured and the rest are reported in "off_crop" — split them into separate
captures (one tight image per sentence) instead of one tall blurry image.

Headless by default. For login-gated / anti-bot pages (微信公众号, 小红书…),
reuse your real logged-in Chrome:
    1) start Chrome with  --remote-debugging-port=9222   (or enable it via the
       web-access skill), then
    2) pass  --cdp http://localhost:9222

Usage:
  python capture_source.py --url URL --out shot.png \
      --highlight "5,427个" --highlight "占全球总量的45%" \
      [--title "标题"] [--cdp http://localhost:9222] [--full] \
      [--crop-w 760] [--context 130] [--max-h 1500]

Prints JSON: {"url","out","matched":[...],"unmatched":[...],"off_crop":[...],"ok":bool}
"""
import argparse, base64, json, os, sys
from pathlib import Path
from playwright.sync_api import sync_playwright


def browser_executable():
    configured = os.environ.get("REPORT_WORKPAPER_BROWSER")
    if configured and Path(configured).exists():
        return configured
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return candidate
    return None

# Returns, per target: the red-box bbox (page coords) AND the bounds of the
# nearest BLOCK-level ancestor (the text column / paragraph) so we can crop
# horizontally to the readable column instead of the whole viewport.
HIGHLIGHT_JS = r"""
(targets) => {
  const norm = s => (s||'').replace(/\s+/g,'');
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null);
  const nodes = [];
  let n;
  while (n = walker.nextNode()) {
    const v = n.nodeValue;
    if (!v || !v.trim()) continue;
    const pe = n.parentElement;
    if (!pe) continue;
    const st = getComputedStyle(pe);
    if (st.display === 'none' || st.visibility === 'hidden') continue;
    let visible=true;
    for(let ancestor=pe;ancestor;ancestor=ancestor.parentElement){
      const css=getComputedStyle(ancestor);
      if(css.display==='none'||css.visibility==='hidden'||Number(css.opacity)===0){visible=false;break;}
    }
    if(!visible||!pe.getClientRects().length)continue;
    if (['SCRIPT','STYLE','NOSCRIPT'].includes(pe.tagName)) continue;
    nodes.push(n);
  }
  let g = '';
  const map = [];
  for (const node of nodes) {
    const raw = node.nodeValue;
    for (let i = 0; i < raw.length; i++) {
      if (!/\s/.test(raw[i])) { g += raw[i]; map.push([node, i]); }
    }
  }
  const BLOCK = new Set(['P','DIV','LI','TD','TH','BLOCKQUOTE','SECTION','ARTICLE',
                         'DD','DT','FIGCAPTION','H1','H2','H3','H4','H5','H6','PRE']);
  function blockRect(node){
    let el = node.parentElement;
    while (el && el !== document.body){
      const st = getComputedStyle(el);
      if (BLOCK.has(el.tagName) || st.display==='block' || st.display==='list-item'
          || st.display==='table-cell'){
        const r = el.getBoundingClientRect();
        if (r.width >= 120) return r;   // ignore tiny inline-ish wrappers
      }
      el = el.parentElement;
    }
    return null;
  }
  const out = [];
  for (const t of targets) {
    const tt = norm(t);
    if (!tt) { out.push({t, box:null}); continue; }
    const pos = g.indexOf(tt);
    if (pos < 0) { out.push({t, box:null}); continue; }
    const [sn, so] = map[pos];
    const [en, eo] = map[pos + tt.length - 1];
    const range = document.createRange();
    try { range.setStart(sn, so); range.setEnd(en, eo + 1); }
    catch(e){ out.push({t, box:null}); continue; }
    const rects = [...range.getClientRects()];
    if (!rects.length) { out.push({t, box:null}); continue; }
    let L=1e9,T=1e9,R=-1e9,B=-1e9;
    for (const r of rects) {
      const d = document.createElement('div');
      d.style.cssText = 'position:absolute;left:'+(r.left+scrollX)+'px;top:'+
        (r.top+scrollY)+'px;width:'+r.width+'px;height:'+r.height+
        'px;border:2.5px solid #e60000;background:rgba(255,229,0,.30);'+
        'z-index:2147483647;pointer-events:none;border-radius:2px;'+
        'box-sizing:border-box;';
      d.className = '__wp_hl';
      document.body.appendChild(d);
      L=Math.min(L,r.left+scrollX); T=Math.min(T,r.top+scrollY);
      R=Math.max(R,r.right+scrollX); B=Math.max(B,r.bottom+scrollY);
    }
    const br = blockRect(sn);
    const block = br ? {left: br.left+scrollX, right: br.right+scrollX} : null;
    out.push({t, box:{left:L,top:T,right:R,bottom:B, block}});
  }
  return {out, pageW: document.documentElement.scrollWidth,
          pageH: document.documentElement.scrollHeight, dpr: devicePixelRatio};
}
"""

# ---- crop defaults (tight & readable; see module docstring) ----
CONTEXT = 130      # vertical context above & below the highlighted line(s), CSS px
MAX_H = 1500       # cap crop height (CSS px); beyond this, split into more captures
HPAD = 46          # horizontal context beyond the text column, CSS px
MIN_W = 520        # never crop narrower than this (keeps some context)
MAX_W = 1040       # never crop wider than this (keeps text legible)
FULL_W = 1180      # width for --full / no-box fallback
DPR = 3            # device scale factor: higher = sharper (does NOT change on-screen size)


def capture(url, highlights, out, cdp=None, full=False, width=1200,
            timeout=35000, crop_w=None, context=CONTEXT, max_h=MAX_H):
    with sync_playwright() as pw:
        if cdp:
            browser = pw.chromium.connect_over_cdp(cdp)
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            page = ctx.new_page()
        else:
            executable = browser_executable()
            browser = pw.chromium.launch(executable_path=executable) if executable else pw.chromium.launch()
            ctx = browser.new_context(
                viewport={"width": width, "height": 1000},
                device_scale_factor=DPR,
                user_agent=("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/124.0 Safari/537.36"))
            page = ctx.new_page()
        info = {"url": url, "out": os.path.abspath(out),
                "matched": [], "unmatched": [], "off_crop": [], "ok": False}
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            try:
                page.wait_for_load_state("networkidle", timeout=6000)
            except Exception:
                pass
            page.wait_for_timeout(1200)
            res = page.evaluate(HIGHLIGHT_JS, highlights)
            boxes = [r["box"] for r in res["out"] if r["box"]]
            for r in res["out"]:
                (info["matched"] if r["box"] else info["unmatched"]).append(r["t"])
            # neutralize Playwright's "waiting for fonts" hang on ad-heavy pages
            try:
                page.evaluate(
                    "()=>{try{Object.defineProperty(document,'fonts',"
                    "{configurable:true,value:{ready:Promise.resolve(),"
                    "status:'loaded',check:()=>true,load:()=>Promise.resolve(),"
                    "forEach:()=>{},values:()=>[],addEventListener:()=>{},"
                    "removeEventListener:()=>{}}});}catch(e){}}")
            except Exception:
                pass
            page_w = page.evaluate("()=>document.documentElement.scrollWidth")
            page_h = page.evaluate("()=>document.documentElement.scrollHeight")

            if full or not boxes:
                x0, w = 0.0, float(min(page_w, FULL_W))
                y0, h = 0.0, float(min(page_h, max_h))
            else:
                # ---- vertical: highlighted line(s) + a couple lines context ----
                tops = sorted(b["top"] for b in boxes)
                top0 = tops[0]
                y0 = max(0, top0 - context)
                # include as many highlights as fit under max_h; rest -> off_crop
                kept = [b for b in boxes if b["bottom"] <= y0 + max_h - context]
                if not kept:
                    kept = [min(boxes, key=lambda b: b["top"])]
                bot = max(b["bottom"] for b in kept)
                y1 = min(bot + context, page_h)
                if y1 - y0 > max_h:
                    y1 = y0 + max_h
                h = max(80, y1 - y0)
                # report highlights that fell outside this crop
                kepttext = set()
                for r in res["out"]:
                    b = r["box"]
                    if b and (b["top"] < y0 or b["bottom"] > y1):
                        info["off_crop"].append(r["t"])
                # ---- horizontal: crop to the text column / paragraph block ----
                blocks = [b["block"] for b in kept if b.get("block")]
                if crop_w:
                    cx = sum((b["left"] + b["right"]) / 2 for b in kept) / len(kept)
                    x0 = max(0, cx - crop_w / 2); x1 = min(page_w, cx + crop_w / 2)
                elif blocks:
                    bl = min(bk["left"] for bk in blocks)
                    br = max(bk["right"] for bk in blocks)
                    x0 = max(0, bl - HPAD); x1 = min(page_w, br + HPAD)
                else:
                    bl = min(b["left"] for b in kept)
                    br = max(b["right"] for b in kept)
                    x0 = max(0, bl - HPAD); x1 = min(page_w, br + HPAD)
                w = x1 - x0
                if w < MIN_W:                      # widen symmetrically
                    cx = (x0 + x1) / 2
                    x0 = max(0, cx - MIN_W / 2); w = min(page_w, x0 + MIN_W) - x0
                if w > MAX_W:                      # center on the highlights
                    cx = sum((b["left"] + b["right"]) / 2 for b in kept) / len(kept)
                    x0 = max(0, min(cx - MAX_W / 2, page_w - MAX_W)); w = MAX_W

            for item in res['out']:
                box=item.get('box')
                if box and (box['left']<x0 or box['right']>x0+w or box['top']<y0 or box['bottom']>y0+h):
                    if item['t']not in info['off_crop']:info['off_crop'].append(item['t'])

            cdp_s = page.context.new_cdp_session(page)
            shot = cdp_s.send("Page.captureScreenshot", {
                "format": "png",
                "captureBeyondViewport": True,
                "clip": {"x": float(x0), "y": float(y0), "width": float(w),
                         "height": float(h), "scale": float(DPR)}})
            Path(out).parent.mkdir(parents=True,exist_ok=True)
            with open(out, "wb") as fh:
                fh.write(base64.b64decode(shot["data"]))
            info['capture_ok']=True
            info['requires_manual_review']=not bool(highlights)or bool(info['unmatched'])or bool(info['off_crop'])
            info["ok"] = bool(boxes)and not info['requires_manual_review']
        except Exception as e:
            info["error"] = str(e)
        finally:
            try:
                if not cdp:
                    browser.close()
                else:
                    page.close()
            except Exception:
                pass
        return info


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--highlight", action="append", default=[],
                    help="exact supporting phrase (repeatable)")
    ap.add_argument("--title")
    ap.add_argument("--cdp", help="reuse Chrome via CDP, e.g. http://localhost:9222")
    ap.add_argument("--full", action="store_true", help="full page, no box-crop")
    ap.add_argument("--crop-w", type=float, default=None,
                    help="force horizontal crop width (CSS px), centered on highlights")
    ap.add_argument("--context", type=int, default=CONTEXT,
                    help="vertical context above/below highlight (CSS px)")
    ap.add_argument("--max-h", type=int, default=MAX_H, help="cap crop height (CSS px)")
    a = ap.parse_args()
    info = capture(a.url, a.highlight, a.out, cdp=a.cdp, full=a.full,
                   crop_w=a.crop_w, context=a.context, max_h=a.max_h)
    if a.title:
        info["title"] = a.title
    print(json.dumps(info, ensure_ascii=False))


if __name__ == "__main__":
    main()
