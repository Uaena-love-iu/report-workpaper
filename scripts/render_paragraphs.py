#!/usr/bin/env python3
"""
render_paragraphs.py — Render each report paragraph (from extract_report.py JSON)
into a clean PNG that mimics the report's look: red bold sub-heading, bold lead
sentence, justified body with first-line indent. These become the LEFT column of
the 底稿 (work paper).

One PNG per paragraph. The first paragraph of each subsection is rendered with
its sub-heading on top (matching the analyst's existing 底稿, e.g. image "4.1.1
国内…" + first paragraph in one shot).

Usage:
  python render_paragraphs.py section.json --outdir imgs/ [--width 760]

Output: imgs/<label>_p<n>.png  + updates a manifest fragment printed as JSON:
  {"4.3.1": [{"idx":1,"img":".../4.3.1_p1.png","text":"...","lead":"..."}, ...]}
"""
import argparse, json, os, sys, html
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

CSS = """
*{margin:0;padding:0;box-sizing:border-box}
body{background:#fff}
.page{width:%(width)dpx;background:#fff;padding:18px 22px 22px 22px;
      font-family:"PingFang SC","Microsoft YaHei","Heiti SC",sans-serif;}
.head{color:#C00000;font-weight:700;font-size:21px;line-height:1.5;
      margin-bottom:12px;letter-spacing:.3px;}
.head .num{margin-right:14px;}
p.body{font-size:17px;line-height:1.85;color:#1a1a1a;text-align:justify;
       text-justify:inter-ideograph;text-indent:2em;}
b{font-weight:700;color:#000;}
"""

HTML = """<!doctype html><html><head><meta charset="utf-8">
<style>%(css)s</style></head><body>
<div class="page" id="page">%(inner)s</div></body></html>"""


def runs_to_html(runs, fallback_text):
    if not runs:
        return html.escape(fallback_text)
    out = []
    for r in runs:
        t = html.escape(r["text"])
        out.append(f"<b>{t}</b>" if r.get("bold") else t)
    return "".join(out)


def head_html(label, title):
    return (f'<div class="head"><span class="num">{html.escape(label)}</span>'
            f'{html.escape(title)}</div>')


def render(section, outdir, width):
    os.makedirs(outdir, exist_ok=True)
    manifest = {}
    with sync_playwright() as pw:
        executable = browser_executable()
        browser = pw.chromium.launch(executable_path=executable) if executable else pw.chromium.launch()
        page = browser.new_page(device_scale_factor=3)
        for sub in section["subsections"]:
            label = sub["label"]
            entries = []
            for i, para in enumerate(sub["paragraphs"], 1):
                inner = ""
                if i == 1 and sub.get("kind") != "intro":
                    inner += head_html(label, sub["title"])
                inner += f'<p class="body">{runs_to_html(para["runs"], para["text"])}</p>'
                doc = HTML % {"css": CSS % {"width": width}, "inner": inner}
                page.set_content(doc, wait_until="networkidle")
                el = page.query_selector("#page")
                fn = os.path.join(outdir, f"{label}_p{i}.png")
                el.screenshot(path=fn)
                lead = "".join(r["text"] for r in para["runs"] if r.get("bold"))
                entries.append({"idx": i, "img": os.path.abspath(fn),
                                "text": para["text"], "lead": lead})
            manifest[label] = {"title": sub["title"], "paragraphs": entries}
        browser.close()
    return manifest


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("section_json")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--width", type=int, default=760)
    ap.add_argument("--out-manifest")
    a = ap.parse_args()
    section = json.load(open(a.section_json))
    manifest = render(section, a.outdir, a.width)
    js = json.dumps(manifest, ensure_ascii=False, indent=2)
    if a.out_manifest:
        open(a.out_manifest, "w").write(js)
        n = sum(len(v["paragraphs"]) for v in manifest.values())
        print(f"渲染 {n} 段 -> {a.outdir}")
    else:
        print(js)


if __name__ == "__main__":
    main()
