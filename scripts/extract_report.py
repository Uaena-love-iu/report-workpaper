#!/usr/bin/env python3
"""
extract_report.py — Parse a券商 .docx report into a section/subsection/paragraph
tree (with per-run bold), so the work-paper (底稿) pipeline can render each
paragraph and source it.

Heading levels are detected from styles.xml outlineLvl (0/1/2 -> H1/H2/H3),
NOT hard-coded style ids, so it works across reports. Figure/table captions
are detected by style name (图标题 / 表标题) so the caller knows which
separate "table sheets" are needed.

Usage:
  # 1) list every heading with computed numbering (pick your section)
  python extract_report.py REPORT.docx --list-headings

  # 2) extract one H2 section (matched by keyword) into JSON
  python extract_report.py REPORT.docx --h2-contains 缺电 --label 4.3 \
        --out section.json
"""
import argparse, json, re, sys, zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _val(el, attr=W + "val"):
    return el.get(attr) if el is not None else None


def load_style_levels(z):
    """styleId -> dict(outline=int|None, is_fig=bool, is_tab=bool, is_body=bool)."""
    root = ET.fromstring(z.read("word/styles.xml"))
    out = {}
    for s in root.findall(W + "style"):
        sid = s.get(W + "styleId")
        name_el = s.find(W + "name")
        name = _val(name_el) or ""
        ppr = s.find(W + "pPr")
        olvl = None
        if ppr is not None:
            o = ppr.find(W + "outlineLvl")
            if o is not None and _val(o) is not None:
                olvl = int(_val(o))
        out[sid] = {
            "outline": olvl,
            "is_fig": "图标题" in name or "figure" in name.lower(),
            "is_tab": "表标题" in name or "table" in name.lower(),
            "is_source": "资料来源" in name,
            "name": name,
        }
    return out


def para_style(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None
    st = ppr.find(W + "pStyle")
    return _val(st) if st is not None else None


def _is_bold(rpr):
    if rpr is None:
        return False
    b = rpr.find(W + "b")
    if b is None:
        return False
    v = _val(b)
    return v in (None, "1", "true", "on")


def para_runs(p):
    """Return [(text, bold)] in document order, merging hyperlinks/runs."""
    runs = []
    # iterate over runs anywhere under the paragraph (incl. w:hyperlink)
    for r in p.iter(W + "r"):
        rpr = r.find(W + "rPr")
        bold = _is_bold(rpr)
        txt = "".join(t.text or "" for t in r.iter(W + "t"))
        # include tab as nothing, line breaks as newline
        if txt:
            runs.append([txt, bold])
    # merge consecutive runs with same bold
    merged = []
    for txt, bold in runs:
        if merged and merged[-1][1] == bold:
            merged[-1][0] += txt
        else:
            merged.append([txt, bold])
    return [{"text": t, "bold": b} for t, b in merged]


def para_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def classify(p, levels):
    sid = para_style(p)
    info = levels.get(sid, {})
    olvl = info.get("outline")
    if olvl == 0:
        return "h1"
    if olvl == 1:
        return "h2"
    if olvl == 2:
        return "h3"
    if info.get("is_fig"):
        return "figcap"
    if info.get("is_tab"):
        return "tabcap"
    if info.get("is_source"):
        return "source"
    return "body"


def walk(z, levels):
    """Flat ordered list of {kind,text,runs}."""
    root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    items = []
    for p in body.findall(W + "p"):
        txt = para_text(p).strip()
        kind = classify(p, levels)
        if not txt and kind == "body":
            continue
        items.append({"kind": kind, "text": txt, "runs": para_runs(p)})
    return items


def numbered_headings(items):
    """Assign hierarchical numbers; skip front-matter H1 with no number in TOC.
    We number H1 sequentially but tag front-matter (投资聚焦/目录/插图目录/表格目录)."""
    FRONT = ("投资聚焦", "目录", "插图目录", "表格目录", "摘要")
    h1 = h2 = h3 = 0
    h1_is_front = False
    out = []
    for it in items:
        if it["kind"] == "h1":
            if any(it["text"].startswith(f) or it["text"] == f for f in FRONT):
                h1_is_front = True
                out.append((None, it))
                continue
            h1 += 1
            h2 = h3 = 0
            h1_is_front = False
            out.append((f"{h1}", it))
        elif it["kind"] == "h2" and not h1_is_front:
            h2 += 1
            h3 = 0
            out.append((f"{h1}.{h2}", it))
        elif it["kind"] == "h3" and not h1_is_front:
            h3 += 1
            out.append((f"{h1}.{h2}.{h3}", it))
        else:
            out.append((None, it))
    return out


def list_headings(items):
    for num, it in numbered_headings(items):
        if it["kind"] in ("h1", "h2", "h3"):
            indent = {"h1": "", "h2": "  ", "h3": "    "}[it["kind"]]
            tag = f"[{num}]" if num else "[front]"
            print(f"{indent}{tag} {it['text']}")


def extract_section(items, h2_contains, label):
    """Collect the H2 whose text contains h2_contains, until next H2.
    Group into subsections by H3. Returns dict for JSON."""
    n = len(items)
    start = None
    for i, it in enumerate(items):
        if it["kind"] == "h2" and h2_contains in it["text"]:
            start = i
            break
    if start is None:
        sys.exit(f"未找到包含『{h2_contains}』的二级标题")
    # find end (next h2 or next h1)
    end = n
    for j in range(start + 1, n):
        if items[j]["kind"] in ("h2", "h1"):
            end = j
            break
    section = {
        "label": label,
        "title": items[start]["text"],
        "subsections": [],
    }
    # walk children, bucketing by H3
    cur = {"label": f"{label}.0", "title": "（节首/引言）", "kind": "intro",
           "paragraphs": [], "figures": [], "tables": []}
    sub_idx = 0

    def flush(cur):
        if cur["paragraphs"] or cur["figures"] or cur["tables"]:
            section["subsections"].append(cur)

    for k in range(start + 1, end):
        it = items[k]
        if it["kind"] == "h3":
            flush(cur)
            sub_idx += 1
            cur = {"label": f"{label}.{sub_idx}", "title": it["text"],
                   "kind": "h3", "paragraphs": [], "figures": [], "tables": []}
        elif it["kind"] == "body":
            cur["paragraphs"].append({"text": it["text"], "runs": it["runs"]})
        elif it["kind"] == "figcap":
            cur["figures"].append(it["text"])
        elif it["kind"] == "tabcap":
            cur["tables"].append(it["text"])
        # ignore 'source' notes
    flush(cur)
    return section


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--list-headings", action="store_true")
    ap.add_argument("--h2-contains")
    ap.add_argument("--label", default="X")
    ap.add_argument("--out")
    a = ap.parse_args()

    z = zipfile.ZipFile(a.docx)
    levels = load_style_levels(z)
    items = walk(z, levels)

    if a.list_headings:
        list_headings(items)
        return
    if not a.h2_contains:
        sys.exit("需要 --list-headings 或 --h2-contains")
    section = extract_section(items, a.h2_contains, a.label)
    js = json.dumps(section, ensure_ascii=False, indent=2)
    if a.out:
        open(a.out, "w").write(js)
        n_sub = len(section["subsections"])
        n_par = sum(len(s["paragraphs"]) for s in section["subsections"])
        print(f"写出 {a.out}: {n_sub} 个小节, {n_par} 段正文")
        for s in section["subsections"]:
            print(f"  {s['label']} {s['title']}  "
                  f"({len(s['paragraphs'])}段, 图{len(s['figures'])}, 表{len(s['tables'])})")
    else:
        print(js)


if __name__ == "__main__":
    main()
