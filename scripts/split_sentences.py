#!/usr/bin/env python3
"""
split_sentences.py — Split each paragraph of a section.json into sentences and
emit a per-sentence COVERAGE checklist, so the semi-auto sourcing pass can make
sure EVERY factual sentence gets a source (nothing slips through).

Sentences are split on 。！？ (and Chinese ； kept as its own clause when long).
Each sentence is heuristically tagged:
  - "opinion"  : contains judgement wording; factual premises still require
                 evidence review. This tag never exempts the sentence.
  - "factual"  : carries a checkable fact (number/%/date/名称/事件) -> needs a source.
The tag is a HINT; the human/agent decides. Output is meant to be filled in:
add {"url","title","img","verbatim"} per factual sentence, then feed into the
build manifest (sentence order == left->right order).

Usage:
  python split_sentences.py section.json --out coverage.json
  python split_sentences.py section.json            # pretty print to stdout
"""
import argparse, json, re, sys

OPINION_CUES = ("我们认为", "我们看好", "我们预计", "我们判断", "有望", "看好",
                "预计", "建议关注", "我们", "重视", "优选")
FACT_CUES = ("%", "亿", "万", "GW", "MW", "TWh", "倍", "年", "月", "占",
             "达", "约", "超", "根据", "据", "截至", "份额", "同比", "环比")
NUM = re.compile(r"[0-9０-９]")


def split_sentences(text):
    # split keeping the terminator with the sentence
    parts = re.split(r"(?<=[。！？])", text)
    out = []
    for p in parts:
        s = p.strip()
        if s:
            out.append(s)
    return out


def classify(s):
    has_num = bool(NUM.search(s))
    has_fact = has_num or any(c in s for c in FACT_CUES)
    is_op = any(c in s for c in OPINION_CUES)
    # a sentence can be both (opinion grounded in a number) -> treat as factual
    if has_fact:
        return "factual"
    if is_op:
        return "opinion"
    return "factual"  # default to needing a source (conservative)


def build(section):
    out = {"label": section["label"], "title": section["title"], "subsections": []}
    for sub in section["subsections"]:
        srows = []
        for pi, para in enumerate(sub["paragraphs"], 1):
            sents = split_sentences(para["text"])
            slist = []
            for si, s in enumerate(sents, 1):
                slist.append({
                    "para": pi, "sent": si, "kind": classify(s),
                    "text": s,
                    "review_required": True,
                    "sources": []  # fill: [{"url","title","img","verbatim"}]
                })
            srows.append({"para": pi, "sentences": slist})
        out["subsections"].append(
            {"label": sub["label"], "title": sub["title"], "paragraphs": srows})
    return out


def pretty(cov):
    for sub in cov["subsections"]:
        print(f"\n### {sub['label']} {sub['title']}")
        for para in sub["paragraphs"]:
            print(f"  [P{para['para']}]")
            for s in para["sentences"]:
                flag = "📌需源" if s["kind"] == "factual" else "💭观点"
                print(f"    {flag} P{s['para']}S{s['sent']}: {s['text'][:54]}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("section_json")
    ap.add_argument("--out")
    a = ap.parse_args()
    section = json.load(open(a.section_json))
    cov = build(section)
    if a.out:
        json.dump(cov, open(a.out, "w"), ensure_ascii=False, indent=2)
        nf = sum(1 for sub in cov["subsections"] for p in sub["paragraphs"]
                 for s in p["sentences"] if s["kind"] == "factual")
        nt = sum(1 for sub in cov["subsections"] for p in sub["paragraphs"]
                 for s in p["sentences"])
        print(f"写出 {a.out}: 共 {nt} 句, 其中需找源(factual) {nf} 句")
        pretty(cov)
    else:
        pretty(cov)


if __name__ == "__main__":
    main()
