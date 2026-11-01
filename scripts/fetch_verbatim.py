#!/usr/bin/env python3
"""
fetch_verbatim.py — Get a page's LITERAL text (not a paraphrase) via Jina Reader
(r.jina.ai), so you can pick `--highlight` phrases that actually exist on the page.

WHY THIS EXISTS: capture_source.py draws the red box by matching a phrase against
the live DOM with whitespace-removed substring search. That part works. What used
to break it is UPSTREAM: reading the page with WebFetch to *choose* the phrase —
WebFetch PARAPHRASES, so the phrase you pick may not literally be on the page.
r.jina.ai returns the page's real characters (Readability→Turndown, mechanical,
not an LLM), key-free at 20 RPM (set JINA_API_KEY for 500 RPM + anti-bot proxy).

So the per-sentence flow becomes:
    1) fetch_verbatim.py --url URL          # read the REAL page text
    2) pick short, digit-bearing substrings of that text as --highlight
    3) fetch_verbatim.py --url URL --find "候选短语"   # confirm it will match
    4) capture_source.py --url URL --highlight "候选短语" ...   # red-box screenshot

Validation in step 3 mirrors capture_source.py's HIGHLIGHT_JS norm exactly
(strip ALL whitespace, then substring), so a `matched` phrase here will red-box
there. `fold_only` flags phrases that match only after full/half-width folding —
fix the width (e.g. ５,４２７ → 5,427) or they won't box.

Usage:
  python fetch_verbatim.py --url URL [--out page.md]
  python fetch_verbatim.py --url URL --find "5,427个" --find "占全球总量的45%"

Prints the markdown (default) or, with --find, JSON:
  {"url","http","bytes","matched":[...],"unmatched":[...],"fold_only":[...],"ok"}

Anti-bot pages (403 / "Just a moment…") → falls back: use capture_source.py
--cdp to read the live DOM instead (it handles login-gated / JS-heavy pages).
"""
import argparse, json, os, sys, unicodedata
import urllib.parse, urllib.request, urllib.error

READER = "https://r.jina.ai/"
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
# keep URL structure intact, percent-encode only non-ASCII (e.g. Chinese paths)
URL_SAFE = ":/?#[]@!$&'()*+,;=-._~%"


def norm_ws(s):
    """Mirror capture_source.py HIGHLIGHT_JS: remove ALL whitespace, nothing else."""
    return "".join((s or "").split())


def norm_fold(s):
    """Whitespace-removed AND NFKC-folded (full-width→half-width). Used only to
    detect width mismatches the strict matcher would miss."""
    return "".join(unicodedata.normalize("NFKC", s or "").split())


def fetch(url, key=None, timeout=60):
    req_url = READER + urllib.parse.quote(url, safe=URL_SAFE)
    headers = {"User-Agent": UA, "Accept": "text/markdown",
               "X-Return-Format": "markdown"}
    if key:
        headers["Authorization"] = "Bearer " + key
    req = urllib.request.Request(req_url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8", "replace")
            return resp.status, body
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace") if e.fp else ""
    except Exception as e:
        return 0, "__ERROR__ " + str(e)


def looks_blocked(http, body):
    if http != 200 or body.startswith("__ERROR__"):
        return True
    low = body[:1500].lower()
    return ("just a moment" in low or "enable javascript" in low
            or "captcha" in low) and len(body) < 2000


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True)
    ap.add_argument("--out", help="write markdown here (default: stdout)")
    ap.add_argument("--find", action="append", default=[],
                    help="candidate highlight phrase to validate (repeatable)")
    ap.add_argument("--key", help="Jina API key (else env JINA_API_KEY; else key-free)")
    a = ap.parse_args()
    key = a.key or os.environ.get("JINA_API_KEY")

    http, body = fetch(a.url, key=key)

    if a.find:
        info = {"url": a.url, "http": http, "bytes": len(body.encode("utf-8")),
                "matched": [], "unmatched": [], "fold_only": [], "ok": False}
        if looks_blocked(http, body):
            info["error"] = "blocked/anti-bot; use capture_source.py --cdp to read live DOM"
            print(json.dumps(info, ensure_ascii=False))
            sys.exit(2)
        hay_ws = norm_ws(body)
        hay_fold = norm_fold(body)
        for ph in a.find:
            if norm_ws(ph) and norm_ws(ph) in hay_ws:
                info["matched"].append(ph)
            elif norm_fold(ph) and norm_fold(ph) in hay_fold:
                info["fold_only"].append(ph)   # exists, but width differs → fix it
            else:
                info["unmatched"].append(ph)
        info["ok"] = bool(info["matched"]) and not info["unmatched"]
        print(json.dumps(info, ensure_ascii=False))
        sys.exit(0 if info["ok"] else 1)

    if looks_blocked(http, body):
        print(json.dumps({"url": a.url, "http": http,
                          "error": "blocked/anti-bot; use capture_source.py --cdp",
                          "ok": False}, ensure_ascii=False), file=sys.stderr)
        sys.exit(2)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as fh:
            fh.write(body)
        print(json.dumps({"url": a.url, "http": http, "out": os.path.abspath(a.out),
                          "bytes": len(body.encode("utf-8")), "ok": True},
                         ensure_ascii=False))
    else:
        sys.stdout.write(body)


if __name__ == "__main__":
    main()
