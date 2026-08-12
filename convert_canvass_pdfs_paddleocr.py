#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PaddleOCR-VL (AI Studio cloud) extraction backend for the 2026 AL canvass PDFs.

Why this backend: measured on Perry, PaddleOCR-VL-1.6 reads the vote *digits* at
reference quality (36/38 REP candidate county-totals exact, including the Horn=71
the nuextract baseline misread as 70) — on par with Claude and well above the
local nuextract3 baseline, at local/cheap cost. Unlike the per-page Ollama/Claude
backends, PaddleOCR's AI Studio endpoint takes a whole PDF as one async job and
returns per-page markdown, so a county is one submission, not N page calls.

Like every other backend it only swaps the page-image-to-markdown step: the
markdown is cached per page and fed to convert_canvass_pdfs.parse_page unchanged,
so the same checksum gate and totals-join naming apply. PaddleOCR lays its tables
out differently from nuextract/Claude, so a thin normalizer (`_normalize_md`)
rewrites its output into the structure parse_page expects before caching.

Auth: reads the AI Studio token from $PADDLEOCR_TOKEN or the git-ignored
`.paddleocr_token` file. Get a token from the PaddleOCR AI Studio console.

    python convert_canvass_pdfs_paddleocr.py <pdf> [<pdf> ...] [--validate-only]

Output goes to 2026/counties_paddleocr/ — never the verified 2026/counties/.
It also plugs into repair_canvass_contests.py:  repair <pdf> --model paddleocr
"""

import argparse
import glob
import json
import os
import re
import sys
import time

import pandas as pd

import convert_canvass_pdfs as base

JOB_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
MODEL = "PaddleOCR-VL-1.6"
CACHE_DIR = os.environ.get("PADDLEOCR_CACHE", ".canvass_cache_paddleocr")
OPTIONAL_PAYLOAD = {
    "useDocOrientationClassify": False,
    "useDocUnwarping": False,
    "useChartRecognition": False,
}
POLL_SECONDS = 5


def _token():
    tok = os.environ.get("PADDLEOCR_TOKEN")
    if tok:
        return tok.strip()
    for path in (".paddleocr_token", os.path.expanduser("~/.paddleocr_token")):
        if os.path.exists(path):
            return open(path).read().strip()
    raise SystemExit("No PaddleOCR token: set $PADDLEOCR_TOKEN or create .paddleocr_token")


# ---------------------------------------------------------------------------
# AI Studio job: submit whole PDF -> poll -> per-page markdown
# ---------------------------------------------------------------------------

def _submit_job(pdf_path, token):
    import requests
    headers = {"Authorization": f"bearer {token}"}
    data = {"model": MODEL, "optionalPayload": json.dumps(OPTIONAL_PAYLOAD)}
    with open(pdf_path, "rb") as f:
        r = requests.post(JOB_URL, headers=headers, data=data, files={"file": f}, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"PaddleOCR submit failed ({r.status_code}): {r.text[:400]}")
    return r.json()["data"]["jobId"]


def _poll_job(job_id, token):
    """Block until the job is done; return the JSONL result URL."""
    import requests
    headers = {"Authorization": f"bearer {token}"}
    while True:
        r = requests.get(f"{JOB_URL}/{job_id}", headers=headers, timeout=60)
        r.raise_for_status()
        d = r.json()["data"]
        state = d["state"]
        if state == "done":
            return d["resultUrl"]["jsonUrl"]
        if state == "failed":
            raise RuntimeError(f"PaddleOCR job failed: {d.get('errorMsg')}")
        prog = d.get("extractProgress") or {}
        if state == "running" and "totalPages" in prog:
            print(f"    paddleocr running {prog.get('extractedPages')}/{prog['totalPages']} pages",
                  flush=True)
        time.sleep(POLL_SECONDS)


def _fetch_pages(jsonl_url):
    """Download the JSONL and return per-page markdown text, in page order."""
    import requests
    r = requests.get(jsonl_url, timeout=120)
    r.raise_for_status()
    pages = []
    for line in r.text.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        for res in json.loads(line)["result"]["layoutParsingResults"]:
            pages.append(res["markdown"]["text"])
    return pages


# ---------------------------------------------------------------------------
# Normalize PaddleOCR markdown into the shape parse_page expects
# ---------------------------------------------------------------------------

# PaddleOCR emits HTML tables with styled cells and colspans; parse_page's CELL
# regex already tolerates attributes, but PaddleOCR also (a) wraps whole matrices
# in one <table> with the vertical name-lattice rows and a trailing
# "CANDIDATE PERCENT" row, and (b) sometimes carries the party only as page text
# ("ALABAMA REPUBLICAN P"). parse_page already handles the party-header text and
# CANDIDATE TOTALS rows, and ignores non-numeric lattice rows, so the main thing
# to strip is the CANDIDATE PERCENT row, whose decimals otherwise look like extra
# candidate columns to the totals logic.
_PERCENT_ROW = re.compile(r"<tr[^>]*>(?:(?!</tr>).)*?CANDIDATE\s+PERCENT.*?</tr>", re.S | re.I)

# PaddleOCR is inconsistent about the "(VOTE FOR) N" marker parse_page anchors
# each contest's office on: sometimes it's a standalone line above the table
# (fine), sometimes it's packed as the first cell INSIDE the table — on the same
# text line as the <table> tag. In the inline case the office-anchor position
# ties the table's start, and parse_page's strict "office position < table
# position" lookup drops the contest (office=None), losing e.g. Governor on a
# page it shares with Lt Governor. Hoist an inline marker out to its own line
# just before the table so the anchor sits strictly before it.
_INLINE_VOTEFOR = re.compile(
    r"(<table\b[^>]*>\s*<tr[^>]*>\s*<td[^>]*>)\s*(\(VOTE\s*FOR\)\s*\d*)\s*(</td>)", re.I)


def _normalize_md(md):
    md = _PERCENT_ROW.sub("", md)
    md = _INLINE_VOTEFOR.sub(lambda m: f"{m.group(2)}\n{m.group(1)}{m.group(2)}{m.group(3)}", md)
    return md


# ---------------------------------------------------------------------------
# extract_fn (matches convert_canvass_pdfs.extract_pages' contract)
# ---------------------------------------------------------------------------

def make_extract_fn():
    token = _token()

    def extract_pages(pdf_path, dpi):  # dpi ignored: PaddleOCR rasterizes server-side
        stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
        cache = os.path.join(CACHE_DIR, stem)
        os.makedirs(cache, exist_ok=True)

        # Cache the RAW PaddleOCR markdown and apply _normalize_md() on read, so
        # tightening the normalizer/parser never forces a re-fetch from the API.
        done_marker = os.path.join(cache, ".complete")
        if os.path.exists(done_marker):
            for md in sorted(glob.glob(os.path.join(cache, "p*.md"))):
                page = int(re.search(r"p(\d+)\.md$", md).group(1))
                yield page, _normalize_md(open(md).read())
            return

        print(f"    paddleocr submitting {os.path.basename(pdf_path)} ...", flush=True)
        job_id = _submit_job(pdf_path, token)
        jsonl_url = _poll_job(job_id, token)
        pages = _fetch_pages(jsonl_url)
        for i, raw in enumerate(pages, start=1):
            open(os.path.join(cache, f"p{i:03d}.md"), "w").write(raw)
        open(done_marker, "w").write(str(len(pages)))
        print(f"    paddleocr got {len(pages)} pages", flush=True)
        for i, raw in enumerate(pages, start=1):
            yield i, _normalize_md(raw)

    return extract_pages


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="canvass PDF path(s); globs allowed")
    ap.add_argument("--validate-only", action="store_true")
    args = ap.parse_args()

    extract_fn = make_extract_fn()
    out_dir = os.path.join("2026", "counties_paddleocr")
    county_df = pd.read_csv(base.COUNTY_CSV, dtype=str, keep_default_na=False)
    county_df["votes"] = county_df["votes"].astype(int)

    paths = []
    for p in args.pdfs:
        paths.extend(sorted(glob.glob(p)) if any(c in p for c in "*?[") else [p])
    by_county = {}
    for p in paths:
        if not os.path.exists(p):
            print(f"skip (missing): {p}", file=sys.stderr)
            continue
        by_county.setdefault(base.detect_county(p), []).append(p)

    print(f"model: {MODEL} (PaddleOCR AI Studio)", flush=True)
    tp = tf = 0
    for county_paths in by_county.values():
        a, b = base.process(sorted(county_paths), 200, county_df, args.validate_only,
                            extract_fn=extract_fn, out_dir=out_dir)
        tp += a
        tf += b
    print(f"\n==== {len(by_county)} counties: {tp} verified, {tf} failed ====")
    return 1 if tf else 0


if __name__ == "__main__":
    sys.exit(main())
