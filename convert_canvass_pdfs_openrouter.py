#!/usr/bin/env python3
"""
OpenRouter vision-model extraction — a drop-in alternative to nuextract3 for
testing whether a different (typically larger/cloud) model does better on the
same canvass-matrix PDFs.

Reuses every parsing/checksum/naming/never-drop function from
convert_canvass_pdfs.py completely unchanged via process()'s injectable
extract_fn — only the page-image-to-markdown extraction step is swapped, so
results are directly comparable using the exact same PASS/FAIL checksum
criteria and the same CSV output format. Any future fix to the shared parsing
logic in convert_canvass_pdfs.py benefits both scripts automatically.

Requires an OpenRouter API key: https://openrouter.ai/keys

    export OPENROUTER_API_KEY=sk-or-...
    python convert_canvass_pdfs_openrouter.py <pdf> [<pdf> ...] \
        [--model MODEL] [--dpi 200] [--validate-only]

Model defaults to $OPENROUTER_MODEL or google/gemini-2.5-flash if unset — a
cheap, fast, vision-capable model. Any vision model on
https://openrouter.ai/models works; try e.g. openai/gpt-4o or
qwen/qwen2.5-vl-72b-instruct to compare.
"""

import argparse
import base64
import glob
import os
import re
import sys

import pandas as pd
import requests
from natural_pdf import PDF

import convert_canvass_pdfs as base

API_URL = "https://openrouter.ai/api/v1/chat/completions"
DEFAULT_MODEL = os.environ.get("OPENROUTER_MODEL", "google/gemini-2.5-flash")

# Mirrors the structural conventions convert_canvass_pdfs.parse_page() already
# expects from nuextract3's output, so the exact same regexes/parsing can run
# against a completely different model's transcription unmodified.
PROMPT = """Transcribe this election canvass report page exactly, preserving all structure. Output GitHub-flavored Markdown with these rules:

- Any office/contest title (e.g. "GOVERNOR", "LIEUTENANT GOVERNOR") on its own line, exactly as printed, followed by "(VOTE FOR) 1" on its own line if present.
- If the page says "(CONTINUED FROM PREVIOUS PAGE)", include that exact text on its own line.
- Render every vote table as a proper HTML <table> with <thead> and <tbody>, using one <td> per cell — one candidate's vote count per <td>, never multiple numbers combined into one cell.
- Each precinct row's first cell must be the exact precinct code and name as printed (e.g. "0001 SOME PRECINCT NAME").
- The totals row's first cell must be exactly "CANDIDATE TOTALS", followed by each candidate's total in its own <td>, in the same left-to-right order as the precinct rows' columns.
- Do not summarize, skip, or omit any office, precinct, or candidate — transcribe every one visible on the page, including any you are not fully certain about.
- Do not describe the image or add commentary. Output only the transcribed content.
"""


def _image_data_uri(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f"data:image/png;base64,{b64}"


def make_extract_fn(model, api_key):
    """Build an extract_pages(pdf_path, dpi) matching convert_canvass_pdfs's
    signature, bound to one OpenRouter model/key, with its own cache
    namespace so different models' outputs never collide."""
    model_slug = re.sub(r"[^A-Za-z0-9]+", "_", model)

    def extract_pages(pdf_path, dpi):
        stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
        cache = os.path.join(f".canvass_cache_openrouter_{model_slug}", f"{stem}_{dpi}")
        os.makedirs(cache, exist_ok=True)

        pdf = PDF(pdf_path)
        for i in range(1, len(pdf.pages) + 1):
            md_path = os.path.join(cache, f"p{i:03d}.md")
            if os.path.exists(md_path):
                yield i, open(md_path).read()
                continue
            png = os.path.join(cache, f"p{i:03d}.png")
            if not os.path.exists(png):
                pdf.pages[i - 1].render(resolution=dpi).save(png)

            r = requests.post(
                API_URL,
                headers={"Authorization": f"Bearer {api_key}"},
                json={
                    "model": model,
                    "temperature": 0,
                    "messages": [{
                        "role": "user",
                        "content": [
                            {"type": "text", "text": PROMPT},
                            {"type": "image_url", "image_url": {"url": _image_data_uri(png)}},
                        ],
                    }],
                },
                timeout=180,
            )
            if not r.ok:
                raise RuntimeError(f"OpenRouter request failed ({r.status_code}): {r.text[:500]}")
            data = r.json()
            if "choices" not in data:
                raise RuntimeError(f"OpenRouter response missing choices: {data}")
            txt = data["choices"][0]["message"]["content"] or ""
            open(md_path, "w").write(txt)
            print(f"    p{i:03d} extracted ({model})", flush=True)
            yield i, txt

    return extract_pages


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="canvass PDF path(s); globs allowed")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"OpenRouter model id (default: {DEFAULT_MODEL})")
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--validate-only", action="store_true")
    args = ap.parse_args()

    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        print("OPENROUTER_API_KEY is not set. Get a key at https://openrouter.ai/keys "
              "and run: export OPENROUTER_API_KEY=sk-or-...", file=sys.stderr)
        return 1

    extract_fn = make_extract_fn(args.model, api_key)
    # Never write into the shared 2026/counties/ output — this script is for
    # comparing extraction quality, and a model under test must not silently
    # overwrite an already-verified nuextract3 CSV for the same county.
    model_slug = re.sub(r"[^A-Za-z0-9]+", "_", args.model)
    out_dir = os.path.join("2026", f"counties_openrouter_{model_slug}")

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

    print(f"model: {args.model}", flush=True)
    if not args.validate_only:
        print(f"output: {out_dir}/ (kept separate from 2026/counties/ on purpose)", flush=True)
    tp = tf = 0
    for county_paths in by_county.values():
        a, b = base.process(sorted(county_paths), args.dpi, county_df, args.validate_only,
                            extract_fn=extract_fn, out_dir=out_dir)
        tp += a
        tf += b
    print(f"\n==== {len(by_county)} count{'y' if len(by_county) == 1 else 'ies'} "
          f"({len(paths)} PDF file(s)), model={args.model}: "
          f"{tp} contests verified, {tf} failed ====")
    return 1 if tf else 0


if __name__ == "__main__":
    sys.exit(main())
