#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Anthropic Claude vision extraction — the paid escalation tier for canvass pages
the PaddleOCR backend can't get past the checksum on.

This only swaps the page-image-to-markdown step: every parse/stitch/checksum/
naming/never-drop rule in convert_canvass_pdfs.py is reused unchanged through
process()'s injectable extract_fn, so a Claude run is scored by the exact same
PASS/FAIL checksum gate and writes the identical CSV shape. It plugs into
repair_canvass_contests.py the same way (make_extract_fn), so the contest-level,
checksum-gated merge applies here too — Claude output is only merged into
2026/counties/ when it passes.

Two modes:

  * blind      — send just the page image (default).
  * correction — send the page image *and* the noisy nuextract3 markdown for the
                 same page (from .canvass_cache), asking Claude to correct it.
                 Cheaper and often more accurate than blind OCR on these matrix
                 pages. (Earlier iterations also tried this correction idea via
                 the now-removed hybrid/openrouter/ollama backends.)

Uses Simon Willison's `llm` library (a repo dependency) for the API call, so it
reuses the Anthropic key already stored in `llm keys` — no ANTHROPIC_API_KEY in
the environment required. Confirm the key with `llm keys` and the model list
with `llm models | grep anthropic`.

    python convert_canvass_pdfs_claude.py <pdf> [<pdf> ...] \\
        [--model anthropic/claude-sonnet-4-6] [--dpi 300] [--validate-only]

Default model is $CLAUDE_MODEL or anthropic/claude-sonnet-4-6. Output goes to
2026/counties_claude_<model>/ — never the verified 2026/counties/ directly.
"""

import argparse
import glob
import os
import re
import sys

import pandas as pd
from natural_pdf import PDF

import convert_canvass_pdfs as base

DEFAULT_MODEL = os.environ.get("CLAUDE_MODEL", "anthropic/claude-sonnet-4-6")
MAX_TOKENS = int(os.environ.get("CLAUDE_MAX_TOKENS", "8192"))

# The transcription prompt is identical to the other backends so parse_page()
# sees the same structural contract regardless of which model produced the text.
PROMPT = """Transcribe this election canvass report page exactly, preserving all structure. Output GitHub-flavored Markdown with these rules:

- Any office/contest title (e.g. "GOVERNOR", "LIEUTENANT GOVERNOR") on its own line, exactly as printed, followed by "(VOTE FOR) 1" on its own line if present.
- If the page says "(CONTINUED FROM PREVIOUS PAGE)", include that exact text on its own line.
- Render every vote table as a proper HTML <table> with <thead> and <tbody>, using one <td> per cell — one candidate's vote count per <td>, never multiple numbers combined into one cell.
- Each precinct row's first cell must be the exact precinct code and name as printed (e.g. "0001 SOME PRECINCT NAME").
- The totals row's first cell must be exactly "CANDIDATE TOTALS", followed by each candidate's total in its own <td>, in the same left-to-right order as the precinct rows' columns.
- Do not summarize, skip, or omit any office, precinct, or candidate — transcribe every one visible on the page, including any you are not fully certain about.
- Do not describe the image or add commentary. Output only the transcribed content.
"""

# Correction mode: image + the noisy OCR, asked for a corrected transcription in
# the same contract. (Earlier iterations of this prompt lived in the
# now-removed hybrid backend; the contract itself is unchanged.)
CORRECTION_PROMPT = PROMPT + """
An automated OCR pass produced the transcription below, which may contain wrong
digits, merged/split candidate columns, or missing rows. Use the image as ground
truth to correct it; keep every row and column. Noisy OCR to correct:

{markdown}
"""


def _nuextract_markdown(pdf_path, dpi, page):
    """The baseline nuextract3 markdown for this page, if cached (for correction
    mode). Mirrors base.extract_pages' cache path so it lines up page-for-page."""
    stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
    md = os.path.join(base.CACHE_DIR, f"{stem}_{dpi}", f"p{page:03d}.md")
    return open(md).read() if os.path.exists(md) else None


def make_extract_fn(model, correction=False, dpi_note=""):
    """extract_pages(pdf_path, dpi) bound to one Claude model via `llm`, own cache.

    The api_key argument other backends take is unnecessary here: `llm` reads the
    stored anthropic key itself. The repair driver's make_extract_fn passes
    correction= through; api_key is accepted-and-ignored for signature parity.
    """
    import llm
    llm_model = llm.get_model(model)
    model_slug = re.sub(r"[^A-Za-z0-9]+", "_", model)
    mode = "correction" if correction else "blind"

    def extract_pages(pdf_path, dpi):
        stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
        cache = os.path.join(f".canvass_cache_claude_{model_slug}_{mode}", f"{stem}_{dpi}")
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

            prompt = PROMPT
            if correction:
                noisy = _nuextract_markdown(pdf_path, dpi, i)
                if noisy:
                    prompt = CORRECTION_PROMPT.format(markdown=noisy)

            r = llm_model.prompt(prompt, attachments=[llm.Attachment(path=png)],
                                 temperature=0, max_tokens=MAX_TOKENS)
            txt = r.text() or ""
            open(md_path, "w").write(txt)
            print(f"    p{i:03d} extracted ({model}, {mode})", flush=True)
            yield i, txt

    return extract_pages


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="canvass PDF path(s); globs allowed")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"Claude model id (default: {DEFAULT_MODEL})")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--correction", action="store_true",
                    help="send the cached nuextract3 markdown alongside the image for correction")
    ap.add_argument("--validate-only", action="store_true")
    args = ap.parse_args()

    extract_fn = make_extract_fn(args.model, correction=args.correction)
    model_slug = re.sub(r"[^A-Za-z0-9]+", "_", args.model)
    out_dir = os.path.join("2026", f"counties_claude_{model_slug}")

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

    print(f"model: {args.model} (anthropic, {'correction' if args.correction else 'blind'})", flush=True)
    if not args.validate_only:
        print(f"output: {out_dir}/ (kept separate from 2026/counties/ on purpose)", flush=True)
    tp = tf = 0
    for county_paths in by_county.values():
        a, b = base.process(sorted(county_paths), args.dpi, county_df, args.validate_only,
                            extract_fn=extract_fn, out_dir=out_dir)
        tp += a
        tf += b
    print(f"\n==== {len(by_county)} count{'y' if len(by_county) == 1 else 'ies'}, "
          f"model={args.model}: {tp} verified, {tf} failed ====")
    return 1 if tf else 0


if __name__ == "__main__":
    sys.exit(main())
