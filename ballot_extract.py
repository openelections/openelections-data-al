#!/usr/bin/env python3
"""
Extract the office -> ordered-candidate-list structure from a county sample
ballot PDF, using a Claude vision model via the `llm` library.

Why vision and not pdftotext: these ballots are single-page "poster" layouts
with many offices tiled across overlapping text columns. pdftotext (even with
-layout) interleaves neighboring boxes and merges words across column
boundaries — e.g. "KEN McFEETERS" collides with an adjacent "REPUBLICAN
PRIMARY AND" to produce "KENPRIMARY McFEETERSAND". The candidate order within
an office is exactly what we need and exactly what that scrambling destroys,
so the page has to be read as an image.

Why Claude and not the local Ollama models used elsewhere in this repo:
kimi/qwen were near-perfect on the canvass *grid* pages but fail badly on the
ballot poster layout — they bleed candidates between adjacent columns and
attach them to the wrong office (measured on Lawrence: US Senator's 7
candidates collapsed to 1, the rest leaking into unrelated County Commission
races; a GOP committee race got a State Treasurer candidate). claude-haiku-4.5
reads the multi-column layout cleanly — every office, correct candidates,
correct order — so this module routes ballots there via `llm`.

What this produces, per ballot, is authority data with no vote totals:
    {office_raw: [candidate_name, ...]}   # names in printed ballot order
Ballot order matters because these reports (and match_candidates.py's Pass C)
resolve tied/uncovered columns by position, and Alabama primary ballots list
candidates alphabetically by surname — the same order the canvass columns and
the county CSV use. Ballots also cover strictly-local races (county
committee, sheriff, board of education) that never appear in
2026/20260519__al__primary__county.csv at all, which is where the leftover
"Unverified Candidate" columns come from.

IMPORTANT — the ballot has no vote totals, so its output is NOT self-verifying
the way the canvass extraction is (there, each column's total is checksummed
and joined to the county CSV). Ballot names must be cross-checked downstream —
by candidate count per office and, where possible, against the checksummed
canvass columns — before being trusted, not written blindly. A stray
transcription slip does happen (Lawrence: "McNNIS" for "McINNIS"), which the
fuzzy/total-based matching in match_candidates.py absorbs, but a wholesale
mis-read would not be caught by this module alone.

Model output is cached per (ballot file, model) so re-running to fix parsing
costs nothing. Model defaults to $BALLOT_MODEL or claude-haiku-4.5. Empty or
unparseable responses are retried.

Requires the `llm` library with llm-anthropic configured (an Anthropic API
key). Run under the repo venv: `.venv/bin/python3 ballot_extract.py ...`.

Usage:
    .venv/bin/python3 ballot_extract.py .sample_ballots/Lawrence/*.pdf
    .venv/bin/python3 ballot_extract.py --county Lawrence          # fetch + extract both parties
    .venv/bin/python3 ballot_extract.py --county Lawrence --json    # dump structured result
"""

import argparse
import json
import os
import re
import sys

MODEL = os.environ.get("BALLOT_MODEL", "claude-haiku-4.5")
CACHE_DIR = os.environ.get("BALLOT_EXTRACT_CACHE", ".ballot_extract_cache")
RENDER_DPI = int(os.environ.get("BALLOT_DPI", "200"))
MAX_ATTEMPTS = 3
EMPTY_THRESHOLD = 20

PROMPT = """This image is an official primary election SAMPLE BALLOT for one county. It lists contests (offices), each followed by that office's candidates in printed order. The ballot has a multi-column layout; read each office's candidates only from directly beneath that office's header, and do not let candidates from an adjacent column bleed into a different office.

Extract every contest and its candidates. Output ONLY a JSON object mapping each office title to an ordered array of candidate names, like:

{
  "FOR GOVERNOR": ["KEN McFEETERS", "\\"Alabama\\" WILL SANTIVASCI", "THOMAS TUBERVILLE"],
  "FOR LAWRENCE COUNTY SHERIFF": ["KEITH LIGON", "STACY ROSE"]
}

Rules:
- Use the office title exactly as printed (keep "FOR", district numbers, place numbers, etc.).
- List candidates top-to-bottom in the exact order printed under each office.
- Transcribe candidate names exactly, including quotes, nicknames, initials, and suffixes (Jr., II, etc.).
- Include EVERY office on the ballot, including county commission districts, board of education, party executive committee, and coroner races.
- Do NOT include the "(Vote for One)" line, instructions, or the constitutional amendment / "Yes"/"No" ballot questions.
- Output only the JSON object. No commentary, no markdown fences.
"""


def _render_page(pdf_path, dpi, png_path):
    from natural_pdf import PDF
    pdf = PDF(pdf_path)
    pdf.pages[0].render(resolution=dpi).save(png_path)


def _strip_json(txt):
    """Pull the JSON object out of a model response that may wrap it in prose
    or a ```json fence."""
    txt = txt.strip()
    # drop a leading/trailing markdown fence if present
    txt = re.sub(r"^```(?:json)?\s*", "", txt)
    txt = re.sub(r"\s*```$", "", txt)
    # fall back to the outermost {...} span
    start, end = txt.find("{"), txt.rfind("}")
    if start != -1 and end != -1 and end > start:
        txt = txt[start:end + 1]
    return txt


def extract_ballot(pdf_path, dpi=RENDER_DPI, model=MODEL):
    """Return {office_raw: [candidate, ...]} for one ballot PDF, cached on disk."""
    import llm

    model_slug = re.sub(r"[^A-Za-z0-9]+", "_", model)
    stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
    cache = os.path.join(CACHE_DIR, model_slug)
    os.makedirs(cache, exist_ok=True)
    json_path = os.path.join(cache, f"{stem}.json")
    raw_path = os.path.join(cache, f"{stem}.raw.md")

    if os.path.exists(json_path):
        with open(json_path, encoding="utf-8") as f:
            return json.load(f)

    png = os.path.join(cache, f"{stem}.png")
    if not os.path.exists(png):
        _render_page(pdf_path, dpi, png)

    llm_model = llm.get_model(model)
    parsed, raw = None, ""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        resp = llm_model.prompt(
            PROMPT,
            attachments=[llm.Attachment(path=png)],
            temperature=0.1,
        )
        raw = resp.text() or ""
        if len(raw.strip()) >= EMPTY_THRESHOLD:
            try:
                parsed = json.loads(_strip_json(raw))
                break
            except json.JSONDecodeError:
                if attempt < MAX_ATTEMPTS:
                    print(f"    {stem}: unparseable JSON (attempt {attempt}/{MAX_ATTEMPTS}), retrying...",
                          flush=True)
                continue
        if attempt < MAX_ATTEMPTS:
            print(f"    {stem}: empty response (attempt {attempt}/{MAX_ATTEMPTS}), retrying...", flush=True)

    open(raw_path, "w", encoding="utf-8").write(raw)
    if parsed is None:
        raise RuntimeError(f"could not extract valid JSON from {pdf_path} after {MAX_ATTEMPTS} attempts "
                           f"(raw response saved to {raw_path})")

    # normalize: strip whitespace, drop empty offices
    cleaned = {}
    for office, cands in parsed.items():
        office = office.strip()
        if not isinstance(cands, list):
            continue
        names = [c.strip() for c in cands if isinstance(c, str) and c.strip()]
        if office and names:
            cleaned[office] = names

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, indent=2, ensure_ascii=False)
    print(f"    {stem}: extracted {len(cleaned)} offices", flush=True)
    return cleaned


def extract_county(county, dpi=RENDER_DPI, model=MODEL):
    """Fetch (if needed) and extract both parties for a county.
    Returns {'DEM': {...}, 'REP': {...}}."""
    from fetch_sample_ballots import fetch_ballots
    dem_path, rep_path = fetch_ballots(county)
    return {
        "DEM": extract_ballot(dem_path, dpi=dpi, model=model),
        "REP": extract_ballot(rep_path, dpi=dpi, model=model),
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="*", help="ballot PDF path(s)")
    ap.add_argument("--county", help="fetch + extract both parties for this county")
    ap.add_argument("--model", default=MODEL)
    ap.add_argument("--dpi", type=int, default=RENDER_DPI)
    ap.add_argument("--json", action="store_true", help="print the structured result")
    args = ap.parse_args()

    if args.county:
        result = extract_county(args.county, dpi=args.dpi, model=args.model)
        if args.json:
            print(json.dumps(result, indent=2, ensure_ascii=False))
        else:
            for party, offices in result.items():
                print(f"\n=== {args.county} {party}: {len(offices)} offices ===")
                for office, cands in offices.items():
                    print(f"  {office}: {cands}")
        return 0

    if not args.pdfs:
        ap.print_help()
        return 1

    for pdf in args.pdfs:
        offices = extract_ballot(pdf, dpi=args.dpi, model=args.model)
        if args.json:
            print(json.dumps(offices, indent=2, ensure_ascii=False))
        else:
            print(f"\n=== {os.path.basename(pdf)}: {len(offices)} offices ===")
            for office, cands in offices.items():
                print(f"  {office}: {cands}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
