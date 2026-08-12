#!/usr/bin/env python3
"""
Status/routing report for the 2026 AL primary precinct-CSV pipeline.

This does NOT run any parser or vision model. It answers two questions per
county, cheaply (pdftotext/pdfinfo only, no vision-model calls, no natural_pdf/
ollama dependency — runs under plain system python3):

  1. Does 2026/counties/ already have a clean output for this county?
  2. If not, which existing script should produce it, and is that cheap
     (OCR already cached in .canvass_cache/.precinct_ocr_cache — a rerun costs
     nothing) or does it need a fresh, paid vision-model pass?

Three parsers already exist and should NOT be reimplemented here:
  - convert_precinct_pdfs.py       text-layer EL30 "PRECINCT REPORT" PDFs.
                                    No vision model. Whole-corpus batch script
                                    (no CLI args) — already "applies to
                                    everything" for this format.
  - convert_precinct_pdfs_ocr.py   image-only EL30 PDFs. Vision (Ollama,
                                    default qwen3.5:397b-cloud, override via
                                    $OLLAMA_MODEL).
  - convert_canvass_pdfs.py        image-only "NAME HEADING CANVASS" matrix
                                    PDFs (vertical/rotated candidate headers).
                                    Vision (Ollama nuextract3 by default,
                                    override via $NUEXTRACT_MODEL), or a
                                    drop-in alternate backend via
                                    convert_canvass_pdfs_paddleocr.py
                                    (PaddleOCR-VL) / convert_canvass_pdfs_claude.py
                                    driven through repair_canvass_contests.py.
                                    This is the "positional-column" case, but it
                                    already does better than pure position:
                                    candidate identity comes from joining each
                                    column's checksummed total against
                                    2026/20260519__al__primary__county.csv,
                                    with ballot-order (alphabetical-by-surname)
                                    position used only to break an exact tie.

This script's job is purely to say, for every county, which of those three to
run (or that the source document itself is missing/wrong and no script can
help yet). Vision-model reruns are expensive (~40-170s/page) — nothing here
triggers one. Add --execute yourself when you're ready; it is intentionally
not wired to anything by default.

Usage:
    python3 pipeline_status.py                  # full report, all counties
    python3 pipeline_status.py --county shelby tuscaloosa
    python3 pipeline_status.py --action-only     # skip already-clean counties
"""

import argparse
import csv
import glob
import os
import re
import subprocess

REPO = os.path.dirname(os.path.abspath(__file__))
COUNTY_CSV = os.path.join(REPO, "2026", "20260519__al__primary__county.csv")
COUNTIES_DIR = os.path.join(REPO, "2026", "counties")
GOP_ROOT = os.path.join(REPO, "2026 AL Republican Party Primary Precinct Results")
CANVASS_CACHE = os.path.join(REPO, ".canvass_cache")
PRECINCT_OCR_CACHE = os.path.join(REPO, ".precinct_ocr_cache")

MIN_PRECINCT_PAGES = 5  # fewer pages than this is almost certainly a Blue
                         # Sheet / certification summary, not precinct detail


def norm(s):
    return re.sub(r"[^a-z]", "", s.lower())


# ---------------------------------------------------------------------------
# 1. What does 2026/counties/ already have?
# ---------------------------------------------------------------------------

def load_county_status():
    """{county: (state, extra)} — state in 'missing'/'non-conforming'/'has-placeholders'/'clean'"""
    # county.csv itself has an inconsistent-casing duplicate ("Dekalb" and
    # "DeKalb" both appear) — dedupe by normalized name so it isn't reported twice.
    all_counties = {}
    with open(COUNTY_CSV, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            all_counties.setdefault(norm(r["county"]), r["county"])

    have = {}  # norm(county) -> (real_name, placeholder_count)
    for path in sorted(glob.glob(os.path.join(COUNTIES_DIR, "*.csv"))):
        fname = os.path.basename(path)
        if "democratic" in fname:
            continue  # statewide extract, not a county file
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames
            rows = list(reader)
        if not rows:
            continue
        if fieldnames != ["county", "precinct", "office", "district", "party", "candidate", "votes"]:
            have[norm(rows[0]["county"])] = (rows[0]["county"], "non-conforming")
            continue
        county = rows[0]["county"]
        n_placeholder = sum(
            1 for r in rows if r.get("candidate", "").startswith("Unverified Candidate")
        )
        have[norm(county)] = (county, n_placeholder)

    status = {}
    for key, c in sorted(all_counties.items(), key=lambda kv: kv[1]):
        hit = have.get(key)
        if hit is None:
            status[c] = ("missing", None)
        elif hit[1] == "non-conforming":
            status[c] = ("non-conforming", "wrong column schema, needs regenerating")
        elif hit[1] > 0:
            status[c] = ("has-placeholders", hit[1])
        else:
            status[c] = ("clean", 0)
    return status


# ---------------------------------------------------------------------------
# 2. Cache hits (cheap — OCR already paid for)
# ---------------------------------------------------------------------------

def cache_hit(county):
    """Return 'matrix' | 'el30-image' | None based on existing cache dirs."""
    n = norm(county)
    for entry in os.listdir(CANVASS_CACHE) if os.path.isdir(CANVASS_CACHE) else []:
        if n in norm(entry):
            return "matrix"
    for entry in os.listdir(PRECINCT_OCR_CACHE) if os.path.isdir(PRECINCT_OCR_CACHE) else []:
        if n in norm(entry):
            return "el30-image"
    return None


# ---------------------------------------------------------------------------
# 3. Cheap format classification for counties with no cache yet
# ---------------------------------------------------------------------------

def find_county_folder(county):
    if not os.path.isdir(GOP_ROOT):
        return None
    for entry in os.listdir(GOP_ROOT):
        if entry.lower().startswith(county.lower() + " county"):
            return os.path.join(GOP_ROOT, entry)
    return None


def classify_pdf(path):
    """Return (format, confidence, note) for one PDF, using only pdfinfo/pdftotext."""
    try:
        info = subprocess.run(["pdfinfo", path], capture_output=True, text=True, timeout=30).stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return ("unknown", "low", "pdfinfo failed")
    m_pages = re.search(r"Pages:\s+(\d+)", info)
    m_size = re.search(r"Page size:\s+([\d.]+)\s+x\s+([\d.]+)", info)
    pages = int(m_pages.group(1)) if m_pages else 0
    landscape = False
    if m_size:
        w, h = float(m_size.group(1)), float(m_size.group(2))
        landscape = w > h

    if pages < MIN_PRECINCT_PAGES:
        return ("too-short", "high",
                f"{pages} page(s) — likely a Blue Sheet/certification summary, not precinct detail")

    try:
        txt = subprocess.run(["pdftotext", "-layout", "-l", "3", path, "-"],
                              capture_output=True, text=True, timeout=30).stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        txt = ""

    has_text = bool(txt.strip())
    if has_text and "PRECINCT REPORT" in txt.upper():
        has_precinct_hdr = bool(re.search(r"^\d{4}\s+[A-Z]", txt, re.M))
        conf = "medium" if has_precinct_hdr else "low"
        return ("el30-textlayer", conf,
                "pdftotext finds 'PRECINCT REPORT' + a precinct header — "
                "try convert_precinct_pdfs.py first, it's free; fall back to "
                "the vision path only if its own validation section flags "
                "discrepancies for this county")
    if has_text:
        return ("unknown", "low", "has a text layer but no recognizable EL30 markers")

    # image-only
    if landscape:
        return ("matrix", "medium",
                f"{pages}pp, landscape, no text layer — matches the "
                "'NAME HEADING CANVASS' matrix shape")
    return ("el30-image", "low",
            f"{pages}pp, portrait, no text layer — guessing EL30 image "
            "format by elimination; not confirmed")


def pick_pdf(folder):
    """Prefer a PDF whose name suggests precinct/canvas detail over a bare
    Blue Sheet or certification-only document; skip non-PDF folders."""
    pdfs = sorted(glob.glob(os.path.join(folder, "*.pdf")) + glob.glob(os.path.join(folder, "*.PDF")))
    if not pdfs:
        others = os.listdir(folder)
        img_only = others and all(f.lower().endswith((".jpg", ".jpeg", ".png")) for f in others)
        if img_only:
            return None, f"no PDF at all — {len(others)} raw image file(s) only; " \
                          "needs a new ingestion path (none of the existing scripts take images)"
        return None, "no PDF found in folder"
    ranked = sorted(
        pdfs,
        key=lambda p: (
            0 if re.search(r"canvas|precinct", os.path.basename(p), re.I) else
            1 if not re.search(r"cert|blue sheet", os.path.basename(p), re.I) else 2,
            p,
        ),
    )
    return ranked[0], None


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

ROUTE_COMMAND = {
    "matrix": (
        "python repair_canvass_contests.py repair <pdf> --model paddleocr --only-mismatched   "
        "# PaddleOCR-VL primary; or --model anthropic/claude-sonnet-4-6 as fallback"
    ),
    "el30-image": (
        "python convert_precinct_pdfs_ocr.py <pdf>   "
        "# Ollama qwen3.5:397b-cloud by default; export OLLAMA_MODEL=... to change it"
    ),
    "el30-textlayer": "python convert_precinct_pdfs.py   # no vision model, whole-corpus batch, free",
}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--county", nargs="*", default=None, help="limit to these county names (case-insensitive)")
    ap.add_argument("--action-only", action="store_true", help="omit counties that are already clean")
    args = ap.parse_args()

    status = load_county_status()
    counties = sorted(status)
    if args.county:
        wanted = {norm(c) for c in args.county}
        counties = [c for c in counties if norm(c) in wanted]

    groups = {"free-rerun": [], "vision-matrix": [], "vision-el30": [],
              "try-textlayer-first": [], "needs-source-doc": [], "uncertain": [], "clean": []}

    def label_for(state, extra):
        if state == "has-placeholders":
            return f"{state} ({extra} placeholder rows)"
        if state == "non-conforming":
            return f"{state} ({extra})"
        return state

    for county in counties:
        state, extra = status[county]
        if state == "clean":
            groups["clean"].append((county, "already clean, no placeholders"))
            continue
        label = label_for(state, extra)

        cache = cache_hit(county)
        if cache == "matrix":
            groups["free-rerun"].append(
                (county, f"{label} — .canvass_cache/ already has this "
                         "county's pages OCR'd; rerunning convert_canvass_pdfs.py costs nothing "
                         "(cache hit) and reapplies the current (better) naming/checksum logic"))
            continue
        if cache == "el30-image":
            groups["free-rerun"].append(
                (county, f"{label} — .precinct_ocr_cache/ already has "
                         "this county's pages OCR'd; rerunning convert_precinct_pdfs_ocr.py costs "
                         "nothing (cache hit)"))
            continue

        folder = find_county_folder(county)
        if not folder:
            groups["needs-source-doc"].append((county, f"{label} — no source folder found under {GOP_ROOT!r}"))
            continue
        pdf, note = pick_pdf(folder)
        if pdf is None:
            groups["needs-source-doc"].append((county, f"{label} — {note} (folder: {os.path.basename(folder)})"))
            continue

        fmt, conf, cnote = classify_pdf(pdf)
        rel_pdf = os.path.relpath(pdf, REPO)
        if fmt == "too-short":
            groups["needs-source-doc"].append((county, f"{label} — {rel_pdf}: {cnote}"))
        elif fmt == "matrix":
            groups["vision-matrix"].append((county, f"{label} — {rel_pdf} [{conf} confidence]: {cnote}"))
        elif fmt == "el30-image":
            groups["vision-el30"].append((county, f"{label} — {rel_pdf} [{conf} confidence]: {cnote}"))
        elif fmt == "el30-textlayer":
            groups["try-textlayer-first"].append((county, f"{label} — {rel_pdf} [{conf} confidence]: {cnote}"))
        else:
            groups["uncertain"].append((county, f"{label} — {rel_pdf}: {cnote}"))

    order = [
        ("FREE RERUN (OCR already cached — no vision-model cost)", "free-rerun", "matrix-or-el30-image"),
        ("TRY TEXT-LAYER PARSER FIRST (no vision model needed)", "try-textlayer-first", "el30-textlayer"),
        ("NEEDS VISION — matrix/canvass format", "vision-matrix", "matrix"),
        ("NEEDS VISION — EL30 image format", "vision-el30", "el30-image"),
        ("NEEDS SOURCE DOCUMENT (not a parsing problem)", "needs-source-doc", None),
        ("UNCERTAIN — needs a human look", "uncertain", None),
    ]

    for title, key, route_key in order:
        items = groups[key]
        if not items:
            continue
        print(f"\n=== {title} ({len(items)}) ===")
        for county, msg in items:
            print(f"  {county}: {msg}")
        if route_key and route_key in ROUTE_COMMAND:
            print(f"  -> {ROUTE_COMMAND[route_key]}")
        elif key == "free-rerun":
            print(f"  -> {ROUTE_COMMAND['matrix']}")
            print(f"  -> {ROUTE_COMMAND['el30-image']}")

    if not args.action_only and groups["clean"]:
        print(f"\n=== CLEAN ({len(groups['clean'])}) ===")
        for county, msg in groups["clean"]:
            print(f"  {county}: {msg}")

    total_needing_work = sum(len(groups[k]) for k in groups if k != "clean")
    print(f"\n{total_needing_work} counties need action, {len(groups['clean'])} are clean, "
          f"{len(counties)} total.")
    print("\nNothing above was executed. This script only classifies and routes.")


if __name__ == "__main__":
    main()
