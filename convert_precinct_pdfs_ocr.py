#!/usr/bin/env python3
"""
OCR-based parser for ES&S "PRECINCT REPORT" (EL30) layout PDFs using Ollama.

Uses NaturalPDF to render pages as images and qwen3.5:397b-cloud for OCR,
producing per-county precinct-level CSVs in 2026/counties/.

This handles PDFs that don't have usable text layers for pdftotext —
specifically Choctaw and Hale counties for the 2026 primary.

Usage:
    python convert_precinct_pdfs_ocr.py <pdf> [<pdf> ...] [--dpi 150]
    python convert_precinct_pdfs_ocr.py "2026 AL Republican.../Choctaw County.../Chocraw County Blue Sheet and Canvas Report.pdf"
"""

import argparse
import base64
import glob
import os
import re
import sys
from collections import defaultdict
from io import BytesIO

import pandas as pd
from natural_pdf import PDF
from ollama import chat

# Import parsing functions from convert_precinct_pdfs.py
import convert_precinct_pdfs as base

MODEL = os.environ.get("OLLAMA_MODEL", "qwen3.5:397b-cloud")
CACHE_DIR = ".precinct_ocr_cache"

# OCR prompt tuned for EL30 "PRECINCT REPORT" format
PROMPT = """Transcribe this election precinct report page exactly as printed. Output plain text only.

Include every line in order:
- Precinct header: 4-digit code + name (e.g., "0001 PENNINGTON")
- Statistics: "REGISTERED VOTERS", "BALLOTS CAST", etc.
- Party: "***** (ALABAMA DEMOCRATIC P) *****" or "***** (ALABAMA REPUBLICAN P) *****"
- Office: "GOVERNOR", "LIEUTENANT GOVERNOR", "U.S. SENATOR", etc.
- "(VOTE FOR) 1" lines
- Candidates: "NAME . . . votes percent"
- "Over Votes" / "Under Votes"
- Amendments: "PROPOSED STATEWIDE AMENDMENT" with "Yes" / "No"

Output ONLY the transcribed text, nothing else."""


def extract_pages(pdf_path, dpi):
    """Yield (page_number, text) for every page, caching OCR output on disk."""
    stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
    cache = os.path.join(CACHE_DIR, f"{stem}_{dpi}")
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

        # Use Ollama with qwen3.5:397b-cloud for OCR
        r = chat(
            model=MODEL,
            messages=[
                {"role": "user", "content": PROMPT, "images": [png]},
            ],
        )
        txt = r["message"]["content"]

        open(md_path, "w").write(txt)
        print(f"    p{i:03d} extracted", flush=True)
        yield i, txt


def parse_el30_text(text, county):
    """Parse EL30 precinct report OCR output into rows.

    The OCR output has a different format than pdftotext output:
    - Party markers appear as "ALABAMA DEMOCRATIC P" or "ALABAMA REPUBLICAN P"
    - Candidate rows have format: "NAME . . . votes percent"
    """
    lines = text.splitlines()
    rows = []
    cur_precinct = None
    cur_party = None
    cur_office = None
    office_buffer = []  # Buffer for multi-line office names

    def emit(name, votes):
        if cur_office and cur_party in ("DEM", "REP"):
            office, district = cur_office
            # Clean up candidate name - remove dot leaders
            name = re.sub(r"\s*\.\s*", " ", name).strip()
            name = re.sub(r"\s+", " ", name).strip()
            # Skip Over/Under votes
            if name.lower().startswith("over vote") or name.lower().startswith("under vote"):
                return
            rows.append({
                "county": county,
                "precinct": cur_precinct,
                "office": office,
                "district": district,
                "party": cur_party,
                "candidate": name,
                "votes": votes,
            })

    # Regex patterns
    PRECINCT_RE = re.compile(r"^(\d{4})\s+(.+?)\s*$")
    PARTY_DEM_RE = re.compile(r"ALABAMA\s+DEMOCRAT", re.IGNORECASE)
    PARTY_REP_RE = re.compile(r"ALABAMA\s+REPUBLICAN", re.IGNORECASE)
    VOTEFOR_RE = re.compile(r"\(VOTE FOR\)")
    # Candidate regex - captures name with dot leaders, votes, and optional percentage
    CANDIDATE_RE = re.compile(r"^(.+?)\s+(\d+)\s+([\d.]+)\s*$")
    CANDIDATE_NO_PCT = re.compile(r"^(.+?)\s+(\d+)\s*$")
    # Skip patterns
    SKIP_RE = re.compile(r"^(RUN DATE|STATISTICS|OFFICIAL REPORT|VOTES PERCENT|PRECINCT REPORT|PRIMARY ELECTION|MAY 19|2026 PRIMARY|REPORT-EL|.*PAGE \d+|Page \d+|WITH.*COUNTED|SUMMARY REPORT)$", re.IGNORECASE)

    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # Skip structural lines
        if SKIP_RE.match(stripped):
            i += 1
            continue

        # Precinct header - must start with 4 digits
        pr = PRECINCT_RE.match(stripped)
        if pr and stripped[:4].isdigit() and len(pr.group(1)) == 4:
            cur_precinct = f"{int(pr.group(1))} {pr.group(2).strip()}"
            cur_party = None
            cur_office = None
            office_buffer = []
            i += 1
            continue

        # Party marker
        if PARTY_DEM_RE.search(stripped):
            cur_party = "DEM"
            cur_office = None
            office_buffer = []
            i += 1
            continue
        if PARTY_REP_RE.search(stripped):
            cur_party = "REP"
            cur_office = None
            office_buffer = []
            i += 1
            continue

        # Check for office header followed by (VOTE FOR)
        if cur_party in ("DEM", "REP") and not VOTEFOR_RE.search(stripped):
            # Look ahead for (VOTE FOR)
            for j in range(i+1, min(i+4, n)):
                next_line = lines[j].strip()
                if VOTEFOR_RE.search(next_line):
                    # This is an office header - accumulate office lines
                    office_buffer.append(stripped)
                    cur_office = base.normalize_office(" ".join(office_buffer))
                    office_buffer = []
                    i = j + 1
                    break
            else:
                # Not an office header, try as candidate row
                if cur_office:
                    m = CANDIDATE_RE.match(stripped)
                    if m:
                        emit(m.group(1).strip(), int(m.group(2)))
                        i += 1
                        continue
                    m = CANDIDATE_NO_PCT.match(stripped)
                    if m:
                        emit(m.group(1).strip(), int(m.group(2)))
                        i += 1
                        continue
                i += 1
            continue

        i += 1

    return rows


def detect_county(text_lines, folder):
    """Detect county name from OCR'd text."""
    for line in text_lines[:15]:
        m = re.search(r"([A-Z][A-Za-z .']+?)\s+COUNTY", line)
        if m:
            return base.title_office(m.group(1).strip())
    return folder.split(" County")[0]


def process(pdf_paths, dpi, county_df, out_dir="2026/counties"):
    """Process every PDF into a single output CSV per county."""
    by_county = {}
    for pdf_path in pdf_paths:
        if not os.path.exists(pdf_path):
            print(f"skip (missing): {pdf_path}", file=sys.stderr)
            continue
        folder = os.path.basename(os.path.dirname(pdf_path))
        county = folder.split(" County")[0] if " County" in folder else "Unknown"
        by_county.setdefault(county, []).append(pdf_path)

    all_results = []

    for county, pdfs in by_county.items():
        print(f"\n=== {county}: {len(pdfs)} PDF file(s) ===", flush=True)
        rows = []

        for pdf_path in pdfs:
            for page, txt in extract_pages(pdf_path, dpi):
                # Use custom EL30 parser instead of base parser
                parsed = parse_el30_text(txt, county)
                rows.extend(parsed)
                print(f"    Page {page}: {len(parsed)} rows extracted", flush=True)

        if not rows:
            print(f"  No data extracted from {county}")
            continue

        # Clean candidate names
        for r in rows:
            r["candidate"] = base.clean_candidate_name(r["candidate"])

        # Match to county CSV for exact spelling
        office_str_by_key = defaultdict(lambda: defaultdict(dict))
        cand_str_by_key = defaultdict(dict)
        cand_any_by_key = defaultdict(dict)
        cand_to_office = defaultdict(dict)

        for _, r in county_df.iterrows():
            c, p = r["county"], r["party"]
            ok = base.office_match_key(r["office"])
            office_str_by_key[c][p][ok] = r["office"]
            ck = base.norm_key(r["candidate"])
            cand_str_by_key[c][(p, ok, r["district"], ck)] = r["candidate"]
            cand_any_by_key[c][ck] = r["candidate"]
            cand_to_office[c][(p, ck)] = (r["office"], r["district"], r["candidate"])

        for r in rows:
            ok = base.office_match_key(r["office"])
            cmap = office_str_by_key.get(county, {}).get(r["party"], {})
            if ok in cmap:
                r["office"] = cmap[ok]
            ck = base.norm_key(r["candidate"])
            key = (r["party"], ok, r["district"], ck)
            if key in cand_str_by_key.get(county, {}):
                r["candidate"] = cand_str_by_key[county][key]
            elif ck in cand_any_by_key.get(county, {}):
                r["candidate"] = cand_any_by_key[county][ck]
            elif ck in base.CANDIDATE_ALIASES:
                r["candidate"] = base.CANDIDATE_ALIASES[ck]
            else:
                r["candidate"] = base.title_office(r["candidate"])

            if ok not in cmap:
                co = cand_to_office.get(county, {}).get((r["party"], ck))
                if co:
                    r["office"], r["district"], r["candidate"] = co

        df = pd.DataFrame(rows, columns=["county", "precinct", "office", "district",
                                         "party", "candidate", "votes"])
        df["votes"] = df["votes"].astype(int)
        df = df[df["office"].notna() & df["party"].isin(["DEM", "REP"])]
        df = (df.groupby(["county", "precinct", "office", "district", "party",
                          "candidate"], as_index=False)["votes"].sum())
        df = df.sort_values(["precinct", "party", "office", "district", "candidate"]).reset_index(drop=True)

        # Canonicalize precinct names by code
        df["code"] = df.precinct.str[:4]
        canon = df.groupby("code").precinct.agg(lambda s: s.value_counts().idxmax())
        for code, variants in df.groupby("code").precinct.unique().items():
            if len(variants) > 1:
                print(f"  -> precinct {code} name varied {sorted(variants)}; using {canon[code]!r}")
        df["precinct"] = df.code.map(canon)
        df = df.drop(columns="code")

        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, f"{county.lower().replace(' ', '_').replace('.', '')}.csv")
        df.to_csv(out, index=False)
        print(f"  -> wrote {len(df)} rows to {out}")
        all_results.append((county, len(df), len(df.precinct.unique()), out))

    return all_results


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="precinct report PDF path(s)")
    ap.add_argument("--dpi", type=int, default=150, help="render resolution (default 150)")
    args = ap.parse_args()

    county_df = pd.read_csv(base.COUNTY_CSV, dtype=str, keep_default_na=False)
    county_df["votes"] = county_df["votes"].astype(int)

    paths = []
    for p in args.pdfs:
        paths.extend(sorted(glob.glob(p)) if any(c in p for c in "*?[") else [p])

    if not paths:
        print("No PDF files found", file=sys.stderr)
        return 1

    results = process(paths, args.dpi, county_df)

    print("\n=== Summary ===")
    for county, nrows, nprec, out in results:
        print(f"  {county}: {nrows} rows, {nprec} precincts -> {out}")

    return 0 if results else 1


if __name__ == "__main__":
    sys.exit(main())
