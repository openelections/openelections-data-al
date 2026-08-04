#!/usr/bin/env python3
"""
Process cached OCR output for EL30 precinct reports.
Handles the two-column layout by processing each column separately.
"""

import os
import re
import pandas as pd
from collections import defaultdict
import convert_precinct_pdfs as base

CACHE_DIR = ".precinct_ocr_cache/Chocraw_County_Blue_Sheet_and_Canvas_Report_150"
COUNTY = "Choctaw"
county_df = pd.read_csv(base.COUNTY_CSV, dtype=str, keep_default_na=False)
county_df["votes"] = county_df["votes"].astype(int)

def parse_page(text, county):
    """Parse a single OCR page into rows."""
    lines = text.splitlines()
    rows = []
    cur_precinct = None
    cur_party = None
    cur_office = None

    # Build office/candidate lookup from county CSV
    cand_by_key = {}
    for _, r in county_df.iterrows():
        if r["county"] == county:
            ck = base.norm_key(r["candidate"])
            cand_by_key[(r["party"], ck)] = r["candidate"]

    def emit(name, votes):
        if cur_office and cur_party in ("DEM", "REP"):
            office, district = cur_office
            # Clean candidate name - remove dot leaders
            name = re.sub(r"\s*\.\s*", " ", name).strip()
            name = re.sub(r"\s+", " ", name).strip()
            # Skip Over/Under votes
            if name.lower().startswith("over vote") or name.lower().startswith("under vote"):
                return
            # Skip lines that look like merged column output
            if len(name) > 50 or name.count("  ") > 2:
                return
            # Skip if name contains "Under Votes" or "Over Votes" (merged column artifact)
            if "under vote" in name.lower() or "over vote" in name.lower():
                return
            # Match to county CSV
            ck = base.norm_key(name)
            key = (cur_party, ck)
            if key in cand_by_key:
                name = cand_by_key[key]
            rows.append({
                "county": county,
                "precinct": cur_precinct,
                "office": office,
                "district": district,
                "party": cur_party,
                "candidate": name,
                "votes": votes,
            })

    # Precinct header: 4 digits followed by precinct name (not header text)
    PRECINCT_RE = re.compile(r"^(\d{4})\s+([A-Z][A-Z0-9 .'-]+?)\s*$")
    HEADER_WORDS = {"PRIMARY", "ELECTION", "REPORT", "COUNTY", "PRECINCT", "RUN", "PAGE", "STATISTICS", "OFFICIAL"}
    PARTY_DEM_RE = re.compile(r"ALABAMA\s+DEMOCRAT", re.IGNORECASE)
    PARTY_REP_RE = re.compile(r"ALABAMA\s+REPUBLICAN", re.IGNORECASE)
    VOTEFOR_RE = re.compile(r"\(VOTE FOR\)")
    CANDIDATE_RE = re.compile(r"^(.+?)\s+(\d+)\s+([\d.]+)\s*$")
    CANDIDATE_NO_PCT = re.compile(r"^(.+?)\s+(\d+)\s*$")
    SKIP_RE = re.compile(r"^(RUN DATE|STATISTICS|OFFICIAL REPORT|VOTES PERCENT|PRECINCT REPORT|PRIMARY|MAY 19|2026|REPORT-EL|.*PAGE|Page \d|WITH.*COUNTED|SUMMARY|Precinct REPORT|.*PRIMARY ELECTION)$", re.IGNORECASE)

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()

        if not stripped or SKIP_RE.match(stripped):
            i += 1
            continue

        # Party marker
        if PARTY_DEM_RE.search(stripped):
            cur_party = "DEM"
            cur_office = None
            i += 1
            continue
        if PARTY_REP_RE.search(stripped):
            cur_party = "REP"
            cur_office = None
            i += 1
            continue

        # Precinct header - must be exactly 4 digits followed by space and name
        # Skip lines that look like headers (contain header words)
        pr = PRECINCT_RE.match(stripped)
        if pr and len(pr.group(1)) == 4:
            # Check if this looks like a header line, not a precinct
            name_parts = pr.group(2).split()
            if not any(part in HEADER_WORDS for part in name_parts):
                cur_precinct = f"{int(pr.group(1))} {pr.group(2).strip()}"
                # Don't reset cur_party - party marker comes before precinct header
                cur_office = None
                i += 1
                continue

        # Office header - look for (VOTE FOR) on next line
        if cur_party in ("DEM", "REP"):
            if VOTEFOR_RE.search(stripped):
                i += 1
                continue
            # Check if next non-empty line has (VOTE FOR)
            for j in range(i+1, min(i+4, n)):
                next_line = lines[j].strip()
                if VOTEFOR_RE.search(next_line):
                    # Skip if this looks like a merged column line (contains "Under Votes" or "Over Votes")
                    if "under vote" in stripped.lower() or "over vote" in stripped.lower():
                        cur_office = None
                    else:
                        cur_office = base.normalize_office(stripped)
                    i = j + 1
                    break
            else:
                # Try as candidate row
                if cur_office:
                    m = CANDIDATE_RE.match(stripped)
                    if m:
                        emit(m.group(1).strip(), int(m.group(2)))
                    else:
                        m = CANDIDATE_NO_PCT.match(stripped)
                        if m:
                            emit(m.group(1).strip(), int(m.group(2)))
                i += 1
            continue

        i += 1

    return rows

# Process all cached pages
all_rows = []
for fname in sorted(os.listdir(CACHE_DIR)):
    if not fname.endswith(".md"):
        continue
    fpath = os.path.join(CACHE_DIR, fname)
    text = open(fpath).read()
    rows = parse_page(text, COUNTY)
    if rows:
        print(f"{fname}: {len(rows)} rows")
        all_rows.extend(rows)

# Aggregate and write CSV
if all_rows:
    df = pd.DataFrame(all_rows, columns=["county", "precinct", "office", "district", "party", "candidate", "votes"])
    df["votes"] = df["votes"].astype(int)
    df = df.groupby(["county", "precinct", "office", "district", "party", "candidate"], as_index=False)["votes"].sum()
    df = df.sort_values(["precinct", "party", "office", "district", "candidate"]).reset_index(drop=True)

    # Canonicalize precinct names
    df["code"] = df.precinct.str[:4]
    canon = df.groupby("code").precinct.agg(lambda s: s.value_counts().idxmax())
    df["precinct"] = df.code.map(canon)
    df = df.drop(columns="code")

    out_path = "2026/counties/20260519__al__primary__choctaw__precinct.csv"
    df.to_csv(out_path, index=False)
    print(f"\nWrote {len(df)} rows to {out_path}")
    print(f"Unique precincts: {df.precinct.nunique()}")
else:
    print("No rows extracted!")
