#!/usr/bin/env python3
"""
Process cached OCR output for EL30 precinct reports - v2 with column splitting.
"""

import os
import re
import pandas as pd
import convert_precinct_pdfs as base

CACHE_DIR = ".precinct_ocr_cache/Chocraw_County_Blue_Sheet_and_Canvas_Report_150"
COUNTY = "Choctaw"
county_df = pd.read_csv(base.COUNTY_CSV, dtype=str, keep_default_na=False)
county_df["votes"] = county_df["votes"].astype(int)

# Build candidate lookup
cand_by_key = {}
for _, r in county_df.iterrows():
    if r["county"] == COUNTY:
        ck = base.norm_key(r["candidate"])
        cand_by_key[(r["party"], ck)] = r["candidate"]

def parse_page(text, county):
    """Parse a single OCR page into rows."""
    lines = text.splitlines()
    rows = []
    cur_precinct = None
    cur_party = None
    cur_office = None

    HEADER_WORDS = {"PRIMARY", "ELECTION", "REPORT", "COUNTY", "PRECINCT", "RUN", "PAGE", "STATISTICS", "OFFICIAL", "VOTES", "PERCENT"}
    PARTY_DEM_RE = re.compile(r"ALABAMA\s+DEMOCRAT", re.IGNORECASE)
    PARTY_REP_RE = re.compile(r"ALABAMA\s+REPUBLICAN", re.IGNORECASE)
    VOTEFOR_RE = re.compile(r"\(VOTE FOR\)")
    PRECINCT_RE = re.compile(r"^(\d{4})\s+([A-Z][A-Z0-9 .'-]+?)\s*$")
    # Candidate: name followed by votes and percentage
    CANDIDATE_RE = re.compile(r"^(.+?)\s+(\d+)\s+([\d.]+)\s*$")

    def try_emit_candidate(line, party, precinct, office):
        """Try to parse a candidate line and emit a row."""
        m = CANDIDATE_RE.match(line)
        if not m:
            return None

        name = m.group(1).strip()
        votes = int(m.group(2))

        # Clean name - remove dot leaders
        name = re.sub(r"\s*\.\s*", " ", name).strip()
        name = re.sub(r"\s+", " ", name).strip()

        # Skip Over/Under votes
        if name.lower().startswith(("over vote", "under vote")):
            return None

        # Skip if name is too long (merged columns)
        if len(name) > 45:
            return None

        # Skip if name contains header words
        name_upper = name.upper()
        if any(word in name_upper for word in HEADER_WORDS):
            return None

        # Match to county CSV
        ck = base.norm_key(name)
        key = (party, ck)
        if key in cand_by_key:
            name = cand_by_key[key]

        return {
            "county": county,
            "precinct": precinct,
            "office": office[0] if office else None,
            "district": office[1] if office else None,
            "party": party,
            "candidate": name,
            "votes": votes,
        }

    i = 0
    n = len(lines)

    while i < n:
        line = lines[i].rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # Skip header lines
        if any(hw in stripped.upper() for hw in ["RUN DATE", "OFFICIAL REPORT", "STATISTICS", "PRECINCT REPORT", "PRIMARY ELECTION", "MAY 19", "2026 PRIMARY"]):
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

        # Precinct header
        pr = PRECINCT_RE.match(stripped)
        if pr and len(pr.group(1)) == 4:
            name_parts = pr.group(2).split()
            if not any(part in HEADER_WORDS for part in name_parts):
                cur_precinct = f"{int(pr.group(1))} {pr.group(2).strip()}"
                cur_office = None
                i += 1
                continue

        # Office header - look for (VOTE FOR) on next lines
        if cur_party in ("DEM", "REP"):
            if VOTEFOR_RE.search(stripped):
                i += 1
                continue

            # Check if next non-empty line has (VOTE FOR)
            found_office = False
            for j in range(i+1, min(i+4, n)):
                next_line = lines[j].strip()
                if VOTEFOR_RE.search(next_line):
                    # Skip if office name contains noise
                    if "votes" not in stripped.lower() and "percent" not in stripped.lower() and "under" not in stripped.lower() and "over" not in stripped.lower():
                        cur_office = base.normalize_office(stripped)
                    found_office = True
                    i = j + 1
                    break

            if not found_office and cur_office:
                # Try to parse as candidate
                row = try_emit_candidate(stripped, cur_party, cur_precinct, cur_office)
                if row:
                    rows.append(row)
                i += 1
            elif not found_office:
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
