#!/usr/bin/env python3
"""
Text-layer shortcut (plan step 2): parse the GOP precinct-result PDFs that have a
usable text layer with `pdftotext -layout`, producing per-county precinct-level
CSVs in 2026/counties/.

Only the ES&S "PRECINCT REPORT" (EL30) layout is parsed here — it's regular:
each precinct is a header line `NNNN <NAME>` at column 0, followed by a
STATISTICS block, then `********** (ALABAMA DEMOCRATIC|REPUBLICAN P) **********`
sections, each with office headers (`<OFFICE>` then `(VOTE FOR) 1`) and candidate
rows (`<name> . . . <votes> <percent>`). The landscape "NAME HEADING CANVASS"
matrices are NOT parsed by this shortcut (deferred to the LLM step per the plan).

Offices/candidates are normalized to match 2026/20260519__al__primary__county.csv
so aggregated precinct totals can be validated against the county CSV.
"""

import os
import re
import subprocess
import sys
from collections import defaultdict

import pandas as pd

GOP_ROOT = "2026 AL Republican Party Primary Precinct Results"
COUNTY_CSV = "2026/20260519__al__primary__county.csv"
OUT_DIR = "2026/counties"

# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------

# Uppercase-keyed map for statewide offices (PDF office text is uppercase).
STATEWIDE_OFFICE_MAP = {
    "GOVERNOR": "Governor",
    "LIEUTENANT GOVERNOR": "Lieutenant Governor",
    "UNITED STATES SENATOR": "U.S. Senate",
    "ATTORNEY GENERAL": "Attorney General",
    "SECRETARY OF STATE": "Secretary of State",
    "STATE TREASURER": "State Treasurer",
    "STATE AUDITOR": "State Auditor",
    "COMMISSIONER OF AGRICULTURE AND INDUSTRIES": "Commissioner of Agriculture and Industries",
}

US_HOUSE_RE = re.compile(r"(\d+)(?:st|nd|rd|th)?\s+CONGRESSIONAL DISTRICT", re.IGNORECASE)
DISTRICT_RE = re.compile(r"Distri?ct(?:\s+No\.?)?\s*(\d+)", re.IGNORECASE)
LOWER_WORDS = {"of", "and", "the", "a", "an", "for"}


def title_office(s):
    """Title-case an office name, keeping small connector words lowercase."""
    words = s.title().split()
    out = []
    for i, w in enumerate(words):
        bare = w.rstrip(".,")
        if i > 0 and bare.lower() in LOWER_WORDS:
            # lowercase the connector but preserve a trailing punctuation mark
            w = w.lower()
        out.append(w)
    return " ".join(out)


def normalize_office(raw):
    """Map a raw PDF office header to (office, district) matching the county CSV."""
    name = " ".join(raw.split())
    # Some counties' canvass software abbreviates office titles more than
    # others (Limestone: "US SENATOR", "AG AND INDUSTRIES", "DIST NO 8" in
    # place of the fully-spelled-out forms every other county in this batch
    # used) — expand the known abbreviations before any pattern match below,
    # so these still resolve to the same office as the spelled-out version.
    # Applied to `name` (not just its upper-cased form) because DISTRICT_RE
    # is matched against `name` further down — "DIST" alone doesn't match
    # that pattern (it requires "Dist(r)ict"), so leaving name unexpanded
    # would keep the district number unresolved even once the office itself
    # matches.
    for pat, repl in [(r"\bUS\b", "United States"), (r"\bAG AND INDUSTRIES\b", "Agriculture and Industries"),
                       (r"\bDIST\b", "District")]:
        name = re.sub(pat, repl, name, flags=re.IGNORECASE)
    up = name.upper()

    if up in STATEWIDE_OFFICE_MAP:
        return STATEWIDE_OFFICE_MAP[up], ""

    m = US_HOUSE_RE.search(up)
    if m:
        return "U.S. House", m.group(1)

    if up.startswith("STATE SENATOR") or up.startswith("STATE SENATE"):
        dm = DISTRICT_RE.search(name)
        return "State Senate", (dm.group(1) if dm else "")

    if up.startswith("STATE REPRESENTATIVE"):
        dm = DISTRICT_RE.search(name)
        return "State House", (dm.group(1) if dm else "")

    if "STATE BOARD OF EDUCATION" in up:
        dm = DISTRICT_RE.search(name)
        return "State Board of Education", (dm.group(1) if dm else "")

    # State Democratic Executive Committee: the PDF splits the header across two
    # lines and carries a gender tag + district (e.g. "STATE DEMOCRATIC EXECUTIVE
    # COMMITTEE (FEMALE), DISTRICT NO. 29"). The county CSV collapses all SDEC
    # races to one office with a blank district, so do the same here.
    if "STATE DEMOCRATIC EXECUTIVE COMMITTEE" in up:
        return "State Democratic Executive Committee", ""

    # County-local / other: title-case, blank district.
    return title_office(name), ""


def norm_key(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def office_match_key(s):
    """Tolerant office key for matching precinct offices to county-CSV offices.

    The PDF and the XLSX workbook spell a few offices differently:
      * "Public Service Commission" (PDF/REP) vs "Public Service Commissioner" (DEM XLSX)
      * "Place No. 1" (PDF/REP) vs "Place 1" (DEM XLSX)
    Local county-commission/school-board races add a third inconsistency: some
    sources prefix "Member, " and spell "District No. N", others give the bare
    office with just "District N" — e.g. "Member, Marion County Commission,
    District No. 1" (canvass PDF) vs "Marion County Commission, District 1"
    (county CSV), same race. Collapse all three so the same race keys equal
    across sources.
    """
    k = norm_key(s)
    k = k.replace("COMMISSIONER", "COMMISSION")
    k = k.replace("PLACENO", "PLACE")
    if k.startswith("MEMBER"):
        k = k[len("MEMBER"):]
    k = re.sub(r"DISTRICTNO(\d+)", r"DISTRICT\1", k)
    return k


# Candidate names that the precinct PDF and the county XLSX spell differently
# but that refer to the same person (verified: vote totals are equal).
CANDIDATE_ALIASES = {
    "COREYHILL": "Cory Hill",  # Ag Commissioner — PDF "Corey Hill", XLSX "Cory Hill"
}


# Candidate rows: <name> ... <votes> [<percent>]. Percent may be absent (0-vote).
CAND_VOTES_PCT = re.compile(r"^(.*?)\s+(\d+)\s+(\d*\.\d+)\s*$")
CAND_VOTES_ONLY = re.compile(r"^(.*?)\s+(\d+)\s*$")

# Structural lines to ignore when they appear mid-block.
STRUCTURAL_RE = re.compile(
    r"^(RUN DATE|RUN OATE|STATISTICS|OFFICIAL REPORT|VOTES PERCENT|"
    r"PRECINCT REPORT|PRIMARY ELECTION|MAY 19|2026 PRIMARY|REPORT-EL|"
    r".*PAGE \d+|Page \d+ of \d+|Instrument|Election Book|"
    r".*Judge of Prob|.*Total Fees|.*Total Due|.*Total Ta)",
    re.IGNORECASE,
)

PARTY_MARKER_RE = re.compile(r"\*{5,}\s*\(([^)]+)\)")
PRECINCT_HEADER_RE = re.compile(r"^(\d{4})\s+(.+?)\s*$")
VOTEFOR_RE = re.compile(r"\(VOTE FOR\)")


def detect_county(text_lines, folder):
    m = re.search(r"PRECINCT REPORT\s+(.+?)\s+COUNTY\s+OFFICIAL REPORT",
                  "\n".join(text_lines[:10]))
    if m:
        return title_office(m.group(1).strip())
    return folder.split(" County")[0]


def parse_pdf_text(text, county):
    """Parse pdftotext -layout output into precinct rows."""
    lines = text.splitlines()
    rows = []
    cur_precinct = None
    cur_party = None      # "DEM" | "REP" | "NON" | None
    cur_office = None     # (office, district) or None

    def emit(name, votes):
        office, district = cur_office
        rows.append({
            "county": county,
            "precinct": cur_precinct,
            "office": office,
            "district": district,
            "party": cur_party,
            "candidate": name,
            "votes": votes,
        })

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()

        # Precinct header: NNNN <NAME> at column 0.
        if PRECINCT_HEADER_RE.match(line) and line[:1].isdigit():
            m = PRECINCT_HEADER_RE.match(line)
            cur_precinct = f"{int(m.group(1))} {m.group(2).strip()}"
            cur_party = None
            cur_office = None
            i += 1
            continue

        # Party section marker.
        pm = PARTY_MARKER_RE.search(stripped)
        if pm:
            tag = pm.group(1).upper()
            if "DEMOCRAT" in tag:
                cur_party = "DEM"
            elif "REPUBLICAN" in tag:
                cur_party = "REP"
            else:  # NONPARTISAN, etc.
                cur_party = "NON"
            cur_office = None
            i += 1
            continue

        if not stripped:
            i += 1
            continue

        # Office header: this line is an office name iff the next non-empty line
        # is "(VOTE FOR) ...", or the line itself contains "(VOTE FOR)". Office
        # headers can span multiple lines (e.g. SDEC: "STATE DEMOCRATIC EXECUTIVE
        # COMMITTEE (FEMALE)," then "DISTRICT NO. 29"), so gather the consecutive
        # non-empty lines above this one (back to a blank line) as the office name.
        if VOTEFOR_RE.search(stripped):
            i += 1
            continue
        j = i + 1
        while j < n and not lines[j].strip():
            j += 1
        if j < n and VOTEFOR_RE.search(lines[j].strip()):
            office_lines = [stripped]
            k = i - 1
            while k >= 0 and lines[k].strip() \
                    and not PARTY_MARKER_RE.search(lines[k].strip()) \
                    and not (PRECINCT_HEADER_RE.match(lines[k]) and lines[k][:1].isdigit()) \
                    and not STRUCTURAL_RE.match(lines[k].strip()):
                office_lines.insert(0, lines[k].strip())
                k -= 1
            office_name = " ".join(office_lines)
            cur_office = normalize_office(office_name)
            i = j + 1
            continue

        # Candidate row (only inside a DEM/REP office block).
        if cur_office and cur_party in ("DEM", "REP") and not STRUCTURAL_RE.match(stripped):
            m = CAND_VOTES_PCT.match(line)
            if m:
                name = m.group(1).strip()
                votes = int(m.group(2))
                emit(name, votes)
            else:
                m = CAND_VOTES_ONLY.match(line)
                if m:
                    name = m.group(1).strip()
                    votes = int(m.group(2))
                    emit(name, votes)
        i += 1

    return rows


def clean_candidate_name(name):
    """Strip dot-leader noise from a candidate name."""
    # Collapse runs of dots and surrounding spaces.
    name = re.sub(r"\s*\.\s*", " ", name).strip()
    name = re.sub(r"\s{2,}", " ", name).strip()
    return name


def main():
    county_df = pd.read_csv(COUNTY_CSV, dtype=str, keep_default_na=False)
    # Per-county, party-aware lookups keyed by tolerant office/candidate keys.
    #   office_str_by_key[county][party][office_match_key] = office_str
    #   cand_str_by_key[county][(party, office_match_key, district, cand_key)] = cand_str
    #   cand_any_by_key[county][cand_key] = cand_str   (looser fallback)
    office_str_by_key = defaultdict(lambda: defaultdict(dict))
    cand_str_by_key = defaultdict(dict)
    cand_any_by_key = defaultdict(dict)
    # cand_to_office[county][(party, cand_key)] = (office_str, district, cand_str)
    # lets a precinct row borrow the county CSV's office label when the office
    # strings differ between sources but the candidate (and votes) match — e.g.
    # Calhoun's SREC race, labeled "Calhoun County" in the PDF but "Blount County,
    # Place No. 5" in the XLSX.
    cand_to_office = defaultdict(dict)
    for _, r in county_df.iterrows():
        c, p = r["county"], r["party"]
        ok = office_match_key(r["office"])
        office_str_by_key[c][p][ok] = r["office"]
        ck = norm_key(r["candidate"])
        cand_str_by_key[c][(p, ok, r["district"], ck)] = r["candidate"]
        cand_any_by_key[c][ck] = r["candidate"]
        cand_to_office[c][(p, ck)] = (r["office"], r["district"], r["candidate"])

    os.makedirs(OUT_DIR, exist_ok=True)

    pdfs = []
    for root, _dirs, files in os.walk(GOP_ROOT):
        for fn in files:
            if fn.lower().endswith(".pdf"):
                pdfs.append(os.path.join(root, fn))
    pdfs.sort()

    report = []
    for pdf in pdfs:
        folder = os.path.basename(os.path.dirname(pdf))
        try:
            txt = subprocess.run(["pdftotext", "-layout", pdf, "-"],
                                 capture_output=True, text=True, check=True).stdout
        except subprocess.CalledProcessError:
            continue
        if not txt.strip():
            continue
        text_lines = txt.splitlines()
        # Only parse files that actually contain EL30 precinct reports.
        if not any("PRECINCT REPORT" in l for l in text_lines):
            continue
        n_precincts = sum(1 for l in text_lines if PRECINCT_HEADER_RE.match(l) and l[:1].isdigit())
        if n_precincts == 0:
            continue

        county = detect_county(text_lines, folder)
        rows = parse_pdf_text(txt, county)
        if not rows:
            continue

        # Clean candidate names + match to county CSV exact spelling.
        for r in rows:
            r["candidate"] = clean_candidate_name(r["candidate"])
        for r in rows:
            ok = office_match_key(r["office"])
            # Align office string to county CSV exact spelling (party-aware).
            cmap = office_str_by_key.get(county, {}).get(r["party"], {})
            if ok in cmap:
                r["office"] = cmap[ok]
            ck = norm_key(r["candidate"])
            key = (r["party"], ok, r["district"], ck)
            if key in cand_str_by_key.get(county, {}):
                r["candidate"] = cand_str_by_key[county][key]
            elif ck in cand_any_by_key.get(county, {}):
                r["candidate"] = cand_any_by_key[county][ck]
            elif ck in CANDIDATE_ALIASES:
                r["candidate"] = CANDIDATE_ALIASES[ck]
            else:
                r["candidate"] = title_office(r["candidate"])

            # Office label differs between sources but candidate matches the
            # county CSV (same votes) — borrow the county CSV's office + district
            # so the precinct file stays consistent with the county file.
            if ok not in cmap:
                co = cand_to_office.get(county, {}).get((r["party"], ck))
                if co:
                    r["office"], r["district"], r["candidate"] = co

        df = pd.DataFrame(rows, columns=["county", "precinct", "office", "district",
                                         "party", "candidate", "votes"])
        df["votes"] = df["votes"].astype(int)
        # Drop rows we couldn't attach to an office/party (shouldn't happen).
        df = df[df["office"].notna() & df["party"].isin(["DEM", "REP"])]
        # Several SDEC races (gender × district) collapse to one office, so their
        # Over/Under Votes rows collide; aggregate by the full row key (sum) to
        # dedup and produce the collapsed contest's total over/under votes.
        df = (df.groupby(["county", "precinct", "office", "district", "party",
                          "candidate"], as_index=False)["votes"].sum())
        df = df.sort_values(["precinct", "party", "office", "district", "candidate"]).reset_index(drop=True)

        out_path = os.path.join(OUT_DIR, f"{county.lower().replace(' ', '_').replace('.', '')}.csv")
        df.to_csv(out_path, index=False)
        report.append((county, pdf, len(df), n_precincts, out_path))

    print("Parsed precinct CSVs:")
    for county, pdf, nrows, nprec, out in report:
        print(f"  {county}: {nrows} rows, {nprec} precincts -> {out}")
    if not report:
        print("  (no EL30 precinct-report PDFs with a text layer were found)")

    # ---- Validation: aggregate precinct totals vs county CSV ----
    print("\nValidation against", COUNTY_CSV)
    pseudocands = {"Over Votes", "Under Votes", "Over votes", "Under votes"}
    any_disc = False
    for county, pdf, nrows, nprec, out in report:
        pdf = out
        pc = pd.read_csv(pdf, dtype=str, keep_default_na=False)
        pc["votes"] = pc["votes"].astype(int)
        # aggregate real candidates only (exclude over/under votes)
        agg = (pc[~pc.candidate.isin(pseudocands)]
               .groupby(["office", "district", "party", "candidate"])["votes"].sum())
        cc = county_df[county_df.county == county]
        cc = cc.groupby(["office", "district", "party", "candidate"])["votes"].sum().astype(int)
        # compare on normalized keys
        def k(idx):
            office, district, party, candidate = idx
            return (office_match_key(office), district, party, norm_key(candidate))
        agg_map = {k(idx): v for idx, v in agg.items()}
        cc_map = {k(idx): v for idx, v in cc.items()}
        keys = set(agg_map) | set(cc_map)
        disc = []
        for key in sorted(keys):
            a = agg_map.get(key, 0)
            c = cc_map.get(key, 0)
            if a != c:
                disc.append((key, a, c))
        status = "OK" if not disc else f"{len(disc)} discrepancies"
        print(f"  {county}: {status}  (precinct rows={len(pc)}, county rows={len(cc)})")
        if disc:
            any_disc = True
            for key, a, c in disc[:20]:
                print(f"      precinct={a}  county={c}  diff={a-c}  {key}")
            if len(disc) > 20:
                print(f"      ... and {len(disc)-20} more")
    if not any_disc:
        print("  All parsed counties: precinct totals match county CSV exactly.")


if __name__ == "__main__":
    main()