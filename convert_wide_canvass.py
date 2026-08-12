#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Parser for the "wide precinct-report" canvass format (EL40-style), the third
Alabama layout — distinct from the NAME HEADING one-office-per-page matrix that
convert_canvass_pdfs.parse_page handles.

In this format a single wide table carries MANY offices at once: the header row
groups columns by office (`<td colspan=N>GOVERNOR</td>...`), then a block of
vertical name-lattice rows, then one row per precinct (precinct name, then a
vote count per candidate column spanning every office), and finally a `TOTALS`
row (NOT the "CANDIDATE TOTALS" label the other format uses). Cherokee, Bibb,
Clay, Fayette, Mobile and the combined Blue-Sheet+Canvas PDFs use this.

The OCR of this layout is structurally noisy in two distinct ways, and the
parser is built around defeating both with the printed TOTALS row as an anchor:

1. **Colspans lie.** (Cherokee p001 claims Governor spans 7 columns; it has 3
   candidates.) So offices are located, not trusted: for each header office we
   know its authoritative candidate-total multiset from 2026/..__county.csv, and
   we find the contiguous slice of the TOTALS row equal to that multiset —
   first at the running cursor, else by scanning (resync). A local race with no
   authority entry no longer loses every office after it; the next known office
   re-anchors by scan. An office whose totals appear nowhere (or ambiguously) is
   skipped and reported, never guessed.

2. **Precinct rows drift.** Individual rows gain/lose a cell mid-row, so a
   fixed slice misaligns later offices row by row. Per office, each row is
   allowed a small horizontal shift, and coordinate-descent picks the shift
   vector that makes every column sum exactly to the printed TOTALS slice.
   Exact equality or the contest is left failing — the same
   checksum-or-nothing rule as the rest of the pipeline.

Candidate names then come from the same totals-join as everywhere else
(name_columns), never from the vertical lattice.
"""

import re

import pandas as pd

from convert_precinct_pdfs import normalize_office, office_match_key


# ---------------------------------------------------------------------------
# Low-level table extraction
# ---------------------------------------------------------------------------

def _table_rows(md):
    """Every <tr> as a list of plain-text cell strings."""
    rows = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", md, re.S):
        cells = [re.sub(r"<[^>]*>", " ", c).replace("\n", " ").strip()
                 for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S)]
        rows.append(cells)
    return rows


def _office_sequence(md):
    """Offices in left-to-right header order, from the colspan-tagged header
    cell labels. The colspan NUMBERS are ignored on purpose (they lie)."""
    m = re.search(r"<tr[^>]*>(.*?)</tr>", md, re.S)
    if not m:
        return []
    offices = []
    for cell in re.findall(r"<t[dh][^>]*?colspan=\"?\d+\"?[^>]*>(.*?)</t[dh]>", m.group(1), re.S):
        name = re.sub(r"<[^>]*>", " ", cell).replace("\n", " ")
        # PaddleOCR sometimes emits a LITERAL backslash-n inside header text
        # ("UNITED STATES REPRESENTATIVE \n 3RD CONGRESS") — not a newline.
        name = name.replace("\\n", " ")
        name = re.sub(r"\s+", " ", name).strip()
        # OCR truncates "3RD CONGRESSIONAL DISTRICT" to "3RD CONGRESS"; expand it
        # here (wide format only) so normalize_office's US_HOUSE_RE can match.
        name = re.sub(r"(\d+(?:ST|ND|RD|TH))\s+CONGRESS(?:IONAL)?(?:\s+DIST(?:RICT)?)?\.?$",
                      r"\1 CONGRESSIONAL DISTRICT", name, flags=re.I)
        if name and re.search(r"[A-Z]{3,}", name):
            offices.append(name)
    return offices


_NUMERIC = re.compile(r"^-?\d+$")


def _precinct_and_totals(rows):
    """(precinct_rows, totals_values). A precinct row's first cell is a real
    name (not the one-letter-per-cell lattice) and carries >=3 integers."""
    prec, totals = [], None
    for r in rows:
        if not r:
            continue
        head = r[0].strip()
        vals = [int(c.replace(",", "")) for c in r[1:] if _NUMERIC.match(c.replace(",", ""))]
        if head.upper() == "TOTALS":
            totals = vals
        elif re.search(r"[A-Za-z]{2,}", head) and len(vals) >= 3:
            if not re.fullmatch(r"[A-Z]( [A-Z])*", head):
                prec.append((head, vals))
    return prec, totals


# ---------------------------------------------------------------------------
# Authority anchors
# ---------------------------------------------------------------------------

def _authority_totals(county, office, district, party, county_df):
    """{candidate: votes} for one contest from the authoritative county file."""
    key = office_match_key(office)
    pool = county_df[(county_df.county == county) & (county_df.party == party)
                     & (county_df.office.map(office_match_key) == key)]
    if district:
        pool = pool[pool.district.astype(str) == str(district)]
    return {r.candidate: int(r.votes) for _, r in pool.iterrows()}


# ---------------------------------------------------------------------------
# Per-row shift alignment
# ---------------------------------------------------------------------------

def _align_rows(prec, start, width, target, max_shift=2, max_iter=60):
    """Choose a small horizontal shift per precinct row so that each of the
    `width` columns sums exactly to `target`. Coordinate descent from the
    all-zero shift vector; returns (residual_error, aligned_rows, shifts).

    aligned_rows: [(name, [v0..v_{width-1}])] with None for cells that fall
    outside a short row. Exact success is residual_error == 0.
    """
    options = []
    for name, vals in prec:
        opts = {}
        for s in range(-max_shift, max_shift + 1):
            lo = start + s
            if lo < 0:
                continue
            ext = vals[lo:lo + width]
            if len(ext) == width:
                opts[s] = ext
        if not opts:
            ext = vals[start:start + width]
            opts = {0: ext + [None] * (width - len(ext))}
        options.append((name, opts))

    shifts = [0 if 0 in o[1] else sorted(o[1])[0] for o in options]

    def err(sh):
        sums = [0] * width
        for (name, opts), s in zip(options, sh):
            for j, v in enumerate(opts[s]):
                if v is not None:
                    sums[j] += v
        return sum(abs(sums[j] - target[j]) for j in range(width) if target[j] is not None)

    E = err(shifts)
    for _ in range(max_iter):
        if E == 0:
            break
        improved = False
        for i, (name, opts) in enumerate(options):
            if len(opts) == 1:
                continue
            best_s, best_E = shifts[i], E
            for s in opts:
                if s == shifts[i]:
                    continue
                old = shifts[i]
                shifts[i] = s
                e = err(shifts)
                shifts[i] = old
                if e < best_E:
                    best_s, best_E = s, e
            if best_s != shifts[i]:
                shifts[i] = best_s
                E = best_E
                improved = True
        if not improved:
            break

    aligned = [(name, opts[s]) for (name, opts), s in zip(options, shifts)]
    return E, aligned, shifts


# ---------------------------------------------------------------------------
# Page parse
# ---------------------------------------------------------------------------

def parse_wide_page(md, county, party, county_df):
    """Parse one wide-format page into contest dicts compatible with
    convert_canvass_pdfs.checksum() / name_columns().

    Every emitted contest's TOTALS slice equals its authoritative candidate-
    total multiset by construction (that's how it was located), so a downstream
    checksum PASS means the precinct rows are internally consistent with the
    document AND the document agrees with the certified county totals.
    """
    rows = _table_rows(md)
    offices = _office_sequence(md)
    prec, totals = _precinct_and_totals(rows)
    notes = []
    if not offices or not prec or not totals:
        return [], notes

    contests = []
    cursor, lost = 0, False
    for off in offices:
        office_n, district = normalize_office(off)
        auth = _authority_totals(county, office_n, district, party, county_df)
        n = len(auth)
        if n == 0:
            notes.append(f"{off}: no authority entry ({party}) — position lost, will resync")
            lost = True
            continue
        M = sorted(auth.values())
        start = None
        if not lost and cursor + n <= len(totals) and sorted(totals[cursor:cursor + n]) == M:
            start = cursor
        else:
            hits = [q for q in range(len(totals) - n + 1) if sorted(totals[q:q + n]) == M]
            if len(hits) == 1:
                start = hits[0]
                if lost or start != cursor:
                    notes.append(f"{off}: resynced to column {start}")
            elif len(hits) > 1:
                notes.append(f"{off}: ambiguous totals position {hits} — skipped")
                lost = True
                continue
            else:
                notes.append(f"{off}: authoritative totals not found in TOTALS row — skipped")
                lost = True
                continue
        cursor, lost = start + n, False

        target = totals[start:start + n]
        E, aligned, shifts = _align_rows(prec, start, n, target)
        nshift = sum(1 for s in shifts if s)
        if E == 0 and nshift:
            notes.append(f"{off}: aligned {nshift} row(s) by per-row shift")
        if E > 0:
            notes.append(f"{off}: residual column error {E} after alignment (will FAIL checksum)")
        contests.append({"office": off, "party": party, "district": district,
                         "prec": [(name, "", vals) for name, vals in aligned],
                         "totals": list(target), "pages": []})
    return contests, notes


def _denormalize(office, district):
    """Authority offices are already normalized, but downstream naming re-runs
    normalize_office (which reads the district back OUT of the office string).
    Re-encode districted offices so that round-trips, keeping the district."""
    if not district:
        return office
    return {
        "U.S. House": f"UNITED STATES REPRESENTATIVE, {district} CONGRESSIONAL DISTRICT",
        "State House": f"STATE REPRESENTATIVE, DISTRICT NO. {district}",
        "State Senate": f"STATE SENATOR, DISTRICT NO. {district}",
        "State Board of Education": f"STATE BOARD OF EDUCATION, DISTRICT NO. {district}",
    }.get(office, office)


def scan_wide_contests(md, county, party, county_df, offices_hint=None):
    """Locate contests on a wide DATA page with no reliable office headers
    (Mobile splits headers onto the odd page and data+TOTALS onto the even one,
    and the pairing is unreliable). Instead of trusting headers, scan the printed
    TOTALS row for every authoritative contest of this party whose candidate
    total multiset appears as a contiguous, unique slice — then align precinct
    rows to it. The multiset-plus-exact-alignment gate makes a coincidental match
    astronomically unlikely for contests of >=3 candidates; 2-candidate contests
    are only accepted when their slice position is unique on the page.

    Returns (contests, notes) in the same shape as parse_wide_page.
    """
    rows = _table_rows(md)
    prec, totals = _precinct_and_totals(rows)
    if not prec or not totals:
        return [], []

    # candidate authoritative contests for this party (office, district) -> totals
    pool = county_df[county_df.party == party]
    contests_meta = {}
    for (office, district), grp in pool.groupby([pool.office, pool.district]):
        if county not in set(grp.county):
            continue
        vals = sorted(grp[grp.county == county].votes.astype(int).tolist())
        if len(vals) >= 2:
            contests_meta[(office, str(district or ""))] = vals

    used = [False] * len(totals)
    contests, notes = [], []
    # larger contests first — their long multisets pin position unambiguously and
    # claim their columns before short 2-candidate races can grab them by accident
    for (office, district), M in sorted(contests_meta.items(), key=lambda kv: -len(kv[1])):
        n = len(M)
        hits = [q for q in range(len(totals) - n + 1)
                if sorted(totals[q:q + n]) == M and not any(used[q:q + n])]
        if len(hits) != 1:
            continue
        start = hits[0]
        target = totals[start:start + n]
        E, aligned, shifts = _align_rows(prec, start, n, target)
        if E != 0:
            continue
        for k in range(start, start + n):
            used[k] = True
        contests.append({"office": _denormalize(office, district), "party": party,
                         "district": district or "",
                         "prec": [(name, "", vals) for name, vals in aligned],
                         "totals": list(target), "pages": []})
    return contests, notes


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------

def is_wide_page(md):
    """A wide page has colspan-grouped office headers and a bare 'TOTALS' row,
    and lacks the 'CANDIDATE TOTALS' label the NAME HEADING format uses."""
    if re.search(r"CANDIDATE\s+TOTALS", md, re.I):
        return False
    has_groups = len(_office_sequence(md)) >= 2
    has_totals = bool(re.search(r"<t[dh][^>]*>\s*TOTALS\s*</t[dh]>", md, re.I))
    return has_groups and has_totals
