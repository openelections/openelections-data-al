#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vision path (plan step 3): convert the "NAME HEADING CANVASS" matrix PDFs — the
ones `convert_precinct_pdfs.py` can't touch — into precinct-level CSVs.

These pages are image-only scans laid out as a matrix: precincts down the left,
candidates across the top, with candidate names printed *vertically*, one
character per row, stair-stepped across columns. There is no usable text layer.

Pipeline:

  render page (natural-pdf)  ->  nuextract3 content mode  ->  HTML tables
      ->  stitch contests across page breaks
      ->  checksum each contest against its printed CANDIDATE TOTALS row
      ->  name each column by joining its total to the county CSV
      ->  OpenElections CSV

Two things make this trustworthy without hand-built ground truth:

1. **Checksum gate.** Every contest prints a CANDIDATE TOTALS row. Extracted
   precinct rows must sum to it, per column. Contests that fail are reported and
   (by default) not written. This is the same idea as `src/total_checksum.py`.

2. **Names come from data, not from OCR.** The model reads the vote digits
   near-perfectly but does *not* decode the vertical name block. It doesn't need
   to: each column's total uniquely identifies its candidate against that
   county's roster in the county CSV. Measured on Butler, 14 of 15 contests have
   totals that uniquely identify every candidate; ties are reported, not guessed.

Requires a running Ollama with `numind/nuextract3:q4_k_m`.
NOTE: use the GGUF tag. The `:latest`/`:bf16` builds are MLX-packaged and emit
pure token garbage on Apple Silicon (Ollama 0.32.1).

Usage:
    python convert_canvass_pdfs.py <pdf> [<pdf> ...] [--dpi 200] [--validate-only]
    python convert_canvass_pdfs.py "2026 AL Republican.../Butler.../Butler County Canvas Report.pdf"


TODO: Houston, Lauderdale, Madison and Shelby
"""

import argparse
import difflib
import glob
import os
import re
import sys

import pandas as pd

from convert_precinct_pdfs import (
    COUNTY_CSV,
    OUT_DIR,
    norm_key,
    normalize_office,
    office_match_key,
    title_office,
)

MODEL = os.environ.get("NUEXTRACT_MODEL", "numind/nuextract3:q4_k_m")
CACHE_DIR = os.environ.get("CANVASS_CACHE", ".canvass_cache")

# Output files follow the OpenElections naming convention:
#   YYYYMMDD__al__<type>__<county>__precinct.csv
# Derive the election part from the county CSV's own name so the two stay in
# sync if COUNTY_CSV ever points at a different election.
ELECTION_PREFIX = re.sub(r"__county\.csv$", "", os.path.basename(COUNTY_CSV))

# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def extract_pages(pdf_path, dpi):
    """Yield (page_number, markdown) for every page, caching model output on disk.

    The model call is by far the slowest step (~40-170s/page), so cache keyed on
    pdf + page + dpi. Re-running to fix parsing logic then costs nothing.
    """
    from natural_pdf import PDF
    from ollama import chat

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
        # NuMind's documented "content extraction" invocation.
        r = chat(
            model=MODEL,
            messages=[
                {"role": "mode", "content": "content"},
                {"role": "user", "content": "", "images": [png]},
            ],
            think=False,
            options={"temperature": 0.2},
        )
        txt = r.message.content
        open(md_path, "w").write(txt)
        print(f"    p{i:03d} extracted", flush=True)
        yield i, txt


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

# [^>]* tolerates HTML attributes on the opening tag (e.g. <td style="...">,
# <th align="right">) — some counties' tables are styled, and a bare-tag regex
# silently fails to extract those cells at all, dropping whole contests with no
# error (checksum() sees zero rows and returns None, so nothing even prints).
CELL = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.S)
ROW = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S)
TABLE = re.compile(r"<table[^>]*>(.*?)</table>", re.S)
# Most counties number precincts 0001, 0002... but Lawrence uses bare 3-digit
# codes (001, 002...) — {3,4} is greedy, so a real 4-digit code still matches
# in full; only a genuinely 3-digit code falls back to 3. Every capture site
# below zero-pads to 4 digits so mixed-width sources still sort/join
# consistently with the rest of the output.
PRECINCT = re.compile(r"^(\d{3,4})\s*(.*)")
# Rare OCR misread: a leading "0" occasionally comes out as a visually similar
# capital letter instead of a digit ("J001" for "0001", Covington) — the row
# then misses the strict 4-digit match entirely and the whole precinct's data
# vanishes from output, not just its label. Narrow on purpose (single
# confusable letter + exactly 3 digits) to avoid matching content that isn't
# actually a precinct row.
PRECINCT_OCR_FALLBACK = re.compile(r"^([OQDJ])(\d{3})\s*(.*)")
VOTEFOR = re.compile(r"\(VOTE\s*FOR\)", re.I)
# Alternate office anchor for the "DISTRICT CANVASS" report layout (e.g.
# Tuscaloosa), which — unlike the "NAME HEADING CANVASS" layout — prints no
# "(VOTE FOR) 1" line under each office. Instead the office title sits just
# above a bare "<counted> OF <total> Precinct" line (e.g. "54 OF 55 Precinct"),
# so that marker serves the same role _office_above() keys off of. Anchored to
# the line start so it does NOT also match a summary page's parenthetical
# "(WITH 55 OF 55 Precinct COUNTED)".
PRECINCT_MARKER = re.compile(r"^\s*\d+\s+OF\s+\d+\s+Precinct", re.I)
CONTINUED = re.compile(r"CONTINUED FROM PREVIOUS PAGE", re.I)
# The model sometimes emits the totals row as plain text instead of a table row.
TOTALS_TEXT = re.compile(r"^CANDIDATE\s+TOTALS((?:\s+[\d,]+)+)\s*$", re.I | re.M)
# A contest split across a page break sometimes has its continuation rendered
# as plain text lines instead of an HTML table (Covington/Coffee/Marengo: the
# table-based parser sees nothing there at all, since it only ever looks
# inside <table> tags — this isn't a mis-parse, the rows are just invisible).
# "CODE NAME... NUM NUM..." — name is non-greedy so it absorbs only as much
# as needed to leave a pure trailing run of space-separated integers.
PLAIN_PRECINCT_TEXT = re.compile(r"^(\d{3,4})\s+(.+?)\s+((?:-?\d+\s+)*-?\d+)\s*$", re.M)
TOTALS_CELL = re.compile(r"CANDIDATE\s+TOTALS", re.I)
CHROME = re.compile(
    r"NAME HEADING|CANDIDATE HEADING|OFFICAL|OFFICIAL REPORT|RUN DATE|VOTE FOR|"
    r"CONTINUED|SUMMARY REPORT|REPORT-EL|PRIMARY ELECTION|ALABAMA\s+(REPUBLICAN|DEMOCRATIC)|"
    # DISTRICT CANVASS page chrome: the layout prints "DISTRICT CANVASS",
    # "PRINTED <date>", and a "PAGE 011.011.01" line above each office. Without
    # skipping these, _office_above() latches onto "PAGE 011.011.01" as the
    # office on continuation pages where the real title isn't repeated.
    r"DISTRICT CANVASS|PRINTED|PAGE\s+\d",
    re.I,
)
# DISTRICT CANVASS tables prepend per-precinct statistics columns — Registered
# Voters, Ballots Cast, Turnout Percentage — before the candidate columns. They
# are real numbers (not blanks), so the leading-None stripping can't remove them;
# they must be dropped by header detection or they masquerade as the first two or
# three candidates and wreck the column alignment and checksum.
STAT_COL_RE = re.compile(r"REGISTERED\s+VOTERS|BALLOTS\s+CAST|TURNOUT", re.I)
CAPS_LINE = re.compile(r"^#*\s*([A-Z][A-Z0-9 ,.'&\"/()-]{4,})\s*$")
# Some counties (e.g. Bullock) never print an inline (DEM)/(REP) marker at all —
# party is declared once, as page-header text like "ALABAMA REPUBLICAN P".
PARTY_HEADER_RE = re.compile(r"ALABAMA\s+(REPUBLICAN|DEMOCRATIC)", re.I)


def _num(s):
    s = str(s).strip().replace(",", "")
    # A stray trailing period is a real OCR artifact (Henry's US Senate totals
    # row had "27." for a plain 27) — without stripping it this cell parses as
    # unparseable, landing a None in a non-trailing position of an otherwise
    # all-numeric row, which is exactly the shape checksum()'s positional
    # alignment depends on getting right.
    s = s.rstrip(".")
    return int(s) if re.fullmatch(r"-?\d+", s) else None


def _strip_leading_none(vals):
    """Drop a leading run of unparseable cells.

    When the model can't transcribe a candidate-name header it sometimes emits
    a <figure> placeholder instead, and the data table that follows gets one
    spurious empty leading <td> per row that doesn't correspond to any real
    candidate column — confirmed by checking the printed CANDIDATE TOTALS row,
    which carries the same leading blank. Left in place this shifts every real
    column over by one and corrupts the checksum; worse, a contest split across
    a page break can have the placeholder on only one side, so raw column index
    stops lining up between the two halves entirely. Stripping only a *leading*
    run (never middle/trailing) is deliberately conservative: every observed
    case was a leading artifact, and None elsewhere already means "unparsed
    cell" and is excluded from sums as-is.
    """
    i = 0
    while i < len(vals) and vals[i] is None:
        i += 1
    return vals[i:]


DISTRICT_ONLY_RE = re.compile(r"DISTRICT\s+NO\.?\s*\d+\.?", re.I)


def _office_above(lines, idx):
    """The office is the caps line just above '(VOTE FOR)'.

    Anchoring on (VOTE FOR) matters: simply taking the nearest ALL-CAPS line
    picks up fragments of the vertical candidate-name block ("D S ST", "GE EE").
    """
    for j in range(idx - 1, max(-1, idx - 8), -1):
        s = lines[j].strip()
        # the model sometimes wraps the office header in markdown bold
        # ("**UNITED STATES REPRESENTATIVE...**"); strip it before matching,
        # or the header is invisible here and a contest split across a page
        # break fails to stitch (the office on this page never gets detected).
        s = re.sub(r"^\*+|\*+$", "", s).strip()
        if not s or CHROME.search(s):
            continue
        m = CAPS_LINE.match(s)
        if not m:
            continue
        name = m.group(1).strip()
        # reject single-letter-per-word noise from the vertical name block
        if re.fullmatch(r"[A-Z](\s+[A-Z])*", name):
            continue
        if DISTRICT_ONLY_RE.fullmatch(name):
            # State Democratic/Republican Executive Committee races print
            # across two lines — "STATE DEMOCRATIC EXECUTIVE COMMITTEE
            # (FEMALE)," then "DISTRICT NO. 89" — and a bare district number
            # isn't a real office on its own (normalize_office() has nothing
            # to map it to, so it fails to resolve against the county CSV
            # entirely). Climb one more line for the real title and join them,
            # matching convert_precinct_pdfs.py's handling of the same source
            # quirk in the text-layer parser.
            for k in range(j - 1, max(-1, j - 4), -1):
                s2 = re.sub(r"^\*+|\*+$", "", lines[k].strip()).strip()
                if not s2:
                    continue
                if CHROME.search(s2):
                    break
                m2 = CAPS_LINE.match(s2)
                if m2:
                    return f"{m2.group(1).strip()} {name}"
                break
            return name
        return name
    return None


def parse_page(txt, page):
    """Return one dict per data block on the page, in document order.

    Two quirks of the model's output force this to look beyond the table itself:

    * The (DEM)/(REP) markers often land in a *header-only* table holding the
      vertical name block, which carries no rows and would otherwise be dropped
      along with the party. So party is taken from the nearest marker anywhere
      at or before the table's end, not just from inside it.
    * A CANDIDATE TOTALS row is sometimes emitted as plain text outside any
      table, and for a contest continued across a page break it can land before
      the page's first table. Plain-text totals are therefore collected as
      standalone blocks positioned in document order.

    HTML whitespace entities (&nbsp;, &emsp;) are normalized to real spaces up
    front — Geneva's plain-text "CANDIDATE TOTALS&nbsp;&nbsp;...2054&nbsp;..."
    used literal entity text as column spacing instead of real whitespace
    characters, so TOTALS_TEXT's \\s+ never matched it at all and the contest
    silently had no totals. Anything relying on \\s+ downstream benefits.

    Markdown bold markers are also stripped globally here — Morgan's
    plain-text totals line wrapped every number in "**" ("CANDIDATE TOTALS
    **480** **321** ..."), which TOTALS_TEXT's [\\d,]+ can't see through
    (asterisks aren't digits), so all 14 of that document's REP contests had
    no totals at all. _office_above() already stripped **-wrapped office
    headers locally; doing it here too covers every other use of \\d+/\\s+
    against the raw text, not just that one call site.
    """
    txt = re.sub(r"&nbsp;|&emsp;", " ", txt)
    txt = txt.replace("**", "")
    lines = txt.split("\n")
    offices, off = [], 0
    for i, ln in enumerate(lines):
        # Two office anchors for two report layouts: "(VOTE FOR)" (NAME HEADING
        # CANVASS) and a line-leading "N OF M Precinct" (DISTRICT CANVASS). A
        # given page uses one or the other, and their line shapes don't overlap,
        # so recognizing both here is additive — it never double-detects on the
        # layouts that already worked.
        if VOTEFOR.search(ln) or PRECINCT_MARKER.search(ln):
            offices.append((off, _office_above(lines, i)))
        off += len(ln) + 1

    # every party marker on the page, with its position — either an inline
    # (DEM)/(REP) tag on a table, or a page-header declaration like
    # "ALABAMA REPUBLICAN P" for counties that never tag tables individually.
    markers = [(m.start(), m.group(1)) for m in re.finditer(r"\((DEM|REP)\)", txt)]
    markers += [
        (m.start(), "REP" if m.group(1).upper().startswith("REP") else "DEM")
        for m in PARTY_HEADER_RE.finditer(txt)
    ]
    markers.sort()

    def office_at(pos):
        o = None
        for p, name in offices:
            if p < pos and name:
                o = name
        return o

    def party_at(pos):
        p = None
        for mpos, party in markers:
            if mpos <= pos:
                p = party
        return p

    events = []
    table_spans = []
    for m in TABLE.finditer(txt):
        tbl = m.group(1)
        table_spans.append((m.start(), m.end()))
        prec, totals = [], None
        # Sometimes the totals LABEL lands in a <thead> header cell instead of
        # a <tbody> row (confirmed on Coosa: <thead><th>CANDIDATE TOTALS</th>...),
        # while the actual VALUES are still a normal-looking <tbody> row — just
        # with an empty label, since the label text got displaced into the
        # header. Detect that shape once per table so the row loop below can
        # fall back to it only when nothing else claimed the totals.
        thead_m = re.search(r"<thead>(.*?)</thead>", tbl, re.S)
        totals_in_thead = bool(thead_m and TOTALS_CELL.search(thead_m.group(1)))
        # DISTRICT CANVASS: count the leading statistics columns (Registered
        # Voters / Ballots Cast / Turnout) from the header so they can be sliced
        # off every data and totals row below. NAME HEADING tables have none, so
        # lead_skip stays 0 and nothing changes for them.
        lead_skip = len(STAT_COL_RE.findall(thead_m.group(1))) if thead_m else 0
        for tr in ROW.findall(tbl):
            cells = [re.sub(r"<.*?>", "", c).strip() for c in CELL.findall(tr)]
            if not cells:
                continue
            label = cells[0]
            # Usually one candidate's value per <td>. Some tables (Covington's
            # US House 1st CD) instead pack a whole row's values into a single
            # cell as a space-separated string ("239 92 260 133 64 42 150") —
            # _num() can't parse that as one integer, so the entire cell
            # silently became nothing. Split multi-token cells into separate
            # values; a normal single-value or empty cell has nothing to
            # split, so this changes nothing for the common case (an empty
            # cell must stay exactly one None slot, not disappear, or it
            # breaks the leading/all-blank column logic elsewhere).
            raw_tokens = []
            for c in cells[1:]:
                parts = c.split()
                raw_tokens.extend(parts) if len(parts) > 1 else raw_tokens.append(c)
            raw_vals = [_num(t) for t in raw_tokens]
            if lead_skip:
                # drop the leading RegVoters/BallotsCast/Turnout columns before
                # any None-stripping, so candidate columns start at index 0
                raw_vals = raw_vals[lead_skip:]
            vals = _strip_leading_none(raw_vals)
            pm = PRECINCT.match(label)
            if not pm:
                pm_ocr = PRECINCT_OCR_FALLBACK.match(label)
                if pm_ocr:
                    prec.append(("0" + pm_ocr.group(2), pm_ocr.group(3).strip(), vals))
                    continue
            if pm:
                prec.append((pm.group(1).zfill(4), pm.group(2).strip(), vals))
            elif TOTALS_CELL.search(label):
                totals = vals
            elif (label == "" and len(cells) > 1 and TOTALS_CELL.search(cells[1])):
                # A phantom leading empty cell sometimes pushes "CANDIDATE
                # TOTALS" itself into the second <td> instead of the first
                # (Geneva) — <td></td><td>CANDIDATE TOTALS</td><td>314</td>...
                # No other row in the same table showed this shape (precinct
                # rows aren't shifted), just the totals row, so re-derive vals
                # from cells[2:] rather than the cells[1:] already computed
                # above for the normal case.
                #
                # Guard against a lookalike that isn't a totals row at all: a
                # <thead> row shaped <th></th><th>CANDIDATE TOTALS</th>
                # <th>CANDIDATE PERCENT</th> (Pike) matches this same shape —
                # empty first cell, "CANDIDATE TOTALS" second — but its
                # remaining cells are more header text, not numbers. Left
                # unguarded this claims `totals` with garbage/empty content
                # and blocks the real tbody totals row from ever being tried
                # by the totals_in_thead fallback below (whose own `totals is
                # None` check would otherwise catch it). Require at least one
                # real number, exactly like that fallback already does.
                raw2 = []
                for c in cells[2:]:
                    parts = c.split()
                    raw2.extend(parts) if len(parts) > 1 else raw2.append(c)
                candidate_totals = _strip_leading_none([_num(t) for t in raw2])
                if any(t is not None for t in candidate_totals):
                    totals = candidate_totals
            elif (totals_in_thead and totals is None and label == ""
                    and any(v is not None for v in vals)):
                totals = vals
        if prec or totals:
            events.append((m.start(), {
                "office": office_at(m.start()), "party": party_at(m.end()),
                "prec": prec, "totals": totals, "page": page,
                "continued": bool(CONTINUED.search(txt[max(0, m.start() - 600):m.start()])),
            }))

    # plain-text totals rows that fall outside every table
    for tm in TOTALS_TEXT.finditer(txt):
        if any(a <= tm.start() < b for a, b in table_spans):
            continue
        events.append((tm.start(), {
            "office": office_at(tm.start()), "party": party_at(tm.start()),
            "prec": [], "totals": [int(x) for x in tm.group(1).split()],
            "page": page, "continued": True,
        }))

    # plain-text precinct rows that fall outside every table — a page-break
    # continuation sometimes renders as bare "CODE NAME NUM NUM..." lines
    # instead of continuing the HTML table (Covington/Coffee/Marengo). One
    # event per line is enough: stitch() merges consecutive same-office
    # blocks regardless of granularity, so there's no need to group runs.
    for pm2 in PLAIN_PRECINCT_TEXT.finditer(txt):
        if any(a <= pm2.start() < b for a, b in table_spans):
            continue
        vals = _strip_leading_none([_num(x) for x in pm2.group(3).split()])
        events.append((pm2.start(), {
            "office": office_at(pm2.start()), "party": party_at(pm2.start()),
            "prec": [(pm2.group(1).zfill(4), pm2.group(2).strip(), vals)], "totals": None,
            "page": page, "continued": True,
        }))

    return [b for _, b in sorted(events, key=lambda e: e[0])]


def _office_fuzzy_match(a, b):
    """True if two office strings likely name the same contest despite an OCR
    spelling slip on one page (Escambia: "MEMBER, ESCAMBIA COUNTY COMMISSION,
    DISTRICT NO. 3" on the page that opens the contest, "MEMBER, ESCAMBA
    COUNTY COMMISSION, DISTRICT NO. 3" — missing a letter — on its
    continuation page). Only ever consulted as a fallback when the model's
    own "(CONTINUED FROM PREVIOUS PAGE)" marker is present — that's the real
    evidence these belong together; the fuzzy text match is just confirming
    it wasn't actually a same-looking but different contest that happened to
    follow it. Digits must match exactly: a Place/District number difference
    (PSC Place No. 1 vs Place No. 2) means a genuinely different race and
    must never be treated as a spelling variant of the same one.
    """
    digits = lambda s: re.findall(r"\d+", s)
    if digits(a) != digits(b):
        return False
    strip = lambda s: re.sub(r"\d+", "", s.upper())
    return difflib.SequenceMatcher(None, strip(a), strip(b)).ratio() >= 0.9


def stitch(blocks):
    """Merge table blocks into contests.

    Contests run across page breaks (marked '(CONTINUED FROM PREVIOUS PAGE)'), so
    precinct rows must be accumulated before checksumming. A contest is a
    contiguous run sharing one office — keeping DEM GOVERNOR and REP GOVERNOR
    separate, since they are never adjacent.
    """
    contests = []
    for b in blocks:
        if b["office"] is None:
            # statistics table (registered voters / ballots cast) — no candidates
            continue
        prev = contests[-1] if contests else None
        same = (
            prev is not None
            and (b["continued"] or prev["totals"] is None)
            and (prev["office"] == b["office"]
                 or (b["continued"] and _office_fuzzy_match(prev["office"], b["office"])))
        )
        if same:
            cur = prev
        else:
            cur = {"office": b["office"], "party": b["party"],
                   "prec": [], "totals": None, "pages": []}
            contests.append(cur)
        cur["prec"].extend(b["prec"])
        if b["totals"] is not None:
            cur["totals"] = b["totals"]
        if b["party"] and not cur["party"]:
            cur["party"] = b["party"]
        if b["page"] not in cur["pages"]:
            cur["pages"].append(b["page"])
    return contests


def _drop_empty_columns(contest):
    """Drop raw column positions that never hold data in any precinct row.

    Some tables fragment the candidate-name header into far more raw columns
    than real candidates — Coosa's Secretary of State table has 11 raw value
    columns for 3 real candidates, 8 of them blank in every single precinct
    row. Left alone, positional column indexing lines up real vote data with
    phantom columns and the checksum fails not because the data is wrong but
    because the alignment is. A column that is None in every precinct row
    genuinely has no data anywhere, so compressing it out (preserving the
    left-to-right order of the columns that remain) recovers the true
    candidate values in their original order — the same left-to-right order
    the totals row and name_columns()'s join already rely on.

    Mutates contest["prec"] in place so every downstream consumer (checksum,
    name_columns, the row-writing loop) sees the corrected columns, not just
    the checksum comparison — and always rebuilds every row to the same
    length, even when the populated-column set is already contiguous. Skipping
    the rebuild in that case (an earlier version did, as an optimization) is
    unsafe: a single precinct the model failed to transcribe at all comes back
    as an empty vals list, and if downstream code takes any one row's raw
    length as authoritative (checksum() does, for ncols), that empty row can
    silently claim the whole contest has zero real columns.
    """
    rows = [v for _, _, v in contest["prec"]]
    if not rows:
        return
    width = max(len(r) for r in rows)
    if width == 0:
        return
    populated = [i for i in range(width) if any(i < len(r) and r[i] is not None for r in rows)]
    if not populated:
        return
    contest["prec"] = [
        (code, name, [v[i] if i < len(v) else None for i in populated])
        for code, name, v in contest["prec"]
    ]


def checksum(contest):
    """Compare summed precinct rows to the printed CANDIDATE TOTALS row.

    The (post-compression) precinct-row width is authoritative for the column
    count, not a count of totals that happen to be non-None. Those are not
    the same thing whenever a total is unparseable somewhere other than the
    very end — confirmed on Henry's US Senate: a 7-column totals row had "27."
    (trailing-period OCR artifact) at position 5, one column short of the end.
    Counting the 6 non-None values and slicing the first 6 raw positions
    (totals[:6]) grabbed positions 0-5 — dropping the real 7th candidate's
    total entirely and leaving the unparseable None sitting at the end of the
    slice where it looked, misleadingly, like ordinary trailing padding.
    Comparing position-by-position against the full-width totals row (short
    positions treated as unverifiable, not padded to look like a match) keeps
    every real column in its true place regardless of where the gap is.
    """
    _drop_empty_columns(contest)
    rows = [v for _, _, v in contest["prec"]]
    if not rows:
        return None
    if contest["totals"] is None:
        return {"status": "no-totals", "ncols": max(len(r) for r in rows)}
    ncols = len(rows[0])
    totals = contest["totals"]
    printed = [totals[i] if i < len(totals) else None for i in range(ncols)]
    computed = [sum(r[i] for r in rows if i < len(r) and r[i] is not None)
                for i in range(ncols)]
    dupes = len(rows) - len({p for p, _, _ in contest["prec"]})
    mismatches = [i for i in range(ncols) if printed[i] is not None and computed[i] != printed[i]]
    ok = not mismatches and not dupes
    contest["totals"] = printed  # keep name_columns() aligned with this same positional mapping
    return {"status": "PASS" if ok else "FAIL", "ncols": ncols,
            "computed": computed, "printed": printed, "dupes": dupes}


# ---------------------------------------------------------------------------
# Naming columns from the county CSV
# ---------------------------------------------------------------------------

SUFFIXES = {"JR", "JR.", "SR", "SR.", "II", "III", "IV"}


def _surname_key(name):
    """Sort key approximating ballot order, which is alphabetical by surname."""
    parts = [p for p in re.split(r"\s+", name.strip().strip('"')) if p]
    while len(parts) > 1 and parts[-1].upper().rstrip(".") in {s.rstrip(".") for s in SUFFIXES}:
        parts.pop()
    return norm_key(parts[-1]) if parts else norm_key(name)


def _totals_pool(county, party, office, district, county_df):
    """(votes -> [candidate names]) for one county/party/office/district."""
    key = office_match_key(office)
    pool = county_df[
        (county_df.county == county)
        & (county_df.party == party)
        & (county_df.office.map(office_match_key) == key)
    ]
    if district:
        pool = pool[pool.district.astype(str) == str(district)]
    by_total = {}
    for _, r in pool.iterrows():
        by_total.setdefault(int(r.votes), []).append(r.candidate)
    return by_total


def infer_party(contest, county, county_df):
    """Infer party from which party's candidates the printed totals match.

    Some documents drop the party signal entirely — no inline (DEM)/(REP) tag
    and no page-header declaration (see Dale: the model emitted an image
    placeholder for a whole header block, taking the party tag with it). Rather
    than dropping the contest, reuse the same totals-join used for candidate
    names: a primary contest is single-party, so whichever party's candidate
    pool for this office the printed totals match is almost certainly the real
    party. This is an inference from data, not a fact read off the page, so the
    caller must log it — never treat it as equivalent to a real marker.

    Returns the inferred party, or None if the match is empty or tied between
    both parties (never guess when the evidence doesn't clearly favor one).
    """
    office, district = normalize_office(contest["office"])
    totals = contest["totals"][: contest["ncols"]]
    scores = {}
    for party in ("DEM", "REP"):
        by_total = _totals_pool(county, party, office, district, county_df)
        scores[party] = sum(1 for t in totals if t in by_total)
    dem, rep = scores["DEM"], scores["REP"]
    if dem == rep:
        return None
    return "DEM" if dem > rep else "REP"


def name_columns(contest, county, party, county_df):
    """Map each column index to a candidate name via its printed total.

    The vision model does not decode the vertical name block, so identity comes
    from authoritative data instead: within one contest a candidate's county-wide
    total is (almost always) unique, so the totals row is a join key.

    Ballot questions have no candidates in the county CSV; their two columns are
    the Yes/No tallies.

    Ties (two candidates with the same county total) are resolved by position:
    these reports print candidates in ballot order, which is alphabetical by
    surname, so N tied columns map in order onto the N tied names sorted the same
    way. That is an inference, not a checksum, so it is always reported.

    Returns (names, (office, district), notes); names[i] is None where the join
    could not be resolved, which the caller must skip rather than guess.
    """
    office, district = normalize_office(contest["office"])
    totals = contest["totals"][: contest["ncols"]]

    if re.match(r"PROPOSED\b", contest["office"], re.I):
        # ballot measure: columns are YES / NO in printed order
        if len(totals) == 2:
            return ["Yes", "No"], (office, district), ["ballot measure: columns assumed Yes/No"]
        return [None] * len(totals), (office, district), ["ballot measure: unexpected column count"]

    by_total = _totals_pool(county, party, office, district, county_df)
    names, notes = [None] * len(totals), []
    tied = {}
    for i, tot in enumerate(totals):
        cands = by_total.get(tot, [])
        if len(cands) == 1:
            names[i] = cands[0]
        elif len(cands) > 1:
            tied.setdefault(tot, []).append(i)
        else:
            notes.append(f"col{i+1}: total {tot} not in county CSV")

    for tot, cols in tied.items():
        cands = by_total[tot]
        if len(cols) == len(cands):
            ordered = sorted(cands, key=_surname_key)
            for col, nm in zip(sorted(cols), ordered):
                names[col] = nm
            notes.append(
                f"cols {[c+1 for c in sorted(cols)]}: total {tot} tied, "
                f"assigned by ballot order -> {ordered}"
            )
        else:
            notes.append(f"cols {[c+1 for c in cols]}: total {tot} ties {cands} (unresolved)")

    return names, (office, district), notes


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def detect_county(pdf_path):
    """County name from the containing folder, e.g. 'Butler County Blue Sheet...'."""
    folder = os.path.basename(os.path.dirname(pdf_path))
    m = re.match(r"([A-Za-z .']+?)\s+County", folder)
    return m.group(1).strip() if m else folder.split()[0]


def process(pdf_paths, dpi, county_df, validate_only, extract_fn=extract_pages, out_dir=OUT_DIR):
    """Process every PDF belonging to one county into a single output CSV.

    Most counties are one PDF. Geneva is the known exception: its office-level
    results are split across 13 separate single-page PDFs rather than bundled
    into one file, so all of a county's PDFs must be extracted and stitched
    together before the checksum/naming/output steps, or all but the last file
    processed would be silently discarded.

    extract_fn is injectable (defaulting to this module's Ollama/nuextract3
    extract_pages) so an alternate backend — PaddleOCR-VL
    (convert_canvass_pdfs_paddleocr.py) or Anthropic Claude
    (convert_canvass_pdfs_claude.py) — can reuse every parsing/checksum/
    naming/never-drop rule below unchanged. Swap only the extraction step and
    every future fix to this function benefits both instead of needing to be
    ported between two copies.

    out_dir is likewise injectable, defaulting to the real 2026/counties/
    output directory: an alternate backend being evaluated for quality must
    not silently overwrite an already-verified CSV in that directory just
    because it processes the same county.
    """
    county = detect_county(pdf_paths[0])
    names = ", ".join(os.path.basename(p) for p in pdf_paths)
    print(f"\n=== {county}: {names} ===", flush=True)

    blocks = []
    for pdf_path in pdf_paths:
        for page, txt in extract_fn(pdf_path, dpi):
            blocks.extend(parse_page(txt, page))
    contests = stitch(blocks)

    rows, n_pass, n_fail, n_fail_written, unresolved = [], 0, 0, 0, []
    for c in contests:
        res = checksum(c)
        if res is None:
            continue
        c["ncols"] = res["ncols"]
        tag = res["status"]
        label = f"{c['office'][:44]:44s} {c['party'] or '   '} pp{c['pages']}"
        if tag == "no-totals":
            # No printed totals row exists anywhere for this contest, so there
            # is no vote count to join candidate identity against at all —
            # unlike a FAIL (below), there's nothing here to write with any
            # confidence in *which* column is which candidate.
            n_fail += 1
            print(f"  [{tag:>9}] {label}")
            continue
        verified = tag == "PASS"
        if verified:
            n_pass += 1
        else:
            # FAIL: the printed totals row exists and candidate identity can
            # still be resolved against it, the precinct rows just don't sum to
            # it. Per user direction, write it anyway rather than dropping real
            # vote data — but count and flag it distinctly so it's never
            # confused with checksum-verified output.
            n_fail += 1
            n_fail_written += 1
            print(f"  [    FAIL] {label}  (writing anyway — checksum mismatch below)")
            print(f"              computed={res['computed']}")
            print(f"              printed ={res['printed']}"
                  + (f"  dupes={res['dupes']}" if res["dupes"] else ""))
            unresolved.append(f"{c['office']}: CHECKSUM FAILED, included anyway "
                               f"(computed={res['computed']} printed={res['printed']})")
        party = c["party"]
        is_measure = bool(re.match(r"PROPOSED\b", c["office"], re.I))
        if is_measure:
            # The canvass prints no (DEM)/(REP) marker for ballot questions; any
            # party here is leakage from the preceding contest, so drop it rather
            # than assert an attribution the source doesn't make.
            party = None
        elif not party:
            # No marker anywhere in the source text for this contest (inline tag,
            # nor page-header declaration). Rather than drop real vote data,
            # infer party from which party's candidates the totals match — see
            # infer_party()'s docstring for why this is safe to log but not to
            # treat as a read fact.
            party = infer_party(c, county, county_df)
            if party:
                unresolved.append(f"{c['office']}: no party marker, inferred {party} from vote totals")
            else:
                unresolved.append(f"{c['office']}: no party marker, could not infer (dropped)")
                print(f"  [{'PASS' if verified else 'FAIL':>9}] {label}  (no party, inference failed — skipped)")
                continue
        names, (office, district), notes = name_columns(c, county, party, county_df)
        inferred_tag = f"  [party inferred: {party}]" if not c["party"] and not is_measure else ""
        status_tag = "" if verified else "  [CHECKSUM FAILED]"
        n_unresolved_names = sum(1 for n in names if n is None)
        print(f"  [{'PASS' if verified else 'FAIL':>9}] {label} precincts={len(c['prec'])} cols={res['ncols']}"
              + (f"  FLAGGED (included)={n_unresolved_names}" if n_unresolved_names else "")
              + inferred_tag + status_tag)
        for n in notes:
            unresolved.append(f"{c['office']} [{party}] {n}")
        for code, pname, vals in c["prec"]:
            for i, nm in enumerate(names):
                if i >= len(vals) or vals[i] is None:
                    continue
                # A column whose candidate couldn't be matched to the county
                # CSV is still real vote data — some contests genuinely aren't
                # in the reference file at all (very local races), and a
                # single missed total shouldn't erase the other columns'
                # votes. Include it under a clearly-flagged placeholder rather
                # than silently dropping it; unlike a normal candidate name,
                # this one is not verified against anything.
                candidate = nm if nm is not None else f"Unverified Candidate {i + 1}"
                rows.append({
                    "county": county,
                    "precinct": f"{code} {pname}".strip(),
                    "office": office,
                    "district": district,
                    "party": party or "",
                    "candidate": candidate,
                    "votes": vals[i],
                })

    print(f"  -> contests: {n_pass} verified, {n_fail} failed checksum"
          + (f" ({n_fail_written} written anyway, {n_fail - n_fail_written} had no totals to join at all)"
             if n_fail else ""))
    if unresolved:
        print(f"  -> {len(unresolved)} unresolved column(s):")
        for u in unresolved[:10]:
            print(f"       {u}")

    if rows and not validate_only:
        df = pd.DataFrame(rows, columns=["county", "precinct", "office", "district",
                                         "party", "candidate", "votes"])
        # The checksum gate validates digits, not labels, so a precinct's *name*
        # can vary between contests when OCR wobbles (e.g. MCLAINS/MCCLAINS).
        # The 4-digit code is stable, so canonicalize each code to its most
        # common spelling and report any code that needed it.
        df["code"] = df.precinct.str[:4]
        canon = df.groupby("code").precinct.agg(lambda s: s.value_counts().idxmax())
        for code, variants in df.groupby("code").precinct.unique().items():
            if len(variants) > 1:
                print(f"  -> precinct {code} name varied {sorted(variants)}; "
                      f"using {canon[code]!r}")
        df["precinct"] = df.code.map(canon)
        df = df.drop(columns="code")
        df = df.sort_values(["precinct", "party", "office", "district", "candidate"])
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, f"{ELECTION_PREFIX}__{county.lower().replace(' ', '_')}__precinct.csv")
        df.to_csv(out, index=False)
        print(f"  -> wrote {len(df)} rows to {out}")
    return n_pass, n_fail


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+", help="canvass PDF path(s); globs allowed")
    ap.add_argument("--dpi", type=int, default=200,
                    help="render resolution (default 200; lower is faster, less accurate)")
    ap.add_argument("--validate-only", action="store_true",
                    help="run the checksum gate but don't write CSVs")
    args = ap.parse_args()

    county_df = pd.read_csv(COUNTY_CSV, dtype=str, keep_default_na=False)
    county_df["votes"] = county_df["votes"].astype(int)

    paths = []
    for p in args.pdfs:
        paths.extend(sorted(glob.glob(p)) if any(c in p for c in "*?[") else [p])

    by_county = {}
    for p in paths:
        if not os.path.exists(p):
            print(f"skip (missing): {p}", file=sys.stderr)
            continue
        by_county.setdefault(detect_county(p), []).append(p)

    tp = tf = 0
    for county_paths in by_county.values():
        a, b = process(sorted(county_paths), args.dpi, county_df, args.validate_only)
        tp += a
        tf += b
    print(f"\n==== {len(by_county)} count{'y' if len(by_county)==1 else 'ies'} "
          f"({len(paths)} PDF file(s)): {tp} contests verified, {tf} failed ====")
    return 1 if tf else 0


if __name__ == "__main__":
    sys.exit(main())
