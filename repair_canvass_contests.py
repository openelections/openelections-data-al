#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Verification-gated repair driver for the 2026 AL primary canvass CSVs.

Motivation: OCR of the NAME HEADING / DISTRICT CANVASS matrix PDFs is uneven.
`convert_canvass_pdfs.py` already validates every contest against its printed
CANDIDATE TOTALS row (the checksum gate) and names columns from the county CSV
rather than trusting OCR'd names. That makes any extraction *provably* right or
wrong per contest, so we can escalate: re-extract the *failing* contests with a
higher render resolution and/or a stronger vision model, and only merge a
contest back into 2026/counties/ when its precinct rows now sum to the printed
totals.

This module never regenerates a whole county in place. It:

  1. analyze()  — runs extract -> parse -> stitch -> checksum for a county's
                  PDF(s) with a chosen backend, returning structured per-contest
                  results (rows + checksum status) WITHOUT writing anything.
                  Mirrors convert_canvass_pdfs.process()'s core exactly, reusing
                  its parse_page/stitch/checksum/name_columns so any future fix
                  there benefits this too.
  2. worklist() — the subset of analyze()'s results that FAIL or have no totals.
  3. merge_passing() — for each contest that PASSES the checksum in a new run,
                  replace ONLY that contest's rows (matched by office/district/
                  party) in the existing county CSV. A passing contest never
                  regresses a contest that isn't in the new run.

Backends (extract_fn, all with the signature convert_canvass_pdfs expects,
yielding (page_number, markdown)):

  * nuextract    — the repo baseline (convert_canvass_pdfs.extract_pages).
  * paddleocr    — PaddleOCR-VL via convert_canvass_pdfs_paddleocr (primary).
  * a claude model via convert_canvass_pdfs_claude.make_extract_fn(model),
    optionally at a higher --dpi (300) than the cached 200.

Usage:
    # Just report each county's per-contest checksum status from a backend:
    python repair_canvass_contests.py analyze <county-folder-or-pdf> \\
        [--model paddleocr] [--dpi 300]

    # Re-extract failing contests with a stronger model and merge the ones that
    # now pass into 2026/counties/ (writes; prints every replacement):
    python repair_canvass_contests.py repair <county-folder-or-pdf> \\
        --model paddleocr --dpi 300 [--dry-run]

A "county-folder-or-pdf" may be a single PDF, a glob, or a county folder under
"2026 AL Republican Party Primary Precinct Results/" (all its PDFs are used).
"""

import argparse
import glob
import os
import re
import sys

import pandas as pd

import convert_canvass_pdfs as base
from convert_canvass_pdfs import (
    COUNTY_CSV,
    ELECTION_PREFIX,
    OUT_DIR,
    checksum,
    detect_county,
    infer_party,
    name_columns,
    parse_page,
    stitch,
)
from convert_precinct_pdfs import normalize_office, office_match_key

CSV_COLUMNS = ["county", "precinct", "office", "district", "party", "candidate", "votes"]


# ---------------------------------------------------------------------------
# Analysis: structured per-contest results, no writing
# ---------------------------------------------------------------------------

def resolve_dupes(contest, res):
    """When a contest has duplicate precinct codes, collapse each code to one row,
    choosing sum-vs-keep-first by which reproduces the printed CANDIDATE TOTALS.

    A precinct code that appears twice is almost always one physical precinct
    whose row straddled a page break: sometimes each half carries a *partial*
    count (Baldwin GOVERNOR: 88+22, 113+41 — sum is right), sometimes the whole
    row is simply reprinted verbatim on both pages (equal-valued pairs — keep one
    is right). We do not guess: checksum() has already pinned contest["totals"]
    to the printed positional totals, so we try both reducers and accept the one
    whose per-column sums match the printed row exactly. If neither matches, the
    duplication isn't the (only) problem and the contest stays FAIL for OCR
    repair.

    On success, mutates contest["prec"] to the collapsed rows and returns
    (new_status, note); otherwise returns (res["status"], None).
    """
    prec = contest["prec"]
    codes = [c for c, _, _ in prec]
    n_dupe = len(codes) - len(set(codes))
    if n_dupe == 0 or contest.get("totals") is None:
        return res["status"], None
    ncols = contest["ncols"]
    printed = contest["totals"]

    def collapse(reducer):
        by = {}
        order = []
        for code, nm, vals in prec:
            if code not in by:
                by[code] = [nm, list(vals)]
                order.append(code)
            else:
                ex = by[code][1]
                width = max(len(ex), len(vals))
                merged = []
                for i in range(width):
                    a = ex[i] if i < len(ex) else None
                    b = vals[i] if i < len(vals) else None
                    merged.append(reducer(a, b))
                by[code][1] = merged
        return [(code, by[code][0], by[code][1]) for code in order]

    def col_sums(rows):
        return [sum(r[2][i] for r in rows if i < len(r[2]) and r[2][i] is not None)
                for i in range(ncols)]

    def matches(rows):
        cs = col_sums(rows)
        return all(printed[i] is None or cs[i] == printed[i] for i in range(ncols))

    sum_reduce = lambda a, b: (a or 0) + (b or 0) if (a is not None or b is not None) else None
    keep_reduce = lambda a, b: a if a is not None else b

    for label, reducer in [("sum", sum_reduce), ("keep-first", keep_reduce)]:
        collapsed = collapse(reducer)
        if matches(collapsed):
            contest["prec"] = collapsed
            return "PASS", f"collapsed {n_dupe} split precinct row(s) by {label}"
    return res["status"], None


def stitch_carry(blocks):
    """Like convert_canvass_pdfs.stitch(), but carries the current office/party
    onto blocks the parser left office=None. Some canvasses (Madison via
    PaddleOCR) print a contest's 81-precinct matrix as several tables and repeat
    the office header only on the first; base stitch() drops the continuation
    tables (office=None) as statistics, losing most of the vote. Carrying the
    last-seen office/party recovers them. Wrong carries are harmless: the pooled
    precinct rows then won't sum to the printed CANDIDATE TOTALS, so the contest
    FAILs the checksum and is not merged (and never beats a real PASS, since the
    caller keeps best-status per contest key)."""
    contests = []
    cur_office = cur_party = None
    for b in blocks:
        office = b["office"] or cur_office
        party = b["party"] or cur_party
        if b["office"]:
            cur_office = b["office"]
        if b["party"]:
            cur_party = b["party"]
        if office is None:
            continue
        prev = contests[-1] if contests else None
        # A precinct code repeating is the reliable signal that a new contest has
        # begun even though the office header was missed/carried: within one
        # contest each precinct appears once, so a code already present in the
        # current contest means this block belongs to the next one. Without this,
        # carry over-pools an adjacent contest's tables (Madison Lt Gov DEM
        # ballooning to 92 of 81 precincts).
        repeats = (prev is not None
                   and {c for c, _, _ in b["prec"]} & {c for c, _, _ in prev["prec"]})
        same = (prev is not None and prev["office"] == office and not repeats
                and (prev["totals"] is None or b["continued"] or not b["office"]))
        if same:
            cur = prev
        else:
            cur = {"office": office, "party": party, "prec": [], "totals": None, "pages": []}
            contests.append(cur)
        cur["prec"].extend(b["prec"])
        if b["totals"] is not None:
            cur["totals"] = b["totals"]
        if party and not cur["party"]:
            cur["party"] = party
        if b["page"] not in cur["pages"]:
            cur["pages"].append(b["page"])
    return contests


def analyze(pdf_paths, dpi, extract_fn, county_df, county=None):
    """Return [(contest, checksum_result)] for a county's PDFs, no side effects.

    Reproduces convert_canvass_pdfs.process()'s extract->parse->stitch->checksum
    core so callers get structured results (contest dict + checksum dict) instead
    of the printed summary + written CSV that process() produces. checksum()
    mutates the contest in place (drops all-blank columns, pins the positional
    totals), exactly as process() relies on, so downstream name_columns() sees
    the same aligned columns.

    A contest that fails only because of duplicated (page-break-split) precinct
    rows is repaired here via resolve_dupes() and its status upgraded to PASS,
    with the reducer recorded in res["dedup_note"] — this is the same no-OCR fix
    the repo applied by hand to Lamar/Lee, now gated by the printed totals.

    Pages in the wide multi-office format (see convert_wide_canvass) are routed
    to that parser instead of parse_page when `county` is given — it needs the
    county to anchor office columns on the authoritative totals. Each wide page
    is tried under both parties (these reports are single-party per page; the
    authority-multiset gate makes the wrong party yield nothing). The same
    (office, district, party) can repeat across wide pages, so wide results are
    deduplicated keeping the best status — otherwise merge_passing would see a
    key collision and refuse the contest entirely.
    """
    import convert_wide_canvass as wide

    # Materialize pages so wide-format handling can stitch a header-only page
    # (office colspans, no data) onto the following data page (precincts+TOTALS,
    # no headers) — the Mobile layout splits them across a page break.
    pages = [(page, txt) for pdf_path in pdf_paths for page, txt in extract_fn(pdf_path, dpi)]

    _has_totals = lambda md: bool(re.search(r">\s*TOTALS\s*<", md, re.I))

    # A wide "contest group" starts at a page with >=2 colspan office headers and
    # runs through the page that carries the TOTALS row — Mobile splits one
    # contest's offices, precincts and TOTALS across up to a few pages; Cherokee
    # keeps them on one. Concatenate the group so parse_wide_page/scan see the
    # office headers, every precinct row, and the printed TOTALS together.
    wide_inputs, normal_pages, i = [], [], 0
    while i < len(pages):
        page, txt = pages[i]
        if county is not None and len(wide._office_sequence(txt)) >= 2:
            md, j = txt, i
            while not _has_totals(md) and j + 1 < len(pages) and j - i < 3:
                j += 1
                md += "\n" + pages[j][1]
            wide_inputs.append((page, md))
            i = j + 1
            continue
        if county is not None and wide.is_wide_page(txt):
            wide_inputs.append((page, txt))
        else:
            normal_pages.append((page, txt))
        i += 1

    blocks = []
    wide_best = {}  # (office_key, district, party) -> (contest, res)

    def _consider(c, res, store):
        """Keep the best (PASS-first, then most precincts) contest per key."""
        c["ncols"] = res["ncols"]
        office, district = normalize_office(c["office"])
        key = (office_match_key(office), str(district or ""), c.get("party"))
        cur = store.get(key)
        better = (cur is None
                  or (cur[1]["status"] != "PASS" and res["status"] == "PASS")
                  or (cur[1]["status"] == res["status"] and len(c["prec"]) > len(cur[0]["prec"])))
        if better:
            store[key] = (c, res)

    for page, md in wide_inputs:
        for party in ("REP", "DEM"):
            # Try both: header-anchored parse (needs precincts to sum to the
            # page TOTALS) and the headerless authority scan (locates each
            # contest by its total multiset). Both are checksum-gated below.
            found = list(wide.parse_wide_page(md, county, party, county_df)[0])
            found += wide.scan_wide_contests(md, county, party, county_df)[0]
            for c in found:
                res = checksum(c)
                if res is None:
                    continue
                c["ncols"] = res["ncols"]
                c["pages"] = [page]
                office, district = normalize_office(c["office"])
                key = (office_match_key(office), str(district or ""), party)
                cur = wide_best.get(key)
                better = (cur is None
                          or (cur[1]["status"] != "PASS" and res["status"] == "PASS")
                          or (cur[1]["status"] == res["status"]
                              and len(c["prec"]) > len(cur[0]["prec"])))
                if better:
                    wide_best[key] = (c, res)
    for page, txt in normal_pages:
        blocks.extend(parse_page(txt, page))

    # Two stitch strategies, unioned best-wins per contest key so neither can
    # regress the other: the base stitch(), and stitch_carry() which additionally
    # carries an office/party forward onto headerless continuation tables (a long
    # precinct matrix PaddleOCR splits into several tables, repeating the office
    # header only on the first — the base stitch drops the rest as office=None).
    norm_best = {}
    for source in (stitch(blocks), stitch_carry(blocks)):
        for c in source:
            res = checksum(c)
            if res is None:
                continue
            c["ncols"] = res["ncols"]
            if res["status"] == "FAIL" and res.get("dupes"):
                new_status, note = resolve_dupes(c, res)
                if new_status != res["status"]:
                    res["status"] = new_status
                    res["dedup_note"] = note
            _consider(c, res, norm_best)

    results = list(norm_best.values())
    results.extend(wide_best.values())
    return results


def contest_to_rows(contest, county, county_df):
    """Build OpenElections rows for one checksummed contest.

    Factored out of convert_canvass_pdfs.process()'s inner loop verbatim in
    behavior: resolve party (marker, ballot-measure blanking, or totals-based
    inference), name columns via the county-CSV totals join, and emit one row
    per (precinct, resolved-column) with a non-None vote.

    Returns (rows, party, notes). rows is a list of dicts in CSV_COLUMNS order.
    notes carries the same provenance strings process() logs to
    VERIFICATION_NEEDED (party inference, ballot-order tie assignment, unmatched
    totals) so the repair log can preserve them.
    """
    notes = []
    party = contest["party"]
    is_measure = bool(re.match(r"PROPOSED\b", contest["office"], re.I))
    if is_measure:
        party = None
    elif not party:
        party = infer_party(contest, county, county_df)
        if party:
            notes.append(f"{contest['office']}: no party marker, inferred {party} from vote totals")
        else:
            notes.append(f"{contest['office']}: no party marker, could not infer (dropped)")
            return [], None, notes

    names, (office, district), name_notes = name_columns(contest, county, party, county_df)
    for n in name_notes:
        notes.append(f"{contest['office']} [{party}] {n}")

    rows = []
    for code, pname, vals in contest["prec"]:
        for i, nm in enumerate(names):
            if i >= len(vals) or vals[i] is None:
                continue
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
    return rows, party, notes


# ---------------------------------------------------------------------------
# Contest-level merge into an existing county CSV
# ---------------------------------------------------------------------------

def _contest_key(office, district, party):
    """Identity used to line up a re-extracted contest with rows already in the
    county CSV: normalized office-match key + district + party. office_match_key
    absorbs the spelling/abbreviation drift between OCR runs so the same race
    matches even when one run misread a letter of the title."""
    return (office_match_key(office), str(district or ""), party or "")


def _existing_dupe_contest_keys(existing):
    """Contest keys (office/district/party) that still carry duplicated precinct
    rows in the current CSV — the residual dupes the authority pass couldn't
    reach (local races whose candidates aren't in the county-level file)."""
    if existing.empty:
        return set()
    ex = existing.copy()
    ex["_code"] = ex["precinct"].str.slice(0, 4)
    dup = ex[ex.duplicated(subset=["office", "district", "party", "_code", "candidate"], keep=False)]
    return {_contest_key(r["office"], r["district"], r["party"]) for _, r in dup.iterrows()}


def mismatched_contest_keys(csv_path, county, county_df):
    """Contest keys (office/district/party) that currently disagree with the
    authoritative county totals — the surgical target set for --only-mismatched."""
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    if not {"office", "district", "party", "candidate", "votes"}.issubset(df.columns):
        return set()
    df["votes"] = pd.to_numeric(df["votes"], errors="coerce")
    keys = set()
    for (office, district, party), grp in df.groupby(["office", "district", "party"], sort=False):
        auth = _authoritative_totals(county, office, district, party, county_df)
        if not auth:
            continue
        got = grp.dropna(subset=["votes"]).groupby("candidate")["votes"].sum().to_dict()
        if any(c in auth and int(v) != auth[c] for c, v in got.items()):
            keys.add(_contest_key(office, district, party))
    return keys


def merge_passing(county, results, county_df, out_dir=OUT_DIR, dry_run=False,
                  accept_status=("PASS",), dedup_only=False, require_existing_dupes=False,
                  restrict_keys=None):
    """Replace each newly-passing contest's rows in the county CSV.

    For every contest in `results` whose checksum status is in accept_status,
    build its rows and swap them for the existing rows of the same
    (office, district, party) in 2026/counties/<county>. Contests not present
    in `results` are left untouched, so a targeted re-run can only ever add
    verified data — never regress a contest it didn't touch.

    Returns (n_merged, log_lines).
    """
    out = os.path.join(out_dir, f"{ELECTION_PREFIX}__{county.lower().replace(' ', '_').replace('.', '')}__precinct.csv")
    existing = (pd.read_csv(out, dtype=str, keep_default_na=False)
                if os.path.exists(out) else pd.DataFrame(columns=CSV_COLUMNS))

    log = []
    dupe_keys = _existing_dupe_contest_keys(existing) if require_existing_dupes else None

    # Collect candidate merges keyed by contest identity first, so we can detect
    # when several DISTINCT contests collapsed onto one key (e.g. three "State
    # Republican Executive Committee, Baldwin County" places the OCR gave no
    # distinguishing district/place to). Writing all of them under one key would
    # re-create duplicate (precinct, Unverified Candidate N) rows — the very
    # thing we're removing — so those keys are refused, not merged.
    pending = {}  # key -> list of (office, district, party, res, frame, notes)
    for contest, res in results:
        if res["status"] not in accept_status:
            continue
        if require_existing_dupes:
            office, district = normalize_office(contest["office"])
            if _contest_key(office, district, contest.get("party")) not in dupe_keys:
                continue
        if dedup_only and not res.get("dedup_note"):
            continue
        rows, party, notes = contest_to_rows(contest, county, county_df)
        if not rows:
            log.append(f"  SKIP {contest['office']}: produced no rows ({'; '.join(notes) or 'party unresolved'})")
            continue
        office, district = rows[0]["office"], rows[0]["district"]
        key = _contest_key(office, district, party)
        # --only-mismatched: leave already-correct contests (and their resolved
        # names) untouched; merge a re-OCR'd PASS only where the CSV currently
        # disagrees with the authoritative totals.
        if restrict_keys is not None and key not in restrict_keys:
            continue
        pending.setdefault(key, []).append(
            (office, district, party, res, pd.DataFrame(rows, columns=CSV_COLUMNS), notes))

    new_frames = []
    drop_keys = set()
    merged = 0
    for key, items in pending.items():
        if len(items) > 1:
            office, district, party = items[0][:3]
            log.append(f"  REFUSE {office}" + (f" dist {district}" if district else "")
                       + f" [{party or ''}]: {len(items)} distinct contests share this "
                       "office key (OCR gave no place/district to tell them apart) — "
                       "not merged, needs contest separation")
            continue
        office, district, party, res, frame, notes = items[0]
        drop_keys.add(key)
        new_frames.append(frame)
        merged += 1
        log.append(f"  MERGE [{res['status']}] {office}"
                   + (f" dist {district}" if district else "")
                   + f" [{party or ''}]  {len(frame)} rows")
        if res.get("dedup_note"):
            log.append(f"        note: {res['dedup_note']}")
        for n in notes:
            log.append(f"        note: {n}")

    if not merged:
        return 0, log

    # Preserve the precinct LABEL the CSV already uses for each (party, code).
    # A repaired contest's rows carry the precinct name from its own source file,
    # but the same 4-digit code can be labeled differently across a county's REP
    # and DEM canvass files (Baldwin code 0029 is "LILLIAN COMM. CTR" in the REP
    # file, "PERDIDO BEACH VFD" in the DEM file), and the existing CSV already
    # settled on one label per code. Changing a precinct's name is out of scope
    # for a vote-data repair and would look like a spurious edit, so keep the
    # established label and only swap the numbers underneath it.
    if not existing.empty:
        ex = existing.copy()
        ex["_code"] = ex["precinct"].str.slice(0, 4)
        # Only meaningful where precincts carry 4-DIGIT codes (0001 LITTLE RIVER...).
        # Counties keyed by bare names (Mobile: "COLLIER ELEM") would cross-map
        # distinct precincts sharing a 4-char prefix ("ST. ELMO"/"ST. PAUL"), so
        # numeric-coded rows are the only ones mapped.
        ex = ex[ex["_code"].str.fullmatch(r"\d{4}")]
        label_map = (ex.groupby(["party", "_code"]).precinct
                     .agg(lambda s: s.value_counts().idxmax()).to_dict()) if len(ex) else {}
        for f in new_frames:
            codes = f["precinct"].str.slice(0, 4)
            f["precinct"] = [label_map.get((p, c), orig) if re.fullmatch(r"\d{4}", c) else orig
                             for p, c, orig in zip(f["party"], codes, f["precinct"])]

    if not existing.empty:
        exist_keys = existing.apply(
            lambda r: _contest_key(r["office"], r["district"], r["party"]), axis=1)
        kept = existing[~exist_keys.isin(drop_keys)]
    else:
        kept = existing

    combined = pd.concat([kept] + new_frames, ignore_index=True)
    combined = combined.sort_values(["precinct", "party", "office", "district", "candidate"])

    if dry_run:
        log.append(f"  (dry-run) would write {len(combined)} rows to {out} "
                   f"(was {len(existing)}; {len(existing) - len(kept)} rows replaced)")
        return merged, log

    os.makedirs(out_dir, exist_ok=True)
    combined.to_csv(out, index=False)
    log.append(f"  wrote {len(combined)} rows to {out} "
               f"({len(existing) - len(kept)} rows replaced by {sum(len(f) for f in new_frames)} new)")
    return merged, log


# ---------------------------------------------------------------------------
# OCR-free dedup gated by the authoritative county-level totals
# ---------------------------------------------------------------------------

def _authoritative_totals(county, office, district, party, county_df):
    """{candidate: county-wide votes} from 2026/..__county.csv for one contest."""
    key = office_match_key(office)
    pool = county_df[
        (county_df.county == county)
        & (county_df.party == party)
        & (county_df.office.map(office_match_key) == key)
    ]
    if district:
        pool = pool[pool.district.astype(str) == str(district)]
    return {r.candidate: int(r.votes) for _, r in pool.iterrows()}


def dedup_county_csv(csv_path, county, county_df, dry_run=False):
    """Collapse duplicated precinct rows in one county CSV, choosing sum vs
    keep-one per contest by which reproduces the authoritative county-level
    totals in 2026/..__county.csv — no OCR involved.

    Two physically different situations both show up as a repeated (precinct,
    candidate) key: a page-break split whose halves are *partial* counts (sum is
    right) and a row simply reprinted verbatim (keep one is right). Summing an
    exact duplicate would silently double real votes (Lamar Governor: authoritative
    60, summed 120), so we never guess: for each contest we compute both
    collapses and accept the reducer whose per-candidate county totals match the
    authoritative file exactly. If neither matches (or the contest's candidates
    aren't in the authoritative file), the contest is left untouched and
    reported, not force-collapsed.

    Returns (n_contests_fixed, n_rows_removed, log_lines).
    """
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    required = {"precinct", "office", "district", "party", "candidate", "votes"}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        return 0, 0, [f"  SKIP {os.path.basename(csv_path)}: non-conforming "
                      f"(missing {sorted(missing)})"]
    df["votes"] = df["votes"].astype(int)
    df["_code"] = df["precinct"].str.slice(0, 4)

    out_frames, log, fixed, removed = [], [], 0, 0
    for (office, district, party), grp in df.groupby(["office", "district", "party"], sort=False):
        dup_mask = grp.duplicated(subset=["_code", "candidate"], keep=False)
        if not dup_mask.any():
            out_frames.append(grp)
            continue

        auth = _authoritative_totals(county, office, district, party, county_df)

        def collapse(reducer):
            rows = []
            for (code, cand), sub in grp.groupby(["_code", "candidate"], sort=False):
                label = sub["precinct"].value_counts().idxmax()
                rows.append({**sub.iloc[0].to_dict(),
                             "precinct": label, "candidate": cand,
                             "votes": reducer(sub["votes"].tolist())})
            return pd.DataFrame(rows)

        def totals(frame):
            return frame.groupby("candidate")["votes"].sum().to_dict()

        variants = {"sum": collapse(sum),
                    "keep-one": collapse(lambda v: v[0] if len(set(v)) == 1 else None)}

        chosen = None
        for label, frame in variants.items():
            if frame["votes"].isna().any():
                continue  # keep-one only valid when the duplicate copies are equal
            t = totals(frame)
            checkable = [c for c in t if c in auth]
            if checkable and all(t[c] == auth[c] for c in checkable):
                chosen = (label, frame)
                break

        if chosen is None:
            out_frames.append(grp)
            log.append(f"  UNRESOLVED {office}"
                       + (f" d{district}" if district else "")
                       + f" [{party}]: {int(dup_mask.sum())} dup rows, "
                       f"neither sum nor keep-one matches authoritative totals "
                       f"(candidates in auth: {sum(1 for c in grp.candidate.unique() if c in auth)}/"
                       f"{grp.candidate.nunique()})")
            continue

        label, frame = chosen
        n_removed = len(grp) - len(frame)
        removed += n_removed
        fixed += 1
        out_frames.append(frame)
        log.append(f"  {label:8s} {office}" + (f" d{district}" if district else "")
                   + f" [{party}]: collapsed {n_removed} row(s)")

    if not fixed:
        return 0, 0, log

    combined = pd.concat(out_frames, ignore_index=True).drop(columns="_code")
    combined = combined.sort_values(["precinct", "party", "office", "district", "candidate"])
    if not dry_run:
        combined.to_csv(csv_path, index=False)
    log.append(f"  {'(dry-run) ' if dry_run else ''}{os.path.basename(csv_path)}: "
               f"{fixed} contest(s) collapsed, {removed} row(s) removed "
               f"({len(df)} -> {len(combined)})")
    return fixed, removed, log


# ---------------------------------------------------------------------------
# Backends / county resolution
# ---------------------------------------------------------------------------

def resolve_pdfs(target):
    """Accept a PDF, a glob, or a county folder; return its PDF paths."""
    if os.path.isdir(target):
        return sorted(glob.glob(os.path.join(target, "*.pdf")))
    if any(c in target for c in "*?["):
        return sorted(glob.glob(target))
    return [target]


def make_extract_fn(model, correction=False):
    """Resolve a backend by model name:

      * None / "nuextract" -> the repo's nuextract3 baseline (cache reuse).
      * "claude*"          -> Anthropic Claude via convert_canvass_pdfs_claude
                              (uses the stored `llm` anthropic key); the paid
                              escalation tier.
      * anything else      -> error (the ollama/hybrid/openrouter backends were
                              removed in the cleanup; see 2026/OCR_TOOLCHAIN.md).
    """
    if model in (None, "nuextract", "nuextract3"):
        return base.extract_pages
    if model == "paddleocr":
        # PaddleOCR-VL via the AI Studio cloud job API (whole-PDF submission).
        from convert_canvass_pdfs_paddleocr import make_extract_fn as paddle_extract
        return paddle_extract()
    if "claude" in model:
        # Routed through the `llm` library, which supplies the stored anthropic
        # key itself — no ANTHROPIC_API_KEY needed. Accepts bare "claude-*" or a
        # fully-qualified "anthropic/claude-*" model id.
        from convert_canvass_pdfs_claude import make_extract_fn as claude_extract
        return claude_extract(model, correction=correction)
    raise SystemExit(
        f"unknown model {model!r}: use 'paddleocr', 'nuextract', or a claude "
        f"model (e.g. 'anthropic/claude-sonnet-4-6'). The ollama/hybrid/openrouter "
        f"backends were removed; see 2026/OCR_TOOLCHAIN.md.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _load_county_df():
    df = pd.read_csv(COUNTY_CSV, dtype=str, keep_default_na=False)
    df["votes"] = df["votes"].astype(int)
    return df


def cmd_analyze(args):
    county_df = _load_county_df()
    extract_fn = make_extract_fn(args.model)
    for target in args.targets:
        pdfs = resolve_pdfs(target)
        if not pdfs:
            print(f"skip (no PDFs): {target}", file=sys.stderr)
            continue
        county = detect_county(pdfs[0])
        print(f"\n=== {county} ({os.path.basename(target)}) model={args.model or 'nuextract'} dpi={args.dpi} ===")
        results = analyze(pdfs, args.dpi, extract_fn, county_df, county=county)
        n_pass = n_fail = n_none = 0
        for c, res in results:
            st = res["status"]
            n_pass += st == "PASS"
            n_fail += st == "FAIL"
            n_none += st == "no-totals"
            if st != "PASS":
                lbl = f"{c['office'][:46]:46s} {c['party'] or '   '} pp{c.get('pages')}"
                extra = ""
                if st == "FAIL":
                    extra = f"\n        computed={res['computed']}\n        printed ={res['printed']}"
                    if res.get("dupes"):
                        extra += f"  dupes={res['dupes']}"
                print(f"  [{st:>9}] {lbl}{extra}")
        print(f"  -> {n_pass} PASS, {n_fail} FAIL, {n_none} no-totals")


def cmd_repair(args):
    county_df = _load_county_df()
    extract_fn = make_extract_fn(args.model, correction=getattr(args, "correction", False))
    for target in args.targets:
        pdfs = resolve_pdfs(target)
        if not pdfs:
            print(f"skip (no PDFs): {target}", file=sys.stderr)
            continue
        county = detect_county(pdfs[0])
        print(f"\n=== REPAIR {county} model={args.model or 'nuextract'} dpi={args.dpi} ===")
        restrict = None
        if args.only_mismatched:
            out = os.path.join(OUT_DIR, f"{ELECTION_PREFIX}__{county.lower().replace(' ', '_').replace('.', '')}__precinct.csv")
            restrict = mismatched_contest_keys(out, county, county_df) if os.path.exists(out) else set()
            print(f"  --only-mismatched: {len(restrict)} contest(s) currently disagree with authoritative totals")
        results = analyze(pdfs, args.dpi, extract_fn, county_df, county=county)
        accept = ("PASS",) if not args.accept_fail else ("PASS", "FAIL")
        n_merged, log = merge_passing(county, results, county_df,
                                      dry_run=args.dry_run, accept_status=accept,
                                      dedup_only=args.dedup_only,
                                      require_existing_dupes=args.fix_dupes,
                                      restrict_keys=restrict)
        for line in log:
            print(line)
        print(f"  -> {n_merged} contest(s) {'would be ' if args.dry_run else ''}merged")


def cross_check_county(csv_path, county, county_df):
    """Compare each candidate's precinct-sum in the CSV to the authoritative
    county-level total. Any mismatch is a real extraction error (dropped/added
    precinct, misread digit, unresolved dup) — found with no OCR at all, for
    every candidate the authoritative 2026/..__county.csv knows about.

    Returns (n_ok, n_mismatch, n_uncheckable, mismatch_lines).
    """
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    required = {"office", "district", "party", "candidate", "votes"}
    if not required.issubset(df.columns):
        return 0, 0, 0, [f"  non-conforming (missing {sorted(required - set(df.columns))})"]
    df["votes"] = pd.to_numeric(df["votes"], errors="coerce")

    n_ok = n_bad = n_unk = 0
    lines = []
    n_badvotes = int(df["votes"].isna().sum())
    if n_badvotes:
        lines.append(f"  WARNING: {n_badvotes} row(s) with non-integer votes (excluded from sums)")
    df = df.dropna(subset=["votes"])
    df["votes"] = df["votes"].astype(int)
    for (office, district, party), grp in df.groupby(["office", "district", "party"], sort=False):
        auth = _authoritative_totals(county, office, district, party, county_df)
        if not auth:
            n_unk += grp["candidate"].nunique()
            continue
        got = grp.groupby("candidate")["votes"].sum().to_dict()
        for cand, tot in got.items():
            if cand not in auth:
                n_unk += 1
                continue
            if tot == auth[cand]:
                n_ok += 1
            else:
                n_bad += 1
                lines.append(f"  MISMATCH {office}" + (f" d{district}" if district else "")
                             + f" [{party}] {cand}: CSV={tot} authoritative={auth[cand]} (Δ{tot-auth[cand]:+d})")
    return n_ok, n_bad, n_unk, lines


def cmd_crosscheck(args):
    county_df = _load_county_df()
    targets = args.targets or sorted(glob.glob(os.path.join(OUT_DIR, "*__precinct.csv")))
    T_ok = T_bad = T_unk = 0
    bad_counties = []
    for csv_path in targets:
        m = re.search(r"__primary__(.+?)__precinct\.csv$", os.path.basename(csv_path))
        if not m:
            continue
        county = _slug_to_county(m.group(1), county_df)
        ok, bad, unk, lines = cross_check_county(csv_path, county, county_df)
        T_ok += ok; T_bad += bad; T_unk += unk
        if bad or (lines and not ok):
            bad_counties.append(county)
            print(f"\n=== {county}: {ok} OK, {bad} MISMATCH, {unk} uncheckable ===")
            for line in lines[: args.limit]:
                print(line)
            if len(lines) > args.limit:
                print(f"  ... +{len(lines)-args.limit} more")
    print(f"\n==== authoritative cross-check: {T_ok} candidate-totals OK, "
          f"{T_bad} MISMATCH across {len(bad_counties)} counties, {T_unk} uncheckable "
          f"(candidate not in county-level file) ====")


def _slug_to_county(slug, county_df):
    """Map a filename slug (st_clair) to the county_df's spelling (St. Clair)."""
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
    want = norm(slug)
    for c in county_df.county.unique():
        if norm(c) == want:
            return c
    return slug.replace("_", " ").title()


def cmd_dedup(args):
    county_df = _load_county_df()
    targets = args.targets or sorted(glob.glob(os.path.join(OUT_DIR, "*__precinct.csv")))
    tot_fixed = tot_removed = 0
    for csv_path in targets:
        m = re.search(r"__primary__(.+?)__precinct\.csv$", os.path.basename(csv_path))
        if not m:
            print(f"skip (unrecognized name): {csv_path}", file=sys.stderr)
            continue
        county = _slug_to_county(m.group(1), county_df)
        fixed, removed, log = dedup_county_csv(csv_path, county, county_df, dry_run=args.dry_run)
        if fixed or log:
            print(f"\n=== dedup {county} ===")
            for line in log:
                print(line)
        tot_fixed += fixed
        tot_removed += removed
    print(f"\n==== {tot_fixed} contest(s) collapsed, {tot_removed} row(s) removed"
          f"{' (dry-run)' if args.dry_run else ''} ====")


def cmd_reconcile(args):
    import reconcile_two_read as rec
    county_df = _load_county_df()
    for pdf in args.pdfs:
        county = detect_county(pdf)
        n, log = rec.reconcile_county(pdf, county, county_df, dpi=args.dpi, dry_run=args.dry_run)
        for line in log:
            print(line)
        print(f"  -> {n} contest(s) reconciled\n")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    d = sub.add_parser("dedup", help="OCR-free: collapse duplicated precinct rows, "
                                     "gated by authoritative county totals")
    d.add_argument("targets", nargs="*", help="county CSV paths (default: all in 2026/counties)")
    d.add_argument("--dry-run", action="store_true")
    d.set_defaults(func=cmd_dedup)

    x = sub.add_parser("crosscheck", help="OCR-free: compare precinct-sums to authoritative "
                                          "county totals; report every mismatch")
    x.add_argument("targets", nargs="*", help="county CSV paths (default: all in 2026/counties)")
    x.add_argument("--limit", type=int, default=8, help="max mismatch lines per county")
    x.set_defaults(func=cmd_crosscheck)

    rc = sub.add_parser("reconcile", help="two-read reconciliation for digit-noise mismatches "
                                          "(second high-DPI sliced OCR + subset-sum to certified totals)")
    rc.add_argument("pdfs", nargs="+", help="county canvass PDF path(s)")
    rc.add_argument("--dpi", type=int, default=400)
    rc.add_argument("--dry-run", action="store_true")
    rc.set_defaults(func=cmd_reconcile)

    a = sub.add_parser("analyze", help="report per-contest checksum status (no writes)")
    a.add_argument("targets", nargs="+")
    a.add_argument("--model", default=None,
                   help="backend: 'paddleocr', 'nuextract' (default), or a claude model id")
    a.add_argument("--dpi", type=int, default=200)
    a.set_defaults(func=cmd_analyze)

    r = sub.add_parser("repair", help="re-extract failing contests and merge passing ones")
    r.add_argument("targets", nargs="+")
    r.add_argument("--model", default=None,
                   help="backend: 'paddleocr', 'nuextract' (default), or a claude model id")
    r.add_argument("--dpi", type=int, default=300)
    r.add_argument("--correction", action="store_true",
                   help="claude backend only: send cached nuextract markdown for correction")
    r.add_argument("--dry-run", action="store_true")
    r.add_argument("--dedup-only", action="store_true",
                   help="merge only contests fixed purely by collapsing duplicated "
                        "split rows (no OCR change); safest, no name churn")
    r.add_argument("--fix-dupes", action="store_true",
                   help="merge checksum-clean re-extractions ONLY for contests that "
                        "still have duplicate rows in the CSV (residual local races)")
    r.add_argument("--only-mismatched", action="store_true",
                   help="merge ONLY contests that currently disagree with authoritative "
                        "county totals; leaves correct contests (and their names) alone")
    r.add_argument("--accept-fail", action="store_true",
                   help="also merge FAIL contests (default: only checksum-PASS)")
    r.set_defaults(func=cmd_repair)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
