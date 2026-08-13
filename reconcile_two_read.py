#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Two-read reconciliation for canvass contests that fail the authoritative
cross-check by *digit noise* (not structure).

PaddleOCR-VL is deterministic: re-OCRing the same page image reproduces the same
~1% of misread digit cells, so a plain re-run never surfaces them. But rendering
the page differently — higher DPI, sliced into overlapping halves so glyphs are
larger and tables shorter — yields a genuinely independent read that errs on
*different* cells. Where the two reads disagree is exactly the uncertain set, and
the certified county totals say which combination is right:

  for each candidate column, deficit = certified_total - sum(read_A column);
  find the subset of that column's A/B disagreements whose value changes sum to
  the deficit, and apply those B readings.

Every accepted value is one of the two actual OCR readings — nothing is
interpolated — and a contest is only written when *every* column closes to the
certified total exactly (the same checksum-or-nothing rule as the rest of the
pipeline). First proven by hand on Mobile Governor/Lt Gov/US Senate; this
generalizes it.

Usage:
    python reconcile_two_read.py <county-pdf> [--dpi 400] [--dry-run]
    # or via the repair driver:  repair_canvass_contests.py reconcile <pdf>
"""

import itertools
import os
import re

import pandas as pd

import repair_canvass_contests as R
from convert_canvass_pdfs import checksum, infer_party
from convert_precinct_pdfs import normalize_office, office_match_key

SECOND_READ_DIR = os.environ.get(
    "RECONCILE_DIR",
    "/private/tmp/claude-501/-Users-dwillis-code-openelections-data-al/"
    "b0dbc1f2-9aa5-4156-b79a-ff80867441d8/scratchpad/second_reads")

_norm = lambda s: re.sub(r"[^A-Z0-9]", "", str(s).upper())


# ---------------------------------------------------------------------------
# Second read: re-render at high DPI, sliced into overlapping halves
# ---------------------------------------------------------------------------

def build_second_read_pdf(pdf_path, county, dpi=400, overlap=0.12):
    """Render every page at `dpi`, split each into overlapping top/bottom halves,
    and assemble a new PDF under SECOND_READ_DIR/<County> County/ (so
    convert_canvass_pdfs.detect_county resolves the right county). Cached by
    mtime-independent path; returns the PDF path."""
    from natural_pdf import PDF
    from PIL import Image

    subdir = os.path.join(SECOND_READ_DIR, f"{county} County")
    os.makedirs(subdir, exist_ok=True)
    out = os.path.join(subdir, f"{_norm(county).lower()}_2read.pdf")
    if os.path.exists(out):
        return out

    pdf = PDF(pdf_path)
    strips = []
    for pg in pdf.pages:
        img = pg.render(resolution=dpi).convert("RGB")
        w, h = img.size
        ov = int(h * overlap)
        strips.append(img.crop((0, 0, w, h // 2 + ov)))
        strips.append(img.crop((0, h // 2 - ov, w, h)))
    strips[0].save(out, save_all=True, append_images=strips[1:], resolution=100.0)
    return out


# ---------------------------------------------------------------------------
# Column target: printed totals -> certified candidate totals
# ---------------------------------------------------------------------------

def _certified_totals(county, office, district, party, county_df):
    key = office_match_key(office)
    pool = county_df[(county_df.county == county) & (county_df.party == party)
                     & (county_df.office.map(office_match_key) == key)]
    if district:
        pool = pool[pool.district.astype(str) == str(district)]
    return sorted(int(v) for v in pool.votes)


def _map_columns(hint, certified, tol_frac=0.02, tol_min=3):
    """Map each column (a per-column vote figure: printed total or read-A sum) to
    the certified value it most likely is — exact match first, then the unique
    nearest unused certified value within tolerance. None if any column is
    ambiguous."""
    if hint is None or len(hint) != len(certified):
        return None
    remaining = list(certified)
    targets = [None] * len(hint)
    for i, p in enumerate(hint):
        if p in remaining:
            targets[i] = p
            remaining.remove(p)
    for i, p in enumerate(hint):
        if targets[i] is not None or p is None:
            continue
        cands = [v for v in remaining if abs(v - p) <= max(tol_min, int(tol_frac * max(v, 1)))]
        if len(set(cands)) == 1:
            targets[i] = cands[0]
            remaining.remove(cands[0])
        else:
            return None
    return targets if all(t is not None for t in targets) else None


def _match_columns_loose(hint, certified, tol_frac=0.05, tol_min=8):
    """Like _map_columns but allows `hint` to have MORE columns than `certified`
    — the extras are spurious OCR columns (a colspan artifact with a tiny sum and
    a blank printed total, common on Madison's continuation pages). Returns
    (keep_indices, unmatched_certified): keep_indices are the hint columns that
    map to a certified value, in original order; unmatched_certified is the list
    of certified values no column matched (non-empty means we can't reconcile)."""
    remaining = list(certified)
    match = {}
    for i, p in enumerate(hint):
        if p is not None and p in remaining:
            match[i] = p
            remaining.remove(p)
    for i, p in enumerate(hint):
        if i in match or p is None:
            continue
        cands = [v for v in remaining if abs(v - p) <= max(tol_min, int(tol_frac * max(v, 1)))]
        if len(set(cands)) == 1:
            match[i] = cands[0]
            remaining.remove(cands[0])
    return sorted(match), remaining


def _column_targets(printed, certified, sums_a=None):
    """Target certified value per column. Try the printed TOTALS row first (it's
    the document's own claim), then fall back to read-A's per-column sums with a
    looser tolerance — the printed row is a single OCR line and is sometimes
    itself misread or a different width, whereas the column sums come from 80+
    precinct rows and are only off by the same small digit noise we're fixing."""
    t = _map_columns(printed, certified)
    if t is not None:
        return t
    if sums_a is not None:
        return _map_columns(sums_a, certified, tol_frac=0.05, tol_min=8)
    return None


# ---------------------------------------------------------------------------
# Reconcile one contest
# ---------------------------------------------------------------------------

def _subset_delta(disagreements, deficit, max_r=4):
    """Smallest set of (idx, delta) whose deltas sum to deficit; None if none."""
    if deficit == 0:
        return []
    for r in range(1, min(max_r, len(disagreements)) + 1):
        for combo in itertools.combinations(disagreements, r):
            if len({i for i, _ in combo}) == r and sum(d for _, d in combo) == deficit:
                return combo
    return None


def reconcile_contest(cA, cB, targets):
    """Return (reconciled_prec_rows, notes) if every column closes to its target
    using A-vs-B disagreements, else (None, notes). cA/cB are contest dicts with
    prec rows [(name, _, vals)] in the same physical column order."""
    ncol = len(targets)
    A = {}
    for name, _, vals in cA["prec"]:
        if len(vals) >= ncol:
            A.setdefault(_norm(name), (name, list(vals[:ncol])))
    B = {}
    for name, _, vals in cB["prec"]:
        if len(vals) >= ncol:
            B.setdefault(_norm(name), list(vals[:ncol]))

    notes, applied = [], 0
    grid = {k: [(x if x is not None else 0) for x in v[1]] for k, v in A.items()}
    for j in range(ncol):
        deficit = targets[j] - sum(grid[k][j] for k in grid)
        if deficit == 0:
            continue
        disagreements = [(k, (B[k][j] if B[k][j] is not None else 0) - grid[k][j])
                         for k in grid
                         if k in B and B[k][j] is not None and B[k][j] != grid[k][j]]
        hit = _subset_delta(disagreements, deficit)
        if hit is None:
            notes.append(f"col{j}: deficit {deficit} not reconcilable "
                         f"({len(disagreements)} disagreements)")
            return None, notes
        for k, d in hit:
            grid[k][j] += d
            applied += 1
    rows = [(A[k][0], "", grid[k]) for k in grid]
    notes.append(f"reconciled: {applied} cell substitution(s) across {ncol} columns, "
                 f"{len(grid)} precincts")
    return rows, notes


# ---------------------------------------------------------------------------
# County driver
# ---------------------------------------------------------------------------

def reconcile_county(pdf_path, county, county_df, dpi=400, dry_run=False):
    ef = R.make_extract_fn("paddleocr")
    csv = os.path.join(R.OUT_DIR, f"{R.ELECTION_PREFIX}__"
                       f"{county.lower().replace(' ', '_').replace('.', '')}__precinct.csv")
    mism = R.mismatched_contest_keys(csv, county, county_df) if os.path.exists(csv) else set()
    if not mism:
        return 0, [f"{county}: no authoritative mismatches to reconcile"]

    def _party(c):
        return c.get("party") or (infer_party(c, county, county_df)
                                   if c.get("totals") is not None else None)

    readA = {}
    for c, res in R.analyze([pdf_path], dpi, ef, county_df, county=county):
        office, district = normalize_office(c["office"])
        party = _party(c)
        c["party"] = party
        readA[(office_match_key(office), str(district or ""), party)] = c

    log = [f"{county}: {len(mism)} mismatched contest(s); building second read..."]
    second = build_second_read_pdf(pdf_path, county, dpi=dpi)
    readB = {}
    for c, res in R.analyze([second], dpi, ef, county_df, county=county):
        office, district = normalize_office(c["office"])
        party = _party(c)
        c["party"] = party
        readB.setdefault((office_match_key(office), str(district or ""), party), c)

    to_merge = []
    for key in mism:
        cA = readA.get(key)
        if cA is None or cA.get("totals") is None:
            log.append(f"  SKIP {key}: no read-A contest/totals")
            continue
        office, district = normalize_office(cA["office"])
        certified = _certified_totals(county, office, district, cA.get("party"), county_df)
        ncol_a = len(cA["totals"]) if cA.get("totals") else 0
        sums_a = ([sum(v[i] for _, _, v in cA["prec"] if i < len(v) and v[i] is not None)
                   for i in range(ncol_a)] if ncol_a else None)
        # Spurious-column handling: read-A sometimes has MORE columns than the
        # certified file has candidates (a colspan artifact on continuation pages
        # leaves a phantom column with a tiny sum and a blank printed total). The
        # certified file is the authority on candidate count, so drop the columns
        # that don't map to any certified value before reconciling.
        dropped = []
        if sums_a is not None and len(sums_a) > len(certified):
            keep, unmatched = _match_columns_loose(sums_a, certified)
            if unmatched or len(keep) != len(certified):
                log.append(f"  SKIP {office} [{cA.get('party')}]: columns don't map to "
                           f"certified totals ({len(sums_a)} cols, {len(certified)} certified)")
                continue
            dropped = [i for i in range(len(sums_a)) if i not in keep]
            cA["prec"] = [(code, nm, [v[i] if i < len(v) else None for i in keep])
                          for code, nm, v in cA["prec"]]
            cA["totals"] = [cA["totals"][i] if i < len(cA["totals"]) else None for i in keep]
            sums_a = [sums_a[i] for i in keep]
            ncol_a = len(keep)
        targets = _column_targets(cA["totals"], certified, sums_a=sums_a)
        if targets is None:
            log.append(f"  SKIP {office} [{cA.get('party')}]: columns don't map to certified totals")
            continue
        cB = readB.get(key)
        if cB is None:
            # No second read for this contest — but if read-A's own per-column sums
            # already close to the certified targets (deficit 0 everywhere), the
            # data is correct as-is and the merge is the same 0-substitution case
            # the two-read path produces when readB happens to agree. Only fall
            # back to this when readA is already exact; otherwise the absence of
            # a confirming readB is a real gap and we skip.
            if sums_a is not None and len(sums_a) == len(targets) \
                    and all(sums_a[j] == targets[j] for j in range(len(targets))):
                rows = [(nm, "", list(vals[:len(targets)])) for _, nm, vals in cA["prec"]
                        if len(vals) >= len(targets)]
                notes = [f"reconciled: 0 cell substitution(s) across {len(targets)} columns, "
                         f"{len(rows)} precincts (read-A already certified; no second read)"]
            else:
                log.append(f"  SKIP {office} [{cA.get('party')}]: contest absent from second read")
                continue
        else:
            if dropped:
                cB["prec"] = [(code, nm, [v[i] if i < len(v) else None for i in keep])
                              for code, nm, v in cB["prec"]]
            if dropped:
                log.append(f"  note: {office} [{cA.get('party')}]: dropped spurious column(s) "
                           f"{dropped} (no matching certified candidate)")
            rows, notes = reconcile_contest(cA, cB, targets)
        if rows is None:
            log.append(f"  FAIL {office} [{cA.get('party')}]: {notes[-1]}")
            continue
        rec = {"office": cA["office"], "party": cA.get("party"), "district": district,
               "prec": rows, "totals": targets, "pages": cA.get("pages", [])}
        res = checksum(rec)
        rec["ncols"] = res["ncols"]
        if res["status"] != "PASS":
            log.append(f"  FAIL {office} [{cA.get('party')}]: post-reconcile checksum {res['status']}")
            continue
        log.append(f"  OK   {office} [{cA.get('party')}]: {notes[-1]}")
        to_merge.append((rec, res))

    if to_merge and not dry_run:
        n, mlog = R.merge_passing(county, to_merge, county_df)
        log += mlog
    elif to_merge:
        log.append(f"  (dry-run) would merge {len(to_merge)} reconciled contest(s)")
    return len(to_merge), log


def main():
    import argparse
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("--dpi", type=int, default=400)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    from convert_canvass_pdfs import detect_county
    county_df = R._load_county_df()
    for pdf in args.pdfs:
        county = detect_county(pdf)
        n, log = reconcile_county(pdf, county, county_df, dpi=args.dpi, dry_run=args.dry_run)
        for line in log:
            print(line)
        print(f"  -> {n} contest(s) reconciled\n")


if __name__ == "__main__":
    main()
