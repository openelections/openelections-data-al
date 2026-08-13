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

def build_second_read_pdf(pdf_path, county, dpi=400, overlap=0.12, split=False,
                          page_indices=None, max_bytes=2_000_000):
    """Render pages at `dpi` into a new PDF under SECOND_READ_DIR/<County>
    County/ (so convert_canvass_pdfs.detect_county resolves the right county).
    Cached by mtime-independent path; returns the PDF path.

    PaddleOCR rasterizes server-side, so `dpi` is not the OCR resolution — it is
    the resolution of the intermediate natural_pdf render that gets re-assembled
    into a PDF and re-rasterized by PaddleOCR. Independence from the first read
    comes from that double rasterization (natural_pdf render -> image -> PDF ->
    PaddleOCR), a different pixel path than reading the original PDF directly.

    By default pages are rendered WHOLE (split=False). The earlier
    split-into-overlapping-halves strategy enlarged glyphs but dropped the
    colspan office header from the bottom half of wide DISTRICT CANVASS pages
    and the office heading from the bottom half of NAME HEADING pages, so the
    second read came back missing precincts the first read had — and the A-vs-B
    disagreement scan (which matches precincts by name across the two reads) had
    no rows to disagree on. That is why Lt Gov showed "0 disagreements" and AG
    was absent from the second read entirely. Rendering whole keeps every page's
    header and all its precinct rows, so readB covers the same precincts as
    readA and the disagreement scan actually has cells to work on. Pass
    split=True to restore the half-page behavior for experimentation.

    `page_indices` (1-indexed) restricts the second read to just those pages —
    used for the targeted 2read so a 40+ page county PDF doesn't have to be
    uploaded in full (a 28 MB whole-page render timed out the 120s submit).
    Rendering only the pages that carry a mismatched office keeps the upload
    small AND lets those few pages be rendered at full 400 dpi for maximum
    independence. None means render every page.
    """
    from natural_pdf import PDF
    from PIL import Image

    subdir = os.path.join(SECOND_READ_DIR, f"{county} County")
    os.makedirs(subdir, exist_ok=True)
    base = f"{_norm(county).lower()}_2read"
    if page_indices:
        import hashlib
        h = hashlib.md5(",".join(map(str, sorted(page_indices))).encode()).hexdigest()[:8]
        base = f"{base}_{h}"
    out = os.path.join(subdir, f"{base}.pdf")
    if os.path.exists(out):
        return out

    pdf = PDF(pdf_path)
    wanted = set(page_indices) if page_indices else None
    render_pages = [pg for idx, pg in enumerate(pdf.pages, start=1)
                    if wanted is None or idx in wanted]
    if not render_pages:
        return None
    # A 400 dpi whole-page render is ~0.95 MB/page on these canvass pages, and
    # the PaddleOCR upload times out somewhere past ~2 MB (a 5.5 MB 6-page
    # render consistently aborted mid-write; a 1.95 MB 2-page render uploaded
    # cleanly). Scale the render dpi down so the assembled PDF stays under
    # max_bytes — independence comes from the re-rasterization, not from raw
    # dpi, so a somewhat lower dpi still yields a usefully independent read
    # while keeping the upload small enough to succeed. Floor at 150 dpi.
    n = len(render_pages) * (2 if split else 1)
    scale = (max_bytes / (n * 950_000)) ** 0.5
    render_dpi = max(150, min(dpi, int(400 * scale))) if n else dpi
    strips = []
    for pg in render_pages:
        img = pg.render(resolution=render_dpi).convert("RGB")
        if not split:
            strips.append(img)
            continue
        w, h = img.size
        ov = int(h * overlap)
        strips.append(img.crop((0, 0, w, h // 2 + ov)))
        strips.append(img.crop((0, h // 2 - ov, w, h)))
    strips[0].save(out, save_all=True, append_images=strips[1:], resolution=100.0)
    return out


# Keyword per normalized office key for locating the page(s) a contest is on in
# the readA OCR cache. Chosen to be distinctive (no office keyword is a substring
# of another office's header) so the targeted 2read picks up exactly the right
# pages.
_OFFICE_PAGE_KEYWORD = {
    "ATTORNEYGENERAL": "ATTORNEY",
    "LIEUTENANTGOVERNOR": "LIEUTENANT",
    "GOVERNOR": "GOVERNOR",
    "SECRETARYOFSTATE": "SECRETARY OF STATE",
    "STATETREASURER": "TREASURER",
    "STATEAUDITOR": "AUDITOR",
    "USSENATE": "UNITED STATES SENATOR",
    "USHOUSE": "REPRESENTATIVE",
    "COMMISSIONEROFAGRICULTUREANDINDUSTRIES": "AGRICULTURE",
    "PUBLICSERVICECOMMISSION": "PUBLIC SERVICE COMMISSION",
    "STATESUPERINTENDENT": "SUPERINTENDENT",
    "STATEBOARDOFEDUCATION": "STATE BOARD OF EDUCATION",
}


def _mismatch_page_indices(pdf_path, mism, readA):
    """1-indexed source-PDF page numbers whose readA OCR mentions any
    mismatched office that readA actually found (so the targeted 2read covers
    the contests the two-read can close). Offices absent from readA are skipped
    — the two-read matches precincts by name across readA and readB, so a
    contest with no readA entry can't be reconciled that way regardless. Returns
    None (-> render all pages) if the cache can't be found or no page matches."""
    import glob
    stem = re.sub(r"[^A-Za-z0-9]+", "_", os.path.splitext(os.path.basename(pdf_path))[0])
    cache = os.path.join(".canvass_cache_paddleocr", stem)
    if not os.path.isdir(cache):
        return None
    needles = []
    for k in mism:
        if readA.get(k) is None:
            continue  # two-read needs a readA entry; skip offices not in readA
        kw = _OFFICE_PAGE_KEYWORD.get(k[0])
        if kw:
            needles.append(kw.upper())
    if not needles:
        return None
    idx = []
    for md_path in sorted(glob.glob(os.path.join(cache, "p*.md"))):
        pg = int(re.search(r"p(\d+)\.md$", md_path).group(1))
        md = open(md_path).read().upper()
        if any(n in md for n in needles):
            idx.append(pg)
    return sorted(set(idx)) or None


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

def _col_sums(c):
    """Per-column sum of a contest's precinct rows (None cells excluded)."""
    prec = c.get("prec") or []
    ncol = max((len(v) for _, _, v in prec), default=0)
    return [sum(v[i] for _, _, v in prec if i < len(v) and v[i] is not None)
            for i in range(ncol)]


def _reada_closes(cA, county, county_df):
    """True if read-A's per-column sums (after spurious-column dropping) already
    equal the certified targets — i.e. the contest needs no second read, just a
    merge. Non-mutating (computes on a copy of the column sums), so it can be
    called up front to decide whether to build the 2read at all."""
    if cA is None or cA.get("totals") is None:
        return False
    office, district = normalize_office(cA["office"])
    certified = _certified_totals(county, office, district, cA.get("party"), county_df)
    if not certified:
        return False
    sums = list(cA["totals"])
    if len(sums) > len(certified):
        keep, unmatched = _match_columns_loose(sums, certified)
        if unmatched or len(keep) != len(certified):
            return False
        sums = [sums[i] for i in keep]
    targets = _column_targets(list(cA["totals"]), certified, sums_a=sums)
    if targets is None:
        return False
    return sums == targets


def _rekey_by_mismatch_party(store, mism, county, county_df):
    """Re-stamp each contest's party from the authoritative mismatch keys.

    mismatched_contest_keys returns (office_key, district, party) from the
    certified county file, so its party is authoritative. infer_party is only a
    guess from the OCR'd column sums, and it returns None when digit noise makes
    a sum not exactly hit a certified candidate total — which leaves the contest
    keyed under party=None and invisible to the REP mismatch lookup (it then
    SKIPs with a misleading "no read-A contest" instead of being reconciled).

    This re-stamps the ONE contest per (office, district) that belongs to the
    mismatch party with that party and re-keys the store so readA/readB.get(key)
    resolves. Contests not in the mismatch set keep their inferred party.

    Collision-aware: when two contests share an (office, district) but different
    parties (e.g. Tuscaloosa Lt Gov — a 9-col REP NAME-HEADING contest on
    p019/p020 and a 2-col DEM wide contest on p032), only the one that actually
    belongs to the mismatch party is re-stamped; the other keeps its inferred
    party. The earlier blanket re-stamp collapsed both onto the mismatch party
    and the DEM contest (fewer cols) overwrote the REP one in the dict, so the
    REP contest reported 'columns don't map' against the wrong certified set.
    The belonging contest is chosen by: (1) already inferred as the mismatch
    party; else (2) whose column sums map to that party's certified candidate
    totals (loose, spurious columns dropped); else (3) a None-party fallback.
    """
    from collections import defaultdict
    mism_od = {(k[0], k[1]): k[2] for k in mism}
    groups = defaultdict(list)
    for key, c in store.items():
        groups[(key[0], key[1])].append((key, c))
    out = {}
    for (ok, dist), items in groups.items():
        if (ok, dist) not in mism_od:
            for key, c in items:
                out[key] = c
            continue
        P = mism_od[(ok, dist)]
        chosen = None
        for key, c in items:  # (1) already inferred as P
            if key[2] == P:
                chosen = c
                break
        if chosen is None:  # (2) column sums map to P's certified totals
            office_name = items[0][1]["office"]
            o, dn = normalize_office(office_name)
            cert = _certified_totals(county, o, dn, P, county_df)
            for key, c in items:
                sums = c.get("totals") or []
                if not sums:
                    continue
                keep, unmatched = _match_columns_loose(list(sums), cert)
                if not unmatched and len(keep) == len(cert):
                    chosen = c
                    break
        if chosen is None:  # (3) None-party fallback (best-effort)
            for key, c in items:
                if key[2] is None:
                    chosen = c
                    break
        for key, c in items:
            if c is chosen:
                c["party"] = P
                out[(ok, dist, P)] = c
            else:
                out[key] = c  # keep under its own (inferred) party
    return out


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
        # DISTRICT CANVASS tables (Tuscaloosa) carry no printed CANDIDATE TOTALS
        # row — the totals live on a separate SUMMARY page. The per-column
        # precinct sums are the document's own claim, so stand them in for
        # totals: party inference and the column->candidate mapping both join on
        # these sums exactly as they join on a printed totals row.
        if c.get("totals") is None and c.get("prec"):
            sums = _col_sums(c)
            c["totals"] = sums
            c["ncols"] = len(sums)
        party = _party(c)
        c["party"] = party
        readA[(office_match_key(office), str(district or ""), party)] = c
    # Re-stamp parties from the authoritative mismatch keys so a contest whose
    # column sums don't EXACTLY hit a certified total (infer_party -> None) is
    # still reconciled against its REP mismatch key instead of silently dropped.
    readA = _rekey_by_mismatch_party(readA, mism, county, county_df)

    log = [f"{county}: {len(mism)} mismatched contest(s); building second read..."]
    # The second read only helps contests whose read-A does NOT already close to
    # certified (the A-vs-B disagreement path). If every mismatched contest that
    # read-A found already closes, skip the 2read entirely — no OCR upload, no
    # timeout. This is the common case when a clean re-OCR already matches the
    # certified totals and only needs the merge.
    need_second = any(readA.get(k) is not None and not _reada_closes(readA.get(k), county, county_df)
                      for k in mism)
    readB = {}
    if not need_second:
        log.append("  note: read-A already closes every mismatched contest; "
                   "skipping second read")
    else:
        # Targeted 2read: render only the pages carrying a mismatched office
        # that readA found, so the upload stays small (a whole-county 400 dpi
        # render can be 25+ MB and time out the submit) and those pages keep
        # full 400 dpi for maximum read independence.
        page_indices = _mismatch_page_indices(pdf_path, mism, readA)
        if page_indices:
            log.append(f"  note: targeted 2read of {len(page_indices)} page(s): "
                       f"{page_indices}")
        second = build_second_read_pdf(pdf_path, county, dpi=dpi,
                                       page_indices=page_indices)
        # If the 2read OCR upload fails — a transient service timeout on large
        # multi-page 2read PDFs is the common cause — don't crash the whole
        # reconcile: readB stays empty and the readA-already-certified fallback
        # below still merges the contests that don't need a confirming read.
        # Contests that genuinely need the second read are simply SKIPped this
        # run and can be retried once the service is back.
        try:
            for c, res in R.analyze([second], dpi, ef, county_df, county=county):
                office, district = normalize_office(c["office"])
                if c.get("totals") is None and c.get("prec"):
                    sums = _col_sums(c)
                    c["totals"] = sums
                    c["ncols"] = len(sums)
                party = _party(c)
                c["party"] = party
                readB.setdefault((office_match_key(office), str(district or ""), party), c)
        except Exception as e:
            log.append(f"  note: second read OCR failed ({type(e).__name__}); "
                       f"continuing with read-A only where it already closes")
        readB = _rekey_by_mismatch_party(readB, mism, county, county_df)

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
