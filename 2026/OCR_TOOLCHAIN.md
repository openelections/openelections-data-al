# 2026 AL Primary — OCR repair toolchain

How the 2026 precinct CSVs are produced, verified, and repaired. Everything runs
with the project interpreter: `.venv/bin/python` (or `uv run python`).

## The load-bearing idea: verify against the certified county file

`2026/20260519__al__primary__county.csv` holds the **certified county-level
totals** (per office/district/party/candidate). It is the ground truth. Two gates
use it, and nothing is written that fails them:

1. **Checksum gate** — a contest's precinct rows must sum to its printed
   CANDIDATE TOTALS / TOTALS row (from the canvass image).
2. **Authoritative cross-check** — each candidate's summed precinct votes must
   equal the certified county total. This needs *no OCR* and is the primary
   scoreboard.

Candidate **names come from a totals-join**, never from OCR of the vertical
name-lattice: within a contest a candidate's county total uniquely identifies its
column.

## Tools (all in the repo root)

| File | What it does |
|---|---|
| `repair_canvass_contests.py` | The driver. Subcommands: `crosscheck`, `dedup`, `analyze`, `repair`, `reconcile`. |
| `convert_canvass_pdfs.py` | Base NAME-HEADING parser + checksum + name-join (pre-existing). |
| `convert_canvass_pdfs_paddleocr.py` | **Primary OCR backend** — PaddleOCR-VL via the AI Studio job API (whole-PDF, one job). Token in git-ignored `.paddleocr_token`. |
| `convert_canvass_pdfs_claude.py` | Claude backend via the `llm` library (uses the stored `anthropic` key; no env var). Escalation/fallback OCR. |
| `convert_wide_canvass.py` | Parser for the wide multi-office "precinct-report" format (Cherokee/Clay/Fayette/Mobile): authority-anchored column location + per-row shift alignment. |
| `reconcile_two_read.py` | Two-read reconciliation for clean digit-noise contests (second high-DPI sliced OCR + subset-sum to certified totals). |

`repair_canvass_contests.analyze()` auto-detects page format and routes: base
NAME-HEADING parser, `stitch_carry` (pools headerless continuation tables), the
wide parser, and the headerless authority scan — unioned best-status-per-contest
so no path regresses another.

## Common commands

```bash
# Scoreboard: every county's precinct-sums vs certified totals (no OCR)
.venv/bin/python repair_canvass_contests.py crosscheck
.venv/bin/python repair_canvass_contests.py crosscheck 2026/counties/<one>.csv --limit 50

# OCR-free dedup (sum vs keep-one chosen by certified totals)
.venv/bin/python repair_canvass_contests.py dedup [--dry-run]

# Repair a county with PaddleOCR; merge only checksum-PASS contests.
#   --only-mismatched : touch only contests that disagree with certified totals (safest)
#   --dry-run         : show what would merge
.venv/bin/python repair_canvass_contests.py repair "<path/to/County Canvas Report.pdf>" \
    --model paddleocr --only-mismatched [--dry-run]

# Same, but Claude as the OCR (needs the `llm` anthropic key)
.venv/bin/python repair_canvass_contests.py repair "<pdf>" --model anthropic/claude-sonnet-4-6 --only-mismatched

# Two-read reconciliation (clean digit noise only)
.venv/bin/python repair_canvass_contests.py reconcile "<pdf>" [--dry-run]

# Inspect one county's per-contest checksum status without writing
.venv/bin/python repair_canvass_contests.py analyze "<pdf>" --model paddleocr
```

Every `repair`/`reconcile` write is checksum-gated and (with `--only-mismatched`)
contest-scoped, so a run can only add verified data. Run them **sequentially**
(one process) — concurrent writers race on the CSVs.

## Session results (2026-08-12)

- **2,640 candidate-totals verified** against the certified file (from 2,378), **115
  mismatches**, **62 county files** (from 55).
- New fully/partly verified counties via PaddleOCR + the wide parser: Cherokee,
  Clay, Fayette, **Mobile** (13 contests, incl. Gov/LtGov/Senate recovered by
  two-read reconciliation — see `mobile_reconciliation_notes.md`), Conecuh, Greene,
  Crenshaw.
- Madison 40→25 via `stitch_carry` (recovers dropped continuation tables) + one
  reconciliation.

## Known residual & unhandled

- **~115 mismatches are heterogeneous structural errors**, not uniform digit noise:
  overcounts / merged candidate columns in older CSVs (e.g. Shelby Ag Commissioner
  ≈2×), shared misreads (both reads agree but wrong), contests the OCR won't parse
  cleanly. Each needs case-by-case diagnosis; no single tool clears them. Current
  list: `2026/authoritative_crosscheck.md`.
- **Bibb** — per-office pages with no totals row anywhere (no checksum anchor).
- **Randolph, Wilcox** — summary-only PDFs; no precinct data exists (skip).
- **Autauga** — 105-page text-layer EL30 "Precinct Report"; use
  `convert_precinct_pdfs.py` (no OCR needed), not this toolchain.
- Ops: 400-DPI second reads trip PIL's decompression-bomb warning; a very large
  slice set (Baldwin, 108 strips) hit a PaddleOCR API error — cap resolution or
  batch pages for big counties.

## Gotchas

- `src/verifier.py` is broken on Python 3.12 (`open(...,'rU')`); validate with the
  `crosscheck` subcommand or the external `openelections/openelections-data-tests`.
- `.canvass_cache*` dirs are keyed by PDF basename; reuse requires matching names.
