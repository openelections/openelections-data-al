# 2026 AL Primary — OCR Session Summary

## What we set out to do

OCR the 2026 AL primary canvasses accurately, given uneven results from
Ollama/nuextract. The reframe that made everything work: **stop trusting OCR,
verify against the certified county file** (`2026/20260519__al__primary__county.csv`).
Every contest is provably right or wrong, so accuracy became a gated pipeline
rather than a hope.

## Results

| | Start | End |
|---|---|---|
| Candidate-totals verified vs certified | 2,378 | **2,640** |
| Mismatches | 174 | **115** |
| County files | 55 | **62** |

New counties that had **zero** data, now verified: **Cherokee, Clay, Fayette,
Mobile** (all 13 contests, including Gov/LtGov/Senate), plus partial **Conecuh,
Greene, Crenshaw**. Madison went 40→25. Lamar/Lowndes fully de-duplicated. All
checksum-gated — no unverified data was written.

## What was built (all documented in `2026/OCR_TOOLCHAIN.md`)

- **`repair_canvass_contests.py`** — the driver: `crosscheck` (OCR-free
  scoreboard), `dedup` (authority-gated sum-vs-keep-one), `repair`, `analyze`,
  `reconcile`.
- **`convert_canvass_pdfs_paddleocr.py`** — PaddleOCR-VL via the AI Studio API,
  the primary OCR (reference-quality digits, one job per PDF).
- **`convert_canvass_pdfs_claude.py`** — Claude via the `llm` key,
  escalation/fallback.
- **`convert_wide_canvass.py`** — the wide multi-office format (authority-anchored
  column location + per-row drift alignment).
- **`reconcile_two_read.py`** — two-read reconciliation for clean digit noise.
- Plus `stitch_carry` (recovers dropped continuation tables) folded into the
  driver.

## The honest bottom line on the residual

The last 115 mismatches are **not** a single fixable class — they're heterogeneous
structural errors (merged/duplicated candidate columns in older CSVs, shared
misreads two reads can't break, contests the OCR won't parse). The general tools
were built and proved; they correctly *refuse* these rather than guess. Clearing
them is case-by-case work, not another broad tool. The full list with deltas is in
`2026/authoritative_crosscheck.md`, and provenance/caveat notes are in
`2026/mobile_reconciliation_notes.md`.

## Loose ends worth knowing

- **Autauga** is a text-layer EL30 report → belongs to `convert_precinct_pdfs.py`,
  no OCR needed (quick win if you want it).
- **Bibb** (no totals row) and **Randolph/Wilcox** (summary-only, no precinct data)
  are genuinely unhandled/out-of-scope.
- Nothing is committed — all changes are in the working tree for review.
  `src/verifier.py` is broken on Python 3.12; use `crosscheck` or the external
  test repo.
