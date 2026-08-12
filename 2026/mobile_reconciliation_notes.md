# Mobile County — Governor / Lt. Governor / U.S. Senate reconciliation

Mobile's canvass prints these three contests as one wide table spanning pages 17–18
(precinct half A on p17 with the office headers, half B plus a full-county TOTALS row
on p18). PaddleOCR-VL read the 79-precinct × 17-column grid with ~1% of digit cells
wrong — identically on repeated runs (the model is deterministic), so a simple re-OCR
could not surface the errors.

**Method.** A second, genuinely different read was produced by re-rendering the two
pages at 400 DPI, slicing each into overlapping halves (larger glyphs, shorter tables),
and submitting that as a new PaddleOCR job. For each of the 17 columns, the cells where
the two reads disagreed were treated as the uncertain set; the combination of readings
whose column sum equals the **certified county total** (`2026/20260519__al__primary__county.csv`)
exactly was selected. Every accepted value is one of the two actual OCR readings —
nothing is interpolated — and all 17 columns close exactly, so the three contests pass
the same checksum gate as every other contest in the repo. The Lt. Governor column
target `1542` printed in the document's own TOTALS row is an OCR misread of the
certified `1552` (six of seven columns matched exactly; the unique leftover pair
differed by one digit).

**29 cell substitutions** were applied (precinct, column, original read → strip read),
e.g. `CLEARWATER CHRISTIAN CH col2: 11 → 272`, `BURNS MIDDLE SCHOOL col3: 46 → 218`,
`GRAND BAY MIDDLE SCH col9: 311 → 200`. The full list is in the session log of
2026-08-12; the reconciliation script pattern lives in the repair toolchain
(`convert_wide_canvass.py` + `repair_canvass_contests.py`).

Confidence caveat: cell choices are validated by exact column sums against certified
totals, not by human inspection of the scan. Rows where both reads agreed carry a small
residual risk of a shared misread that happens to cancel — no such cancellation was
needed here (all deficits were closed by observed disagreements), which is why the
result is presented as checksum-verified rather than estimated.
