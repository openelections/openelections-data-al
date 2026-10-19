# 2026 AL Primary — Discrepancy Summary

One-page orientation to every remaining problem in the 2026 primary data, organized by race. Full detail lives in
`VERIFICATION_NEEDED.md` (per-county narratives), `primary_spreadsheet_comparison_report.md` (county-by-county xls comparison),
and `primary_comparison_cells/<county>__cell_diffs.csv` (per-precinct diff tables). Nothing described here has been edited
into a CSV yet unless the entry says so.

## Sources

| Source | What it is | Role |
|---|---|---|
| `2026/counties/*__precinct.csv` | the deliverable, OCR/rebuilt from county canvass PDFs | what we publish |
| `2026/20260519__al__primary__county.csv` | certified county totals (per candidate) | what the crosscheck validates against |
| `2026_Primary_Election/*.xls` (new) | SoS machine-exported precinct spreadsheets, parsed by `parse_2026_primary_spreadsheets.py` | independent authoritative source (no OCR, not the preliminary canvass) |

The authoritative crosscheck is **sum-based** (precinct-sum vs certified county total) — so it cannot see per-precinct
cell errors inside a contest whose totals are right. The SoS xls comparison is **cell-based** and caught those.
Headline xls-vs-CSV numbers: 8 of 63 counties match cell-for-cell; 315 contest blocks exist only in the xls
(amendments, DEM U.S. Senate, DEM State House, SDEC/SREC, local committee races); ~2,100 cell diffs concentrated in
5 structural counties (DeKalb, Lawrence, Mobile, Lamar, Franklin) — DeKalb, Lawrence and Lamar's counts were key-format
phantoms since renamed away (2026-09-30), leaving Mobile (fixed 2026-09-30) and Franklin (fixed 2026-09-30) — Franklin's 4 residual
columns + 7 gap candidates + fire districts repaired from the xls; the 4 canvass-OCR-unrecoverable counties
(Calhoun, Choctaw, Randolph, Wilcox) now have complete parsed CSVs from the xls.

## How the OCR'd results compare to the state's xls, and where they can't

The two sources answer the same question from different documents. Our CSVs are read — mostly by PaddleOCR — from the
**county canvass PDFs**: scanned printouts of per-precinct results, printed by each county. The SoS xls is the state's
own machine export of the same election, per precinct. Nobody transcribed the xls; it has no OCR in it. So when the two
disagree, the question is almost never "did the OCR misread a number" and usually "do these two documents even claim the
same thing."

**On the contests the canvass actually prints, OCR matches the xls very closely.** Of the 63 counties with a CSV, 8 match
the xls cell-for-cell in every contest and precinct. Across all counties, about 4,300 candidate totals (precinct sums vs
certified) match at 98.8%. The cell diffs that do exist on shared contests cluster in five counties — Mobile, DeKalb,
Lawrence, Lamar, Franklin — and every one of those clusters turned out to be something other than an OCR misread:
key-format phantoms (DeKalb, Lawrence, Lamar), a column-sum reconciliation that shuffled cells while keeping totals
(Mobile), and canvass-grid errors the canvass's own printed totals contradict (Franklin). We fixed all five per-cell.
When an OCR value was genuinely wrong — single transposed cells in Hale, Walker, Cherokee, Jefferson, Fayette — the
crosscheck caught it, and re-reading the PDF fixed it. On shared ground, the OCR pipeline now reads the canvasses about
as faithfully as they can be read.

**OCR cannot match the xls in three ways, and no better OCR would change any of them:**

1. **The canvass doesn't print what the xls carries.** The xls has 315 contest blocks our CSVs lack — the two statewide
   amendments (absent from 23 and 19 counties respectively), DEM U.S. Senate and DEM State House in several counties,
   most SDEC/SREC detail, and the local committee and fire-district races. This isn't a reading failure; the county
   simply never printed those contests on the pages we scan. The xls is the only source for them.
2. **The canvass is the wrong document for late-counted votes.** Our canvass scans are preliminary election-night prints.
   The certified totals count later absentee and provisional ballots that the print doesn't distribute to precincts.
   Reading harder doesn't recover a vote the page never assigns. Elmore is the proof by exception: its REVISED canvass —
   a later run from the same county — moved each of those votes to its proper `PROVISIONAL` row, and the OCR read it
   fine. Tuscaloosa (28) and Perry (14) have no such revised print, so only the xls closes them.
3. **The documents themselves disagree.** Nine residual pairs are cases where the canvass and the xls agree with each
   other against the certified county total (Crenshaw Wess −60, Lauderdale AG −27, Tallapoosa Ensler −10, and the +1/−1
   singles), plus Elmore's one three-way. There the OCR is faithfully reporting a real conflict between the county's
   print and the state's certification, and we keep the canvass value and document the conflict rather than pick a winner.

**One more asymmetry cuts the other way: the xls is not an oracle.** In several places the xls carries its own defects —
DeKalb's ABS/early-vote distribution is scrambled in the xls but correct (canvass-verified) in our CSV, and in Franklin the
per-precinct cells of the canvass grid were wrong while both the canvass print and the xls agreed on the right values. The
xls reads cleanly, but it is one more document with opinions, not ground truth. The working rule that has held all year:
treat the canvass as the per-precinct authority, the certified county file as the totals authority, and use the xls to
fill what the canvass doesn't carry and to catch what neither check can see — cell-level errors inside a contest whose
totals happen to add up.

---

## The 52 residuals at a glance

Sum-based residual pairs (precinct CSV ≠ certified county total), now classified against the xls. *(2026-09-30: was 72 —
the three Mobile PSC1 `[25,9,11]` residuals closed when the wide-race fix rewrote those cells from the xls; one new two-source
residual surfaced in their place, Mobile LtGov REP Pate Δ−10. Later the same day, Franklin's 4 closed from the xls → 66.
2026-10-01: Elmore's 14 Δ−1 residuals closed from the OCR'd REVISED canvass (not the xls) — all 14 are exactly a
`0031 PROVISIONAL` cell the preliminary print left at 0 that the final amended run distributes — → 52.)*

- **41 are xls==certified — all except Tuscaloosa's 28 and Perry's 14 are now closed** (Mobile PSC1 3 and Franklin 4,
  both 2026-09-30; Elmore's 14 on 2026-10-01 via the county's REVISED canvass, which agrees with the xls-vs-certified
  verdict). The remaining 42 (Tuscaloosa 28 incl. the Hill name key + Perry 14) are closable per-precinct the same way.
- **9 are xls==canvass** (Crenshaw, Greene, Lauderdale, Limestone, St. Clair, Tallapoosa, Winston, plus the new Mobile
  LtGov REP Pate Δ−10: canvass text layer and the SoS xls both total 1,542 vs certified 1,552): the SoS export
  itself agrees with our CSV against the certified county total. Genuine source discrepancies — keep documented.
- **1 is a 3-way disagreement**: Elmore PSC1 Oden (xls 2,270 = revised canvass 2,270 > preliminary canvass 2,269 >
  certified 2,260) — the revised canvass confirms it is not another provisional cell.
- **1 is a name collision**: Tuscaloosa AgComm Hill (SoS spells "Corey Hill", CSV "Cory Hill"; SoS total 3,954 = certified).

Race-organized detail below; residual pairs are folded into each race's county bullets.

---

## By race

### U.S. Senate (14 residuals + Mobile's 369 cell diffs — the biggest single race)
- **Mobile (369 cells) — FIXED (2026-09-30):** all REP + DEM per-precinct cells now read from source (REP from the
  canvass text-layer plan, DEM from the SoS xls, which sums to certified here). See the Mobile county-row entry below.
- **Perry** — all 4 DEM candidates short by −2…−10 (late abs/prov); xls==certified, closable per-precinct.
- **Tuscaloosa** — 4 REP candidates short by −1…−4; xls==certified, closable.
- **Elmore** — Moore Δ−1 **closed 2026-10-01** (REVISED canvass: the 0031 PROVISIONAL cell). **Franklin** — Hudson Δ+1 — **closed 2026-09-30** from the xls (see
  the Franklin county row below). **Tallapoosa** — no Senate residual, but Ensler in LtGov below.
- Small +/-1 residuals where xls==canvass: **Crenshaw** Wess Δ−60, **Greene** Wess Δ+1, **Limestone** Deas Δ+1,
  **Winston** Wheeler Δ+1 — two independent sources (canvass + SoS export) agree against the certified total; keep.

### Lieutenant Governor (9 residuals + Mobile's 304 cell diffs)
- **Mobile (304 cells) — FIXED (2026-09-30)** — same fix as Senate (REP canvass text-layer, DEM xls). One residual
  remains: REP Rick Pate canvass+xls agree at 1,542 vs certified 1,552 (documented, kept canvass-faithful).
- **Perry** — Ensler −12, Perryman −9; xls==certified, closable. **Tuscaloosa** — 5 small REP deltas; xls==certified.
- **Elmore** — Allen Δ−1 **closed 2026-10-01** (REVISED canvass). **Tallapoosa** — Ensler Δ−10, but here xls==canvass (source discrepancy, keep).

### Governor (7 residuals + Mobile's 244 cell diffs)
- **Mobile (244 cells) — FIXED (2026-09-30)** — same fix as Senate; no residual (REP Gov totals match certified).
- **Perry** — 5 DEM deltas (Jones −20 etc.); xls==certified, closable. **Elmore** — Tuberville Δ−1 **closed 2026-10-01**.
- **Crenshaw** — Jones Δ+40, xls==canvass (source discrepancy, keep). **Lamar (57 cells)** — all missing rows,
  part of Lamar's legacy-format compensating block — FIXED 2026-09-30 by the key rename (no residual; see the
  county table below).
- **Jefferson** — Governor is closed; Jefferson's one remaining issue is elsewhere (amendments, below).

### Public Service Commission, Place 1 (9 residuals; both dual-key counties fixed 2026-09-30)
- **Lawrence** — its defect was a **synthetic parallel `1001–1031` unit block**: the csv builder double-listed the
  county's 31 unit names, re-coding the second block `10NN NAME`. 4 wide offices (AgComm, PSC1 REP, PSC1 DEM) and parts
  of 2 more (Sheriff units 0005–0031, Commission D3 all but 0013) lived entirely in that block. **FIXED (2026-09-30)** —
  326 rows re-keyed to the coded `00NN` forms; every cell equals the SoS xls per-cell and every sum equals certified.
- **Mobile — FIXED (2026-09-30):** all 80 units × 3 DEM candidates now carry the xls per-precinct values (which sum to
  certified exactly); the [25, 9, 11] candidate-total gaps are gone, and RISE CHURCH rows were inserted. Residual: none.
- **Perry** — 3 DEM deltas (−9/−7/−3), xls==certified, closable. **Tuscaloosa** — Oden −2, Gentry −7, closable
  (Gentry 8,404 = certified). **Elmore Oden +9 is the sole 3-way disagreement** (xls = revised canvass 2,270 > preliminary canvass 2,269 > certified 2,260) — keep.

### PSC Place 2 (5 residuals, all small)
- **Tuscaloosa** — 3 deltas, xls==certified, closable. **Franklin** — Beeker +3 — **closed 2026-09-30** from the xls.
- **Elmore** — Zeigler Δ−1 **closed 2026-10-01** (REVISED canvass).

### Commissioner of Agriculture and Industries (5 residuals; Lawrence's 91 missing cells fixed 2026-09-30)
- **Lawrence's 91 cells** — its AgComm block lives under the synthetic `10NN` keys; **FIXED (2026-09-30)** (see PSC1 above).
- **Tuscaloosa** — 3 deltas incl. the "Cory vs Corey Hill" name collision (SoS total 3,954 = certified); closable with
  the spelling noted. **Elmore** −1 **closed 2026-10-01**; **Franklin** McInnis −2 **closed 2026-09-30** from the xls.

### Attorney General (6 residuals; DeKalb's 135 missing cells fixed 2026-09-30)
- **DeKalb (135 missing cells)** — dual-key defect **FIXED (2026-09-30)**: 46 bare-name precinct keys renamed to their
  coded `NNNN NAME` forms; all 368 AG cells verified against the canvass per-cell.
- **Lawrence's AgComm/PSC1/Sheriff/Commission blocks sat in a synthetic `10NN` key block** — **FIXED (2026-09-30)**,
  326 rows re-keyed to `00NN`; the "missing cells" label was a phantom (data existed, under the wrong keys — all values
  already equal the SoS xls per-cell).
- **Lauderdale Δ−27 and St. Clair Δ+1** — xls==canvass, genuine discrepancies, keep. **Tuscaloosa** — 3 deltas
  (Mitchell 4,097, Robertson 5,940), xls==certified, closable. **Elmore** Δ−1 **closed 2026-10-01** (REVISED canvass).

### State Auditor (3 residuals; DeKalb's 90 missing cells fixed 2026-09-30) / State Treasurer / Secretary of State
- **DeKalb (90 missing cells)** — dual-key defect **FIXED (2026-09-30)**, same 46-key rename. **Tuscaloosa** (Sorrell/Chen −5 each, Lolley −3, Boozer −7,
  Dobson −5, Horn −4), **Elmore** (5 contests × Δ−1 — **closed 2026-10-01**, REVISED canvass), **Lamar** (6 SoS cells) — **fixed by the 2026-09-30 key
  rename** (the cells were present, under bare/legacy/minority-garbled keys); the remaining SoS-REP distribution
  difference in Lamar is one of its two documented key-preserving scrambles, see the anomalies entry.
- **Sumter** — SoS/Treasurer/Auditor REP present, but the county's **entire DEM side is missing** from the CSV
  while the xls has it complete (see "Party coverage" below).

### Statewide Amendments 1 & 2 (126 cell diffs + 23/19 counties missing the contest entirely)
- Missing entirely from **23** counties (Amendment 2) and **19** (Amendment 1); xls supplies Yes/No per precinct.
- **Jefferson** — the closed county's one remaining gap: its ABSENTEE/PROVISIONAL amendment rows are absent
  (~1,798 votes; xls Yes 91,735 vs CSV 89,937). All other amendment counties' blocks are simply absent, not partial.

### SDEC / SREC / State Board of Ed / DEM State House (the missing-contest block)
- **SDEC (Female)/(Male)**: entirely absent from **35/34 counties** — the xls has every one; biggest single
  completeness win available.
- **State Board of Education**: absent from **12 counties** (xls has it).
- **DEM U.S. Senate**: absent from 6 counties (incl. the 4 no-CSV ones); **DEM State House** from 7;
  **local executive-committee races** (county DEM/Republican executive committees, fire districts, sheriffs)
  make up most of the rest of the 315 xls-only blocks.

### U.S. House
- **Tuscaloosa** D4 deltas (−2/−1), xls==certified, closable; **Elmore** Dixon Δ−1 **closed 2026-10-01**; absent from 4 counties.

### Other residuals worth listing once
- Elmore's county/local Δ−1s: Circuit Court Judge Carter, District Court Judge Richardson, SREC Wright/Nick Carter —
  **all closed 2026-10-01** from the REVISED canvass. **Franklin County Executive Committee Pl 6 Lawler +21** — xls==certified (the canvass
  grid, not the print, was wrong); **closed 2026-09-30** from the xls (6 cells, net −21).

---

## By county (the structural items races alone don't show)

| County | Cell diffs | Diagnosis |
|---|---|---|
| **Mobile** | — | **Four wide races FIXED (2026-09-30).** Historical diagnosis: the 2026-08 column-sum-only reconciliation had shuffled distributions across precincts in Senate (369), LtGov (304), Governor (244) and DEM PSC1 (90 cells) while grand totals stayed equal, and all 7 wide blocks were missing the `RISE CHURCH` run with SEALS values hidden under the combined `'JAMES SEALS JR. PARK &amp; RISE CHURCH'` display name. Fix, applied per-precinct with per-cell validation: **REP wide** (Gov/LtGov/Senate, 80 units × 17) rebuilt from the canvass text-layer cell plan; **DEM wide + DEM PSC1** (80 units × 15) overwritten with the SoS xls cells, which were re-verified to sum to the certified Mobile county totals exactly for every one of these races. Totals: 586 REP + 323 DEM value edits, 32 combined→`JAMES SEALS JR. PARK &amp;` renames + 145 in-block display-name normalizations, 32 `RISE CHURCH` rows inserted (17 REP + 15 DEM; xls covers RISE CHURCH while the canvass scanned pages lack it). Post-fix validation: 0 per-cell mismatches, 80 runs per block × 7 blocks, all 4 CI tests + verifier.py PASS, sum-vs-certified diff across all Mobile candidates = exactly 1 — **REP LtGov Rick Pate 1,542 vs certified 1,552 (Δ−10)**, where canvass text and xls independently agree at 1,542 (two-source, documented, kept canvass-faithful per no-manufacture). Out-of-scope garbles remain in the untouched narrow blocks (`SUNLIGHT AUDITORUM` 17 rows, `MEADOWAKE ELEM.` 7 rows etc.) and at row 6803+ (DEM SDEC/State House combined-name + garbled displays). |
| **DeKalb** | 360 → 0 | **Dual-key defect FIXED (2026-09-30):** 46 contest blocks under bare-name precinct keys renamed to their coded `NNNN NAME` forms (368 rows); the 3 affected offices (AG, AgComm, State Auditor) verified cell-for-cell against the canvass, so the CSV is the authority and the SoS xls's scrambled distribution for these offices stands documented. |
| **Lawrence** | 313 → 0 | **Synthetic `10NN` key block FIXED (2026-09-30):** 31 unit names were double-listed, the second block re-coded `1001–1031`; 4 wide offices (AgComm, PSC1 REP + DEM) and parts of 2 more (Sheriff 0005–0031, Commission D3 minus 0013) lived there. 326 rows re-keyed to `00NN`; the 313 "cell diffs" were compare-script key phantoms — **every cell equals the SoS xls per-cell and every sum equals certified** (no value edits). |
| **Lamar** | 94 → 0 | **FIXED (2026-09-30):** 48 bare/legacy-`2009`–`2027`/minority-garbled keys renamed to the file's own coded `NNNN` forms (212 rows, no value edits; LF-only file). 2 sum-preserving distribution scrambles remain, kept canvass-faithful + documented (LtGov DEM Perryman 4-unit; SoS REP `CHU.`↔`VERNON WEST`). |
| **Franklin** | 74 → 0 | **FIXED (2026-09-30):** 77 value edits + 231 inserts from the SoS xls (whose every checkable candidate total == certified) — the 4 residual cells (McInnis, Hudson, Beeker, and Lawler ×6 cells net −21), the scramble-corrupted amendment A1/A2 columns (canvass-OCR: PROVISIONAL carried 70 Yes votes, two units swapped, one zeroed), the 7 "gap" candidates (PSC1 DEM Ramsey/Northrop, PSC2 Woodall/Zeigler, U.S. House Barnes, Exec Comm Pl 7 Davis/Nale), Amendment 2 Yes, 3 missing amendment cells, and the Belgreen + Frog Pond fire-district elections (20 rows, blank party). 1869 → 2100 rows; every certified candidate total now == the CSV sum. Remaining xls diffs are office-string **spelling variants only**. |
| **Elmore** | 15 → 1 | **FIXED (2026-10-01, REVISED-canvass OCR):** the 14 Δ−1 residuals were post-canvass-provisional — the county's final amended run (`Elmore County Canvass Report REVISED.pdf`, RUN DATE 05/27/26) distributes each as a `0031 PROVISIONAL` vote — 14 cells edited 0→1 with the revised blocks summing to certified on all 14 and matching the CSV per-cell on the other 30 units. Remaining: PSC1 Oden Δ+9, a genuine 3-way (canvass 2,269 / xls + revised 2,270 / certified 2,260), kept canvass-faithful. |
| Calhoun / Choctaw / Randolph / Wilcox | — (no CSV before) | **Superseded**: complete parsed CSVs now exist from the xls; the "OCR unrecoverable" narratives no longer block. |
| **Sumter** | — | Party coverage "REP only" superseded: xls has the full DEM side (Governor, LtGov, U.S. Senate, PSC1, Sheriff, Commission 1/4/5, District Court Judge) + amendments. |
| Jefferson | small | County is closed; only gap is the ~1,798 absentee/provisional amendment votes. |

## What this means as a punch list

1. **Closable from the xls with per-precinct provenance (~42 residuals)**: Tuscaloosa (all) and Perry (all) are what's left;
   Elmore's 14 were closed 2026-10-01 from the OCR'd REVISED canvass instead (the xls verdict xls==certified was correct,
   but the county's own final run read the cells directly — every one a `0031 PROVISIONAL` 0→1). Just import the xls cells.
   *(Franklin's 4 and Mobile PSC1's 3 were done 2026-09-30.)*
2. **Keep documented (10)**: the 9 two-source-agree discrepancies + Elmore Oden.
3. **New completeness work** all from one source: amendments (23 counties), SDEC (35), State Board of Ed (12),
   DEM U.S. Senate (6), DEM State House (7), Sumter DEM, local committee races.
4. **Cell-level data-quality fixes** (sum checks alone will never catch these): ~~the Mobile per-precinct scramble~~
   (done 2026-09-30, four wide races); the DeKalb and Lawrence key-block fixes (done 2026-09-30 — DeKalb 46 bare keys
   renamed to coded `NNNN NAME`, Lawrence's synthetic `1001–1031` block re-keyed to `00NN`); then Lamar's structural key block
   (done 2026-09-30) and ~~Franklin's residual cells + scrambled amendment columns~~ (done 2026-09-30, per-cell from the xls).
5. **Retire**: the Choctaw/Randolph/Wilcox/Calhoun "unrecoverable" entries, Sumter REP-only (done 2026-09-30 below/see xls parse),
   and Franklin's "canvass-internal" explanation (superseded 2026-09-30 — the xls fixed the 4 cells).