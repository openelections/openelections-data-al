# 2026 AL Primary — Rows Needing Verification

Generated from 55 processed county CSVs in `2026/counties` (the `20260519__al__primary__democratic__precinct.csv` file is not a county and is excluded — see anomalies below). Three sections follow: **party coverage** (does this county's CSV contain both parties' results or only one), **failed-checksum contests** (real data written, but the extracted precinct rows don't sum to the document's own printed total), and **unresolved candidate columns** (real vote counts written under a placeholder name because no candidate in the county CSV matched that column's total). Neither of the last two categories was dropped from the CSVs — this doc is a punch list for reviewing them against the source PDFs, not a list of missing data.

The party-coverage and unresolved-candidate tables below were regenerated directly from the current CSVs in `2026/counties`; the failed-checksum section is carried forward from the conversion run logs (it cannot be reconstructed from the CSVs alone). Since the last revision, Calhoun, Elmore, Franklin, Hale, Jefferson, Macon, Montgomery, Sumter, Talladega, Tallapoosa, Tuscaloosa, Walker were added, Morgan was removed (no file), and several previously split "merged-column" rows (Marshall, St. Clair, Etowah, Clarke) were consolidated in the CSVs into single per-candidate totals; Shelby resolved five former placeholders. St. Clair was regenerated from the Canvas Report PDF with all Republican precinct results.

## Data-quality anomalies in the CSVs themselves

- **Democratic (file):** `20260519__al__primary__democratic__precinct.csv` is not a county file — its `county` column reads `Democratic` for every row (a statewide Democratic extract referencing precincts in multiple counties). Excluded from the per-county tables below.
- **Franklin:** non-conforming CSV — columns are ['county', 'precinct', 'office', 'votes'] (expected the 7-column OpenElections format); no party/candidate data, so it is excluded from the unresolved-candidate table and its party coverage is unknown.
- **Montgomery:** contains a stray non-DEM/REP party value ['party'] (an embedded duplicate header row mid-file — should be removed).
- **St. Clair:** regenerated from Canvas Report PDF (Republican results only); no Democratic section exists in the source document.
- **Madison:** 6 candidate-totals still mismatch the certified county file after OCR repair (down from 25). Two Public Service Commission contests remain: **PSC Place 1 REP** (Oden Δ-3552, Gentry Δ-8047) is structural column drift — Gentry votes bled into the Oden column across continuation pages; the two-read reconciliation's 10 disagreements don't sum to the 670 deficit, so it needs per-row column realignment, not digit-noise repair. **PSC Place 2 REP** (Andrews Δ-90, Woodall Δ-128, Beeker Δ-126, Zeigler Δ-227) is already 97-98% correct in the CSV; the CANVAS read is broken on Zeigler (4981 vs 11995) and the Democratic-PDF has correct printed totals but only 44 of 81 precincts recovered (stitch failure), so improving it needs per-column sourcing or a DEM-PDF stitch fix. Every other Madison contest is authority-verified.
- **Tuscaloosa — canvass-vs-certified source discrepancy (2026-08-13):** all 10 REP statewide contests in `20260519__al__primary__tuscaloosa__precinct.csv` were merged from the "Blue Sheet and Canvas Report" PDF with real candidate names and checksum-PASS parses. The 7 wide DISTRICT CANVASS contests (Attorney General, Secretary of State, State Treasurer, State Auditor, Commissioner of Agriculture and Industries, PSC Place 1, PSC Place 2) and the 2 NAME HEADING contests (Lieutenant Governor, U.S. Senate) each cover all 55 precincts. U.S. House 4 covers 33 precincts — Tuscaloosa is split between AL-04 (33 precincts, ~83k registered) and AL-07 (the other 22 precincts, ~63k registered); the canvass's REP U.S. House page itemizes only the AL-04 precincts, and the certified county file likewise records only the AL-04 REP race (Aderholt/Barnes, no district 7), so 33 precincts is the complete AL-04 set, not a shortfall. However, the canvass PDF prints candidate totals ~10 votes SHORT of the certified county file (`20260519__al__primary__county.csv`) for every one of these 10 contests — e.g. AG `[2145,4094,5935]` vs certified `[2147,4097,5940]`; SoS `[8090,2438,1068]` vs `[8095,2442,1068]`; Treasurer `[8802,3135]` vs `[8809,3138]`; U.S. House `[5580,1151]` vs `[5582,1152]`. Five independent OCR reads (whole-page 200/400 DPI, crop+upscale, + totals-row reads) all reproduce the SAME short digits for the wide contests — the digits are unambiguous, so this is NOT an OCR error fixable by re-OCR. The consistent ~10-vote deficit across all contests is the signature of ~10 late absentee/provisional ballots present in the certified county total but NOT distributed to precincts in the preliminary canvass PDF. These 10 contests are therefore canvass-faithful but will show a small residual mismatch in the authoritative crosscheck (`repair_canvass_contests.py mismatched_contest_keys` — note the county name is case-sensitive, `"Tuscaloosa"`) until a certified precinct-level source (SoS website / corrected canvass / county election office) is found and used to close the ~10-vote gap. **Governor** was reconciled to certified exactly (0 deficit) via two-read reconciliation — its digits were ambiguous across renders, unlike the 10 above; the committed per-precinct values are readB's readings (the crosscheck is sum-based, so per-precinct accuracy is unverified). (Resolved: U.S. House district now renders as the integer `4` in the CSV, not `4.0`.)
- **Morgan — Governor REP Tuberville column dropped (2026-08-13):** the original conversion of `20260519__al__primary__morgan__precinct.csv` kept Thomas Tuberville's per-precinct votes for only precincts 0001–0005 (564 votes); for precincts 0006–0041 his vote total had been shoved into the `candidate` field with `votes=0` (36 garbage rows whose "candidate" was a bare number, e.g. `88` for precinct 0006). Ken McFeeters (1639) and "Alabama" Will Santivasci (862) were already exact across all 41 precincts. Tuberville's 41 per-precinct values were recovered from the PaddleOCR `morgan_2read` cache (a higher-granularity re-render of the Morgan Canvas Report): page p017 holds precincts 0001–0035 cleanly in a 6-cell-per-row, two-precincts-per-row layout `[name1, McF1, Sant1, Tuber1, name2, "McF2 Sant2 Tuber2"]`, and page p018 supplies 0036–0041 (ABSENTEE/PROVISIONAL) where Tuberville's value is always the large trailing number. Column order `[McFeeters, Santivasci, Tuberville]` was confirmed by the p018 `CANDIDATE TOTALS` row `[1639, 862, 13596]` and by the p017 McFeeters/Santivasci per-precinct values matching the existing CSV with 0 mismatches. Recovered Tuberville sum = 13596 = certified exactly (Morgan's canvass matches certified with no ~10-vote deficit, unlike Tuscaloosa). The 36 garbage rows were removed and 41 clean Tuberville rows inserted; Governor now passes the authoritative crosscheck. **Remaining Morgan mismatch: Superintendent, Morgan County Board of Education (REP)** — still flagged by `mismatched_contest_keys`; not yet investigated.
- **Conecuh — U.S. Senate & PSC Place 2 REP dropped precincts (2026-08-13):** the original conversion of `20260519__al__primary__conecuh__precinct.csv` had dropped one column from each of these per-office-page contests: **U.S. Senate REP** carried only 21 of 29 precincts (0001–0008 missing) with all 7 candidate totals running ~30–40% short of certified, plus a spurious `Unverified Candidate 8` (31 votes) parse artifact; **PSC Place 2 REP** carried only 16 of 29 precincts (0017–0029 missing) with all 4 candidate totals at ~55% of certified (Zeigler 371 vs 667, Beeker 142 vs 257, Woodall 126 vs 226, Andrews 78 vs 156) — the classic half-the-precincts signature of a dropped page-break column. Both contests were fully re-extracted from the PaddleOCR `conecuh_2read` cache (a higher-granularity re-render that splits each contest onto its own page pair): USSenate REP = p016 (0001–0008) + p017 (0009–0029), 7 columns `[Burton, Deas, Hudson, Marshall, Moore, Murphy, Walker]` confirmed by the p017 `CANDIDATE TOTALS` row `[53, 33, 249, 345, 654, 22, 114]` = certified exactly; PSC2 REP = p028 (0001–0016) + p029 (0017–0029), 4 columns `[Andrews, Beeker, Woodall, Zeigler]` confirmed by the p029 `CANDIDATE TOTALS` row `[156, 257, 226, 667]` = certified exactly. Column order was further validated against the existing CSV overlap (PSC2 0001–0016: 0 per-precinct mismatches; USSenate 0009–0029: 0 mismatches except 2 tiny Rodney Walker values the 2read corrects, confirmed by the 2read totals matching certified). The spurious `Unverified Candidate 8` was dropped. Both contests now total certified exactly across all 29 precincts; Conecuh has **0 crosscheck mismatches**.
- **Shelby — Commissioner of Agriculture & Industries REP mislabeled column (2026-08-13):** the CSV carried 3 rows per precinct for this contest but only 2 candidate names — `Christina Woerner McInnis` appeared twice and `Cory Hill` once, with Jack Williams missing entirely. The first "McInnis" row per precinct was actually Williams' votes (column mislabeled): the first-McInnis precinct sum was 6963 = certified Williams, the second-McInnis sum was 8210 = certified McInnis, Hill was 7240 = certified. Confirmed per-precinct against the PaddleOCR `shelby_2read` cache (p033 precincts 0001–0028 + p034 precincts 0010–0041, ballot order `[Hill, McInnis, Williams]` verified by the p034 `CANDIDATE TOTALS` row `[7240, 8210, 6963]` = certified exactly): 0 per-precinct mismatches across all 40 precincts. Fix was a pure rename of the first McInnis row per precinct → `Jack Williams` (no vote values changed). Shelby now has **0 crosscheck mismatches**. (The Shelby State Republican Executive Committee placeholders in the unresolved table below are a separate local race not present in the certified county file.)
- **Monroe — LtGov & U.S. Senate REP dropped/misattributed precincts (2026-08-13):** two REP contests had precinct gaps in `20260519__al__primary__monroe__precinct.csv`. **Lieutenant Governor REP** was missing precinct 0004 (28/29 precincts); the deficit exactly equaled 0004's canvass values, so the other 28 were correct and only 0004 needed adding. **U.S. Senate REP** had two problems: precinct 0001 was missing entirely, and precincts 0023 + 0024 had been merged — precinct 0023 carried 14 rows (each candidate twice), where the first occurrence was 0023's real value and the second was 0024's value, with 0024 absent as its own precinct. Both were recovered from the PaddleOCR `monroe_2read` cache (Blue Sheet/Canvas Report format): LtGov 0004 from p025 (`[56,5,2,11,6,4,52]` in ballot order `[Allen, Bishop, Childress, Pate, Tankersley, Nicole, Wahl]`, confirmed by the p026 LtGov `CANDIDATE TOTALS` row `[933,93,70,206,28,92,916]` = certified); USSenate 0001 from p026's second contest (precincts 0001–0008 with names lost in OCR — 0001 = `[1,0,9,37,28,1,1]`, ballot order `[Burton, Deas, Hudson, Marshall, Moore, Murphy, Walker]` confirmed by the p027 `CANDIDATE TOTALS` row `[88,59,440,783,861,43,76]` = certified) and 0024 from p027. The fix deduped 0023 to its first-occurrence rows, added 0024 and 0001 as separate precincts (0001's precinct name taken from the canonical Governor-DEM row since the canvass lost it). Both contests now total certified exactly across all 29 precincts; Monroe has **0 crosscheck mismatches**.
- **Houston — Attorney General REP dropped precincts 0001–0005 (2026-08-13):** `20260519__al__primary__houston__precinct.csv` had AG REP for only 25 of 30 precincts (0006–0030); the other REP contests (Governor, U.S. Senate, Lieutenant Governor) all had 30 and matched certified, so only AG REP's first 5 precincts were missing. The deficit was uniform across the 3 candidates at ~84% of certified. Recovered from the PaddleOCR `houston_2read` cache: p030's bottom table holds AG REP precincts 0001–0005 (3 columns), and p031 holds the continuation 0006–0030 with an identical header. Column order `[Pamela L. Casey, Jay Mitchell, Katherine Robertson]` was confirmed two ways — (a) the p031 column values for 0006–0030 exactly match the CSV's existing (correct) rows, and (b) each 0001–0005 column sum equals exactly one candidate's deficit (Casey 83+30+47+92+74=326=2083−1757; Mitchell 152+87+79+152+133=603=3439−2836; Robertson 221+57+111+192+172=753=4861−4108), and the three deficits are distinct so the assignment is unambiguous. Inserted 15 rows (5 precincts × 3 candidates) after each precinct's U.S. Senate DEM block, matching the 0006 office ordering. AG REP now totals certified exactly across all 30 precincts; Houston has **0 crosscheck mismatches**. (Houston canvass = certified exactly, no ~10-vote deficit.)
- **Crenshaw — U.S. Senate DEM rename + Wheeler digit fix + Wess source discrepancy (2026-08-13):** `20260519__al__primary__crenshaw__precinct.csv` had two issues in this contest. (1) The 4th DEM candidate was labeled `Unverified Candidate 3` (431 votes) — the canvass PDF's rotated name header didn't OCR, but the votes are **Everett Wess**'s; renamed all 17 rows. (2) **Mark S. Wheeler II** precinct 0011 (COUNTY COURTHOUSE) was `19` in the CSV but the canvass PDF shows `18` — confirmed by two independent PaddleOCR reads (`Crenshaw_cert_2` primary read and `crenshaw_2read`): both prints' `CANDIDATE TOTALS` row says Wheeler `91`, and only `18` at 0011 makes the precincts sum to 91 (the primary read's per-precinct `19` was a mis-OCR inconsistent with its own totals row; the 2read resolved it). Fixed 19→18, so Wheeler totals 91 = certified. After these fixes Larriett (123), Sweetser (51), and Wheeler (91) all match certified exactly. **Remaining residual: Everett Wess 431 vs certified 491** — a 60-vote shortfall that is NOT an extraction error: both independent OCR reads agree the canvass PDF prints 431 (per-precinct cells sum exactly to the printed `CANDIDATE TOTALS` 431, internally consistent), while the certified county file records 491. Unlike Tuscaloosa's ~10-vote deficit, this gap falls on a single candidate with every other Crenshaw contest matching certified exactly, so it is not a uniform late-absentee discrepancy; it is a preliminary-canvass-vs-certified discrepancy specific to Wess, accepted as a residual until a certified precinct-level source closes it. Crosscheck will show 1 mismatch (Wess) for this contest.

## Party coverage by county

Which party sections actually appear in each county's source data — a county showing only REP isn't missing data, its document simply has no DEM section (confirmed on Bullock, for example: no Democratic contest exists anywhere in the file). `?` means neither party could be confirmed for that county, worth a manual look.

| County | REP | DEM | Coverage |
|---|---|---|---|
| Baldwin | ✓ | ✓ | Both |
| Barbour | ✓ | ✓ | Both |
| Blount | ✓ | ✓ | Both |
| Bullock | ✓ |  | REP only |
| Butler | ✓ | ✓ | Both |
| Calhoun | ✓ | ✓ | Both |
| Chambers | ✓ | ✓ | Both |
| Chilton | ✓ | ✓ | Both |
| Clarke | ✓ | ✓ | Both |
| Cleburne | ✓ | ✓ | Both |
| Coffee | ✓ | ✓ | Both |
| Colbert | ✓ | ✓ | Both |
| Coosa | ✓ | ✓ | Both |
| Covington | ✓ | ✓ | Both |
| Cullman | ✓ | ✓ | Both |
| Dale | ✓ | ✓ | Both |
| Dallas | ✓ | ✓ | Both |
| DeKalb | ✓ | ✓ | Both |
| Elmore | ✓ | ✓ | Both |
| Escambia | ✓ | ✓ | Both |
| Etowah | ✓ | ✓ | Both |
| Franklin |  |  | ? (no party data) |
| Geneva | ✓ | ✓ | Both |
| Hale | ✓ |  | REP only |
| Henry | ✓ | ✓ | Both |
| Houston | ✓ | ✓ | Both |
| Jackson | ✓ | ✓ | Both |
| Jefferson | ✓ | ✓ | Both |
| Lamar | ✓ | ✓ | Both |
| Lauderdale | ✓ | ✓ | Both |
| Lawrence | ✓ |  | REP only |
| Lee | ✓ | ✓ | Both |
| Limestone | ✓ | ✓ | Both |
| Lowndes | ✓ | ✓ | Both |
| Macon | ✓ | ✓ | Both |
| Madison | ✓ | ✓ | Both |
| Marengo | ✓ | ✓ | Both |
| Marion | ✓ | ✓ | Both |
| Marshall | ✓ | ✓ | Both |
| Monroe | ✓ | ✓ | Both |
| Montgomery | ✓ | ✓ | Both |
| Perry | ✓ |  | REP only |
| Pickens | ✓ | ✓ | Both |
| Pike | ✓ | ✓ | Both |
| Russell | ✓ | ✓ | Both |
| Shelby | ✓ | ✓ | Both |
| St. Clair | ✓ |  | REP only |
| Talladega | ✓ | ✓ | Both |
| Tallapoosa | ✓ | ✓ | Both |
| Tuscaloosa | ✓ |  | REP only |
| Walker | ✓ | ✓ | Both |
| Washington | ✓ | ✓ | Both |
| Winston | ✓ | ✓ | Both |

**48 counties have both parties, 6 REP-only, 0 DEM-only, 1 with no party data** (of 55 county files).

## Failed-checksum contests

| County | Office | Party | Computed | Printed | Note |
|---|---|---|---|---|---|
| Blount | PROPOSED STATEWIDE AMENDMENT 2 |  | [4790, 765, 3507] | [4790, 4272, None] | checksum mismatch (written anyway) |
| Clarke | COMMISSIONER OF AGRICULTURE AND INDUSTRIES | REP | [27, 25, 36] | [578, 711, 684] | checksum mismatch (written anyway) |
| Clarke | LIEUTENANT GOVERNOR | REP | [913, 59, 32, 111, 13, 69, 940, 24] | [913, 59, 32, 111, 13, 69, 964, None] | checksum mismatch (written anyway) |
| Clarke | STATE AUDITOR | REP |  |  | no printed totals row found at all |
| Clarke | STATE DEMOCRATIC EXECUTIVE COMMITTEE (FEMALE), DISTRICT NO. 65 | DEM | [682, 173, 156] | [682, None, 329] | checksum mismatch (written anyway) |
| Cleburne | GOVERNOR | DEM | [13, 8, 6, 113, 11, 4] | [13, 8, 6, 115, 12, 4] | checksum mismatch (written anyway) |
| Cleburne | LIEUTENANT GOVERNOR | DEM | [81, 66] | [1, 0] | checksum mismatch (written anyway) |
| Coffee | GOVERNOR | DEM | [259, 109, 119, 1068, 159, 24] | [259, 109, 119, 1069, 159, 24] | checksum mismatch (written anyway) |
| Coffee | PROPOSED STATEWIDE AMENDMENT 2 |  | [3053, 1939] | [4835, 3154] | checksum mismatch (written anyway) |
| Coffee | PROPOSED STATEWIDE AMENDMENT 2 |  |  |  | no printed totals row found at all |
| Colbert | GOVERNOR | DEM | [502, 77, 74, 1962, 84, 16] | [502, None, 77, 74, 1962, None] | checksum mismatch (written anyway) |
| Colbert | PUBLIC SERVICE COMMISSION, PLACE NO. 1 | DEM | [1495, 376, 19, 628] | [1495, 376, None, 647] | checksum mismatch (written anyway) |
| Dallas | PUBLIC SERVICE COMMISSION, PLACE NO. 1 | REP | [1078, 84, 197] | [1078, 281, None] | checksum mismatch (written anyway) |
| Escambia | GOVERNOR | DEM | [373, 152, 63, 1052, 28, 13] | [373, None, 152, 63, 1052, 28] | checksum mismatch (written anyway) |
| Escambia | PROPOSED STATEWIDE AMENDMENT 2 |  |  |  | no printed totals row found at all |
| Etowah | ETOWAH COUNTY SHERIFF | REP | [1509, 9800] | [1543, 9931] | checksum mismatch (written anyway) |
| Etowah | MEMBER, ETOWAH COUNTY COMMISSION, DISTRICT NO. 4 | REP | [1043, 1232] | [1049, 1237] | checksum mismatch (written anyway) |
| Henry | ATTORNEY GENERAL | REP | [638, 940, 1358] | [638, 940, 1359] | checksum mismatch (written anyway) |
| Henry | MEMBER, HENRY COUNTY BOARD OF EDUCATION, DISTRICT NO. 2 | REP | [299, 369] | [426, 514] | checksum mismatch (written anyway) |
| Jackson | PROPOSED STATEWIDE AMENDMENT 2 |  |  |  | no printed totals row found at all |
| Lawrence | COMMISSIONER OF AGRICULTURE AND INDUSTRIES | REP | [335, 203, 242] | [1986, 1127, 1297] | checksum mismatch (written anyway) |
| Lawrence | STATE REPUBLICAN EXECUTIVE COMMITTEE, LAWRENCE COUNT | REP | [1700, 2576, 626] | [1700, 3202, None] | checksum mismatch (written anyway) |
| Lawrence | UNITED STATES SENATOR | REP | [116, 119, 954, 1367, 1863, 79, 222] | [116, 64, 1012, 1421, 1924, 82, 237] | checksum mismatch (written anyway) |
| Lee | COMMISSIDNER OF AGRICULTURE AND INDUSTRIES | REP | [3875, 3954, 2791] | [5518, 5519, 3996] | checksum mismatch (written anyway) |
| Lee | COMMISSIONER OF AGRICULTURE AND INDUSTRIES | REP |  |  | no printed totals row found at all |
| Lee | LEE COUNTY CORONER | REP | [5051, 8726] | [5051, 9843] | checksum mismatch (written anyway) |
| Lee | LEE COUNTY CORONER | REP |  |  | no printed totals row found at all |
| Lee | LIEUTENANT GOVERNOR | REP | [5068, 811, 479, 1792, 633, 1376, 6189] | [811, 479, 1928, 633, 1376, 6189, None] | checksum mismatch (written anyway) |
| Lee | MEMBER, LEE COUNTY COMMISSION, DISTRICT NO. 4 | REP |  |  | no printed totals row found at all |
| Lee | MEMBER. LEE COUNTY COMMISSION, DISTRICT NO. 4 | REP | [410, 254] | [3145, 2463] | checksum mismatch (written anyway) |
| Lee | STATE REPRESENTATIVE, DISTRICT NO. 82 | DEM |  |  | no printed totals row found at all |
| Lee | STATE REPUBLICAN EXECUTIVE COMMITTEE, LEE COUNTY, PL | REP |  |  | no printed totals row found at all |
| Limestone | CHAIRMAN, LIMESTONE COUNTY COMMISSION | REP | [792, 2832, 1443, 0, 337] | [1615, 5401, 3726, None, 680] | checksum mismatch (written anyway) |
| Limestone | CHAIRMAN. LIMESTONE COUNTY COMMISSION | REP |  |  | no printed totals row found at all |
| Limestone | MEMBER, STATE BOARD OF EDUCATION, DIST NO 8 | REP |  |  | no printed totals row found at all |
| Limestone | MEMBER. STATE BOARD OF EDUCATION, DIST NO 8 | REP | [2865, 3004, 1497, 1, 1363] | [3929, 3837, 1907, 1, 1748] | checksum mismatch (written anyway) |
| Limestone | STATE REPUBLICAN EXEC COMM, LIMESTONE CO - PL NO 2 | REP | [1529, 4304, 1589, 2, 2296] | [1885, 5046, 1890, 2, 2599] | checksum mismatch (written anyway) |
| Limestone | STATE REPUBLICAN EXEC COMM, LIMESTONE CO - PL NO 3 | REP | [7838, 1447, 0, 295, 1842] | [7838, 1447, None, None, 2137] | checksum mismatch (written anyway) |
| Limestone | STATE REPUBLICAN EXEC COMM, LIMESTONE CO - PL NO 4 | REP |  |  | no printed totals row found at all |
| Limestone | STATE REPUBLICAN EXEC COMM. LIMESTONE CO - PL NO 4 | REP | [3160, 3661, 0, 2510] | [3908, 4459, None, 3055] | checksum mismatch (written anyway) |
| Marengo | GOVERNOR | REP | [85, 29, 796, 47] | [45, 26, 843, None] | checksum mismatch (written anyway) |
| Marengo | MEMBER, MARENGO COUNTY COMMISSION, DISTRICT NO. 5 | DEM | [474, 396, 80] | [513, 394, None] | checksum mismatch (written anyway) |
| Monroe | LIEUTENANT GOVERNOR | REP | [877, 88, 68, 195, 22, 88, 864] | [933, 93, 70, 206, 28, 92, 916] | checksum mismatch (written anyway) |
| Monroe | UNITED STATES SENATOR | REP | [87, 59, 431, 746, 833, 42, 75] | [88, 59, 440, 783, 861, 43, 76] | checksum mismatch (written anyway) |
| Morgan | ATTORNEY GENERAL | REP |  |  | no printed totals row found at all |
| Morgan | COMMISSIONER OF AGRICULTURE AND INDUSTRIES |  |  |  | no printed totals row found at all |
| Morgan | COUNTY | REP |  |  | no printed totals row found at all |
| Morgan | MEMBER, MORGAN COUNTY COMMISSION, DISTRICT NO. 1 | REP |  |  | no printed totals row found at all |
| Morgan | MORGAN COUNTY REVENUE COMMISSIONER | REP |  |  | no printed totals row found at all |
| Morgan | MORGAN COUNTY SHERIFF | REP |  |  | no printed totals row found at all |
| Morgan | PROPOSED STATEWIDE AMENDMENT 1 |  |  |  | no printed totals row found at all |
| Morgan | PROPOSED STATEWIDE AMENDMENT 2 |  |  |  | no printed totals row found at all |
| Morgan | PUBLIC SERVICE COMMISSION, PLACE NO. 2 | REP |  |  | no printed totals row found at all |
| Morgan | SECRETARY OF STATE |  |  |  | no printed totals row found at all |
| Morgan | STATE AUDITOR |  |  |  | no printed totals row found at all |
| Morgan | STATE REPRESENTATIVE, DISTRICT NO. 8 | REP |  |  | no printed totals row found at all |
| Morgan | STATE TREASURER | REP |  |  | no printed totals row found at all |
| Morgan | SUPERINTENDENT, MORGAN COUNTY BOARD OF EDUCATION COUNTY | REP |  |  | no printed totals row found at all |
| Perry | SECRETARY OF STATE | REP | [283, 70, 28] | [283, 71, 28] | checksum mismatch (written anyway) |
| Pickens | LIEUTENANT GOVERNOR | DEM | [1149, 584] | [1162, 595] | checksum mismatch (written anyway) |
| Russell | RUSSELL COUNTY, ALABAMA MAY 19, 2026 |  |  |  | no printed totals row found at all |
| Russell | SECRETARY OF STATE | REP | [1271, 446, 189] | [1271, None, 446] | checksum mismatch (written anyway) |
| Washington | ATTORNEY GENERAL | REP | [829, 1060, 1415] | [829, None, 1060] | checksum mismatch (written anyway) |
| Washington | COMMISSIONER OF AGRICULTURE AND INDUSTRIES | REP | [770, 347, 686, 414, 1110] | [770, 761, 1796, None, None] | checksum mismatch (written anyway) |
| Washington | LIEUTENANT GOVERNOR | REP | [1433, 112, 160, 179, 43, 180, 849, 466] | [1433, 112, 160, 179, 43, 180, 1315, None] | checksum mismatch (written anyway) |
| Winston | MEMBER, WINSTON COUNTY COMMISSION, DISTRICT NO. 2 | REP |  |  | no printed totals row found at all |
| Winston | WINSTON COUNTY SHERIFF | REP | [332, 2259] | [988, 3241] | checksum mismatch (written anyway) |

**67 contests** across 21 counties.

> **Resolved (not data errors):** Lamar SECRETARY OF STATE, Lee LEE COUNTY SHERIFF, and Lee STATE AUDITOR were flagged here with `computed == printed` — the failure was a duplicate precinct row, not a numeric mismatch. Lamar precinct `0007 CREWS COMMUNITY CHU.` and Lee precinct `0002 COVINGTON PARK COMM.` had each been split across a page break into two rows per candidate; the duplicate rows have been merged (votes summed) so each `(precinct, candidate)` is now unique. Contest totals are unchanged; `verifier.py` no longer flags duplication on these files.

> **Counties not listed above** either passed the checksum check or had their status recorded only in the conversion script's run log (not recoverable from the CSVs alone). That includes the 12 counties added since this section was written — Calhoun, Elmore, Franklin, Hale, Jefferson, Macon, Montgomery, Sumter, Talladega, Tallapoosa, Tuscaloosa, and Walker — for which no checksum analysis has been run yet. Morgan, which appeared here previously, has been removed: no `20260519__al__primary__morgan__precinct.csv` exists in `2026/counties`. St. Clair was regenerated from the Canvas Report PDF; checksums have not been run against the new file.

## Unresolved candidate columns ("Unverified Candidate N")

One row per candidate *column* that never matched anyone in the county CSV (county-wide total shown, not per-precinct — the precinct-level rows are in the CSVs themselves under this placeholder name). This is real, checksummed vote data; only the candidate's identity is unconfirmed.

As of this revision, `match_candidates.py` (repo root) was run against all county CSVs: it builds an authority table of official candidate names and county-wide vote totals from `2026/20260519__al__primary__county.csv` (primary) and the SoS "Final Primary Results by County" xlsx (secondary, REP only), matches each contest by office-name similarity plus vote-total overlap, and renames placeholder/mislabeled candidate columns only where a match is provably unique (exact total match, or — when column and candidate counts line up exactly — a close ordered match). It also re-verified every already-named candidate against the official totals and auto-corrected exact-total swaps (e.g. two candidates' vote totals transposed by an earlier ballot-order patch). Every action it took is logged in `2026/candidate_match_report.md`, including contests it could not resolve (reported as MISMATCH or UNRESOLVED rather than guessed). This pass resolved 125 of the previously-listed placeholder columns and corrected 26 previously mislabeled ones; the table below reflects what's left.

**124 candidate columns** across 53 county/office combinations in 22 counties. (Since the last revision, 6 placeholder columns were resolved by authority-total matching — Washington LtGov `Unverified Candidate 8`→John Wahl and AgComm `Unverified Candidate 4/5`→McInnis/Williams, Clarke LtGov `Unverified Candidate 8`→John Wahl, Baldwin State House d=96 `Unverified Candidate 2`→Matt Simpson, Conecuh AG `Unverified Candidate 3`→Katherine Robertson, Conecuh State Auditor `Unverified Candidate 2`→Andrew Sorrell — each an exact or split-column match to the certified county total; Washington now has 0 crosscheck mismatches.)

| County | Office | District | Party | Column | County-wide total | Precincts w/ data |
|---|---|---|---|---|---|---|
| Baldwin | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 3 | 4379 | 51 |
| Baldwin | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 4 | 898 | 14 |
| Baldwin | Public Service Commission, Place No. 1 |  | REP | Unverified Candidate 3 | 1447 | 23 |
| Baldwin | Secretary of State |  | REP | Unverified Candidate 4 | 660 | 21 |
| Baldwin | State Republican Executive Committee, Baldwin County |  | REP | Unverified Candidate 1 | 65105 | 65 |
| Baldwin | State Republican Executive Committee, Baldwin County |  | REP | Unverified Candidate 2 | 45541 | 65 |
| Barbour | Barbour County Democratic Executive Committee (Female), District No. 1 |  | DEM | Unverified Candidate 1 | 213 | 6 |
| Barbour | Barbour County Democratic Executive Committee (Female), District No. 1 |  | DEM | Unverified Candidate 2 | 415 | 6 |
| Barbour | Barbour County Democratic Executive Committee (Female), District No. 1 |  | DEM | Unverified Candidate 3 | 83 | 6 |
| Blount | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 1 | 4790 | 26 |
| Blount | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 2 | 765 | 6 |
| Blount | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 3 | 3507 | 20 |
| Clarke | State Democratic Executive Committee |  | DEM | Unverified Candidate 2 | 511 | 19 |
| Colbert | Public Service Commission, Place No. 1 |  | DEM | Unverified Candidate 3 | 19 | 1 |
| Cullman | Cullman County Revenue Commissioner |  | REP | Unverified Candidate 3 | 38 | 1 |
| Cullman | State Board of Education | 6 | REP | Unverified Candidate 4 | 33 | 1 |
| Cullman | U.S. House | 4 | DEM | Unverified Candidate 3 | 123 | 11 |
| Cullman | U.S. House | 4 | REP | Unverified Candidate 4 | 24 | 1 |
| Dallas | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 3 | 212 | 19 |
| Dallas | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 4 | 93 | 12 |
| Dallas | Dallas County Democratic Executive Committee (Male), District No. 3 |  | DEM | Unverified Candidate 1 | 523 | 12 |
| Dallas | Dallas County Democratic Executive Committee (Male), District No. 3 |  | DEM | Unverified Candidate 2 | 895 | 12 |
| Dallas | Dallas County Democratic Executive Committee (Male), District No. 3 |  | DEM | Unverified Candidate 3 | 761 | 12 |
| Dallas | Dallas County Democratic Executive Committee (Male), District No. 3 |  | DEM | Unverified Candidate 4 | 633 | 12 |
| Dallas | Dallas County Democratic Executive Committee (Male), District No. 3 |  | DEM | Unverified Candidate 5 | 560 | 12 |
| Dallas | Dallas County Sheriff |  | DEM | Unverified Candidate 1 | 3177 | 31 |
| Dallas | Dallas County Sheriff |  | DEM | Unverified Candidate 2 | 3823 | 31 |
| Dallas | Public Service Commission, Place No. 1 |  | REP | Unverified Candidate 3 | 197 | 24 |
| Escambia | State Senate | 22 | REP | Unverified Candidate 1 | 2047 | 31 |
| Escambia | State Senate | 22 | REP | Unverified Candidate 2 | 1261 | 31 |
| Etowah | State Republican Executive Committee, Etowah County, |  | REP | Unverified Candidate 1 | 14117 | 21 |
| Etowah | State Republican Executive Committee, Etowah County, |  | REP | Unverified Candidate 2 | 5696 | 21 |
| Houston | Member, Houston County Commission, District No. 1 |  | DEM | Unverified Candidate 1 | 693 | 11 |
| Houston | Member, Houston County Commission, District No. 1 |  | DEM | Unverified Candidate 2 | 986 | 11 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 1 | 9654 | 33 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 2 | 7 | 1 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 3 | 3575 | 33 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 4 | 3193 | 25 |
| Lauderdale | State Republican Executive Committee, Lauderdale Cou |  | REP | Unverified Candidate 1 | 16051 | 33 |
| Lauderdale | State Republican Executive Committee, Lauderdale Cou |  | REP | Unverified Candidate 2 | 11607 | 33 |
| Lawrence | Lawrence County Sheriff |  | REP | Unverified Candidate 2 | 154 | 4 |
| Lawrence | Lawrence County Sheriff |  | REP | Unverified Candidate 3 | 2870 | 27 |
| Lawrence | State Republican Executive Committee, Lawrence Count |  | REP | Unverified Candidate 2 | 2576 | 20 |
| Lawrence | State Republican Executive Committee, Lawrence Count |  | REP | Unverified Candidate 3 | 626 | 11 |
| Lee | Commissidner of Agriculture and Industries |  | REP | Unverified Candidate 1 | 3875 | 18 |
| Lee | Commissidner of Agriculture and Industries |  | REP | Unverified Candidate 2 | 3954 | 18 |
| Lee | Commissidner of Agriculture and Industries |  | REP | Unverified Candidate 3 | 2791 | 18 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 1 | 314 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 2 | 296 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 3 | 193 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 4 | 286 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 5 | 448 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 6 | 912 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 7 | 197 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 8 | 554 | 16 |
| Lee | Member, Lee County Commission, District No. 5 |  | DEM | Unverified Candidate 1 | 1182 | 15 |
| Lee | Member, Lee County Commission, District No. 5 |  | DEM | Unverified Candidate 2 | 1786 | 15 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 1 | 2953 | 27 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 2 | 1139 | 27 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 3 | 2581 | 27 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 4 | 1406 | 27 |
| Limestone | Lieutenant Governor |  | REP | Unverified Candidate 8 | 4 | 28 |
| Limestone | Lieutenant Governor |  | REP | Unverified Candidate 9 | 431 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 7 | 2024 | 28 |
| Limestone | U.S. Senate |  | REP | Unverified Candidate 8 | 8 | 28 |
| Limestone | U.S. Senate |  | REP | Unverified Candidate 9 | 481 | 28 |
| Lowndes | Lowndes County Coroner |  | DEM | Unverified Candidate 1 | 2177 | 14 |
| Lowndes | Lowndes County Coroner |  | DEM | Unverified Candidate 2 | 448 | 14 |
| Lowndes | Member, Lowndes County Board of Education, District No. 1 |  | DEM | Unverified Candidate 1 | 324 | 6 |
| Lowndes | Member, Lowndes County Board of Education, District No. 1 |  | DEM | Unverified Candidate 2 | 321 | 6 |
| Lowndes | Member, Lowndes County Board of Education. District No. 5 |  | DEM | Unverified Candidate 1 | 236 | 5 |
| Lowndes | Member, Lowndes County Board of Education. District No. 5 |  | DEM | Unverified Candidate 2 | 396 | 5 |
| Lowndes | Member, Lowndes County Commission, District No. 1 |  | DEM | Unverified Candidate 1 | 211 | 6 |
| Lowndes | Member, Lowndes County Commission, District No. 1 |  | DEM | Unverified Candidate 2 | 391 | 6 |
| Lowndes | Member, Lowndes County Commission, District No. 1 |  | DEM | Unverified Candidate 3 | 109 | 6 |
| Lowndes | State Democratic Executive Committee |  | DEM | Unverified Candidate 1 | 1161 | 14 |
| Lowndes | State Democratic Executive Committee |  | DEM | Unverified Candidate 2 | 825 | 14 |
| Madison | Governor |  | DEM | Unverified Candidate 7 | 2 | 3 |
| Madison | Lieutenant Governor |  | REP | Unverified Candidate 8 | 419 | 3 |
| Madison | Madison County Democratic Executive Committee (Femal District No. 4 |  | DEM | Unverified Candidate 1 | 3559 | 16 |
| Madison | Madison County Democratic Executive Committee (Femal District No. 4 |  | DEM | Unverified Candidate 2 | 2582 | 16 |
| Madison | Madison County Democratic Executive Committee (Femal District No. 4 |  | DEM | Unverified Candidate 3 | 3991 | 16 |
| Madison | Madison County Democratic Executive Committee (Femal District No. 4 |  | DEM | Unverified Candidate 4 | 1411 | 16 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 1 |  | DEM | Unverified Candidate 1 | 2227 | 15 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 1 |  | DEM | Unverified Candidate 2 | 2140 | 15 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 1 |  | DEM | Unverified Candidate 3 | 1204 | 15 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 1 |  | DEM | Unverified Candidate 4 | 1537 | 15 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 4 |  | DEM | Unverified Candidate 1 | 4683 | 16 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 4 |  | DEM | Unverified Candidate 2 | 1757 | 16 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 4 |  | DEM | Unverified Candidate 3 | 2804 | 16 |
| Madison | Madison County Democratic Executive Committee (Male) District No. 4 |  | DEM | Unverified Candidate 4 | 1996 | 16 |
| Madison | State House | 25 | DEM | Unverified Candidate 3 | 70 | 1 |
| Madison | State Senate | 2 | DEM | Unverified Candidate 4 | 102 | 1 |
| Marengo | Governor |  | REP | Unverified Candidate 4 | 47 | 1 |
| Marengo | Marengo County Coroner |  | DEM | Unverified Candidate 1 | 2519 | 20 |
| Marengo | Marengo County Coroner |  | DEM | Unverified Candidate 2 | 1222 | 20 |
| Marengo | Marengo County Revenue Commissioner |  | DEM | Unverified Candidate 1 | 1670 | 20 |
| Marengo | Marengo County Revenue Commissioner |  | DEM | Unverified Candidate 2 | 2759 | 20 |
| Marengo | Member, Marengo County Commission, District No. 3 |  | DEM | Unverified Candidate 1 | 443 | 6 |
| Marengo | Member, Marengo County Commission, District No. 3 |  | DEM | Unverified Candidate 2 | 476 | 6 |
| Marengo | Member, Marengo County Commission, District No. 5 |  | DEM | Unverified Candidate 1 | 474 | 9 |
| Marengo | Member, Marengo County Commission, District No. 5 |  | DEM | Unverified Candidate 2 | 396 | 9 |
| Marengo | Member, Marengo County Commission, District No. 5 |  | DEM | Unverified Candidate 3 | 80 | 1 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 1 | 20341 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 2 | 20761 | 31 |
| Pickens | Member, Pickens County Commission, District No. 4 |  | DEM | Unverified Candidate 1 | 253 | 8 |
| Pickens | Member, Pickens County Commission, District No. 4 |  | DEM | Unverified Candidate 2 | 371 | 8 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 1 | 38585 | 40 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 3 | 23384 | 40 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 4 | 50 | 1 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 1 | 222 | 51 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 2 | 227 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 3 | 4201 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 4 | 3149 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 5 | 4020 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 6 | 103 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 7 | 241 | 51 |
| Washington | State House | 65 | REP | Unverified Candidate 2 | 1351 | 11 |
| Washington | State House | 65 | REP | Unverified Candidate 3 | 1078 | 10 |
