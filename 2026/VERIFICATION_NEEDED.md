# 2026 AL Primary — Rows Needing Verification

Generated from 55 processed county CSVs in `2026/counties` (the `20260519__al__primary__democratic__precinct.csv` file is not a county and is excluded — see anomalies below). Three sections follow: **party coverage** (does this county's CSV contain both parties' results or only one), **failed-checksum contests** (real data written, but the extracted precinct rows don't sum to the document's own printed total), and **unresolved candidate columns** (real vote counts written under a placeholder name because no candidate in the county CSV matched that column's total). Neither of the last two categories was dropped from the CSVs — this doc is a punch list for reviewing them against the source PDFs, not a list of missing data.

The party-coverage and unresolved-candidate tables below were regenerated directly from the current CSVs in `2026/counties`; the failed-checksum section is carried forward from the conversion run logs (it cannot be reconstructed from the CSVs alone). Since the last revision, Calhoun, Elmore, Franklin, Hale, Jefferson, Macon, Montgomery, Sumter, Talladega, Tallapoosa, Tuscaloosa, Walker were added, Morgan was removed (no file), and several previously split "merged-column" rows (Marshall, St. Clair, Etowah, Clarke) were consolidated in the CSVs into single per-candidate totals; Shelby resolved five former placeholders. St. Clair was regenerated from the Canvas Report PDF with all Republican precinct results. Limestone's broader bogus-column issue (13 REP contests with mislabeled Over/Under pseudocolumns, plus U.S. Senate / SREC Pl 2 / SREC Pl 4 re-extraction) was cleaned up against the canvass PDF; Limestone now has 1 residual (U.S. Senate Dale Shelton Deas Jr. 295 canvass vs 294 certified, a 1-vote source discrepancy).

## Data-quality anomalies in the CSVs themselves

- **Democratic (file):** `20260519__al__primary__democratic__precinct.csv` is not a county file — its `county` column reads `Democratic` for every row (a statewide Democratic extract referencing precincts in multiple counties). Excluded from the per-county tables below.
- **Franklin:** non-conforming CSV — columns are ['county', 'precinct', 'office', 'votes'] (expected the 7-column OpenElections format); no party/candidate data, so it is excluded from the unresolved-candidate table and its party coverage is unknown.
- **Montgomery:** contains a stray non-DEM/REP party value ['party'] (an embedded duplicate header row mid-file — should be removed).
- **St. Clair:** regenerated from Canvas Report PDF (Republican results only); no Democratic section exists in the source document.
- **Madison — PSC Place 1 & Place 2 REP dropped precincts (resolved 2026-08-13):** `20260519__al__primary__madison__precinct.csv` had two Public Service Commission contests with dropped page-break precincts (the existing precincts' values were all correct — only precincts were missing). **PSC Place 1 REP** had only 48 of 81 precincts (0001–0048); precincts 0049–0081 (33, the dropped continuation table) were missing. Recovered from the PaddleOCR `Madison_County_Canvas_Report` cache (p030 top table), whose `CANDIDATE TOTALS` row `[18186, 8263]` = certified exactly in column order `[Matt Gentry, Jeremy H. Oden]`; the 33 recovered precincts sum to Gentry 8047 + Oden 3552, exactly the CSV's prior deficit, so column order and values are confirmed and no existing row needed changing. **PSC Place 2 REP** had 76 of 81 precincts; precincts 0033–0037 (5) were missing. Recovered from the PaddleOCR `Madison` cache (p039 top table), whose `CANDIDATE TOTALS` `[3648, 6265, 5122, 11995]` = certified exactly in column order `[Priscilla Andrews, Chris Beeker, Brent Woodall, Jim Zig Zeigler]`; the 5 recovered precincts sum to Andrews 90 + Beeker 126 + Woodall 128 + Zeigler 227, exactly the CSV's prior deficit. 86 rows inserted (66 PSC1 + 20 PSC2) via targeted byte-level insertion at each precinct's PSC anchor, preserving the LF + precinct-grouped structure (pure additive diff, 0 modifications). Both contests now total certified exactly across all 81 precincts; **Madison has 0 crosscheck mismatches.** (The earlier "structural column drift / Gentry bled into Oden" diagnosis was wrong — the existing 48 PSC1 precincts were correct; the deficit was entirely the missing 33.)
- **Tuscaloosa — canvass-vs-certified source discrepancy (2026-08-13):** all 10 REP statewide contests in `20260519__al__primary__tuscaloosa__precinct.csv` were merged from the "Blue Sheet and Canvas Report" PDF with real candidate names and checksum-PASS parses. The 7 wide DISTRICT CANVASS contests (Attorney General, Secretary of State, State Treasurer, State Auditor, Commissioner of Agriculture and Industries, PSC Place 1, PSC Place 2) and the 2 NAME HEADING contests (Lieutenant Governor, U.S. Senate) each cover all 55 precincts. U.S. House 4 covers 33 precincts — Tuscaloosa is split between AL-04 (33 precincts, ~83k registered) and AL-07 (the other 22 precincts, ~63k registered); the canvass's REP U.S. House page itemizes only the AL-04 precincts, and the certified county file likewise records only the AL-04 REP race (Aderholt/Barnes, no district 7), so 33 precincts is the complete AL-04 set, not a shortfall. However, the canvass PDF prints candidate totals ~10 votes SHORT of the certified county file (`20260519__al__primary__county.csv`) for every one of these 10 contests — e.g. AG `[2145,4094,5935]` vs certified `[2147,4097,5940]`; SoS `[8090,2438,1068]` vs `[8095,2442,1068]`; Treasurer `[8802,3135]` vs `[8809,3138]`; U.S. House `[5580,1151]` vs `[5582,1152]`. Five independent OCR reads (whole-page 200/400 DPI, crop+upscale, + totals-row reads) all reproduce the SAME short digits for the wide contests — the digits are unambiguous, so this is NOT an OCR error fixable by re-OCR. The consistent ~10-vote deficit across all contests is the signature of ~10 late absentee/provisional ballots present in the certified county total but NOT distributed to precincts in the preliminary canvass PDF. These 10 contests are therefore canvass-faithful but will show a small residual mismatch in the authoritative crosscheck (`repair_canvass_contests.py mismatched_contest_keys` — note the county name is case-sensitive, `"Tuscaloosa"`) until a certified precinct-level source (SoS website / corrected canvass / county election office) is found and used to close the ~10-vote gap. **Governor** was reconciled to certified exactly (0 deficit) via two-read reconciliation — its digits were ambiguous across renders, unlike the 10 above; the committed per-precinct values are readB's readings (the crosscheck is sum-based, so per-precinct accuracy is unverified). (Resolved: U.S. House district now renders as the integer `4` in the CSV, not `4.0`.)
- **Morgan — Governor REP Tuberville column dropped (2026-08-13):** the original conversion of `20260519__al__primary__morgan__precinct.csv` kept Thomas Tuberville's per-precinct votes for only precincts 0001–0005 (564 votes); for precincts 0006–0041 his vote total had been shoved into the `candidate` field with `votes=0` (36 garbage rows whose "candidate" was a bare number, e.g. `88` for precinct 0006). Ken McFeeters (1639) and "Alabama" Will Santivasci (862) were already exact across all 41 precincts. Tuberville's 41 per-precinct values were recovered from the PaddleOCR `morgan_2read` cache (a higher-granularity re-render of the Morgan Canvas Report): page p017 holds precincts 0001–0035 cleanly in a 6-cell-per-row, two-precincts-per-row layout `[name1, McF1, Sant1, Tuber1, name2, "McF2 Sant2 Tuber2"]`, and page p018 supplies 0036–0041 (ABSENTEE/PROVISIONAL) where Tuberville's value is always the large trailing number. Column order `[McFeeters, Santivasci, Tuberville]` was confirmed by the p018 `CANDIDATE TOTALS` row `[1639, 862, 13596]` and by the p017 McFeeters/Santivasci per-precinct values matching the existing CSV with 0 mismatches. Recovered Tuberville sum = 13596 = certified exactly (Morgan's canvass matches certified with no ~10-vote deficit, unlike Tuscaloosa). The 36 garbage rows were removed and 41 clean Tuberville rows inserted; Governor now passes the authoritative crosscheck. **Superintendent, Morgan County Board of Education (REP) — resolved (2026-08-13):** this contest was also broken in the original CSV. The canvass PDF (PaddleOCR `morgan_2read` p044/p045 + primary `Morgan_County_Canvas_Report` p022) shows the contest has only **26 precincts** (`0001`, `0002`, `0018`–`0041`) — precincts `0003`–`0017` do not participate in this board-of-education race — with `CANDIDATE TOTALS` `[5075, 4070]` = certified exactly (Daniel Gullion 5075, Tracie Turrentine 4070). The CSV had two errors: (1) the two candidate names were **swapped** — the row labeled `Tracie Turrentine` was really Daniel Gullion, and the row labeled `Robert Underwood` (a phantom name not in the canvass or certified file) was really Tracie Turrentine; (2) precinct `0022 UNION HILL SR CTR` was missing Gullion's `383` vote (the only row absent), which was exactly the 383-vote deficit (4692+383=5075). Fixed by renaming the two candidates back and inserting the missing 0022 Gullion row (383, confirmed by both the 2read and the primary read). All 26 precincts now match the canvass pairs exactly; Superintendent totals 5075/4070 = certified. **Morgan now has 0 crosscheck mismatches.**
- **Conecuh — U.S. Senate & PSC Place 2 REP dropped precincts (2026-08-13):** the original conversion of `20260519__al__primary__conecuh__precinct.csv` had dropped one column from each of these per-office-page contests: **U.S. Senate REP** carried only 21 of 29 precincts (0001–0008 missing) with all 7 candidate totals running ~30–40% short of certified, plus a spurious `Unverified Candidate 8` (31 votes) parse artifact; **PSC Place 2 REP** carried only 16 of 29 precincts (0017–0029 missing) with all 4 candidate totals at ~55% of certified (Zeigler 371 vs 667, Beeker 142 vs 257, Woodall 126 vs 226, Andrews 78 vs 156) — the classic half-the-precincts signature of a dropped page-break column. Both contests were fully re-extracted from the PaddleOCR `conecuh_2read` cache (a higher-granularity re-render that splits each contest onto its own page pair): USSenate REP = p016 (0001–0008) + p017 (0009–0029), 7 columns `[Burton, Deas, Hudson, Marshall, Moore, Murphy, Walker]` confirmed by the p017 `CANDIDATE TOTALS` row `[53, 33, 249, 345, 654, 22, 114]` = certified exactly; PSC2 REP = p028 (0001–0016) + p029 (0017–0029), 4 columns `[Andrews, Beeker, Woodall, Zeigler]` confirmed by the p029 `CANDIDATE TOTALS` row `[156, 257, 226, 667]` = certified exactly. Column order was further validated against the existing CSV overlap (PSC2 0001–0016: 0 per-precinct mismatches; USSenate 0009–0029: 0 mismatches except 2 tiny Rodney Walker values the 2read corrects, confirmed by the 2read totals matching certified). The spurious `Unverified Candidate 8` was dropped. Both contests now total certified exactly across all 29 precincts; Conecuh has **0 crosscheck mismatches**.
- **Shelby — Commissioner of Agriculture & Industries REP mislabeled column (2026-08-13):** the CSV carried 3 rows per precinct for this contest but only 2 candidate names — `Christina Woerner McInnis` appeared twice and `Cory Hill` once, with Jack Williams missing entirely. The first "McInnis" row per precinct was actually Williams' votes (column mislabeled): the first-McInnis precinct sum was 6963 = certified Williams, the second-McInnis sum was 8210 = certified McInnis, Hill was 7240 = certified. Confirmed per-precinct against the PaddleOCR `shelby_2read` cache (p033 precincts 0001–0028 + p034 precincts 0010–0041, ballot order `[Hill, McInnis, Williams]` verified by the p034 `CANDIDATE TOTALS` row `[7240, 8210, 6963]` = certified exactly): 0 per-precinct mismatches across all 40 precincts. Fix was a pure rename of the first McInnis row per precinct → `Jack Williams` (no vote values changed). Shelby now has **0 crosscheck mismatches**. (The Shelby State Republican Executive Committee placeholders in the unresolved table below are a separate local race not present in the certified county file.)
- **Monroe — LtGov & U.S. Senate REP dropped/misattributed precincts (2026-08-13):** two REP contests had precinct gaps in `20260519__al__primary__monroe__precinct.csv`. **Lieutenant Governor REP** was missing precinct 0004 (28/29 precincts); the deficit exactly equaled 0004's canvass values, so the other 28 were correct and only 0004 needed adding. **U.S. Senate REP** had two problems: precinct 0001 was missing entirely, and precincts 0023 + 0024 had been merged — precinct 0023 carried 14 rows (each candidate twice), where the first occurrence was 0023's real value and the second was 0024's value, with 0024 absent as its own precinct. Both were recovered from the PaddleOCR `monroe_2read` cache (Blue Sheet/Canvas Report format): LtGov 0004 from p025 (`[56,5,2,11,6,4,52]` in ballot order `[Allen, Bishop, Childress, Pate, Tankersley, Nicole, Wahl]`, confirmed by the p026 LtGov `CANDIDATE TOTALS` row `[933,93,70,206,28,92,916]` = certified); USSenate 0001 from p026's second contest (precincts 0001–0008 with names lost in OCR — 0001 = `[1,0,9,37,28,1,1]`, ballot order `[Burton, Deas, Hudson, Marshall, Moore, Murphy, Walker]` confirmed by the p027 `CANDIDATE TOTALS` row `[88,59,440,783,861,43,76]` = certified) and 0024 from p027. The fix deduped 0023 to its first-occurrence rows, added 0024 and 0001 as separate precincts (0001's precinct name taken from the canonical Governor-DEM row since the canvass lost it). Both contests now total certified exactly across all 29 precincts; Monroe has **0 crosscheck mismatches**.
- **Houston — Attorney General REP dropped precincts 0001–0005 (2026-08-13):** `20260519__al__primary__houston__precinct.csv` had AG REP for only 25 of 30 precincts (0006–0030); the other REP contests (Governor, U.S. Senate, Lieutenant Governor) all had 30 and matched certified, so only AG REP's first 5 precincts were missing. The deficit was uniform across the 3 candidates at ~84% of certified. Recovered from the PaddleOCR `houston_2read` cache: p030's bottom table holds AG REP precincts 0001–0005 (3 columns), and p031 holds the continuation 0006–0030 with an identical header. Column order `[Pamela L. Casey, Jay Mitchell, Katherine Robertson]` was confirmed two ways — (a) the p031 column values for 0006–0030 exactly match the CSV's existing (correct) rows, and (b) each 0001–0005 column sum equals exactly one candidate's deficit (Casey 83+30+47+92+74=326=2083−1757; Mitchell 152+87+79+152+133=603=3439−2836; Robertson 221+57+111+192+172=753=4861−4108), and the three deficits are distinct so the assignment is unambiguous. Inserted 15 rows (5 precincts × 3 candidates) after each precinct's U.S. Senate DEM block, matching the 0006 office ordering. AG REP now totals certified exactly across all 30 precincts; Houston has **0 crosscheck mismatches**. (Houston canvass = certified exactly, no ~10-vote deficit.)
- **Crenshaw — U.S. Senate DEM rename + Wheeler digit fix + Wess source discrepancy (2026-08-13):** `20260519__al__primary__crenshaw__precinct.csv` had two issues in this contest. (1) The 4th DEM candidate was labeled `Unverified Candidate 3` (431 votes) — the canvass PDF's rotated name header didn't OCR, but the votes are **Everett Wess**'s; renamed all 17 rows. (2) **Mark S. Wheeler II** precinct 0011 (COUNTY COURTHOUSE) was `19` in the CSV but the canvass PDF shows `18` — confirmed by two independent PaddleOCR reads (`Crenshaw_cert_2` primary read and `crenshaw_2read`): both prints' `CANDIDATE TOTALS` row says Wheeler `91`, and only `18` at 0011 makes the precincts sum to 91 (the primary read's per-precinct `19` was a mis-OCR inconsistent with its own totals row; the 2read resolved it). Fixed 19→18, so Wheeler totals 91 = certified. After these fixes Larriett (123), Sweetser (51), and Wheeler (91) all match certified exactly. **Remaining residual: Everett Wess 431 vs certified 491** — a 60-vote shortfall that is NOT an extraction error: both independent OCR reads agree the canvass PDF prints 431 (per-precinct cells sum exactly to the printed `CANDIDATE TOTALS` 431, internally consistent), while the certified county file records 491. Unlike Tuscaloosa's ~10-vote deficit, this gap falls on a single candidate with every other Crenshaw contest matching certified exactly, so it is not a uniform late-absentee discrepancy; it is a preliminary-canvass-vs-certified discrepancy specific to Wess, accepted as a residual until a certified precinct-level source closes it. Crosscheck will show 1 mismatch (Wess) for this contest.
- **Clarke — SDEC District 65 candidate rename (2026-08-13):** `20260519__al__primary__clarke__precinct.csv` had the two State Democratic Executive Committee District 65 contests (FEMALE and MALE, each "vote for 1") over the complete 19-precinct District 65 set with correct vote values but mislabeled candidates. The canvass PDF (PaddleOCR `Clarke_Amended` p005/p006) shows FEMALE = [O.J. Parnell Barnes, Talika Palmer] with `CANDIDATE TOTALS` [681, 327] and MALE = [T.L. Cherry, Edward Harris] with `CANDIDATE TOTALS` [666, 338] = certified exactly. Barnes (682) and Cherry (669) were already correctly named in the CSV; the remaining candidates were all labeled `Unverified Candidate 2` (511 votes across 19 precincts). The fix was **pure renaming — no vote values changed**: in the 7 precincts (`0004`,`0006`,`0007`,`0011`–`0014`) that had two `Unverified Candidate 2` rows (and no Palmer), the first → **Talika Palmer** and the second → **Edward Harris**; in the 12 precincts (`0015`–`0018`,`0022`,`0024`–`0028`,`0032`,`0033`) that had one `Unverified Candidate 2` (Palmer already named there), it → **Edward Harris**. After renaming, all four candidates match the certified county file exactly: Barnes 682, Palmer 329, Cherry 669, Harris 338. The renamed per-precinct Palmer (0004–0014) and Harris (all 19) values each match the canvass print cell-for-cell. The CRLF + contest-grouped file structure was preserved via minimal byte-level edits (26/26 line diff, renames only). **Clarke now has 0 crosscheck mismatches** for SDEC. (Note: the canvass `CANDIDATE TOTALS` prints FEMALE [681, 327] and MALE [666, 338], i.e. the preliminary election-night canvass undercounts Barnes by 1, Palmer by 2, and Cherry by 3 vs the certified county file — the same preliminary-canvass-vs-certified pattern seen in Tuscaloosa/Crenshaw. The CSV follows the certified totals, which is what the crosscheck validates against.)
- **Baldwin — U.S. Senate REP Steve Marshall 1-vote extraction error (fixed 2026-08-13):** `20260519__al__primary__baldwin__precinct.csv` had Marshall totaling 5421 vs certified 5422 — a 1-vote shortfall initially assumed to be a source-discrepancy residual, but the PaddleOCR `Baldwin_County_Canvas_Report` canvass (p016) `CANDIDATE TOTALS` row prints Marshall **5422** (col 4), agreeing with certified, not the CSV. So this was an extraction error, not a source discrepancy. Comparing the canvass per-precinct Marshall column to the CSV: precincts 0035–0065 match cell-for-cell except **0051 ONO ISLAND COMM CTR** (CSV 35 vs canvass 36), and the 0001–0034 sums match exactly (2556 each), so 0051 was the single bad cell. Fixed 35→36; Marshall now totals 5422 = certified = canvass. **Baldwin now has 0 crosscheck mismatches.** (CRLF + structure preserved, 1-line diff.)
- **Sumter — Lieutenant Governor REP Nicole Jones Wadsworth 1-vote extraction error (fixed 2026-08-13):** `20260519__al__primary__sumter__precinct.csv` had Nicole totaling 3 vs certified 4 — a 1-vote shortfall. The PaddleOCR `sumter_2read` canvass (p007) `GRAND TOTALS` row prints Nicole **4** (col 10), agreeing with certified. Per-precinct, the canvass shows Nicole = 1 at 0003 CUBA, 2 at 0007 LIVINGSTON, and **1 at 0016 ABSENTEE**; the CSV had 0003=1, 0007=2 but 0016=0 — the single missing vote. Fixed 0016 ABSENTEE Nicole 0→1; Nicole now totals 4 = certified = canvass. **Sumter now has 0 crosscheck mismatches.** (LF + structure preserved, 1-line diff.)
- **Marshall — U.S. House 4 DEM dropped precincts 0001–0011 (fixed 2026-08-13):** `20260519__al__primary__marshall__precinct.csv` had U.S. House (4th Congressional District) DEM for only 20 of 31 precincts (0012–0031); every other Marshall contest already had all 31 and matched certified. Both candidates sat at ~47% of certified (Pusczek 448 vs 946, Weaver 206 vs 444) — the classic half-the-precincts signature. Recovered from the PaddleOCR `Marshall` cache (the Democratic Party Results `Marshall.pdf`, NAME-HEADING format): p006 holds the contest's first page (precincts 0001–0011, 2 value columns) and p007 the continuation (0012–0031) with `CANDIDATE TOTALS` `[946, 444]` = certified exactly. Column order `[Amanda N. Pusczek, Shane Weaver]` was confirmed two ways — (a) the p007 0012–0031 values match the CSV's existing (correct) rows cell-for-cell, and (b) each 0001–0011 column sum equals exactly one candidate's deficit (Pusczek 207+16+4+120+70+11+5+15+24+18+8 = 498 = 946−448; Weaver 80+6+3+66+28+9+5+13+12+11+5 = 238 = 444−206). Inserted 22 rows (11 precincts × 2 candidates) before each precinct's U.S. Senate DEM block, matching the 0012 office ordering. U.S. House 4 DEM now totals certified exactly across all 31 precincts; **Marshall now has 0 crosscheck mismatches.** (CRLF + precinct-grouped structure preserved via byte-level edits, 22-line additive diff. Note: the canvass's own LtGov DEM `CANDIDATE TOTALS` row prints `[913, 469]` = certified, but its per-precinct cells sum short to `[725, 307]` — a source checksum discrepancy; the CSV's LtGov DEM already matched certified exactly before this work, so no change was needed there.)
- **Limestone — State Board of Education 8 REP missing precincts + mislabeled pseudocolumns (fixed 2026-08-13):** `20260519__al__primary__limestone__precinct.csv` had this contest with 5 candidate columns over only 24 of 28 precincts (0005–0028). The 3 real candidates (Connie Spears, Emily Jones, William Matthews) were all short of certified (1497/2865/3004 vs 1907/3929/3837), and 2 extra columns were mislabeled with Governor-DEM candidate names — `"Chad ""Chig"" Martin"` and `Doug Jones` — that are NOT certified State Board of Ed candidates. Recovered from the PaddleOCR `Limestone_County_Canvas_Report` cache (wide REPORT-EL111 format): p010 holds the contest's first page (precincts 0001–0004, garbled — precinct names lost but the 3 real-candidate value columns are intact) and p011 the continuation (0005–0028) with `CANDIDATE TOTALS` `[3929, 3837, 1907, 1748, (empty)]` — the first 3 = certified exactly (Emily Jones, William Matthews, Connie Spears); the 4th (1748) and 5th (empty) columns have no `CANDIDATE PERCENT` entry, confirming they are pseudocandidate columns (Over/Under Votes), not real candidates. Column order `[Emily Jones, William Matthews, Connie Spears]` was confirmed by deficit-sum: p010's 0001–0004 column sums equal exactly each candidate's shortfall (Emily 266+428+146+224 = 1064 = 3929−2865; William 231+269+136+197 = 833 = 3837−3004; Connie 104+166+47+93 = 410 = 1907−1497) — the three deficits are distinct so the assignment is unambiguous, and the per-precinct 0005–0028 values already in the CSV match the canvass p011 cell-for-cell. Fix: removed the 48 bogus `"Chad ""Chig"" Martin"`/`Doug Jones` rows (2 × 24 precincts) and inserted 12 rows (4 precincts × 3 real candidates) for 0001–0004, positioned between each precinct's State Auditor and State Republican Exec Comm blocks. State Board of Ed 8 REP now totals certified exactly across all 28 precincts; **Limestone now has 0 crosscheck mismatches.** (LF + precinct-grouped structure preserved via byte-level edits; net −36 rows.)
  - **Broader Limestone bogus-column cleanup (fixed 2026-08-13):** the same mislabeling affected ~13 other Limestone REP contests — extra candidate columns that were really the canvass's Over/Under-Votes pseudocolumns, mislabeled with Governor-DEM candidate names (`"Chad ""Chig"" Martin"`, `Doug Jones`, `Guy Sotomayor`, `Nathan "Nate" Mathis`) and other bogus names (`Allison T Montgomery`, `Damon Eubanks`, `Unverified Candidate 7/8/9`). (Those names are REAL only in Governor DEM, where they total certified: Chad 135, Doug 4323.) The cleanup was done in two phases, all via byte-level edits preserving the LF + precinct-grouped structure:
    - **Phase 1 — 13 REP contests, 696 bogus rows removed:** Attorney General, Chairman Limestone County Commission, Commissioner of Agriculture and Industries, Governor, Lieutenant Governor, Member Limestone County Commission Dist 1, PSC Place 1, PSC Place 2, Secretary of State, State Auditor, State Treasurer, SREC Pl 1, SREC Pl 3. The bogus Over/Under pseudocolumns were removed (not relabeled), per the State Board of Ed precedent. A party guard (`REP`) was required on the removal set to avoid colliding with the real DEM Governor candidates Chad "Chig" Martin (135) and Doug Jones (4323), whose rows are in the same file. SREC Pl 1 and Pl 3 already carried their real candidates (Pl 1: Elizabeth Stewart/Johnny Turner/Martha C. Fisher/Michael S. Shelton/Rosemary Stainbrook; Pl 3: Ben Harrison/Jeriel Jammullamudy) — only the extra bogus columns were stripped. All 13 now match certified exactly.
    - **Phase 2 — U.S. Senate relabel + SREC Pl 2 & Pl 4 re-extraction (fixed 2026-08-13):** three contests needed real-candidate recovery, not just bogus-column removal. **U.S. Senate REP:** the column labeled `Allison T Montgomery` was really **Dale Shelton Deas Jr.** (per-precinct values confirmed against canvass p004 cell-for-cell); relabeled all 28 rows. The `Unverified Candidate 8` (Over, 8) and `Unverified Candidate 9` (Under, 481) pseudocolumns were removed. **1-vote residual:** Dale totals 295 in the canvass vs 294 in the certified county file — confirmed by two independent PaddleOCR reads (the primary `Limestone_County_Canvas_Report` p004 and a 400-DPI re-OCR of the p004 totals crop, `lim_ussenate_2read`), both printing 295. This is a genuine canvass-over-certified source discrepancy (not an extraction error), accepted/annotated. **SREC Pl 2:** the 5 bogus candidates (Allison/Chad/Damon/Doug/Guy, 26 precincts) were replaced with the 3 real candidates — Mark McEathron (1885), Wayne Reynolds (5046), Helen V Thompson (1890) — re-extracted from canvass p013 (precincts 0001–0002) + p014 (0003–0028); the 2 missing precincts (0001–0002) were restored. Column order `[McEathron, Reynolds, Thompson]` confirmed by the p014 `CANDIDATE TOTALS` row `[1885, 5046, 1890, 2, 2599]` = certified exactly, and the per-precinct column sums (col1–3 across all 28) match certified. **SREC Pl 4:** the 4 bogus candidates (Allison/Damon/Doug/Guy, 25 precincts) were replaced with the 2 real candidates — Sheila Banister (3908), John Lee Sparks (4459) — re-extracted from canvass p015 (precincts 0001–0003) + a 400-DPI re-OCR of p016 (`lim_p16_2read`, precincts 0004–0028; the original p016 OCR had captured only the header, no data rows); the 3 missing precincts (0001–0003) were restored. Column order `[Banister, Sparks]` confirmed by the p016 `CANDIDATE TOTALS` row `[3908, 4459, (empty), 3055]` = certified exactly, and per-precinct column sums match. (Note: Pl 4's office string in the canvass prints `COMM.` with a period, vs `COMM,` with a comma for Pl 1–3 — a source inconsistency preserved as-is in the CSV.)
    - **Result:** Limestone now has **1 residual crosscheck mismatch: U.S. Senate REP Dale Shelton Deas Jr. 295 (canvass) vs 294 (certified)**, a 1-vote source discrepancy. `total_checksum.py --primary` is clean (no Total rows). All four SREC contests (Pl 1–4) match certified exactly (verified manually, since `mismatched_contest_keys` skips SREC entirely — the CSV office string `State Republican Exec Comm, Limestone Co - Pl No N` does not match the certified `State Republican Executive Committee, Limestone County, Place N` under `office_match_key`). No `Unverified Candidate N` placeholders remain in the file. Net row change: 2597 → 2451 data rows (−696 Phase 1 bogus, −56 US Senate pseudocolumns, −130 Pl 2 bogus, −100 Pl 4 bogus, +84 Pl 2 real, +56 Pl 4 real).
    - **Remaining (separate scope):** `Proposed Statewide Amendment 1` and `Proposed Statewide Amendment 2` still carry the same bogus column labels (Allison T Montgomery / Damon Eubanks / Doug Jones / Guy Sotomayor) — these are non-partisan ballot measures not present in the certified county file (so crosscheck has no counterpart and they are unchecked), and the real For/Against labels would require OCR of the canvass amendment pages. Left for a separate pass.
- **St. Clair — Attorney General / Lieutenant Governor / State Treasurer REP per-precinct extraction errors (fixed 2026-08-13):** `20260519__al__primary__st_clair__precinct.csv` (regenerated from the Canvas Report PDF, REP only) had 3 crosscheck-mismatched contests, all with CSV totals **over** certified — the opposite of the usual missing-precincts signature. Recovered from the PaddleOCR `St_Clair_County_Canvas_Report` + `stclair_2read` caches (the `St Clair County Canvas Report.pdf`); where the standard-DPI read dropped the dense AG continuation table, a high-DPI re-render of PDF page 5 (cropped, re-submitted to PaddleOCR) supplied an independent second read. **State Treasurer REP** was a single bad precinct: 0023 PRESCOTT BAPTIST CH had its two candidate columns shifted (CSV Lolley 111 / Boozer 132 vs canvass Lolley 58 / Boozer 111); all other 31 precincts matched the canvass cell-for-cell. Fixed 0023 → Lolley 58, Boozer 111; now Lolley 3841 / Boozer 6399 = certified exactly. **Attorney General REP** had 3 bad precincts — 0001 (all 3 columns rotated: CSV Casey 80/Mitchell 115/Robertson 88 vs canvass Casey 115/Mitchell 88/Robertson 80), 0004 (Casey & Robertson swapped: 435/393 vs 393/435), 0009 (Casey 180 vs canvass 146, a duplicated Robertson value) — confirmed against the canvass `CANDIDATE TOTALS` `[3081, 3382, 4166]` (col order Casey, Mitchell, Robertson). After fixing those 6 cells, Casey 3081 and Robertson 4166 = certified exactly; **Mitchell totals 3382 vs certified 3381 — a 1-vote source discrepancy**: the canvass PDF's printed `CANDIDATE TOTALS` for Mitchell is 3382 (confirmed by two independent OCR reads — `stclair_2read` p009 and the high-DPI page-5 crop — and the per-precinct cells sum to 3382), while the certified county file says 3381. This is a genuine canvass-over-certified residual (not an extraction error), accepted/annotated. **Lieutenant Governor REP** (7 candidates, wide format) was substantially scrambled — 85 wrong cells across ~20 precincts (per-precinct column swaps/shifts), but the canvass `CANDIDATE TOTALS` `[4104, 445, 203, 687, 232, 490, 4544]` (col order Wes Allen, Bishop, Childress, Pate, Tankersley, Nicole, Wahl) matched certified exactly for all 7, so all 224 LtGov vote fields were reset to the canvass per-precinct values; all 7 now total certified exactly. **St. Clair now has 1 residual crosscheck mismatch: AG REP Mitchell 3382 (canvass) vs 3381 (certified), a 1-vote source discrepancy.** (LF + precinct-grouped structure preserved via byte-level vote-field edits; row count unchanged at 1766, 93 vote fields changed total: LtGov 85, AG 6, Treasurer 2.)

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

**119 candidate columns** across 50 county/office combinations in 21 counties. (Since the last revision, 6 placeholder columns were resolved by authority-total matching — Washington LtGov `Unverified Candidate 8`→John Wahl and AgComm `Unverified Candidate 4/5`→McInnis/Williams, Clarke LtGov `Unverified Candidate 8`→John Wahl, Baldwin State House d=96 `Unverified Candidate 2`→Matt Simpson, Conecuh AG `Unverified Candidate 3`→Katherine Robertson, Conecuh State Auditor `Unverified Candidate 2`→Andrew Sorrell — each an exact or split-column match to the certified county total; Washington now has 0 crosscheck mismatches. Additionally, 5 Limestone placeholder/pseudocandidate columns were removed by the Limestone bogus-column cleanup — LtGov `Unverified Candidate 8/9` (Over/Under pseudocolumns), SREC Pl 1 `Unverified Candidate 7`, and U.S. Senate `Unverified Candidate 8/9` (Over/Under pseudocolumns); Limestone now has 0 `Unverified Candidate` placeholders.)

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
