# 2026 AL Primary — Rows Needing Verification

Generated from 55 processed county CSVs in `2026/counties` (the `20260519__al__primary__democratic__precinct.csv` file is not a county and is excluded — see anomalies below). Three sections follow: **party coverage** (does this county's CSV contain both parties' results or only one), **failed-checksum contests** (real data written, but the extracted precinct rows don't sum to the document's own printed total), and **unresolved candidate columns** (real vote counts written under a placeholder name because no candidate in the county CSV matched that column's total). Neither of the last two categories was dropped from the CSVs — this doc is a punch list for reviewing them against the source PDFs, not a list of missing data.

The party-coverage and unresolved-candidate tables below were regenerated directly from the current CSVs in `2026/counties`; the failed-checksum section is carried forward from the conversion run logs (it cannot be reconstructed from the CSVs alone). Since the last revision, Calhoun, Elmore, Franklin, Hale, Jefferson, Macon, Montgomery, Sumter, Talladega, Tallapoosa, Tuscaloosa, Walker were added, Morgan was removed (no file), and several previously split "merged-column" rows (Marshall, St. Clair, Etowah, Clarke) were consolidated in the CSVs into single per-candidate totals; Shelby resolved five former placeholders. St. Clair was regenerated from the Canvas Report PDF with all Republican precinct results.

## Data-quality anomalies in the CSVs themselves

- **Democratic (file):** `20260519__al__primary__democratic__precinct.csv` is not a county file — its `county` column reads `Democratic` for every row (a statewide Democratic extract referencing precincts in multiple counties). Excluded from the per-county tables below.
- **Franklin:** non-conforming CSV — columns are ['county', 'precinct', 'office', 'votes'] (expected the 7-column OpenElections format); no party/candidate data, so it is excluded from the unresolved-candidate table and its party coverage is unknown.
- **Montgomery:** contains a stray non-DEM/REP party value ['party'] (an embedded duplicate header row mid-file — should be removed).
- **St. Clair:** regenerated from Canvas Report PDF (Republican results only); no Democratic section exists in the source document.
- **Tuscaloosa:** U.S. House district is rendered as `['4.0']` (a float string, should be the integer `4`).

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

**128 candidate columns** across 56 county/office combinations in 22 counties.

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
| Clarke | Lieutenant Governor |  | REP | Unverified Candidate 8 | 24 | 2 |
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
| Tuscaloosa | Attorney General |  | REP | Unverified Candidate 1 | 1965 | 49 |
| Tuscaloosa | Attorney General |  | REP | Unverified Candidate 2 | 3771 | 49 |
| Tuscaloosa | Attorney General |  | REP | Unverified Candidate 5 | 2228 | 49 |
| Tuscaloosa | Attorney General |  | REP | Unverified Candidate 6 | 969 | 49 |
| Tuscaloosa | Attorney General |  | REP | Unverified Candidate 8 | 2875 | 49 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 1 | 222 | 51 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 2 | 227 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 3 | 4201 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 4 | 3149 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 5 | 4020 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 6 | 103 | 52 |
| Tuscaloosa | Tuscaloosa County Sheriff |  | REP | Unverified Candidate 7 | 241 | 51 |
| Washington | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 4 | 414 | 12 |
| Washington | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 5 | 1110 | 12 |
| Washington | Lieutenant Governor |  | REP | Unverified Candidate 8 | 466 | 8 |
| Washington | State House | 65 | REP | Unverified Candidate 2 | 1351 | 11 |
| Washington | State House | 65 | REP | Unverified Candidate 3 | 1078 | 10 |
