# 2026 AL GOP Primary — Rows Needing Verification

Generated from 43 processed counties. Three sections below: **party coverage** (does this county's PDF contain both parties' results or only one), **failed-checksum contests** (real data written, but the extracted precinct rows don't sum to the document's own printed total), and **unresolved candidate columns** (real vote counts written under a placeholder name because no candidate in the county CSV matched that column's total). Neither of the last two categories was dropped from the CSVs — this doc is a punch list for reviewing them against the source PDFs, not a list of missing data.

## Party coverage by county

Which party sections actually appear in each county's source PDF — a county showing only REP isn't missing data, its document simply has no DEM section (confirmed on Bullock, for example: no Democratic contest exists anywhere in the file). `?` means neither party could be confirmed for that county at all, worth a manual look.

| County | REP | DEM | Coverage |
|---|---|---|---|
| Butler | ✓ | ✓ | Both |
| Bullock | ✓ |  | REP only |
| Coosa | ✓ | ✓ | Both |
| Dale | ✓ | ✓ | Both |
| Perry | ✓ |  | REP only |
| Lowndes | ✓ | ✓ | Both |
| Marion | ✓ | ✓ | Both |
| Henry | ✓ | ✓ | Both |
| Pickens | ✓ | ✓ | Both |
| Limestone | ✓ | ✓ | Both |
| St. Clair | ✓ |  | REP only |
| Marengo | ✓ | ✓ | Both |
| Clarke | ✓ | ✓ | Both |
| Lamar | ✓ | ✓ | Both |
| Coffee | ✓ | ✓ | Both |
| Colbert | ✓ | ✓ | Both |
| Covington | ✓ | ✓ | Both |
| Cleburne | ✓ | ✓ | Both |
| Escambia | ✓ | ✓ | Both |
| Marshall | ✓ | ✓ | Both |
| Barbour | ✓ | ✓ | Both |
| Etowah | ✓ | ✓ | Both |
| Lawrence | ✓ |  | REP only |
| Monroe | ✓ | ✓ | Both |
| Pike | ✓ | ✓ | Both |
| Dallas | ✓ | ✓ | Both |
| Russell | ✓ | ✓ | Both |
| Blount | ✓ | ✓ | Both |
| Geneva | ✓ | ✓ | Both |
| Winston | ✓ | ✓ | Both |
| Washington | ✓ | ✓ | Both |
| Lee | ✓ | ✓ | Both |
| Morgan | ✓ | ✓ | Both |
| Chilton | ✓ | ✓ | Both |
| Jackson | ✓ | ✓ | Both |
| Baldwin | ✓ | ✓ | Both |
| Chambers | ✓ | ✓ | Both |
| Cullman | ✓ | ✓ | Both |
| DeKalb | ✓ | ✓ | Both |
| Houston | ✓ | ✓ | Both |
| Lauderdale | ✓ | ✓ | Both |
| Madison | ✓ | ✓ | Both |
| Shelby | ✓ | ✓ | Both |

**38 counties have both parties, 5 REP-only, 0 DEM-only** (of 43 processed).

*Updated: Added Democratic precinct results to Geneva, Lamar, Limestone, Lowndes, Marshall, and Shelby from Democratic Party Results PDFs.*

## Failed-checksum contests

> **Note:** Baldwin, Chambers, Cullman, DeKalb, Houston, Lauderdale, Madison, and Shelby are not listed in this section. Their checksum status (computed precinct sums vs. the PDF's printed totals row) is captured in the conversion script's run log, not in the written CSVs, so it could not be recovered from the CSV output alone. The party-coverage and unresolved-candidate sections below were updated from the CSVs.

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
| St. Clair | COMMISSIONER OF AGRICULTURE AND INDUSTRIES | REP | [3630, 3308] | [3630, 3024] | checksum mismatch (written anyway) |
| St. Clair | UNITED STATES REPRESENTATIVE, 3RD CONGRESSIONAL DISTRICT | REP | [2093, 6008, 2586] | [2093, 8594, None] | checksum mismatch (written anyway) |
| Washington | ATTORNEY GENERAL | REP | [829, 1060, 1415] | [829, None, 1060] | checksum mismatch (written anyway) |
| Washington | COMMISSIONER OF AGRICULTURE AND INDUSTRIES | REP | [770, 347, 686, 414, 1110] | [770, 761, 1796, None, None] | checksum mismatch (written anyway) |
| Washington | LIEUTENANT GOVERNOR | REP | [1433, 112, 160, 179, 43, 180, 849, 466] | [1433, 112, 160, 179, 43, 180, 1315, None] | checksum mismatch (written anyway) |
| Winston | MEMBER, WINSTON COUNTY COMMISSION, DISTRICT NO. 2 | REP |  |  | no printed totals row found at all |
| Winston | WINSTON COUNTY SHERIFF | REP | [332, 2259] | [988, 3241] | checksum mismatch (written anyway) |

**69 contests** across 22 counties.

> **Resolved (not data errors):** Lamar SECRETARY OF STATE, Lee LEE COUNTY SHERIFF, and Lee STATE AUDITOR were flagged here with `computed == printed` — the failure was a duplicate precinct row, not a numeric mismatch. Lamar precinct `0007 CREWS COMMUNITY CHU.` and Lee precinct `0002 COVINGTON PARK COMM.` had each been split across a page break into two rows per candidate; the duplicate rows have been merged (votes summed) so each `(precinct, candidate)` is now unique. Contest totals are unchanged; `verifier.py` no longer flags duplication on these files.

## Unresolved candidate columns ("Unverified Candidate N")

One row per candidate *column* that never matched anyone in the county CSV (county-wide total shown, not per-precinct — the precinct-level rows are in the CSVs themselves under this placeholder name). This is real, checksummed vote data; only the candidate's identity is unconfirmed.

| County | Office | District | Party | Column | County-wide total | Precincts w/ data |
|---|---|---|---|---|---|---|
| Baldwin | Circuit Court Judge, 28Th Judicial Circuit, Place No |  | REP | Unverified Candidate 1 | 11243 | 65 |
| Baldwin | Circuit Court Judge, 28Th Judicial Circuit, Place No |  | REP | Unverified Candidate 2 | 4475 | 21 |
| Baldwin | Circuit Court Judge, 28Th Judicial Circuit, Place No |  | REP | Unverified Candidate 3 | 14431 | 44 |
| Baldwin | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 3 | 4379 | 51 |
| Baldwin | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 4 | 898 | 14 |
| Baldwin | District Court Judge, Baldwin County, Place No. 1 |  | REP | Unverified Candidate 3 | 390 | 1 |
| Baldwin | District Court Judge, Baldwin County, Place No. 1 |  | REP | Unverified Candidate 4 | 555 | 1 |
| Baldwin | Public Service Commission, Place No. 1 |  | REP | Unverified Candidate 3 | 1447 | 23 |
| Baldwin | Secretary of State |  | REP | Unverified Candidate 4 | 660 | 21 |
| Baldwin | State House | 96 | REP | Unverified Candidate 2 | 2796 | 12 |
| Baldwin | State Republican Executive Committee, Baldwin County |  | REP | Unverified Candidate 1 | 65105 | 65 |
| Baldwin | State Republican Executive Committee, Baldwin County |  | REP | Unverified Candidate 2 | 45541 | 65 |
| Baldwin | State Republican Executive Committee, Baldwin County |  | REP | Unverified Candidate 3 | 16548 | 65 |
| Barbour | Barbour County Democratic Executive Committee (Female), District No. 1 |  | DEM | Unverified Candidate 1 | 213 | 6 |
| Barbour | Barbour County Democratic Executive Committee (Female), District No. 1 |  | DEM | Unverified Candidate 2 | 415 | 6 |
| Barbour | Barbour County Democratic Executive Committee (Female), District No. 1 |  | DEM | Unverified Candidate 3 | 83 | 6 |
| Blount | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 1 | 4790 | 26 |
| Blount | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 2 | 765 | 6 |
| Blount | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 3 | 3507 | 20 |
| Blount | State Republican Executive Committee, Blount County, |  | REP | Unverified Candidate 1 | 3159 | 26 |
| Blount | State Republican Executive Committee, Blount County, |  | REP | Unverified Candidate 2 | 4635 | 26 |
| Chilton | Circuit Court Judge, 19Th Judicial Circuit, Place No |  | REP | Unverified Candidate 1 | 2338 | 20 |
| Chilton | Circuit Court Judge, 19Th Judicial Circuit, Place No |  | REP | Unverified Candidate 2 | 1477 | 20 |
| Chilton | Circuit Court Judge, 19Th Judicial Circuit, Place No |  | REP | Unverified Candidate 3 | 1948 | 20 |
| Clarke | Lieutenant Governor |  | REP | Unverified Candidate 8 | 24 | 2 |
| Clarke | State Democratic Executive Committee |  | DEM | Unverified Candidate 2 | 173 | 7 |
| Clarke | State Democratic Executive Committee |  | DEM | Unverified Candidate 2 | 338 | 19 |
| Cleburne | Lieutenant Governor |  | DEM | Unverified Candidate 1 | 81 | 15 |
| Cleburne | Lieutenant Governor |  | DEM | Unverified Candidate 2 | 66 | 15 |
| Colbert | Governor |  | DEM | Unverified Candidate 2 | 77 | 36 |
| Colbert | Governor |  | DEM | Unverified Candidate 6 | 16 | 36 |
| Colbert | Lieutenant Governor |  | REP | Unverified Candidate 7 | 2442 | 36 |
| Colbert | Public Service Commission, Place No. 1 |  | DEM | Unverified Candidate 3 | 19 | 1 |
| Covington | Member, Covington County Board of Education, Place N |  | REP | Unverified Candidate 1 | 2109 | 27 |
| Covington | Member, Covington County Board of Education, Place N |  | REP | Unverified Candidate 2 | 565 | 27 |
| Covington | Member, Covington County Board of Education, Place N |  | REP | Unverified Candidate 3 | 699 | 27 |
| Covington | Member, Covington County Board of Education, Place N |  | REP | Unverified Candidate 4 | 452 | 27 |
| Cullman | City of Cullman |  | REP | Unverified Candidate 1 | 708 | 6 |
| Cullman | City of Cullman |  | REP | Unverified Candidate 2 | 1316 | 6 |
| Cullman | City of Cullman |  | REP | Unverified Candidate 3 | 1246 | 6 |
| Cullman | City of Cullman |  | REP | Unverified Candidate 4 | 1145 | 6 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 1 | 6780 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 2 | 4207 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 3 | 7541 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 4 | 6469 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 5 | 5936 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 6 | 3744 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 7 | 5335 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 8 | 4476 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 9 | 5508 | 51 |
| Cullman | Cullman County Republican Executive Committee, At La |  | REP | Unverified Candidate 10 | 4015 | 51 |
| Cullman | Cullman County Republican Executive Committee, East |  | REP | Unverified Candidate 1 | 5860 | 51 |
| Cullman | Cullman County Republican Executive Committee, East |  | REP | Unverified Candidate 2 | 3002 | 51 |
| Cullman | Cullman County Republican Executive Committee, East |  | REP | Unverified Candidate 3 | 5177 | 51 |
| Cullman | Cullman County Republican Executive Committee, East |  | REP | Unverified Candidate 4 | 6238 | 51 |
| Cullman | Cullman County Republican Executive Committee, West |  | REP | Unverified Candidate 1 | 4607 | 51 |
| Cullman | Cullman County Republican Executive Committee, West |  | REP | Unverified Candidate 2 | 4248 | 51 |
| Cullman | Cullman County Republican Executive Committee, West |  | REP | Unverified Candidate 3 | 5916 | 51 |
| Cullman | Cullman County Republican Executive Committee, West |  | REP | Unverified Candidate 4 | 5266 | 51 |
| Cullman | Cullman County Revenue Commissioner |  | REP | Unverified Candidate 3 | 38 | 1 |
| Cullman | Cullman County Revenue Commissioner |  | REP | Unverified Candidate 4 | 54 | 1 |
| Cullman | Fairview / District No. 2 |  | REP | Unverified Candidate 1 | 683 | 9 |
| Cullman | Fairview / District No. 2 |  | REP | Unverified Candidate 2 | 841 | 9 |
| Cullman | Hanceville / District No. 4 |  | REP | Unverified Candidate 1 | 1055 | 12 |
| Cullman | Hanceville / District No. 4 |  | REP | Unverified Candidate 2 | 923 | 12 |
| Cullman | Secretary of State |  | REP | Unverified Candidate 4 | 46 | 1 |
| Cullman | Secretary of State |  | REP | Unverified Candidate 5 | 24 | 1 |
| Cullman | Secretary of State |  | REP | Unverified Candidate 6 | 15 | 1 |
| Cullman | State Board of Education | 6 | REP | Unverified Candidate 3 | 56 | 1 |
| Cullman | State Board of Education | 6 | REP | Unverified Candidate 4 | 33 | 1 |
| Cullman | State Republican Executive Committee, Cullman County |  | REP | Unverified Candidate 1 | 15444 | 51 |
| Cullman | State Republican Executive Committee, Cullman County |  | REP | Unverified Candidate 2 | 7701 | 51 |
| Cullman | State Republican Executive Committee, Cullman County |  | REP | Unverified Candidate 3 | 3526 | 51 |
| Cullman | U.S. House | 4 | DEM | Unverified Candidate 3 | 342 | 49 |
| Cullman | U.S. House | 4 | REP | Unverified Candidate 3 | 71 | 1 |
| Cullman | U.S. House | 4 | DEM | Unverified Candidate 4 | 1 | 1 |
| Cullman | U.S. House | 4 | REP | Unverified Candidate 4 | 24 | 1 |
| Cullman | U.S. House | 4 | DEM | Unverified Candidate 5 | 0 | 1 |
| Cullman | West Point / District No. 1 |  | REP | Unverified Candidate 1 | 1313 | 10 |
| Cullman | West Point / District No. 1 |  | REP | Unverified Candidate 2 | 886 | 10 |
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
| DeKalb | Attorney General |  | REP | Unverified Candidate 1 | 1830 | 46 |
| DeKalb | Attorney General |  | REP | Unverified Candidate 2 | 2517 | 46 |
| DeKalb | Attorney General |  | REP | Unverified Candidate 3 | 2219 | 46 |
| DeKalb | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 1 | 2318 | 46 |
| DeKalb | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 2 | 1482 | 46 |
| DeKalb | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 3 | 2663 | 46 |
| DeKalb | Dekalb County Republican Executive Committee, At Lar |  | REP | Unverified Candidate 1 | 2797 | 46 |
| DeKalb | Dekalb County Republican Executive Committee, At Lar |  | REP | Unverified Candidate 2 | 1999 | 46 |
| DeKalb | Dekalb County Republican Executive Committee, At Lar |  | REP | Unverified Candidate 3 | 1652 | 46 |
| DeKalb | Dekalb County Republican Executive Committee, Place District No. 1 |  | REP | Unverified Candidate 1 | 1533 | 13 |
| DeKalb | Dekalb County Republican Executive Committee, Place District No. 1 |  | REP | Unverified Candidate 2 | 1546 | 13 |
| DeKalb | Dekalb County Republican Executive Committee, Place District No. 2 |  | REP | Unverified Candidate 1 | 1189 | 14 |
| DeKalb | Dekalb County Republican Executive Committee, Place District No. 2 |  | REP | Unverified Candidate 2 | 1742 | 14 |
| DeKalb | Dekalb County Republican Executive Committee, Place District No. 4 |  | REP | Unverified Candidate 1 | 2620 | 11 |
| DeKalb | Dekalb County Republican Executive Committee, Place District No. 4 |  | REP | Unverified Candidate 2 | 1761 | 11 |
| DeKalb | Dekalb County Republican Executive Committee. Place District No. 2 |  | REP | Unverified Candidate 1 | 1092 | 14 |
| DeKalb | Dekalb County Republican Executive Committee. Place District No. 2 |  | REP | Unverified Candidate 2 | 380 | 14 |
| DeKalb | Governor |  | REP | Unverified Candidate 1 | 546 | 46 |
| DeKalb | Governor |  | REP | Unverified Candidate 2 | 283 | 46 |
| DeKalb | Governor |  | REP | Unverified Candidate 3 | 6290 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 1 | 2502 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 2 | 147 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 3 | 122 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 4 | 785 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 5 | 74 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 6 | 331 | 46 |
| DeKalb | Lieutenant Governor |  | REP | Unverified Candidate 7 | 2831 | 46 |
| DeKalb | Member, Dekalb County Commission, District No. 4 |  | REP | Unverified Candidate 1 | 1246 | 11 |
| DeKalb | Member, Dekalb County Commission, District No. 4 |  | REP | Unverified Candidate 2 | 739 | 11 |
| DeKalb | Member, Dekalb County Commission, District No. 4 |  | REP | Unverified Candidate 3 | 303 | 11 |
| DeKalb | Member, Dekalb County Commission. District No. 3 |  | REP | Unverified Candidate 1 | 752 | 14 |
| DeKalb | Member, Dekalb County Commission. District No. 3 |  | REP | Unverified Candidate 2 | 638 | 14 |
| DeKalb | President, Dekalb County Commission |  | REP | Unverified Candidate 1 | 2194 | 46 |
| DeKalb | President, Dekalb County Commission |  | REP | Unverified Candidate 2 | 4544 | 46 |
| DeKalb | Public Service Commission, Place No. 1 |  | REP | Unverified Candidate 1 | 4364 | 46 |
| DeKalb | Public Service Commission, Place No. 1 |  | REP | Unverified Candidate 2 | 1659 | 46 |
| DeKalb | Public Service Commission, Place No. 2 |  | REP | Unverified Candidate 1 | 690 | 46 |
| DeKalb | Public Service Commission, Place No. 2 |  | REP | Unverified Candidate 2 | 1424 | 46 |
| DeKalb | Public Service Commission, Place No. 2 |  | REP | Unverified Candidate 3 | 1814 | 46 |
| DeKalb | Public Service Commission, Place No. 2 |  | REP | Unverified Candidate 4 | 2211 | 46 |
| DeKalb | Secretary of State |  | REP | Unverified Candidate 1 | 3832 | 46 |
| DeKalb | Secretary of State |  | REP | Unverified Candidate 2 | 1428 | 46 |
| DeKalb | Secretary of State |  | REP | Unverified Candidate 3 | 909 | 46 |
| DeKalb | State Auditor |  | REP | Unverified Candidate 1 | 1903 | 46 |
| DeKalb | State Auditor |  | REP | Unverified Candidate 2 | 4413 | 46 |
| DeKalb | State Board of Education | 8 | REP | Unverified Candidate 1 | 2779 | 46 |
| DeKalb | State Board of Education | 8 | REP | Unverified Candidate 2 | 1746 | 46 |
| DeKalb | State Board of Education | 8 | REP | Unverified Candidate 3 | 1471 | 46 |
| DeKalb | State Senate | 10 | REP | Unverified Candidate 1 | 1308 | 19 |
| DeKalb | State Senate | 10 | REP | Unverified Candidate 2 | 545 | 19 |
| DeKalb | State Treasurer |  | REP | Unverified Candidate 1 | 4416 | 46 |
| DeKalb | State Treasurer |  | REP | Unverified Candidate 2 | 1883 | 46 |
| DeKalb | U.S. House | 4 | REP | Unverified Candidate 1 | 5246 | 46 |
| DeKalb | U.S. House | 4 | REP | Unverified Candidate 2 | 1747 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 1 | 120 | 46 |
| DeKalb | U.S. Senate |  | DEM | Unverified Candidate 2 | 253 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 2 | 68 | 46 |
| DeKalb | U.S. Senate |  | DEM | Unverified Candidate 3 | 193 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 3 | 1949 | 46 |
| DeKalb | U.S. Senate |  | DEM | Unverified Candidate 4 | 262 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 4 | 2065 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 5 | 2600 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 6 | 48 | 46 |
| DeKalb | U.S. Senate |  | REP | Unverified Candidate 7 | 174 | 46 |
| Escambia | Governor |  | DEM | Unverified Candidate 2 | 152 | 31 |
| Escambia | State Senate | 22 | REP | Unverified Candidate 1 | 2047 | 31 |
| Escambia | State Senate | 22 | REP | Unverified Candidate 2 | 1261 | 31 |
| Etowah | Etowah County School District |  | REP | Unverified Candidate 1 | 2761 | 19 |
| Etowah | Etowah County School District |  | REP | Unverified Candidate 2 | 5552 | 19 |
| Etowah | State Republican Executive Committee, Etowah County, |  | REP | Unverified Candidate 1 | 6059 | 21 |
| Etowah | State Republican Executive Committee, Etowah County, |  | REP | Unverified Candidate 1 | 8058 | 21 |
| Etowah | State Republican Executive Committee, Etowah County, |  | REP | Unverified Candidate 2 | 3641 | 21 |
| Etowah | State Republican Executive Committee, Etowah County, |  | REP | Unverified Candidate 2 | 2055 | 21 |
| Etowah | U.S. House | 3 | REP | Unverified Candidate 2 | 9361 | 21 |
| Jackson | School Board |  | REP | Unverified Candidate 1 | 1898 | 32 |
| Jackson | School Board |  | REP | Unverified Candidate 2 | 2571 | 32 |
| Lawrence | Lawrence County Sheriff |  | REP | Unverified Candidate 2 | 154 | 4 |
| Lawrence | Lawrence County Sheriff |  | REP | Unverified Candidate 3 | 2870 | 27 |
| Lawrence | Member, Lawrence County Gop Executive Committee, District No. 3 |  | REP | Unverified Candidate 1 | 525 | 9 |
| Lawrence | Member, Lawrence County Gop Executive Committee, District No. 3 |  | REP | Unverified Candidate 2 | 362 | 9 |
| Lawrence | State Republican Executive Committee, Lawrence Count |  | REP | Unverified Candidate 1 | 1700 | 31 |
| Lawrence | State Republican Executive Committee, Lawrence Count |  | REP | Unverified Candidate 2 | 2576 | 20 |
| Lawrence | State Republican Executive Committee, Lawrence Count |  | REP | Unverified Candidate 3 | 626 | 11 |
| Lee | Commissidner of Agriculture and Industries |  | REP | Unverified Candidate 1 | 3875 | 18 |
| Lee | Commissidner of Agriculture and Industries |  | REP | Unverified Candidate 2 | 3954 | 18 |
| Lee | Commissidner of Agriculture and Industries |  | REP | Unverified Candidate 3 | 2791 | 18 |
| Lee | District No. B3 |  | DEM | Unverified Candidate 1 | 1415 | 10 |
| Lee | District No. B3 |  | DEM | Unverified Candidate 2 | 685 | 10 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 1 | 314 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 2 | 296 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 3 | 193 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 4 | 286 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 5 | 448 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 6 | 912 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 7 | 197 | 16 |
| Lee | Lee County Democratic Executive Committee (Female). District No. 4 |  | DEM | Unverified Candidate 8 | 554 | 16 |
| Lee | Lieutenant Governor |  | REP | Unverified Candidate 7 | 6189 | 27 |
| Lee | Member, Lee County Commission, District No. 5 |  | DEM | Unverified Candidate 1 | 1182 | 15 |
| Lee | Member, Lee County Commission, District No. 5 |  | DEM | Unverified Candidate 2 | 1786 | 15 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 1 | 2953 | 27 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 2 | 1139 | 27 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 3 | 2581 | 27 |
| Lee | U.S. Senate |  | DEM | Unverified Candidate 4 | 1406 | 27 |
| Limestone | Attorney General |  | REP | Unverified Candidate 4 | 1 | 28 |
| Limestone | Attorney General |  | REP | Unverified Candidate 5 | 722 | 28 |
| Limestone | Chairman, Limestone County Commission |  | REP | Unverified Candidate 4 | 0 | 17 |
| Limestone | Chairman, Limestone County Commission |  | REP | Unverified Candidate 5 | 337 | 17 |
| Limestone | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 4 | 1 | 28 |
| Limestone | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 5 | 1540 | 28 |
| Limestone | Governor |  | REP | Unverified Candidate 4 | 3 | 28 |
| Limestone | Governor |  | REP | Unverified Candidate 5 | 342 | 28 |
| Limestone | Lieutenant Governor |  | REP | Unverified Candidate 8 | 4 | 28 |
| Limestone | Lieutenant Governor |  | REP | Unverified Candidate 9 | 431 | 28 |
| Limestone | Member, Limestone County Commission, District No 1 |  | REP | Unverified Candidate 4 | 0 | 12 |
| Limestone | Member, Limestone County Commission, District No 1 |  | REP | Unverified Candidate 5 | 349 | 12 |
| Limestone | Proposed Statewide Amendment 1 |  |  | Unverified Candidate 1 | 14147 | 28 |
| Limestone | Proposed Statewide Amendment 1 |  |  | Unverified Candidate 2 | 2612 | 28 |
| Limestone | Proposed Statewide Amendment 1 |  |  | Unverified Candidate 3 | 13 | 28 |
| Limestone | Proposed Statewide Amendment 1 |  |  | Unverified Candidate 4 | 565 | 28 |
| Limestone | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 1 | 8760 | 28 |
| Limestone | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 2 | 7525 | 28 |
| Limestone | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 3 | 3 | 28 |
| Limestone | Proposed Statewide Amendment 2 |  |  | Unverified Candidate 4 | 1049 | 28 |
| Limestone | Public Service Commission, Place No 1 |  | REP | Unverified Candidate 3 | 1 | 28 |
| Limestone | Public Service Commission, Place No 1 |  | REP | Unverified Candidate 4 | 1783 | 28 |
| Limestone | Public Service Commission, Place No 2 |  | REP | Unverified Candidate 5 | 2 | 28 |
| Limestone | Public Service Commission, Place No 2 |  | REP | Unverified Candidate 6 | 1695 | 28 |
| Limestone | Secretary of State |  | REP | Unverified Candidate 4 | 0 | 28 |
| Limestone | Secretary of State |  | REP | Unverified Candidate 5 | 1514 | 28 |
| Limestone | State Auditor |  | REP | Unverified Candidate 3 | 0 | 28 |
| Limestone | State Auditor |  | REP | Unverified Candidate 4 | 1230 | 28 |
| Limestone | State Board of Education | 8 | REP | Unverified Candidate 4 | 1 | 24 |
| Limestone | State Board of Education | 8 | REP | Unverified Candidate 5 | 1363 | 24 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 1 | 675 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 2 | 1481 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 3 | 521 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 4 | 1462 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 5 | 5255 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 6 | 4 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 1 |  | REP | Unverified Candidate 7 | 2024 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 2 |  | REP | Unverified Candidate 1 | 1529 | 26 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 2 |  | REP | Unverified Candidate 2 | 4304 | 26 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 2 |  | REP | Unverified Candidate 3 | 1589 | 26 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 2 |  | REP | Unverified Candidate 4 | 2 | 26 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 2 |  | REP | Unverified Candidate 5 | 2296 | 26 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 3 |  | REP | Unverified Candidate 1 | 7838 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 3 |  | REP | Unverified Candidate 2 | 1447 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 3 |  | REP | Unverified Candidate 3 | 0 | 2 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 3 |  | REP | Unverified Candidate 4 | 295 | 28 |
| Limestone | State Republican Exec Comm, Limestone Co - Pl No 3 |  | REP | Unverified Candidate 5 | 1842 | 26 |
| Limestone | State Republican Exec Comm. Limestone Co - Pl No 4 |  | REP | Unverified Candidate 1 | 3160 | 25 |
| Limestone | State Republican Exec Comm. Limestone Co - Pl No 4 |  | REP | Unverified Candidate 2 | 3661 | 25 |
| Limestone | State Republican Exec Comm. Limestone Co - Pl No 4 |  | REP | Unverified Candidate 3 | 0 | 25 |
| Limestone | State Republican Exec Comm. Limestone Co - Pl No 4 |  | REP | Unverified Candidate 4 | 2510 | 25 |
| Limestone | State Treasurer |  | REP | Unverified Candidate 3 | 1 | 28 |
| Limestone | State Treasurer |  | REP | Unverified Candidate 4 | 1283 | 28 |
| Limestone | U.S. Senate |  | REP | Unverified Candidate 2 | 295 | 28 |
| Limestone | U.S. Senate |  | REP | Unverified Candidate 8 | 8 | 28 |
| Limestone | U.S. Senate |  | REP | Unverified Candidate 9 | 481 | 28 |
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
| Marshall | County Sch |  | REP | Unverified Candidate 1 | 3337 | 31 |
| Marshall | County Sch |  | REP | Unverified Candidate 1 | 2873 | 31 |
| Marshall | County Sch |  | REP | Unverified Candidate 2 | 2610 | 31 |
| Marshall | County Sch |  | REP | Unverified Candidate 2 | 2613 | 31 |
| Marshall | County Sch |  | REP | Unverified Candidate 3 | 1221 | 31 |
| Marshall | County Sch |  | REP | Unverified Candidate 3 | 699 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 1 | 3789 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 1 | 8908 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 1 | 4143 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 1 | 3501 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 2 | 6586 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 2 | 2230 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 2 | 5481 | 31 |
| Marshall | State Republican Executive Committee, Marshall Count |  | REP | Unverified Candidate 2 | 6464 | 31 |
| Monroe | Monroe County Sheriff |  | REP | Unverified Candidate 2 | 1697 | 29 |
| Pickens | Member, Pickens County Commission, District No. 1 |  | REP | Unverified Candidate 1 | 246 | 9 |
| Pickens | Member, Pickens County Commission, District No. 1 |  | REP | Unverified Candidate 2 | 418 | 9 |
| Pickens | Member, Pickens County Commission, District No. 3 |  | REP | Unverified Candidate 1 | 472 | 6 |
| Pickens | Member, Pickens County Commission, District No. 3 |  | REP | Unverified Candidate 2 | 495 | 6 |
| Pickens | Member, Pickens County Commission, District No. 4 |  | DEM | Unverified Candidate 1 | 253 | 8 |
| Pickens | Member, Pickens County Commission, District No. 4 |  | DEM | Unverified Candidate 2 | 371 | 8 |
| Pickens | Member, Pickens County Commission, District No. 5 |  | REP | Unverified Candidate 1 | 377 | 7 |
| Pickens | Member, Pickens County Commission, District No. 5 |  | REP | Unverified Candidate 2 | 316 | 7 |
| Pike | Governor |  | DEM | Unverified Candidate 2 | 263 | 36 |
| Pike | Governor |  | DEM | Unverified Candidate 3 | 169 | 36 |
| Pike | Governor |  | DEM | Unverified Candidate 4 | 1711 | 36 |
| Pike | Governor |  | DEM | Unverified Candidate 5 | 45 | 36 |
| Pike | Governor |  | DEM | Unverified Candidate 6 | 26 | 36 |
| Russell | Lieutenant Governor |  | REP | Unverified Candidate 2 | 118 | 20 |
| Russell | Lieutenant Governor |  | REP | Unverified Candidate 3 | 50 | 20 |
| Russell | Lieutenant Governor |  | REP | Unverified Candidate 4 | 140 | 20 |
| Russell | Lieutenant Governor |  | REP | Unverified Candidate 5 | 38 | 20 |
| Russell | Lieutenant Governor |  | REP | Unverified Candidate 6 | 125 | 20 |
| Russell | Lieutenant Governor |  | REP | Unverified Candidate 7 | 1063 | 20 |
| Russell | Secretary of State |  | REP | Unverified Candidate 2 | 446 | 20 |
| Russell | State Senate | 27 | REP | Unverified Candidate 2 | 543 | 7 |
| St. Clair | Attorney General |  | REP | Unverified Candidate 2 | 3382 | 32 |
| St. Clair | County (All) |  | REP | Unverified Candidate 1 | 4275 | 25 |
| St. Clair | County (All) |  | REP | Unverified Candidate 2 | 2321 | 25 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 1 | 4751 | 32 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 1 | 3792 | 32 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 1 | 6643 | 32 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 2 | 2850 | 32 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 2 | 6037 | 32 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 2 | 2842 | 32 |
| St. Clair | State Republican Executive Committee, St.Clair Count |  | REP | Unverified Candidate 3 | 1859 | 32 |
| St. Clair | U.S. House | 3 | REP | Unverified Candidate 3 | 2586 | 7 |
| Washington | Attorney General |  | REP | Unverified Candidate 2 | 1060 | 21 |
| Washington | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 4 | 414 | 12 |
| Washington | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 5 | 1110 | 12 |
| Washington | Governor |  | REP | Unverified Candidate 2 | 189 | 21 |
| Washington | Governor |  | REP | Unverified Candidate 3 | 3129 | 21 |
| Washington | Governor |  | DEM | Unverified Candidate 6 | 11 | 20 |
| Washington | Lieutenant Governor |  | REP | Unverified Candidate 8 | 466 | 8 |
| Washington | State Auditor |  | REP | Unverified Candidate 2 | 2197 | 21 |
| Washington | State House | 65 | REP | Unverified Candidate 2 | 1351 | 11 |
| Washington | State House | 65 | REP | Unverified Candidate 3 | 1078 | 10 |
| Winston | U.S. Senate |  | DEM | Unverified Candidate 4 | 45 | 20 |
| Houston | Member, Houston County Commission, District No. 1 |  | DEM | Unverified Candidate 1 | 693 | 11 |
| Houston | Member, Houston County Commission, District No. 1 |  | DEM | Unverified Candidate 2 | 986 | 11 |
| Houston | State Democratic Executive Committee |  | DEM | Unverified Candidate 2 | 805 | 9 |
| Lauderdale | Attorney General |  | REP | Unverified Candidate 3 | 3147 | 33 |
| Lauderdale | Commissioner of Agriculture Ano Industries |  | REP | Unverified Candidate 1 | 3087 | 33 |
| Lauderdale | Commissioner of Agriculture Ano Industries |  | REP | Unverified Candidate 2 | 2658 | 33 |
| Lauderdale | Commissioner of Agriculture Ano Industries |  | REP | Unverified Candidate 3 | 2168 | 33 |
| Lauderdale | Lauderdale Cdunty Republican Executive Committee, Pl |  | REP | Unverified Candidate 1 | 2022 | 15 |
| Lauderdale | Lauderdale Cdunty Republican Executive Committee, Pl |  | REP | Unverified Candidate 2 | 795 | 15 |
| Lauderdale | Lauderdale County Coroner |  | REP | Unverified Candidate 2 | 2431 | 11 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 1 | 9654 | 33 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 2 | 7 | 1 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 3 | 3575 | 33 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 4 | 3193 | 25 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl District No. 1 |  | REP | Unverified Candidate 1 | 6318 | 18 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl District No. 1 |  | REP | Unverified Candidate 2 | 5750 | 18 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl District No. 2 |  | REP | Unverified Candidate 1 | 3426 | 22 |
| Lauderdale | Lauderdale County Republican Executive Committee, Pl District No. 2 |  | REP | Unverified Candidate 2 | 1990 | 22 |
| Lauderdale | Lauoerdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 1 | 3966 | 33 |
| Lauderdale | Lauoerdale County Republican Executive Committee, Pl |  | REP | Unverified Candidate 2 | 3745 | 33 |
| Lauderdale | Public Service Commission, Place No. 1 |  | DEM | Unverified Candidate 2 | 644 | 33 |
| Lauderdale | State Republican Executive Committee, Lauderdale Cou |  | REP | Unverified Candidate 1 | 16051 | 33 |
| Lauderdale | State Republican Executive Committee, Lauderdale Cou |  | REP | Unverified Candidate 2 | 11607 | 33 |
| Lauderdale | State Republican Executive Committee, Lauderdale Cou |  | REP | Unverified Candidate 3 | 2950 | 16 |
| Madison | Districts 4 & 6, Place No. 2 |  | REP | Unverified Candidate 1 | 2207 | 31 |
| Madison | Districts 4 & 6, Place No. 2 |  | REP | Unverified Candidate 2 | 2473 | 31 |
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
| Madison | Madison County Republican Executive Committee, District No. 3 |  | REP | Unverified Candidate 1 | 3053 | 17 |
| Madison | Madison County Republican Executive Committee, District No. 3 |  | REP | Unverified Candidate 2 | 2208 | 17 |
| Madison | State Democratic Executive Committee |  | DEM | Unverified Candidate 2 | 1265 | 16 |
| Madison | State House | 25 | DEM | Unverified Candidate 3 | 70 | 1 |
| Madison | State Republican Executive Committee, District No. 1 |  | REP | Unverified Candidate 1 | 1345 | 15 |
| Madison | State Republican Executive Committee, District No. 1 |  | REP | Unverified Candidate 2 | 2544 | 15 |
| Madison | State Republican Executive Committee, District No. 2 |  | REP | Unverified Candidate 1 | 1066 | 14 |
| Madison | State Republican Executive Committee, District No. 2 |  | REP | Unverified Candidate 2 | 1127 | 14 |
| Madison | State Republican Executive Committee, District No. 2 |  | REP | Unverified Candidate 3 | 1053 | 14 |
| Madison | State Republican Executive Committee, District No. 3 |  | REP | Unverified Candidate 1 | 2182 | 17 |
| Madison | State Republican Executive Committee, District No. 3 |  | REP | Unverified Candidate 2 | 3077 | 17 |
| Madison | State Senate | 2 | DEM | Unverified Candidate 4 | 102 | 1 |
| Shelby | Commissioner of Agriculture and Industries |  | REP | Unverified Candidate 2 | 8210 | 40 |
| Shelby | Member, Shelby County Board of Education, Place No. District No. 2 |  | REP | Unverified Candidate 1 | 6649 | 39 |
| Shelby | Member, Shelby County Board of Education, Place No. District No. 2 |  | REP | Unverified Candidate 2 | 7308 | 39 |
| Shelby | Shelby County Sheriff |  | REP | Unverified Candidate 1 | 7072 | 40 |
| Shelby | Shelby County Sheriff |  | REP | Unverified Candidate 2 | 17070 | 40 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 1 | 38585 | 40 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 2 | 47116 | 40 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 3 | 23384 | 40 |
| Shelby | State Republican Executive Committee, Shelby County, |  | REP | Unverified Candidate 4 | 50 | 1 |
| Shelby | Superintendent, Shelby County Board of Education District No. 2 |  | REP | Unverified Candidate 1 | 7152 | 39 |
| Shelby | Superintendent, Shelby County Board of Education District No. 2 |  | REP | Unverified Candidate 2 | 7517 | 39 |

**380 candidate columns** across 146 county/office combinations.

> **Note:** Houston, Lauderdale, Madison, and Shelby (the four counties added since the last revision) were updated from the CSV output alone, so — like Baldwin, Chambers, Cullman, and DeKalb above — their unresolved-candidate rows use the merged-column convention: when several contest columns in the PDF share one office label, they are summed into a single row rather than split into separate columns. Perry's two former `Member, Perry County Commission, District No. 3` rows were removed: the current Perry CSV contains no such office and no `Unverified Candidate` entries, so those placeholders have been resolved.
