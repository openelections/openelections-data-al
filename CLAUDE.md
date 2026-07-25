# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

This is an [OpenElections](https://openelections.net/) data repository for Alabama. It holds **pre-processed, precinct-level election results** as CSVs plus the Python scripts that download raw files from the Alabama Secretary of State and convert them into the normalized OpenElections format. The checked-in CSVs (in `2012/`, `2014/`, `2016/`, `2017/`, `2018/`, `2020/`) are the deliverable; the scripts regenerate them.

## The OpenElections CSV format (load-bearing convention)

Every output CSV uses exactly these columns and this ordering:

```
county,precinct,office,district,party,candidate,votes
```

- **Filenames** encode the election: `YYYYMMDD__al__<type>__precinct.csv` (or `...__al__<type>.csv` for county-level). `<type>` is `general`, `primary`, or `special` (special can combine, e.g. `20171212__al__special__general__precinct.csv`). The `src/verifier.py` parses this filename to derive state and county.
- **Offices** are normalized to a closed set (e.g. `President`, `U.S. Senate`, `U.S. House`, `Governor`, `State Senate`, `State House`). `convert_spreadsheets_to_csv.py` maps raw office strings via `office_map`; `src/verifier.py` enforces `validOffices`.
- **Districts** are integers for districted offices (`U.S. House`, `State Senate`, `State House`); blank otherwise.
- **Pseudocandidates** (`Write-ins`, `Over Votes`, `Under Votes`, `Total`, `Registered Voters`) are recognized as special non-candidate rows.

When changing conversion logic, keep output CSVs conformant — the CI tests and `src/verifier.py` validate against this format.

## Common commands

```bash
pip install -r requirements.txt

# Download + unzip raw SoS zip files into data/AL/ (data/ is gitignored)
python file_download_unzipper.py

# Convert one election's raw files into a normalized CSV (pass an election dir + output path)
python convert_spreadsheets_to_csv.py data/AL/<election-dir> <output.csv>

# Validate a produced CSV against the OpenElections format (schema, offices, districts, parties, uniqueness)
python src/verifier.py <path-to-csv>            # single file
python src/verifier.py 2016/*.csv --singleError # stop after first error per file

# Check that precinct/candidate "Total" rows equal the sum of component rows (checksum)
python src/total_checksum.py <path-to-csv>
python src/total_checksum.py <path-to-csv> --primary   # treat as primary (party per office)
```

`verifier.py` and `total_checksum.py` take `nargs='+'` paths, so you can pass globs/many files.

## CI / tests

There is no in-repo test suite. Validation runs in GitHub Actions (`.github/workflows/`) via the external repo `openelections/openelections-data-tests` (pinned at `v2.2.0`), running four checks against changed CSVs: `duplicate_entries`, `file_format`, `missing_values`, `vote_breakdown_totals`. The "Changed Files" workflow only runs against `*.csv` files added in a PR and posts a comment on failure. To run these locally you'd need to clone `openelections/openelections-data-tests` and invoke its `run_tests.py` the same way the workflows do.

## Architecture

Two-stage pipeline: **download → convert → verify**.

- `file_download_unzipper.py` — reads `alabama_general_precinct_files.csv` (election name + zip URL), downloads each to `data/AL/`, unzips. Note the 2004 file is an `.exe` that is really a zip; the script rewrites the extension. To add a new election, add a row to `alabama_general_precinct_files.csv`.
- `convert_spreadsheets_to_csv.py` — the substantive logic. `XLSProcessor` takes a directory of per-county files for one election and emits one statewide normalized CSV. Alabama's SoS files come in several heterogeneous formats, dispatched by inspecting the first cell of the spreadsheet:
  - `process_contest_title_excel_file` — "Contest Title" header format (2016 xls and some earlier).
  - `process_TOC_excel_file` — "Table of Contents" multi-sheet format (many 2014 files); only relevant sheets (President / U.S. House) are parsed via `relevant_sheets`.
  - `process_blank_header_excel_file` — blank/office-header format (many 2014 files); transposes, forward-fills offices, splits district from office and party from candidate. Contains a known Clay-2014 data hack.
  - `process_csv_file` — newer raw-CSV format with fixed column names.
  - All branches funnel through `populateOfficesAndDistricts` (splits district numbers off office strings) and `normalizeOfficesAndCandidates` (title-cases offices, applies `office_map`, drops non-statewide offices, normalizes pseudocandidates). `completeColumnNames` defines final column order.
- `src/verifier.py` — post-hoc validation of any output CSV. Uses `__new__` to dispatch to a subclass based on the filename (`General*`, `Primary*`, `Special*`, `*Precinct` vs. county-level). County-level verifiers relax the `precinct`-required constraint.
- `src/total_checksum.py` — recomputes candidate/precinct totals and compares to reported `Total` rows.

`requirements.txt` pins very old versions (pandas 0.19.2, numpy 1.12.0); note `pandas.DataFrame.append` (used in `process_TOC_excel_file`) was removed in pandas 2.0, so running conversion against modern pandas will require adjusting that branch.