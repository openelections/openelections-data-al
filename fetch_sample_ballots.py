#!/usr/bin/env python3
"""
Fetch both parties' 2026 AL primary sample ballot PDFs for a county from the
Alabama Secretary of State.

Source page: https://www.sos.alabama.gov/alabama-votes/2026-primary-election-sample-ballots

The county -> filename mapping below was scraped directly from that page's
live DOM (not derived from a naming convention), because the filenames are
NOT uniform — a few counties break the usual "<County> - <Party>.pdf" pattern
(missing spaces around the dash, or a dropped period):
    Geneva-Rep.pdf, Russell-Dem.pdf, Shelby -Rep.pdf, "St Clair" (no period).
Deriving URLs from county names programmatically would silently 404 on these
four. Refresh the table by re-running this against the live page if the SoS
ever re-publishes ballots under different names:
    Array.from(document.querySelectorAll('a[href*="sample-ballots"]'))
        .map(a => a.getAttribute('href'))

Downloads go through `curl`, not Python's requests/urllib3: on this machine
Python's SSL context can't verify sos.alabama.gov's certificate chain (a
missing intermediate in the certifi bundle), while curl — using the macOS
system keychain — verifies it fine. `requests` and even the Claude web-fetch
tool both fail with a certificate-verify error against this exact host; curl
does not.

Usage:
    python3 fetch_sample_ballots.py <county> [<county> ...]
    python3 fetch_sample_ballots.py --all
    python3 fetch_sample_ballots.py --list          # show county names this script knows

As a library:
    from fetch_sample_ballots import fetch_ballots
    dem_path, rep_path = fetch_ballots("Lawrence")   # downloads if not cached, else instant
"""

import argparse
import os
import subprocess
import sys
import urllib.parse

BASE_URL = "https://www.sos.alabama.gov/sites/default/files/sample-ballots/2026/pri"
CACHE_DIR = os.environ.get("SAMPLE_BALLOT_CACHE", ".sample_ballots")

# county name (matching 2026/20260519__al__primary__county.csv's spelling) ->
# (dem_filename, rep_filename) as they actually appear on the SoS site.
BALLOT_FILES = {
    'Autauga': ('Autauga - Dem.pdf', 'Autauga - Rep.pdf'),
    'Baldwin': ('Baldwin - Dem.pdf', 'Baldwin - Rep.pdf'),
    'Barbour': ('Barbour - Dem.pdf', 'Barbour - Rep.pdf'),
    'Bibb': ('Bibb - Dem.pdf', 'Bibb - Rep.pdf'),
    'Blount': ('Blount - Dem.pdf', 'Blount - Rep.pdf'),
    'Bullock': ('Bullock - Dem.pdf', 'Bullock - Rep.pdf'),
    'Butler': ('Butler - Dem.pdf', 'Butler - Rep.pdf'),
    'Calhoun': ('Calhoun - Dem.pdf', 'Calhoun - Rep.pdf'),
    'Chambers': ('Chambers - Dem.pdf', 'Chambers - Rep.pdf'),
    'Cherokee': ('Cherokee - Dem.pdf', 'Cherokee - Rep.pdf'),
    'Chilton': ('Chilton - Dem.pdf', 'Chilton - Rep.pdf'),
    'Choctaw': ('Choctaw - Dem.pdf', 'Choctaw - Rep.pdf'),
    'Clarke': ('Clarke - Dem.pdf', 'Clarke - Rep.pdf'),
    'Clay': ('Clay - Dem.pdf', 'Clay - Rep.pdf'),
    'Cleburne': ('Cleburne - Dem.pdf', 'Cleburne - Rep.pdf'),
    'Coffee': ('Coffee - Dem.pdf', 'Coffee - Rep.pdf'),
    'Colbert': ('Colbert - Dem.pdf', 'Colbert - Rep.pdf'),
    'Conecuh': ('Conecuh - Dem.pdf', 'Conecuh - Rep.pdf'),
    'Coosa': ('Coosa - Dem.pdf', 'Coosa - Rep.pdf'),
    'Covington': ('Covington - Dem.pdf', 'Covington - Rep.pdf'),
    'Crenshaw': ('Crenshaw - Dem.pdf', 'Crenshaw - Rep.pdf'),
    'Cullman': ('Cullman - Dem.pdf', 'Cullman - Rep.pdf'),
    'Dale': ('Dale - Dem.pdf', 'Dale - Rep.pdf'),
    'Dallas': ('Dallas - Dem.pdf', 'Dallas - Rep.pdf'),
    'DeKalb': ('DeKalb - Dem.pdf', 'DeKalb - Rep.pdf'),
    'Elmore': ('Elmore - Dem.pdf', 'Elmore - Rep.pdf'),
    'Escambia': ('Escambia - Dem.pdf', 'Escambia - Rep.pdf'),
    'Etowah': ('Etowah - Dem.pdf', 'Etowah - Rep.pdf'),
    'Fayette': ('Fayette - Dem.pdf', 'Fayette - Rep.pdf'),
    'Franklin': ('Franklin - Dem.pdf', 'Franklin - Rep.pdf'),
    'Geneva': ('Geneva - Dem.pdf', 'Geneva-Rep.pdf'),
    'Greene': ('Greene - Dem.pdf', 'Greene - Rep.pdf'),
    'Hale': ('Hale - Dem.pdf', 'Hale - Rep.pdf'),
    'Henry': ('Henry - Dem.pdf', 'Henry - Rep.pdf'),
    'Houston': ('Houston - Dem.pdf', 'Houston - Rep.pdf'),
    'Jackson': ('Jackson - Dem.pdf', 'Jackson - Rep.pdf'),
    'Jefferson': ('Jefferson - Dem.pdf', 'Jefferson - Rep.pdf'),
    'Lamar': ('Lamar - Dem.pdf', 'Lamar - Rep.pdf'),
    'Lauderdale': ('Lauderdale - Dem.pdf', 'Lauderdale - Rep.pdf'),
    'Lawrence': ('Lawrence - Dem.pdf', 'Lawrence - Rep.pdf'),
    'Lee': ('Lee - Dem.pdf', 'Lee - Rep.pdf'),
    'Limestone': ('Limestone - Dem.pdf', 'Limestone - Rep.pdf'),
    'Lowndes': ('Lowndes - Dem.pdf', 'Lowndes - Rep.pdf'),
    'Macon': ('Macon - Dem.pdf', 'Macon - Rep.pdf'),
    'Madison': ('Madison - Dem.pdf', 'Madison - Rep.pdf'),
    'Marengo': ('Marengo - Dem.pdf', 'Marengo - Rep.pdf'),
    'Marion': ('Marion - Dem.pdf', 'Marion - Rep.pdf'),
    'Marshall': ('Marshall - Dem.pdf', 'Marshall - Rep.pdf'),
    'Mobile': ('Mobile - Dem.pdf', 'Mobile - Rep.pdf'),
    'Monroe': ('Monroe - Dem.pdf', 'Monroe - Rep.pdf'),
    'Montgomery': ('Montgomery - Dem.pdf', 'Montgomery - Rep.pdf'),
    'Morgan': ('Morgan - Dem.pdf', 'Morgan - Rep.pdf'),
    'Perry': ('Perry - Dem.pdf', 'Perry - Rep.pdf'),
    'Pickens': ('Pickens - Dem.pdf', 'Pickens - Rep.pdf'),
    'Pike': ('Pike - Dem.pdf', 'Pike - Rep.pdf'),
    'Randolph': ('Randolph - Dem.pdf', 'Randolph - Rep.pdf'),
    'Russell': ('Russell-Dem.pdf', 'Russell - Rep.pdf'),
    'Shelby': ('Shelby - Dem.pdf', 'Shelby -Rep.pdf'),
    'St. Clair': ('St Clair - Dem.pdf', 'St Clair - Rep.pdf'),
    'Sumter': ('Sumter - Dem.pdf', 'Sumter - Rep.pdf'),
    'Talladega': ('Talladega - Dem.pdf', 'Talladega - Rep.pdf'),
    'Tallapoosa': ('Tallapoosa - Dem.pdf', 'Tallapoosa - Rep.pdf'),
    'Tuscaloosa': ('Tuscaloosa - Dem.pdf', 'Tuscaloosa - Rep.pdf'),
    'Walker': ('Walker - Dem.pdf', 'Walker - Rep.pdf'),
    'Washington': ('Washington - Dem.pdf', 'Washington - Rep.pdf'),
    'Wilcox': ('Wilcox - Dem.pdf', 'Wilcox - Rep.pdf'),
    'Winston': ('Winston - Dem.pdf', 'Winston - Rep.pdf'),
}

_NORM = {k.lower().replace('.', '').replace(' ', ''): k for k in BALLOT_FILES}


def resolve_county(name):
    """Case/punctuation-tolerant lookup ('st clair', 'St. Clair', 'saint clair'
    all fail cleanly except the two real spellings) -> canonical county name."""
    key = name.lower().replace('.', '').replace(' ', '')
    if key in _NORM:
        return _NORM[key]
    raise KeyError(f"unknown county {name!r}; known: {sorted(BALLOT_FILES)}")


def _download(filename, dest_path):
    url = BASE_URL + "/" + urllib.parse.quote(filename)
    # curl, not requests: see module docstring for why.
    result = subprocess.run(
        ["curl", "-sL", "-f", "--retry", "2", url, "-o", dest_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0 or not os.path.exists(dest_path) or os.path.getsize(dest_path) == 0:
        if os.path.exists(dest_path):
            os.remove(dest_path)
        raise RuntimeError(
            f"failed to fetch {url!r} (curl exit {result.returncode}): {result.stderr.strip()}"
        )
    return dest_path


def fetch_ballots(county, cache_dir=None, force=False):
    """Return (dem_path, rep_path) for a county, downloading into cache_dir
    (default CACHE_DIR / .sample_ballots) only if not already cached."""
    county = resolve_county(county)
    cache_dir = cache_dir or CACHE_DIR
    county_dir = os.path.join(cache_dir, county.replace('.', '').replace(' ', '_'))
    os.makedirs(county_dir, exist_ok=True)

    dem_fname, rep_fname = BALLOT_FILES[county]
    paths = []
    for fname in (dem_fname, rep_fname):
        dest = os.path.join(county_dir, fname)
        if force or not os.path.exists(dest):
            _download(fname, dest)
        paths.append(dest)
    return tuple(paths)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("counties", nargs="*", help="county name(s)")
    ap.add_argument("--all", action="store_true", help="fetch every county")
    ap.add_argument("--list", action="store_true", help="list known county names and exit")
    ap.add_argument("--force", action="store_true", help="re-download even if cached")
    args = ap.parse_args()

    if args.list:
        for c in sorted(BALLOT_FILES):
            print(c)
        return 0

    targets = sorted(BALLOT_FILES) if args.all else args.counties
    if not targets:
        ap.print_help()
        return 1

    failures = []
    for county in targets:
        try:
            dem_path, rep_path = fetch_ballots(county, force=args.force)
            print(f"{county}:")
            print(f"  DEM -> {dem_path}")
            print(f"  REP -> {rep_path}")
        except (KeyError, RuntimeError) as e:
            print(f"{county}: FAILED — {e}", file=sys.stderr)
            failures.append(county)

    if failures:
        print(f"\n{len(failures)} failed: {failures}", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
