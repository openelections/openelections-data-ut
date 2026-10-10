#!/usr/bin/env python3
"""
Parse the Utah 2024 primary election JSON export into a precinct-level CSV.

Source: export-primary06252024.json (Utah Lieutenant Governor results export),
which has the same shape as the general-election export:
    localResults -> counties -> ballotItems -> ballotOptions -> precinctResults

Only `localResults` is processed. The top-level `results` key also carries
precinct-level data for statewide contests, but those same precinct results are
already present per-county inside `localResults`, so processing both would
duplicate statewide contests.

Primary contest names are party-prefixed and inconsistically cased
("Republican For US Senate", "REP County Commission Seat C",
"COUNTY COUNCIL 2 (REP)", "DAVIS SCHOOL BOARD DISTRICT 1"), so this script
normalizes them before delegating office/district extraction to json_parser.

Run with uv:

    uv run python parsers/parse_primary.py \\
        "/path/to/export-primary06252024.json" \\
        2024/20240625__ut__primary__precinct.csv

Output columns: county,precinct,office,district,party,candidate,votes
"""

import csv
import json
import os
import re
import sys

# json_parser lives in the same directory; running this script by path puts
# parsers/ on sys.path[0], so this import works under `uv run python parsers/...`.
from json_parser import (
    parse_office_district,
    parse_candidate_party,
    _aggregate_write_ins,
)

# --- primary-specific contest-name normalization ----------------------------

# Leading "Republican For " / "Republican for " / "Democratic For " / etc.
_LEADING_PARTY = re.compile(
    r'^(?:Republican|Democratic|Democrat|Libertarian|Independent American|'
    r'Independent|Constitution|Green|United Utah)\s+[Ff]or\s+'
)
# Leading bare party code: "REP County Commission Seat C"
_LEADING_PARTY_CODE = re.compile(r'^(?:REP|DEM|IAP|LIB|CON|GRN|UUP|IND)\s+')
# Trailing party code: "COUNTY COUNCIL 2 (REP)"
_TRAILING_PARTY_CODE = re.compile(
    r'\s*\((?:REP|DEM|IAP|LIB|CON|GRN|UUP|IND)\)\s*$'
)
# The word "District"/"Districts", stripped once a district number is extracted.
_DISTRICT_WORD = re.compile(r'\bDistricts?\b', re.IGNORECASE)
# Separators left behind after stripping words/digits.
_SEPARATOR = re.compile(r'\s*[-–—]\s*')
_TRAILING_HASH = re.compile(r'#\s*$')
# Seat designators like "#4" (as opposed to geographic "District 4"): the number
# is a seat number, so it stays in the office name and district is left empty.
_SEAT_HASH = re.compile(r'#\s*(\d+)')

# Offices whose names json_parser's office_map standardizes; these are kept
# verbatim. Everything else (county/local offices) is title-cased.
STANDARD_OFFICES = {
    "President", "U.S. House", "U.S. Senate", "Governor",
    "Attorney General", "State Auditor", "State Treasurer",
    "State Senate", "State House", "Registered Voters", "Ballots Cast",
}


def normalize_primary_contest(name):
    """Strip party designators and standardize federal office tokens."""
    name = (name or '').strip()
    name = _LEADING_PARTY.sub('', name)
    name = _LEADING_PARTY_CODE.sub('', name)
    name = _TRAILING_PARTY_CODE.sub('', name)
    name = name.replace('US Senate', 'U.S. Senate').replace('US House', 'U.S. House')
    return name.strip()


def parse_primary_office(contest_name, contest_type):
    """
    Return (office, district) for a primary contest name.

    Delegates digit extraction and standard-office mapping to
    json_parser.parse_office_district, then strips a trailing "District" word
    when a district number was pulled out (so "Davis School Board District 1"
    -> office "Davis School Board", district "1") and title-cases county/local
    offices that aren't part of the standard map.
    """
    name = normalize_primary_contest(contest_name)

    # Seat designators ("#4") are seats, not geographic districts: keep the
    # number in the office name ("Seat #4" -> "Seat 4") and leave district empty.
    if _SEAT_HASH.search(name):
        office = _SEAT_HASH.sub(r'\1', name)
        office = _SEPARATOR.sub(' ', office)
        office = ' '.join(office.split())
        if office not in STANDARD_OFFICES:
            office = office.title()
        return office, ''

    office, district = parse_office_district(name, contest_type)

    if district:
        office = _DISTRICT_WORD.sub('', office)
    office = _SEPARATOR.sub(' ', office)
    office = _TRAILING_HASH.sub('', office)
    office = ' '.join(office.split())

    if office not in STANDARD_OFFICES:
        office = office.title()
    return office, district


def normalize_party_field(political_party):
    """
    The primary export uses full party names with mixed case and a trailing
    " PARTY" suffix ("REPUBLICAN PARTY ", "Republican", "Nonpartisan"), unlike
    the general export's clean codes. Normalize so json_parser's party_map
    matches: strip, drop " PARTY", uppercase -> "REPUBLICAN", "DEMOCRATIC", etc.
    """
    pp = (political_party or '').strip().upper()
    pp = pp.replace(' PARTY', '').strip()
    return pp


def process_county_data(county_data, county_name):
    """Process one county's ballotItems into a list of result rows."""
    if county_name.endswith(' County'):
        county_name = county_name[:-7]

    results = []
    for contest in county_data.get('ballotItems', []):
        contest_type = contest.get('contestType', 'Candidate')
        office, district = parse_primary_office(contest.get('name', ''), contest_type)

        for option in contest.get('ballotOptions', []):
            candidate, party = parse_candidate_party(
                option.get('name', ''), normalize_party_field(option.get('politicalParty', ''))
            )
            for precinct_result in option.get('precinctResults', []):
                precinct_name = precinct_result.get('name', '')
                if not precinct_name:
                    continue
                vote_count = precinct_result.get('voteCount', 0)
                if vote_count is None or vote_count == '':
                    vote_count = ''
                results.append({
                    'county': county_name,
                    'precinct': precinct_name,
                    'office': office,
                    'district': district if district else '',
                    'party': party if party else '',
                    'candidate': candidate,
                    'votes': vote_count,
                })

    return results


def process_json_file(json_file_path, output_csv_path):
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    counties = data.get('localResults', [])
    print(f"Found {len(counties)} counties in localResults")

    all_results = []
    for county_data in counties:
        county_name = county_data.get('name', '')
        print(f"  Processing {county_name}...")
        all_results.extend(process_county_data(county_data, county_name))

    all_results = _aggregate_write_ins(all_results)

    fieldnames = ['county', 'precinct', 'office', 'district', 'party', 'candidate', 'votes']
    os.makedirs(os.path.dirname(output_csv_path) or '.', exist_ok=True)
    with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in all_results:
            writer.writerow(row)

    print(f"\nProcessed {len(all_results)} rows")
    print(f"Output written to: {output_csv_path}")
    return all_results


def print_office_summary(results):
    """Print the distinct (office, district) pairs for normalization review."""
    seen = {}
    for row in results:
        key = (row['office'], row['district'])
        seen[key] = seen.get(key, 0) + 1
    print(f"\n=== distinct (office, district) pairs: {len(seen)} ===")
    for (office, district), n in sorted(seen.items()):
        print(f"  {office!r:50} district={district!r:6}  ({n} rows)")


def main():
    if len(sys.argv) < 2:
        print("Usage: uv run python parsers/parse_primary.py <input.json> [output.csv]")
        sys.exit(1)
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) >= 3 else \
        input_file.replace('.json', '.csv')
    results = process_json_file(input_file, output_file)
    print_office_summary(results)


if __name__ == '__main__':
    main()