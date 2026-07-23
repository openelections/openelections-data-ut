#!/usr/bin/env python3
"""
Parser for Utah election results JSON data.
Converts JSON format from election results system to OpenElections CSV format.

This parser can handle:
1. Single-county JSON files (e.g., wasatch.json with county-level data)
2. Multi-county JSON files (e.g., export-general11052024.json with all counties)
3. Optional filtering to process only specific counties

Usage:
    # Process all counties from a multi-county file:
    python json_parser.py export-general11052024.json all_counties.csv
    
    # Process only Wasatch County:
    python json_parser.py export-general11052024.json wasatch.csv Wasatch
    
    # Process a single-county file:
    python json_parser.py wasatch.json wasatch_output.csv

Output CSV format:
    county,precinct,office,district,party,candidate,votes
"""

import json
import csv
import sys
import re


def parse_office_district(contest_name, contest_type):
    """
    Parse contest name to extract office and district.
    
    Args:
        contest_name: The name of the contest (e.g., "U.S. Senate", "U.S. House 3")
        contest_type: The type of contest (e.g., "Candidate", "BallotMeasure")
    
    Returns:
        tuple: (office, district)
    """
    # Clean up the contest name
    name = contest_name.strip()
    
    # Handle ballot measures and judicial retention
    if contest_type == "BallotMeasure":
        if "be retained" in name.lower():
            # Judicial retention
            return (name, None)
        else:
            # Other ballot measures
            return (name, None)
    
    # Standardize office names
    office_map = {
        "U.S. President": "President",
        "President": "President",
        "U.S. House": "U.S. House",
        "U.S. Senate": "U.S. Senate",
        "Governor": "Governor",
        "Attorney General": "Attorney General",
        "State Auditor": "State Auditor",
        "State Treasurer": "State Treasurer",
        "State Senate": "State Senate",
        "State House": "State House",
        "Registered Voters": "Registered Voters",
        "Ballots Cast": "Ballots Cast",
    }
    
    # Extract district from contest name if present
    # Pattern: "Office District #" or "Office # District"
    district_match = re.search(r'\b(\d+)\b', name)
    
    if district_match:
        district = district_match.group(1)
        # Remove the district number from the office name
        office = re.sub(r'\s*\b\d+\b\s*', ' ', name).strip()
    else:
        office = name
        district = None
    
    # Apply standardization
    for key, standard in office_map.items():
        if key.lower() in office.lower():
            office = standard
            break
    
    return (office, district)


def parse_candidate_party(candidate_name, political_party):
    """
    Parse candidate name and party.
    
    Args:
        candidate_name: The candidate's name
        political_party: The political party from the data
    
    Returns:
        tuple: (cleaned_candidate_name, party_code)
    """
    candidate = candidate_name.strip()
    
    # Map party names to codes
    party_map = {
        "REPUBLICAN": "REP",
        "DEMOCRATIC": "DEM",
        "DEMOCRAT": "DEM",
        "LIBERTARIAN": "LIB",
        "INDEPENDENT AMERICAN": "IAP",
        "INDEPENDENT": "IND",
        "CONSTITUTION": "CON",
        "UNITED UTAH": "UUP",
        "NONPARTISAN": None,
        "UNAFFILIATED": None,
        "GREEN": "GRN",
        "": None
    }
    
    party_upper = political_party.upper() if political_party else ""
    party = party_map.get(party_upper, political_party if political_party else None)
    
    # Handle write-ins
    if "WRITE-IN" in candidate.upper() or candidate.upper() == "WRITE-INS":
        candidate = "Write-ins"
        party = None
    
    # Clean up candidate names - remove extra spaces
    candidate = " ".join(candidate.split())
    
    return (candidate, party)


def process_county_data(county_data, county_name):
    """
    Process a single county's data and return results list.
    
    Args:
        county_data: Dictionary containing county ballot items
        county_name: Name of the county
    
    Returns:
        List of result dictionaries
    """
    # Strip " County" suffix to match expected format
    if county_name.endswith(' County'):
        county_name = county_name[:-7]
    
    results = []

    # Iterate through ballot items (contests)
    for contest in county_data.get('ballotItems', []):
        contest_name = contest.get('name', '')
        contest_type = contest.get('contestType', 'Candidate')

        office, district = parse_office_district(contest_name, contest_type)

        # Iterate through ballot options (candidates or choices)
        for option in contest.get('ballotOptions', []):
            candidate_name = option.get('name', '')
            political_party = option.get('politicalParty', '')

            candidate, party = parse_candidate_party(candidate_name, political_party)

            # Iterate through precinct results
            for precinct_result in option.get('precinctResults', []):
                precinct_name = precinct_result.get('name', '')
                vote_count = precinct_result.get('voteCount', 0)

                # Skip if precinct name is empty or vote count is None
                if not precinct_name:
                    continue

                # Handle empty vote counts
                if vote_count is None or vote_count == '':
                    vote_count = ''

                results.append({
                    'county': county_name,
                    'precinct': precinct_name,
                    'office': office,
                    'district': district if district else '',
                    'party': party if party else '',
                    'candidate': candidate,
                    'votes': vote_count
                })

    # A single contest may list multiple certified write-in candidates as
    # separate ballot options. parse_candidate_party collapses all of them to
    # the single label "Write-ins", which would otherwise emit one row per
    # write-in candidate per precinct -- duplicate rows that differ only in
    # votes. Aggregate those rows by summing their votes per precinct/contest so
    # each precinct has a single "Write-ins" total.
    return _aggregate_write_ins(results)


def _to_int(value):
    """Coerce a vote value to int for summing; blanks/non-numeric become 0."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _aggregate_write_ins(results):
    """
    Merge rows whose non-vote columns match and whose candidate is "Write-ins",
    summing their votes. Other rows are passed through unchanged so genuine
    duplicates are still surfaced by downstream tests.
    """
    merged = []
    write_in_index = {}
    for row in results:
        if row['candidate'] == 'Write-ins':
            key = (row['county'], row['precinct'], row['office'],
                   row['district'], row['party'], row['candidate'])
            pos = write_in_index.get(key)
            if pos is not None:
                merged[pos]['votes'] = _to_int(merged[pos]['votes']) + _to_int(row['votes'])
            else:
                write_in_index[key] = len(merged)
                row['votes'] = _to_int(row['votes'])
                merged.append(row)
        else:
            merged.append(row)
    return merged


def process_json_file(json_file_path, output_csv_path, county_filter=None):
    """
    Process JSON file and convert to CSV format.
    
    Args:
        json_file_path: Path to input JSON file
        output_csv_path: Path to output CSV file
        county_filter: Optional county name to filter (e.g., "Wasatch County"). 
                      If None, processes all counties.
    """
    # Read JSON data
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    all_results = []
    
    # Check if this is a single-county file (like wasatch.json) or multi-county file
    if 'name' in data and 'County' in data.get('name', ''):
        # Single county file
        county_name = data['name']
        if county_filter and county_filter.lower() not in county_name.lower():
            print(f"Skipping {county_name} (filter: {county_filter})")
            return
        
        print(f"Processing {county_name}...")
        all_results.extend(process_county_data(data, county_name))
    
    elif 'localResults' in data:
        # Multi-county file (like export-general11052024.json)
        counties = data['localResults']
        print(f"Found {len(counties)} counties in file")
        
        for county_data in counties:
            county_name = county_data.get('name', '')
            
            # Apply county filter if specified
            if county_filter:
                if county_filter.lower() not in county_name.lower():
                    continue
                print(f"Processing {county_name} (matched filter)...")
            else:
                print(f"Processing {county_name}...")
            
            all_results.extend(process_county_data(county_data, county_name))
    
    elif 'results' in data and 'localResults' not in data:
        # Try processing as if 'results' contains a single county
        county_name = data.get('results', {}).get('name', 'Unknown')
        if county_filter and county_filter.lower() not in county_name.lower():
            print(f"Skipping {county_name} (filter: {county_filter})")
            return
        
        print(f"Processing {county_name}...")
        all_results.extend(process_county_data(data['results'], county_name))
    
    else:
        print("Error: Unable to determine file structure")
        return
    
    # Write to CSV
    if all_results:
        with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['county', 'precinct', 'office', 'district', 'party', 'candidate', 'votes']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for row in all_results:
                writer.writerow(row)
        
        print(f"Processed {len(all_results)} rows")
        print(f"Output written to: {output_csv_path}")
    else:
        print("No results to write")


def main():
    """Main function to run the parser."""
    if len(sys.argv) < 2:
        print("Usage: python json_parser.py <input_json_file> [output_csv_file] [county_filter]")
        print("\nExamples:")
        print("  # Process all counties:")
        print("  python json_parser.py export-general11052024.json all_counties.csv")
        print("\n  # Process only Wasatch County:")
        print("  python json_parser.py export-general11052024.json wasatch.csv Wasatch")
        print("\n  # Process a single-county file:")
        print("  python json_parser.py wasatch.json wasatch_output.csv")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Default output file name based on input
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        output_file = input_file.replace('.json', '.csv')
    
    # Optional county filter
    county_filter = sys.argv[3] if len(sys.argv) >= 4 else None
    
    process_json_file(input_file, output_file, county_filter)


if __name__ == '__main__':
    main()
