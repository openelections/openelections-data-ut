"""Driver for the 2026 Utah primary parse pipeline.

Run summary or precinct CSVs for all 29 counties:

    uv run python3 parsers/run_2026.py summary            # all 29 county CSVs
    uv run python3 parsers/run_2026.py summary box_elder # one county
    uv run python3 parsers/run_2026.py precinct           # all 29 precinct CSVs
    uv run python3 parsers/run_2026.py precinct davis
    uv run python3 parsers/run_2026.py reconcile          # reconcile all
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import p2026_common as C

# Counties handled by the dedicated Salt Lake parser (distinct format).
# All others use the Electionware family parser.
SALT_LAKE = "salt_lake"

# Precinct format families not handled by the sequential Electionware parser:
#   - 2D SOVC table (precincts as rows, candidates as columns): davis, summit, uintah
#   - xlsx Statement of Votes Cast (Salt Lake): salt_lake
SOVC_2D = {"davis", "summit", "uintah"}

ALL_COUNTIES = list(C.COUNTY_DISPLAY.keys())


def _summary_parser(county):
    if county == SALT_LAKE:
        import p2026_saltlake as P
    else:
        import p2026_electionware as P
    return P


def _precinct_parser(county):
    if county == SALT_LAKE:
        import p2026_saltlake as P
    elif county in SOVC_2D:
        import p2026_sovc as P
    elif county == "box_elder":
        import p2026_boxelder as P
    else:
        import p2026_electionware as P
    return P


def run_summary(county):
    P = _summary_parser(county)
    src = C.find_source(county, "summary")
    if src is None:
        print(f"[{county}] no summary source found")
        return
    rows = P.parse_summary(county, src)
    ccsv, _ = C.county_paths(county)
    C.write_csv(rows, ccsv, precinct=False)
    issues = C.sanity_check(ccsv, precinct=False)
    n = len(rows)
    ocr = " (OCR)" if (C.needs_ocr(src) or county in C.FORCE_OCR_COUNTIES) else ""
    if issues:
        print(f"[{county}] summary{ocr}: {n} rows, {len(issues)} ISSUES:")
        for iss in issues[:20]:
            print(f"    {iss}")
        if len(issues) > 20:
            print(f"    ...and {len(issues)-20} more")
    else:
        print(f"[{county}] summary{ocr}: {n} rows, OK")


def run_precinct(county):
    P = _precinct_parser(county)
    src = C.find_source(county, "precinct")
    if src is None:
        print(f"[{county}] no precinct source found")
        return
    rows = P.parse_precinct(county, src)
    _, pcsv = C.county_paths(county)
    # Fix OCR space-loss in candidate names against the summary spelling
    # (e.g. 'Danielgardner' -> 'Daniel Gardner').
    ccsv, _ = C.county_paths(county)
    if ccsv.exists():
        rows, nfixed = C.reconcile_candidate_names(rows, ccsv)
        if nfixed:
            print(f"[{county}] name-reconcile: fixed {nfixed} candidate name(s)")
    C.write_csv(rows, pcsv, precinct=True)
    issues = C.sanity_check(pcsv, precinct=True)
    n = len(rows)
    ocr = " (OCR)" if (C.needs_ocr(src) or county in C.FORCE_OCR_COUNTIES) else ""
    if issues:
        print(f"[{county}] precinct{ocr}: {n} rows, {len(issues)} ISSUES:")
        for iss in issues[:25]:
            print(f"    {iss}")
        if len(issues) > 25:
            print(f"    ...and {len(issues)-25} more")
    else:
        print(f"[{county}] precinct{ocr}: {n} rows, OK")


def run_reconcile(county):
    ccsv, pcsv = C.county_paths(county)
    if not pcsv.exists():
        print(f"[{county}] reconcile: no precinct CSV")
        return
    C.reconcile(county, ccsv, pcsv)


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return
    kind = args[0]
    rest = args[1:]
    counties = rest if rest else ALL_COUNTIES
    if kind == "summary":
        for c in counties:
            try:
                run_summary(c)
            except Exception as e:
                import traceback
                print(f"[{c}] summary FAILED: {e}")
                traceback.print_exc()
    elif kind == "precinct":
        for c in counties:
            try:
                run_precinct(c)
            except Exception as e:
                import traceback
                print(f"[{c}] precinct FAILED: {e}")
                traceback.print_exc()
    elif kind == "reconcile":
        for c in counties:
            try:
                run_reconcile(c)
            except Exception as e:
                print(f"[{c}] reconcile FAILED: {e}")
    else:
        print(f"unknown kind {kind!r}")


if __name__ == "__main__":
    main()