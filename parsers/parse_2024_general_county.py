#!/usr/bin/env python3
"""
Parse the Utah 2024 general election statewide canvass PDF into a
county-level results CSV.

The canvass is a scanned/rendered PDF (most pages carry no extractable text
layer), and its tables mix multiple contests side-by-side on a single page
(e.g. four U.S. House districts, or five judicial retention races), with
some contests spanning more than one page (write-in candidate columns
continue onto the next page with no repeated header). Reconstructing that
layout from raw text/word positions is unreliable, so this script renders
each page to an image with NaturalPDF and asks a vision-capable LLM (via the
`llm` library and the llm-openrouter plugin) to read the table(s) on that
page into structured data. The LLM is not asked to do arithmetic or
normalization beyond reading what's printed; office/party normalization and
running-mate stripping happen in Python afterward.

Requires an OpenRouter API key configured for the `llm` tool:

    llm keys set openrouter

Run with uv:

    uv run python parsers/parse_2024_general_county.py

which is equivalent to:

    uv run python parsers/parse_2024_general_county.py \\
        "/Users/dwillis/code/openelections-sources-ut/2024 General Election/2024-General-Election-Statewide-Canvass-2.pdf" \\
        2024/20241105__ut__general__county.csv \\
        --model openrouter/qwen/qwen3.7-plus

Output columns: county,office,district,party,candidate,votes
"""

import argparse
import csv
import difflib
import io
import os
import re
import sys
from typing import List, Optional

import llm
import natural_pdf as npdf
from pydantic import BaseModel

DEFAULT_INPUT = (
    "/Users/dwillis/code/openelections-sources-ut/2024 General Election/"
    "2024-General-Election-Statewide-Canvass-2.pdf"
)
DEFAULT_OUTPUT = "2024/20241105__ut__general__county.csv"
DEFAULT_MODEL = "openrouter/qwen/qwen3.7-plus"
DEFAULT_RESOLUTION = 250
DEFAULT_RETRIES = 4

FIELDNAMES = ["county", "office", "district", "party", "candidate", "votes"]

# Offices whose printed page title should collapse to this canonical name.
# Matched case-insensitively against the LLM's reported office string.
OFFICE_ALIASES = {
    "u.s. president and vice president": "President",
    "president and vice president": "President",
    "president": "President",
    "governor, lieutenant governor": "Governor",
    "governor and lieutenant governor": "Governor",
    "governor": "Governor",
    "u.s. senate": "U.S. Senate",
    "u.s. house": "U.S. House",
    "attorney general": "Attorney General",
    "state auditor": "State Auditor",
    "state treasurer": "State Treasurer",
    "state senate": "State Senate",
    "utah senate": "State Senate",
    "state house": "State House",
    "utah house": "State House",
    "state school board": "State School Board",
    "state board of education": "State School Board",
}

# Fallback fuzzy-match targets for the handful of standard (non-judicial,
# non-amendment) offices, to catch any OCR/transcription typo that
# OFFICE_ALIASES' exact lookup would miss ("State Audtitor" was one such
# typo, since corrected here and in the emitted CSV).
STANDARD_OFFICES = sorted(set(OFFICE_ALIASES.values()))

# Offices where the printed candidate name pairs a top-of-ticket candidate
# with a running mate ("NAME, RUNNING MATE"); only the top of the ticket is
# kept, matching this repo's existing convention for presidential rows.
RUNNING_MATE_OFFICES = {"President", "Governor"}

# Matches "WRITE-IN", "WRITE IN", or "WRITEIN" (the canvass hyphenates this
# inconsistently from page to page).
WRITE_IN_RE = re.compile(r"WRITE\s*-?\s*IN")

SYSTEM_PROMPT = """You are reading one page of the official 2024 Utah General \
Election statewide canvass. Each results page is a grid: one or more \
contests laid out as groups of columns (a contest's candidates), with a row \
per county down the left side, plus "Total Votes Cast Per Candidate", \
"% of Votes", and "Total Votes Cast in Contest" summary rows at the bottom.

Some pages are NOT results tables: a cover/certification page, a table of \
contents, a "Total Ballots Cast"/voter turnout table, a colored county map \
graphic, or the "Single County Races (For information only)" page. For any \
of those, set is_results_table to false and return an empty contests list.

A page that continues a contest from the previous page (extra write-in \
candidate columns, no title/header repeated at the top) should still have \
is_results_table set to true; leave that contest's office and district as \
null/empty so the caller knows to reuse the previous page's contest.

For each genuine results contest on the page, report:
- office: the contest name, normalized as follows:
  * "U.S. President and Vice President" -> "President"
  * "Governor, Lieutenant Governor" -> "Governor"
  * "U.S. House District N" -> office "U.S. House", district "N"
  * "Utah Senate District N" -> office "State Senate", district "N"
  * "Utah House District N" -> office "State House", district "N"
  * "State Board of Education District N" -> office "State School Board", \
district "N"
  * "U.S. Senate", "Attorney General", "State Auditor", "State Treasurer" \
unchanged, district empty
  * A judicial retention contest (column header names a judge and a \
court, with Yes/No sub-columns) -> office "Retain " followed by the \
judge's name as printed, then " - " then the court/position as printed, \
e.g. "Retain Matthew B. Durrant - Supreme Court of Utah"; district empty
  * A constitutional amendment (For/Against sub-columns) -> office exactly \
as printed, e.g. "Constitutional Amendment B"; district empty
- district: the district number as a string, or empty/null if none
- candidates: in left-to-right column order, each with:
  * candidate: the name exactly as printed in the column header (keep the \
full "NAME, RUNNING MATE" text if two names are printed together; do not \
shorten it yourself). For judicial retention use "Yes"/"No". For \
constitutional amendments use whatever the two options are labeled \
(typically "For"/"Against").
  * party: the party code printed in parentheses under the candidate name \
(e.g. "REP", "DEM", "LIB", "CON", "GRN", "UNA", "IAP", "UUP"), or "" if no \
party is printed for that column (this is common for write-in candidates)
- counties: one entry per county that has a value printed in this contest's \
columns. IMPORTANT: if every cell for a county is blank/shaded for this \
contest (meaning that county is not part of this district), OMIT that \
county entirely -- do not invent a row for it. When several contests share \
the page (e.g. four U.S. House districts side by side), a county is \
usually populated in only ONE of those column groups and blank in the \
others -- do not carry a county's numbers from the column group where it \
has data over into a different contest's column group just because that \
county's row has values somewhere else on the page. Check this contest's \
own columns for that county's row in isolation. If a cell literally shows \
0, include that county with vote 0. Give the county name exactly as \
printed (e.g. "Beaver County"). votes must be a list of integers in the \
same left-to-right order as candidates, with commas in numbers like 1,208 \
read as 1208.
- totals: the "Total Votes Cast Per Candidate" row for this contest, as a \
list of integers in the same order as candidates (for cross-checking; omit \
if that row isn't present).

Read every number carefully and precisely -- these are official certified \
vote totals."""


class Candidate(BaseModel):
    candidate: str
    party: str = ""


class CountyVotes(BaseModel):
    county: str
    votes: List[int]


class Contest(BaseModel):
    office: Optional[str] = None
    district: Optional[str] = None
    candidates: List[Candidate] = []
    counties: List[CountyVotes] = []
    totals: List[int] = []


class PageResult(BaseModel):
    is_results_table: bool
    contests: List[Contest] = []


class RepairResult(BaseModel):
    counties: List[CountyVotes] = []
    totals: List[int] = []


def build_repair_prompt(office, district, candidates):
    contest_label = f"{office} District {district}" if district else office
    candidate_list = "\n".join(
        f"{i + 1}. {candidate} ({party or 'no party printed'})"
        for i, (candidate, party) in enumerate(candidates)
    )
    return f"""This page has multiple contests side by side. Focus ONLY on \
the "{contest_label}" contest's column group -- ignore every other \
contest's columns entirely, even for counties that have data there.

This contest's candidates, left to right:
{candidate_list}

For each of Utah's 29 counties, report a county ONLY if THIS SPECIFIC \
column group has a printed, non-blank/non-shaded value for that county. Do \
not include a county just because it has a value in a different contest's \
columns elsewhere on the page -- check this contest's own columns for that \
county's row in isolation. If a cell in this column group literally shows \
0, include that county with vote 0. votes must list integers in the exact \
candidate order above. Also report the "Total Votes Cast Per Candidate" \
row for just these columns, in the same order, if it's present on the \
page."""


def repair_contest(model, png_bytes, office, district, candidates, log, page_num, retries=3):
    """
    Re-query the model with a prompt narrowed to a single contest, for a
    contest whose full-page reading didn't sum to its printed total. This
    avoids the page-level failure mode where a county's data from a
    neighboring contest's column group gets attributed to this one.
    """
    prompt = build_repair_prompt(office, district, candidates)
    attachment = llm.Attachment(type="image/png", content=png_bytes)
    for attempt in range(retries + 1):
        try:
            response = model.prompt(prompt, attachments=[attachment], schema=RepairResult)
            result = RepairResult.model_validate_json(response.text())
        except Exception as exc:
            log(f"  ! page {page_num}: repair attempt failed for "
                f"{office} {district or ''} ({exc})")
            continue
        fake_contest = Contest(
            office=office, district=district,
            candidates=[Candidate(candidate=c, party=p) for c, p in candidates],
            counties=result.counties, totals=result.totals,
        )
        if result.counties and contest_totals_match(fake_contest):
            return result
    return None


def clean_county(name):
    name = " ".join(name.split())
    if name.lower().endswith(" county"):
        name = name[: -len(" county")]
    return name


def normalize_office(raw_office, raw_district):
    office = (raw_office or "").strip()
    district = (raw_district or "").strip()
    if not office:
        return office, district

    canonical = OFFICE_ALIASES.get(office.lower())
    if canonical:
        return canonical, district

    # "Retain ..." and "Constitutional Amendment ..." office strings carry a
    # judge/amendment name we don't want to fuzzy-match away; leave them as
    # transcribed.
    lower = office.lower()
    if lower.startswith("retain ") or lower.startswith("constitutional amendment"):
        return office, district

    close = difflib.get_close_matches(office, STANDARD_OFFICES, n=1, cutoff=0.6)
    if close:
        return close[0], district
    return office, district


def normalize_candidate(raw_candidate, raw_party, office):
    candidate = " ".join((raw_candidate or "").strip().split())
    party = (raw_party or "").strip().upper()

    if WRITE_IN_RE.search(candidate.upper()):
        return "Write-ins", ""

    if office in RUNNING_MATE_OFFICES and "," in candidate:
        candidate = candidate.split(",", 1)[0].strip()

    return candidate, party


def render_page_png_bytes(page, resolution):
    image = page.render(resolution=resolution)
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return buf.getvalue()


def extract_page(model, png_bytes):
    attachment = llm.Attachment(type="image/png", content=png_bytes)
    response = model.prompt(SYSTEM_PROMPT, attachments=[attachment], schema=PageResult)
    return PageResult.model_validate_json(response.text())


def contest_totals_match(contest):
    if not contest.totals or len(contest.totals) != len(contest.candidates):
        return True
    sums = [0] * len(contest.candidates)
    for county_votes in contest.counties:
        if len(county_votes.votes) != len(contest.candidates):
            continue
        for idx, votes in enumerate(county_votes.votes):
            sums[idx] += votes
    return sums == contest.totals


def result_looks_valid(result):
    """
    A trustworthy page result: classified as a results table, with at least
    one contest that actually has data (the model occasionally emits a
    stray empty contest object), and -- when a "Total Votes Cast Per
    Candidate" row was read -- summed county votes matching it.
    """
    if not result.is_results_table:
        return False
    real_contests = [c for c in result.contests if c.candidates and c.counties]
    if not real_contests:
        return False
    return all(contest_totals_match(c) for c in real_contests)


def extract_page_with_retry(model, png_bytes, log, page_num, retries=DEFAULT_RETRIES):
    """
    Vision extraction is occasionally non-deterministic: a genuine results
    page can come back misclassified as is_results_table=false, with a
    stray empty contest, or with numbers that don't sum to the page's own
    printed total. Retry a few times looking for an internally-consistent
    read before giving up and returning the last attempt as a best effort.
    """
    last_result = None
    for attempt in range(retries + 1):
        try:
            result = extract_page(model, png_bytes)
        except Exception as exc:
            log(f"  ! page {page_num}: LLM extraction failed ({exc})")
            continue
        last_result = result
        if result_looks_valid(result):
            return result
        if attempt < retries:
            log(f"  page {page_num}: extraction looked inconsistent, "
                f"retrying ({attempt + 1}/{retries})...")
    if last_result is not None and not result_looks_valid(last_result):
        log(f"  ! page {page_num}: keeping best-effort result after "
            f"retries; check warnings above for this page")
    return last_result


def process_pdf(pdf_path, output_csv, model_name, resolution, start, end,
                 retries=DEFAULT_RETRIES, log=print):
    model = llm.get_model(model_name)
    pdf = npdf.PDF(pdf_path)
    total_pages = len(pdf.pages)
    end = total_pages if end is None else min(end, total_pages)

    rows = []
    current_office = None
    current_district = ""

    for i in range(start, end):
        page = pdf.pages[i]
        log(f"Page {i + 1}/{total_pages}: rendering + analyzing...")
        png_bytes = render_page_png_bytes(page, resolution)

        result = extract_page_with_retry(model, png_bytes, log, i + 1, retries=retries)
        if result is None:
            log(f"  ! page {i + 1}: no usable response after retries; skipping")
            continue

        if not result.is_results_table or not result.contests:
            log("  (not a results table, skipped)")
            continue

        for contest in result.contests:
            if not contest.candidates or not contest.counties:
                # Stray/placeholder contest object (seen occasionally ahead
                # of a page's real contest); not a continuation marker.
                continue

            office, district = normalize_office(contest.office, contest.district)
            if office:
                current_office, current_district = office, district
            elif current_office is not None:
                office, district = current_office, current_district
            else:
                log(f"  ! page {i + 1}: contest has no office and no prior "
                    f"context; skipping")
                continue

            candidates = [
                normalize_candidate(c.candidate, c.party, office)
                for c in contest.candidates
            ]

            counties_data = contest.counties
            if not contest_totals_match(contest):
                log(f"  page {i + 1}: {office} {district or ''} - totals "
                    f"don't match, attempting targeted repair...")
                repaired = repair_contest(model, png_bytes, office, district,
                                           candidates, log, i + 1)
                if repaired is not None:
                    counties_data = repaired.counties
                    log(f"  page {i + 1}: {office} {district or ''} - "
                        f"repair succeeded")
                else:
                    log(f"  ! page {i + 1}: {office} {district or ''} - "
                        f"repair failed, keeping original (flagged) numbers")

            candidate_sums = {cand: 0 for cand, _party in candidates}

            for county_votes in counties_data:
                county = clean_county(county_votes.county)
                if len(county_votes.votes) != len(candidates):
                    log(f"  ! page {i + 1}: {office} {district} - column "
                        f"count mismatch for {county}, skipping that county")
                    continue
                for (candidate, party), votes in zip(candidates, county_votes.votes):
                    rows.append({
                        "county": county,
                        "office": office,
                        "district": district,
                        "party": party,
                        "candidate": candidate,
                        "votes": votes,
                    })
                    candidate_sums[candidate] += votes

            if contest.totals and len(contest.totals) == len(candidates):
                for (candidate, _party), reported_total in zip(candidates, contest.totals):
                    summed = candidate_sums.get(candidate, 0)
                    if summed != reported_total:
                        log(f"  ! page {i + 1}: {office} {district or ''} - "
                            f"{candidate!r} county votes still sum to "
                            f"{summed}, page total is {reported_total}")

            log(f"  {office} {district or ''}: {len(counties_data)} "
                f"counties x {len(candidates)} candidates")

    os.makedirs(os.path.dirname(output_csv) or ".", exist_ok=True)
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)

    log(f"\nWrote {len(rows)} rows to {output_csv}")
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_pdf", nargs="?", default=DEFAULT_INPUT,
                         help="Path to the statewide canvass PDF")
    parser.add_argument("output_csv", nargs="?", default=DEFAULT_OUTPUT,
                         help="Path to write the county-level results CSV")
    parser.add_argument("--model", default=DEFAULT_MODEL,
                         help=f"llm model id to use (default: {DEFAULT_MODEL})")
    parser.add_argument("--resolution", type=int, default=DEFAULT_RESOLUTION,
                         help=f"Page render resolution in DPI (default: {DEFAULT_RESOLUTION})")
    parser.add_argument("--start", type=int, default=0,
                         help="First page index to process (0-based, default: 0)")
    parser.add_argument("--end", type=int, default=None,
                         help="Last page index (exclusive) to process (default: all pages)")
    parser.add_argument("--retries", type=int, default=DEFAULT_RETRIES,
                         help="Extra extraction attempts per page when the read looks "
                              f"inconsistent (default: {DEFAULT_RETRIES})")
    args = parser.parse_args()

    if not os.path.exists(args.input_pdf):
        print(f"Input PDF not found: {args.input_pdf}", file=sys.stderr)
        sys.exit(1)

    process_pdf(
        args.input_pdf,
        args.output_csv,
        args.model,
        args.resolution,
        args.start,
        args.end,
        retries=args.retries,
    )


if __name__ == "__main__":
    main()
