"""Parser for the Electionware "Summary Results Report" / precinct formats used
by most Utah counties in the 2026 primary.  Works on both embedded-text PDFs
(pdftotext -layout) and scanned PDFs (PaddleOCR), via the unified visual-row
extraction in p2026_common.

Contest-header variants handled:
    Republican for U.S. House District 3
    REP U.S. House District 2
    DEM UTAH STATE SENATE DISTRICT 5
    Republican Candidate for U.S. House District 3
    REP Republican for Congressional District 3
    State School Board District 14            (nonpartisan, no party prefix)

Vote-column variants handled:
    CELESTE MALOY          1,425                      (single TOTAL column)
    CELESTE MALOY          5,366   68.23%             (TOTAL + VOTE %)
"""

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import p2026_common as C

# --------------------------------------------------------------------------- #
# Statistics
# --------------------------------------------------------------------------- #

_STAT_LABELS = {
    "Registered Voters": C.REGISTERED_VOTERS,
    "Ballots Cast": C.BALLOTS_CAST,
    "Ballots Cast Blank": C.BALLOTS_CAST_BLANK,
}

_PARTY_SUFFIX = {
    "TOTAL": "", "REPUBLICAN": "REP", "DEMOCRATIC": "DEM", "DEMOCRAT": "DEM",
    "NONPARTISAN": "NP", "NONPARTISAN N": "NP", "LIBERTARIAN": "LIB",
    "INDEPENDENT AMERICAN": "IAP", "CONSTITUTION": "CON", "UNITED UTAH": "UUP",
    "INDEPENDENT": "IND", "GREEN": "GRN", "UNAFFILIATED": "",
}

_NUM = r"[\d,]+"


def _first_number(text):
    m = re.search(_NUM, text)
    return m.group(0) if m else ""


_STAT_LABEL_RES = [
    (re.compile(r"^REGISTER\s*ED\s*VOTERS\b"), C.REGISTERED_VOTERS),
    (re.compile(r"^BALLOTS\s*CAST\s*BLANK\b"), C.BALLOTS_CAST_BLANK),
    (re.compile(r"^BALLOTS\s*CAST\b"), C.BALLOTS_CAST),
]


def _lev(a, b):
    """Levenshtein distance between two short uppercase tokens."""
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if abs(la - lb) > 3:
        return 99
    if la > lb:
        a, b = b, a
        la, lb = lb, la
    prev = list(range(lb + 1))
    for i in range(1, la + 1):
        cur = [i] + [0] * lb
        ca = a[i - 1]
        for j in range(1, lb + 1):
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1,
                         prev[j - 1] + (ca != b[j - 1]))
        prev = cur
    return prev[lb]


def _has_token(toks, target, max_dist):
    """True if any token is within `max_dist` edits of `target`."""
    return any(_lev(t, target) <= max_dist for t in toks)


def _party_from_suffix(toks):
    """Extract a party code from the trailing token(s) of a statistics label."""
    if not toks:
        return ""
    for n in (1, 2):
        if len(toks) >= n:
            key = " ".join(toks[-n:])
            if key in _PARTY_SUFFIX:
                return _PARTY_SUFFIX[key]
    return ""


def _parse_statistics_line_fuzzy(norm):
    """OCR-tolerant fallback: classify a statistics line by anchor tokens within
    a small edit distance of the canonical words.  Handles garbles such as
    'Bailots Cast', 'Eallots Casl', 'Ballols Casl', 'Regislered Volers',
    'Registeled Volers' Republican', 'Reglstered Voters. Republican'."""
    # Turnout / percentage lines are not captured statistics rows.
    if "TURNOUT" in norm or "%" in norm:
        return None
    vm = re.search(_NUM, norm)
    if not vm:
        return None
    votes = vm.group(0)
    label_text = re.sub(r"\s+", " ", norm[:vm.start()].strip())
    if not label_text:
        return None
    toks = label_text.split()
    party = _party_from_suffix(toks)
    # Ballots Cast Blank -- "BLANK" must match exactly: the common surname
    # "BLACK" is one edit away and would otherwise be misclassified.
    if _has_token(toks, "BLANK", 0):
        return C.BALLOTS_CAST_BLANK, "", votes
    # Registered Voters -- require both a VOTERS-ish and a REGISTERED-ish token
    # so a stray candidate surname near one word cannot trigger it.
    if (_has_token(toks, "VOTERS", 1) or _has_token(toks, "VOTER", 1)) \
       and (_has_token(toks, "REGISTERED", 2) or _has_token(toks, "REGISTER", 2)):
        return C.REGISTERED_VOTERS, party, votes
    # Ballots Cast -- the "Ballots" word is the distinctive anchor.
    if _has_token(toks, "BALLOTS", 1) or _has_token(toks, "BALLOT", 1):
        return C.BALLOTS_CAST, party, votes
    return None


def parse_statistics_line(text):
    """Return (label, party_code, votes_str) if text is a statistics row, else None.

    Tries an exact-label match first (clean text), then falls back to fuzzy
    anchor-token matching for OCR-garbled labels (e.g. 'Bailots Cast - Total',
    'Regislered Volers - Republican', 'Eallots Casl - Republican')."""
    norm = re.sub(r"\s+", " ", text).strip().upper().replace("'", " ")
    norm = re.sub(r"\s*-\s*", " ", norm)  # " - " separator between label and suffix
    for pat, out in _STAT_LABEL_RES:
        m = pat.match(norm)
        if m:
            rest = norm[m.end():].lstrip(" -").strip()
            votes = ""
            vm = re.search(_NUM, rest)
            if vm:
                votes = vm.group(0)
                rest = rest[:vm.start()].strip()
            party = _PARTY_SUFFIX.get(rest, "")
            return out, party, votes
    return _parse_statistics_line_fuzzy(norm)


# --------------------------------------------------------------------------- #
# Contest header parsing
# --------------------------------------------------------------------------- #

_CONTEST_PARTY_PREFIX = [
    (re.compile(r"^(REP|DEM|LIB|IAP|CON|UUP|IND|GRN)\s+(.*)$"), None),
    (re.compile(r"^(REPUBLICAN|DEMOCRATIC|LIBERTARIAN|INDEPENDENT AMERICAN|CONSTITUTION|UNITED UTAH)\s+(?:CANDIDATE\s+)?FOR\s+(.*)$", re.IGNORECASE), None),
]


def parse_contest_header(text):
    """Return (office_raw, district, party_code) for a contest header line, or None."""
    raw = re.sub(r"\s+", " ", text).strip()
    if not raw:
        return None
    up = raw.upper()
    # Skip obvious non-headers
    if re.match(r"^(VOTE FOR|TOTAL|CANDIDATE|TIMES CAST|STATISTICS|SUMMARY RESULTS|"
                r"OFFICIAL|PRIMARY ELECTION|JUNE|PAGE|RESULTS|VOTER TURNOUT|2026|"
                r"ELECTION DAY|PRECINCTS|ABSENTEE|REGISTERED VOTERS|BALLOTS CAST|"
                r"TOTAL VOTES|CONTEST TOTALS|OVER ?VOTES|UNDER ?VOTES)", up):
        return None
    # Must not contain a trailing vote number (those are candidate/result lines)
    # but may contain digits in a district.
    party = ""
    office = raw
    # Code prefix: "REP U.S. House District 2"
    m = re.match(r"^(REP|DEM|LIB|IAP|CON|UUP|IND|GRN)\b\s*(.*)$", up)
    code = None
    if m:
        code = m.group(1)
        office = m.group(2).strip() or raw
    # "Republican for X" / "Republican Candidate for X"
    m2 = re.match(r"^(REPUBLICAN|DEMOCRATIC|LIBERTARIAN|INDEPENDENT AMERICAN|CONSTITUTION|UNITED UTAH)\s+(?:CANDIDATE\s+)?FOR\s+(.*)$", up)
    if m2:
        party = C.normalize_party(m2.group(1))
        office = m2.group(2).strip() or office
    elif code:
        party = code
    # If office still begins with "Republican for ..." (Utah: "REP Republican for Congressional District 3")
    m3 = re.match(r"^(REPUBLICAN|DEMOCRATIC|LIBERTARIAN|INDEPENDENT AMERICAN|CONSTITUTION|UNITED UTAH)\s+FOR\s+(.*)$", office.upper())
    if m3:
        party = C.normalize_party(m3.group(1))
        office = m3.group(2)
    office = re.sub(r"\s*\(VOTE\s+FOR\s*\d+\)\s*", "", office, flags=re.IGNORECASE)
    office = re.sub(r"\s*-\s*(REP|DEM|LIB|IAP|CON|UUP|IND|GRN)\s*$", "", office, flags=re.IGNORECASE)
    district = C.parse_district(office)
    office_norm = C.normalize_office(office)
    return office_norm, district, party


# --------------------------------------------------------------------------- #
# Line classification within a contest
# --------------------------------------------------------------------------- #

_SKIP_PREFIXES = (
    "VOTE FOR", "OTE FOR", "VATE FOR", "VOLER TURNOUT", "VOTER TURNOUT", "TIMES CAST",
    "TOTAL VOTES", "OTAL VOTES", "TOTAL VOTE", "CONTEST TOTALS", "ONTEST TOTALS",
    "ELECTION DAY", "PRECINCTS ", "ABSENTEE", "PAGE", "RESULTS -",
    "RESULTS REPORT", "2026 PRIMARY", "PRIMARY ELECTION", "JUNE", "OFFICIAL",
    "SUMMARY RESULTS", "STATISTICS", "CANDIDATE", "REGISTERED VOTERS",
    "BALLOTS CAST", "VOTERS CAST", "ELECTION SUMMARY", "FULL ELECTION",
    "CANVASS", "SPARE", "LECTION SUMMARY", "ECTION SUMMARY",
)
_SKIP_EXACT = {"TOTAL", "TOTAL.", "STATISTICS", "CANDIDATE", "TOTAL VOTE %",
               "VOTE %", "TOTALS", "OFFICIAL RESULTS", "OFFICAL ELECTION RESULTS"}


def _is_contest_totals_fuzzy(up):
    """Catch OCR-garbled 'Contest Totals <votes>' lines (e.g. 'Conlesl Totals',
    'Contesl Toials', 'Cootest Totals') that slip past the exact skip prefixes.
    A two-token label where one token is near CONTEST and the other near TOTALS."""
    label = re.sub(r"\s*[\d,]+\s*%?\s*$", "", up).strip()
    toks = label.split()
    if len(toks) != 2:
        return False
    return (_has_token(toks, "CONTEST", 2)
            and (_has_token(toks, "TOTALS", 2) or _has_token(toks, "TOTAL", 1)))

# Candidate line: name (letters/punctuation, no digits) then the FIRST number
# is the TOTAL vote column; trailing vote-breakdown columns (Election Day /
# Early / Mail / Spare 1..3) and a trailing vote-% are ignored.  We do not
# anchor to end-of-line because multi-column counties append breakdown cols.
_CAND_RE = re.compile(r"^(?P<name>['\"]?[A-Za-z][A-Za-z .'\-\"]*?)\s+(?P<votes>[\d,]+)(?:\s|$)")
# Over/Under votes: likewise take the first number, ignore trailing breakdowns.
_OU_RE = re.compile(r"^(?P<kind>Overvotes|Undervotes|Over\s+Votes|Under\s+Votes)\s+(?P<votes>[\d,]+)", re.IGNORECASE)

# Per-contest column-header rows and their wrapped continuations, e.g.
#     TOTAL       Election    Early   Mail   Spare 1   Spare 2   Spare 3
#                     Day    Voting
# These are not contests or candidates; skip them.
_COL_HEADER_TOKENS = {"DAY", "VOTING", "DAY VOTING", "TOTAL", "TOTAL.", "TOTALS"}


def _is_col_header(up):
    # Collapse internal whitespace so "DAY    VOTING" == "DAY VOTING".
    up = re.sub(r"\s+", " ", up).strip()
    if up in _COL_HEADER_TOKENS:
        return True
    if up.startswith("TOTAL VOTE"):
        return True
    # The full column-header line: contains ELECTION plus a vote-method column.
    if "ELECTION" in up and any(k in up for k in ("SPARE", "MAIL", "EARLY", "VOTING")):
        return True
    return False


def _is_skip(text):
    up = re.sub(r"\s+", " ", text.strip()).upper()
    if up in _SKIP_EXACT:
        return True
    if _is_col_header(up):
        return True
    if up.startswith(_SKIP_PREFIXES):
        return True
    return _is_contest_totals_fuzzy(up)


def _clean_name(name):
    # Strip leading indent markers / quotes, and any inner double-quotes left
    # by quoted nicknames (e.g. '"CJ" CHRISTINA HERNANDEZ' -> 'CJ CHRISTINA HERNANDEZ').
    n = name.strip().strip('"').strip()
    n = re.sub(r'^[""\']+', "", n).strip()
    n = n.replace('"', "")
    return n


# --------------------------------------------------------------------------- #
# Summary parser
# --------------------------------------------------------------------------- #

def parse_summary(county, pdf_path):
    return _parse_rows(county, pdf_path, precinct_mode=False)


def parse_precinct(county, pdf_path):
    """Parse a per-precinct sequential results report.  Each precinct section
    begins with a precinct-id line (a line starting with a digit) followed by a
    Statistics block and the same contest/candidate lines as the summary.  The
    summary line-classification logic is reused; precinct ids are detected and
    threaded into every row."""
    return _parse_rows(county, pdf_path, precinct_mode=True)


# A precinct-id line that starts with a digit (the common case, e.g. "01BV01:1",
# "0801 EMERY", "10-1C").  Alphabetic precinct names like "Federal" are caught
# instead by the Statistics-block lookback below.
_PRECINCT_ID_RE = re.compile(r"\d+\s*OF\s*\d+")

# Page-header/footer keywords that can never be part of a precinct CODE.
# Note "COUNTY"/"UTAH" are intentionally excluded: real precinct names can
# contain them (e.g. "26-120 County NE", "Utah County ...").  Date/county
# header lines are still rejected by the comma check in the name path and by
# the other keywords here ("JUNE"/"2026"/"ELECTION"/"PRIMARY"/...).
_ID_HEADER_RE = re.compile(
    r"ELECTION|PRIMARY|SUMMARY|RESULTS|OFFICIAL|PAGE\b|REPORT|VOTER|"
    r"TURNOUT|2026|JUNE|STATISTIC|CANVASS|ABSENTEE|CONTEST|TOTALS?\b|"
    r"OVERVOTES|UNDERVOTES|OVER\s+VOTES|UNDER\s+VOTES")


def _looks_like_precinct_id(t):
    """A line that starts with a digit and is a precinct identifier (e.g.
    "01BV01:1", "0801 EMERY", "10-1C", "1911", "1913B Suppressed",
    "26-120 County NE").  In the sequential Electionware format the only
    digit-start lines are precinct codes and "N of M" turnout fragments, so
    any digit-start line that is not a percentage, an "of" fragment, or a
    page header is treated as a precinct id.  Alphabetic
    precinct names like "Federal" are caught by the Statistics-block lookback.
    """
    s = t.strip()
    if not s or not s[0].isdigit():
        return False
    up = s.upper()
    if "%" in up or _PRECINCT_ID_RE.search(up):
        return False
    # Reject page headers/footers that happen to start with a digit, e.g.
    # "2026 Primary Election ... CANVASS", "2026 Republican Primary Election".
    # Real precinct codes do not contain these words.  (COUNTY/UTAH are not
    # checked here -- precinct names can contain them.)
    if _ID_HEADER_RE.search(up):
        return False
    return True


# Header/footer lines the Statistics-block lookback must NOT mistake for a
# precinct name (page headers like "June 23, 2026 Carbon County, Utah",
# "Summary Results Report OFFICIAL RESULTS", "Precinct Summary - ...").
_HEADER_RE = re.compile(
    r"COUNTY|ELECTION|SUMMARY|RESULTS|OFFICIAL|PAGE\b|UTAH|PRIMARY|"
    r"REPORT|VOTER|TURNOUT|2026|JUNE|STATISTIC|PRECINCT\b|CANVASS|ABSENTEE|"
    r"CONTEST|TOTALS?\b")


def _looks_like_precinct_name(t):
    """True if a line plausibly names a precinct (for the alphabetic-name
    lookback path).  Rejects headers/footers, date lines, and result rows.

    A real candidate line has a multi-word name or a sizable vote count.
    A precinct name like "Annabella 1" / "Richfield 11" is a single word
    followed by a small precinct-split number, which we keep."""
    s = t.strip()
    if not s or "," in s or len(s) > 50:
        return False
    if _HEADER_RE.search(s.upper()):
        return False
    m = _CAND_RE.match(s)
    if m:
        votes = m.group("votes").replace(",", "")
        if len(m.group("name").split()) >= 2 or int(votes) > 50:
            return False
    if _OU_RE.match(s):
        return False
    return any(c.isalpha() for c in s)


def _parse_rows(county, pdf_path, precinct_mode=False):
    from collections import deque
    pages = C.extract_rows(pdf_path, county)
    display = C.COUNTY_DISPLAY.get(county, county)
    rows = []
    office, district, party = "", "", ""
    precinct = ""
    # Ring buffer of recent raw lines, used to recover the precinct id as the
    # line immediately preceding each Statistics block (skipping the "TOTAL"
    # column header).  This catches alphabetic precinct names ("Federal") that
    # the digit-start fast path misses.
    recent = deque(maxlen=8)

    for page in pages:
        n = len(page)
        for idx in range(n):
            t = page[idx][2].strip()
            if not t:
                continue
            # Next non-empty line in this page, for look-ahead.
            nxt = ""
            for j in range(idx + 1, min(idx + 3, n)):
                s = page[j][2].strip()
                if s:
                    nxt = s
                    break
            if precinct_mode:
                up = re.sub(r"\s+", " ", t).upper()
                # Continuation pages re-emit the precinct id at the top (after
                # the page header) without a "TOTAL"/"Statistics" follower, so
                # the look-ahead below would miss it and parse it as a
                # candidate row of the previous contest.  If the line matches
                # the current precinct, it is that re-emission -- skip it.
                if precinct and t == precinct:
                    recent.append(t)
                    continue
                # A Statistics block marks the start of a precinct section; the
                # precinct id is the nearest preceding non-column-header line.
                # The line is often "Statistics" but may carry a trailing
                # token ("Statistics  TOTAL"), so match by prefix.
                if up.startswith("STATISTIC"):
                    # If a precinct id was already captured just above (by the
                    # digit-start or name+number look-ahead paths) and is still
                    # the nearest id in the recent window, this Statistics block
                    # belongs to that precinct -- keep it.  Otherwise the lookback
                    # below would reject a multi-word name like "Oak City 6"
                    # (it looks like a candidate) and reset precinct to "",
                    # dropping the whole precinct's rows.
                    if precinct and precinct in recent:
                        office, district, party = "", "", ""
                        recent.append(t)
                        continue
                    new_precinct = None
                    for prev in reversed(recent):
                        pu = re.sub(r"\s+", " ", prev).upper()
                        if pu in ("", "TOTAL", "TOTAL.", "STATISTICS", "STATISTIC") \
                           or pu.startswith("STATISTIC"):
                            continue
                        # A digit-start precinct code preceding the block is
                        # the precinct id -- stop here rather than scanning back
                        # into the previous precinct's result lines.
                        if _looks_like_precinct_id(prev):
                            new_precinct = prev
                            break
                        if not _looks_like_precinct_name(prev):
                            continue
                        new_precinct = prev
                        break
                    # If no plausible precinct name precedes this Statistics
                    # block, it is the county-summary section (which has no
                    # precinct id) -- reset to "" so its rows are skipped by the
                    # empty-precinct guard rather than tagged onto the last
                    # real precinct and doubling its sums.
                    precinct = new_precinct if new_precinct is not None else ""
                    office, district, party = "", "", ""
                    recent.append(t)
                    continue
                # Digit-start precinct-id line (also covers continuation pages,
                # where the id is re-emitted without a Statistics block).
                if _looks_like_precinct_id(t):
                    precinct = t
                    office, district, party = "", "", ""
                    recent.append(t)
                    continue
                # Alpha-start precinct id like "Abraham 14" / "Annabella 1": a
                # name+number line immediately followed by the "TOTAL" column
                # header (or "Statistics").  Without this it would be mis-parsed
                # as a candidate row of the previous contest.  "Vote For N" lines
                # are also name+number+TOTAL but are excluded via _is_skip, and
                # contest headers ("State School Board District 14") are excluded
                # via _looks_like_header (they contain DISTRICT/FOR/party codes).
                nu = re.sub(r"\s+", " ", nxt).upper().strip()
                if _CAND_RE.match(t) and not _is_skip(t) \
                   and not _looks_like_header(t) \
                   and (nu in ("TOTAL", "TOTAL.") or nu.startswith("STATISTIC")):
                    precinct = t
                    office, district, party = "", "", ""
                    recent.append(t)
                    continue
                recent.append(t)
            # In precinct mode, every emitted row must belong to a known
            # precinct.  The county-totals / "Contest Totals" summary section
            # that follows all precinct sections has no precinct id, so
            # precinct is "" there -- skip those rows (they would otherwise
            # double the per-precinct sums).
            if precinct_mode and not precinct:
                continue
            # Statistics line?
            stat = parse_statistics_line(t)
            if stat:
                label, sp, svotes = stat
                # Skip zero-value nonpartisan registration noise
                if sp == "NP" and C.to_int(svotes) == 0:
                    continue
                rows.append(C.meta_row(display, label, sp, svotes, precinct=precinct))
                continue
            if _is_skip(t):
                continue
            # Over/Under votes within a contest
            ou = _OU_RE.match(t)
            if ou:
                kind = "Over Votes" if "OVER" in ou.group("kind").upper() else "Under Votes"
                rows.append(C.overunder_row(display, office, district, party, kind,
                                             ou.group("votes"), precinct=precinct))
                continue
            # Contest header (no trailing number)?
            if not re.search(r"[\d,]+\s*%?\s*$", t) or _looks_like_header(t):
                hdr = parse_contest_header(t)
                if hdr:
                    office, district, party = hdr
                    continue
                # ambiguous: a line with no number that isn't a recognized header
                # -> treat as contest header if it has alpha content
                if re.search(r"[A-Za-z]{3,}", t) and not re.search(r"[\d]", t):
                    hdr = parse_contest_header(t)
                    if hdr and hdr[0]:
                        office, district, party = hdr
                    continue
            # Candidate line
            cm = _CAND_RE.match(t)
            if cm:
                name = _clean_name(cm.group("name"))
                # Reject if "name" is actually a skip label that slipped through
                if _is_skip(name):
                    continue
                rows.append(C.candidate_row(display, office, district, party,
                                             C.title_case_name(name), cm.group("votes"),
                                             precinct=precinct))
                continue
    return rows


def _looks_like_header(text):
    up = text.upper()
    # Headers often contain "FOR" or "DISTRICT" or known office keywords and no comma-separated votes
    if (" FOR " in up or " DISTRICT" in up or up.startswith("REP ")
            or up.startswith("DEM ") or "CANDIDATE FOR" in up):
        return True
    # OCR-garbled "District" (e.g. "oistrict", D->O) still marks a contest header
    # and must NOT be mistaken for a precinct name+number.
    return _has_token(re.sub(r"\s+", " ", up).split(), "DISTRICT", 1)