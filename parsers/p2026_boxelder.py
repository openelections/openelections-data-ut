"""Parser for Box Elder County's 2026 primary precinct PDF.

Box Elder's precinct report is a scanned "Custom Table Report" (zero embedded
text -> PaddleOCR at 400 DPI).  Layout: precincts as ROWS (codes down the left
at x~160), columns grouped by contest.  Each contest prints N candidate vote
columns followed by TWO stat columns -- "Total Votes Cast" and "Contest Total"
-- with NO separate Over/Under columns (over+under are folded into Contest
Total, so per-precinct Over/Under are not recoverable and are omitted).

Sections (confirmed from the 9-page OCR):
  pages 0-1 (top)   STATISTICS -- 52 precincts, 7 meta columns:
                      Registered Voters total/REP/NP, Ballots Cast total/REP/NP,
                      Ballots Cast Blank.  A "Totals" row at the bottom carries
                      the county totals (30270/21802/8468/10223/9428/795/5)
                      which identify which x-column is which meta field.
  pages 1(bot)+2+3  U.S. House 2 + Commissioner Seat A (both REP, side by side)
  pages 4-5         Commissioner Seat B + Sheriff (both REP, side by side)
  page 6            School Board District 3 (nonpartisan, 4 candidates)
  page 7            School Board District 4 (nonpartisan, 6 candidates)
  page 8            School Board District 7 (nonpartisan, 3 candidates)

Candidate identities are APPLIED FROM THE SUMMARY (the county summary CSV):
OCR'd column-header name fragments are clustered by x-position and matched to
summary candidates by letter-set Jaccard overlap, so the column order -- which
differs from the summary order for the multi-candidate school-board contests --
is handled correctly.

Values are RIGHT-ALIGNED: a 1-digit value's left edge sits ~25px right of a
5-digit value's left edge in the same column.  Column assignment uses a 90px
x-tolerance (column gaps are ~210px).  Each precinct's values sit within ~6px
of the precinct-code's y (the next precinct is ~70px away), so a 18px y-band
matches values to their precinct reliably.

OCR occasionally misses a single small value (seen on the School Board 4 page).
When a precinct has exactly one candidate column without a value and the
"Total Votes Cast" column IS present, the missing value is inferred as
TVC - sum(assigned candidates).
"""

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import p2026_common as C

CODE_RE = re.compile(r"^0\d[A-Z0-9-]{3,}$")

# Value x-assignment tolerance.  Values are right-aligned: a 1-digit value's
# left edge is ~25px right of a 5-digit value's left edge in the same column.
# Column gaps are ~210px, so 90px never merges adjacent columns.
COL_TOL = 90
# Precinct-code y tolerance: a precinct's code and its values sit within ~6px
# of each other; the next precinct is ~70px away.  18px is a safe band.
CODE_Y_TOL = 18
# Gap that starts a new column when clustering value x-positions.  Within a
# column the left-edge span (1-digit .. 5-digit values) is ~50px; column gaps
# are ~210px.  120px separates columns without splitting a sparse column.
COL_GAP = 120
# Tolerance for matching a header name-cluster x-center to a value-column
# center (headers are left-of-center, values right-aligned, but within ~30px).
NAME_COL_TOL = 60
# Tolerance for matching a stat label ("Total Votes Cast"/"Contest Total") x
# to a value-column center (label left edge vs right-aligned values, ~40px).
STAT_LABEL_TOL = 80
# Jaccard threshold for matching an OCR'd name cluster to a summary candidate.
NAME_JACCARD = 0.5
# Gap for clustering header name fragments into one candidate.  Name parts
# (first/middle/last) stack vertically and span up to ~110px in x; adjacent
# candidates are ~150px apart.  90px gap-based clustering keeps a 3-part name
# together without merging neighbors.
NAME_CLUSTER_GAP = 90

# Header tokens that are structural / office / party words, not candidate-name
# parts.  Used to isolate candidate-name clusters from the contest header band.
_NAME_STOP = {
    "TOTAL", "TOTALS", "VOTES", "CAST", "VOTE", "CONTEST", "FOR", "OVERVOTES",
    "UNDERVOTES", "OVER", "UNDER", "REPUBLICAN", "REPUBLICANS", "DEMOCRAT",
    "DEMOCRATS", "DEMOCRATIC", "NONPARTISAN", "PARTISAN", "REGISTERED",
    "BALLOTS", "TURNOUT", "VOTER", "VOTERS", "BLANK", "STATISTICS", "STATISTIC",
    "PAGE", "OFFICIAL", "RESULTS", "CUSTOM", "TABLE", "REPORT", "PRIMARY",
    "ELECTION", "ELECTIONWARE", "COUNTY", "JUNE", "FEDERAL", "PRECINCT",
    "CANDIDATE", "U", "S", "US", "HOUSE", "SENATE", "DISTRICT", "DIST",
    "SEAT", "COMMISSIONER", "SHERIFF", "SCHOOL", "BOARD", "BOX", "ELDER",
    "ONE", "OF", "AM", "PM",
}

# Statistics meta-column targets: (office label, party).  The county-total
# value that identifies each column is read from the summary at runtime.
_META_TARGETS = [
    (C.REGISTERED_VOTERS, ""),
    (C.REGISTERED_VOTERS, "REP"),
    (C.REGISTERED_VOTERS, "NP"),
    (C.BALLOTS_CAST, ""),
    (C.BALLOTS_CAST, "REP"),
    (C.BALLOTS_CAST, "NP"),
    (C.BALLOTS_CAST_BLANK, ""),
]


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def _letters(s):
    """Set of lowercase letters in s (spaces/punct stripped)."""
    return set(re.sub(r"[^a-z]", "", (s or "").lower()))


def _jaccard(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _val(text):
    """Extract an integer value from a token, tolerating OCR artifacts
    (trailing fullwidth colons, etc.).  Rejects percentages and dates."""
    if "%" in text or "." in text or "/" in text:
        return None
    digits = re.sub(r"[^\d]", "", text)
    if not digits:
        return None
    return int(digits)


def _column_centers(x_positions):
    """Cluster x-positions into column centers (gap > COL_GAP starts a new
    column)."""
    xs = sorted(set(round(x) for x in x_positions))
    clusters = []
    for x in xs:
        if clusters and x - clusters[-1][-1] <= COL_GAP:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    return [round(sum(c) / len(c)) for c in clusters]


def _nearest(x, centers):
    best, best_d = None, 1e9
    for c in centers:
        d = abs(c - x)
        if d < best_d:
            best_d, best = d, c
    return best, best_d


# --------------------------------------------------------------------------- #
# Summary loading
# --------------------------------------------------------------------------- #

def _load_summary(summary_csv):
    """Return (meta_totals, contests).
    meta_totals: {(office, party): value} for the 7 statistics targets.
    contests: [(office, district, party, [candidate_names])] in summary order,
              excluding Over/Under Votes and meta rows."""
    rows = C._read_csv(summary_csv)
    meta_totals = {}
    contests = []
    cur = None
    _META = {C.REGISTERED_VOTERS, C.BALLOTS_CAST, C.BALLOTS_CAST_BLANK}
    for r in rows:
        office, district, party, cand = (
            r["office"], r["district"], r["party"], r["candidate"])
        if office in _META:
            meta_totals[(office, party)] = C.num(r["votes"])
            continue
        if cand in ("Over Votes", "Under Votes", ""):
            continue
        key = (office, district, party)
        if cur is None or (cur[0], cur[1], cur[2]) != key:
            cur = [office, district, party, []]
            contests.append(cur)
        cur[3].append(cand)
    return meta_totals, [(c[0], c[1], c[2], c[3]) for c in contests]


# --------------------------------------------------------------------------- #
# Section detection
# --------------------------------------------------------------------------- #

def _row_text(page, y, band=10):
    """Join all tokens on a visual row (y +/- band) into one uppercase string."""
    parts = [l["text"] for l in page if abs(l["y"] - y) <= band]
    return " ".join(parts).upper()


def _section_markers(pages):
    """Find section-header marker rows across all pages.
    Returns list of (page, y, kind) where kind in {'STAT','CONTEST'}, with
    same-page markers within 50px collapsed (two side-by-side contest headers
    are one section)."""
    raw = []
    for p, page in enumerate(pages):
        for l in page:
            if l["x"] > 3000:  # skip right-margin header/footer
                continue
            t = l["text"].upper()
            y = round(l["y"])
            if "STATISTICS" in t and l["x"] > 500:
                raw.append((p, y, "STAT"))
            elif ("REPUBLICAN FOR" in t
                  or ("SCHOOL BOARD" in t and "DISTRICT" in t)) \
                    and l["x"] < 2500:
                raw.append((p, y, "CONTEST"))
    raw.sort()
    # collapse same-page markers within 50px
    out = []
    for m in raw:
        if out and out[-1][0] == m[0] and m[1] - out[-1][1] <= 50:
            continue
        out.append(m)
    return out


def _is_stop_token(text):
    """True if the token is a structural/office header (every alpha part is a
    stop-word or single letter), not a candidate name.  Catches multi-word
    office titles like 'Republican for Box Elder County' and 'Commissioner
    Seat B' that single-word _NAME_STOP lookups miss."""
    parts = re.split(r"[\s.:/]+", text.upper())
    parts = [p for p in parts if p]
    if not parts:
        return True
    for p in parts:
        if len(p) <= 1:
            continue  # single char -- treat as stop
        if p not in _NAME_STOP:
            return False  # found a real name part -> candidate
    return True


def _name_clusters(header_tokens):
    """Candidate-name clusters from header tokens (list of (y, x, text)).
    Returns [(center_x, letters_set)] sorted by x.  Tokens that are purely
    structural/office header words (per _is_stop_token) are excluded.
    Clustering is gap-based on consecutive x-positions (not center-distance),
    so a 3-part name like 'KRISHA PETERSEN DECOURSEY' spanning ~110px stays
    one cluster."""
    name_tokens = []
    for y, x, text in header_tokens:
        if not re.search(r"[A-Za-z]", text):
            continue
        if _is_stop_token(text):
            continue
        letters = _letters(text)
        if not letters:
            continue
        name_tokens.append((round(x), letters))
    if not name_tokens:
        return []
    name_tokens.sort(key=lambda t: t[0])
    clusters = [[name_tokens[0]]]
    for x, letters in name_tokens[1:]:
        if x - clusters[-1][-1][0] <= NAME_CLUSTER_GAP:
            clusters[-1].append((x, letters))
        else:
            clusters.append([(x, letters)])
    result = []
    for cl in clusters:
        cx = round(sum(t[0] for t in cl) / len(cl))
        letters = set()
        for _, l in cl:
            letters |= l
        result.append((cx, letters))
    return result


def _match_names(clusters, contests):
    """Match name clusters to summary candidates, best-match-first greedy.
    All (cluster, candidate) pairs with Jaccard >= NAME_JACCARD are sorted by
    score descending; each is assigned only if neither the cluster nor the
    candidate is already taken.  This ensures e.g. a Lisonbee cluster matches
    'Kapianne Lisonbee' (0.82) before 'Ginger Douglas' can claim it (0.54).
    Returns dict {contest_idx: [(cluster_x, candidate_name)]} and matched xs."""
    pairs = []
    for ci, (office, district, party, cands) in enumerate(contests):
        for cname in cands:
            clet = _letters(cname)
            if not clet:
                continue
            for cx, cletters in clusters:
                j = _jaccard(clet, cletters)
                if j >= NAME_JACCARD:
                    pairs.append((j, cx, ci, cname))
    pairs.sort(key=lambda p: -p[0])
    result = {}
    matched_xs = set()
    used_cands = set()  # (ci, cname)
    for (j, cx, ci, cname) in pairs:
        if cx in matched_xs or (ci, cname) in used_cands:
            continue
        matched_xs.add(cx)
        used_cands.add((ci, cname))
        result.setdefault(ci, []).append((cx, cname))
    return result, matched_xs


def _detect_sections(pages, contests=None):
    """Return list of sections, each:
      kind ('STAT'|'CONTEST'), ranges (list of (page, y_lo, y_hi)).
    STAT markers merge across pages (the statistics table spans pages 0-1 and
    its identifying Totals row is on page 1).  Each CONTEST marker is its own
    section — contest pages are self-contained (each reprints its full header
    and carries its own precinct rows), so no cross-page merging is needed and
    no contest-identity matching is required here (that happens per-section in
    _parse_contest via name-cluster matching)."""
    markers = _section_markers(pages)
    if not markers:
        return []

    sections = []
    cur = None
    for p, y, kind in markers:
        if kind == "STAT" and cur and cur["kind"] == "STAT":
            cur["ranges"].append((p, y, None))
        else:
            cur = {"kind": kind, "ranges": [(p, y, None)]}
            sections.append(cur)

    # Fill y_hi: next marker's y on the same page, else None (page end).
    all_markers = [(p, y) for p, y, _ in markers]
    for sec in sections:
        filled = []
        for (p, y_lo, _) in sec["ranges"]:
            y_hi = None
            for (mp, my) in all_markers:
                if mp == p and my > y_lo:
                    y_hi = my
                    break
            filled.append((p, y_lo, y_hi))
        sec["ranges"] = filled
    return sections


def _section_tokens(sec, pages):
    """Collect all tokens within a section's page/y ranges.  Each token is
    tagged with its source page (for same-page precinct matching)."""
    out = []
    for (p, y_lo, y_hi) in sec["ranges"]:
        page = pages[p]
        for l in page:
            if l["y"] < y_lo:
                continue
            if y_hi is not None and l["y"] >= y_hi:
                continue
            out.append((p, l["y"], l["x"], l["text"]))
    return out


# --------------------------------------------------------------------------- #
# Value/precinct extraction within a section
# --------------------------------------------------------------------------- #

def _codes_and_values(tokens, x_max=3400):
    """Return (codes, values) from section tokens.
    codes: list of (page, y, x, text) matching CODE_RE, x<250.
    values: list of (page, y, x, int) where _val() succeeds, 250<=x<x_max."""
    codes = []
    values = []
    for (p, y, x, text) in tokens:
        if x < 250:
            if CODE_RE.match(text):
                codes.append((p, y, x, text))
        elif x < x_max:
            v = _val(text)
            if v is not None:
                values.append((p, y, x, v))
    return codes, values


def _assign_values(codes, values, centers):
    """Assign each value to a precinct code (same page, nearest y within
    CODE_Y_TOL) and a column (nearest center within COL_TOL).
    Returns {code_text: {center: value}}."""
    by_page = {}
    for (p, y, x, v) in values:
        by_page.setdefault(p, []).append((y, x, v))
    result = {}
    for (p, cy, cx, ctext) in codes:
        row = {}
        pvals = by_page.get(p, [])
        for (vy, vx, v) in pvals:
            if abs(vy - cy) > CODE_Y_TOL:
                continue
            cc, d = _nearest(vx, centers)
            if d <= COL_TOL:
                row[cc] = v  # last write wins (right-aligned values cluster to 1)
        result[ctext] = row
    return result


# --------------------------------------------------------------------------- #
# Statistics section
# --------------------------------------------------------------------------- #

def _parse_statistics(sec, tokens, meta_totals, county):
    """Parse a STATISTICS section.  Returns list of meta rows (no precinct
    column, since these are county-level stats printed per precinct for
    reconciliation)."""
    codes, values = _codes_and_values(tokens, x_max=2600)  # exclude % columns
    if not values:
        return []
    centers = _column_centers([v[2] for v in values])

    # Identify the Totals row and match its values to meta targets by value,
    # then to column centers by x.
    totals_y = None
    for (p, y, x, text) in tokens:
        if x < 250 and text.upper().startswith("TOTAL"):
            totals_y = y
            break
    # meta column -> (office, party)
    col_meta = {}
    if totals_y is not None:
        # collect totals-row value tokens (within +/- 20px, covers the 2
        # visual sub-rows that the wide totals row wraps into)
        totals_vals = []
        for (p, y, x, v) in values:
            if abs(y - totals_y) <= 20:
                totals_vals.append((x, v))
        # match each meta target value to a totals value, then to column center
        used = set()
        for (office, party) in _META_TARGETS:
            target = meta_totals.get((office, party))
            if target is None:
                continue
            best_x, best_d = None, 1e9
            for (tx, tv) in totals_vals:
                if (tx, tv) in used:
                    continue
                if tv == target:
                    d = 0
                else:
                    continue
                # nearest column center to this totals value's x
                cc, cd = _nearest(tx, centers)
                if cd < best_d:
                    best_d, best_x = cd, cc
                    used.add((tx, tv))
            if best_x is not None:
                col_meta[best_x] = (office, party)

    # Fallback: if totals matching didn't identify all 7, assign by value-to-
    # column (sum each column across precincts, match sums to meta totals).
    if len(col_meta) < len(_META_TARGETS):
        col_sums = {c: 0 for c in centers}
        assigned = _assign_values(codes, values, centers)
        for ctext, row in assigned.items():
            for cc, v in row.items():
                col_sums[cc] = col_sums.get(cc, 0) + v
        used_offices = set(col_meta.values())
        for (office, party) in _META_TARGETS:
            target = meta_totals.get((office, party))
            if target is None or (office, party) in used_offices:
                continue
            best_c, best_d = None, 1e9
            for c in centers:
                if c in col_meta:
                    continue
                d = abs(col_sums.get(c, 0) - target)
                if d < best_d:
                    best_d, best_c = d, c
            if best_c is not None and best_d == 0:
                col_meta[best_c] = (office, party)
                used_offices.add((office, party))

    rows = []
    assigned = _assign_values(codes, values, centers)
    for ctext, row in assigned.items():
        for cc, (office, party) in col_meta.items():
            v = row.get(cc)
            if v is None:
                v = 0
            rows.append(C.meta_row(county, office, party, v, precinct=ctext))
    return rows


# --------------------------------------------------------------------------- #
# Contest section
# --------------------------------------------------------------------------- #

def _parse_contest(sec, tokens, contests, county):
    """Parse a CONTEST section.  Returns list of candidate rows."""
    codes, values = _codes_and_values(tokens)
    if not values:
        return []
    centers = _column_centers([v[2] for v in values])

    # Header name clusters (from tokens above the first data row on each page
    # in the section; candidate names only appear in headers).
    header_tokens = []
    section_pages = {p for (p, _, _) in sec["ranges"]}
    first_data_y = {}
    for p in section_pages:
        page_codes = [c[1] for c in codes if c[0] == p]
        if page_codes:
            first_data_y[p] = min(page_codes)
    for (p, y, x, text) in tokens:
        if p in first_data_y and y < first_data_y[p] - 20:
            header_tokens.append((y, x, text))
    clusters = _name_clusters(header_tokens)

    # Stat columns: find "Total Votes Cast" / "Contest Total" label x-positions,
    # match to nearest column centers.  PaddleOCR sometimes drops the space
    # ("TotalVotes Cast"), so match on individual words, not the full phrase.
    stat_centers = set()
    for (p, y, x, text) in tokens:
        t = text.upper()
        if ("TOTAL" in t and "VOTES" in t) or ("CONTEST" in t and "TOTAL" in t):
            cc, d = _nearest(x, centers)
            if d <= STAT_LABEL_TOL:
                stat_centers.add(cc)

    # Match name clusters to summary candidates -> candidate per column.
    matched, _ = _match_names(clusters, contests)
    # Build {center: (contest_idx, office, district, party, candidate)}.
    col_cand = {}
    for ci, hits in matched.items():
        office, district, party, _ = contests[ci]
        for (cx, cname) in hits:
            cc, d = _nearest(cx, centers)
            if d <= NAME_COL_TOL and cc not in stat_centers:
                col_cand[cc] = (ci, office, district, party, cname)

    # Group candidate columns by contest, find each contest's TVC column
    # (the stat column immediately right of the contest's rightmost candidate).
    contest_cols = {}  # ci -> list of candidate centers
    for cc, (ci, *_rest) in col_cand.items():
        contest_cols.setdefault(ci, []).append(cc)
    contest_tvc = {}  # ci -> TVC center
    for ci, ccols in contest_cols.items():
        rightmost = max(ccols)
        right_stats = sorted(c for c in stat_centers if c > rightmost)
        if right_stats:
            contest_tvc[ci] = right_stats[0]

    # TVC column centers (for missing-value inference): the stat column nearest
    # each candidate group.  We find TVC as the stat column whose values, summed,
    # equal the sum of candidate columns + over/under.  Simpler: TVC is the
    # first stat column to the right of the rightmost candidate in each contest.
    # For inference we just need "the TVC column for this precinct's contest" --
    # since each precinct row has all columns, TVC is the stat column that is
    # NOT Contest Total.  We identify per-contest TVC below.

    assigned = _assign_values(codes, values, centers)

    rows = []
    for ctext, row in assigned.items():
        for ci, ccols in contest_cols.items():
            tvc_center = contest_tvc.get(ci)
            tvc = row.get(tvc_center) if tvc_center else None
            assigned_sum = sum(row.get(c, 0) for c in ccols)
            missing = [c for c in ccols if c not in row]
            # If exactly one candidate column is missing a value and TVC is
            # present, infer the missing value (handles OCR-missed digits and
            # blank zero cells alike): missing = TVC - sum(assigned candidates).
            if len(missing) == 1 and tvc is not None:
                row[missing[0]] = tvc - assigned_sum
            for c in ccols:
                info = col_cand.get(c)
                if not info:
                    continue
                _, office, district, party, cname = info
                v = row.get(c, 0)
                rows.append(C.candidate_row(county, office, district, party,
                                            cname, v, precinct=ctext))
    return rows


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #

def parse_precinct(county, pdf_path):
    """Parse Box Elder's scanned precinct PDF into OpenElections precinct rows."""
    summary_csv = C.county_paths(county)[0]
    pages = C.ocr_pdf(pdf_path, county, dpi=400)
    meta_totals, contests = _load_summary(summary_csv)
    rows = []
    for sec in _detect_sections(pages, contests):
        tokens = _section_tokens(sec, pages)
        if sec["kind"] == "STAT":
            rows += _parse_statistics(sec, tokens, meta_totals, county)
        else:
            rows += _parse_contest(sec, tokens, contests, county)
    return rows