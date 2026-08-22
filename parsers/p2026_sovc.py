"""Parser for the 2D "Statement of Votes Cast" (SOVC) precinct format used by
Davis, Summit, and Uintah counties in the 2026 primary.

These precinct reports are wide tables: each row is a precinct, and columns
are grouped by contest (candidate vote columns followed by Total Votes Cast /
Overvotes / Undervotes, repeated for each contest on the page).  The contest
structure (office/district/party and candidate identities) is taken from that
county's already-parsed summary CSV; the SOVC table is consumed positionally
-- the N value columns of a contest are its N candidates (in the SOVC's own
left-to-right order, which can differ from the summary) plus Total (skipped),
Over Votes, and Under Votes.

Because the SOVC prints candidate names vertically above their columns in a
jumbled order (last name / first name / middle), we do NOT assume the summary
candidate order.  Instead the candidate name tokens printed in the header are
clustered by x-position and matched to summary candidates by token-set
overlap, so each value column is assigned to the correct candidate regardless
of the column order.

Page 1 of these reports is a STATISTICS table (Registered Voters / Ballots
Cast / Ballots Cast Blank / Turnout per precinct, broken out by party group);
it is parsed into meta rows.  Contest SOVC tables follow on later pages.

Text is extracted with pdfplumber word coordinates (embedded-text PDFs).
"""

import csv
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import p2026_common as C

import pdfplumber

_PARTY_CODES = ("REP", "DEM", "LIB", "IAP", "CON", "UUP", "IND", "GRN")
# Contest headers use either the short code ("REP U.S. House ...") or the full
# party name ("Republican Candidate for U.S. House ...").  Map every spelling to
# its code; the regex matches any as a standalone token (longer first so
# "DEMOCRATIC" is tried before "DEM" -- the \S boundaries already prevent
# partial matches, but ordering is belt-and-suspenders).
_PARTY_TOKEN_MAP = {
    "REP": "REP", "REPUBLICAN": "REP", "REPUBLICANS": "REP",
    "DEM": "DEM", "DEMOCRAT": "DEM", "DEMOCRATS": "DEM", "DEMOCRATIC": "DEM",
    "LIB": "LIB", "LIBERTARIAN": "LIB",
    "IAP": "IAP", "AMERICAN": "IAP",
    "CON": "CON", "CONSTITUTION": "CON",
    "UUP": "UUP", "UNITED": "UUP",
    "IND": "IND", "INDEPENDENT": "IND",
    "GRN": "GRN", "GREEN": "GRN", "GREENS": "GRN",
}
_PARTY_TOKEN_RE = re.compile(
    r"(?<!\S)(" + "|".join(sorted(_PARTY_TOKEN_MAP, key=len, reverse=True))
    + r")(?!\S)", re.IGNORECASE)
# Office keywords that must also appear on a contest-header line (guards the
# header search against matching a candidate whose surname happens to be a
# party word like "Green").
_OFFICE_KW = ("HOUSE", "SENATE", "DISTRICT", "COUNTY", "COUNCIL",
              "COMMISSION", "GOVERNOR", "PRESIDENT", "ATTORNEY", "TREASURER",
              "AUDITOR", "CLERK", "SCHOOL", "COURT", "CANDIDATE", "MAYOR",
              "SHERIFF", "RECORDER", "SURVEYOR", "ASSESSOR")
# A value token: digits with optional commas, NO percent sign.
_VAL_RE = re.compile(r"^[\d,]+$")

# Header tokens that are structural, not candidate-name parts.
_NAME_STOP = {
    "TOTAL", "TOTALS", "VOTES", "CAST", "VOTE", "FOR", "OVERVOTES",
    "UNDERVOTES", "OVER", "UNDER", "BLANK", "REGISTERED", "BALLOTS",
    "TURNOUT", "VOTER", "VOTERS", "REPUBLICAN", "REPUBLICANS", "DEMOCRAT",
    "DEMOCRATIC", "DEMOCRATS", "NONPARTISAN", "NONPARTISA", "PARTISAN",
    "NP", "N", "VOTE%", "STATISTICS", "STATISTIC", "OF", "PAGE",
}


def _load_summary_contests(summary_csv):
    """Return ordered list of contests: (office, district, party, [candidates]).
    Candidates are in summary order; Over/Under/meta rows are excluded."""
    rows = C._read_csv(summary_csv)
    contests = []
    cur = None
    for r in rows:
        office, district, party, cand = r["office"], r["district"], r["party"], r["candidate"]
        if office in (C.REGISTERED_VOTERS, C.BALLOTS_CAST, C.BALLOTS_CAST_BLANK):
            continue
        if cand in ("Over Votes", "Under Votes", ""):
            continue
        key = (office, district, party)
        if cur is None or (cur[0], cur[1], cur[2]) != key:
            cur = [office, district, party, []]
            contests.append(cur)
        cur[3].append(cand)
    return [(c[0], c[1], c[2], c[3]) for c in contests]


def _contest_key(office, district, party):
    return (office, district, party)


def _name_tokens(name):
    """Uppercase alpha tokens of a candidate name, for set matching."""
    return set(re.findall(r"[A-Za-z]+", name.upper()))


def _split_header_contests(words):
    """Split contest-header words into ordered [(office_raw, party), ...].

    Each contest begins at a party-token word (code or full name).  Contests
    sit side-by-side on the page, so a contest's wrapped header lines (e.g.
    "Council District 4" printed below "Summit County") are associated with it
    by x-position: a contest region spans from its party token's x to the next
    party token's x.  Words within a region are read in (top, x) order, and a
    leading "Candidate for " (Summit style) is stripped from the office text."""
    marks = sorted((w["x0"], w["text"]) for w in words
                   if w["text"].upper() in _PARTY_TOKEN_MAP)
    contests = []
    for i, (px, ptxt) in enumerate(marks):
        party = _PARTY_TOKEN_MAP[ptxt.upper()]
        hi = marks[i + 1][0] if i + 1 < len(marks) else 1e9
        region = [w for w in words
                  if px <= w["x0"] < hi and w["text"].upper() not in _PARTY_TOKEN_MAP]
        region.sort(key=lambda w: (round(w["top"]), w["x0"]))
        office = " ".join(w["text"] for w in region)
        office = re.sub(r"\s+", " ", office).strip()
        office = re.sub(r"^Candidate\s+for\s+", "", office,
                        flags=re.IGNORECASE).strip()
        if office:
            contests.append((office, party))
    return contests


def _row_words_grouped(page):
    """Return rows of word dicts grouped by 'top' (y), sorted top-to-bottom,
    each row sorted by x0."""
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False,
                               extra_attrs=[])
    rows = {}
    for w in words:
        key = round(w["top"] / 3) * 3
        rows.setdefault(key, []).append(w)
    return [sorted(rows[t], key=lambda w: w["x0"]) for t in sorted(rows)]


def _column_x_clusters(data_rows_words):
    """Return (col_xs, first_val_x): the sorted x0 positions of value columns
    (numeric, non-% tokens) across the given rows, clustered to the nearest
    few pixels.  The precinct-code column (a numeric token at the far left,
    e.g. "24" at x~23) is dropped: it is separated from the real value columns
    by a large gap."""
    xs = []
    for words in data_rows_words:
        for w in words:
            if _VAL_RE.match(w["text"]):
                xs.append(round(w["x0"]))
    if not xs:
        return [], None
    xs.sort()
    clusters = []
    for x in xs:
        # 20px tolerance: the "Totals" row's numbers land a few px off the
        # data rows' numbers, but real adjacent columns are >=~37px apart.
        if clusters and x - clusters[-1][-1] <= 20:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    col_xs = [round(sum(c) / len(c)) for c in clusters]
    # Drop the precinct-code column: a low-x cluster followed by a big gap.
    if col_xs and col_xs[0] < 80:
        for i in range(1, len(col_xs)):
            if col_xs[i] - col_xs[i - 1] > 60:
                col_xs = col_xs[i:]
                break
    return col_xs, (col_xs[0] if col_xs else None)


def _candidate_header_clusters(row_groups, header_idx, first_data_idx):
    """Collect candidate-name tokens from the header region (between the
    contest header line and the first data row).  Cluster by x-position so
    each candidate's name parts group together.  Returns a list of
    (cluster_x, token_set) sorted by x."""
    buckets = {}  # cluster_x -> list of (x, top, text)
    for i in range(header_idx + 1, first_data_idx):
        for w in row_groups[i]:
            t = w["text"].upper()
            # keep alphabetic tokens (incl. single-letter initials) that are
            # not structural labels
            if not re.search(r"[A-Za-z]", t):
                continue
            tok = t.strip(".")
            if tok in _NAME_STOP:
                continue
            x = round(w["x0"])
            # find/merge nearest cluster within 15px
            placed = False
            for cx in list(buckets):
                if abs(x - cx) <= 15:
                    buckets[cx].append((x, w["top"], tok))
                    # nudge center
                    placed = True
                    break
            if not placed:
                buckets[x] = [(x, w["top"], tok)]
    out = []
    for cx, items in buckets.items():
        items.sort(key=lambda it: (it[1], it[0]))
        toks = set(it[2] for it in items)
        mean_x = round(sum(it[0] for it in items) / len(items))
        out.append((mean_x, toks))
    out.sort(key=lambda c: c[0])
    return out


def _match_candidate(col_x, clusters, cand_token_sets, used):
    """Match a value column to a summary candidate by nearest header token
    cluster, then by token-set overlap.  `cand_token_sets` is list of
    (candidate_name, token_set).  `used` is a set of already-assigned
    candidate indexes.  Returns the candidate name or None."""
    # nearest cluster by x
    best_cluster = None
    best_dist = 1e9
    for cx, toks in clusters:
        d = abs(cx - col_x)
        if d < best_dist:
            best_dist = d
            best_cluster = toks
    if best_cluster is not None and best_dist <= 40:
        # best token-set overlap among unused candidates
        best_j = -1.0
        best_idx = None
        for i, (name, ts) in enumerate(cand_token_sets):
            if i in used:
                continue
            if not ts:
                continue
            inter = len(best_cluster & ts)
            if inter == 0:
                continue
            j = inter / len(best_cluster | ts)
            if j > best_j:
                best_j = j
                best_idx = i
        if best_idx is not None:
            used.add(best_idx)
            return cand_token_sets[best_idx][0]
    return None


def _is_statistics_page(row_groups):
    """True if the page begins with a STATISTICS table."""
    for words in row_groups[:6]:
        for w in words:
            if w["text"].upper() in ("STATISTICS", "STATISTIC"):
                return True
    return False


def _is_totals_precinct(precinct):
    """True if a precinct name is the report's Totals/summary row."""
    s = re.sub(r"\s+", " ", precinct).strip()
    if not s:
        return False
    first = s.split()[0].upper().strip(".,")
    return first in ("TOTALS", "TOTAL")


# --- STATISTICS page ------------------------------------------------------- #

_PARTY_WORD = {
    "TOTAL": "", "TOTALS": "",
    "REPUBLICAN": "REP", "REPUBLICANS": "REP",
    "DEMOCRAT": "DEM", "DEMOCRATS": "DEM", "DEMOCRATIC": "DEM",
    "NONPARTISAN": "NP", "NONPARTISA": "NP", "UNAFFILIATED": "",
    "LIBERTARIAN": "LIB", "INDEPENDENT": "IND",
    "AMERICAN": "IAP", "CONSTITUTION": "CON", "UNITED": "UUP",
    "GREEN": "GRN",
}


def _parse_statistics_page(row_groups, display, col_xs, first_val_x):
    """Parse a STATISTICS table page into per-precinct meta rows.

    Columns are grouped by x-gaps; each group is one stat type (Registered
    Voters / Ballots Cast / Ballots Cast Blank) spanning Total + per-party
    columns. Party is assigned per-column by matching each party-label word to
    its *nearest* column (greedy) — columns with no label land on "" (Total).
    This handles both layouts: Uintah marks the Total column with a "Total"
    label, while Summit labels every column in a group "Registered" and
    distinguishes party only by "Republican"/"Democratic" labels sitting
    between columns."""
    rows = []
    if not col_xs:
        return rows
    # collect header words (label + party) above the first data row
    label_words = []   # (x, text) Registered/Voters/Ballots/Cast/Blank/Turnout
    party_words = []   # (x, party_code)
    first_data_idx = None
    for i, words in enumerate(row_groups):
        if any(_VAL_RE.match(w["text"]) and w["x0"] >= first_val_x - 5 for w in words):
            first_data_idx = i
            break
    if first_data_idx is None:
        return rows
    for i in range(first_data_idx):
        for w in row_groups[i]:
            up = w["text"].upper().strip(".")
            x = round(w["x0"])
            if up in ("REGISTERED", "VOTERS", "VOTER", "BALLOTS", "CAST",
                      "BLANK", "TURNOUT"):
                label_words.append((x, up))
            if up in _PARTY_WORD:
                party_words.append((x, _PARTY_WORD[up]))

    # Group value columns by x-gap (>30px starts a new group). Each group is
    # one stat type spanning Total + per-party columns.
    groups = []
    for cx in col_xs:
        if groups and cx - groups[-1][-1] <= 30:
            groups[-1].append(cx)
        else:
            groups.append([cx])

    # Party per column: assign each party-label word to its nearest column
    # (greedy); a column keeps the closest label. Unlabeled columns -> "".
    col_party = {cx: "" for cx in col_xs}
    col_party_d = {cx: 1e9 for cx in col_xs}
    for px, p in party_words:
        best_cx, best_d = None, 1e9
        for cx in col_xs:
            d = abs(cx - px)
            if d < best_d:
                best_d, best_cx = d, cx
        if best_cx is not None and best_d <= 40 \
           and best_d < col_party_d[best_cx]:
            col_party[best_cx] = p
            col_party_d[best_cx] = best_d

    # Stat type per group; Blank/Turnout override per column.
    col_class = {}
    for grp in groups:
        reg_d = min((abs(lx - cx) for cx in grp for lx, t in label_words
                     if t in ("REGISTERED", "VOTERS", "VOTER")), default=1e9)
        bal_d = min((abs(lx - cx) for cx in grp for lx, t in label_words
                     if t == "BALLOTS"), default=1e9)
        if min(reg_d, bal_d) > 60:
            # no stat label near this group (e.g. a stray/Turnout cluster)
            for cx in grp:
                col_class[cx] = None
            continue
        grp_label = C.REGISTERED_VOTERS if reg_d <= bal_d else C.BALLOTS_CAST
        for cx in grp:
            blank_d = min((abs(lx - cx) for lx, t in label_words
                           if t == "BLANK"), default=1e9)
            turn_d = min((abs(lx - cx) for lx, t in label_words
                          if t == "TURNOUT"), default=1e9)
            if turn_d <= 20 and turn_d < min(reg_d, bal_d, blank_d):
                col_class[cx] = None
            elif blank_d <= 20:
                col_class[cx] = (C.BALLOTS_CAST_BLANK, "")
            else:
                col_class[cx] = (grp_label, col_party[cx])

    left_bound = col_xs[0] - 15
    pending_name = ""
    started = False
    for words in row_groups[first_data_idx:]:
        name_words = [w for w in words if w["x0"] < left_bound]
        val_words = [w for w in words if _VAL_RE.match(w["text"])]
        val_words.sort(key=lambda w: w["x0"])
        # Once stats data has begun, a row with alphabetic text in the candidate
        # region (x>=150) and no numeric values is contest content (candidate
        # names / Overvotes / contest Total).  Some Summit pages flow a second
        # stats sub-table after a contest section; stop here so its values --
        # which align with these columns and would duplicate precincts already
        # captured -- are not emitted as bogus Registered/Ballots rows.
        if started and not val_words and any(
                re.search(r"[A-Za-z]", w["text"]) and w["x0"] >= 150 for w in words):
            break
        if len(val_words) < 2:
            if name_words:
                pending_name = " ".join(w["text"] for w in name_words).strip()
            continue
        precinct = (" ".join(w["text"] for w in name_words).strip()
                    if name_words else pending_name)
        if not precinct or _is_totals_precinct(precinct):
            pending_name = ""
            continue
        started = True
        for w in val_words:
            x = round(w["x0"])
            best_cx, best_d = None, 1e9
            for cx in col_xs:
                d = abs(cx - x)
                if d < best_d:
                    best_d, best_cx = d, cx
            if best_cx is None or best_d > 15:
                continue
            cls = col_class.get(best_cx)
            if cls is None:
                continue
            label, party = cls
            rows.append(C.meta_row(display, label, party, w["text"],
                                    precinct=precinct))
        pending_name = ""
    return rows


# --- Contest SOVC page ---------------------------------------------------- #

def _parse_contest_page(row_groups, contests_by_key, display):
    """Parse one contest SOVC page; return list of rows."""
    if not row_groups:
        return []
    candidate_first_rows = [words for words in row_groups
                            if len([w for w in words if _VAL_RE.match(w["text"])]) >= 4]
    if not candidate_first_rows:
        return []
    col_xs, first_val_x = _column_x_clusters(candidate_first_rows)
    if first_val_x is None:
        return []

    first_data_idx = None
    for i, words in enumerate(row_groups):
        nums = [w for w in words if _VAL_RE.match(w["text"]) and w["x0"] >= first_val_x - 5]
        # >=4 numerics: a real data row has >=4 value columns; this skips the
        # "VOTE FOR 1" header row (two lone "1"s) and contest-header digits.
        if len(nums) >= 4:
            first_data_idx = i
            break
    if first_data_idx is None:
        return []

    header_idx = None
    for i in range(first_data_idx - 1, -1, -1):
        line = " ".join(w["text"] for w in row_groups[i])
        up = line.upper()
        # A contest header carries a party token AND an office keyword (the
        # keyword guards against matching a candidate row whose surname is a
        # party word).
        if _PARTY_TOKEN_RE.search(up) and any(k in up for k in _OFFICE_KW):
            header_idx = i
            break
    if header_idx is None:
        return []
    # Collect header words from header_idx through the "VOTE FOR" row (the
    # office text may wrap across several lines, and wrapped lines associate
    # with their contest by x-position, handled in _split_header_contests).
    header_words = []
    for i in range(header_idx, first_data_idx):
        line = " ".join(w["text"] for w in row_groups[i])
        up = line.upper()
        if i > header_idx and "VOTE" in up and "FOR" in up:
            break
        header_words.extend(row_groups[i])

    # (office, party) -> [(district, cands)] for a district fallback: some
    # headers omit the district number (e.g. "State House District" with the
    # "59" cut off at the page edge), so the (office, district, party) lookup
    # misses.  Fall back to the unique summary contest with that office+party.
    by_office_party = {}
    for (o, d, p), cands in contests_by_key.items():
        by_office_party.setdefault((o, p), []).append((d, cands))

    page_contests = []  # (office, district, party, [candidates])
    for office_raw, party in _split_header_contests(header_words):
        office = C.normalize_office(office_raw)
        district = C.parse_district(office_raw)
        cands = contests_by_key.get(_contest_key(office, district, party), [])
        if not cands:
            opts = by_office_party.get((office, party), [])
            if len(opts) == 1:
                district, cands = opts[0]
        page_contests.append((office, district, party, cands))

    expected = sum(len(c[3]) + 3 for c in page_contests)
    if expected == 0 or len(col_xs) < expected:
        return []

    # Stat-column labels (Overvotes/Undervotes/Total) in the header region.
    # Their x-positions classify the 3 trailing columns of each contest, since
    # the column order differs by county (Uintah: cand,Total,Over,Under;
    # Summit: cand,Over,Under,Total).
    stat_label_xs = {"OVER": [], "UNDER": [], "TOTAL": []}
    for i in range(header_idx + 1, first_data_idx):
        for w in row_groups[i]:
            up = w["text"].upper().strip(".")
            x = round(w["x0"])
            if up in ("OVERVOTES", "OVER"):
                stat_label_xs["OVER"].append(x)
            elif up in ("UNDERVOTES", "UNDER"):
                stat_label_xs["UNDER"].append(x)
            elif up == "TOTAL":
                stat_label_xs["TOTAL"].append(x)

    def _stat_role(cx):
        best_role, best_d = None, 1e9
        for role in ("OVER", "UNDER", "TOTAL"):
            for lx in stat_label_xs[role]:
                d = abs(lx - cx)
                if d < best_d:
                    best_d, best_role = d, role
        return best_role

    # Build candidate-name header clusters for the whole page.
    clusters = _candidate_header_clusters(row_groups, header_idx, first_data_idx)

    # Per-column role map: col_x -> (office, district, party, role, name).
    # role in {"cand","over","under","total"}; name set for candidate columns.
    col_role = {}
    col_offset = 0
    for office, district, party, cands in page_contests:
        n = len(cands)
        cand_cols = col_xs[col_offset:col_offset + n]
        stat_cols = col_xs[col_offset + n:col_offset + n + 3]
        col_offset += n + 3
        cand_token_sets = [(cn, _name_tokens(cn)) for cn in cands]
        used = set()
        for cx in cand_cols:
            name = _match_candidate(cx, clusters, cand_token_sets, used)
            col_role[cx] = (office, district, party, "cand", name)
        # fallback: fill unmatched candidate columns with remaining summary
        # candidates (in summary order) so every column gets a name.
        remaining = [cn for i, (cn, _) in enumerate(cand_token_sets)
                     if i not in used]
        ri = 0
        for cx in cand_cols:
            _o, _d, _p, _r, name = col_role[cx]
            if name is None and ri < len(remaining):
                col_role[cx] = (_o, _d, _p, _r, remaining[ri])
                ri += 1
        role_map = {"OVER": "over", "UNDER": "under", "TOTAL": "total"}
        for cx in stat_cols:
            col_role[cx] = (office, district, party,
                            role_map.get(_stat_role(cx), "total"), None)

    # Read data rows, matching each value word to its nearest column role.
    # This handles precincts that only participate in one of the page's
    # contests (suppressed/absent columns) -- each present value emits its row,
    # missing columns simply contribute nothing.
    rows = []
    pending_name = ""
    # Precinct names live in the left column (x < ~100); value columns start at
    # first_val_x (~188).  A contest-totals row is split across two text rows --
    # a "Totals" label in the name column, then the totals values on the next
    # row, left-shifted ~10px (e.g. "2,437" @179).  Using first_val_x-50 for the
    # name boundary leaves that left-shifted value in the gap (neither name nor
    # value), so the values row inherits the "Totals" pending name and is
    # skipped by _is_totals_precinct instead of being emitted as a bogus
    # precinct whose totals double the candidate counts.
    name_x_max = first_val_x - 50
    for words in row_groups[first_data_idx:]:
        name_words = [w for w in words if w["x0"] < name_x_max]
        val_words = [w for w in words if _VAL_RE.match(w["text"])
                     and w["x0"] >= first_val_x - 5]
        val_words.sort(key=lambda w: w["x0"])
        if not val_words:
            if name_words:
                pending_name = " ".join(w["text"] for w in name_words).strip()
            continue
        precinct = (" ".join(w["text"] for w in name_words).strip()
                    if name_words else pending_name)
        if not precinct or _is_totals_precinct(precinct):
            pending_name = ""
            continue
        for w in val_words:
            x = round(w["x0"])
            best_cx, best_d = None, 1e9
            for cx in col_xs:
                d = abs(cx - x)
                if d < best_d:
                    best_d, best_cx = d, cx
            if best_cx is None or best_d > 15:
                continue
            entry = col_role.get(best_cx)
            if entry is None:
                continue
            office, district, party, role, name = entry
            if role == "cand":
                if name in (None, "Over Votes", "Under Votes"):
                    continue
                rows.append(C.candidate_row(display, office, district, party,
                                             name, w["text"], precinct=precinct))
            elif role == "over":
                rows.append(C.overunder_row(display, office, district, party,
                                            "Over Votes", w["text"],
                                            precinct=precinct))
            elif role == "under":
                rows.append(C.overunder_row(display, office, district, party,
                                            "Under Votes", w["text"],
                                            precinct=precinct))
            # "total" role is not emitted
        pending_name = ""
    return rows


# --- Davis wide SOVC ------------------------------------------------------- #
# Davis prints one very wide table: a Statistics section (Registered Voters,
# Total Ballots Cast, Turnout, Election Day/Provisional/Early/Curbside/By Mail)
# on the left, then ~11 contests side-by-side, each with only candidate vote
# columns (no per-contest Over/Under/Total).  The header (contests + candidate
# names + stats labels) appears only on page 1; pages 2-7 are continuation data
# rows reusing the same column x-positions.  pdfplumber reads the data values and
# contest headers normally, but char-reverses every token in the candidate-name
# / stats-label band (e.g. "EEBNOSIL" = "LISONBEE"); reversing each token recovers
# the text.  Each candidate's name tokens cluster at their column's x-position,
# so we identify a column by its token SET matched to the summary candidates --
# the in-column reading order is unreliable and unnecessary.

_DAVIS_STATS_VOCAB = {
    "REGISTERED", "VOTERS", "VOTER", "TOTAL", "BALLOTS", "CAST", "TURNOUT",
    "ELECTION", "DAY", "PROVISIONAL", "EARLY", "CURBSIDE", "BY", "MAIL",
    "PRECINCT", "BLANK",
}
# Contest-header words (printed normally, above the reversed candidate band)
# that must not be mistaken for candidate-name tokens.
_DAVIS_HEADER_STOP = {
    "REP", "DEM", "LIB", "IAP", "CON", "UUP", "IND", "GRN", "U", "S", "US",
    "STATE", "HOUSE", "SENATE", "COUNTY", "COMMISSIONER", "SHERIFF", "CLERK",
    "DISTRICT", "SEAT", "CANDIDATE", "FOR", "VOTE", "ONE", "STATISTICS",
}


def _davis_row_groups(page):
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    rows = {}
    for w in words:
        key = round(w["top"] / 3) * 3
        rows.setdefault(key, []).append(w)
    return [sorted(rows[t], key=lambda w: w["x0"]) for t in sorted(rows)]


def _davis_first_data_top(row_groups):
    """top of first data row: a row with >=10 numeric tokens at x>80."""
    for words in row_groups:
        n = sum(1 for w in words if _VAL_RE.match(w["text"]) and w["x0"] > 80)
        if n >= 10:
            return round(words[0]["top"] / 3) * 3
    return None


def _davis_token_set(text):
    return set(re.findall(r"[A-Za-z]+", text.upper()))


def _davis_match(tokens, cands):
    """Best summary candidate (name, office, district, party) for a token set,
    by Jaccard overlap; None below 0.6."""
    best, best_j = None, 0.0
    for name, ts, office, district, party in cands:
        if not ts or not tokens:
            continue
        inter = len(tokens & ts)
        if inter == 0:
            continue
        j = inter / len(tokens | ts)
        if j > best_j:
            best_j, best = j, (name, office, district, party)
    return best if best_j >= 0.6 else None


def _davis_build_col_map(row_groups, cands):
    """Build {col_x: role} from page 1's header.
    role is ("cand", office, district, party, name) or ("meta", label, party).
    """
    first_top = _davis_first_data_top(row_groups)
    if first_top is None:
        return {}
    # The "VOTE FOR" row separates the (normal) contest headers above from the
    # (reversed) candidate-name / stats-label band below.
    vote_top = 0
    for words in row_groups:
        top = round(words[0]["top"] / 3) * 3
        if top >= first_top:
            break
        if any(w["text"].upper() == "VOTE" for w in words):
            vote_top = top

    # Candidate-name clusters by x (tokens reversed to real text).
    clusters = {}  # cluster_x -> list of (x, top, token)
    stats_labels = []  # (x, label) for stats columns
    for words in row_groups:
        top = round(words[0]["top"] / 3) * 3
        if top <= vote_top or top >= first_top:
            continue
        for w in words:
            tx = w["text"]
            if not re.search(r"[A-Za-z]", tx):
                continue
            real = tx[::-1].upper().strip(".")
            x = round(w["x0"])
            if x < 254:
                if real in _DAVIS_STATS_VOCAB:
                    stats_labels.append((x, real))
                continue
            if real in _DAVIS_HEADER_STOP:
                continue
            placed = False
            for cx in list(clusters):
                if abs(x - cx) <= 12:
                    clusters[cx].append((x, top, real))
                    placed = True
                    break
            if not placed:
                clusters[x] = [(x, top, real)]

    col_map = {}
    for cx, items in clusters.items():
        toks = set()
        for it in items:
            toks |= _davis_token_set(it[2])
        match = _davis_match(toks, cands)
        if match:
            name, office, district, party = match
            mean_x = round(sum(it[0] for it in items) / len(items))
            col_map[mean_x] = ("cand", office, district, party,
                               C.title_case_name(name))

    # Stats columns: Registered Voters + Total Ballots Cast.  Identify by label
    # x, then snap to the nearest stats value column.
    stats_val_xs = []
    for words in row_groups:
        top = round(words[0]["top"] / 3) * 3
        if top < first_top:
            continue
        for w in words:
            if _VAL_RE.match(w["text"]) and w["x0"] < 254:
                stats_val_xs.append(round(w["x0"]))
    stats_val_xs = sorted(set(stats_val_xs))

    def _nearest_label(label_set):
        xs = [x for x, lab in stats_labels if lab in label_set]
        return min(xs) if xs else None

    def _nearest_val(target_x):
        if target_x is None or not stats_val_xs:
            return None
        return min(stats_val_xs, key=lambda x: abs(x - target_x))

    reg_x = _nearest_val(_nearest_label({"REGISTERED", "VOTERS"}))
    bal_x = _nearest_val(_nearest_label({"BALLOTS", "TOTAL", "CAST"}))
    if reg_x is not None:
        col_map[reg_x] = ("meta", C.REGISTERED_VOTERS, "")
    if bal_x is not None and bal_x != reg_x:
        col_map[bal_x] = ("meta", C.BALLOTS_CAST, "")
    return col_map


def _davis_nearest_key(x, col_map):
    best_k, best_d = None, 1e9
    for k in col_map:
        d = abs(k - x)
        if d < best_d:
            best_d, best_k = d, k
    return best_k if best_d <= 15 else None


def _davis_is_totals(precinct):
    s = re.sub(r"\s+", " ", precinct).strip().upper()
    return s in ("TOTALS", "TOTAL", "COUNTY TOTALS", "COUNTY TOTAL")


def _davis_parse_page(row_groups, col_map, display):
    """Parse one page's data rows by assigning each value to its nearest
    left-column (x<80) label by vertical distance.  A precinct's name may
    print a few pixels above OR below its value row, so the sequential
    name-then-values assumption is unreliable; nearest-label assignment
    handles both orders.  A value whose nearest label is NOT a precinct code
    (e.g. the report's trailing "TOTAL" row, whose values are the county-wide
    sums) is skipped, so those totals are not re-emitted as a precinct's votes.
    """
    rows = []
    if not col_map:
        return rows
    first_top = _davis_first_data_top(row_groups)
    if first_top is None:
        return rows
    all_words = [w for words in row_groups for w in words]
    labels = [(round(w["top"]), w["text"])
              for w in all_words if w["x0"] < 80 and w["text"]]
    if not labels:
        return rows
    for w in all_words:
        if not _VAL_RE.match(w["text"]) or w["x0"] < 80:
            continue
        vt = round(w["top"])
        if vt < first_top - 6:
            continue
        ntop, ntext = min(labels, key=lambda lab: abs(lab[0] - vt))
        if abs(ntop - vt) > 10:
            continue
        if not ntext[0].isdigit():
            continue  # nearest label is TOTAL / a header, not a precinct code
        precinct = ntext
        if _davis_is_totals(precinct):
            continue
        k = _davis_nearest_key(round(w["x0"]), col_map)
        if k is None:
            continue
        role = col_map[k]
        if role[0] == "meta":
            _, label, party = role
            rows.append(C.meta_row(display, label, party, w["text"],
                                   precinct=precinct))
        else:
            _, office, district, party, name = role
            rows.append(C.candidate_row(display, office, district, party,
                                         name, w["text"], precinct=precinct))
    return rows


def _parse_precinct_davis(display, pdf_path, contests):
    # Global candidate list with token sets for column matching.
    cands = []
    for office, district, party, cand_names in contests:
        for cn in cand_names:
            cands.append((cn, _davis_token_set(cn), office, district, party))

    rows = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        col_map = {}
        for i, page in enumerate(pdf.pages):
            row_groups = _davis_row_groups(page)
            if not row_groups:
                continue
            if i == 0 or not col_map:
                col_map = _davis_build_col_map(row_groups, cands)
            rows.extend(_davis_parse_page(row_groups, col_map, display))
    return rows


def parse_precinct(county, pdf_path):
    display = C.COUNTY_DISPLAY.get(county, county)
    summary_csv, _ = C.county_paths(county)
    contests = _load_summary_contests(summary_csv)
    if county == "davis":
        return _parse_precinct_davis(display, pdf_path, contests)
    contests_by_key = {}
    for office, district, party, cands in contests:
        contests_by_key.setdefault(_contest_key(office, district, party), cands)

    rows = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            row_groups = _row_words_grouped(page)
            if not row_groups:
                continue
            if _is_statistics_page(row_groups):
                cand_rows = [w for w in row_groups
                             if len([x for x in w if _VAL_RE.match(x["text"])]) >= 4]
                col_xs, fv = _column_x_clusters(cand_rows)
                if fv is not None:
                    rows.extend(_parse_statistics_page(row_groups, display,
                                                        col_xs, fv))
            else:
                rows.extend(_parse_contest_page(row_groups, contests_by_key, display))
    return rows


if __name__ == "__main__":
    county = sys.argv[1] if len(sys.argv) > 1 else "uintah"
    src = C.find_source(county, "precinct")
    rows = parse_precinct(county, src)
    _, pcsv = C.county_paths(county)
    C.write_csv(rows, pcsv, precinct=True)
    issues = C.sanity_check(pcsv, precinct=True)
    print(f"{county}: {len(rows)} rows, issues={len(issues)}")
    for iss in issues[:20]:
        print(" ", iss)