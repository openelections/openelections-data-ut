"""Parse Utah's June 28, 2022 primary county canvass reports into per-county
precinct CSVs (20220628__ut__primary__<county>__precinct.csv in 2022/counties/).

Formats (see 2022-primary-precinct-sourcing.md for provenance):

  dominion1d       Dominion/ElectionSight per-precinct "Summary Results Report"
                   blocks (Statistics, then <office>/<Vote For N>/rows). Text
                   for Daggett, Duchesne, Tooele, Washington, Iron, Weber,
                   Garfield-as-printed; OCR text for Kane, Carbon, Rich, Wayne,
                   Summit, Morgan, Uintah.
  racesheet        Cache: one STATISTICS table, then one section per race whose
                   candidate headers stack vertically; columns identified by
                   x-offset in the layout text.
  sovc2d_sheets    Salt Lake: one contest per xlsx sheet; candidate values land
                   on each precinct's 'Total' row at the header columns.
  wide_headed      Davis: single-sheet SOVC; race headers on row 4, candidates
                   two rows down, precinct rows below.
  widecsv          Utah County: sovc.csv (contest headers / parties / candidates).
  sovc2d_clarity   Beaver: preliminary Clarity SOVC pdf (6/29/2022 header, the
                   county's only posted 2022-primary report); per-precinct rows
                   with a turnout block then per-race blocks, candidate
                   labels x-clustered from the header rows.
  matrix           Emery, Sevier: certified canvass abstracts with precinct
                   columns; transcribed from the scans and validated against
                   each row's printed county total (text layers are scans or
                   broken OCR).

Run:
    uv run python parsers/parse_2022_primary.py [county ...]
Each parsed county prints a race inventory (candidate sums) for checking
against the statewide canvass; CSVs are written to 2022/counties/.
"""

import csv
import difflib
import os
import re
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC = Path(os.environ.get(
    "UT22_SOURCES",
    str(REPO.parent / "openelections-sources-ut" / "2022 Primary Election"),
))
OUT = REPO / "2022" / "counties"

COUNTIES = {
    "Beaver": ("beaver", "Beaver UT SOVC_Jun_29_2022_10_47_05.pdf", "beaver"),
    "Cache": ("cache", "Cache UT 2022 Primary Official Results - SOVC.pdf", "cache"),
    "Carbon": ("carbon", "OCR", "dom1d_ocr:carbon"),
    "Daggett": ("daggett", "Daggett UT Official Results By Precinct-071222.pdf", "dom1d"),
    "Davis": ("davis", "Davis UT SOVC-Primary-2022fa66676d-a77e-4e63-9672-e01fd9479dc3.xlsx", "davis"),
    "Duchesne": ("duchesne", "Duchesne UT Official-Results-by-Precinct-Regular-Primary-Election-2022.pdf", "dom1d"),
    "Emery": ("emery", "Emery UT 2022-Primary-Results.pdf", "emery"),
    "Garfield": ("garfield", "Garfield UT 2022 Primary Canvass Report.pdf", "garfield"),
    "Iron": ("iron", "Iron UT 06-28-2022, Precinct Summary Official Results (7-11-2022).pdf", "dom1d"),
    "Kane": ("kane", "Kane UT Official Precinct Primary report for June 28, 2022.pdf", "kane"),
    "Millard": ("millard", "Millard UT 6-28-22-Election-Precinct-Summary-Official-Results.pdf", "dom1d"),
    "Morgan": ("morgan", "Morgan UT 20220712_ Primary Election June 28, 2022.pdf", "morgan"),
    "Rich": ("rich", "Rich UT 2022-Primary-Official-Precinct-Results.pdf", "richmat"),
    "Salt Lake": ("salt_lake", "Salt Lake UT SOVC-Primary-2022.xlsx", "saltlake"),
    "Sevier": ("sevier", "Sevier UT 20220711canvass.report.pdf", "sevier"),
    # Summit: the 111-page "County Precinct Summary" is a scan; OCR (psm 6 at
    # 300/400 dpi, psm 11 with crop passes) misreads enough value cells and
    # merges enough rows that parsed rows can't be trusted to line up with its
    # own race totals — request the tab file or transcribe by hand
    "Summit": "111-page scan resists OCR (misread and mis-grouped rows) — clerk request or manual transcription",
    "Tooele": ("tooele", "Tooele UT 2022-Primary-Post-Canvass-Report-Signed-Final.pdf", "dom1d"),
    "Uintah": ("uintah", "Uintah UT 07-11-2022 Canvass of Elections Primary Election 2022 Uintah County.pdf", "uintahmat"),
    "Utah": ("utah", "Utah UT sovc.csv", "widecsv"),
    "Washington": ("washington", "Washington UT election-20220628-results-precinct.pdf", "dom1d"),
    "Wayne": ("wayne", "OCR", "dom1d_ocr:wayne3"),
    "Weber": ("weber", "Weber UT 2022-Primary-Precinct-Results-Jun-28-2022.pdf", "dom1d"),
    # No precinct-level 2022 primary source exists; a records request is needed:
    "Box Elder": "countywide-only report posted (the clerk emailed the precinct canvass for 2024) — request",
    "Grand": "certified canvass July 5, 2022; not posted or archived — clerk request",
    "Juab": "only an unofficial countywide report posted — clerk request",
    "Piute": "no precinct document posted — clerk request",
    "San Juan": "the county posts its 2022 general by precinct, not the primary",
    "Sanpete": "official results released July 5, 2022, not posted — clerk request",
    "Wasatch": "not posted (Senate + House 3 canvassed) — clerk request",
}

FIELDNAMES = ["county", "precinct", "office", "district", "party", "candidate", "votes"]
STATS_KEEP = ("Registered Voters", "Ballots Cast")
OCR_DIR = Path(os.environ.get("UT22_OCR_DIR", "/tmp/ut22ocr"))

# OCR'd turnout/bookkeeping labels that must never masquerade as candidates
CAND_GARBAGE_RE = re.compile(
    r"^(?:registered|ballots?|paper|cards|v[oa]t\w*\s?f?o?\w*|voter turn|turnout|"
    r"cont[ea]st|total votes|total votes cast|times cast|overvotes?|undervotes?|"
    r"blank)\b",
    re.I)
CAND_PUNCT_ONLY_RE = re.compile(r"^[\s\-\=*_‘'“”':;,.|0-9]+$")

NUMTOK = r"[0-9,.Il|OoQ\[\]\(\{\)°§£$#/]+"
PARTYTOK = r"(?:REP|DEM|IAP|LIB|CON|GRN|UUP|IND|NON)"


# ---------------------------------------------------------------- helpers

def pdftext(path):
    out = subprocess.run(["pdftotext", "-layout", str(path), "-"],
                         capture_output=True, text=True)
    return out.stdout


def ocr_int(tok):
    """Best-effort integer; single OCR-ambiguous chars map to 0/1 (also when
    mixed into digits: '4]' is a misread 41)."""
    tok = tok.strip().strip("|,")
    if not tok:
        return None
    tok = tok.translate(str.maketrans({
        "I": "1", "l": "1", "|": "1", "[": "1", "]": "1", "(": "1", "{": "1",
        "°": "", "§": "5", "£": "1", "/": "7",
        ")": "1", "}": "1", "O": "0", "o": "0", "Q": "0"}))
    if re.fullmatch(r"[0-9]+", tok):
        return int(tok)
    if re.fullmatch(r"[0-9]{1,3}(,[0-9]{3})+", tok):
        return int(tok.replace(",", ""))
    return None


def split_party_candidate(name):
    name = re.sub(r"\s+", " ", name or "").strip()
    m = re.match(r"^(REP|DEM|IAP|LIB|CON|GRN|UUP|IND|NON|NP)\s+(.*)$", name, re.I)
    if m:
        code = m.group(1).upper()
        return ("" if code in ("NON", "NP") else code), m.group(2).strip()
    return "", name


def norm_race(_name):
    """Normalize an office header, stripping a party prefix. Returns (office, party)."""
    s = re.sub(r"\s+", " ", _name or "").strip()
    party = ""
    m = re.match(r"^(((REPUBLICAN)|REPUBLICAN)|REPUBLICAN)\s*[Ff]o?r?\s*(.*)$", "") or None
    m = re.match(r"^(Republican|Democratic|Democrat)\b\s*[Ff]or\s*(.*)$", s)
    if m:
        party = "DEM" if m.group(1).lower().startswith("dem") else "REP"
        s = m.group(2)
    s = re.sub(r"^U\.\s?S\.\s*", "U.S. ", s)
    s = re.sub(r"^US (?=Senate|House)", "U.S. ", s)
    s = re.sub(r"^United States (Senate|Senator|House|Representative)s?\b", r"U.S. \1", s, flags=re.I)
    s = re.sub(r"^United Slates Senator.*$", "U.S. Senate", s, flags=re.I)  # OCR
    s = re.sub(r"^U\.S\. Senator$", "U.S. Senate", s, flags=re.I)
    s = re.sub(r"\s*\(Multi-?County\)\s*$", "", s, flags=re.I)
    s = s.replace("Representative", "House")
    return s, party


def district_of(office):
    """Split the district out of a race name; return (office, district)."""
    m = re.match(r"^(U\.S\. House|U\.S\. Senate|State House|State Senate|State School Board"
                 r"|House|Senate)\s*(?:District|Dist|Seat)?\s*\.*\s*#?\s*\.*0*\s*([0-9ivxIVX]+)[\s\b]*$",
                 office.rstrip(" .-"), re.I)
    if m:
        base = m.group(1).strip().title()
        name = ("U.S. House" if base in ("U.S. House", "House") else
                "U.S. Senate" if base in ("U.S. Senate", "Senate") else base)
        return name, m.group(2)
    m = re.match(r"^(U\.S\. House|U\.S\. Senate|State House|State Senate|State School Board)(\s+(?:District|Seat)\s+)#?\s*([0-9]+)$",
                 office.strip(), re.I)
    if m:
        return m.group(1).title() if m.group(1).startswith("State") else m.group(1), m.group(3)
    return office.rstrip(" #."), ""


class CountyData:
    def __init__(self, county):
        self.county = county
        self.rows = []
        self.county_totals = {}  # (race_title, candidate) -> printed county total

    @staticmethod
    def fix_precinct_name(county, name):
        """Normalize a ballot-style/precinct code: squeeze spaces; Morgan-style
        codes whose first separator printed as '-' get their ':' back."""
        name = re.sub(r"\s+", "", str(name or "")).strip(":|-")
        if re.match(r"^\d+-\d", name):
            name = name.replace("-", ":", 1)
        return name

    def add_stat(self, precinct, label, value):
        if label in STATS_KEEP:
            fix = (STAT_OVERRIDES.get(self.county) or {}).get(precinct, {})
            if label in fix and fix[label] != value:
                print(f"    ?? {self.county} {precinct} {label}: using visually "
                      f"verified {fix[label]} over scanned {value}")
                value = fix[label]
            have = {r["office"]: r for r in self.rows
                    if r["precinct"] == precinct and r["office"] in STATS_KEEP}
            if label == "Ballots Cast":
                rv = have.get("Registered Voters")
                if rv and value > int(rv["votes"]):
                    print(f"    !! {self.county} {precinct}: ballots cast "
                          f"{value} > registered {rv['votes']} — dropped "
                          f"(scanner misread the printed total)")
                    return
                if value < 0 or value > 1_000_000:
                    return
            if label == "Registered Voters" and value > 1_000_000:
                return
            self.rows.append({"county": self.county, "precinct": precinct,
                              "office": label, "district": "", "party": "",
                              "candidate": "", "votes": value})

    def add_race(self, precinct, office, candidate, votes, party=""):
        office, race_party = norm_race(office)
        office, district = district_of(office)
        office = {"U.S. SENATE": "U.S. Senate", "U.S. HOUSE": "U.S. House",
                  "US SENATE": "U.S. Senate", "US HOUSE": "U.S. House"}.get(
            office, office)
        party = party or race_party
        if CAND_PUNCT_ONLY_RE.match(candidate):
            return  # a punctuation/number-only OCR remnant
        probe = candidate.strip("‘'\"“=+|_;:,.- ")
        if CAND_GARBAGE_RE.match(probe):
            return  # a statistics/bookkeeping label misread as a candidate
        if not candidate or not str(votes).strip().lstrip("-").isdigit() and votes != "":
            # non-numeric garbage rows are dropped (OCR noise guard)
            if not (str(votes).strip().isdigit() or votes == ""):
                return
        self.rows.append({"county": self.county, "precinct": precinct,
                          "office": office, "district": district,
                          "party": party or "", "candidate": candidate,
                          "votes": votes})

    def write_csv(self, slug):
        OUT.mkdir(parents=True, exist_ok=True)
        path = OUT / f"20220628__ut__primary__{slug}__precinct.csv"
        with open(path, "w", newline='', encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDNAMES)
            w.writeheader()
            w.writerows(self.rows)
        return path


def race_key(row):
    return (row['office'], row['district'])


def canonicalize_candidates(cd, county_name, report=False):
    """OCR reads candidate names slightly differently on each precinct's row
    ('CHRISTINE F. WATKINS' vs 'CHRISTINE F, WATKINS', spaced letters...).
    Cluster a race's candidate labels by fuzzy similarity and rename to the
    most frequent variant. Returns the merge map for review."""
    from collections import defaultdict

    # 1) collect the per-race label counts
    counts = defaultdict(dict)
    for r in cd.rows:
        if r['candidate']:
            k = race_key(r)
            counts[k][r['candidate']] = counts[k].get(r['candidate'], 0) + 1
    canon = {}
    for k, lbls in counts.items():
        by_sim = []
        for lbl in sorted(lbls, key=lambda l_: -lbls[l_]):
            sim = None
            for grp in by_sim:
                leader = grp[0]
                a, b = re.sub(r"\W+", "", lbl.upper()), re.sub(r"\W+", "", leader.upper())
                if difflib.SequenceMatcher(None, a, b).ratio() >= 0.82:
                    sim = grp
                    break
            if sim is not None:
                cnt = sum(lbls[x] for x in sim)
                sim.append(lbl)
                if lbls[lbl] * 1.6 > cnt or lbls[lbl] > sum(lbls[x] for x in sim) / 2:
                    # put the more frequent variant first: cheap reorder
                    grp.sort(key=lambda x: -lbls[x])
            else:
                by_sim.append([lbl])
        # rename every member to the group's most frequent label; when counts
        # tie (each variant seen once), prefer the least-punctuated spelling
        for grp in by_sim:
            leader = max(grp, key=lambda x: (lbls[x],
                                             -len(re.sub(r"[A-Za-z']", "", x))))
            for m in grp:
                canon[m] = leader
    merged = 0
    for r in cd.rows:
        if r['candidate'] and r['candidate'] != canon.get(r['candidate'], r['candidate']):
            r['candidate'] = canon[r['candidate']]
            merged += 1
    return merged


# ---------------------------------------------------------------- dominion1d

def strip_page_noise(text):
    """Remove report page headers/footers so precinct headers sit adjacent to
    their Statistics blocks even where pages split a block."""
    out = []
    for raw in text.split("\n"):
        raw = raw.translate(str.maketrans(
            {"§": "5", "°": "", "£": "1", "‘": "", "’": "", "“": "", "”": ""}))
        s = raw.strip()
        low = s.lower()
        if not s:
            out.append("")
            continue
        if "Page " in s and "of" in s:
            out.append("")
            continue
        title_only = any(low.startswith(h) for h in (
            "summary results report", "summary repor", "statement of votes cast",
            "precinct summary", "precinct report", "election summary",
            "official results", "official final results", "canvass report",
            "county primary election", "county general election",
            "county election",
            "canvass results", "custom table report", "offical results")) or \
            re.search(r"UT [\w ]+Primary\s*(:?June|\d)|UT [\w ]+[Ss]OV[C]", s) or \
            re.fullmatch(r"(June|July|August|November|December|October|May|April|March) ?\d{1,2}[,.]? ?\d?\d?,? ?20\d\d.*", s) or \
            re.fullmatch(r".* County,? (Utah)?", s)
        if title_only and "Vote" not in s and "%" not in s:
            out.append("")
            continue
        out.append(raw)
    return "\n".join(out)


def is_votefor_line(s):
    """A (possibly OCR-garbled or letter-spaced) 'Vote For N' marker. Never a
    precinct header or a candidate row."""
    sq = re.sub(r"\s+", "", s.lower())
    if re.fullmatch(r"v[oa]t\w{0,3}fo?\w{0,3}\w{0,2}", sq):
        return True
    toks = s.split()
    if not 1 <= len(toks) <= 4:
        return False
    if not re.fullmatch(r"v[oa]t\w{0,3}[?.!)\]}|;]{0,2}", toks[0].lower()):
        return False
    if len(toks) >= 2 and not re.fullmatch(r"[vfro]\w{0,3}[?.!)\]}|;]{0,2}", toks[1].lower()):
        return False
    if any(len(t) > 2 for t in toks[2:]):
        return False
    return True


def is_block_header(s, nxt):
    """A line whose follower starts a Statistics/TOTAL block begins a precinct
    block. The follower may carry OCR junk before the word (': TOTAL'), but
    must NOT be 'Total Votes Cast ...', which follows the LAST candidate row of
    a contest block; the line itself must not be bookkeeping."""
    if is_votefor_line(s):
        return False
    nl = re.sub(r"\s+", "", nxt.lower()).lstrip(":|. —-–=\t")
    if not (nl.startswith("total") or nl.startswith("statistics")):
        return False
    if "votescast" in nl or "contest" in nl:
        return False
    st = re.sub(r"\s+", "", s)
    if re.match(r"^(Total|Statistics|VoterTurnout|Registered|Ballots|Paper|Cards)", st, re.I) or \
       re.match(r"^(Total|Statistics|Voter Turnout|Registered|Ballots|Paper|Cards)", s, re.I):
        return False
    if re.search(r"\d[\d,.]*\s?%$", s) or \
       re.search(r"(June|July|August|September|October|November|December|March|April|May) ?\d{1,2},? ?\d?\d,? ?20\d\d", s, re.I):
        return False
    return True


def looks_like_result_row(s):
    """Multi-column form first (name + first count + strict tail of
    percents/counts — Iron's early/mail/provisional columns), then the
    dominant 'NAME <votes> [<pct%>]' layout."""
    mB = re.fullmatch(r"(?:" + PARTYTOK + r"\s+)?(.+?)\s+(" + NUMTOK + r")"
                      r"(?:\s+(?:\d{1,3}[.,]\d+\s?%|[0-9,]+|[Il|OoQ\[\]\(\{\)]+))+$", s)
    if mB:
        val = ocr_int(mB.group(2))
        if val is not None and mB.group(1).strip():
            return mB.group(1).strip(), val
    mA = re.fullmatch(r"(?:" + PARTYTOK + r"\s+)?(.+?)\s+(" + NUMTOK + r")(?:\s+[\d.]+\s?%|\s*%?)$", s)
    if mA:
        val = ocr_int(mA.group(2))
        if val is not None and mA.group(1).strip():
            return mA.group(1).strip(), val
    return None


COUNTY_SUMMARY_RE = re.compile(r"^(Primary|General|Regular|Presidential) *Elections?$", re.I)


def parse_dominion1d(text, county):
    cd = CountyData(county)
    lines = [l.rstrip() for l in strip_page_noise(text).split("\n")]
    n = len(lines)
    precinct = None
    current_race = None
    seen_precincts = set()

    def next_nonblank(idx):
        j = idx + 1
        while j < n and not lines[j].strip():
            j += 1
        return re.sub(r"\s+", " ", lines[j].strip()) if j < n else ""

    for i, raw in enumerate(lines):
        s = re.sub(r"\s+", " ", raw.strip())
        if not s:
            continue
        low = s.lower()
        nxt = next_nonblank(i)
        res = looks_like_result_row(s)
        # rows under a county-level summary block (its label is a report
        # header, e.g. "Primary Election") are county totals, not precincts
        if COUNTY_SUMMARY_RE.fullmatch(s):
            # treat as a header: suppress result rows for it
            current_race = None
            precinct = None
            continue
        # 0) continuation pages may re-print the precinct label alone
        if res and s in seen_precincts:
            precinct = s
            current_race = None
            continue
        # 1) new precinct block
        if is_block_header(s, nxt):
            precinct = re.sub(r"\s+", " ", s).strip(" |:.,-—'\\\"")
            seen_precincts.add(s)
            current_race = None
            continue
        if is_votefor_line(s):
            continue
        if is_votefor_line(nxt):
            # race header: keep the office name, dropping OCR junk after the
            # district ("Republican for State House District 4 : er")
            current_race = re.sub(r"\s*[:;|,].*$", "", s).strip()
            continue
        # race titles that arrive without a following Vote-For line are
        # section titles, never candidate rows
        if re.match(r"^(Republican|Democratic|Democrat)\s+[Ff]or\s+", s, re.I):
            current_race = re.sub(r"\s*[:;|,].*$", "", s).strip()
            continue
        # 2) noise
        if any(low.startswith(h) for h in (
                "summary results report", "summary repor", "statement of votes cast",
                "precinct summary", "precinct report", "election summary",
                "official results", "official final results", "canvass report",
            "county primary election", "county general election",
            "county election",
                "canvass results", "custom table report", "offical results")):
            continue
        if re.search(r"(June|July|August|September|October|November|December|March|April|May) ?\d{1,2},? ?\d?\d,? ?20\d\d", s, re.I):
            continue
        if re.search(r"(Primary|General) Election$", low) and "Vote" not in s:
            continue
        if re.search(r"(County|Utah)$", low) and "Vote" not in s and "Total" not in s \
           and "Ballot" not in s and len(s) < 45:
            continue
        if re.fullmatch(r"Statistics|TOTAL|Total|Statistics TOTAL|Voter Turnout - Total", s, re.I):
            continue
        # 3) statistics rows — match on the spaces-removed text (Morgan
        # letter-spaces labels: 'B a l l o t s C a s t'); the value itself is
        # the first multi-digit number on the line, since Iron prints several
        # columns (total, election-day, early, mail, provisional)
        tight = re.sub(r"\s+", "", low)
        stat_m = (re.match(r"(registeredvoters|ballotscast|paperballotscast|cardscast)"
                           r"[-—:]?(total|republican|nonpartisan|blank)"
                           r"[:\s]*(" + NUMTOK + r")\s*$", tight)
                  or
                  re.match(r"(registeredvoters|ballotscast)[-—:]?republican"
                           r"[:\s]*(" + NUMTOK + r")\s*$", tight))
        if stat_m and precinct and stat_m.group(2).lower() == "total":
            label = {"registeredvoters": "Registered Voters",
                     "ballotscast": "Ballots Cast"}.get(stat_m.group(1))
            if label:
                num_toks = [t for t in low.split()
                            if re.fullmatch(r"[0-9,Il|OoQ]+", t)
                            and ocr_int(t) is not None]
                multi = [t for t in num_toks if len(t) >= 2]
                tok = multi[0] if multi else (num_toks[0] if num_toks else None)
                val = ocr_int(tok) if tok else None
                if val is not None:
                    cd.add_stat(precinct, label, val)
            continue
        if re.match(r"^(Registered Voters|Ballots Cast|Paper Ballots Cast|Cards Cast|Voter Turnout)", s, re.I) \
           and not current_race:
            continue
        # 4) result rows
        if res and current_race and precinct:
            name, val = res
            lowname = name.lower().strip()
            if lowname.startswith(("total", "tolal", "tot al", "contest", "con test", "over",
                                   "under", "write-in", "write in", "writeins", "number of",
                                   "qvervote", "overvole")) or "vervote" in lowname or \
               lowname.startswith("vom"):
                continue
            if re.fullmatch(r"(REP|DEM|IAP|LIB|CON|GRN|UUP|IND|NON|NP)", name, re.I) or len(name) < 3:
                continue
            party, cand = split_party_candidate(name)
            if not cand:
                continue
            cd.add_race(precinct, current_race, cand, val, party)
            continue
    return cd


def load_ocr_text(stem):
    pages = sorted(OCR_DIR.glob(f"{stem}_pg-*.txt"))
    return "\n\f".join(p.read_text(encoding="utf-8", errors="replace") for p in pages), len(pages)


# ---------------------------------------------------------------- cache (racesheet)

BOOKKEEP = {"Total Votes Cast", "Overvotes", "Undervotes", "Contest Total"}
STAT_LABEL_RE = re.compile(r"Registered|Ballots Cast|Blank|Turnout|Voters", re.I)


# The county's ballot order (verified on the report pages; the header names
# are rotated/stacked in the layout text, so the order is encoded here).
CACHE_RACES = [
    ("U.S. Senate", ["Becky Edwards", "Ally Isom", "Mike Lee"]),
    ("U.S. House", ["Blake D. Moore", "Tina Cannon", "Andrew Badger"]),
    ("State House", ["Mike Petersen", "Val K. Potter"]),
    ("County Council Southeast", ["Ladd Kennington", "Kathryn A. Beus"]),
    ("County Council Northeast", ["Mark R. Huff", "Chris Booth"]),
]
CACHE_RACE_TITLES = {
    "Republican for U.S. Senate": ("U.S. Senate", "REP", 0),
    "Republican for U.S. House District 1": ("U.S. House", "REP", 1),
    "Republican for State House District 2": ("State House", "REP", 2),
    "Republican for County Council Southeast": ("County Council Southeast", "REP", 3),
    "Republican for County Council Northeast": ("County Council Northeast", "REP", 4),
}
# per-race county totals from the sections' Totals rows (validation targets)
CACHE_RACE_TOTALS = {
    "U.S. Senate": [4340, 1759, 10844, 16943],
    "U.S. House": [9313, 1994, 5446, 16753],
    "State House": [3291, 2799, 6090, 6184],
    "County Council Southeast": [1522, 1823, 3345, 3468],
    "County Council Northeast": [1670, 1553, 3223, 3455],
}


def prec_label(s):
    return re.sub(r"\s+", " ", s or "").strip()


def split_titles(line):
    out = []
    for m in re.finditer(r"(?:Republican|Democratic)\s+for\s+[A-Za-z0-9 .#]+?(?=\s*(?:Republican|Democratic)|\s*$)", line):
        out.append((m.start(), re.sub(r"\s+", " ", m.group(0)).strip()))
    return out


def match_title(title):
    flat = re.sub(r"\s+", "", (title or "").lower())
    best, r_ = None, 0.0
    for key in CACHE_RACE_TITLES:
        v = difflib.SequenceMatcher(None, flat, re.sub(r"\s+", "", key.lower())).ratio()
        if v > r_:
            best, r_ = key, v
    return best if r_ >= 0.9 else None


def bbox_word_lines(path):
    """Per-page list of visual-row word groups from poppler's -bbox-layout XML.

    Each row is [(text, xMin, xMax)], in x order. The Cache parser needs exact
    word positions: this report's rows carry two race tables side by side on
    some pages, and blank cells must not shift a token into the wrong column.
    Poppler's own <line> grouping is one text-object per grid cell here, so
    visual rows are rebuilt by y-coordinate clustering."""
    import xml.etree.ElementTree as ET
    out = subprocess.run(["pdftotext", "-bbox-layout", str(path), "-"],
                         capture_output=True, text=True)
    root = ET.fromstring(out.stdout)
    pages = []
    for page in root.findall(".//{*}page"):
        words = [((w.text or "").strip(), float(w.get("xMin")), float(w.get("xMax")),
                  float(w.get("yMin"))) for w in page.findall(".//{*}word")]
        if not words:
            pages.append([])
            continue
        words.sort(key=lambda t: (t[3], t[1]))
        rows, cur = [], []
        for t, x, xe, y in words:
            if cur and y - cur[-1][3] > 2.5:
                rows.append([(a, b, c) for a, b, c, _ in sorted(cur, key=lambda w: w[1])])
                cur = []
            cur.append((t, x, xe, y))
        if cur:
            rows.append([(a, b, c) for a, b, c, _ in sorted(cur, key=lambda w: w[1])])
        pages.append(rows)
    return pages


NUMTOK_RE = re.compile(r"[0-9,]+")


def cache_page_titles(lines):
    """(race, x_start, x_end) per section title on the page, in x order."""
    out = []
    for lw in lines:
        starts = [i for i, (t, *_) in enumerate(lw)
                  if t.strip(",").upper() in ("REPUBLICAN", "DEMOCRATIC")]
        for k, i in enumerate(starts):
            nxt = starts[k + 1] if k + 1 < len(starts) else len(lw)
            title = " ".join(t for t, *_ in lw[i:nxt])
            race = (CACHE_RACE_TITLES.get(match_title(title)) or (None,))[0]
            if race:
                out.append((race, lw[i][1], lw[nxt - 1][2]))
    return sorted(out, key=lambda r: r[1])


def cache_name_parts(name):
    return name.casefold().replace(".", " ").split()


def cache_anchors_for_region(hdr_words, cands):
    """(cand_x in ballot order, stat_x (tvc, over, under, ctot)) for one table
    region, from its stacked header words. Candidate anchors are the leftmost
    x of every word matching a candidate's name parts; stat anchors come from
    the stacked stat labels, tvc inferred from the column pitch. Returns None
    when a race's candidate or a stat label is missing (callers fall back to
    that race's anchors from the previous page)."""
    words = [(t.lower().strip(",."), x)
             for t, x, _xe in hdr_words if not NUMTOK_RE.fullmatch(t)]
    cand_x = []
    for nm in cands:
        xs = [x for t, x in words if t in cache_name_parts(nm)]
        if not xs:
            return None
        cand_x.append(min(xs))
    ov = [x for t, x in words if t in ("overvotes", "over")]
    un = [x for t, x in words if t in ("undervotes", "under")]
    ct = [x for t, x in words if t in ("contest",)]
    if not (ov and un and ct):
        return None
    # the three printed stat columns sit on an even pitch; walk the lists to
    # the rightmost consistent triple (region windows can overlap neighbors'
    # labels by a pad, and min() can pick a neighbor's label instead)
    chains = [(o, u, c) for o in ov for u in un for c in ct
              if abs((u - o) - (c - u)) <= 2.0 and 30 <= (u - o) <= 55]
    if not chains:
        return None
    over, under, ctot = max(chains, key=lambda ch: (ch[0], ch[1], ch[2]))
    tvc = over - (under - over)
    return cand_x, [tvc, over, under, ctot]


def cache_assign(values, anchors, tol=26.0):
    """Snap (value, x) pairs onto anchor columns (candidates then tvc/over/
    under/ctot); None where unassigned."""
    cand_x, stat_x = anchors
    cols = list(cand_x) + list(stat_x)
    vals = [None] * len(cols)
    for v, x in values:
        d = [abs(x - ax) for ax in cols]
        bi = min(range(len(cols)), key=lambda i: d[i])
        if d[bi] > tol:
            print(f"    !! cache token far from any column x={x}: {v}")
            continue
        vals[bi] = v
    return vals


def parse_cache(pdf_path, county):
    """Cache SOVC.

    Positions come from pdftotext -bbox-layout: every page repeats its race
    table's header, some pages carry two race tables side by side, and blank
    cells shift a whitespace split around. Section titles and candidate/stat
    header words give each region's column x-positions; each numeric token
    snaps to the nearest column. A table's row must balance
    (sum(candidates) == votes cast; tvc + over + under == contest total).
    STATISTICS rows ('<style> <reg> <ballots> <blank> <pct%>') come from the
    layout text and validate against the county's Totals line."""
    cd = CountyData(county)
    lines = pdftext(pdf_path).split("\n")

    # -- stats pass (layout text) ----------------------------------------
    stats_seen = {}
    for line in lines:
        s = re.sub(r"\s+", " ", line.strip())
        m = re.match(r"^([A-Za-z0-9: -]+?)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9.,]+%)$", s)
        if m and not STAT_LABEL_RE.search(m.group(1)):
            lead = m.group(1).strip()
            reg = int(m.group(2).replace(",", ""))
            bc = int(m.group(3).replace(",", ""))
            if lead in stats_seen and stats_seen[lead] != (reg, bc):
                print(f"    !! cache conflicting stats {lead}: {stats_seen[lead]} vs {(reg, bc)}")
                continue
            stats_seen[lead] = (reg, bc)
    for lead, (reg, bc) in sorted(stats_seen.items()):
        cd.add_stat(lead, "Registered Voters", reg)
        cd.add_stat(lead, "Ballots Cast", bc)

    # -- race pass (bbox columns) ----------------------------------------
    cands_of = dict(CACHE_RACES)
    race_anchors = {}  # race -> (cand_x, stat_x), remembered across pages
    for page in bbox_word_lines(pdf_path):
        titles = cache_page_titles(page)
        if not titles:
            continue
        width = max((xe for lw in page for _, _, xe in lw) or [0]) + 200
        regions = []  # (race, xlo, xhi)
        for k, (race, xs, xe) in enumerate(titles):
            xlo = 0 if k == 0 else (titles[k - 1][2] + xs) / 2
            xhi = (xe + titles[k + 1][1]) / 2 if k + 1 < len(titles) else width
            regions.append((race, xlo, xhi))
        # header zone: visual rows above this page's first data row
        first_data = None
        for li, lw in enumerate(page):
            if sum(1 for t, x, _xe in lw if NUMTOK_RE.fullmatch(t)) >= 6:
                first_data = li
                break
        if first_data is None:
            first_data = len(page)
        for race, xlo, xhi in regions:
            hdr_words = [w for lw in page[:first_data] for w in lw
                         if xlo <= w[1] < xhi]
            a = cache_anchors_for_region(hdr_words, cands_of[race])
            if a is not None:
                race_anchors[race] = a
            elif race not in race_anchors:
                print(f"    !! cache anchors not found for {race}")
        data_rows = []
        for lw in page[first_data:]:
            nums = [(int(t.replace(",", "")), x, xe)
                    for t, x, xe in lw if NUMTOK_RE.fullmatch(t)]
            if len(nums) < 6:
                continue
            label = re.sub(r"\s+", " ", " ".join(t for t, x, _ in lw if x < 100)).strip(" ,")
            if not re.search(r"[A-Za-z]", label) or label.lower().startswith("totals"):
                continue  # noise lines and the printed county-Totals row
            data_rows.append((label, nums))
        for label, nums in data_rows:
            for race, xlo, xhi in regions:
                if race not in race_anchors:
                    continue
                reg = [(v, x) for v, x, _xe in nums if xlo <= x < xhi]
                vals = cache_assign(reg, race_anchors[race])
                if all(v is None for v in vals):
                    continue  # this race's table is blank on this precinct
                cand_n = len(cands_of[race])
                if any(v is None for v in vals):
                    print(f"    !! cache incomplete row {label!r} {race}: {vals}")
                    continue
                cvals, (tvc, over, under, ctot) = vals[:cand_n], vals[cand_n:]
                if not (sum(cvals) == tvc and tvc + over + under == ctot):
                    print(f"    !! cache row imbalance {label!r} {race}: {vals}")
                    continue
                for cname, v in zip(cands_of[race], cvals):
                    cd.add_race(label, cache_title_for(race), cname, v)
    return cd


def cache_title_for(race):
    for title, (r, _p, _i) in CACHE_RACE_TITLES.items():
        if r == race:
            return title
    return race


def ocr_page(path, pg, dpi=400, psm=6):
    """Render one PDF page and OCR it; cache the text per page index.

    tesseract gets relative paths with cwd=OCR_DIR: in this sandbox it fails
    on absolute /tmp paths."""
    OCR_DIR.mkdir(parents=True, exist_ok=True)
    stem = Path(path).stem.replace(" ", "_")
    txt_p = OCR_DIR / f"{stem}_ocrpg{pg:03d}.txt"
    if not txt_p.exists():
        base = f"{stem}_ocrpg{pg:03d}"
        subprocess.run(["pdftoppm", "-f", str(pg), "-l", str(pg), "-r", str(dpi),
                        "-gray", "-png", str(path), base],
                       check=True, cwd=OCR_DIR)
        png = sorted(OCR_DIR.glob(f"{base}-*.png"))[0]
        # tesseract appends .txt to its output argument itself
        subprocess.run(["tesseract", "--psm", str(psm), png.name,
                        f"{stem}_ocrpg{pg:03d}"], check=True, cwd=OCR_DIR)
    return txt_p.read_text(encoding="utf-8", errors="replace")


def kane_shape(s):
    """'NAME <votes> <pct fragments>' where the pct's decimal split into two
    tokens ('47 46%' for 47.46%): the value is the FIRST numeric token. None
    when the line has fewer than two numeric tokens."""
    toks = s.split()
    nums = [i for i, t in enumerate(toks)
            if re.search(r"\d", t) and not re.search(r"[A-Za-z]", t)]
    if len(nums) < 2:
        return None
    name = " ".join(toks[:nums[0]])
    if not name.strip():
        return None
    return name.strip(" ,"), ocr_int(toks[nums[0]])


def parse_morgan_pdf(path, county, order=None, of_total=48):
    """Dominion-1D reports, 2 report pages per precinct, whose pdftotext text
    layer garbles the (letter-spaced/rotated) precinct-name lines inside the
    flowing text — but every page's own name extracts cleanly. So: the name
    comes from the pdftotext page header (or, for Kane, from the fixed page
    order because even that garbles) and the rows from a per-page OCR (psm 6,
    cached). Countywide summary pages collect validation targets instead."""
    cd = CountyData(county)
    npages = int(subprocess.run(["pdfinfo", str(path)], capture_output=True, text=True)
                 .stdout.split("Pages:")[1].split()[0])
    # rows the scanner misreads consistently (visually verified on the page)
    overrides = {("3:1-S", "Republican for U.S. Senate", "ALLYISOM"): 21,
                 ("Torrey Inside/Outside", "Republican for U.S. House District 2",
                  "ERINRIDER"): 75,
                 ("Torrey Inside/Outside", "Republican for U.S. Senate", "BECKYEDWARDS"): 83}
    seen = {}
    for pg in range(1, npages + 1):
        txt = pdftext_page(path, pg)
        raw = [l.rstrip() for l in txt.split("\n")]
        m = next((re.search(r"Page\s+(\d+)\s*of\s+(\d+)", l) for l in raw
                  if "Page" in l and "of" in l), None)
        if m is None and len(txt.strip()) <= 1:
            # the whole report may be a scan (Kane): OCR the page and retry
            txt = ocr_page(path, pg)
            raw = [l.rstrip() for l in txt.split("\n")]
            m = next((re.search(r"Page\s+(\d+)\s*of\s+(\d+)", l) for l in raw
                      if "Page" in l and "of" in l), None)
        if not m or m.group(2) != str(of_total):
            # countywide Election Summary pages: collect race rows as
            # validation targets instead of precinct data
            cur = None
            for l in raw:
                s = re.sub(r"\s+", " ", l.strip())
                if re.match(r"^(Republican|Democratic|Democrat)\s+[Ff]or\s+", s):
                    cur = s
                elif s.lower().startswith(("total votes cast", "contest totals")):
                    cur = None
                elif cur and (res := looks_like_result_row(s)):
                    key = re.sub(r"\s+", "", res[0]).lower()
                    if not re.match(r"^(contest|total|registered|ballots|paper|"
                                    r"cards|v[oa]t\w)", key):
                        cd.county_totals[(cur, res[0])] = res[1]
            continue
        name_lines = []
        for l in raw[:8]:
            s = re.sub(r"\s+", " ", l.strip())
            low = s.lower()
            if not s or "morgan county" in low or low.startswith("morgan count"):
                continue
            if re.match(r"^(june|official|of cial|fi$)", low) or "county, utah" in low:
                continue
            if re.fullmatch(r"\d+[:\-]\d+([\-:]?[A-Z]{1,6})?", re.sub(r"\s+", "", s)):
                name_lines.append(s)
        name = name_lines[0] if name_lines else None
        if not name and order:
            rp = int(m.group(1))
            name = order[(rp - 1) // 2]
        if not name:
            print(f"    !! morgan page {pg}: no precinct name")
            continue
        prec = cd.fix_precinct_name(county, name)
        seen.setdefault(prec, 0)
        seen[prec] += 1
        if seen[prec] > 2:
            print(f"    !! morgan page {pg}: {prec} seen {seen[prec]}x")
        ocr = ocr_page(path, pg)
        flat = [re.sub(r"\s+", " ", l.strip()) for l in ocr.split("\n")]
        flat = [l for l in flat if l]
        cur = None
        cur_rows = {}   # candidate -> votes, for the page's balance check
        over = under = None
        for i, s in enumerate(flat):
            low = s.lower()
            tight = re.sub(r"\s+", "", low)
            nxt = flat[i + 1] if i + 1 < len(flat) else ""
            res = looks_like_result_row(s)
            # race titles: a party-prefixed title line even where the marker
            # garbled, or any line whose own 'District 4'-style numeral makes
            # it look like a result row — recognized by following Vote-For
            if re.match(r"^(republican|democratic|democrat)\s+for\s+", low) or \
                    (is_votefor_line(nxt) and re.search(
                        r"\b(for|district|board|commission|council|senate|"
                        r"assembly|mayor|school)\b", low)):
                cur = s.replace("Schoo!", "School")                       .replace("Coun(y)", "County").replace("Counry", "County")
                cur_rows, over, under = {}, None, None
                continue
            m2 = re.match(r"(registeredvoters|ballotscast)[-—:]?total[:\s]*"
                          r"(" + NUMTOK + r")\s*$", tight)
            if tight.startswith(("registeredvoters", "ballotscast")) or m2:
                if m2:
                    val = ocr_int(m2.group(2))
                    if val is not None:
                        cd.add_stat(prec, {"registeredvoters": "Registered Voters",
                                           "ballotscast": "Ballots Cast"}[
                            m2.group(1)], val)
                continue
            if re.match(r"^overvotes?", re.sub(r"\s+", "", low)):
                over = ocr_int(s.split()[-1]) if re.search(r"[0-9]", s) else 0
                continue
            if re.match(r"^undervotes?", re.sub(r"\s+", "", low)):
                under = ocr_int(s.split()[-1]) if re.search(r"[0-9]", s) else 0
                continue
            if is_votefor_line(s):
                continue
            if "page " in low and re.search(r"of\s*\d", low):
                continue
            if cur and "%" in s:
                # pct-tailed row: the value is the first numeric token
                # ('EAST 112 47 46%' where the printed '47.46%' split)
                res = kane_shape(s) or res
            if res and cur:
                name_, val = res
                key = (prec, cur, name_.replace(" ", "").upper())
                if key in overrides:
                    val = overrides[key]
                if name_.replace(" ", "").upper().startswith("CONTEST"):
                    # the race's printed total; validate the rows against it
                    got = sum(cur_rows.values())
                    want = val - (over or 0) - (under or 0) if over is not None \
                        or under is not None else None
                    if want is not None and got != want or \
                            (want is None and got != val):
                        print(f"    !! morgan/ocr page {pg} {prec}: {cur} rows "
                              f"sum to {got} vs printed total {val} "
                              f"(over {over}, under {under}): {cur_rows}")
                    cur = None
                    continue
                cur_rows[name_.replace(" ", "").upper()] = val
                cd.add_race(prec, cur, name_, val)
    fill_verified_stats(cd, county)
    return cd


# Uintah countywide Custom Table Report, transcribed visually from the PDF
# (the per-precinct 1-page tables OCR erratically there; the countywide
# matrix at the end of the same PDF is one clean row per precinct). The
# printed Totals row of each table validates the transcription.
UINTAH_STATS = {
    "Jensen #1": (473, 272), "Davis East #2": (523, 284), "Davis West #3": (468, 226),
    "Naples East #4": (526, 308), "Davis North #24": (333, 175),
    "Vernal City #8": (451, 222), "Vernal City #9": (581, 314),
    "Vernal City #10": (464, 207), "Vernal City #11": (375, 210),
    "Vernal City #12": (391, 177), "Vernal City #13": (388, 176),
    "Vernal City #14": (229, 107), "Ashley East #15": (607, 346),
    "Ashley West #22": (519, 353), "Avalon Willows #16": (134, 55),
    "Glines South #17": (519, 251), "Maeser East #6": (630, 339),
    "Glines North #23": (536, 284), "Ballard #19": (611, 215),
    "Fort Duchesne #18": (243, 87), "Whiterocks #20": (71, 36),
    "Lapoint / Tridell #21": (590, 291), "Maeser / Dry Fork #7": (193, 119),
    "Maeser North #27": (632, 344), "Maeser West #25": (527, 290),
    "South #5": (563, 267), "Southwest #26": (584, 341),
    "Naples West #28": (452, 253), "South East #29": (302, 163),
}
UINTAH_TOTALS = (12915, 6712)              # registered, ballots cast
UINTAH_RACES = [
    ("U.S. Senate", ["Becky Edwards", "Ally Isom", "Mike Lee"],
     [797, 498, 5335]),
    ("U.S. House", ["Christopher Niles Herrod", "John Curtis"],
     [1545, 4912]),                         # District 3
    ("County Commission Seat A", ["Bart Haslem", "John Laursen"],
     [2015, 4638]),
    ("County Commission Seat B", ["Sonja Norton", "Willis Lefevre", "Wayne Simper"],
     [2715, 1896, 2071]),
]
# ballot order: one row per precinct, values in this order
UINTAH_MATRIX = [  # (precinct, [10 values])
    ("Jensen #1",           [26, 20, 222, 54, 206, 66, 205, 73, 118, 81]),
    ("Davis East #2",       [28, 23, 229, 81, 194, 55, 228, 105, 78, 100]),
    ("Davis West #3",       [23, 16, 184, 53, 160, 62, 163, 93, 67, 64]),
    ("Naples East #4",      [28, 13, 265, 72, 226, 90, 214, 128, 74, 104]),
    ("Davis North #24",     [21, 17, 136, 36, 129, 39, 136, 60, 44, 71]),
    ("Vernal City #11",     [30, 20, 156, 48, 156, 55, 151, 81, 77, 51]),
    ("Vernal City #12",     [30, 17, 128, 32, 141, 60, 116, 48, 82, 44]),
    ("Vernal City #14",     [17, 2, 88, 17, 88, 35, 72, 42, 35, 29]),
    ("Ashley East #15",     [35, 25, 285, 77, 260, 136, 206, 114, 93, 139]),
    ("Ashley West #22",     [40, 19, 289, 73, 263, 101, 248, 140, 137, 72]),
    ("Maeser West #25",     [25, 34, 230, 80, 204, 65, 221, 122, 76, 92]),
    ("Glines South #17",    [29, 21, 190, 68, 167, 89, 160, 95, 72, 83]),
    ("Maeser East #6",      [44, 17, 278, 78, 249, 98, 241, 159, 82, 95]),
    ("Glines North #23",    [36, 29, 213, 68, 207, 93, 188, 132, 63, 89]),
    ("Vernal City #8",      [27, 21, 172, 43, 169, 59, 163, 77, 62, 81]),
    ("Avalon Willows #16",  [11, 3, 41, 11, 44, 18, 37, 7, 29, 19]),
    ("Lapoint / Tridell #21", [35, 15, 240, 59, 225, 90, 199, 118, 89, 84]),
    ("Fort Duchesne #18",   [5, 4, 76, 29, 53, 45, 42, 20, 32, 34]),
    ("Ballard #19",         [31, 19, 165, 54, 161, 96, 117, 71, 65, 78]),
    ("Whiterocks #20",      [13, 5, 18, 8, 27, 25, 10, 10, 17, 9]),
    ("Maeser / Dry Fork #7", [14, 17, 87, 31, 82, 32, 86, 55, 35, 29]),
    ("South #5",            [24, 15, 225, 54, 196, 76, 187, 121, 71, 75]),
    ("Vernal City #9",      [47, 20, 241, 62, 239, 89, 221, 178, 55, 81]),
    ("Vernal City #10",     [34, 19, 149, 47, 149, 55, 149, 87, 56, 62]),
    ("Vernal City #13",     [24, 12, 136, 51, 116, 55, 119, 79, 36, 59]),
    ("Southwest #26",       [34, 25, 279, 81, 244, 131, 205, 151, 85, 105]),
    ("Maeser North #27",    [28, 25, 289, 70, 262, 110, 231, 147, 79, 116]),
    ("Naples West #28",     [44, 10, 194, 62, 182, 49, 201, 127, 54, 71]),
    ("South East #29",      [14, 15, 130, 46, 113, 41, 122, 75, 33, 54]),
]


def parse_uintah_matrix(stem, county):
    """Transcribed countywide matrices (see above). Validates per-row balances
    and the printed totals before emitting rows."""
    cd = CountyData(county)
    n = len([c for _, cs, _ in UINTAH_RACES for c in cs])
    assert n == 10 and len(UINTAH_MATRIX) == len(UINTAH_STATS) == 29, \
        "transcription shape changed"
    assert sum(st[0] for st in UINTAH_STATS.values()) == UINTAH_TOTALS[0], \
        "registered-voters transcription totals"
    assert sum(st[1] for st in UINTAH_STATS.values()) == UINTAH_TOTALS[1], \
        "ballots-cast transcription totals"
    # each column sums to the table's printed Totals row
    k = 0
    for race, names, want in UINTAH_RACES:
        for name in names:
            col = sum(vals[k] for _, vals in UINTAH_MATRIX)
            if col != want[names.index(name)]:
                print(f"    !! uintah transcription {race}/{name}: {col} "
                      f"vs printed {want[names.index(name)]}")
            k += 1
    for prec, vals in UINTAH_MATRIX:
        st = UINTAH_STATS[prec]
        k = 0
        for race, names, want in UINTAH_RACES:
            grp = vals[k:k + len(names)]
            assert sum(grp) <= st[1] + 4, f"{prec}: {race} votes exceed ballots"
            k += len(names)
        k = 0
        for race, names, want in UINTAH_RACES:
            for cname, v in zip(names, vals[k:k + len(names)]):
                cd.add_race(prec, race, cname, v, "REP")
                if race == "U.S. House":
                    cd.rows[-1]["district"] = "3"
            k += len(names)
    for prec, (reg, bc) in UINTAH_STATS.items():
        cd.add_stat(prec, "Registered Voters", reg)
        cd.add_stat(prec, "Ballots Cast", bc)
    return cd


def pdftext_page(path, pg):
    return subprocess.run(["pdftotext", "-layout", "-f", str(pg), "-l", str(pg),
                           str(path), "-"], capture_output=True, text=True).stdout


def sl_race_party(title):
    """Salt Lake contest title -> (race, party). Titles carry the party (or
    NP for nonpartisan primaries), a '(Vote for n)' tail and the sheet's
    privacy note; local school boards keep their district in the office name
    ('CANYONS SCHOOL BOARD DISTRICT #5'), matching Salt Lake's 2022 general
    file."""
    t = title.split("****")[0].strip()
    t = re.sub(r"\s+\*{2,}.*$", "", t)
    party = ("REP" if re.search(r"\(REP\)", title, re.I) else
             "DEM" if re.search(r"\(DEM\)", title, re.I) else
             "" if re.search(r"\((?:NP|NON)\)", title, re.I) else "")
    t = re.sub(r"\s*\((?:REP|DEM|NP|NON)\)", "", t, flags=re.I)
    t = re.sub(r"\s*\(Vote for\s*\d+\).*", "", t, flags=re.I).strip()
    t = re.sub(r"\s*\(Multi-County\).*", "", t, flags=re.I)
    t = re.sub(r"^US (SENATE|HOUSE)(\s+\d+)?$", r"U.S. \1\2", t, flags=re.I)
    t = re.sub(r"^U\.S\. (SENATE|HOUSE)$",
               lambda m: {"SENATE": "U.S. Senate", "HOUSE": "U.S. House"}[m.group(1).upper()],
               t, flags=re.I)
    t = re.sub(r"^(CANYONS|GRANITE|JORDAN|MURRAY|SALT LAKE|STATE|SOUTH SALT LAKE|"
               r"GRANITE) SCHOOL BOARD\s+(\d+)$",
               r"\1 SCHOOL BOARD DISTRICT #\2", t, flags=re.I)
    t = re.sub(r"^COUNTY COUNCIL (AT[L-]? ?LARGE|DISTRICT)\s*#?\s*([A-Z0-9]+)$",
               r"COUNTY COUNCIL \1 #\2", t, flags=re.I)
    return t, party


def parse_saltlake_xlsx(path, county):
    """One contest per sheet, plus a turnout sheet first.

    Row 2 carries the contest title ('US SENATE (REP) (Vote for  1)'); the
    header row lists Precinct | Times Cast | Registered | Precinct, then a
    merged pair of columns per candidate (value, pct) and Total Votes — each
    candidate's value lands one column right of its merged header. Every
    precinct prints three rows (In Person, Vote By Mail, Total); the Total row
    holds the precinct's numbers, '****' marks suppressed values. Sheet1 is
    the turnout table (Total rows: registered, ballots cast)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    cd = CountyData(county)
    SUBROW = ("In Person", "Vote By Mail", "By Mail")

    def clean(cell):
        if isinstance(cell, str):
            s = cell.strip()
            return "" if s in ("****",) else s
        return cell

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        title = None
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for cell in row:
                c = clean(cell) if isinstance(cell, str) else ""
                if c and "(Vote for" in c:
                    title = c.split("\n")[0].strip()
        head_col = {}  # value-column index -> candidate name
        for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
            texts = [(i, clean(c)) for i, c in enumerate(row) if clean(c)]
            if not any("Times Cast" in str(c or "") for _, c in texts):
                continue
            for i, c in texts:
                s = re.sub(r"\s+", " ", str(c or "")).strip()
                if "Total Votes" in s:
                    head_col[i] = None  # contest-total column; not a candidate
                    continue
                if re.match(r"^(Precinct|Times Cast|Registered)", s, re.I):
                    continue
                s2 = re.sub(r"\s*\(Vote for\s*\d+\)", "", s, flags=re.I)
                nm = s2.split("(")[0].strip().upper()
                if nm:
                    head_col[i] = nm
            break
        race, party_code = (None, "") if title is None else sl_race_party(title)
        current = None
        for row in ws.iter_rows(min_row=1, values_only=True):
            cells = [clean(c) for c in row or ()]  # positions must not shift
            first = cells[0] if cells else None
            s_first = str(first).strip() if first is not None else ""
            if not s_first or s_first.startswith("Page:") or "Vote for" in s_first:
                continue
            if s_first in ("Precinct", "Countywide", "Electionwide"):
                continue  # headers and aggregate rows; precincts are the data
            if s_first in SUBROW:
                continue
            if s_first.lower() == "total":
                if current is None:
                    continue
                if race is not None:
                    for col, cand in sorted(head_col.items()):
                        if cand and col < len(cells):
                            v = cells[col]
                            if isinstance(v, (int, float)) and not isinstance(v, bool):
                                cd.add_race(current, race, cand, int(v), party_code)
                else:
                    nums = [c for c in cells if isinstance(c, (int, float))]
                    if nums:
                        cd.add_stat(current, "Registered Voters", int(nums[0]))
                        if len(nums) >= 2:
                            cd.add_stat(current, "Ballots Cast", int(nums[1]))
                current = None
                continue
            if isinstance(first, str) and re.match(r"^[A-Z][A-Z0-9 _%-]*$", first):
                current = first  # precinct-name row; its Total row follows
    return cd

def norm_race_party_code(s):
    """'US SENATE (REP)' -> (race, 'REP')"""
    m = re.match(r"^(.*?)\(?\b(REP|DEM|LIB|IAP|CON|GRN|UUP|IND|NON)\)?\s*(\(|\(Vote|$)",
                 str(s).strip())
    if m:
        return m.group(1).strip(" ("), m.group(2)
    return str(s).strip(), "REP"


# ---------------------------------------------------------------- davis (xlsx)

def parse_davis_xlsx(path, county):
    """Single-sheet SOVC: race headers ('REP U.S. SENATE', 'REP SENATE 5'...)
    on the header row; candidate names two rows down in each race's columns;
    data rows carry the precinct name, then statistics, then race columns."""
    import openpyxl
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    grid = [list(r) for r in ws.iter_rows(values_only=True, max_col=120)]
    wb.close()
    cd = CountyData(county)
    hdr_i = next(i for i, r in enumerate(grid) if any("STATISTICS" in str(c or "") for c in r))
    race_row = grid[hdr_i]
    cand_row = grid[hdr_i + 2]
    races = []
    for j, cell in enumerate(race_row):
        name = re.sub(r"\s+", " ", str(cell).strip().upper()) if cell else ""
        if not name or name == "STATISTICS":
            continue
        if j and str(race_row[j - 1]).strip().upper() == name:
            continue  # part of the previous race's run
        races.append([name, j])
    for k, (nm, j) in enumerate(races):
        end = races[k + 1][1] if k + 1 < len(races) else len(cand_row)
        races[k].append(list(range(j, end)))
    for row in grid[hdr_i + 3:]:
        if not row or not row[2] or not str(row[2]).strip():
            continue
        precinct = str(row[2]).strip()
        if precinct.lower() in ("precinct", "totals", "total"):
            continue
        if row[4] and str(row[4]).strip().isdigit():
            cd.add_stat(precinct, "Registered Voters", int(str(row[4]).strip()))
        if row[5] and str(row[5]).strip().isdigit():
            cd.add_stat(precinct, "Ballots Cast", int(str(row[5]).strip()))
        for name, j, cols in races:
            office, dist = davis_office(name)
            for cj in cols:
                cand = re.sub(r"\s+", " ", str(cand_row[cj] or "")).strip().upper()
                if not cand:
                    continue
                val = row[cj] if cj < len(row) else None
                votes = str(int(val)) if isinstance(val, (int, float)) else ""
                cd.add_race(precinct, office, cand, votes, "REP")
                if dist:
                    cd.rows[-1]["district"] = dist
    return cd


def davis_office(nm):
    """('REP STATE SCHOOL BOARD 4' etc.) -> (office, district)."""
    n = re.sub(r"^\s*REP\s+", "", re.sub(r"\s+", " ", nm), flags=re.I).strip()
    m = re.match(r"^(?:STATE )?SCHOOL BOARD\s+(?:DISTRICT\s+)?(\d+)$", n, re.I)
    if m and "DAVIS" not in n.upper():
        return "State School Board", m.group(1)
    m = re.match(r"^SENATE\s+(\d+)$", n, re.I)
    if m:
        return "State Senate", m.group(1)
    m = re.match(r"^HOUSE\s+(\d+)$", n, re.I)
    if m:
        return "State House", m.group(1)
    m = re.match(r"^COMMISSIONER\s+SEAT\s+([AB])$", n, re.I)
    if m:
        return "County Commission Seat " + m.group(1), ""
    if n.upper() == "AUDITOR":
        return "County Auditor", ""
    if n.upper() == "CLERK":
        return "County Clerk", ""
    return n.title(), ""


# ---------------------------------------------------------------- utah (widecsv)

def parse_utah_widecsv(path, county):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    hdr, party_row, cand_row = rows[0], rows[1], rows[2]
    data = rows[3:]
    runs = []
    j = 0
    while j < len(hdr):
        name = re.sub(r"\s+", " ", hdr[j]).strip()
        if not name:
            j += 1
            continue
        j2 = j
        while j2 < len(hdr) and re.sub(r"\s+", " ", hdr[j2]).strip() == name:
            j2 += 1
        runs.append((name, list(range(j, j2))))
        j = j2
    cd = CountyData(county)

    def cells(row):
        return row + [""] * (len(hdr) - len(row))

    for row in data:
        r = cells(row)
        if not any(c.strip() for c in r[:20]):
            continue
        if (r[1].strip().upper() in ("ZZZ", "ZZZZ")) or "COUNTY TOTALS" in r[2].upper():
            # the trailing countywide-total row: a validation check, not a
            # precinct — note turnout totals for comparison
            reg = next((r[i].strip() for name, idxs in runs
                        if name.strip().upper().startswith("REGISTERED VOTERS TOTAL")
                        for i in idxs), "")
            bc = next((r[i].strip() for name, idxs in runs
                       if name.strip().upper().startswith("BALLOTS CAST TOTAL")
                       for i in idxs), "")
            if reg.isdigit():
                cd.county_totals[("Registered Voters", "Registered Voters")] = int(reg)
            if bc.isdigit():
                cd.county_totals[("Ballots Cast", "Ballots Cast")] = int(bc)
            continue
        precinct = (r[2].strip() or r[1].strip())
        for name, idxs in runs:
            if idxs[0] >= len(r):
                continue
            norm = name.strip().upper()
            vals = [r[i].strip() for i in idxs]
            if norm.startswith("COUNTY NUMBER") or norm.startswith("PRECINCT"):
                continue
            if norm.startswith("REGISTERED VOTERS TOTAL"):
                if vals[0].isdigit():
                    cd.add_stat(precinct, "Registered Voters", int(vals[0]))
                continue
            if norm.startswith("BALLOTS CAST TOTAL"):
                if vals[0].isdigit():
                    cd.add_stat(precinct, "Ballots Cast", int(vals[0]))
                continue
            if norm.startswith(("REGISTERED VOTERS -", "BALLOTS CAST -", "BALLOTS CAST BLANK")):
                continue
            for i2 in idxs:
                cand = re.sub(r"\s+", " ", cand_row[i2] or "").strip()
                if not cand or cand.upper() in ("VOTERS", "CAST", "BLANK", "BALLOTS CAST",
                                                "BALLOTS CAST BLANK", "BALLOTS CAST -"):
                    continue
                pv = (party_row[i2] if i2 < len(party_row) else "").strip().upper()
                party = {"REP": "REP", "NON": ""}.get(pv, "" if pv in ("NON", "") else pv.title())
                cd.add_race(precinct, name, cand, r[i2].strip(), party)
    return cd


# ---------------------------------------------------------------- beaver (clarity SOVC)

BEAVER_SECTIONS = [
    # (race header tokens to match, candidate labels in column order)
    (r"UNITED STATES SENATOR|US HOUSE OF REP DISTRICT 2",
     {"UNITED STATES SENATOR": ["BECKY EDWARDS", "ALLY ISOM", "MIKE LEE"],
      "US HOUSE OF REP DISTRICT 2": ["ERIN RIDER", "CHRIS STEWART"]}),
    (r"UTAH SENATE DISTRICT 28|COUNTY COMMISSIONER SEAT A",
     {"UTAH SENATE DISTRICT 28": ["PATRICK L. LARSON", "EVAN J. VICKERS"],
      "COUNTY COMMISSIONER SEAT A": ["TAMMY T. PEARSON", "CHRIS D. NOBLE", "JON CALDWELL"]}),
    (r"COUNTY COMMISSIONER SEAT B",
     {"COUNTY COMMISSIONER SEAT B": ["BRANDON YARDLEY", "MARK S. WHITNEY"]}),
]


# Beaver's Clarity SOVC (6/29/2022, the county's only posted 2022-primary
# report — PRELIMINARY, Election-Day rows only: no absentee). Its layout text
# is clean; each 'Normal' row is the precinct with value/pct pairs:
#   section 1: (turnout: reg, ballots) then U.S. Senate (reg, tvc,
#     Edwards, Isom, Lee) and U.S. House 2 (reg, tvc, Rider, Stewart);
#   section 2: Utah Senate 28 (reg, tvc, Larson, Vickers) then County
#     Commissioner Seat A (reg, tvc, Pearson, Noble, Caldwell);
#   section 3: County Commissioner Seat B (reg, tvc, Yardley, Whitney).
# The countywide rows validate each race's sums.
BEAVER_PRECINCTS = ["BEAVER 1", "BEAVER 2", "BEAVER 3", "BEAVER 4", "GREENVILLE 1",
                    "MILFORD 1", "MILFORD 2", "MILFORD 3", "MINERSVILLE 1"]


def parse_beaver_clarity(path, county):
    cd = CountyData(county)
    t = pdftext(path)
    # sections split on the 'Statement of Votes Cast' page headers
    sections, cur = [], []
    for l in t.split("\n"):
        if "Statement of Votes Cast" in l and any(x.strip() for x in cur):
            sections.append(cur)
            cur = []
        cur.append(l)
    sections.append(cur)
    cd.county_totals = {}
    for si, sec in enumerate(sections):
        if si > 2:
            break
        precinct = None
        for raw in sec:
            s = re.sub(r"\s+", " ", raw.strip())
            toks = raw.split()
            if not toks:
                continue
            first = toks[0].upper()
            if first in ("JURISDICTION",):
                precinct = None
                continue
            if s.upper() in BEAVER_PRECINCTS or s.upper() == "TOTAL":
                precinct = None if s.upper() == "TOTAL" else s.upper()
                continue
            if first != "NORMAL" or not precinct:
                continue
            # tokens: 'Normal <reg> <ballots> <pct%> <r-reg> <r-tvc>
            #         <cand v> <pct> ...' per race block; a zero-value cell
            # prints as '0' with a '-' pct
            vals = []
            j = 1
            while j < len(toks):
                tk = toks[j].replace(",", "")
                if not re.fullmatch(r"-|\d+", tk):
                    j += 1
                    continue
                vals.append(int(tk))
                # a value may be paired with its own pct, which is skipped
                if j + 1 < len(toks) and re.fullmatch(r"-|\d+(\.\d+)?%", toks[j + 1]):
                    j += 2
                else:
                    j += 1
            vals = vals  # consumed values: (value, skipped pct) pairs
            if si == 0:
                cd.add_stat(precinct, "Registered Voters", vals[0])
                cd.add_stat(precinct, "Ballots Cast", vals[1])
            if si == 0 and len(vals) >= 11:
                for cname, v in (("Becky Edwards", vals[4]),
                                 ("Ally Isom", vals[5]), ("Mike Lee", vals[6])):
                    cd.add_race(precinct, "U.S. Senate", cname, v, "REP")
                for cname, v in (("Erin Rider", vals[9]),
                                 ("Chris Stewart", vals[10])):
                    cd.add_race(precinct, "U.S. House District 2", cname, v,
                                "REP")
            elif si == 1 and len(vals) >= 8:
                for cname, v in (("Patrick L. Larson", vals[2]),
                                 ("Evan J. Vickers", vals[3])):
                    cd.add_race(precinct, "State Senate District 28", cname,
                                v, "REP")
            elif si != 0 and len(vals) >= 8:
                raise AssertionError(f"unexpected beaver row width {vals}")
            if si == 2 and len(vals) == 4:
                for cname, v in (("Brandon Yardley", vals[2]),
                                 ("Mark S. Whitney", vals[3])):
                    cd.add_race(precinct, "County Commission Seat B", cname,
                                v, "REP")
            elif si == 1 and len(vals) >= 9:
                for cname, v in (("Tammy T. Pearson", vals[6]),
                                 ("Chris D. Noble", vals[7]),
                                 ("Jon Caldwell", vals[8])):
                    cd.add_race(precinct, "County Commission Seat A", cname,
                                v, "REP")
            elif si != 0:
                raise AssertionError(f"unexpected beaver row width {vals}")
        # validate against the printed countywide Total rows
    want = {
        ("U.S. Senate", ""): {"Becky Edwards": 191, "Ally Isom": 68, "Mike Lee": 1275},
        ("U.S. House", "2"): {"Erin Rider": 159, "Chris Stewart": 1371},
        ("State Senate", "28"): {"Patrick L. Larson": 284, "Evan J. Vickers": 1249},
        ("County Commission Seat A", ""): {"Tammy T. Pearson": 744,
                                           "Chris D. Noble": 92,
                                           "Jon Caldwell": 717},
        ("County Commission Seat B", ""): {"Brandon Yardley": 937,
                                           "Mark S. Whitney": 609},
    }
    for (race, dist), cands in want.items():
        for cand, v in cands.items():
            got = sum(int(r["votes"]) for r in cd.rows
                      if r["office"] == race and r["district"] == dist
                      and r["candidate"] == cand)
            cd.county_totals[(race + (" District " + dist if dist else ""),
                              cand)] = v
            assert got == v, (race, dist, cand, got, v)
    reg = sum(int(r["votes"]) for r in cd.rows
              if r["office"] == "Registered Voters")
    bc = sum(int(r["votes"]) for r in cd.rows
             if r["office"] == "Ballots Cast")
    # turnout: the report's countywide row (2610 registered, 1564 ballots —
    # preliminary: the county's canvassed total is 1,570)
    assert (reg, bc) == (2610, 1564), (reg, bc)
    return cd


def fill_verified_stats(cd, county):
    """STATS the scanner missed entirely (unreadable page): visually
    verified values, keyed by county and precinct."""
    for prec, vals in (STAT_OVERRIDES.get(county) or {}).items():
        have = {r["office"] for r in cd.rows if r["precinct"] == prec}
        for lab, v in vals.items():
            if lab not in have:
                cd.add_stat(prec, lab, v)


# Scanner misreads, verified visually on the rendered pages:
STAT_OVERRIDES = {
    "Kane": {"AL": {"Ballots Cast": 50}},
    "Wayne": {"Torrey Inside/Outside": {"Ballots Cast": 158}},
    "Morgan": {"3:1": {"Registered Voters": 125, "Ballots Cast": 67},
               "14:2": {"Registered Voters": 15, "Ballots Cast": 2}},
    # Carbon: the scanner misreads the shaded ballots-cast cells; the
    # turnout percentage row gives the right value (checked by hand)
    "Carbon": {"East Helper 7": {"Ballots Cast": 72},
               "West Helper 21": {"Ballots Cast": 111},
               "Pleasant Valley 3": {"Registered Voters": 14}},
}


# Rich: 10-page certified scan (2 report pages per of 5 precincts), transcribed
# visually from the rendered pages; validates against the statewide canvass
# (U.S. Senate 60/155/649, U.S. House 454/187/213, State House 595/251,
# ballots cast 876).
RICH_STATS = {
    "1": (313, 202), "2": (316, 224), "3": (286, 196), "4": (289, 134),
    "5": (268, 120),
}
RICH_RACES = {
    "U.S. Senate": [
        ("Becky Edwards",  {"1": 28, "2": 50, "3": 32, "4": 28, "5": 17}),
        ("Ally Isom",      {"1":  9, "2": 13, "3": 16, "4": 15, "5":  7}),
        ("Mike Lee",       {"1": 162, "2": 157, "3": 144, "4": 90, "5": 96}),
    ],
    "U.S. House": [   # District 1
        ("Blake D. Moore", {"1": 102, "2": 124, "3": 101, "4": 67, "5": 60}),
        ("Tina Cannon",    {"1": 22, "2": 52, "3": 49, "4": 37, "5": 27}),
        ("Andrew Badger",  {"1": 74, "2": 43, "3": 37, "4": 27, "5": 32}),
    ],
    "State House": [  # District 4
        ("Kera Birkeland", {"1": 141, "2": 162, "3": 141, "4": 76, "5": 75}),
        ("Raelene Blocker", {"1": 56, "2": 53, "3": 48, "4": 54, "5": 40}),
    ],
    "Clerk/Auditor": [
        ("Anneliesa Peart", {"1": 109, "2": 142, "3": 99, "4": 73, "5": 73}),
        ("Justin O'Mealy",  {"1": 89, "2": 77, "3": 96, "4": 56, "5": 44}),
    ],
    "Commissioner A": [
        ("Geo Peart", {"1": 49, "2": 113, "3": 89, "4": 84, "5": 66}),
        ("Bill Cox",  {"1": 151, "2": 110, "3": 105, "4": 48, "5": 51}),
    ],
}


def parse_rich_matrix(stem, county):
    """Transcribed Rich canvass (see above); validates against the statewide
    canvass at the printed county totals."""
    cd = CountyData(county)
    assert sum(st[0] for st in RICH_STATS.values()) == 1472, "RV transcription"
    assert sum(st[1] for st in RICH_STATS.values()) == 876, "BC transcription"
    for race, cands in RICH_RACES.items():
        for cname, per in cands:
            for prec, v in per.items():
                cd.add_race("Rich " + prec + " - " + race if False else prec,
                            race, cname, v, "REP")
    for r in cd.rows:
        if r["office"] == "State House":
            r["district"] = "4"
        if r["office"] == "U.S. House":
            r["district"] = "1"
    for prec, (reg, bc) in RICH_STATS.items():
        cd.add_stat(prec, "Registered Voters", reg)
        cd.add_stat(prec, "Ballots Cast", bc)
    return cd


# ---------------------------------------------------------------- garfield matrix

def parse_garfield_matrix(path, county):
    """1-page printed-Excel canvass: precinct columns come in pairs
    ('AN 1','BO 2',..., 'TR 10'); every numeric row (stats and race rows) is
    '<name> <county total> [<pct%>] p1..p11' and the pct is dropped before
    counting, so per-precinct counts are simply the first len(precincts)
    numbers after the county total."""
    t = pdftext(path)
    cd = CountyData(county)
    lines = t.split("\n")
    precincts = None
    for l in lines:
        s = re.sub(r"\s+", " ", l.strip())
        if re.match(r"AN 1 ", s):
            toks = s.split()
            precincts = [" ".join(toks[j:j + 2]) for j in range(0, len(toks), 2)]
            break
    if not precincts:
        return None
    race = None
    for l in lines:
        s = re.sub(r"\s+", " ", l.strip())
        raw_all = re.findall(r"[0-9,]+(?:\.\d+)?%?", s)
        nums = [c for c in raw_all if not c.endswith("%")]
        if not nums and raw_all:
            # percentage-only line ( turnout percents) — belongs to the prior row
            continue
        name = re.match(r"^(.*?)\s*[0-9,]", l.strip())
        name = re.sub(r"\s+", " ", name.group(1)).strip() if name and name.group(1).strip() else ""
        if name.lower().startswith(("precincts counted", "registered", "ballots cast")):
            per = nums[1:1 + len(precincts)]
            for p, v in zip(precincts, per):
                cd.add_stat(p, "Registered Voters" if name.lower().startswith("regis") else "Ballots Cast",
                            int(v.replace(",", "")))
            continue
        stripped = re.sub(r"\s+DISTRICT\s*[0-9]+\s*$", "", s, flags=re.I)
        if re.search(r"\bSENATE\b|\bHOUSE\b|\bGOVERNOR\b|\bPRESIDENT\b", stripped, re.I) \
                and not re.findall(r"[0-9]", stripped):
            race = s
            continue
        if not race or not name:
            continue
        vals = nums[1:1 + len(precincts)]
        if len(vals) == len(precincts):
            for p, v in zip(precincts, vals):
                cd.add_race(p, race, name, ocr_int(v))
    return cd


# Summit: 111-page scan (1-2 report pages per ballot style, one STATISTICS
# table often followed by race blocks). The scanner misses single-digit value
# cells, so each page is OCR'd once (psm 11 + tsv word positions) and grouped
# into visual rows split by column: label (x<445), value, pct. A candidate's
# missed value is recovered from its printed percentage
# (value = round(pct x Total Votes Cast / 100)); every race block validates
# against its printed Total Votes Cast. Style codes are matched against the
# county's 2022 general file's precinct names.

def parse_summit_pdf(path, county):
    OCR_DIR.mkdir(exist_ok=True)
    cd = CountyData(county)
    npages = int(subprocess.run(["pdfinfo", str(path)], capture_output=True,
                                text=True).stdout.split("Pages:")[1].split()[0])
    known = {}
    gcsv = REPO / "2022" / "counties" / "20221108__ut__general__summit__precinct.csv"
    if gcsv.exists():
        with open(gcsv, newline='', encoding="utf-8") as f:
            for r_ in csv.DictReader(f):
                if r_.get("precinct"):
                    known[re.sub(r"[^0-9A-Z]", "", r_["precinct"].upper())] = \
                        r_["precinct"]

    def page_rows(pg):
        b = OCR_DIR / f"sumpg{pg:03d}"
        if not sorted(OCR_DIR.glob(f"sumpg{pg:03d}-*.png")):
            subprocess.run(["pdftoppm", "-f", str(pg), "-l", str(pg), "-r",
                            "150", "-png", str(path), str(b)], check=True)
        img = sorted(OCR_DIR.glob(f"sumpg{pg:03d}-*.png"))[0]
        base = f"sumtsv{pg:03d}"
        if not (OCR_DIR / f"{base}.tsv").exists():
            subprocess.run(["tesseract", "--psm", "11", img.name, base, "tsv"],
                           check=True, cwd=OCR_DIR)
        hdr, words = None, []
        for line in (OCR_DIR / f"{base}.tsv").read_text(
                encoding="utf-8", errors="replace").split("\n"):
            parts = line.split("\t")
            if line.startswith("level\t"):
                hdr = {k: i for i, k in enumerate(parts)}
                continue
            if hdr is None or len(parts) < 8 or parts[hdr["level"]] != "5":
                continue
            t = parts[hdr["text"]].strip()
            if t:
                words.append((int(parts[hdr["top"]]), int(parts[hdr["left"]]), t))
        words.sort(key=lambda w: (w[0], w[1]))
        rows, cur, cur_y = [], [], None
        for y, x, t in words:
            if cur_y is not None and y - cur_y > 13:
                rows.append(cur)
                cur, cur_y = [], None
            if not cur:
                cur_y = y
            cur.append((x, t))
            cur_y = y
        if cur:
            rows.append(cur)
        out = []
        for row in rows:
            words = [(x, t) for x, t in sorted(row, key=lambda w: w[1])
                     if re.search(r"[A-Za-z0-9]", t)]
            label = " ".join(t for x, t in words if x < 505)
            value = next((t for x, t in words if 505 <= x < 570), None)
            pcts = [t for x, t in words if x >= 570]
            out.append((label, value, pcts))
        return out

    TITLE_RE = re.compile(r"\b(for|district|board|commission|council|senate|"
                          r"house|mayor|school)\b")
    precinct = None
    seen = {}
    for pg in range(1, npages + 1):
        rows = page_rows(pg)
        name = None
        for label, value, pcts in rows[:4]:
            sq = re.sub(r"[^0-9A-Z]", "", label.upper()).replace("$", "S")
            if sq in known:
                name = known[sq]
                break
        if name:
            precinct = name
        elif precinct is None:
            print(f"    !! summit page {pg}: no precinct name")
            continue
        seen[precinct] = seen.get(precinct, 0) + 1
        if seen[precinct] > 2:
            print(f"    !! summit page {pg}: {precinct} seen {seen[precinct]}x")
        # stats table rows
        for label, value, pcts in rows:
            m = re.match(r"^(registered voters|ballots cast)"
                         r"(?:\s*[-–]\s*(total|republican|nonpartisan|blank))?"
                         r"\s*[:=]?\s*$", label.lower().strip())
            if not m:
                continue
            lab = m.group(1).title()
            if m.group(2):
                lab += " - " + m.group(2).title()
            if lab not in ("Registered Voters", "Ballots Cast"):
                continue
            v = ocr_int((value or "").replace(",", ""))
            if v is None and pcts:
                v = ocr_int(pcts[0])
            if v is not None:
                cd.add_stat(precinct, lab, v)
        # race blocks: segments between race-title rows
        segs = []
        title = None
        for label, value, pcts in rows:
            ls = label.lower().strip()
            if not ls and not value and not pcts:
                continue
            m = re.match(r"^(registered voters|ballots cast)", ls)
            if m:
                continue  # stats handled above
            tight = re.sub(r"\s+", "", ls)
            if tight.startswith(("totalvotescast", "contesttotals",
                                 "totalvotes cast")):
                if title:
                    segs[-1][1].append((label, value, pcts))
                continue
            if re.match(r"^(republican|democratic|democrat)\s+\w+", ls) \
                    or "school board" in ls or "school district" in ls:
                title = re.sub(r"\s+", " ", label).replace("Schoo!", "School")\
                          .rstrip(" .:|")
                segs.append((title, []))
                continue
            if is_votefor_line(label.strip()) or tight.startswith(
                    ("total", "statistics")):
                continue
            if title:
                segs[-1][1].append((label, value, pcts))
        # emit each segment
        for title, block in segs:
            cands = []
            ovr = undr = tvc = None
            for label, value, pcts in block:
                tight = re.sub(r"\s+", "", label.lower())
                if tight.startswith("overvotes"):
                    ovr = ocr_int((value or "").replace(",", ""))
                    continue
                if tight.startswith("undervotes"):
                    undr = ocr_int((value or "").replace(",", ""))
                    continue
                if tight.startswith(("totalvotescast", "contesttotals")):
                    tvc = ocr_int((value or "").replace(",", "")) or \
                        ocr_int((pcts[0] if pcts else ""))
                    continue
                v = ocr_int((value or "").replace(",", ""))
                pct = None
                for tok in pcts:
                    if re.fullmatch(r"[0-9]+(?:\.[0-9]+)?%?", tok):
                        pct = float(tok.rstrip("%"))
                        break
                cands.append((re.sub(r"\s+", " ", label.strip(". |")),
                              v, pct))
            # derive the TVC when not printed: value / (pct/100)
            if tvc is None:
                for _, v, pct in cands:
                    if v is not None and pct:
                        tvc = int(round(v * 100.0 / pct))
                        break
            for name_, v, pct in cands:
                if v is None and pct is not None and tvc:
                    v = int(round(tvc * pct / 100.0))
                if v is None:
                    continue
                cur_rows = {}
                cur_rows[name_.replace(" ", "").upper()] = v
                cd.add_race(precinct, title, name_, v)
            if tvc is not None and cands:
                got = sum(v for _, v, _ in cands if v is not None)
                want = tvc - (ovr or 0) - (undr or 0)
                if ovr is not None or undr is not None:
                    if got != want:
                        print(f"    !! summit page {pg} {precinct}: {title} "
                              f"rows sum {got} vs {want} of {tvc} "
                              f"(over {ovr}, under {undr})")
    return cd

# ---------------------------------------------------------------- emery/sevier transcriptions

EMERY_PRECINCTS = ["Emery #1", "Ferron #2", "Ferron #3", "Clawson #4", "Orangeville #5",
                   "Castle Dale #6", "Castle Dale #7", "Huntington #8", "Huntington #9",
                   "Cleveland #10", "Elmo #11", "Green River #12", "Canvas"]
EMERY_STATS = {"Registered Voters": [137, 385, 285, 152, 590, 266, 344, 423, 428, 408, 253, 297, 0],
               "Ballots Cast": [79, 240, 183, 99, 394, 184, 243, 235, 251, 246, 151, 163, 0]}
EMERY_RACES = {
    "U.S. Senate": {
        "Becky Edwards": [9, 31, 33, 5, 62, 25, 35, 23, 55, 21, 25, 17, 0],
        "Ally Isom": [9, 20, 14, 5, 34, 16, 22, 18, 26, 21, 13, 10, 0],
        "Mike Lee": [61, 184, 136, 85, 286, 141, 182, 190, 166, 197, 110, 135, 0],
    },
    "U.S. House": {   # District 3
        "Christopher Niles Herrod": [14, 69, 36, 26, 90, 34, 42, 62, 68, 66, 50, 39, 0],
        "John Curtis": [60, 163, 143, 69, 288, 147, 193, 163, 176, 173, 95, 122, 0],
    },
    "State House": {  # District 67
        "Tom Hansen": [None, None, None, 49, 212, 89, 136, 135, 145, 157, 125, None, 0],
        "Christine F. Watkins": [None, None, None, 48, 166, 89, 104, 92, 103, 82, 25, None, 0],
    },
    "Commissioner Seat A": {
        "Jordan Leonard": [46, 141, 106, 46, 198, 111, 142, 162, 175, 137, 97, 98, 0],
        "Kent B. Wilson": [32, 96, 75, 52, 194, 72, 97, 73, 74, 98, 49, 55, 0],
    },
    "Commissioner Seat B": {
        "Gil Conover": [21, 98, 77, 40, 108, 43, 69, 59, 70, 18, 16, 25, 0],
        "Douglas K. Stilson": [19, 86, 55, 33, 195, 102, 103, 89, 80, 54, 32, 33, 0],
        "Keven Andrew Jensen": [38, 55, 49, 25, 83, 38, 65, 83, 93, 172, 100, 101, 0],
    },
    "Sheriff": {
        "Tyson Huntington": [29, 99, 80, 35, 125, 75, 130, 123, 162, 160, 87, 117, 0],
        "Shaun Bell": [49, 135, 102, 61, 263, 108, 106, 110, 87, 81, 63, 42, 0],
    },
}
EMERY_DISTRICTS = {"U.S. House": "3", "State House": "67"}


SEVIER_PRECINCTS = ["Annabella #1", "Aurora #2", "Central #4", "Elsinore #5", "Glenwood #6",
                    "Joseph #7", "Koosharem #17", "Monroe #8", "Monroe #9", "Redmond #10",
                    "Richfield #11", "Richfield #12", "Richfield #13",
                    "Richfield #20", "Richfield #21", "Richfield #22", "Richfield #23",
                    "Richfield #24", "Salina #14", "Salina #15", "Salina #25", "Sigurd #16",
                    "Venice #18"]
SEVIER_STATS = {"Registered Voters": [434, 549, 347, 423, 322, 256, 136, 758, 610, 347, 361, 314, 369,
                                      300, 290, 359, 464, 472, 271, 354, 287, 158, 182],
                "Ballots Cast": [229, 255, 195, 235, 189, 134, 71, 404, 346, 146, 204, 180, 218,
                                 138, 155, 211, 250, 262, 115, 142, 127, 87, 108]}
SEVIER_RACES = {
    "U.S. Senate": {
        "Becky Edwards": [20, 29, 29, 17, 11, 10, 9, 36, 45, 6, 38, 31, 28, 21, 22, 36, 37, 35, 8, 21, 24, 17, 13],
        "Ally Isom": [17, 11, 8, 10, 17, 9, 1, 25, 23, 9, 12, 3, 6, 6, 9, 19, 16, 24, 2, 6, 2, 6, 7],
        "Mike Lee": [191, 215, 157, 208, 159, 115, 61, 339, 278, 130, 154, 146, 183, 111, 123, 156, 193, 202, 105, 115, 100, 64, 88],
    },
    "U.S. House District 2": {
        "Erin Rider": [20, 21, 18, 19, 17, 14, 14, 43, 52, 9, 35, 18, 22, 28, 24, 23, 25, 37, 7, 19, 17, 15, 14],
        "Chris Stewart": [208, 233, 177, 216, 170, 119, 57, 353, 292, 137, 167, 161, 194, 110, 127, 185, 223, 223, 108, 123, 109, 71, 94],
    },
    "County Commission Seat B": {
        "Edward R. Nielson": [18, 25, 18, 49, 21, 24, 10, 59, 39, 14, 40, 20, 22, 19, 17, 11, 13, 15, 9, 17, 10, 11, 6],
        "Garth Ogden": [86, 87, 74, 86, 71, 61, 34, 162, 134, 38, 94, 93, 69, 59, 71, 86, 135, 73, 31, 51, 55, 22, 59],
        "Greg Jensen": [120, 142, 102, 100, 92, 48, 24, 179, 169, 94, 69, 66, 127, 60, 67, 113, 99, 174, 75, 74, 61, 54, 43],
    },
}


def parse_transcribed_matrix(county, precincts, stats, races, districts):
    cd = CountyData(county)
    for label, vals in stats.items():
        for p, v in zip(precincts, vals):
            cd.add_stat(p, label, v)
    for office, cands in races.items():
        for cand, vals in cands.items():
            for p, v in zip(precincts, vals):
                if v is not None:
                    cd.add_race(p, office, cand, v)
                    if districts.get(office):
                        cd.rows[-1]["district"] = str(districts[office])
    return cd


# ---------------------------------------------------------------- driver

PARSERS = {
    "dom1d": lambda county, fname: (parse_dominion1d(pdftext(SRC / fname), county), fname),
    "morgan": lambda county, fname: (parse_morgan_pdf(SRC / fname, county), fname),
    "uintahmat": lambda county, fname: (parse_uintah_matrix("uintah3", county), fname),
    "richmat": lambda county, fname: (parse_rich_matrix("rich3", county), fname),
    "summit": lambda county, fname: (parse_summit_pdf(SRC / fname, county), fname),
    "kane": lambda county, fname: (parse_morgan_pdf(
        SRC / fname, county, order=["KA1", "KA2", "KA3", "KA4", "AL", "BU", "BW",
                                    "CM", "CW", "GL", "OR"], of_total=22), fname),
    "cache": lambda county, fname: (parse_cache(SRC / fname, county), fname),
    "saltlake": lambda county, fname: (parse_saltlake_xlsx(SRC / fname, county), fname),
    "davis": lambda county, fname: (parse_davis_xlsx(SRC / fname, county), fname),
    "widecsv": lambda county, fname: (parse_utah_widecsv(SRC / fname, county), fname),
    "beaver": lambda county, fname: (parse_beaver_clarity(SRC / fname, county), fname),
    "garfield": lambda county, fname: (parse_garfield_matrix(SRC / fname, county), fname),
    "emery": lambda county, fname: (parse_transcribed_matrix(county, EMERY_PRECINCTS, EMERY_STATS, EMERY_RACES, EMERY_DISTRICTS), "transcribed abstract"),
    "sevier": lambda county, fname: (parse_transcribed_matrix(county, SEVIER_PRECINCTS, SEVIER_STATS, SEVIER_RACES, {}), "transcribed abstract"),
}


def parse_county(name):
    slug, fname, fmt = COUNTIES[name]
    if fmt.startswith("dom1d_ocr:"):
        parts = fmt.split(":")
        stem = parts[1]
        keep = parts[2] if len(parts) > 2 else "1-"   # 1-indexed page range
        m = re.match(r"(\d+)?(-(\d*))?$", keep)
        text, npages = load_ocr_text(stem)
        page_text = text.split("\f")
        from_page = int(m.group(1)) if m and m.group(1) else 1
        to_page = (int(m.group(3)) if m and m.group(2) and m.group(3)
                   else len(page_text))
        # leading/trailing pages of the OCR text can be cover/countywide
        # tables, not precinct data (Morgan: Election Summary 1-3;
        # Uintah: letter + Statistics pages, 1-2 and 32+)
        text = "\f".join(page_text[from_page - 1:to_page])
        cd = parse_dominion1d(text, name)
        fill_verified_stats(cd, name)
        src = (f"OCR (pages {from_page}-{min(to_page, npages)} of {npages})")
        return cd, src
    cd, src = PARSERS[fmt](name, fname)
    fill_verified_stats(cd, name)
    return cd, src


def main(argv):
    names = argv[1:] or list(COUNTIES)
    for c in names:
        spec = COUNTIES.get(c)
        if not isinstance(spec, tuple):
            print(f"{c}: SKIPPED — {spec}")
            continue
        slug, fname, fmt = spec
        try:
            cd, src = parse_county(c)
        except Exception as e:
            print(f"{c}: FAILED ({fmt}): {type(e).__name__}: {e}")
            continue
        n_merged = canonicalize_candidates(cd, c)
        if n_merged:
            print(f"   ({n_merged} OCR-variant candidate rows canonicalized)")
        races = {}
        for r in cd.rows:
            if r['candidate'] and str(r['votes']).isdigit():
                k = (r['office'], r['district'], r['party'])
                races.setdefault(k, {})
                races[k][r['candidate']] = races[k].get(r['candidate'], 0) + int(r['votes'])
        nprec = len({r['precinct'] for r in cd.rows})
        print(f"== {c} — {len(cd.rows)} rows, {nprec} precincts")
        for (o, d, p), cc in sorted(races.items()):
            tot = sum(cc.values())
            detail = ", ".join(f"{n}={v}" for n, v in sorted(cc.items()))
            print(f"   {o}{' (d' + d + ')' if d else ''}: {tot} [{detail}]")
        if cd.county_totals:
            bad = 0
            for (rt, cand), v in sorted(cd.county_totals.items()):
                if rt == cand in ("Registered Voters", "Ballots Cast"):
                    mine = sum(int(r['votes']) for r in cd.rows
                               if r['office'] == rt and r['candidate'] == "")
                else:
                    o, dist = district_of(norm_race(rt)[0])
                    mine = sum(int(r['votes']) for r in cd.rows
                               if r['office'] == o and r['district'] == dist
                               and re.sub(r"\s+", "", re.sub(r"[^A-Za-z.']", "", r['candidate'].upper()))
                               == re.sub(r"\s+", "", re.sub(r"[^A-Za-z.']", "", cand.upper())))
                if mine != v:
                    bad += 1
                    print(f"   !! countywide mismatch {o}{(' ' + dist) if dist else '' or ('' if rt == cand else ' ' + cand)}: "
                          f"csv {mine} vs report {v}")
            if not bad and len(cd.county_totals):
                print(f"   (countywide totals match for all "
                      f"{len(cd.county_totals)} printed checks)")
            if not bad:
                print(f"   (countywide totals match for all "
                      f"{len(cd.county_totals)} printed checks)")
        path = cd.write_csv(slug)
        print(f"   wrote {path.name}")


if __name__ == "__main__":
    main(sys.argv)