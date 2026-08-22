"""Shared framework for parsing Utah 2026 primary PDFs into OpenElections CSVs.

Handles text extraction (pdftotext for embedded-text PDFs, PaddleOCR 3.x for
scanned PDFs), normalization of parties/offices/districts/candidate names, row
construction, CSV writing, and per-county reconciliation of precinct sums
against summary totals.

Run via the repo's `uv` environment:
    uv run python3 parsers/p2026_<county>.py
"""

import csv
import json
import os
import re
import subprocess
from pathlib import Path

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #

REPO_ROOT = Path(__file__).resolve().parent.parent
SOURCES_DIR = Path.home() / "code" / "openelections-sources-ut" / "2026 Primary Election"
OUT_DIR = REPO_ROOT / "2026" / "counties"
CACHE_DIR = REPO_ROOT / ".paddleocr_cache"
ELECTION_DATE = "20260623"

# Minimum embedded-text chars to consider a PDF digitally-generated (no OCR).
TEXT_CHAR_THRESHOLD = 50

# Map of normalized county name -> display name used in the `county` column.
# Most are the same; listed for those with multi-word names.
COUNTY_DISPLAY = {
    "beaver": "Beaver", "box_elder": "Box Elder", "cache": "Cache",
    "carbon": "Carbon", "daggett": "Daggett", "davis": "Davis",
    "duchesne": "Duchesne", "emery": "Emery", "garfield": "Garfield",
    "grand": "Grand", "iron": "Iron", "juab": "Juab", "kane": "Kane",
    "millard": "Millard", "morgan": "Morgan", "piute": "Piute",
    "rich": "Rich", "salt_lake": "Salt Lake", "san_juan": "San Juan",
    "sanpete": "Sanpete", "sevier": "Sevier", "summit": "Summit",
    "tooele": "Tooele", "uintah": "Uintah", "utah": "Utah",
    "wasatch": "Wasatch", "washington": "Washington", "wayne": "Wayne",
    "weber": "Weber",
}

# --------------------------------------------------------------------------- #
# Normalization
# --------------------------------------------------------------------------- #

PARTY_MAP = {
    "REPUBLICAN": "REP", "REP": "REP",
    "DEMOCRATIC": "DEM", "DEMOCRAT": "DEM", "DEM": "DEM",
    "LIBERTARIAN": "LIB", "LIB": "LIB",
    "INDEPENDENT AMERICAN": "IAP", "IAP": "IAP",
    "CONSTITUTION": "CON", "CON": "CON",
    "UNITED UTAH": "UUP", "UUP": "UUP",
    "INDEPENDENT": "IND", "IND": "IND",
    "GREEN": "GRN", "GRN": "GRN",
    "NONPARTISAN": "", "UNAFFILIATED": "", "": "",
}

ALLOWED_PARTIES = {"REP", "DEM", "LIB", "IAP", "CON", "UUP", "IND", "GRN", "NP", ""}


def normalize_party(raw):
    if raw is None:
        return ""
    key = re.sub(r"\s+", " ", str(raw).strip().upper())
    key = key.replace("(", "").replace(")", "")
    if key in PARTY_MAP:
        return PARTY_MAP[key]
    # Fallback: try without extra tokens
    for token in key.split():
        if token in PARTY_MAP:
            return PARTY_MAP[token]
    return ""


# Office-name standardization.  Keys are matched as substrings (uppercased).
OFFICE_MAP = [
    (r"U\.?S\.?\s*PRESIDENT", "President"),
    (r"U\.?S\.?\s*SENATE", "U.S. Senate"),
    (r"U\.?S\.?\s*HOUSE", "U.S. House"),
    (r"CONGRESSIONAL\s+DISTRICT", "U.S. House"),
    (r"U\.?S\.?\s*SENATOR", "U.S. Senate"),
    (r"GOVERNOR", "Governor"),
    (r"ATTORNEY\s+GENERAL", "Attorney General"),
    (r"STATE\s+AUDITOR", "State Auditor"),
    (r"STATE\s*TREASURER", "State Treasurer"),
    (r"STATE\s+SENATE", "State Senate"),
    (r"STATE\s+HOUSE", "State House"),
    (r"STATE\s+HOUSE\s+OF\s+REPRESENTATIVES", "State House"),
    (r"STATE\s+SCHOOL\s+BOARD", "State School Board"),
    (r"STATE\s+BOARD\s+OF\s+EDUCATION", "State School Board"),
]

_DISTRICT_RE = re.compile(r"(?:DISTRICT|DIST)\s*#?\s*(\d+)", re.IGNORECASE)
_HOUSE_DIST_RE = re.compile(r"(?:HOUSE|SENATE|SCHOOL\s+BOARD)\s+DISTRICT\s*(\d+)", re.IGNORECASE)
_TRAILING_NUM_RE = re.compile(r"\bDISTRICT\s+(\d+)\b", re.IGNORECASE)


def parse_district(contest):
    """Extract a district string (e.g. '3') from a contest name, or ''."""
    if not contest:
        return ""
    m = _TRAILING_NUM_RE.search(contest) or _HOUSE_DIST_RE.search(contest) or _DISTRICT_RE.search(contest)
    if m:
        return m.group(1)
    # "U.S. House 3" style
    m = re.search(r"HOUSE\s+(\d+)", contest, re.IGNORECASE)
    if m:
        return m.group(1)
    return ""


def normalize_office(contest, contest_type="Candidate"):
    """Standardize an office name from a contest label."""
    if not contest:
        return ""
    name = re.sub(r"\s+", " ", contest).strip()
    up = name.upper()
    if contest_type == "BallotMeasure":
        return name  # ballot measures / retention keep their text
    for pat, std in OFFICE_MAP:
        if re.search(pat, up):
            return std
    # Strip party prefix like "Republican for " / "Democratic for "
    name = re.sub(r"^(REPUBLICAN|DEMOCRATIC|LIBERTARIAN|INDEPENDENT\s+AMERICAN|CONSTITUTION|UNITED\s+UTAH|NONPARTISAN)\s+FOR\s+", "", name, flags=re.IGNORECASE)
    # Strip trailing district/vote-for noise
    name = re.sub(r"\s*\(VOTE\s+FOR\s+\d+\)\s*", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s*-\s*(REP|DEM|LIB|IAP|CON|UUP|IND|GRN)\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+DISTRICT\s+\d+\s*$", "", name, flags=re.IGNORECASE)
    name = name.strip()
    # County/local offices not in the map: title-case the printed name
    if name and name.isupper():
        name = name.title()
    return name


_MC_MAC = re.compile(r"\b(Ma?c)([a-z])")
_O_PREFIX = re.compile(r"\bO'([a-z])")


def title_case_name(raw):
    """Title-case a candidate name, handling Mc/Mac and O' prefixes.

    Short all-caps tokens are preserved as-is -- these are initials or
    nicknames like "CJ", "AJ", "II", "D" that .title() would wrongly lower to
    "Cj"/"Aj"/"Ii".  The common "JR"/"SR" suffixes are deliberately excluded so
    they title-case to "Jr"/"Sr" as expected for Junior/Senior.
    """
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""
    # Already mixed-case (e.g. from JSON-style sources) -> leave as-is
    if s != s.upper():
        return s
    # Tokens to preserve verbatim after .title(): short all-caps alpha tokens
    # (1-2 letters) that are NOT the Jr/Sr suffixes.
    preserve = {tok for tok in s.split()
                if tok.isalpha() and tok.isupper() and 1 <= len(tok) <= 2
                and tok not in ("JR", "SR")}
    s = s.title()
    # Mc/Mac corrections: Mcadams -> McAdams
    s = _MC_MAC.sub(lambda m: m.group(1) + m.group(2).upper(), s)
    s = _O_PREFIX.sub(lambda m: "O'" + m.group(1).upper(), s)
    if preserve:
        def _keep_caps(m):
            w = m.group(0)
            return w.upper() if w.upper() in preserve else w
        s = re.sub(r"\b[A-Za-z]{1,2}\b", _keep_caps, s)
    return s


def to_int(value):
    """Coerce a vote value to int; blanks / percentages / None -> '' ."""
    if value is None:
        return ""
    s = str(value).strip().replace(",", "")
    if s == "" or s == "-" or "%" in s:
        return ""
    try:
        return int(s)
    except ValueError:
        return ""


def num(value):
    """Coerce to int for arithmetic (blanks -> 0)."""
    v = to_int(value)
    return v if v != "" else 0


# --------------------------------------------------------------------------- #
# Text extraction
# --------------------------------------------------------------------------- #

def _pdf_page_count(pdf_path):
    out = subprocess.run(["pdfinfo", str(pdf_path)], capture_output=True, text=True)
    for line in out.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split(":")[1].strip())
    return 0


def pdf_text_chars(pdf_path):
    out = subprocess.run(["pdftotext", str(pdf_path), "-"], capture_output=True, text=True)
    return len(re.sub(r"\s", "", out.stdout))


def needs_ocr(pdf_path):
    return pdf_text_chars(pdf_path) < TEXT_CHAR_THRESHOLD


def pdftotext_layout(pdf_path):
    """Return list of pages; each page is a list of (lineno, x, text)."""
    out = subprocess.run(["pdftotext", "-layout", str(pdf_path), "-"],
                         capture_output=True, text=True)
    pages = []
    for page_text in out.stdout.split("\f"):
        lines = []
        for i, line in enumerate(page_text.splitlines()):
            if not line.strip():
                continue
            x = len(line) - len(line.lstrip(" "))
            lines.append({"y": i, "x": x, "text": line.rstrip(), "score": 1.0})
        pages.append(lines)
    return pages


def _render_pages(pdf_path, dpi=300, pad=80):
    """Render each PDF page to a PNG (padded with a white margin to avoid
    left-edge OCR cropping).  Returns list of image Paths."""
    import tempfile
    tmp = Path(tempfile.mkdtemp(prefix="ut2026_"))
    prefix = tmp / "page"
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
                   check=True, capture_output=True)
    try:
        from PIL import Image, ImageOps
    except ImportError:
        return sorted(prefix.parent.glob("page-*.png"))
    out = []
    for png in sorted(prefix.parent.glob("page-*.png")):
        im = Image.open(png).convert("RGB")
        im = ImageOps.expand(im, border=pad, fill="white")
        padded = png.with_suffix(".pad.png")
        im.save(padded)
        out.append(padded)
    return out


def ocr_pdf(pdf_path, county, dpi=150):
    """OCR a scanned PDF with PaddleOCR 3.x.  Returns list of pages; each page
    is a list of line dicts {y, x, text, score} sorted top-to-bottom.

    The cache is keyed by county AND source filename, so a county's summary
    and precinct PDFs (both scanned) don't collide and overwrite each other's
    cached pages.
    """
    slug = Path(pdf_path).stem.replace(" ", "_").replace(",", "")
    cache = CACHE_DIR / county / slug
    cache.mkdir(parents=True, exist_ok=True)
    npages = _pdf_page_count(pdf_path)

    pages = []
    # Reuse a single OCR instance across pages
    ocr = None
    images = None
    for p in range(npages):
        cache_file = cache / f"page_{p:03d}.json"
        if cache_file.exists():
            pages.append(json.loads(cache_file.read_text()))
            continue
        if ocr is None:
            from paddleocr import PaddleOCR
            # PP-OCRv6 "small" models: ~3x faster than the default "medium" on CPU
            # with no loss of accuracy on these clean printed scans. The fast-
            # config flags (no doc orientation/unwarping/textline orientation)
            # cut init time; text_det_unclip_ratio=2.0 expands detection boxes so
            # the leftmost glyph of left-margin lines is not cropped ("DANIEL"
            # not "ANIEL"). text_det_limit_side_len=720 + batch 16 speed detection
            # and recognition; 150 DPI is sufficient at this print quality.
            ocr = PaddleOCR(lang="en", use_textline_orientation=False,
                            use_doc_orientation_classify=False,
                            use_doc_unwarping=False,
                            text_det_unclip_ratio=2.0,
                            text_recognition_batch_size=16,
                            text_det_limit_side_len=720,
                            text_detection_model_name="PP-OCRv6_small_det",
                            text_recognition_model_name="PP-OCRv6_small_rec")
            images = _render_pages(pdf_path, dpi=dpi)
        img = images[p]
        res = ocr.predict(str(img))
        d = res[0].json["res"]
        lines = []
        for txt, poly, score in zip(d["rec_texts"], d["dt_polys"], d["rec_scores"]):
            if not txt or not txt.strip():
                continue
            xs = [pt[0] for pt in poly]
            ys = [pt[1] for pt in poly]
            lines.append({"y": round(min(ys)), "x": round(min(xs)),
                           "text": txt.strip(), "score": round(float(score), 3)})
        lines.sort(key=lambda l: (l["y"], l["x"]))
        cache_file.write_text(json.dumps(lines))
        pages.append(lines)
    return pages


# Counties whose source PDFs have an embedded text layer that is garbled
# (a broken ToUnicode map: letters render correctly on screen but extract as
# wrong characters, e.g. "County" -> "couniy", "JOE" -> "JEO", "7" -> "1").  The
# visual rendering is fine, so OCR on the rasterized pages gives clean text --
# force OCR for these even though pdftotext finds plenty of (garbled) chars.
FORCE_OCR_COUNTIES = {"sevier"}


def extract(pdf_path, county):
    """Extract per-page lines from a PDF, choosing pdftotext or PaddleOCR.
    Returns (pages, is_ocr) where pages is a list of lists of line dicts."""
    if county in FORCE_OCR_COUNTIES or needs_ocr(pdf_path):
        return ocr_pdf(pdf_path, county), True
    return pdftotext_layout(pdf_path), False


def visual_rows(page, is_ocr, y_tol=18):
    """Merge a page's line dicts into visual rows (one line of text per visual
    row).  For pdftotext each line is already a row.  For OCR, lines at nearly
    the same y are merged (text sorted left-to-right) so a candidate name and
    its vote count, which OCR emits as two boxes, become one row."""
    if not is_ocr:
        return [(l["y"], l["x"], l["text"]) for l in page]
    rows = []
    for line in page:
        y = line["y"]
        placed = False
        for r in rows:
            if abs(r[0] - y) <= y_tol:
                r[2].append((line["x"], line["text"]))
                placed = True
                break
        if not placed:
            rows.append([y, line["x"], [(line["x"], line["text"])]])
    out = []
    for y, _, parts in rows:
        parts.sort()
        text = "  ".join(t for _, t in parts).strip()
        out.append((y, min(p[0] for p in parts), text))
    return out


def extract_rows(pdf_path, county):
    """Return list of pages; each page is a list of (y, x, text) visual rows."""
    pages, is_ocr = extract(pdf_path, county)
    return [visual_rows(p, is_ocr) for p in pages]


# --------------------------------------------------------------------------- #
# Row construction & CSV writing
# --------------------------------------------------------------------------- #

PRECINCT_HEADER = ["county", "precinct", "office", "district", "party", "candidate", "votes"]
COUNTY_HEADER = ["county", "office", "district", "party", "candidate", "votes"]

# Labels used as `office` for statistics rows.
REGISTERED_VOTERS = "Registered Voters"
BALLOTS_CAST = "Ballots Cast"
BALLOTS_CAST_BLANK = "Ballots Cast Blank"


def candidate_row(county, office, district, party, candidate, votes, precinct=None):
    row = {"county": county, "office": office, "district": district,
           "party": party, "candidate": candidate, "votes": to_int(votes)}
    if precinct is not None:
        row = {"county": county, "precinct": precinct, **row}
    return row


def meta_row(county, label, party, votes, precinct=None):
    row = {"county": county, "office": label, "district": "", "party": party,
           "candidate": "", "votes": to_int(votes)}
    if precinct is not None:
        row = {"county": county, "precinct": precinct, **row}
    return row


def overunder_row(county, office, district, party, kind, votes, precinct=None):
    """kind = 'Over Votes' or 'Under Votes'."""
    return candidate_row(county, office, district, party, kind, votes, precinct)


def _sort_key(row, precinct):
    if precinct:
        return (row.get("county", ""), row.get("precinct", ""), row.get("office", ""),
                str(row.get("district", "")), row.get("party", ""), row.get("candidate", ""))
    return (row.get("county", ""), row.get("office", ""),
            str(row.get("district", "")), row.get("party", ""), row.get("candidate", ""))


def write_csv(rows, path, precinct=False):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    header = PRECINCT_HEADER if precinct else COUNTY_HEADER
    rows = sorted(rows, key=lambda r: _sort_key(r, precinct))
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=header)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in header})
    return path


def county_paths(county):
    """Return (county_csv, precinct_csv) output paths for a normalized county.
    Naming: 20260623__ut__primary__<county>__county.csv / __precinct.csv
    (double-underscore separators throughout)."""
    base = f"{ELECTION_DATE}__ut__primary__{county}"
    return OUT_DIR / f"{base}__county.csv", OUT_DIR / f"{base}__precinct.csv"


def find_source(county, kind):
    """Find a source file (pdf or xlsx) in SOURCES_DIR for the county and kind
    ('summary' or 'precinct').  Classification uses filename keywords first
    (precinct/sovc/sov => precinct; summary/canvass/results => summary), then
    falls back to size (smaller file = summary, larger = precinct) for counties
    whose filenames don't carry those keywords."""
    county_prefix = COUNTY_DISPLAY.get(county, county.replace("_", " "))
    files = [f for f in SOURCES_DIR.iterdir()
             if f.name.lower().startswith(county_prefix.lower() + " ut")
             and f.suffix.lower() in (".pdf", ".xlsx")]
    if not files:
        return None

    def classify(f):
        low = f.name.lower()
        is_precinct = any(k in low for k in ("precinct", "precint", "sovc", "sov ", "state canvass precinct"))
        is_summary = any(k in low for k in ("summary", "canvass", "results."))
        # "State Canvass" alone (no precinct) is a summary; precinct canvass is precinct
        if "state canvass" in low and "precinct" not in low:
            is_summary = True
        return is_precinct, is_summary

    precinct_hits = [f for f in files if classify(f)[0]]
    summary_hits = [f for f in files if classify(f)[1] and not classify(f)[0]]
    if kind == "precinct":
        if precinct_hits:
            return max(precinct_hits, key=lambda f: f.stat().st_size)
        # No keyword match: largest file is the precinct report
        return max(files, key=lambda f: f.stat().st_size)
    if kind == "summary":
        if summary_hits:
            return min(summary_hits, key=lambda f: f.stat().st_size)
        # No keyword match: smallest file is the summary
        return min(files, key=lambda f: f.stat().st_size)
    return None


# --------------------------------------------------------------------------- #
# Reconciliation
# --------------------------------------------------------------------------- #

def _read_csv(path):
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def reconcile(county, county_csv, precinct_csv, verbose=True):
    """Compare precinct vote sums to summary totals.  Returns list of
    mismatch dicts."""
    crows = _read_csv(county_csv)
    prows = _read_csv(precinct_csv)
    display = COUNTY_DISPLAY.get(county, county)

    # Aggregate precinct by (office, district, party, candidate)
    agg = {}
    for r in prows:
        key = (r["office"], r["district"], r["party"], r["candidate"])
        agg[key] = agg.get(key, 0) + num(r["votes"])

    summary = {}
    for r in crows:
        key = (r["office"], r["district"], r["party"], r["candidate"])
        summary[key] = num(r["votes"])

    _META_OFFICES = {REGISTERED_VOTERS, BALLOTS_CAST, BALLOTS_CAST_BLANK}
    _SUPPLEMENTAL_CANDIDATES = {"Over Votes", "Under Votes"}
    mismatches = []
    all_keys = set(summary) | set(agg)
    for key in sorted(all_keys):
        s = summary.get(key, 0)
        p = agg.get(key, 0)
        in_summary = key in summary
        # Statistics meta rows: only compare when the summary actually has
        # that (office, party) breakdown.  Precinct STATISTICS pages often
        # carry Registered/Ballots detail the summary PDF omits.
        if key[0] in _META_OFFICES and not in_summary:
            continue
        # Over/Under Votes are optional; summaries often omit them while the
        # precinct SOVC includes them.  Don't flag precinct-only Over/Under.
        if key[3] in _SUPPLEMENTAL_CANDIDATES and not in_summary:
            continue
        if s != p:
            mismatches.append({"office": key[0], "district": key[1],
                               "party": key[2], "candidate": key[3],
                               "summary": s, "precinct_sum": p,
                               "diff": p - s})

    if verbose:
        if mismatches:
            print(f"\n[{display}] RECONCILE MISMATCHES ({len(mismatches)}):")
            for m in mismatches[:40]:
                print(f"  {m['office']} | d={m['district']} | {m['party']} | "
                      f"{m['candidate']}: summary={m['summary']} precinct={m['precinct_sum']} (diff {m['diff']:+d})")
            if len(mismatches) > 40:
                print(f"  ...and {len(mismatches) - 40} more")
        else:
            print(f"[{display}] reconcile OK: precinct sums match summary totals "
                  f"({len(summary)} summary rows, {len(agg)} precinct groups)")
    return mismatches


# --------------------------------------------------------------------------- #
# Candidate-name reconciliation (fix OCR space-loss against summary names)
# --------------------------------------------------------------------------- #

def _letters(s):
    """Lowercase letters only, spaces/punctuation stripped — for fuzzy
    space-loss comparison (e.g. 'Danielgardner' vs 'DanielGardner')."""
    return re.sub(r"[^a-z]", "", (s or "").lower())


def reconcile_candidate_names(rows, county_csv):
    """Fix precinct candidate names that don't match any summary candidate
    for their contest but are a space-loss variant of exactly one.

    OCR sometimes drops the space inside a name ('Daniel Gardner' ->
    'Danielgardner').  The summary file has the correct spelling, so for each
    precinct row whose candidate isn't a summary candidate, compare
    letters-only; if exactly one summary candidate in the same contest shares
    those letters, adopt the summary spelling.  Returns a new row list.
    """
    crows = _read_csv(county_csv)
    _META = {REGISTERED_VOTERS, BALLOTS_CAST, BALLOTS_CAST_BLANK}
    # contest -> {letters: [summary candidate names]}
    contest_cands = {}
    name_set = {}  # contest -> set(summary candidate names)
    for r in crows:
        if r["office"] in _META:
            continue
        c = r["candidate"]
        if c in ("Over Votes", "Under Votes", ""):
            continue
        key = (r["office"], str(r["district"]), r["party"])
        contest_cands.setdefault(key, {}).setdefault(_letters(c), []).append(c)
        name_set.setdefault(key, set()).add(c)

    fixed_rows = []
    fixed = 0
    for r in rows:
        key = (r.get("office", ""), str(r.get("district", "")), r.get("party", ""))
        cand = r.get("candidate", "")
        cands = name_set.get(key)
        # Only attempt a fix when the name is NOT already a summary candidate
        # and the contest actually has summary candidates to match against.
        if cands and cand not in cands:
            lc = _letters(cand)
            if lc:
                matches = contest_cands.get(key, {}).get(lc)
                if matches and len(set(matches)) == 1:
                    r = dict(r)
                    r["candidate"] = matches[0]
                    fixed += 1
        fixed_rows.append(r)
    return fixed_rows, fixed


# --------------------------------------------------------------------------- #
# Sanity checks
# --------------------------------------------------------------------------- #

def sanity_check(path, precinct=False):
    """Return list of issues found in a written CSV."""
    rows = _read_csv(path)
    issues = []
    seen = set()
    for i, r in enumerate(rows, start=2):
        if not r.get("county"):
            issues.append(f"line {i}: empty county")
        if precinct and not r.get("precinct"):
            issues.append(f"line {i}: empty precinct")
        if not r.get("office"):
            issues.append(f"line {i}: empty office")
        # Over/Under/meta rows may have blank candidate; candidate rows should not
        if r.get("party") and r["party"] not in ALLOWED_PARTIES:
            issues.append(f"line {i}: bad party '{r['party']}'")
        v = r.get("votes", "")
        if v != "" and not re.fullmatch(r"\d+", str(v)):
            issues.append(f"line {i}: non-integer votes '{v}'")
        if precinct:
            key = (r["county"], r["precinct"], r["office"], r["district"],
                   r["party"], r["candidate"])
        else:
            key = (r["county"], r["office"], r["district"], r["party"], r["candidate"])
        if key in seen:
            issues.append(f"line {i}: duplicate row {key}")
        seen.add(key)
    return issues