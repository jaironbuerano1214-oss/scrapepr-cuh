"""
Alumni Employment Tracer - Automated Data Lookup
=================================================
Reads names from 'list of names.xlsx', matches them against existing
responses in the alumni form responses xlsx, then searches LinkedIn/web
for any alumni not yet found. Outputs 'alumni_tracer_output.xlsx' with
all columns from the tracer form.

FILTER: Only alumni who studied at one of the accepted universities are
        processed.  All others are silently skipped.

Accepted universities (case-insensitive, substring match):
  - Pampanga State University  (PSU / PamSU)
  - Don Honorio Ventura State University  (DHVSU / Don Honorio)

Usage:
    python alumni_tracer.py

Requirements:
    pip install -r requirements.txt
"""

import re
import time
import unicodedata
import logging
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Optional: set to True to enable live web scraping (requires selenium) ──
ENABLE_WEB_SCRAPING = True

# ── Input / output file paths ───────────────────────────────────────────────
NAMES_FILE      = "list of names.xlsx"
RESPONSES_FILE  = "College of Computing Studies - Alumni Employment Tracer Form (Responses) (2).xlsx"
OUTPUT_FILE     = "alumni_tracer_output.xlsx"

# ── University filter ────────────────────────────────────────────────────────
# Any alumni whose school field (col C in the names file, or scraped from the
# web) does NOT contain at least one of these substrings will be SKIPPED.
# The responses file is school-specific (it belongs to the college of the
# accepted universities), so every row in it is automatically accepted.
ACCEPTED_UNIVERSITIES: list[str] = [
    "pampanga state university",
    "pamsu",
    "psu",
    "don honorio ventura state university",
    "dhvsu",
    "don honorio",
]

# ── Logging ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Column headers for the output xlsx ──────────────────────────────────────
OUTPUT_COLUMNS = [
    "LAST NAME",
    "GIVEN NAME",
    "MIDDLE INITIAL",
    "YEAR BATCH",                   # from list of names
    "SCHOOL / UNIVERSITY",          # confirmed school affiliation
    "PRESENT ADDRESS",
    "CONTACT NUMBER",
    "PROGRAM TAKEN",
    "YEAR OF GRADUATION",
    "EMPLOYMENT STATUS",
    "IF SELF-EMPLOYED: BUSINESS NAME & DESCRIPTION",
    "PURSUING FURTHER STUDIES?",
    "FURTHER STUDIES DETAILS",
    "SECTOR",
    "EMPLOYER NAME / CONTACT",
    "WORK LOCATION",
    "EMPLOYER CLASSIFICATION",
    "CHALLENGES IN LOOKING FOR A JOB",
    "SKILLS/COMPETENCIES LACKING AFTER GRADUATION",
    "CONSENT",
    "DATA SOURCE",                  # 'existing_response' | 'web_search' | 'not_found'
    "LINKEDIN PROFILE",             # LinkedIn profile URL (if found via web search)
]

# ── Mapping: response-file column letter (0-based index) → output column ────
# Based on the header row of the responses file:
# A=Timestamp B=Email C=Email2 D=LAST NAME E=GIVEN NAME F=MIDDLE INITIAL
# G=PRESENT ADDRESS H=CONTACT NUMBER I=PROGRAM TAKEN J=YEAR OF GRADUATION
# K=EMPLOYMENT STATUS L=SELF-EMPLOYED BUSINESS M=FURTHER STUDIES
# N=FURTHER STUDIES DETAILS O=SECTOR P=EMPLOYER Q=WORK LOCATION
# R=EMPLOYER CLASSIFICATION S=CHALLENGES T=SKILLS LACKING U=CONSENT
RESP_COL = {
    "last_name":          3,   # D (0-based)
    "given_name":         4,   # E
    "middle_initial":     5,   # F
    "address":            6,   # G
    "contact":            7,   # H
    "program":            8,   # I
    "grad_year":          9,   # J
    "emp_status":         10,  # K
    "self_emp_biz":       11,  # L
    "further_studies":    12,  # M
    "further_details":    13,  # N
    "sector":             14,  # O
    "employer":           15,  # P
    "work_location":      16,  # Q
    "emp_classification": 17,  # R
    "challenges":         18,  # S
    "skills_lacking":     19,  # T
    "consent":            20,  # U
}


# ────────────────────────────────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────────────────────────────────

def normalize(text: str) -> str:
    """Lower-case, strip accents, collapse whitespace for fuzzy matching."""
    if not text:
        return ""
    text = str(text)
    # Remove accents
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", text).strip().lower()


def parse_name(raw: str):
    """
    Split 'Surname, Given Name Middle.' into components.
    Returns (last, given, middle).
    """
    raw = raw.strip()
    if "," in raw:
        parts = raw.split(",", 1)
        last = parts[0].strip()
        rest = parts[1].strip()
        # Middle initial is usually a single letter with optional '.'
        mi_match = re.search(r"\b([A-Za-z])\.$", rest)
        if mi_match:
            middle = mi_match.group(1).upper()
            given = rest[: mi_match.start()].strip()
        else:
            # Try trailing single character without dot
            words = rest.split()
            if words and len(words[-1]) == 1 and words[-1].isalpha():
                middle = words[-1].upper()
                given = " ".join(words[:-1])
            else:
                middle = ""
                given = rest
    else:
        # No comma — treat whole thing as last name
        last = raw
        given = ""
        middle = ""
    return last.strip(), given.strip(), middle.strip()


def row_to_dict(ws_row, col_map: dict) -> dict:
    """Convert an openpyxl row (tuple of cells) to a dict using col_map."""
    result = {}
    for key, idx in col_map.items():
        cell = ws_row[idx] if idx < len(ws_row) else None
        result[key] = str(cell.value).strip() if cell and cell.value is not None else ""
    return result


def is_accepted_university(school: str) -> bool:
    """
    Return True if `school` matches at least one of the accepted universities.
    Matching is case-insensitive substring search after accent-stripping.
    An empty school string returns False.
    """
    if not school:
        return False
    normed = normalize(school)
    return any(kw in normed for kw in ACCEPTED_UNIVERSITIES)


def canonical_school_name(school: str) -> str:
    """
    Map any accepted variant to the canonical full name, or return the
    original string if it already looks like a full name.
    """
    normed = normalize(school)
    if any(kw in normed for kw in ("don honorio ventura state university", "dhvsu", "don honorio")):
        return "Don Honorio Ventura State University (DHVSU)"
    if any(kw in normed for kw in ("pampanga state university", "pamsu", "psu")):
        return "Pampanga State University (PSU / PamSU)"
    return school.strip()


# ────────────────────────────────────────────────────────────────────────────
# Load existing responses
# ────────────────────────────────────────────────────────────────────────────

def load_existing_responses(filepath: str) -> list[dict]:
    """Return list of response dicts from the alumni form responses file.

    The responses file belongs to the College of Computing Studies of the
    accepted universities, so every row is implicitly from those schools.
    """
    log.info("Loading existing responses from: %s", filepath)
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    responses = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        data = row_to_dict(row, RESP_COL)
        if not data.get("last_name") and not data.get("given_name"):
            continue
        responses.append(data)
    log.info("  Loaded %d responses.", len(responses))
    return responses


# ────────────────────────────────────────────────────────────────────────────
# Load name list
# ────────────────────────────────────────────────────────────────────────────

# Column layout of 'list of names.xlsx' (0-based):
#   Col 0 – Name  (e.g. "Dela Cruz, Juan M.")
#   Col 1 – Batch (e.g. "2023-2024")
#   Col 2 – School / University  (optional; if blank, the alumni is assumed to
#             belong to one of the accepted universities because the names file
#             itself is scoped to those institutions)
_NAMES_COL_SCHOOL = 2

def load_name_list(filepath: str) -> list[dict]:
    """Return list of {raw, last, given, middle, batch, school} from the names file.

    Alumni whose school column is non-empty AND does NOT match the accepted
    universities are skipped with a warning.  Alumni with an empty school
    column are assumed to be from the accepted universities (the names list
    is itself scoped to those schools).
    """
    log.info("Loading name list from: %s", filepath)
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    names = []
    skipped = 0
    for row in ws.iter_rows(min_row=1, values_only=True):
        raw_name = row[0]
        batch    = row[1] if len(row) > 1 else ""
        school   = row[_NAMES_COL_SCHOOL] if len(row) > _NAMES_COL_SCHOOL else None

        if not raw_name:
            continue

        school_str = str(school).strip() if school else ""

        # ── University filter ────────────────────────────────────────────────
        if school_str:
            # A school is explicitly specified — validate it.
            if not is_accepted_university(school_str):
                log.warning(
                    "  SKIPPED (not an accepted university): %s  [school=%s]",
                    raw_name, school_str,
                )
                skipped += 1
                continue
            confirmed_school = canonical_school_name(school_str)
        else:
            # No school column — the names file is scoped to the accepted
            # universities, so we label it as "PSU / DHVSU (assumed)".
            confirmed_school = "Pampanga State University / DHVSU (assumed)"

        last, given, middle = parse_name(str(raw_name))
        names.append({
            "raw":    str(raw_name).strip(),
            "last":   last,
            "given":  given,
            "middle": middle,
            "batch":  str(batch).strip() if batch else "",
            "school": confirmed_school,
        })

    log.info("  Loaded %d names (%d skipped — not an accepted university).", len(names), skipped)
    return names


# ────────────────────────────────────────────────────────────────────────────
# Name matching
# ────────────────────────────────────────────────────────────────────────────

def match_score(alumni: dict, response: dict) -> float:
    """
    Return a score 0-1 indicating how well a name-list entry matches
    a response record.

    Rules:
    - Last name must be an exact normalized match. Score = 0 otherwise.
    - Given name first token must also match to confirm the right person
      (prevents cross-matching siblings or people with the same last name).
    - Score 1.0 = exact last + exact given
    - Score 0.85 = exact last + first given-name token matches
    - Score 0.7  = exact last + no given name available on one side
    - Score 0.0  = last mismatch OR given-name first-token mismatch
    """
    last_a = normalize(alumni["last"])
    last_r = normalize(response.get("last_name", ""))

    given_a = normalize(alumni["given"])
    given_r = normalize(response.get("given_name", ""))

    # ── Last name: exact match only ──────────────────────────────────────────
    if last_a != last_r:
        return 0.0

    # ── Given name check ─────────────────────────────────────────────────────
    if not given_a or not given_r:
        # One side has no given name — last name alone is sufficient
        return 0.7

    if given_a == given_r:
        return 1.0

    # First token must match (e.g. "Alfred" vs "Ferwinkle" → reject)
    if given_a.split()[0] == given_r.split()[0]:
        return 0.85

    # First tokens differ → not the same person
    return 0.0


def find_best_match(alumni: dict, responses: list[dict]) -> dict | None:
    """Find the best matching response for an alumni entry, or None."""
    best = None
    best_score = 0.0
    for resp in responses:
        s = match_score(alumni, resp)
        if s > best_score:
            best_score = s
            best = resp
    # Accept only matches where last name + at least partial given name matched
    if best_score >= 0.7:
        return best
    return None


# ────────────────────────────────────────────────────────────────────────────
# Web search (optional, requires selenium + chromedriver)
# ────────────────────────────────────────────────────────────────────────────

# Keywords used both in the search query and to validate scraped snippets.
_SCHOOL_QUERY_TERMS = (
    "Pampanga State University OR "
    "Don Honorio Ventura State University OR DHVSU OR PamSU"
)

def web_search_alumni(alumni: dict) -> dict | None:
    """
    Try to find alumni info via a Google/LinkedIn search scoped to the
    accepted universities.

    Returns a partial dict if anything useful is found AND the result can be
    confirmed as belonging to PSU / DHVSU.  Returns None otherwise.

    Requires: selenium, webdriver-manager  (pip install selenium webdriver-manager)
    """
    if not ENABLE_WEB_SCRAPING:
        return None

    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.common.by import By
        from selenium.webdriver.chrome.options import Options
        from webdriver_manager.chrome import ChromeDriverManager
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError:
        log.warning("selenium / webdriver-manager not installed — skipping web search.")
        return None

    # Scope the query to the accepted universities
    name_query = (
        f"{alumni['given']} {alumni['last']} "
        f"({_SCHOOL_QUERY_TERMS}) Philippines LinkedIn"
    )
    log.info("  Web-searching: %s", name_query)

    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )

    driver = None
    result = {}
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(f"https://www.google.com/search?q={name_query.replace(' ', '+')}")
        time.sleep(2)

        snippets = driver.find_elements(By.CSS_SELECTOR, "div.VwiC3b")
        all_text = " ".join(s.text for s in snippets[:10])

        # ── University filter: only accept if a school keyword appears ───────
        if not is_accepted_university(all_text):
            log.info(
                "  Web result for %s does not mention an accepted university — discarded.",
                alumni["raw"],
            )
            return None

        # Determine which school was mentioned
        for snippet in snippets[:10]:
            if is_accepted_university(snippet.text):
                result["school"] = canonical_school_name(snippet.text)
                break
        if "school" not in result:
            result["school"] = alumni.get("school", "PSU / DHVSU (web-confirmed)")

        # Try to get LinkedIn snippet from Google results
        for s in snippets[:5]:
            text = s.text
            if any(kw in text.lower() for kw in ["software", "engineer", "developer",
                                                   "analyst", "manager", "specialist",
                                                   "philippines", "pampanga"]):
                result["employer"] = text[:200]
                break

        # Try to grab a LinkedIn profile URL from search result links
        for link in driver.find_elements(By.CSS_SELECTOR, "a"):
            href = link.get_attribute("href") or ""
            if "linkedin.com/in/" in href:
                result["linkedin_url"] = href
                break

        # Check if any snippet mentions employment status
        if "self-employed" in all_text.lower() or "freelance" in all_text.lower():
            result["emp_status"] = (
                "FREE LANCE WORKER (You have no technical employer but conducts project-based works.)"
            )
        elif "unemployed" in all_text.lower():
            result["emp_status"] = (
                "UNEMPLOYED (You are currently on a break, and you are still looking for a job.)"
            )

    except Exception as e:
        log.warning("  Web search error for %s: %s", alumni["raw"], e)
    finally:
        if driver:
            driver.quit()

    time.sleep(1)  # polite delay
    return result if result else None


# ────────────────────────────────────────────────────────────────────────────
# Build output row
# ────────────────────────────────────────────────────────────────────────────

NOT_FOUND_VALUE = "No data found"

# The responses file is scoped to the accepted universities' College of
# Computing Studies; label all existing-response rows accordingly.
_RESPONSES_FILE_SCHOOL = "Pampanga State University / DHVSU (form response)"


def build_output_row(alumni: dict, response: dict | None, web_data: dict | None) -> list:
    """Compose the OUTPUT_COLUMNS row for a single alumni entry."""
    if response:
        source = "existing_response"
        school = alumni.get("school") or _RESPONSES_FILE_SCHOOL
        grad_year = response.get("grad_year", NOT_FOUND_VALUE)
        # Clean up year — it may be "2024.0" from numeric cell
        try:
            grad_year = str(int(float(grad_year)))
        except (ValueError, TypeError):
            pass
        return [
            response.get("last_name",          NOT_FOUND_VALUE),
            response.get("given_name",          NOT_FOUND_VALUE),
            response.get("middle_initial",      NOT_FOUND_VALUE),
            alumni["batch"],
            school,
            response.get("address",             NOT_FOUND_VALUE),
            response.get("contact",             NOT_FOUND_VALUE),
            response.get("program",             NOT_FOUND_VALUE),
            grad_year,
            response.get("emp_status",          NOT_FOUND_VALUE),
            response.get("self_emp_biz",        NOT_FOUND_VALUE),
            response.get("further_studies",     NOT_FOUND_VALUE),
            response.get("further_details",     NOT_FOUND_VALUE),
            response.get("sector",              NOT_FOUND_VALUE),
            response.get("employer",            NOT_FOUND_VALUE),
            response.get("work_location",       NOT_FOUND_VALUE),
            response.get("emp_classification",  NOT_FOUND_VALUE),
            response.get("challenges",          NOT_FOUND_VALUE),
            response.get("skills_lacking",      NOT_FOUND_VALUE),
            response.get("consent",             NOT_FOUND_VALUE),
            source,
            NOT_FOUND_VALUE,        # LINKEDIN PROFILE (not in form responses)
        ]

    # No existing response — try web data
    last, given, middle = alumni["last"], alumni["given"], alumni["middle"]
    school = alumni.get("school", NOT_FOUND_VALUE)

    if web_data:
        source = "web_search"
        return [
            last,
            given,
            middle,
            alumni["batch"],
            web_data.get("school", school),
            web_data.get("address",            NOT_FOUND_VALUE),
            web_data.get("contact",            NOT_FOUND_VALUE),
            web_data.get("program",            NOT_FOUND_VALUE),
            web_data.get("grad_year",          NOT_FOUND_VALUE),
            web_data.get("emp_status",         NOT_FOUND_VALUE),
            web_data.get("self_emp_biz",       NOT_FOUND_VALUE),
            web_data.get("further_studies",    NOT_FOUND_VALUE),
            web_data.get("further_details",    NOT_FOUND_VALUE),
            web_data.get("sector",             NOT_FOUND_VALUE),
            web_data.get("employer",           NOT_FOUND_VALUE),
            web_data.get("work_location",      NOT_FOUND_VALUE),
            web_data.get("emp_classification", NOT_FOUND_VALUE),
            web_data.get("challenges",         NOT_FOUND_VALUE),
            web_data.get("skills_lacking",     NOT_FOUND_VALUE),
            web_data.get("consent",            NOT_FOUND_VALUE),
            source,
            web_data.get("linkedin_url",       NOT_FOUND_VALUE),
        ]

    # Truly not found
    return [
        last, given, middle, alumni["batch"], school,
        NOT_FOUND_VALUE, NOT_FOUND_VALUE, NOT_FOUND_VALUE,
        NOT_FOUND_VALUE, NOT_FOUND_VALUE, NOT_FOUND_VALUE,
        NOT_FOUND_VALUE, NOT_FOUND_VALUE, NOT_FOUND_VALUE,
        NOT_FOUND_VALUE, NOT_FOUND_VALUE, NOT_FOUND_VALUE,
        NOT_FOUND_VALUE, NOT_FOUND_VALUE, NOT_FOUND_VALUE,
        "not_found",
        NOT_FOUND_VALUE,            # LINKEDIN PROFILE
    ]


# ────────────────────────────────────────────────────────────────────────────
# Excel output formatting
# ────────────────────────────────────────────────────────────────────────────

# Colours
C_HEADER_BG   = "2D5FA3"   # dark blue
C_HEADER_FG   = "FFFFFF"   # white
C_FOUND_BG    = "E8F5E9"   # light green  (existing response)
C_WEB_BG      = "FFF9C4"   # light yellow (web search)
C_NOTFOUND_BG = "FFCDD2"   # light red    (not found)
C_ALT_BG      = "F7F8FA"   # alternate row


def make_fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def make_border() -> Border:
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def write_output_xlsx(rows: list[list], filepath: str, skipped: int = 0):
    """Write the output rows to a nicely formatted xlsx file."""
    log.info("Writing output to: %s", filepath)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alumni Tracer Results"

    # ── Header row ───────────────────────────────────────────────────────────
    ws.append(OUTPUT_COLUMNS)
    header_fill   = make_fill(C_HEADER_BG)
    header_font   = Font(bold=True, color=C_HEADER_FG, size=11)
    header_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_border = make_border()

    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = header_align
        cell.border    = header_border

    ws.row_dimensions[1].height = 40

    # ── Data rows ────────────────────────────────────────────────────────────
    fill_found    = make_fill(C_FOUND_BG)
    fill_web      = make_fill(C_WEB_BG)
    fill_notfound = make_fill(C_NOTFOUND_BG)
    fill_alt      = make_fill(C_ALT_BG)
    fill_white    = make_fill("FFFFFF")
    border        = make_border()
    wrap_align    = Alignment(vertical="top", wrap_text=True)

    for i, row_data in enumerate(rows, start=2):
        ws.append(row_data)
        source = row_data[-1] if row_data else "not_found"

        if source == "existing_response":
            row_fill = fill_found
        elif source == "web_search":
            row_fill = fill_web
        elif source == "not_found":
            row_fill = fill_notfound
        else:
            row_fill = fill_alt if i % 2 == 0 else fill_white

        for cell in ws[i]:
            cell.fill      = row_fill
            cell.border    = border
            cell.alignment = wrap_align

    # ── Column widths ────────────────────────────────────────────────────────
    col_widths = [
        20, 25, 15, 12, 42,          # LAST, GIVEN, MI, BATCH, SCHOOL
        35, 18, 15, 12, 45, 50,      # ADDRESS … EMPLOYMENT STATUS, SELF-EMP
        12, 50, 25, 50, 15, 35,      # FURTHER … WORK LOCATION, EMP CLASS
        80, 80, 10, 18, 45,          # CHALLENGES, SKILLS, CONSENT, SOURCE, LINKEDIN
    ]
    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # ── Freeze header ────────────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Summary sheet ────────────────────────────────────────────────────────
    ws_summary = wb.create_sheet("Summary")
    total     = len(rows)
    found     = sum(1 for r in rows if r[-1] == "existing_response")
    web       = sum(1 for r in rows if r[-1] == "web_search")
    not_found = sum(1 for r in rows if r[-1] == "not_found")

    summary_data = [
        ["Metric", "Count", "Percentage"],
        ["Total alumni processed", total, "100%"],
        ["Found in existing responses", found,
         f"{found/total*100:.1f}%" if total else "0%"],
        ["Found via web search", web,
         f"{web/total*100:.1f}%" if total else "0%"],
        ["Not found", not_found,
         f"{not_found/total*100:.1f}%" if total else "0%"],
        ["Skipped (not accepted university)", skipped, "—"],
    ]

    for i, row in enumerate(summary_data, start=1):
        ws_summary.append(row)
        for cell in ws_summary[i]:
            cell.border = make_border()
            if i == 1:
                cell.fill      = make_fill(C_HEADER_BG)
                cell.font      = Font(bold=True, color=C_HEADER_FG)
                cell.alignment = Alignment(horizontal="center")

    ws_summary.column_dimensions["A"].width = 40
    ws_summary.column_dimensions["B"].width = 12
    ws_summary.column_dimensions["C"].width = 12

    wb.save(filepath)
    log.info("  Saved: %s", filepath)
    return total, found, web, not_found


# ────────────────────────────────────────────────────────────────────────────
# Main
# ────────────────────────────────────────────────────────────────────────────

def main():
    # ── Validate inputs ──────────────────────────────────────────────────────
    for fp in (NAMES_FILE, RESPONSES_FILE):
        if not Path(fp).exists():
            log.error("File not found: %s", fp)
            raise SystemExit(1)

    # ── Load data ────────────────────────────────────────────────────────────
    name_list = load_name_list(NAMES_FILE)
    responses = load_existing_responses(RESPONSES_FILE)

    # Count how many were skipped in the loader
    all_raw_count = 0
    wb_check = openpyxl.load_workbook(NAMES_FILE, data_only=True)
    for row in wb_check.active.iter_rows(min_row=1, values_only=True):
        if row[0]:
            all_raw_count += 1
    wb_check.close()
    skipped_count = all_raw_count - len(name_list)

    # ── Process each alumni ──────────────────────────────────────────────────
    output_rows: list[list] = []
    stats = {"found": 0, "web": 0, "not_found": 0}

    log.info("Matching %d alumni against %d responses …", len(name_list), len(responses))
    for alumni in name_list:
        log.info("  Processing: %s  [%s]", alumni["raw"], alumni["school"])

        # 1. Try exact/fuzzy match in existing responses
        match = find_best_match(alumni, responses)
        if match:
            log.info("    ✓ Found in existing responses")
            stats["found"] += 1
            output_rows.append(build_output_row(alumni, match, None))
            continue

        # 2. Try web search (only if enabled)
        web_data = None
        if ENABLE_WEB_SCRAPING:
            web_data = web_search_alumni(alumni)
            if web_data:
                log.info("    ~ Found via web search")
                stats["web"] += 1
                output_rows.append(build_output_row(alumni, None, web_data))
                continue

        # 3. Not found
        log.info("    ✗ Not found")
        stats["not_found"] += 1
        output_rows.append(build_output_row(alumni, None, None))

    # ── Write output ─────────────────────────────────────────────────────────
    total, found, web, not_found = write_output_xlsx(
        output_rows, OUTPUT_FILE, skipped=skipped_count
    )

    # ── Final report ─────────────────────────────────────────────────────────
    print()
    print("=" * 65)
    print("  ALUMNI TRACER — RESULTS")
    print("  University filter: PSU / DHVSU only")
    print("=" * 65)
    print(f"  Names in input file          : {all_raw_count}")
    print(f"  Skipped (wrong university)   : {skipped_count}")
    print(f"  Total alumni processed       : {total}")
    print(f"  Found (existing data)        : {found}  ({found/total*100:.1f}%)" if total else
          f"  Found (existing data)        : {found}")
    print(f"  Found (web search)           : {web}  ({web/total*100:.1f}%)" if total else
          f"  Found (web search)           : {web}")
    print(f"  Not found                    : {not_found}  ({not_found/total*100:.1f}%)" if total else
          f"  Not found                    : {not_found}")
    print("=" * 65)
    print(f"  Output saved to: {OUTPUT_FILE}")
    print()


if __name__ == "__main__":
    main()
