"""
Pillar article stat extractor.

Scans public pillar-cluster article pages and flags statistics likely to
need an annual refresh: numbers tied to an explicit year (e.g. "$106,950
in 2024") as well as numbers attached to a stats keyword even with no
year in the same sentence (e.g. "median pay is $106,950").

Usage:
    python extract_stats.py <urls.txt | urls_dir> [output.xlsx]

Output defaults to "Pillar Stat Review - <Mon DD YYYY>.xlsx" (today's date)
in the current directory if no output path is given.

If given a single file: one article URL per line (blank lines and lines
starting with # are ignored), treated as one unnamed pillar.

If given a directory: every *.txt file inside is treated as one pillar
(the filename, minus extension, is the pillar name), each holding its
own list of article URLs in the same format. This is the normal mode
for the recurring quarterly skim, e.g.:

    urls/biomedical-engineering.txt
    urls/journalism.txt
    urls/law-school.txt
    urls/premed.txt
"""

import re
import sys
import time
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HEADERS = {"User-Agent": "Mozilla/5.0 (pillar-stats-extractor)"}

YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
NUMBER_RE = re.compile(
    r"\$\s?\d[\d,]*(\.\d+)?"       # dollar amounts
    r"|\d[\d,]*(\.\d+)?\s?%"       # percentages
    r"|\b\d{1,3}(,\d{3})+\b"       # comma-grouped numbers (enrollment, etc.)
)

SOURCE_PATTERNS = [
    ("IPEDS", re.compile(r"\bIPEDS\b", re.I)),
    ("BLS", re.compile(r"Bureau of Labor Statistics|\bBLS\b", re.I)),
    ("U.S. News & World Report", re.compile(r"U\.?S\.? News", re.I)),
    ("Niche", re.compile(r"\bNiche\b")),
    ("PayScale", re.compile(r"PayScale", re.I)),
]

STAT_TYPE_PATTERNS = [
    ("Wage", re.compile(r"salary|salaries|wage|pay\b|earn|income", re.I)),
    ("Outlook", re.compile(r"outlook|growth|demand|projected|job growth|decline", re.I)),
    ("Enrollment", re.compile(r"enroll|completions?|graduates?", re.I)),
    ("Cost/Tuition", re.compile(r"tuition|cost\b|price", re.I)),
    ("Ranking", re.compile(r"rank(ed|ing)?|top \d", re.I)),
]

# Catches stat-bearing sentences even when no year is stated nearby, e.g.
# "median pay is $106,950" (no "2024" in the sentence).
STAT_KEYWORD_RE = re.compile(
    r"|".join(p.pattern for _, p in STAT_TYPE_PATTERNS)
    + r"|median|average|mean\b|according to",
    re.I,
)

# Lead-in sentences that introduce a table/stat rather than citing one
# themselves, e.g. "Below you'll find the top 10 colleges...". These get
# flagged as false positives if they happen to mention a threshold number
# ("20,000 total enrollment") near a stats keyword ("tuition").
INTRO_PHRASE_RE = re.compile(
    r"below you.ll find|you.ll find the|here are the|the following table"
    r"|the table below|as shown below|see the table",
    re.I,
)

NOISE_CLASS_HINTS = ("breadcrumb", "webform", "menu", "nav")
SOURCE_LINE_RE = re.compile(r"Source\s*:.*?[.!?]", re.I)

# Periods inside these should not be treated as sentence boundaries.
_PROTECTED_ABBREVIATIONS = [
    "U.S.A.", "U.S.", "U.K.", "Ph.D.", "M.D.", "vs.", "etc.",
    "approx.", "Dr.", "Mr.", "Ms.", "Mrs.", "Jr.", "Sr.", "St.",
    "Inc.", "Ltd.", "Co.",
]
_ABBREV_PLACEHOLDER = "․"  # one dot leader -- visually a period, not matched by [.!?]


def guess_source(text):
    for label, pattern in SOURCE_PATTERNS:
        if pattern.search(text):
            return label
    return "Unknown (needs review)"


def guess_stat_type(text):
    for label, pattern in STAT_TYPE_PATTERNS:
        if pattern.search(text):
            return label
    return "Other"


def suggested_cadence(source_label):
    if source_label in ("IPEDS", "BLS"):
        return "Annual"
    return "Review"


def latest_bls_wage_year(today=None):
    # BLS OEWS publishes new annual wage estimates each spring (historically
    # ~April), covering May-of-prior-year data. Heuristic, not a live check
    # of BLS's release calendar.
    today = today or date.today()
    return today.year - 1 if today.month >= 4 else today.year - 2


def latest_ipeds_completions_year(today=None):
    # IPEDS Completions data for academic year YYYY-(YY+1) is broadly
    # available roughly mid-the-following-year. Heuristic, not a live check
    # of IPEDS's release calendar. Year returned is the academic year's
    # *start* year (matches how these articles cite it, e.g. "2023-24").
    today = today or date.today()
    return today.year - 1 if today.month >= 7 else today.year - 2


def assess_staleness(source_label, years_str, today=None):
    if source_label not in ("BLS", "IPEDS"):
        return "Source unclear — manual check"
    if not years_str:
        return "No year found — manual check"

    year_ints = [int(y) for y in years_str.split(", ") if y.strip().isdigit()]
    if not year_ints:
        return "No year found — manual check"

    cited_year = max(year_ints)
    latest_expected = (
        latest_bls_wage_year(today) if source_label == "BLS"
        else latest_ipeds_completions_year(today)
    )

    if cited_year < latest_expected:
        return (f"Likely stale — article cites {cited_year}, "
                f"~{latest_expected} data should now be available")
    if cited_year == latest_expected:
        return "Current — matches latest expected release"
    return "Current — cites data newer than heuristic expects (verify)"


def is_noise(el):
    ident = " ".join((el.get("class") or []) + [el.get("id") or ""]).lower()
    return any(hint in ident for hint in NOISE_CLASS_HINTS)


def split_sentences(text):
    protected = text
    for abbr in _PROTECTED_ABBREVIATIONS:
        protected = protected.replace(abbr, abbr.replace(".", _ABBREV_PLACEHOLDER))
    parts = re.split(r"(?<=[.!?])\s+", protected)
    return [p.replace(_ABBREV_PLACEHOLDER, ".") for p in parts]


US_STATES = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
    "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
    "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada",
    "New Hampshire", "New Jersey", "New Mexico", "New York",
    "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
    "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
    "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
    "West Virginia", "Wisconsin", "Wyoming",
]


def guess_stat_name(quote, stat_type, block_type):
    """Best-effort short label for the stat -- a starting point to hand-edit,
    not a substitute for human review."""
    if block_type == "Table":
        header = quote.split(" | ")[0].strip()
        return header or "Data table"

    geography = None
    if re.search(r"United States|\bU\.S\.", quote):
        geography = "US"
    else:
        for state in US_STATES:
            if re.search(rf"\b{re.escape(state)}\b", quote):
                geography = state
                break

    base = stat_type if stat_type != "Other" else "Stat"
    return f"{base} ({geography})" if geography else base


def find_main_content(soup):
    article = soup.find("article")
    if article:
        return article
    return soup.find("main") or soup.find("body")


def fetch_article(url):
    resp = requests.get(url, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    content = find_main_content(soup)
    if content is None:
        return url, []

    title_tag = content.find("h1") or soup.title
    title = title_tag.get_text(strip=True) if title_tag else url

    for tag in content.find_all(["script", "style", "form", "nav"]):
        tag.decompose()

    rows = []
    seen_tables = set()

    blocks = content.find_all(["p", "li", "h2", "h3", "h4", "table"])
    for block in blocks:
        if is_noise(block):
            continue

        if block.name == "table":
            if id(block) in seen_tables:
                continue
            seen_tables.add(id(block))

            table_text = block.get_text(" | ", strip=True)

            context_text = table_text
            node = block
            hops = 0
            climbs = 0
            while hops < 3:
                nxt = node.find_next_sibling()
                if nxt is None:
                    # table may be wrapped in a layout div (e.g. div.table-responsive)
                    # with no sibling of its own; climb one level and retry.
                    if climbs < 2 and node.parent is not None and node.parent.name in ("div", "section"):
                        node = node.parent
                        climbs += 1
                        continue
                    break
                nxt_text = nxt.get_text(" ", strip=True)
                if nxt_text:
                    context_text += " || " + nxt_text
                    if re.search(r"source\s*:", nxt_text, re.I):
                        break
                node = nxt
                hops += 1

            years = sorted(set(m.group(0) for m in YEAR_RE.finditer(context_text)))
            if not years and not STAT_KEYWORD_RE.search(context_text) and not NUMBER_RE.search(context_text):
                continue

            quote = table_text
            if len(quote) > 300:
                quote = quote[:300] + "..."

            tail = context_text[len(table_text):]
            source_line_match = SOURCE_LINE_RE.search(tail)
            merge_key = (source_line_match.group(0).strip().lower()
                         if source_line_match else (tail or None))

            rows.append({
                "block_type": "Table",
                "quote": quote,
                "years": ", ".join(years),
                "source": guess_source(context_text),
                "stat_type": guess_stat_type(context_text),
                "_merge_key": merge_key,
            })
            continue

        text = block.get_text(" ", strip=True)
        if not text:
            continue

        for sentence in split_sentences(text):
            if INTRO_PHRASE_RE.search(sentence):
                continue
            has_number = NUMBER_RE.search(sentence)
            has_year = YEAR_RE.search(sentence)
            if has_number and (has_year or STAT_KEYWORD_RE.search(sentence)):
                years = sorted(set(m.group(0) for m in YEAR_RE.finditer(sentence)))
                quote = sentence if len(sentence) <= 300 else sentence[:300] + "..."
                rows.append({
                    "block_type": "Paragraph",
                    "quote": quote,
                    "years": ", ".join(years),
                    "source": guess_source(sentence),
                    "stat_type": guess_stat_type(sentence),
                    "_merge_key": None,
                })

    return title, merge_sibling_tables(rows)


def merge_sibling_tables(rows):
    """Adjacent Table rows that share the same trailing source/caption text
    (e.g. two tables both captioned "Source: IPEDS, 2023-24 completions...")
    are presenting one dataset split across tables for display -- collapse
    them into a single row."""
    merged = []
    for row in rows:
        key = row.get("_merge_key")
        if (key and merged and merged[-1].get("block_type") == "Table"
                and merged[-1].get("_merge_key") == key):
            merged[-1]["quote"] += " /// " + row["quote"]
        else:
            merged.append(row)
    for row in merged:
        row.pop("_merge_key", None)
    return merged


def load_urls(path):
    urls = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#"):
                urls.append(line)
    return urls


def pillar_name_from_filename(path):
    return path.stem.replace("-", " ").replace("_", " ").title()


def gather_pillar_url_lists(input_path):
    """Returns list of (pillar_name, [urls])."""
    path = Path(input_path)
    if path.is_dir():
        pillars = []
        for txt_file in sorted(path.glob("*.txt")):
            pillars.append((pillar_name_from_filename(txt_file), load_urls(txt_file)))
        return pillars
    return [(pillar_name_from_filename(path), load_urls(path))]


def default_output_path():
    return f"Pillar Stat Review - {date.today():%b %d %Y}.xlsx"


def stat_fingerprint(row):
    """Identifies "the same stat" by its actual numbers/year/source/type
    rather than exact quote text, since the same figure often gets
    paraphrased slightly between articles (e.g. "the U.S." vs "the United
    States", "Bureau of Labor Statistics" vs "U.S. Bureau of Labor
    Statistics") -- exact-text matching alone would miss those as dupes."""
    numbers = tuple(sorted(m.group(0).strip() for m in NUMBER_RE.finditer(row["quote"])))
    return (row["block_type"], row["source"], row["stat_type"], row["years"], numbers)


def dedupe_across_articles(pillar_rows):
    """pillar_rows: list of (row_dict, title, url). Collapses rows that are
    the same underlying stat (same fingerprint) cited in multiple articles
    into one row with all locations combined."""
    by_key = {}
    order = []
    for row, title, url in pillar_rows:
        key = stat_fingerprint(row)
        if key not in by_key:
            by_key[key] = {**row, "locations": []}
            order.append(key)
        by_key[key]["locations"].append((title, url))
    return [by_key[key] for key in order]


def main():
    if len(sys.argv) not in (2, 3):
        print("Usage: python extract_stats.py <urls.txt | urls_dir> [output.xlsx]")
        print(f"  (output defaults to '{default_output_path()}' if omitted)")
        sys.exit(1)

    input_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) == 3 else default_output_path()
    pillar_url_lists = gather_pillar_url_lists(input_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "Flagged stats"
    header = ["Pillar", "SUMMARY", "Found In", "Block Type", "Quote / Context",
              "Detected Year(s)", "Apparent Source", "Stat Type", "Suggested Cadence",
              "Staleness Assessment"]
    ws.append(header)
    for col in range(1, len(header) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    total_flagged = 0
    total_stale = 0
    for pillar_name, urls in pillar_url_lists:
        print(f"=== Pillar: {pillar_name} ({len(urls)} article(s)) ===")
        pillar_rows = []
        for i, url in enumerate(urls, 1):
            print(f"[{i}/{len(urls)}] Fetching {url}")
            try:
                title, rows = fetch_article(url)
            except Exception as exc:
                print(f"  ERROR: {exc}")
                ws.append([pillar_name, "", url, "ERROR", str(exc), "", "", "", "", ""])
                continue

            print(f"  {len(rows)} flagged stat(s)")
            for row in rows:
                pillar_rows.append((row, title, url))

            time.sleep(0.5)  # be polite to the server

        deduped = dedupe_across_articles(pillar_rows)
        total_flagged += len(deduped)
        for row in deduped:
            staleness = assess_staleness(row["source"], row["years"])
            if staleness.startswith("Likely stale"):
                total_stale += 1
            found_in = "; ".join(f"{t} ({u})" for t, u in row["locations"])
            stat_name = guess_stat_name(row["quote"], row["stat_type"], row["block_type"])
            ws.append([
                pillar_name, stat_name, found_in, row["block_type"], row["quote"],
                row["years"], row["source"], row["stat_type"],
                suggested_cadence(row["source"]), staleness,
            ])

    widths = [20, 32, 45, 10, 55, 14, 22, 14, 16, 45]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    wb.save(out_path)
    total_articles = sum(len(urls) for _, urls in pillar_url_lists)
    print(f"\nDone. {total_flagged} unique stat(s) flagged across {total_articles} article(s) "
          f"in {len(pillar_url_lists)} pillar(s).")
    print(f"{total_stale} flagged as likely stale.")
    print(f"Saved to {out_path}")


if __name__ == "__main__":
    main()
