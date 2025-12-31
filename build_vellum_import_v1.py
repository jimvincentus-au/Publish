#!/usr/bin/env python3
"""
build_vellum_import_v1.py

Build Vellum-ready DOCX files (one per week) from canonical Markdown appendix artifacts.

Contract (v1):
- Input:  Publish/Output/Appendices/Week_XX/<week appendix>.md (tolerant discovery)
- Output: Publish/Output/Vellum/Appendices/Week_XX/Democracy_Clock_Year_One_Week_XX_Appendix.docx

Semantic mapping:
- Week title            -> Heading 1
- Category header       -> Heading 2
- Event line "N. ..."   -> List Number
- Summary paragraphs    -> Normal
"""

import argparse
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from typing import cast

from docx import Document
from docx.document import Document as DocxDocument
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# --- Paths (match Publish pipeline conventions) ---
PUBLISH_ROOT = Path("/Volumes/PRINTIFY24/Democracy Clock Automation/Publish")
PUBLISH_LOGS_DIR = PUBLISH_ROOT / "Logs"

DEFAULT_INPUT_ROOT = PUBLISH_ROOT / "Output" / "Scrivener"
DEFAULT_OUTPUT_ROOT = PUBLISH_ROOT / "Output" / "Vellum"

# Preferred appendix JSON root (authoritative provenance: dates + sources)
DEFAULT_STEP3_WEEKS_ROOT = Path("/Volumes/PRINTIFY24/Democracy Clock Automation/Step 3/Weeks")

LOG_NAME = "vellum_import"


def _resolve_input_root(override: Optional[str]) -> Path:
    if override:
        p = Path(override).expanduser()
        return p

    candidates = [
        DEFAULT_INPUT_ROOT,
    ]

    for c in candidates:
        if c.exists() and c.is_dir():
            return c

    raise FileNotFoundError(
        f"Could not resolve appendix input root at expected location: {DEFAULT_INPUT_ROOT}. " 
        "This project assumes Scrivener output as the canonical source. " 
        "Use --input-root only if the project structure has changed."
    )


def _resolve_output_root(override: Optional[str]) -> Path:
    if override:
        return Path(override).expanduser()
    return DEFAULT_OUTPUT_ROOT


# --- Canonical Appendix Categories (MUST NOT DRIFT) ---
CANONICAL_CATEGORIES_ORDER: List[str] = [
    "Power and Authority",
    "Institutions and Governance",
    "Civil Rights and Dissent",
    "Economic Structure",
    "Information, Memory, and Manipulation",
]

# Legacy/alternate headings that must be mapped into the canonical names above.
# NOTE: Keys must match the cleaned heading text produced by _normalize_category_label.
CATEGORY_ALIASES: Dict[str, str] = {
    # canonical self-maps
    "Power and Authority": "Power and Authority",
    "Institutions and Governance": "Institutions and Governance",
    "Civil Rights and Dissent": "Civil Rights and Dissent",
    "Civil Rights & Dissent": "Civil Rights and Dissent",
    "Economic Structure": "Economic Structure",
    "Information, Memory, and Manipulation": "Information, Memory, and Manipulation",
    # common punctuation/connector variants that appear in some appendix JSON
    "Information, Memory and Manipulation": "Information, Memory, and Manipulation",
    "Information, Memory, and Manipulation": "Information, Memory, and Manipulation",
    "Information, Memory & Manipulation": "Information, Memory, and Manipulation",
    "Information, Memory and Manipulation ": "Information, Memory, and Manipulation",

    # older appendix naming (map into the canonical five)
    "Law and Justice": "Institutions and Governance",
    "Elections and Consent": "Civil Rights and Dissent",
    "Equality and Inclusion": "Civil Rights and Dissent",
    "Truth and Information": "Information, Memory, and Manipulation",
}


# --- Source display names (short tags -> formal/human names) ---
# Keep this mapping tight and deterministic; expand as new feeds are added.
SOURCE_DISPLAY_NAMES: Dict[str, str] = {
    # Core feeds
    "guardian": "The Guardian",
    "hcr": "Letters from an American (Heather Cox Richardson)",
    "popinfo": 'Popular Information',

    # Democracy Clock internal / companion feeds
    "meidas": "Meidas Plus",
    "zeteo": "Zeteo",
    "noah": 'Noahpinion',
    "outloud": 'Democracy Outloud',
    "50501": "The 50501 Movement",

    # Government / official sources
    "orders": "White House / Executive Orders",
    "fr": "Federal Register",
    "federalregister": "Federal Register",
    "congress": "Congress.gov",
    "justsecurity": "Just Security",
    "scotus": "Supreme Court of the United States",
    "dhs": "U.S. Department of Homeland Security",
    "doj": "U.S. Department of Justice",
    "dod": "U.S. Department of Defense",

    # Media wires / outlets
    "nyt": "The New York Times",
    "wapo": "The Washington Post",
    "reuters": "Reuters",
    "ap": "Associated Press",
    "bbc": "BBC News",

    # Other
    "noaa": "NOAA",
}

def _format_source_names(keys: List[str]) -> List[str]:
    out: List[str] = []
    for k in keys:
        kk = str(k).strip()
        if not kk:
            continue
        out.append(SOURCE_DISPLAY_NAMES.get(kk, kk))
    # De-dup, preserve order
    seen = set()
    dedup: List[str] = []
    for s in out:
        if s in seen:
            continue
        seen.add(s)
        dedup.append(s)
    return dedup


def _format_iso_date(d: str) -> str:
    """Convert YYYY-MM-DD to 'D Mon YYYY' (e.g., 2025-01-22 -> 22 Jan 2025)."""
    s = str(d).strip()
    m = re.fullmatch(r"(20\d{2})-(\d{2})-(\d{2})", s)
    if not m:
        return s
    y, mo, da = int(m.group(1)), int(m.group(2)), int(m.group(3))
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    if 1 <= mo <= 12:
        return f"{da} {months[mo-1]} {y}"
    return s


def _format_dates(dates: List[str]) -> List[str]:
    out = [_format_iso_date(x) for x in dates]
    # De-dup, preserve order
    seen = set()
    dedup: List[str] = []
    for s in out:
        if s in seen:
            continue
        seen.add(s)
        dedup.append(s)
    return dedup


def setup_logger(level: str) -> logging.Logger:
    """Create a simple file+console logger (mirrors build_wordpress_import_v1 style)."""
    PUBLISH_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_path = PUBLISH_LOGS_DIR / "build_vellum_import_v1.log"

    logger = logging.getLogger(LOG_NAME)
    logger.setLevel(logging.DEBUG if level == "debug" else logging.INFO)

    if not logger.handlers:
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

        ch = logging.StreamHandler()
        ch.setFormatter(fmt)
        logger.addHandler(ch)

    return logger


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Build Vellum-ready DOCX appendix files from canonical Markdown artifacts"
    )
    p.add_argument("--week", type=int, required=True, help="Starting week number (N)")
    p.add_argument("--weeks", type=int, required=True, help="Number of weeks to include (K)")
    p.add_argument("--level", choices=["info", "debug"], default="info", help="Logging level (default: info)")
    p.add_argument(
        "--input-root",
        type=str,
        default=None,
        help="Override input root folder that contains week appendix markdown folders",
    )
    p.add_argument(
        "--output-root",
        type=str,
        default=None,
        help="Override output root folder for Vellum DOCX outputs",
    )
    return p.parse_args()


def _week_dir_names(week: int) -> List[str]:
    """Return common Week folder naming variants used across pipelines."""
    w2 = f"{week:02d}"
    return [
        f"Week {w2}",
        f"Week {week}",
        f"Week_{w2}",
        f"Week_{week}",
        f"Week{w2}",
        f"Week{week}",
        f"Week-{w2}",
        f"Week-{week}",
    ]


def _discover_week_md(input_root: Path, week: int) -> Path:
    """
    Discover the appendix markdown file for the week with tolerant patterns.
    Hard error if ambiguous or missing.
    """
    week_dir: Optional[Path] = None
    assert input_root is not None
    tried_dirs: List[Path] = []

    for name in _week_dir_names(week):
        cand = input_root / name
        tried_dirs.append(cand)
        if cand.exists() and cand.is_dir():
            week_dir = cand
            break

    if week_dir is None:
        # Fallback: any directory under input_root that contains the week number and starts with 'Week'
        # (e.g., 'Week_01', 'Week 01', 'Week01', etc.)
        fallback = []
        for p in input_root.glob("Week*"):
            if p.is_dir() and re.search(rf"\b0?{week}\b", p.name):
                fallback.append(p)
        if len(fallback) == 1:
            week_dir = fallback[0]
        elif len(fallback) > 1:
            names = ", ".join(x.name for x in sorted(fallback))
            raise RuntimeError(
                f"Ambiguous appendix input folder for Week {week:02d} under {input_root}. Found: {names}. "
                "Standardize Week folder naming or adjust discovery rules."
            )
        else:
            try:
                top = [p.name for p in sorted(input_root.iterdir()) if p.is_dir()][:30]
                logger = logging.getLogger(LOG_NAME)
                logger.debug(f"Top-level dirs under input_root: {top}")
            except Exception:
                pass
            tried = ", ".join(str(p) for p in tried_dirs)
            raise FileNotFoundError(
                f"Missing appendix input folder for Week {week:02d} under {input_root}. Tried: {tried}"
            )

    assert week_dir is not None
    # Common patterns observed / expected
    candidates = [
        week_dir / f"week{week:02d}_appendix.md",
        week_dir / f"week{week}_appendix.md",
        week_dir / f"week{week:02d}-appendix.md",
        week_dir / f"week{week}-appendix.md",
        week_dir / f"appendix_week{week:02d}.md",
        week_dir / f"appendix_week{week}.md",
    ]

    for c in candidates:
        if c.exists() and c.is_file():
            return c

    # Fallback: search for a single plausible appendix markdown file
    # Prefer filenames containing "appendix" and the week number
    patterns = [
        f"*appendix*week{week:02d}*.md",
        f"*appendix*week{week}*.md",
        f"*week{week:02d}*appendix*.md",
        f"*week{week}*appendix*.md",
        "*appendix*.md",
    ]

    found: List[Path] = []
    for pat in patterns:
        found.extend([p for p in week_dir.glob(pat) if p.is_file()])

    # De-dup
    found = sorted(list({p.resolve() for p in found}))

    if len(found) == 0:
        raise FileNotFoundError(
            f"Could not find appendix markdown for Week {week:02d} under {week_dir}. "
            f"Tried: {', '.join(str(x.name) for x in candidates)} and fallback globs."
        )

    # If multiple, require explicit disambiguation by tightening patterns upstream.
    if len(found) > 1:
        names = ", ".join(p.name for p in found)
        raise RuntimeError(
            f"Ambiguous appendix markdown for Week {week:02d} under {week_dir}. Found: {names}. "
            "Ensure only one appendix .md exists per week or standardize the filename."
        )

    return found[0]


def _discover_week_appendix_json(week: int, logger: logging.Logger) -> Optional[Path]:
    """
    Preferred: locate Step 7 output JSON for the week:
      /Step 3/Weeks/Week N/events_appendix_weekN.json
    Returns None if missing.
    """
    week_dir = DEFAULT_STEP3_WEEKS_ROOT / f"Week {week}"
    cand = week_dir / f"events_appendix_week{week}.json"
    if cand.exists() and cand.is_file():
        return cand

    cand2 = week_dir / f"events_appendix_week{week:02d}.json"
    if cand2.exists() and cand2.is_file():
        return cand2

    logger.debug(f"No appendix JSON found for Week {week:02d} under {week_dir}; will fall back to markdown.")
    return None


def _output_docx_path(output_root: Path, md_path: Path) -> Path:
    """
    Mirror the input filename, producing '<weekNN>_vellum.docx' directly under output_root.

    Example:
      input:  week04_appendix.md
      output: /Output/Vellum/week04_vellum.docx
    """
    output_root.mkdir(parents=True, exist_ok=True)

    stem = md_path.stem  # e.g., "week04_appendix"
    m = re.match(r"^(week\d{1,3})", stem, flags=re.IGNORECASE)
    base = m.group(1).lower() if m else stem.lower()

    filename = f"{base}_vellum.docx"
    return output_root / filename


def _normalize_category_label(line: str) -> str:
    """
    Normalize possible Markdown category header lines so they can be matched
    against CANONICAL_CATEGORIES.

    Examples handled:
      "## Power and Authority" -> "Power and Authority"
      "**Power and Authority**" -> "Power and Authority"
      "###   Law and Justice" -> "Law and Justice"
    """
    s = line.strip()

    # Strip leading markdown header markers
    s = re.sub(r"^\s*#{1,6}\s*", "", s).strip()

    # Strip surrounding emphasis markers (common in compiled markdown)
    # Remove repeated leading/trailing * or _ characters
    s = re.sub(r"^[\*_]+\s*", "", s)
    s = re.sub(r"\s*[\*_]+$", "", s)

    s = s.strip()

    # Strip common trailing punctuation that often appears in headings
    s = re.sub(r"[:\-–—\s]+$", "", s).strip()

    return s


def _canonicalize_category(line: str) -> Optional[str]:
    """
    Map a raw markdown line to a canonical category name (or None).
    Handles exact matches and prefix matches (e.g., 'Law and Justice — Courts').
    """
    s = _normalize_category_label(line)
    if not s:
        return None

    # Exact match
    if s in CATEGORY_ALIASES:
        return CATEGORY_ALIASES[s]

    # Prefix match (handles suffix descriptors)
    for k, v in CATEGORY_ALIASES.items():
        if s.startswith(k):
            return v

    return None


def _is_category_header(line: str) -> bool:
    return _canonicalize_category(line) is not None


# Appendix event lines are expected to be in the canonical log form:
#   "N. <summary> — <source(s)> — <domain> — <one-line relevance>"
# We therefore require at least one dash separator (em dash preferred) to avoid
# mistakenly treating numbered sublists (e.g., source lists) as events.
_EVENT_RE = re.compile(r"^\s*(\d+)\.\s+(.*)$")


def _is_event_line(text: str) -> bool:
    """Return True if the line looks like a top-level appendix event line."""
    s = text.strip()
    m = _EVENT_RE.match(s)
    if not m:
        return False

    rest = m.group(2).strip()
    # Require at least one separator that is characteristic of the event-log format.
    # Em dash is canonical; allow spaced hyphen as a fallback.
    if "—" in rest or " - " in rest:
        return True

    return False


def _looks_like_week_title(line: str) -> bool:
    # Expect: "# Week 1 Appendix: ..."
    s = line.strip()
    return s.startswith("#") and "Appendix" in s and "Week" in s


def _strip_md_h1(line: str) -> str:
    # "# Title" -> "Title"
    return line.lstrip("#").strip()


def _split_md_paragraphs(lines: List[str]) -> List[str]:
    """
    Convert markdown lines into paragraph blocks.
    Blank lines separate paragraphs.
    """
    paras: List[str] = []
    buf: List[str] = []
    for raw in lines:
        s = raw.rstrip()
        if not s.strip():
            if buf:
                paras.append(" ".join(x.strip() for x in buf).strip())
                buf = []
            continue
        buf.append(s)
    if buf:
        paras.append(" ".join(x.strip() for x in buf).strip())
    return [p for p in paras if p.strip()]


def _safe_list(x) -> List[str]:
    if x is None:
        return []
    if isinstance(x, list):
        return [str(i).strip() for i in x if str(i).strip()]
    if isinstance(x, str):
        s = x.strip()
        return [s] if s else []
    s = str(x).strip()
    return [s] if s else []


def _event_sources_dates(e: Dict) -> Tuple[List[str], List[str]]:
    """
    Return (sources, dates) using Step 7 enrichment fields when present.
    Accept multiple schema variants to be resilient.
    """
    prov = e.get("provenance") or {}
    sources = (
        _safe_list(prov.get("sources"))
        or _safe_list(e.get("source_names"))
        or _safe_list(e.get("sources"))
    )
    dates = (
        _safe_list(prov.get("dates"))
        or _safe_list(e.get("date"))
        or _safe_list(e.get("event_date"))
        or _safe_list(e.get("dates"))
    )
    return sources, dates


def _event_display_text(e: Dict) -> str:
    for k in ("text", "event", "description", "title", "summary"):
        v = e.get(k)
        if isinstance(v, str) and v.strip():
            return v.strip()
    return str(e).strip()



def _format_appendix_event_text(e: Dict) -> str:
    """Produce the list-item text for Word/Vellum from an appendix event dict."""
    actor = e.get("actor")
    action = e.get("action")
    summary_line = e.get("summary_line")

    # Required core: actor + action. Fall back carefully.
    parts: List[str] = []
    if isinstance(actor, str) and actor.strip():
        parts.append(actor.strip())
    if isinstance(action, str) and action.strip():
        # Sentence-style join (no em-dash) so it reads naturally in print.
        if parts:
            parts[-1] = f"{parts[-1]} {action.strip()}"
        else:
            parts.append(action.strip())

    if not parts:
        # Last-resort fallback: use any reasonable textual field; never str(e)
        parts.append(_event_display_text(e))

    # Optional second sentence
    if isinstance(summary_line, str) and summary_line.strip():
        parts.append(summary_line.strip())

    # Sources + dates (formatted)
    sources, dates = _event_sources_dates(e)
    sources_fmt = _format_source_names(sources)
    dates_fmt = _format_dates(dates)

    tail: List[str] = []
    if sources_fmt:
        tail.append("; ".join(sources_fmt))
    if dates_fmt:
        tail.append("; ".join(dates_fmt))

    if tail:
        parts.append("(" + " | ".join(tail) + ")")

    # Join sentences cleanly
    text = " ".join(p.strip() for p in parts if p.strip())
    return re.sub(r"\s+", " ", text).strip()




def build_docx_from_md(md_path: Path, out_path: Path, logger: logging.Logger) -> None:
    raw = md_path.read_text(encoding="utf-8", errors="replace")
    lines = raw.splitlines()

    doc = Document()

    # State machine:
    # - Expect H1
    # - Summary paragraphs until "Week X Events" marker (line contains "Events")
    # - Then categories and numbered events
    idx = 0

    # 1) Title
    # Find first H1-like line; hard error if missing
    title_line = None
    while idx < len(lines):
        if _looks_like_week_title(lines[idx]):
            title_line = _strip_md_h1(lines[idx])
            idx += 1
            break
        idx += 1

    if not title_line:
        raise ValueError(f"{md_path}: No week appendix title found (expected '# Week … Appendix: …').")

    logger.debug(f"Title resolved: {title_line}")
    doc.add_paragraph(title_line, style="Heading 1")

    # 2) Summary block: consume until we hit an Events marker or a category header
    # We treat all content between title and the first category/events marker as summary.
    summary_lines: List[str] = []
    while idx < len(lines):
        line = lines[idx].rstrip()

        # Stop when we encounter a category heading or the explicit events header
        if _is_category_header(line) or re.fullmatch(r"Week\s+\d+\s+Events", line.strip(), flags=re.IGNORECASE):
            break

        summary_lines.append(line)
        idx += 1

    # Write summary paragraphs (Normal)
    summary_paras = _split_md_paragraphs(summary_lines)
    if summary_paras:
        logger.debug(f"Summary paragraphs: {len(summary_paras)}")
        for p in summary_paras:
            doc.add_paragraph(p, style="Normal")
    else:
        logger.info("No summary paragraphs detected (continuing).")

    # 3) Advance past any explicit “Week X Events” header line(s), if present
    while idx < len(lines):
        line = lines[idx].strip()
        if not line:
            idx += 1
            continue
        # Stop at first category heading or first event line
        if _is_category_header(line) or _is_event_line(line):
            break
        # Skip only the explicit week events header
        if re.fullmatch(r"Week\s+\d+\s+Events", line, flags=re.IGNORECASE):
            logger.debug(f"Skipping events header line: {line}")
            idx += 1
            continue
        idx += 1

    # 4) Buffer all content by canonical category, then emit in canonical order
    buffered: Dict[str, List[Tuple[str, str]]] = {c: [] for c in CANONICAL_CATEGORIES_ORDER}
    categories_seen: List[str] = []
    current_category: Optional[str] = None
    in_category = False
    events_count = 0

    while idx < len(lines):
        raw_line = lines[idx]
        line = raw_line.strip()
        idx += 1

        if not line:
            continue

        # Category heading (canonicalize)
        cat = _canonicalize_category(line)
        if cat:
            prev = current_category
            current_category = cat
            in_category = True
            if cat not in categories_seen:
                categories_seen.append(cat)
            # Log only when the category changes, and include line context.
            if prev != cat:
                logger.debug(f"Category @L{idx}: {cat} (raw='{line}')")
            continue


        # Numbered event item (top-level appendix event line)
        if _is_event_line(line):
            m = _EVENT_RE.match(line)
            assert m is not None

            if not in_category or not current_category:
                # If events begin before any category heading, default to first canonical category
                current_category = CANONICAL_CATEGORIES_ORDER[0]
                in_category = True
                if current_category not in categories_seen:
                    categories_seen.append(current_category)
                logger.debug(f"Category (defaulted): {current_category}")

            # At this point, current_category is guaranteed to be a non-empty str
            assert current_category is not None

            # Remove native number; keep only the post-number text
            rest = m.group(2).strip()
            buffered[current_category].append(("list", rest))
            events_count += 1
            continue

        # Continuation / explanatory text line
        if in_category and current_category:
            buffered[current_category].append(("p", line))
        else:
            # Stray text outside category: attach to the first canonical category so nothing is lost
            buffered[CANONICAL_CATEGORIES_ORDER[0]].append(("p", line))

    # Emit categories in canonical order (always), even if empty
    for cat_name in CANONICAL_CATEGORIES_ORDER:
        doc.add_paragraph(cat_name, style="Heading 2")

        # Explicit numbering restart per category with a hanging indent.
        list_index = 1

        for kind, text in buffered[cat_name]:
            if kind == "list":
                p = doc.add_paragraph(style="Normal")
                p.paragraph_format.left_indent = Inches(0.35)
                p.paragraph_format.first_line_indent = Inches(-0.20)
                p.add_run(f"{list_index}. ")
                p.add_run(text)
                list_index += 1
            else:
                doc.add_paragraph(text, style="Normal")

    # 5) Verification logs
    logger.info(f"Input:  {md_path}")
    logger.info(f"Output: {out_path}")
    logger.info(f"Categories seen ({len(categories_seen)}): {', '.join(categories_seen) if categories_seen else '(none)'}")
    logger.info(f"Events emitted: {events_count}")
    if events_count > 200:
        logger.warning(
            f"Events emitted is unusually high for a single week ({events_count}). "
            "This often indicates the parser is still interpreting numbered sublists as events. "
            "Inspect the input markdown for non-event numbered lines and tighten _is_event_line() if needed."
        )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


def build_docx_from_appendix_json(json_path: Path, out_path: Path, week: int, logger: logging.Logger) -> None:
    import json as _json

    data = _json.loads(json_path.read_text(encoding="utf-8", errors="replace"))

    # Heading 1 title
    title = None
    for k in ("title", "week_title", "heading"):
        v = data.get(k)
        if isinstance(v, str) and v.strip():
            title = v.strip()
            break
    if not title:
        title = f"Week {week} Appendix"

    doc = Document()
    doc.add_paragraph(title, style="Heading 1")

    # Optional summary (string)
    summary = data.get("summary")
    if isinstance(summary, str) and summary.strip():
        for p in _split_md_paragraphs(summary.splitlines()):
            doc.add_paragraph(p, style="Normal")

    raw_categories = data.get("categories")
    cat_to_events: Dict[str, List[Dict]] = {c: [] for c in CANONICAL_CATEGORIES_ORDER}

    def add_events(cat_name: str, events: List[Dict]) -> None:
        canon = _canonicalize_category(cat_name) or CATEGORY_ALIASES.get(_normalize_category_label(cat_name)) or cat_name
        canon = CATEGORY_ALIASES.get(canon, canon)
        if canon not in cat_to_events:
            logger.warning(
                f"Week {week:02d}: Unexpected category in appendix JSON: '{cat_name}' (normalized='{canon}'). "
                "Attaching events to first canonical category."
            )
            canon = CANONICAL_CATEGORIES_ORDER[0]
        cat_to_events[canon].extend([e for e in events if isinstance(e, dict)])

    if isinstance(raw_categories, list):
        for c in raw_categories:
            if not isinstance(c, dict):
                continue
            name = c.get("name") or c.get("category") or c.get("heading") or ""
            evs = c.get("events") or []
            if isinstance(name, str) and isinstance(evs, list):
                add_events(name, evs)
    elif isinstance(raw_categories, dict):
        for name, evs in raw_categories.items():
            if isinstance(name, str) and isinstance(evs, list):
                add_events(name, evs)

    # Emit with validation counters
    events_seen = 0
    events_with_sources = 0
    events_with_dates = 0

    # Track source-key coverage for deterministic reporting
    sources_seen_set = set()
    sources_unknown_set = set()

    for cat_name in CANONICAL_CATEGORIES_ORDER:
        doc.add_paragraph(cat_name, style="Heading 2")

        # Explicit numbering restart per category with a hanging indent.
        list_index = 1

        for e in cat_to_events[cat_name]:
            # provenance counters
            sources, dates = _event_sources_dates(e)
            # Track sources seen and unknowns
            for sk in sources:
                sk2 = str(sk).strip()
                if not sk2:
                    continue
                sources_seen_set.add(sk2)
                if sk2 not in SOURCE_DISPLAY_NAMES:
                    sources_unknown_set.add(sk2)
            events_seen += 1
            if sources:
                events_with_sources += 1
            if dates:
                events_with_dates += 1

            line = _format_appendix_event_text(e)

            p = doc.add_paragraph(style="Normal")
            p.paragraph_format.left_indent = Inches(0.35)
            p.paragraph_format.first_line_indent = Inches(-0.20)
            p.add_run(f"{list_index}. ")
            p.add_run(line)
            list_index += 1

    # Week-level warnings (your “suspicious conditions”)
    if events_seen == 0:
        logger.warning(f"Week {week:02d}: Appendix JSON contains zero events: {json_path}")
    else:
        if events_with_dates == 0:
            logger.warning(f"Week {week:02d}: NO appendix events contain dates (JSON): {json_path}")
        elif events_with_dates < events_seen:
            logger.warning(f"Week {week:02d}: {events_seen - events_with_dates} appendix event(s) missing dates (JSON): {json_path}")

        if events_with_sources == 0:
            logger.warning(f"Week {week:02d}: NO appendix events contain sources (JSON): {json_path}")
        elif events_with_sources < events_seen:
            logger.warning(f"Week {week:02d}: {events_seen - events_with_sources} appendix event(s) missing sources (JSON): {json_path}")

    if sources_seen_set:
        seen_sorted = ", ".join(sorted(sources_seen_set))
        logger.info(f"Week {week:02d}: source keys seen: {seen_sorted}")
    if sources_unknown_set:
        unknown_sorted = ", ".join(sorted(sources_unknown_set))
        logger.warning(
            f"Week {week:02d}: UNKNOWN source key(s) with no formal mapping: {unknown_sorted}. "
            "Add them to SOURCE_DISPLAY_NAMES."
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    logger.info(
        f"Week {week:02d}: events={events_seen}, with_sources={events_with_sources}, with_dates={events_with_dates}"
    )
    logger.info(f"Appendix JSON input: {json_path}")


def main() -> None:
    args = parse_args()
    logger = setup_logger(args.level)

    logger.info("Starting Vellum DOCX build (appendix)")
    logger.debug(f"Args: week={args.week}, weeks={args.weeks}, level={args.level}")

    input_root = _resolve_input_root(args.input_root)
    output_root = _resolve_output_root(args.output_root)

    logger.info(f"Input root:  {input_root}")
    logger.info(f"Output root: {output_root}")

    for week in range(args.week, args.week + args.weeks):
        logger.info(f"Processing Week {week:02d}")

        json_path = _discover_week_appendix_json(week, logger)

        if json_path:
            # Output name stays weekNN_vellum.docx even when input is JSON
            md_stub = Path(f"week{week:02d}_appendix.md")
            out_path = _output_docx_path(output_root, md_stub)

            logger.debug(f"Resolved json_path={json_path}")
            logger.debug(f"Resolved out_path={out_path}")

            build_docx_from_appendix_json(json_path, out_path, week, logger)
        else:
            md_path = _discover_week_md(input_root, week)
            out_path = _output_docx_path(output_root, md_path)

            logger.debug(f"Resolved md_path={md_path}")
            logger.debug(f"Resolved out_path={out_path}")

            build_docx_from_md(md_path, out_path, logger)

    logger.info("Done.")


if __name__ == "__main__":
    main()