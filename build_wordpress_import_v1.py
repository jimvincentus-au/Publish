#!/usr/bin/env python3

import argparse
import json
from pathlib import Path
import ast
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
import logging
import re
import html as html_lib

import pandas as pd


DEFAULT_IMAGE_BASE_URL = "https://thedemocracyclock.com/wp-content/uploads"
DEFAULT_SITE_BASE_URL = "https://thedemocracyclock.com"

PUBLISH_ROOT = Path("/Volumes/PRINTIFY24/Democracy Clock Automation/Publish")
WP_OUTPUT_ROOT = PUBLISH_ROOT / "Output" / "Wordpress"
WP_INPUT_ROOT = PUBLISH_ROOT / "Input" / "Wordpress"
PUBLISH_LOGS_DIR = PUBLISH_ROOT / "Logs"

# Column schema for WP All Import (matches the WP All Export layout used for --update)
COLUMN_ORDER: List[str] = [
    "id",
    "Title",
    "Content",
    "Excerpt",
    "Date",
    "Post Type",
    "Permalink",
    "Image URL",
    "Image Filename",
    "Image Path",
    "Image ID",
    "Image Title",
    "Image Caption",
    "Image Description",
    "Image Alt Text",
    "Image Featured",
    "Categories",
    "Tags",
    "Status",
    "Author ID",
    "Author Username",
    "Author Email",
    "Author First Name",
    "Author Last Name",
    "Slug",
    "Format",
    "Template",
    "Parent",
    "Parent Slug",
    "Order",
    "Comment Status",
    "Ping Status",
    "Post Modified Date",
    # ACF / custom fields (single post per week, appendix-parallel fields)
    "Is Current",
    "Week Number",
    "Week Start Date",
    "Week End Date",
    "Week Label",
    "Week Start Minutes",
    "Week Start Time",
    "Week End Minutes",
    "Week End Time",
    "Week Movement Minutes",
    "Week Start Time Display",
    "Week End Time Display",
    "Week Status Tier Label",
    "Week Status Tier Description",
    "Week Status Summary",
    "Clock Time Reference",
    "Week Delta Summary",
    "Clock Moved",
    "Moral Floor Summary",
    "Net Interpretation",
    "Traits Moved Count",
    "Event Count Total",
    "Sources",
    "Appendix Title",
    "Appendix Content",
    "Appendix Excerpt",
    "Appendix Featured Image URL",
    "Appendix Featured Image Filename",
]

# (Removed: CLOCK_STATUS_COLUMN_ORDER)

def _extract_event_count_total(metadata: Dict[str, Any]) -> Any:
    """
    Return total event count using tolerant key lookup.
    (We have seen schema drift across generators.)
    """
    candidates = [
        "event_count_total",
        "events_count_total",
        "event_count",
        "events_total",
        "total_events",
        "event_total",
    ]
    for k in candidates:
        v = metadata.get(k)
        if isinstance(v, (int, float)) and v != "":
            return int(v)
        if isinstance(v, str) and v.strip():
            try:
                return int(float(v.strip()))
            except Exception:
                pass

    # nested fallbacks if present
    v2 = get_nested(metadata, ["counts", "event_count_total"], "")
    if v2 not in ("", None):
        try:
            return int(float(v2))
        except Exception:
            pass

    return ""

def load_event_count_total(week_dir: Path, week: int, logger: Optional[logging.Logger] = None) -> Any:
    def _log(msg: str) -> None:
        if logger is not None:
            logger.debug(msg)

    wk2 = f"{week:02d}"
    wk = str(int(week))

    candidates = [
        week_dir / f"events_appendix_week{wk2}.json",
        week_dir / f"events_appendix_week{wk}.json",
        week_dir / f"events_week{wk2}.json",
        week_dir / f"events_week{wk}.json",
    ]

    # Also allow recursive discovery (in case the file lands in a subfolder later)
    recursive = sorted({
        *week_dir.rglob(f"*events*appendix*week{wk2}*.json"),
        *week_dir.rglob(f"*events*appendix*week{wk}*.json"),
    })

    for p in [*candidates, *recursive]:
        if not p.exists() or not p.is_file():
            continue

        _log(f"Event Count: trying {p}")
        try:
            data = load_json(p)

            # Schema A: {"events": [ ... ]}
            if isinstance(data, dict) and isinstance(data.get("events"), list):
                return len(data["events"])

            # Schema B: {"categories": {"A": {"events":[...]}, ...}} or similar
            if isinstance(data, dict):
                total = 0

                def walk(obj: Any) -> None:
                    nonlocal total
                    if isinstance(obj, dict):
                        if "events" in obj and isinstance(obj["events"], list):
                            total += len(obj["events"])
                        for v in obj.values():
                            walk(v)
                    elif isinstance(obj, list):
                        for v in obj:
                            walk(v)

                walk(data)
                if total:
                    return total

            # Schema C: a bare list of events
            if isinstance(data, list):
                return len(data)

        except Exception as e:
            _log(f"Event Count: failed reading {p}: {e}")

    # If it’s missing, do NOT hard-fail import; just leave blank and log.
    _log(f"Event Count: no events appendix JSON found under {week_dir} for week={week}")
    return ""


def _parse_iso_date(date_str: str) -> Optional[datetime]:
    """Parse an ISO date (YYYY-MM-DD) into a datetime, or return None."""
    if not date_str or not isinstance(date_str, str):
        return None
    s = date_str.strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except Exception:
        return None


def _strip_first_h1(html_text: str) -> str:
    """
    Remove the first <h1>...</h1> from HTML content.

    The <h1> is used to derive the WordPress Post Title and must not
    appear in the rendered content to avoid duplication and layout drift.
    """
    if not html_text or "<h1" not in html_text.lower():
        return html_text

    # Remove the first <h1>...</h1> only
    return re.sub(
        r"<h1[^>]*>.*?</h1>",
        "",
        html_text,
        count=1,
        flags=re.IGNORECASE | re.DOTALL,
    )

def _format_week_label(week_number: int, start_date: str, end_date: str) -> str:
    """Return the canonical Week Label date-range string.

    Locked format rules:
      - Same month/year:     January 25 - 31, 2025
      - Cross-month, same yr: June 30 - July 6, 2025
      - Cross-year:          December 27, 2025 - January 2, 2026

    Notes:
      - Uses single spaces around hyphen: ' - '
      - Full month names
      - No leading zeros on day
    """
    ds = _parse_iso_date(start_date)
    de = _parse_iso_date(end_date)

    # Fallback if dates are missing/unparseable
    if ds is None or de is None:
        return f"Week {week_number:02d}" if week_number is not None else ""

    same_year = ds.year == de.year
    same_month = same_year and (ds.month == de.month)

    if same_month:
        # January 25 - 31, 2025
        return f"{ds.strftime('%B')} {ds.day} - {de.day}, {ds.year}"

    if same_year:
        # June 30 - July 6, 2025
        return f"{ds.strftime('%B')} {ds.day} - {de.strftime('%B')} {de.day}, {ds.year}"

    # Cross-year: December 27, 2025 - January 2, 2026
    return f"{ds.strftime('%B')} {ds.day}, {ds.year} - {de.strftime('%B')} {de.day}, {de.year}"


def _slugify_for_permalink(text: str) -> str:
    """Slugify a title for use in predictable permalinks."""
    s = str(text or "").strip().lower()
    if not s:
        return ""
    # Replace any non-alphanumeric with hyphen
    s = re.sub(r"[^a-z0-9]+", "-", s)
    # Collapse multiple hyphens and trim
    s = re.sub(r"-+", "-", s).strip("-")
    return s


def _predict_post_url(*, site_base_url: str, week_number: int, kind: str, title: str) -> str:
    """Predict the WordPress URL for a week post.

    Expected patterns (per site convention):
      - https://thedemocracyclock.com/week-49-appendix-files-as-instruments-of-power/
      - https://thedemocracyclock.com/week-49-narrative-<slugified-title>/

    NOTE: This is a best-effort fallback used only when an authoritative Permalink is
    not available from WP All Export.
    """
    base = (site_base_url or "").rstrip("/")
    if not base:
        return ""
    slug_title = _slugify_for_permalink(title)
    if not slug_title:
        return ""
    wk = int(week_number)
    prefix = f"week-{wk}-{'appendix' if kind == 'appendix' else 'narrative'}-"
    return f"{base}/{prefix}{slug_title}/"


# (Removed: build_clock_status_row)


def setup_logger(level: str) -> logging.Logger:
    """Create a simple, file+console logger for the WordPress import builder."""
    PUBLISH_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_path = PUBLISH_LOGS_DIR / "build_wordpress_import_v1.log"

    logger = logging.getLogger("wordpress_import")
    logger.setLevel(logging.DEBUG if level == "debug" else logging.INFO)

    # Avoid duplicate handlers if script is called repeatedly in the same process.
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
    parser = argparse.ArgumentParser(
        description="Build WP All Import XLSX from published Democracy Clock artifacts"
    )

    parser.add_argument("--week", type=int, required=True, help="Starting week number (N)")
    parser.add_argument("--weeks", type=int, required=True, help="Number of weeks to include (K)")
    parser.add_argument(
        "--no-current-week",
        action="store_true",
        help="Do not auto-mark the highest week in the run as current; preserve existing Is Current only in --update mode.",
    )
    parser.add_argument(
        "--status",
        choices=["draft", "publish"],
        default="draft",
        help="Post status for import when WP export does not already specify status (default: draft)",
    )
    parser.add_argument(
        "--author-id",
        type=str,
        default="",
        help="Author ID to set for all rows. Default blank leaves author unchanged in --update mode and uses WordPress default on import.",
    )
    # Removed --only
    parser.add_argument(
        "--image-base-url",
        default=DEFAULT_IMAGE_BASE_URL,
        help="Absolute base URL for featured images (default: site uploads root)",
    )
    parser.add_argument(
        "--site-base-url",
        default=DEFAULT_SITE_BASE_URL,
        help="Base site URL used to predict permalinks when WP export permalinks are unavailable (default: site root).",
    )
    parser.add_argument("--output", help="Output XLSX file path (optional; default is auto-generated)")
    parser.add_argument(
        "--skip-missing",
        action="store_true",
        help="Skip weeks with missing WordPress artifacts instead of failing hard.",
    )

    parser.add_argument(
        "--update",
        action="store_true",
        help="Update-mode: read a WP All Export XLSX and carry forward WP-owned fields (id, permalink, slug, etc.).",
    )
    parser.add_argument(
        "--export",
        help="Path to WP All Export XLSX (optional). If omitted in --update mode, the script will use the only active .xlsx found in Publish/Input/Wordpress/.",
    )
    parser.add_argument("--level", choices=["info", "debug"], default="info", help="Logging level (default: info)")

    return parser.parse_args()


def load_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)
    

def load_weekly_analytic_brief(week_dir: Path, week: int) -> Dict[str, Any]:
    """Load weekly_analytic_brief_weekN.json (tolerant of zero-padding)."""
    candidates = [
        week_dir / f"weekly_analytic_brief_week{week:02d}.json",
        week_dir / f"weekly_analytic_brief_week{week}.json",
    ]
    for p in candidates:
        if p.exists():
            return load_json(p)
    return {}

def load_moral_floor_summary(week_dir: Path, week: int, logger: Optional[logging.Logger] = None) -> str:
    """Load a one-paragraph Moral Floor summary with fallbacks.

    Priority:
      1) metadata_weekNN.json["moral_floor"] (handled by caller)
      2) moral_floor_week*.json (fields: summary | moral_floor | text)
      3) moral_floor_text_week*.txt (first non-empty paragraph)

    The Moral Floor artifacts are not guaranteed to live in the week root; we therefore
    search week_dir recursively.
    """

    # Helper to log only when a logger is provided
    def _log(msg: str) -> None:
        if logger is not None:
            logger.debug(msg)

    # Build tolerant patterns
    wk2 = f"{week:02d}"
    wk = str(int(week))

    # 1) JSON candidates (direct names in root)
    json_candidates = [
        week_dir / f"moral_floor_week{wk2}.json",
        week_dir / f"moral_floor_week{wk}.json",
        week_dir / f"moral_floor_week{wk2}_summary.json",
        week_dir / f"moral_floor_week{wk}_summary.json",
    ]

    # 1b) Recursive JSON discovery
    recursive_json = sorted({
        *week_dir.rglob(f"*moral*floor*week{wk2}*.json"),
        *week_dir.rglob(f"*moral*floor*week{wk}*.json"),
    })

    for p in [*json_candidates, *recursive_json]:
        if p.exists() and p.is_file():
            _log(f"Moral Floor: trying JSON {p}")
            try:
                data = load_json(p)
                if isinstance(data, dict):
                    # 1) Direct, top-level string keys
                    for k in ("summary", "moral_floor", "text", "value"):
                        v = data.get(k)
                        if isinstance(v, str) and v.strip():
                            _log(f"Moral Floor: loaded from {p} key={k}")
                            return v.strip()

                    # 2) Known nested locations (current moral_floor_weekN.json structure)
                    nested_candidates = [
                        ("weekly_audit", "summary"),
                        ("weekly_audit", "moral_floor"),
                        ("weekly_audit", "text"),
                    ]
                    for a, b in nested_candidates:
                        v = data.get(a, {})
                        if isinstance(v, dict):
                            s = v.get(b)
                            if isinstance(s, str) and s.strip():
                                _log(f"Moral Floor: loaded from {p} key={a}.{b}")
                                return s.strip()
            except Exception as e:
                _log(f"Moral Floor: JSON read failed for {p}: {e}")

    # 2) TXT candidates (direct names in root)
    txt_candidates = [
        week_dir / f"moral_floor_text_week{wk2}.txt",
        week_dir / f"moral_floor_text_week{wk}.txt",
        week_dir / f"moral_floor_week{wk2}.txt",
        week_dir / f"moral_floor_week{wk}.txt",
    ]

    # 2b) Recursive TXT discovery
    recursive_txt = sorted({
        *week_dir.rglob(f"*moral*floor*week{wk2}*.txt"),
        *week_dir.rglob(f"*moral*floor*week{wk}*.txt"),
    })

    for p in [*txt_candidates, *recursive_txt]:
        if p.exists() and p.is_file():
            _log(f"Moral Floor: trying TXT {p}")
            try:
                raw = p.read_text(encoding="utf-8").strip()
                if not raw:
                    continue
                # First non-empty paragraph
                parts = [x.strip() for x in re.split(r"\n\s*\n", raw) if x.strip()]
                out = parts[0] if parts else raw
                if out.strip():
                    _log(f"Moral Floor: loaded from {p} (first paragraph)")
                    return out.strip()
            except Exception as e:
                _log(f"Moral Floor: TXT read failed for {p}: {e}")

    _log(f"Moral Floor: no artifact found under {week_dir} for week={week}")
    return ""

def get_nested(d: Dict[str, Any], path: List[str], default: Any = "") -> Any:
    cur: Any = d
    for key in path:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


# --- Backup logic helpers ---
BACKUP_TOKEN = "_bkup_"


def is_backup_xlsx(path: Path) -> bool:
    """
    Backup files are named like:
      <stem>_bkup_YYYYMMDD_HHMMSS.xlsx
    These are ignored for default input selection.
    """
    name = path.name.lower()
    return name.endswith(".xlsx") and BACKUP_TOKEN in name


def make_backup_name(active_path: Path) -> Path:
    """Create a timestamped backup name alongside the active XLSX."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = active_path.stem
    return active_path.with_name(f"{stem}{BACKUP_TOKEN}{ts}.xlsx")


def _find_single_xlsx(folder: Path) -> Path:
    """
    Return the single *active* XLSX in the folder.
    Backup files containing BACKUP_TOKEN are ignored.
    """
    folder.mkdir(parents=True, exist_ok=True)

    all_xlsx = sorted([p for p in folder.glob("*.xlsx") if p.is_file()])
    active = [p for p in all_xlsx if not is_backup_xlsx(p)]

    if len(active) == 0:
        raise FileNotFoundError(
            f"No active .xlsx found in {folder}. "
            f"(Backup files like '*{BACKUP_TOKEN}YYYYMMDD_HHMMSS.xlsx' are ignored.)"
        )
    if len(active) > 1:
        names = ", ".join(p.name for p in active)
        raise RuntimeError(
            f"Multiple active .xlsx files found in {folder}; please specify --export. Found: {names}"
        )

    return active[0]


def load_export_xlsx(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def export_index(df: pd.DataFrame) -> Dict[int, Dict[str, Any]]:
    """Index export rows by Week Number (int). For single post per week model."""
    out: Dict[int, Dict[str, Any]] = {}
    for _, r in df.iterrows():
        row = {k: ("" if pd.isna(v) else v) for k, v in r.to_dict().items()}
        try:
            wk = int(row.get("Week Number", "") or row.get("Week Number ", ""))
        except Exception:
            continue
        out[wk] = row
    return out


def _normalize_sources(value: object) -> str:
    """Return a comma-separated sources string with no brackets."""
    if value is None:
        return ""

    if isinstance(value, list):
        parts = [str(x).strip() for x in value if str(x).strip()]
        return ", ".join(parts)

    if isinstance(value, str):
        s = value.strip()
        if not s:
            return ""

        if s.startswith("[") and s.endswith("]"):
            # JSON list
            try:
                parsed = json.loads(s)
                if isinstance(parsed, list):
                    parts = [str(x).strip() for x in parsed if str(x).strip()]
                    return ", ".join(parts)
            except Exception:
                pass

            # Python literal list
            try:
                parsed = ast.literal_eval(s)
                if isinstance(parsed, list):
                    parts = [str(x).strip() for x in parsed if str(x).strip()]
                    return ", ".join(parts)
            except Exception:
                pass

        s = s.replace(";", ",")
        s = " ".join(s.split())
        return s

    return str(value).strip()


def _minutes_after_noon_to_time_str(minutes_after_noon: float | int) -> str:
    """Convert minutes-after-noon to 'h:mm a.m./p.m.' (e.g., '11:57 p.m.')."""
    try:
        m = int(round(float(minutes_after_noon)))
    except Exception:
        return ""

    base = datetime(2000, 1, 1, 12, 0, 0)  # noon anchor
    dt = base + timedelta(minutes=m)

    hour = dt.hour
    minute = dt.minute
    suffix = "a.m." if hour < 12 else "p.m."
    hour12 = hour % 12
    if hour12 == 0:
        hour12 = 12

    return f"{hour12}:{minute:02d} {suffix}"

# --- Clock status tier helpers ---

# Tier config is part of the Publish pipeline and must live alongside this script
# (Publish is intentionally self-contained and has no external Config dependency).

PUBLISH_ROOT_PATH = Path(__file__).resolve().parent
CLOCK_TIER_CONFIG_PATH = PUBLISH_ROOT_PATH / "clock_status_tiers_v1.json"

def _load_clock_tiers() -> List[Dict[str, Any]]:
    """Load clock status tier bins from Publish-local config.

    This file is REQUIRED for publish; absence is a hard error to prevent
    silent drift or untracked changes.
    """
    if not CLOCK_TIER_CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Missing required clock tier config: {CLOCK_TIER_CONFIG_PATH}"
        )

    with CLOCK_TIER_CONFIG_PATH.open("r", encoding="utf-8") as f:
        data = json.load(f)

    # Support multiple schema variants for backward compatibility only.
    # New tier definitions MUST follow the canonical schema and be explicitly versioned.
    #   v1 canonical: {"bins": [...]}
    #   legacy alt:   {"tiers": [...]}
    #   alt:          {"clock_tiers": [...]}
    #   bare list:    [...]
    if isinstance(data, list):
        bins = data
    elif isinstance(data, dict):
        bins = data.get("bins")
        if bins is None:
            bins = data.get("tiers")
        if bins is None:
            bins = data.get("clock_tiers")
    else:
        bins = None

    if not isinstance(bins, list) or not bins:
        keys = sorted(list(data.keys())) if isinstance(data, dict) else []
        raise ValueError(
            "Invalid clock tier config (no tier list found). "
            f"Expected one of: bins/tiers/clock_tiers or a bare list. "
            f"Path={CLOCK_TIER_CONFIG_PATH}; keys={keys}"
        )

    # Basic schema validation (fail fast if required fields are missing)
    required = {"min_minutes", "max_minutes", "label", "description"}
    for i, b in enumerate(bins, start=1):
        if not isinstance(b, dict):
            raise ValueError(
                f"Invalid clock tier config: entry #{i} is not an object/dict (Path={CLOCK_TIER_CONFIG_PATH})"
            )
        missing = required - set(b.keys())
        if missing:
            raise ValueError(
                f"Invalid clock tier config: entry #{i} missing keys {sorted(missing)} (Path={CLOCK_TIER_CONFIG_PATH})"
            )

    return bins

CLOCK_TIERS: List[Dict[str, Any]] = _load_clock_tiers()

def _tier_for_minutes(end_minutes: Optional[float]) -> Tuple[str, str]:
    """
    Map end-of-week minutes-after-noon to a (tier_label, tier_description).
    Returns empty strings if minutes are missing or out of range.
    """
    if end_minutes is None:
        return "", ""

    try:
        m = int(round(float(end_minutes)))
    except Exception:
        return "", ""

    for b in CLOCK_TIERS:
        try:
            if b["min_minutes"] <= m <= b["max_minutes"]:
                return b.get("label", ""), b.get("description", "")
        except Exception:
            continue

    return "", ""


def _extract_carry_forward_fields(carry_forward: Dict[str, Any]) -> Dict[str, Any]:
    """Extract week start/end date + minutes from carry_forward JSON with tolerant keying."""
    out: Dict[str, Any] = {}

    def _get_date(*keys: str) -> str:
        for k in keys:
            v = carry_forward.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return ""

    week_start_date = _get_date("week_start_date", "start_date", "WeekStartDate", "StartDate")
    week_end_date = _get_date("week_end_date", "end_date", "WeekEndDate", "EndDate")

    if (not week_start_date or not week_end_date) and isinstance(carry_forward.get("window"), dict):
        w = carry_forward.get("window")
        if isinstance(w, dict):
            if not week_start_date:
                s = w.get("start")
                if isinstance(s, str) and s.strip():
                    week_start_date = s.strip()
            if not week_end_date:
                e = w.get("end")
                if isinstance(e, str) and e.strip():
                    week_end_date = e.strip()

    def _get_num(*keys: str) -> Optional[float]:
        for k in keys:
            v = carry_forward.get(k)
            if isinstance(v, (int, float)):
                return float(v)
            if isinstance(v, str) and v.strip():
                try:
                    return float(v.strip())
                except Exception:
                    pass
        return None

    start_minutes = _get_num("clock_before_minutes", "week_start_minutes", "start_minutes", "ClockBeforeMinutes")
    end_minutes = _get_num("clock_after_minutes", "week_end_minutes", "end_minutes", "ClockAfterMinutes")

    out["week_start_date"] = week_start_date
    out["week_end_date"] = week_end_date
    out["week_start_minutes"] = start_minutes
    out["week_end_minutes"] = end_minutes
    out["week_start_time"] = _minutes_after_noon_to_time_str(start_minutes) if start_minutes is not None else ""
    out["week_end_time"] = _minutes_after_noon_to_time_str(end_minutes) if end_minutes is not None else ""

    return out



def _derive_wp_title(*, week_number: int, kind: str, metadata: Dict[str, Any], html_text: str) -> str:
    """Derive the canonical WP post Title.

    Rule: Title should match the HTML title displayed in the post body (first <h1>),
    falling back to a deterministic Week-prefixed title.
    """
    m = re.search(r"<h1[^>]*>(.*?)</h1>", html_text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        h1 = m.group(1)
        h1 = re.sub(r"<[^>]+>", "", h1)
        h1 = html_lib.unescape(h1)
        h1 = " ".join(h1.split()).strip()
        if h1:
            return h1

    base = str(metadata.get("title", "") or "").strip()
    return f"Week {week_number}: {base}" if base else f"Week {week_number}"


# Epigraph strip helper for WordPress import
def _strip_epigraph_blocks(html_text: str) -> str:
    """Remove epigraph blocks from generated post HTML.

    WordPress posts should not include epigraphs. In our pipeline, blockquotes
    are used for epigraphs (not for body quoting), so we remove all blockquotes
    EXCEPT those explicitly marked as the subtitle blockquote we generate.

    Removal rules:
      - Remove any element whose class/id includes 'epigraph'.
      - Remove any leading epigraph-style <blockquote> blocks (and optional nearby
        attribution paragraphs) after the title.
      - As a final hard guard, remove ALL remaining <blockquote>...</blockquote>
        blocks unless they contain class 'dc-subtitle'.

    If no matching blocks exist, returns input unchanged.
    """
    if not html_text or "<" not in html_text:
        return html_text

    out = html_text

    # 1) Remove any explicit epigraph containers.
    out = re.sub(
        r"<(?P<tag>div|p|blockquote)[^>]*(?:class|id)=[\"'][^\"']*epigraph[^\"']*[\"'][^>]*>.*?</(?P=tag)>",
        "",
        out,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # 2) Remove any blockquote(s) directly after the first H1 (common structure).
    #    We do this before stripping H1 in some call paths; tolerant either way.
    out = re.sub(
        r"(</h1>\s*)(?:<blockquote[^>]*>.*?</blockquote>\s*)+(?:<p[^>]*>\s*[—\-–].*?</p>\s*)*",
        r"\1",
        out,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # 3) Remove any remaining blockquotes unless they are our subtitle blockquote.
    #    NOTE: subtitle blockquotes are emitted as: <blockquote class=\"dc-subtitle\">...</blockquote>
    out = re.sub(
        r"<blockquote(?![^>]*dc-subtitle)[^>]*>.*?</blockquote>",
        "",
        out,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # Normalize excessive whitespace that can result from removals.
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out


# Subtitle paragraph to blockquote (WordPress import)
def _format_first_subtitle_paragraph_as_blockquote(
    html_text: str,
    *,
    logger: Optional[logging.Logger] = None,
    week_number: Optional[int] = None,
    kind: str = "",
) -> str:
    """Convert the first short subtitle paragraph into an italic blockquote.

    Heuristic (stable for our generators):
      - Ignore leading whitespace.
      - If the first paragraph is an image wrapper (<p><img ...></p>), skip it.
      - Take the next non-empty <p>...</p> as the subtitle if its plain-text length
        is reasonably short (<= 280 chars).

    We emit:
      <blockquote class="dc-subtitle"><em>SUBTITLE</em></blockquote>

    If we cannot confidently identify a subtitle paragraph, we return html_text unchanged.
    """
    if not html_text or "<p" not in html_text.lower():
        return html_text

    def _log(msg: str) -> None:
        if logger is not None:
            logger.debug(msg)

    # Find all paragraph blocks in order.
    paras = list(re.finditer(r"<p[^>]*>.*?</p>", html_text, flags=re.IGNORECASE | re.DOTALL))
    if not paras:
        return html_text

    # Helper: plain text length for a paragraph block
    def _plain_len(p_html: str) -> int:
        s = re.sub(r"<[^>]+>", " ", p_html)
        s = html_lib.unescape(s)
        s = " ".join(s.split()).strip()
        return len(s)

    # Identify candidate subtitle paragraph index
    idx = 0
    # Skip first paragraph if it contains an image
    if idx < len(paras):
        p0 = paras[idx].group(0)
        if re.search(r"<img\b", p0, flags=re.IGNORECASE):
            _log(f"Week {week_number:02d}: subtitle scan: skipping leading image paragraph" if week_number is not None else "Subtitle scan: skipping leading image paragraph")
            idx += 1

    # Find first non-empty paragraph after optional image
    while idx < len(paras):
        p_html = paras[idx].group(0)
        plen = _plain_len(p_html)
        if plen == 0:
            idx += 1
            continue

        # Only treat as subtitle if short enough
        if plen <= 280:
            # Extract inner html
            m = re.match(r"<p[^>]*>(?P<inner>.*)</p>", p_html, flags=re.IGNORECASE | re.DOTALL)
            inner = (m.group("inner") if m else "").strip()

            # If already emphasized, keep inner; otherwise wrap with <em>
            if re.search(r"<em\b", inner, flags=re.IGNORECASE):
                subtitle_inner = inner
            else:
                subtitle_inner = f"<em>{inner}</em>"

            bq = f"<blockquote class=\"dc-subtitle\">{subtitle_inner}</blockquote>"

            start, end = paras[idx].span()
            new_html = html_text[:start] + bq + html_text[end:]

            _log(
                f"Week {week_number:02d}: converted first subtitle paragraph to blockquote (len={plen}) kind={kind}" if week_number is not None
                else f"Converted first subtitle paragraph to blockquote (len={plen}) kind={kind}"
            )
            return new_html

        _log(
            f"Week {week_number:02d}: subtitle scan: first paragraph too long (len={plen}); leaving as-is" if week_number is not None
            else f"Subtitle scan: first paragraph too long (len={plen}); leaving as-is"
        )
        return html_text

    return html_text


def build_week_post_row(
    *,
    week_number: int,
    metadata: Dict[str, Any],
    carry_forward: Dict[str, Any],
    weekly_analytic_brief: Optional[Dict[str, Any]] = None,
    narrative_html_path: Path,
    appendix_html_path: Path,
    post_status: str,
    image_base_url: str,
    post_date: datetime,
    is_current: bool,
    author_id_override: str = "",
    logger: Optional[logging.Logger] = None,
    existing: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    categories = "Week"
    existing = existing or {}
    def _carry(col: str, default: str = "") -> Any:
        v = existing.get(col, default)
        if pd.isna(v):
            return default
        return v
    # Carry-forward WP-owned fields (authoritative from WP export)
    wp_id = str(_carry("id", ""))
    wp_permalink = str(_carry("Permalink", ""))
    wp_date = str(_carry("Date", ""))
    wp_status_existing = str(_carry("Status", ""))
    wp_author_id_existing = str(_carry("Author ID", ""))
    wp_author_id = author_id_override.strip() if str(author_id_override or "").strip() else wp_author_id_existing
    wp_author_username = str(_carry("Author Username", ""))
    wp_author_email = str(_carry("Author Email", ""))
    wp_author_first = str(_carry("Author First Name", ""))
    wp_author_last = str(_carry("Author Last Name", ""))
    wp_slug = str(_carry("Slug", ""))
    wp_format = str(_carry("Format", ""))
    wp_template = str(_carry("Template", ""))
    wp_parent = str(_carry("Parent", ""))
    wp_parent_slug = str(_carry("Parent Slug", ""))
    wp_order = str(_carry("Order", ""))
    wp_comment_status = str(_carry("Comment Status", ""))
    wp_ping_status = str(_carry("Ping Status", ""))
    wp_post_modified = str(_carry("Post Modified Date", ""))
    # Status: always "publish"
    final_status = "publish"
    # Featured image: always narrative featured image
    featured_image_filename = f"week{week_number:02d}-narrative-featured.jpg"
    base = (image_base_url or "").rstrip("/")
    final_image_url = f"{base}/{featured_image_filename}" if base else ""
    # --- Carry-forward-derived week fields (source of truth) ---
    cf = _extract_carry_forward_fields(carry_forward)
    acf_week_number = week_number
    acf_week_start_date = str(cf.get("week_start_date") or "")
    acf_week_end_date = str(cf.get("week_end_date") or "")
    acf_week_start_minutes = cf.get("week_start_minutes")
    acf_week_end_minutes = cf.get("week_end_minutes")
    acf_week_start_time = str(cf.get("week_start_time") or "")
    acf_week_end_time = str(cf.get("week_end_time") or "")
    # Week label
    acf_week_label = _format_week_label(week_number, acf_week_start_date, acf_week_end_date)
    # Movement minutes (prefer carry-forward start/end; fallback to carry_forward clock_delta_minutes; final fallback to metadata)
    delta_minutes: Any = ""
    try:
        if acf_week_start_minutes is not None and acf_week_end_minutes is not None:
            delta_minutes = int(round(float(acf_week_end_minutes) - float(acf_week_start_minutes)))
        else:
            dm = carry_forward.get("clock_delta_minutes", "")
            if dm != "" and dm is not None:
                delta_minutes = int(round(float(dm)))
    except Exception:
        delta_minutes = ""
    if delta_minutes == "" or delta_minutes is None:
        dm2 = metadata.get("clock_movement_minutes", "")
        try:
            if dm2 != "" and dm2 is not None:
                delta_minutes = int(round(float(dm2)))
        except Exception:
            delta_minutes = ""
    # Tier from end-of-week minutes
    tier_label, tier_description = _tier_for_minutes(acf_week_end_minutes)
    # Drift-prevention guard: never allow an unresolved tier to silently import.
    if not tier_label:
        raise ValueError(
            f"Week {week_number:02d}: No tier resolved for end_minutes={acf_week_end_minutes}. "
            f"Check tier definitions in {CLOCK_TIER_CONFIG_PATH}"
        )
    # Status summary: priority: weekly analytic brief → metadata delta_summary.net_interpretation → short synopsis
    week_status_summary = ""
    wab = weekly_analytic_brief or {}
    if isinstance(wab, dict):
        week_status_summary = str(wab.get("summary", "") or "").strip()
    if not week_status_summary:
        ds2 = metadata.get("delta_summary")
        if isinstance(ds2, dict) and "net_interpretation" in ds2:
            week_status_summary = str(ds2.get("net_interpretation", "") or "").strip()
    if not week_status_summary:
        week_status_summary = str(metadata.get("short_synopsis", "") or "").strip()
    # Moral Floor Summary: prefer metadata, then week-dir artifacts
    acf_moral_floor_summary = str(metadata.get("moral_floor", "") or "").strip()
    if logger is not None:
        logger.debug(f"Week {week_number:02d}: metadata.moral_floor length={len(acf_moral_floor_summary)}")
    if not acf_moral_floor_summary:
        week_dir = narrative_html_path.parent
        acf_moral_floor_summary = load_moral_floor_summary(week_dir, week_number, logger=logger)
    if logger is not None:
        logger.debug(
            f"Week {week_number:02d}: resolved Moral Floor Summary length={len(acf_moral_floor_summary)}; "
            f"week_dir={narrative_html_path.parent}"
        )
    # Net Interpretation: prefer metadata.delta_summary.net_interpretation else week_status_summary
    acf_net_interpretation = ""
    dsni = metadata.get("delta_summary")
    if isinstance(dsni, dict) and dsni.get("net_interpretation"):
        acf_net_interpretation = str(dsni.get("net_interpretation") or "").strip()
    else:
        acf_net_interpretation = week_status_summary
    # Traits moved count
    acf_traits_moved_count: Any = ""
    tm = metadata.get("traits_moved")
    if isinstance(tm, dict) and tm.get("value") not in ("", None):
        acf_traits_moved_count = tm.get("value")
    else:
        ds_tm = get_nested(metadata, ["delta_summary", "traits_moved", "value"], "")
        if ds_tm not in ("", None):
            acf_traits_moved_count = ds_tm
    # Clock moved boolean
    clock_moved_flag: Any = ""
    try:
        if delta_minutes == "":
            clock_moved_flag = ""
        else:
            clock_moved_flag = bool(int(delta_minutes) != 0)
    except Exception:
        clock_moved_flag = ""
    acf_clock_movement_minutes = delta_minutes if delta_minutes != "" else metadata.get("clock_movement_minutes", "")
    acf_clock_time_reference = metadata.get("clock_time_reference", "")
    # Week delta summary: notes only from delta_summary.*.note
    ds = metadata.get("delta_summary", "")
    if isinstance(ds, dict):
        notes: List[str] = []
        gtp_note = get_nested(ds, ["grand_total_points", "note"], "")
        tm_note = get_nested(ds, ["traits_moved", "note"], "")
        if gtp_note:
            notes.append(str(gtp_note).strip())
        if tm_note:
            notes.append(str(tm_note).strip())
        acf_week_delta_summary = "\n".join(notes).strip()
    else:
        acf_week_delta_summary = str(metadata.get("week_delta_summary", "") or "")
    try:
        moved_val = float(acf_clock_movement_minutes) if acf_clock_movement_minutes != "" else None
        acf_clock_moved = bool(moved_val and moved_val != 0)
    except Exception:
        acf_clock_moved = ""
    acf_event_count_total = load_event_count_total(narrative_html_path.parent, week_number, logger=logger)
    if acf_event_count_total == "":
        # optional fallback to metadata if you want *something*
        acf_event_count_total = _extract_event_count_total(metadata)
    acf_sources_raw = metadata.get("quote_sources", metadata.get("sources", ""))
    acf_sources = _normalize_sources(acf_sources_raw)
    # Appendix excerpt: falls back to short synopsis if no explicit appendix summary exists
    acf_appendix_excerpt = metadata.get("appendix_summary", metadata.get("short_synopsis", ""))
    # Read narrative HTML
    narrative_html = narrative_html_path.read_text(encoding="utf-8")
    wp_title = _derive_wp_title(
        week_number=week_number,
        kind="narrative",
        metadata=metadata,
        html_text=narrative_html,
    )
    # Sanitize narrative HTML
    if logger is not None:
        logger.debug(
            f"Week {week_number:02d}: pre-sanitize narrative HTML contains blockquote={('<blockquote' in narrative_html.lower())}"
        )
    nar_html = _strip_epigraph_blocks(narrative_html)
    nar_html = _strip_first_h1(nar_html)
    before_len = len(nar_html or "")
    nar_html = _format_first_subtitle_paragraph_as_blockquote(
        nar_html,
        logger=logger,
        week_number=week_number,
        kind="narrative",
    )
    after_len = len(nar_html or "")
    if logger is not None:
        logger.debug(
            f"Week {week_number:02d}: subtitle blockquote pass complete (changed={before_len != after_len}; before={before_len}; after={after_len})"
        )
    # Read appendix HTML and sanitize for ACF
    appendix_html = appendix_html_path.read_text(encoding="utf-8")
    acf_appendix_title = _derive_wp_title(
        week_number=week_number,
        kind="appendix",
        metadata=metadata,
        html_text=appendix_html,
    )
    app_html = _strip_epigraph_blocks(appendix_html)
    app_html = _strip_first_h1(app_html)
    # Do NOT run subtitle conversion for appendix
    # Appendix featured image
    appendix_featured_image_filename = f"week{week_number:02d}-appendix-featured.jpg"
    appendix_featured_image_url = f"{base}/{appendix_featured_image_filename}" if base else ""
    # Display time fields
    week_start_time_display = _minutes_after_noon_to_time_str(acf_week_start_minutes)
    week_end_time_display = _minutes_after_noon_to_time_str(acf_week_end_minutes)
    # Is Current
    is_current_field = "1" if is_current else ""
    return {
        "id": wp_id,
        "Title": wp_title,
        "Content": nar_html,
        "Excerpt": metadata.get("short_synopsis", ""),
        "Date": wp_date if wp_date else post_date.strftime("%Y-%m-%d %H:%M:%S"),
        "Post Type": "post",
        "Permalink": wp_permalink,
        "Image URL": final_image_url,
        "Image Filename": featured_image_filename,
        "Image Path": "",
        "Image ID": "",
        "Image Title": "",
        "Image Caption": "",
        "Image Description": "",
        "Image Alt Text": "",
        "Image Featured": "1" if final_image_url else "",
        "Categories": categories,
        "Tags": "",
        "Status": final_status,
        "Author ID": wp_author_id,
        "Author Username": wp_author_username,
        "Author Email": wp_author_email,
        "Author First Name": wp_author_first,
        "Author Last Name": wp_author_last,
        "Slug": wp_slug,
        "Format": wp_format,
        "Template": wp_template,
        "Parent": wp_parent,
        "Parent Slug": wp_parent_slug,
        "Order": wp_order,
        "Comment Status": wp_comment_status,
        "Ping Status": wp_ping_status,
        "Post Modified Date": wp_post_modified,
        # --- ACF fields ---
        "Is Current": is_current_field,
        "Week Number": acf_week_number,
        "Week Start Date": acf_week_start_date,
        "Week End Date": acf_week_end_date,
        "Week Label": acf_week_label,
        "Week Start Minutes": acf_week_start_minutes,
        "Week Start Time": acf_week_start_time,
        "Week End Minutes": acf_week_end_minutes,
        "Week End Time": acf_week_end_time,
        "Week Movement Minutes": delta_minutes,
        "Week Start Time Display": week_start_time_display,
        "Week End Time Display": week_end_time_display,
        "Week Status Tier Label": tier_label,
        "Week Status Tier Description": tier_description,
        "Week Status Summary": week_status_summary,
        "Clock Time Reference": acf_clock_time_reference,
        "Week Delta Summary": acf_week_delta_summary,
        "Clock Moved": acf_clock_moved,
        "Moral Floor Summary": acf_moral_floor_summary,
        "Net Interpretation": acf_net_interpretation,
        "Traits Moved Count": acf_traits_moved_count,
        "Event Count Total": acf_event_count_total,
        "Sources": acf_sources,
        "Appendix Title": acf_appendix_title,
        "Appendix Content": app_html,
        "Appendix Excerpt": acf_appendix_excerpt,
        "Appendix Featured Image URL": appendix_featured_image_url,
        "Appendix Featured Image Filename": appendix_featured_image_filename,
    }


def main() -> None:
    args = parse_args()
    logger = setup_logger(args.level)
    logger.info("Starting WordPress import build")
    logger.debug(f"Args: {args}")
    export_rows: Dict[int, Dict[str, Any]] = {}
    active_export_name: str = ""
    if args.update:
        export_path = Path(args.export) if args.export else _find_single_xlsx(WP_INPUT_ROOT)
        logger.info(f"Update mode enabled; loading export XLSX: {export_path}")
        canonical_update = args.export is None
        active_export_name = export_path.name
        if canonical_update:
            backup_path = make_backup_name(export_path)
            logger.info(f"Backing up active export XLSX: {active_export_name} -> {backup_path.name}")
            export_path.rename(backup_path)
            export_source: Path = backup_path
        else:
            export_source = export_path
        logger.info(f"Update export source XLSX: {export_source}")
        export_df = load_export_xlsx(export_source)
        export_rows = export_index(export_df)
        logger.info(f"Loaded {len(export_rows)} indexed export rows")
    rows: List[Dict[str, Any]] = []
    for week in range(args.week, args.week + args.weeks):
        logger.info(f"Processing Week {week:02d}")
        week_dir = WP_OUTPUT_ROOT / f"Week {week:02d}"
        if not week_dir.exists():
            msg = f"Missing WordPress output for Week {week:02d}: {week_dir}"
            if args.skip_missing:
                logger.warning(msg + " (skipped)")
                continue
            raise FileNotFoundError(msg)
        metadata_path = week_dir / f"metadata_week{week:02d}.json"
        if not metadata_path.exists():
            msg = f"Missing metadata JSON: {metadata_path}"
            if args.skip_missing:
                logger.warning(msg + " (skipped)")
                continue
            raise FileNotFoundError(msg)
        metadata = load_json(metadata_path)
        carry_forward_path = week_dir / f"carry_forward_week{week:02d}.json"
        if not carry_forward_path.exists():
            msg = f"Missing carry-forward JSON: {carry_forward_path}"
            if args.skip_missing:
                logger.warning(msg + " (skipped)")
                continue
            raise FileNotFoundError(msg)
        carry_forward = load_json(carry_forward_path)
        weekly_analytic_brief = load_weekly_analytic_brief(week_dir, week)
        post_date = datetime.fromtimestamp(metadata_path.stat().st_mtime)
        narrative_html_path = week_dir / f"week{week:02d}-narrative.html"
        appendix_html_path = week_dir / f"week{week:02d}-appendix.html"
        if not narrative_html_path.exists():
            msg = f"Missing narrative HTML: {narrative_html_path}"
            if args.skip_missing:
                logger.warning(msg + " (skipped)")
                continue
            raise FileNotFoundError(msg)
        if not appendix_html_path.exists():
            msg = f"Missing appendix HTML: {appendix_html_path}"
            if args.skip_missing:
                logger.warning(msg + " (skipped)")
                continue
            raise FileNotFoundError(msg)
        if args.no_current_week:
            # In update mode, preserve whatever WP currently has; otherwise leave blank.
            existing_row = export_rows.get(week, {})
            ex_val = str(existing_row.get("Is Current", "")).strip().lower()
            is_current = ex_val in ("1", "true", "yes")
        else:
            is_current = (week == (args.week + args.weeks - 1))
        rows.append(
            build_week_post_row(
                week_number=week,
                metadata=metadata,
                carry_forward=carry_forward,
                weekly_analytic_brief=weekly_analytic_brief,
                narrative_html_path=narrative_html_path,
                appendix_html_path=appendix_html_path,
                post_status=args.status,
                image_base_url=args.image_base_url.rstrip("/"),
                post_date=post_date,
                is_current=is_current,
                author_id_override=args.author_id,
                logger=logger,
                existing=export_rows.get(week),
            )
        )
    df = pd.DataFrame(rows)
    # Enforce canonical column schema/order; fill any missing columns with blanks.
    for col in COLUMN_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMN_ORDER]
    if args.output:
        output_path = Path(args.output)
    else:
        if args.update and (args.export is None):
            if not active_export_name:
                raise RuntimeError("active_export_name not set in canonical update mode")
            output_path = WP_INPUT_ROOT / active_export_name
        else:
            start = args.week
            end = args.week + args.weeks - 1
            output_path = WP_OUTPUT_ROOT / f"wordpress_import_weeks_{start:02d}_{end:02d}.xlsx"
    logger.info(f"Output path set to: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    logger.info(f"Wrote WordPress import file: {output_path}")


if __name__ == "__main__":
    main()