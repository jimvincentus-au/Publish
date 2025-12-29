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
    # ACF / custom fields
    "Week Number",
    "Week Start Date",
    "Week End Date",
    "Narrative Post Slug",
    "Event Count Total",
    "Sources",
    "Appendix Summary",
    "Appendix Post Slug",
    "Clock Movement Minutes",
    "Week Start Minutes",
    "Week Start Time",
    "Week End Minutes",
    "Week End Time",
    "Clock Time Reference",
    "Week Delta Summary",
    "Clock Moved",
    # Narrative clock-status derived fields (deterministic)
    "Clock Status Type",
    "Week Label",
    "Week Movement Minutes",
    "Week Status Tier Label",
    "Week Status Tier Description",
    "Week Status Summary",
    "Moral Floor Summary",
    "Net Interpretation",
    "Traits Moved Count",
]

# Canonical schema for clock_status CPT (for WP All Import)
CLOCK_STATUS_COLUMN_ORDER: List[str] = [
    "id",
    "Title",
    "Content",
    "Excerpt",
    "Date",
    "Post Type",
    "Permalink",
    "Slug",
    "Status",
    # ACF / custom fields for clock_status
    "Clock Status Type",
    "Week Number",
    "Week Label",
    "Week Start Date",
    "Week End Date",
    "Week Start Minutes",
    "Week End Minutes",
    "Clock Movement Minutes",
    "Week Start Time Display",
    "Week End Time Display",
    "Week Status Tier",
    "Week Status Tier Description",
    "Week Status Summary",
    "Narrative Post URL",
    "Appendix Post URL",
]

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


def build_clock_status_row(
    *,
    week_number: int,
    status_type: str,  # 'weekly' | 'current'
    metadata: Dict[str, Any],
    carry_forward: Dict[str, Any],
    post_status: str,
    narrative_post_slug: str,
    appendix_post_slug: str,
) -> Dict[str, Any]:
    cf = _extract_carry_forward_fields(carry_forward)
    week_start_date = cf.get("week_start_date", "")
    week_end_date = cf.get("week_end_date", "")
    week_start_minutes = cf.get("week_start_minutes")
    week_end_minutes = cf.get("week_end_minutes")
    start_display = _minutes_after_noon_to_time_str(week_start_minutes) if week_start_minutes not in (None, "") else ""
    end_display = _minutes_after_noon_to_time_str(week_end_minutes) if week_end_minutes not in (None, "") else ""
    week_label = _format_week_label(week_number, week_start_date, week_end_date)
    clock_movement_minutes = metadata.get("clock_movement_minutes", "")
    if clock_movement_minutes == "" or clock_movement_minutes is None:
        clock_movement_minutes = carry_forward.get("clock_delta_minutes", "")

    # --- Tier computation ---
    end_minutes: Optional[float] = (
        week_end_minutes if isinstance(week_end_minutes, (int, float)) else None
    )
    end_minutes = max(0.0, min(720.0, end_minutes)) if end_minutes is not None else None
    tier_label, tier_description = _tier_for_minutes(end_minutes)

    # Drift-prevention guard: never allow an unresolved tier to silently import.
    if end_minutes is None or not tier_label:
        raise ValueError(
            f"Week {week_number:02d}: No tier resolved for end_minutes={end_minutes}. "
            f"Check tier definitions in {CLOCK_TIER_CONFIG_PATH}"
        )

    # Week status summary
    ds = metadata.get("delta_summary")
    if isinstance(ds, dict) and "net_interpretation" in ds:
        week_status_summary = ds.get("net_interpretation", "")
    else:
        week_status_summary = metadata.get("short_synopsis", "")

    # Title and Slug
    if status_type == "current":
        title = "Current Clock Status"
        slug = "clock-status-current"
    else:
        title = f"Clock Status — Week {week_number:02d}"
        slug = f"clock-status-week{week_number:02d}"

    # Date
    date_field = f"{week_end_date} 00:00:00" if week_end_date else ""

    return {
        "id": "",
        "Title": title,
        "Content": "",
        "Excerpt": "",
        "Date": date_field,
        "Post Type": "clock_status",
        "Permalink": "",
        "Slug": slug,
        "Status": post_status,
        # Clock status classification (locked, derived)
        "Clock Status Type": tier_label,
        "Week Number": week_number,
        "Week Label": week_label,
        "Week Start Date": week_start_date,
        "Week End Date": week_end_date,
        "Week Start Minutes": week_start_minutes if week_start_minutes not in (None, "") else "",
        "Week End Minutes": week_end_minutes if week_end_minutes not in (None, "") else "",
        "Clock Movement Minutes": clock_movement_minutes if clock_movement_minutes not in (None, "") else "",
        "Week Start Time Display": start_display,
        "Week End Time Display": end_display,
        "Week Status Tier": tier_label,
        "Week Status Tier Description": tier_description,
        "Week Status Summary": week_status_summary,
        "Narrative Post URL": "",
        "Appendix Post URL": "",
    }


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
        "--status",
        choices=["draft", "publish"],
        default="draft",
        help="Post status for import when WP export does not already specify status (default: draft)",
    )
    parser.add_argument("--only", choices=["narrative", "appendix"], help="Import only narrative or appendix posts")
    parser.add_argument(
        "--image-base-url",
        default=DEFAULT_IMAGE_BASE_URL,
        help="Absolute base URL for featured images (default: site uploads root)",
    )
    parser.add_argument("--output", help="Output XLSX file path (optional; default is auto-generated)")

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


def export_index(df: pd.DataFrame) -> Dict[Tuple[int, str], Dict[str, Any]]:
    """Index export rows by (Week Number, kind) where kind is narrative|appendix."""
    out: Dict[Tuple[int, str], Dict[str, Any]] = {}
    for _, r in df.iterrows():
        row = {k: ("" if pd.isna(v) else v) for k, v in r.to_dict().items()}
        try:
            wk = int(row.get("Week Number", "") or row.get("Week Number ", ""))
        except Exception:
            continue

        slug = str(row.get("Slug", "") or "").strip().lower()
        cats = str(row.get("Categories", "") or "").strip().lower()

        kind = ""
        if slug.endswith("-narrative"):
            kind = "narrative"
        elif slug.endswith("-appendix"):
            kind = "appendix"
        elif "clock narrative" in cats:
            kind = "narrative"
        elif "event log" in cats:
            kind = "appendix"

        if kind:
            out[(wk, kind)] = row

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

    We treat epigraphs as *presentation* that may be desirable in book/PDF outputs,
    but not in WordPress posts.

    Removal rules (intentionally tolerant):
      - Remove any element whose class/id includes 'epigraph'.
      - Additionally, remove a leading epigraph-style <blockquote> that appears
        immediately after the first </h1> (common formatting in our generators).

    This is a best-effort sanitizer; if no matching epigraph block exists, the
    input is returned unchanged.
    """
    if not html_text or "<" not in html_text:
        return html_text

    out = html_text

    # 1) Remove any explicit epigraph containers.
    #    Covers: <div class="epigraph">, <p class="epigraph">, <blockquote class="epigraph">,
    #    and id variants like id="epigraph".
    out = re.sub(
        r"<(?P<tag>div|p|blockquote)[^>]*(?:class|id)=[\"'][^\"']*epigraph[^\"']*[\"'][^>]*>.*?</(?P=tag)>",
        "",
        out,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # 2) Remove a leading <blockquote> epigraph directly after the first H1.
    #    This targets the common structure:
    #       <h1>...</h1>\n<blockquote>...</blockquote>\n<p>— Author</p>
    #    We remove the blockquote and an immediate attribution paragraph if present.
    out = re.sub(
        r"(</h1>\s*)<blockquote[^>]*>.*?</blockquote>(\s*<p[^>]*>\s*[—\-–].*?</p>)?",
        r"\1",
        out,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # 3) Remove any blockquote that looks like an epigraph attribution quote.
    #    Our WordPress outputs do not use blockquotes for “real” body quoting;
    #    they are epigraph-only. Heuristic: contains an em dash attribution.
    out = re.sub(
        r"<blockquote[^>]*>.*?(?:—|&mdash;|&#8212;).*?</blockquote>",
        "",
        out,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # Normalize excessive blank lines that can result from removals.
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out


def build_post_row(
    *,
    week_number: int,
    kind: str,  # narrative | appendix
    metadata: Dict[str, Any],
    carry_forward: Dict[str, Any],
    weekly_analytic_brief: Optional[Dict[str, Any]] = None,
    html_path: Path,
    post_status: str,
    image_base_url: str,
    post_date: datetime,
    logger: Optional[logging.Logger] = None,
    existing: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    is_narrative = kind == "narrative"

    # Owned taxonomy fields (always overwritten)
    categories = "Clock Narrative" if is_narrative else "Event Log"

    # Tags: intentionally left blank for now (per current workflow decision)
    tags: List[str] = []

    featured_image_filename = (
        f"week{week_number:02d}-"
        f"{'narrative' if is_narrative else 'appendix'}-featured.jpg"
    )

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
    wp_author_id = str(_carry("Author ID", ""))
    wp_author_id = "2"
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

    # Status: if export has it, keep; otherwise use CLI
    final_status = wp_status_existing if wp_status_existing else post_status

    # Featured image URLs are handled inside WordPress (media library / attachment matching).
    # For WP All Import, we leave Image URL blank.
    final_image_url = ""

    # --- Carry-forward-derived week fields (source of truth) ---
    cf = _extract_carry_forward_fields(carry_forward)

    acf_week_number = week_number
    acf_week_start_date = str(cf.get("week_start_date") or "")
    acf_week_end_date = str(cf.get("week_end_date") or "")
    acf_week_start_minutes = cf.get("week_start_minutes")
    acf_week_end_minutes = cf.get("week_end_minutes")
    acf_week_start_time = str(cf.get("week_start_time") or "")
    acf_week_end_time = str(cf.get("week_end_time") or "")
    # --- Deterministic Narrative clock-status fields (for Clock Narrative posts only) ---

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
    if is_narrative and not tier_label:
        raise ValueError(
            f"Week {week_number:02d}: No tier resolved for end_minutes={acf_week_end_minutes}. "
            f"Check tier definitions in {CLOCK_TIER_CONFIG_PATH}"
        )
    # Narrative-only interpretive summary (never populate for appendix posts)
    # Priority: weekly analytic brief → metadata delta_summary.net_interpretation → short synopsis
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

    # New narrative-only ACF fields
    # Moral Floor Summary (narrative-only): prefer metadata, then week-dir artifacts
    acf_moral_floor_summary = str(metadata.get("moral_floor", "") or "").strip()
    if logger is not None:
        logger.debug(f"Week {week_number:02d}: metadata.moral_floor length={len(acf_moral_floor_summary)}")

    if not acf_moral_floor_summary:
        week_dir = html_path.parent
        acf_moral_floor_summary = load_moral_floor_summary(week_dir, week_number, logger=logger)

    if logger is not None:
        logger.debug(
            f"Week {week_number:02d}: resolved Moral Floor Summary length={len(acf_moral_floor_summary)}; "
            f"week_dir={html_path.parent}"
        )

    # Prefer metadata.delta_summary.net_interpretation if present; fallback to the weekly status summary
    acf_net_interpretation = ""
    dsni = metadata.get("delta_summary")
    if isinstance(dsni, dict) and dsni.get("net_interpretation"):
        acf_net_interpretation = str(dsni.get("net_interpretation") or "").strip()
    else:
        acf_net_interpretation = week_status_summary

    # Traits moved count (narrative-only)
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

    # Clock status type is the tier label (single source of truth)
    acf_clock_status_type = tier_label

    # Slugs (for cross-linking ACF fields only; WP Slug itself must never be touched)
    narrative_slug = f"week{week_number:02d}-narrative"
    appendix_slug = f"week{week_number:02d}-appendix"

    # Clock fields
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

    # Appendix-side fields
    acf_event_count_total = _extract_event_count_total(metadata)
    acf_sources_raw = metadata.get("quote_sources", metadata.get("sources", ""))
    acf_sources = _normalize_sources(acf_sources_raw)
    # Appendix-only summary; falls back to short synopsis if no explicit appendix summary exists
    acf_appendix_summary = metadata.get("appendix_summary", metadata.get("short_synopsis", ""))

    # Read HTML once, derive WP Title from the first H1
    html_text = html_path.read_text(encoding="utf-8")

    # 1) Derive WP title from the original HTML (with H1 still present)
    wp_title = _derive_wp_title(
        week_number=week_number,
        kind=kind,
        metadata=metadata,
        html_text=html_text,
    )

    # 2) Then sanitize for WP display
    html_text = _strip_epigraph_blocks(html_text)
    html_text = _strip_first_h1(html_text)

    return {
        "id": wp_id,
        "Title": wp_title,
        "Content": html_text,
        "Excerpt": metadata.get("short_synopsis", ""),
        "Date": wp_date if wp_date else post_date.strftime("%Y-%m-%d %H:%M:%S"),
        "Post Type": "post",
        "Permalink": wp_permalink,

        # Featured image handling for WP All Import (URL-only workflow)
        "Image URL": final_image_url,  # will be blank by design
        "Image Filename": featured_image_filename,
        "Image Path": "",
        "Image ID": "",
        "Image Title": "",
        "Image Caption": "",
        "Image Description": "",
        "Image Alt Text": "",
        "Image Featured": "",

        # Owned taxonomy fields (always overwritten)
        "Categories": categories,
        "Tags": "",  # intentionally blank for now

        # WP-owned fields (carried)
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
        "Week Number": acf_week_number,
        "Week Start Date": acf_week_start_date,
        "Week End Date": acf_week_end_date,

        # Appendix Post Fields
        "Narrative Post Slug": narrative_slug if (not is_narrative) else "",
        "Event Count Total": acf_event_count_total if (not is_narrative) else "",
        "Sources": acf_sources if (not is_narrative) else "",
        "Appendix Summary": acf_appendix_summary if (not is_narrative) else "",

        # Narrative Post Fields
        "Appendix Post Slug": appendix_slug if is_narrative else "",
        "Clock Movement Minutes": acf_clock_movement_minutes if is_narrative else "",
        "Week Start Minutes": acf_week_start_minutes if is_narrative else "",
        "Week Start Time": acf_week_start_time if is_narrative else "",
        "Week End Minutes": acf_week_end_minutes if is_narrative else "",
        "Week End Time": acf_week_end_time if is_narrative else "",
        "Clock Time Reference": acf_clock_time_reference if is_narrative else "",
        "Week Delta Summary": acf_week_delta_summary if is_narrative else "",
        "Clock Moved": acf_clock_moved if is_narrative else "",

        # Narrative derived clock-status fields
        "Clock Status Type": acf_clock_status_type if is_narrative else "",
        "Week Label": acf_week_label if is_narrative else "",
        "Week Movement Minutes": delta_minutes if is_narrative else "",
        "Week Status Tier Label": tier_label if is_narrative else "",
        "Week Status Tier Description": tier_description if is_narrative else "",
        "Week Status Summary": week_status_summary if is_narrative else "",
        "Moral Floor Summary": acf_moral_floor_summary if is_narrative else "",
        "Net Interpretation": acf_net_interpretation if is_narrative else "",
        "Traits Moved Count": acf_traits_moved_count if is_narrative else "",
    }


def main() -> None:
    args = parse_args()
    logger = setup_logger(args.level)

    logger.info("Starting WordPress import build")
    logger.debug(f"Args: {args}")

    export_rows: Dict[Tuple[int, str], Dict[str, Any]] = {}
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
    # Clock status CPT output lists
    clock_weekly_rows: List[Dict[str, Any]] = []
    clock_last_week_payload: Optional[Tuple[int, Dict[str, Any], Dict[str, Any]]] = None

    for week in range(args.week, args.week + args.weeks):
        logger.info(f"Processing Week {week:02d}")
        week_dir = WP_OUTPUT_ROOT / f"Week {week:02d}"

        if not week_dir.exists():
            raise FileNotFoundError(f"Missing WordPress output for Week {week:02d}: {week_dir}")

        metadata_path = week_dir / f"metadata_week{week:02d}.json"
        if not metadata_path.exists():
            raise FileNotFoundError(f"Missing metadata JSON: {metadata_path}")
        metadata = load_json(metadata_path)

        carry_forward_path = week_dir / f"carry_forward_week{week:02d}.json"
        if not carry_forward_path.exists():
            raise FileNotFoundError(f"Missing carry-forward JSON: {carry_forward_path}")
        carry_forward = load_json(carry_forward_path)
        weekly_analytic_brief = load_weekly_analytic_brief(week_dir, week)

        post_date = datetime.fromtimestamp(metadata_path.stat().st_mtime)

        # Always emit a weekly clock_status row
        clock_weekly_rows.append(
            build_clock_status_row(
                week_number=week,
                status_type="weekly",
                metadata=metadata,
                carry_forward=carry_forward,
                post_status=args.status,
                narrative_post_slug=f"week{week:02d}-narrative",
                appendix_post_slug=f"week{week:02d}-appendix",
            )
        )
        clock_last_week_payload = (week, metadata, carry_forward)

        if args.only in (None, "narrative"):
            html_path = week_dir / f"week{week:02d}-narrative.html"
            if not html_path.exists():
                raise FileNotFoundError(f"Missing narrative HTML: {html_path}")

            rows.append(
                build_post_row(
                    week_number=week,
                    kind="narrative",
                    metadata=metadata,
                    carry_forward=carry_forward,
                    weekly_analytic_brief=weekly_analytic_brief,
                    html_path=html_path,
                    post_status=args.status,
                    image_base_url=args.image_base_url.rstrip("/"),
                    post_date=post_date,
                    logger=logger,
                    existing=export_rows.get((week, "narrative")),
                )
            )

        if args.only in (None, "appendix"):
            html_path = week_dir / f"week{week:02d}-appendix.html"
            if not html_path.exists():
                raise FileNotFoundError(f"Missing appendix HTML: {html_path}")

            rows.append(
                build_post_row(
                    week_number=week,
                    kind="appendix",
                    metadata=metadata,
                    carry_forward=carry_forward,
                    weekly_analytic_brief=weekly_analytic_brief,
                    html_path=html_path,
                    post_status=args.status,
                    image_base_url=args.image_base_url.rstrip("/"),
                    post_date=post_date,
                    logger=logger,
                    existing=export_rows.get((week, "appendix")),
                )
            )

    df = pd.DataFrame(rows)

    # Post-pass: fill cross-link ACF slugs using the WordPress Slug values (authoritative)
    # Rule:
    #   - Narrative row gets Appendix Post Slug = appendix row's Slug
    #   - Appendix row gets Narrative Post Slug = narrative row's Slug
    if args.update and not df.empty and "Slug" in df.columns and "Week Number" in df.columns and "Categories" in df.columns:
        for wk in sorted(set(int(x) for x in df["Week Number"].tolist() if str(x).strip() != "")):
            wk_rows = df[df["Week Number"] == wk]
            nar = wk_rows[wk_rows["Categories"].astype(str).str.lower().str.contains("clock narrative")]
            app = wk_rows[wk_rows["Categories"].astype(str).str.lower().str.contains("event log")]

            nar_slug = str(nar.iloc[0]["Slug"]) if len(nar) == 1 else ""
            app_slug = str(app.iloc[0]["Slug"]) if len(app) == 1 else ""

            if nar_slug and app_slug:
                df.loc[nar.index, "Appendix Post Slug"] = app_slug
                df.loc[app.index, "Narrative Post Slug"] = nar_slug

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

    # --- Emit clock_status CPT WP All Import files ---
    # 1. Weekly status file
    df_clock = pd.DataFrame(clock_weekly_rows)
    for col in CLOCK_STATUS_COLUMN_ORDER:
        if col not in df_clock.columns:
            df_clock[col] = ""
    df_clock = df_clock[CLOCK_STATUS_COLUMN_ORDER]
    start = args.week
    end = args.week + args.weeks - 1
    clock_weeks_path = WP_OUTPUT_ROOT / f"wordpress_import_clock_weeks_{start:02d}_{end:02d}.xlsx"
    df_clock.to_excel(clock_weeks_path, index=False)
    logger.info(f"Wrote clock_status weekly import file: {clock_weeks_path}")

    # 2. Current singleton file
    if clock_last_week_payload is None:
        raise RuntimeError("No weeks processed; cannot emit clock_status current import file.")
    last_week, last_metadata, last_carry = clock_last_week_payload
    current_row = build_clock_status_row(
        week_number=last_week,
        status_type="current",
        metadata=last_metadata,
        carry_forward=last_carry,
        post_status=args.status,
        narrative_post_slug=f"week{last_week:02d}-narrative",
        appendix_post_slug=f"week{last_week:02d}-appendix",
    )
    df_current = pd.DataFrame([current_row])
    for col in CLOCK_STATUS_COLUMN_ORDER:
        if col not in df_current.columns:
            df_current[col] = ""
    df_current = df_current[CLOCK_STATUS_COLUMN_ORDER]
    clock_current_path = WP_OUTPUT_ROOT / "wordpress_import_clock_current.xlsx"
    df_current.to_excel(clock_current_path, index=False)
    logger.info(f"Wrote clock_status current import file: {clock_current_path}")


if __name__ == "__main__":
    main()