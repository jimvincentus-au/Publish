#!/usr/bin/env python3
# build_publish_week_v1 — assemble Substack / Scrivener outputs for a given week
#
# v1 goal:
#   * Read Step 3 "Week N" folder (narrative + metadata + assets)
#   * Produce Markdown files under Output/Substack and Output/Scrivener
#   * Keep the script conservative and file-name driven so it can be adapted later
#
# Usage:
#   python build_publish_week_v1 --week 11
#   python build_publish_week_v1 --week 11 --force   # overwrite existing outputs

from __future__ import annotations

# Added for normalization helpers
import argparse
import ast
import json
import logging
import shutil
import re
from datetime import datetime, timedelta
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, cast
# Optional: image conversion/resizing (used only for featured image artifacts)
try:
    from PIL import Image  # type: ignore
except Exception:  # pragma: no cover
    Image = None  # type: ignore

from publish_config_v1 import (
    STEP3_WEEKS_DIR,
    SUBSTACK_OUTPUT_DIR,
    SCRIVENER_OUTPUT_DIR,
    PUBLISH_LOGS_DIR,
)

# ── Output constants ──────────────────────────────────────────────────────────

OUTPUT_DIR = Path(__file__).parent / "Output"
WORDPRESS_OUTPUT_DIR = OUTPUT_DIR / "Wordpress"

# Canonical image output dir for publish artifacts (explicitly under /Publish/Output/Images)
PUBLISH_IMAGES_DIR = OUTPUT_DIR / "Images"

# Featured image normalization (WordPress-friendly)
# For a boxed container with no sidebar, standard featured images typically render at ~1200px.
FEATURED_IMAGE_MAX_WIDTH = 1200
FEATURED_IMAGE_JPG_QUALITY = 72

# ── Logging setup ──────────────────────────────────────────────────────────────


def setup_logger() -> logging.Logger:
    """Create a simple, file+console logger for the publish script."""
    PUBLISH_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_path = PUBLISH_LOGS_DIR / "build_publish_week_v1.log"

    logger = logging.getLogger("publish_week")
    logger.setLevel(logging.INFO)

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


logger = setup_logger()

# ── Data structures ───────────────────────────────────────────────────────────


@dataclass
class WeekPaths:
    week_number: int
    week_dir: Path
    narrative_publish: Path
    narrative_final: Path
    narrative_draft3: Path
    metadata_json: Path
    appendix: Path
    image_wide: Optional[Path]
    image_prompt_wide: Optional[Path]
    image_prompt_square: Optional[Path]
    weekly_brief: Optional[Path] = None
    carry_forward: Optional[Path] = None
    moral_floor: Optional[Path] = None
    appendix_image_wide: Optional[Path] = None


# ── Helpers ───────────────────────────────────────────────────────────────────


def discover_week_paths(week_number: int) -> WeekPaths:
    """
    Resolve all the Step 3 paths we care about for a given week.

    Assumes the Step 3 layout:
        STEP3_WEEKS_DIR / "Week {N}" / ...
        step5_narrative_week{N}_final.txt
        metadata_stack_week{N}.json
        image_wide_week{N}.png
        image_prompt_wide_week{N}.txt
        image_prompt_square_week{N}.txt
    """
    week_label = f"Week {week_number}"
    slug = f"week{week_number}"

    week_dir = STEP3_WEEKS_DIR / week_label
    narrative_publish = week_dir / f"step5_narrative_{slug}_publish.txt"
    narrative_final = week_dir / f"step5_narrative_{slug}_final.txt"
    narrative_draft3 = week_dir / f"step5_narrative_{slug}_draft3.txt"
    metadata_json = week_dir / f"metadata_stack_{slug}.json"
    appendix = week_dir / f"events_appendix_{slug}.json"

    image_wide = week_dir / f"image_wide_{slug}.png"
    if not image_wide.exists():
        image_wide = None

    image_prompt_wide = week_dir / f"image_prompt_wide_{slug}.txt"
    if not image_prompt_wide.exists():
        image_prompt_wide = None

    image_prompt_square = week_dir / f"image_prompt_square_{slug}.txt"
    if not image_prompt_square.exists():
        image_prompt_square = None

    # weekly analytic brief JSON (optional)
    weekly_brief = week_dir / f"weekly_analytic_brief_week{week_number}.json"
    if not weekly_brief.exists():
        weekly_brief = None

    # carry_forward and moral_floor (optional)
    # Note: artifacts are sometimes zero-padded (e.g., week02) and sometimes not (week2).
    slug = f"week{week_number}"

    carry_forward_candidates = [
        week_dir / f"carry_forward_{slug}.json",               # carry_forward_week2.json
        week_dir / f"carry_forward_week{week_number:02d}.json", # carry_forward_week02.json
        week_dir / f"carry_forward_week{week_number}.json",     # carry_forward_week2.json (alternate)
    ]
    carry_forward = next((p for p in carry_forward_candidates if p.exists()), None)

    moral_floor_candidates = [
        week_dir / f"moral_floor_{slug}.json",                 # moral_floor_week2.json (legacy)
        week_dir / f"moral_floor_week{week_number:02d}.json",   # moral_floor_week02.json (current)
        week_dir / f"moral_floor_week{week_number}.json",       # moral_floor_week2.json (alternate)
    ]
    moral_floor = next((p for p in moral_floor_candidates if p.exists()), None)

    # appendix wide image (optional)
    appendix_image_wide = week_dir / f"image_wide_week{week_number}_appendix.png"
    if not appendix_image_wide.exists():
        appendix_image_wide = None

    return WeekPaths(
        week_number=week_number,
        week_dir=week_dir,
        narrative_publish=narrative_publish,
        narrative_final=narrative_final,
        narrative_draft3=narrative_draft3,
        metadata_json=metadata_json,
        appendix=appendix,
        image_wide=image_wide,
        image_prompt_wide=image_prompt_wide,
        image_prompt_square=image_prompt_square,
        weekly_brief=weekly_brief,
        carry_forward=carry_forward,
        moral_floor=moral_floor,
        appendix_image_wide=appendix_image_wide,
    )


def ensure_exists(path: Path, description: str) -> None:
    """Raise a clean error if a required file is missing."""
    if not path.exists():
        raise FileNotFoundError(f"Expected {description} at {path} but it does not exist.")


def load_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def load_metadata(path: Path) -> Dict[str, Any]:
    if not path.exists():
        logger.warning("Metadata file missing at %s; continuing with empty metadata.", path)
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # pragma: no cover - defensive
        logger.error("Failed to parse metadata JSON at %s: %s", path, exc)
        return {}


def first_key(d: Dict[str, Any], keys: list[str], default: str = "") -> str:
    for k in keys:
        value = d.get(k)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return default


# ── Normalization/util helpers ────────────────────────────────────────────────

def _normalize_sources(value: object) -> str:
    """Return a comma-separated sources string with no brackets.

    Accepts:
      - list[str]
      - a comma-separated string
      - a stringified Python list like "['a', 'b']" or "[\"a\", \"b\"]"
    """
    if value is None:
        return ""

    if isinstance(value, list):
        parts = [str(x).strip() for x in value if str(x).strip()]
        return ", ".join(parts)

    if isinstance(value, str):
        s = value.strip()
        if not s:
            return ""

        # Try to parse stringified lists.
        if s.startswith("[") and s.endswith("]"):
            # Try JSON first
            try:
                parsed = json.loads(s)
                if isinstance(parsed, list):
                    parts = [str(x).strip() for x in parsed if str(x).strip()]
                    return ", ".join(parts)
            except Exception:
                pass

            # Fallback: Python literal list
            try:
                parsed = ast.literal_eval(s)
                if isinstance(parsed, list):
                    parts = [str(x).strip() for x in parsed if str(x).strip()]
                    return ", ".join(parts)
            except Exception:
                pass

        # Otherwise treat as a normal string; if it contains newlines or semicolons, normalize to commas.
        s = s.replace(";", ",")
        s = " ".join(s.split())
        return s

    return str(value).strip()


def _minutes_after_noon_to_time_str(minutes_after_noon: float | int) -> str:
    """Convert minutes-after-noon to a human time string like '7:31 p.m.'."""
    try:
        m = int(round(float(minutes_after_noon)))
    except Exception:
        return ""

    # minutes_after_noon is 0..720
    base = datetime(2000, 1, 1, 12, 0, 0)
    dt = base + timedelta(minutes=m)

    hour = dt.hour
    minute = dt.minute

    suffix = "a.m." if hour < 12 else "p.m."
    hour12 = hour % 12
    if hour12 == 0:
        hour12 = 12

    return f"{hour12}:{minute:02d} {suffix}"


def _extract_carry_forward_fields(carry_forward: dict) -> dict[str, object]:
    """Extract start/end dates and minutes from a carry_forward JSON with tolerant keying."""
    out: dict[str, object] = {}

    # Dates: prefer explicit keys; fallback to window.start/end if present.
    def _get_date(*keys: str) -> str:
        for k in keys:
            v = carry_forward.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return ""

    week_start_date = _get_date("week_start_date", "WeekStartDate", "start_date", "StartDate")
    week_end_date = _get_date("week_end_date", "WeekEndDate", "end_date", "EndDate")

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

    # Minutes: prefer canonical names; accept a few alternates.
    def _get_num(*keys: str) -> float | None:
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

    start_minutes = _get_num(
        "clock_before_minutes",
        "week_start_minutes",
        "start_minutes",
        "ClockBeforeMinutes",
    )
    end_minutes = _get_num(
        "clock_after_minutes",
        "week_end_minutes",
        "end_minutes",
        "ClockAfterMinutes",
    )

    out["week_start_date"] = week_start_date
    out["week_end_date"] = week_end_date
    out["week_start_minutes"] = start_minutes
    out["week_end_minutes"] = end_minutes

    # Convenience derived times
    out["week_start_time"] = _minutes_after_noon_to_time_str(start_minutes) if start_minutes is not None else ""
    out["week_end_time"] = _minutes_after_noon_to_time_str(end_minutes) if end_minutes is not None else ""

    return out


def make_scrivener_filename(week: int, meta: Dict[str, Any]) -> str:
    """Return a Scrivener-friendly filename like 'Week 01 - Title'."""
    raw_title = first_key(
        meta,
        ["Title", "title", "scrivener_title", "PostTitle", "post_title", "substack_title"],
        default="",
    ).strip()

    if raw_title and raw_title.lower().startswith("week "):
        base = raw_title
    elif raw_title:
        base = f"Week {week:02d} - {raw_title}"
    else:
        base = f"Week {week:02d}"

    base = re.sub(r"[\\/:*?\"<>|]", "-", base)
    base = re.sub(r"\s+", " ", base).strip()
    return base


def get_week_date_range(meta: Dict[str, Any]) -> str:
    start = first_key(meta, ["WeekStartDate", "week_start", "Week Start", "week_start_date"], default="")
    end = first_key(meta, ["WeekEndDate", "week_end", "Week End", "week_end_date"], default="")
    if start and end:
        return f"{start} - {end}"
    if start:
        return start
    if end:
        return end
    return ""


def _build_event_provenance_index(week_dir: Path, week: int) -> dict[str, dict[str, object]]:
    """Return {event_id: {'date': 'YYYY-MM-DD'|None, 'sources': [..]}} from whatever artifacts exist."""
    candidates: list[Path] = [
        week_dir / f"timeline_week{week}.json",
        week_dir / f"events_week{week}.json",
        week_dir / f"master_event_log_week{week}.json",
    ]
    candidates.extend(sorted(week_dir.glob(f"*week{week}*timeline*.json")))

    index: dict[str, dict[str, object]] = {}

    def _norm_sources(val: object) -> list[str]:
        if val is None:
            return []
        if isinstance(val, str):
            return [val]
        if isinstance(val, list):
            return [str(x) for x in val if x]
        return []

    def _maybe_add(event_id: str | None, date: str | None, sources: object) -> None:
        if not event_id:
            return
        rec = index.get(event_id, {"date": None, "sources": []})
        if date and not rec.get("date"):
            rec["date"] = date
        srcs = _norm_sources(sources)
        if srcs:
            existing_sources = rec.get("sources")
            if not isinstance(existing_sources, list):
                existing_sources = []
            rec["sources"] = list(dict.fromkeys([*existing_sources, *srcs]))
        index[event_id] = rec

    for p in candidates:
        if not p.exists() or not p.is_file():
            continue
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue

        def _walk_list(lst: list[object]) -> None:
            for ev in lst:
                if not isinstance(ev, dict):
                    continue
                eid = ev.get("event_id") or ev.get("id") or ev.get("event")
                date = ev.get("date") or ev.get("event_date") or ev.get("post_date")
                sources = ev.get("sources") or ev.get("source_urls") or ev.get("source") or ev.get("urls")
                _maybe_add(
                    str(eid) if eid is not None else None,
                    str(date) if date is not None else None,
                    sources,
                )

        if isinstance(data, dict):
            for k in ("events", "items", "rows", "timeline", "data"):
                v = data.get(k)
                if isinstance(v, list):
                    _walk_list(v)
        elif isinstance(data, list):
            _walk_list(data)

    return index


def _format_appendix_provenance(date: str | None, sources: list[str]) -> str:
    parts: list[str] = []
    if date:
        parts.append(date)
    if sources:
        parts.append(" ".join(sources))
    if not parts:
        return ""
    return " — " + " — ".join(parts)


def get_epigraphs(meta: Dict[str, Any]) -> list[str]:
    """Return a list of epigraph strings from the metadata."""
    epigraphs: list[str] = []

    def _combine_text_and_source(item: Dict[str, Any]) -> str:
        text_keys = ["text", "Text", "quote", "Quote", "line", "Line", "EpigraphText", "epigraph_text"]
        source_keys = ["source", "Source", "author", "Author", "attribution", "Attribution"]
        text_val = ""
        for k in text_keys:
            v = item.get(k)
            if isinstance(v, str) and v.strip():
                text_val = v.strip()
                break
        source_val = ""
        for k in source_keys:
            v = item.get(k)
            if isinstance(v, str) and v.strip():
                source_val = v.strip()
                break
        if text_val and source_val:
            return f"{text_val} — {source_val}"
        return text_val

    raw_eps = meta.get("Epigraphs") or meta.get("epigraphs")
    if isinstance(raw_eps, list):
        for item in raw_eps:
            if isinstance(item, str) and item.strip():
                epigraphs.append(item.strip())
            elif isinstance(item, dict):
                combined = _combine_text_and_source(item)
                if combined:
                    epigraphs.append(combined)
    elif isinstance(raw_eps, str) and raw_eps.strip():
        for line in raw_eps.splitlines():
            if line.strip():
                epigraphs.append(line.strip())

    for i in range(1, 6):
        for key in (f"Epigraph{i}", f"epigraph{i}"):
            val = meta.get(key)
            if isinstance(val, str) and val.strip():
                epigraphs.append(val.strip())
                break

    if not epigraphs:
        single = meta.get("Epigraph") or meta.get("epigraph")
        if isinstance(single, str) and single.strip():
            epigraphs.append(single.strip())

    return epigraphs


def format_metadata_for_notes(week: int, meta: Dict[str, Any]) -> str:
    """Format metadata JSON into a human-readable notes block."""
    lines: list[str] = []

    title = first_key(meta, ["Title", "title", "PostTitle", "post_title", "substack_title", "scrivener_title"], default="")
    subtitle = first_key(meta, ["Subtitle", "subtitle", "FramingTitle", "framing_title"], default="")
    tagline = first_key(meta, ["Tagline", "tagline", "Hook", "hook", "Tagline/Hook", "tagline_hook"], default="")
    long_synopsis = first_key(meta, ["Long Synopsis", "long_synopsis", "LongSynopsis"], default="")
    short_synopsis = first_key(meta, ["Short Synopsis", "short_synopsis", "ShortSynopsis"], default="")
    seo_title = first_key(meta, ["SEO Title", "seo_title", "SEOTitle"], default="")
    seo_description = first_key(meta, ["SEO Description", "seo_description", "SEODescription"], default="")
    clock_time = first_key(meta, ["ClockTime", "clock_time", "Clock Time Reference", "Clock Time"], default="")
    week_date_range = get_week_date_range(meta)
    epigraphs = get_epigraphs(meta)

    def normalize_list_field(value: Any) -> list[str]:
        if isinstance(value, list):
            return [v.strip() for v in value if isinstance(v, str) and v.strip()]
        if isinstance(value, str) and value.strip():
            parts = [p.strip() for p in value.replace(";", "\n").splitlines() if p.strip()]
            return parts or [value.strip()]
        return []

    category_flags = normalize_list_field(meta.get("Category Flags") or meta.get("category_flags") or meta.get("Categories") or meta.get("categories"))
    key_traits = normalize_list_field(meta.get("Key Traits Referenced") or meta.get("key_traits_referenced") or meta.get("KeyTraits") or meta.get("key_traits"))
    actionable_outcomes = normalize_list_field(meta.get("Actionable Outcomes") or meta.get("actionable_outcomes") or meta.get("Actions"))
    quote_sources = normalize_list_field(meta.get("Quote Sources") or meta.get("quote_sources") or meta.get("Epigraph Sources") or meta.get("epigraph_sources"))
    internal_tags = normalize_list_field(meta.get("InternalTags") or meta.get("Internal Tags") or meta.get("Tags") or meta.get("tags"))
    publishing_status = first_key(meta, ["Publishing Status", "publishing_status", "Status", "status"], default="")
    delta_summary = first_key(meta, ["Delta Summary", "delta_summary", "DeltaSummary"], default="")

    lines.append(f"Metadata for Week {week}")
    lines.append("=" * len(lines[-1]))
    lines.append("")

    if title:
        lines.extend(["Title", "-----", title, ""])
    if subtitle:
        lines.extend(["Subtitle", "--------", subtitle, ""])
    if tagline:
        lines.extend(["Tagline / Hook", "-------------", tagline, ""])

    if clock_time or week_date_range:
        lines.extend(["Clock & Dates", "------------"])
        if clock_time:
            lines.append(f"Clock Time Reference: {clock_time}")
        if week_date_range:
            lines.append(f"Week Date Range: {week_date_range}")
        lines.append("")

    if long_synopsis or short_synopsis:
        lines.extend(["Synopses", "--------"])
        if long_synopsis:
            lines.extend(["Long Synopsis:", long_synopsis, ""])
        if short_synopsis:
            lines.extend(["Short Synopsis:", short_synopsis, ""])

    if seo_title or seo_description:
        lines.extend(["SEO Metadata", "-----------"])
        if seo_title:
            lines.append(f"SEO Title: {seo_title}")
        if seo_description:
            lines.extend(["SEO Description:", seo_description])
        lines.append("")

    if epigraphs:
        lines.extend(["Epigraphs", "---------"])
        for idx, ep in enumerate(epigraphs, start=1):
            lines.append(f"{idx}. {ep}")
        lines.append("")

    if category_flags:
        lines.extend(["Category Flags", "-------------"])
        lines.extend([f"- {x}" for x in category_flags])
        lines.append("")
    if key_traits:
        lines.extend(["Key Traits Referenced", "---------------------"])
        lines.extend([f"- {x}" for x in key_traits])
        lines.append("")
    if actionable_outcomes:
        lines.extend(["Actionable Outcomes", "-------------------"])
        lines.extend([f"- {x}" for x in actionable_outcomes])
        lines.append("")
    if quote_sources:
        lines.extend(["Quote / Epigraph Sources", "------------------------"])
        lines.extend([f"- {x}" for x in quote_sources])
        lines.append("")
    if internal_tags:
        lines.extend(["Internal Tags / Keywords", "------------------------"])
        lines.extend([f"- {x}" for x in internal_tags])
        lines.append("")

    if publishing_status or delta_summary:
        lines.extend(["Publishing Notes", "----------------"])
        if publishing_status:
            lines.append(f"Status: {publishing_status}")
        if delta_summary:
            lines.extend(["Delta Summary:", delta_summary])
        lines.append("")

    # Other metadata
    known_keys = {
        "Title", "title", "PostTitle", "post_title", "substack_title", "scrivener_title",
        "Subtitle", "subtitle", "FramingTitle", "framing_title",
        "Tagline", "tagline", "Hook", "hook", "Tagline/Hook", "tagline_hook",
        "Long Synopsis", "long_synopsis", "LongSynopsis",
        "Short Synopsis", "short_synopsis", "ShortSynopsis",
        "SEO Title", "seo_title", "SEOTitle",
        "SEO Description", "seo_description", "SEODescription",
        "ClockTime", "clock_time", "Clock Time Reference", "Clock Time",
        "WeekStartDate", "week_start", "Week Start", "week_start_date",
        "WeekEndDate", "week_end", "Week End", "week_end_date",
        "Epigraphs", "epigraphs", "Epigraph1", "Epigraph2", "Epigraph3", "Epigraph4", "Epigraph5",
        "epigraph1", "epigraph2", "epigraph3", "epigraph4", "epigraph5",
        "Epigraph", "epigraph",
        "Category Flags", "category_flags", "Categories", "categories",
        "Key Traits Referenced", "key_traits_referenced", "KeyTraits", "key_traits",
        "Actionable Outcomes", "actionable_outcomes", "Actions",
        "Quote Sources", "quote_sources", "Epigraph Sources", "epigraph_sources",
        "InternalTags", "Internal Tags", "Tags", "tags",
        "Publishing Status", "publishing_status", "Status", "status",
        "Delta Summary", "delta_summary", "DeltaSummary",
    }

    other_items: list[str] = []
    for k, v in meta.items():
        if k in known_keys or v is None:
            continue
        if isinstance(v, (str, int, float, bool)):
            other_items.append(f"- {k}: {v}")
        else:
            try:
                v_str = json.dumps(v, ensure_ascii=False)
            except Exception:
                v_str = str(v)
            other_items.append(f"- {k}: {v_str}")

    if other_items:
        lines.extend(["Other Metadata", "--------------"])
        lines.extend(other_items)
        lines.append("")

    return "\n".join(lines)


def build_appendix_from_json(path: Path) -> str:
    """Convert events_appendix JSON into categorized Markdown."""
    ensure_exists(path, "week appendix (events_appendix JSON)")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.error("Failed to parse appendix JSON at %s: %s", path, exc)
        return f"<!-- Failed to parse appendix JSON at {path}: {exc} -->"

    def ev_field(ev: Dict[str, Any], keys: list[str], default: str = "") -> str:
        for k in keys:
            v = ev.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return default

    lines: list[str] = []

    if isinstance(data, dict) and isinstance(data.get("categories"), list):
        categories = data.get("categories") or []
        for cat in categories:
            if not isinstance(cat, dict):
                continue
            cat_name = ev_field(cat, ["name", "Name", "category", "Category"], default="Week Events") or "Week Events"
            events = cat.get("events") or cat.get("Events") or []
            if not isinstance(events, list) or not events:
                continue

            lines.append(f"### {cat_name}")
            lines.append("")

            for idx, ev in enumerate(events, start=1):
                if not isinstance(ev, dict):
                    continue
                date = ev_field(ev, ["date", "Date"])
                actor = ev_field(ev, ["actor", "Actor"])
                action = ev_field(ev, ["action", "Action"])
                summary = ev_field(ev, ["summary_line", "SummaryLine", "summary", "Summary"])
                source = ev_field(ev, ["source", "Source"])
                url = ev_field(ev, ["url", "URL", "link", "Link"])

                label_parts: list[str] = []
                if actor and action:
                    label_parts.append(f"{actor} {action}")
                elif actor:
                    label_parts.append(actor)
                elif action:
                    label_parts.append(action)

                if date:
                    if label_parts:
                        label_parts[-1] = f"{label_parts[-1]} ({date})"
                    else:
                        label_parts.append(f"({date})")

                if source:
                    if label_parts:
                        label_parts[-1] = f"{label_parts[-1]} ({source})"
                    else:
                        label_parts.append(f"({source})")

                label = " ".join(label_parts).strip()
                first_line = f"{idx}. {label}" if label else f"{idx}."
                if summary:
                    first_line = f"{first_line}: {summary}"
                lines.append(first_line)
                if url:
                    lines.append(f"   {url}")
                lines.append("")

            lines.append("")

        while lines and not lines[-1].strip():
            lines.pop()
        return "\n".join(lines).rstrip()

    if isinstance(data, list):
        groups: Dict[str, list[Dict[str, Any]]] = {}
        for ev in data:
            if not isinstance(ev, dict):
                continue
            category = ev_field(ev, ["category", "Category", "domain", "Domain"], default="Other")
            groups.setdefault(category, []).append(ev)

        first_group = True
        for category in sorted(groups.keys()):
            if not first_group:
                lines.append("")
            first_group = False
            lines.append(f"### {category}")
            lines.append("")
            for idx, ev in enumerate(groups[category], start=1):
                date = ev_field(ev, ["date", "Date"])
                actor = ev_field(ev, ["actor", "Actor"])
                action = ev_field(ev, ["action", "Action"])
                summary = ev_field(ev, ["summary_line", "SummaryLine", "summary", "Summary"])
                source = ev_field(ev, ["source", "Source"])
                url = ev_field(ev, ["url", "URL", "link", "Link"])

                label_parts: list[str] = []
                if actor and action:
                    label_parts.append(f"{actor} {action}")
                elif actor:
                    label_parts.append(actor)
                elif action:
                    label_parts.append(action)

                if date:
                    if label_parts:
                        label_parts[-1] = f"{label_parts[-1]} ({date})"
                    else:
                        label_parts.append(f"({date})")

                if source:
                    if label_parts:
                        label_parts[-1] = f"{label_parts[-1]} ({source})"
                    else:
                        label_parts.append(f"({source})")

                label = " ".join(label_parts).strip()
                first_line = f"{idx}. {label}" if label else f"{idx}."
                if summary:
                    first_line = f"{first_line}: {summary}"
                lines.append(first_line)
                if url:
                    lines.append(f"   {url}")
                lines.append("")

        while lines and not lines[-1].strip():
            lines.pop()
        return "\n".join(lines).rstrip()

    lines.append("```json")
    lines.append(json.dumps(data, indent=2, ensure_ascii=False))
    lines.append("```")
    return "\n".join(lines).rstrip()


def build_appendix_plaintext_from_json(path: Path) -> str:
    """Convert appendix markdown into plaintext."""
    md = build_appendix_from_json(path)
    logger.info("Converted appendix JSON at %s into plaintext appendix (%d characters)", path, len(md))
    lines_in = md.splitlines()
    lines_out: list[str] = []
    in_code_block = False

    for line in lines_in:
        stripped = line.lstrip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            lines_out.append(line)
            continue
        if stripped.startswith("### "):
            lines_out.append(stripped[4:].strip())
            continue
        if stripped.startswith("## "):
            lines_out.append(stripped[3:].strip())
            continue
        lines_out.append(line)

    return "\n".join(lines_out).rstrip()


def load_weekly_summary(paths: WeekPaths) -> str:
    if paths.weekly_brief is None:
        logger.warning("No weekly analytic brief JSON found for Week %s; appendix post will omit summary.", paths.week_number)
        return ""
    try:
        data = json.loads(paths.weekly_brief.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.error("Failed to parse weekly analytic brief JSON at %s: %s", paths.weekly_brief, exc)
        return ""
    for key in ["summary", "Summary", "weekly_summary", "WeeklySummary"]:
        val = data.get(key)
        if isinstance(val, str) and val.strip():
            return val.strip()
    logger.warning("Weekly analytic brief JSON for Week %s does not contain a usable summary field.", paths.week_number)
    return ""


def build_substack_appendix_markdown(
    week: int,
    summary: str,
    appendix_text: str,
    meta: Dict[str, Any],
    paths: WeekPaths,
) -> str:
    """Construct a Substack-ready appendix post (separate post)."""
    base_title = first_key(meta, ["Title", "title", "substack_title", "PostTitle", "post_title"], default="").strip()
    if base_title:
        display_title = f"Week {week} Appendix: {base_title}"
    else:
        display_title = f"Week {week} Appendix"

    subtitle = first_key(meta, ["Subtitle", "subtitle", "FramingTitle", "framing_title"], default="")
    clock_time = first_key(meta, ["ClockTime", "clock_time", "Clock Time Reference"], default="")

    tags = meta.get("InternalTags") or meta.get("Tags") or []
    if isinstance(tags, str):
        tags = [t.strip() for t in tags.split(",") if t.strip()]
    tags_str = ", ".join(str(t) for t in tags)

    header_image_name = paths.appendix_image_wide.name if paths.appendix_image_wide is not None else ""

    header_lines = ["---", f'title: "{display_title}"']
    if subtitle:
        header_lines.append(f'subtitle: "{subtitle}"')
    header_lines.append(f"week: {week}")
    if clock_time:
        header_lines.append(f'clock_time: "{clock_time}"')
    if tags_str:
        header_lines.append(f"tags: [{tags_str}]")
    header_lines.append("---")
    header_lines.append("")
    header_lines.append("<!-- Generated by build_publish_week_v1 (appendix post) -->")
    if header_image_name:
        header_lines.append(f"<!-- Header image: {header_image_name} -->")
    header_lines.append("")
    header = "\n".join(header_lines)

    parts: list[str] = [header, f"# {display_title}"]
    if subtitle:
        parts.extend(["", f"*{subtitle}*"])
    if header_image_name:
        parts.extend(["", f"![Header image]({header_image_name})"])
    parts.append("")
    if summary.strip():
        parts.extend([summary.strip(), ""])
    if appendix_text.strip():
        parts.extend([appendix_text.rstrip(), ""])
    return "\n".join(parts) + "\n"


def build_substack_markdown(
    week: int,
    narrative: str,
    appendix: str,
    meta: Dict[str, Any],
    paths: WeekPaths,
    include_appendix: bool = True,
) -> str:
    """Construct Substack-ready Markdown for the main weekly article."""
    base_title = first_key(meta, ["Title", "title", "substack_title", "PostTitle", "post_title"], default="").strip()

    if base_title:
        display_title = base_title if base_title.lower().startswith("week ") else f"Week {week}: {base_title}"
    else:
        display_title = f"Week {week}"

    subtitle = first_key(meta, ["Subtitle", "subtitle", "FramingTitle", "framing_title"], default="")
    clock_time = first_key(meta, ["ClockTime", "clock_time", "Clock Time Reference"], default="")

    tags = meta.get("InternalTags") or meta.get("Tags") or []
    if isinstance(tags, str):
        tags = [t.strip() for t in tags.split(",") if t.strip()]
    tags_str = ", ".join(str(t) for t in tags)

    header_image_name = paths.image_wide.name if paths.image_wide is not None else ""
    long_synopsis = first_key(meta, ["Long Synopsis", "long_synopsis", "LongSynopsis"], default="")
    short_synopsis = first_key(meta, ["Short Synopsis", "short_synopsis", "ShortSynopsis"], default="")
    seo_description = first_key(meta, ["SEO Description", "seo_description", "SEODescription"], default="")
    week_date_range = get_week_date_range(meta)
    epigraphs = get_epigraphs(meta)

    header_lines = ["---", f'title: "{display_title}"']
    if subtitle:
        header_lines.append(f'subtitle: "{subtitle}"')
    header_lines.append(f"week: {week}")
    if clock_time:
        header_lines.append(f'clock_time: "{clock_time}"')
    if tags_str:
        header_lines.append(f"tags: [{tags_str}]")
    header_lines.append("---")
    header_lines.append("")
    header_lines.append("<!-- Generated by build_publish_week_v1 -->")
    if header_image_name:
        header_lines.append(f"<!-- Header image: {header_image_name} -->")
    header_lines.append("")
    header = "\n".join(header_lines)

    parts: list[str] = [header, f"# {display_title}"]
    if subtitle:
        parts.extend(["", f"*{subtitle}*"])
    if header_image_name:
        parts.extend(["", f"![Header image]({header_image_name})"])

    if epigraphs:
        parts.append("")
        for ep in epigraphs:
            parts.append(f"> {ep}")
        parts.append("")

    parts.append(narrative.rstrip())
    parts.append("")

    if include_appendix and appendix.strip():
        events_heading = f"## Week {week} Events ({week_date_range})" if week_date_range else f"## Week {week} Events"
        parts.extend([events_heading, "", appendix.rstrip(), ""])

    if long_synopsis or short_synopsis or seo_description:
        parts.append("<!-- Synopses for cross-posting -->")
        if long_synopsis:
            parts.append(f"Long Synopsis: {long_synopsis}")
        if short_synopsis:
            parts.append(f"Short Synopsis: {short_synopsis}")
        if seo_description:
            parts.append(f"SEO Description: {seo_description}")
        parts.append("")

    return "\n".join(parts) + "\n"


def build_scrivener_markdown(
    week: int,
    narrative: str,
    appendix: str,
    meta: Dict[str, Any],
    paths: WeekPaths,
    include_appendix: bool = True,
) -> str:
    """Construct a Scrivener-friendly plaintext version of the main weekly chapter.

    Scrivener import is plain-text oriented; do not emit Markdown control tokens.
    Rules:
      - Keep title, but do not prefix with '#'
      - Keep subtitle, but do not wrap in '*'
      - Do not include epigraphs in Scrivener output
      - Do not include generation comments
      - Remove blank lines between paragraphs (collapse multiple newlines to single newlines)
    """

    base_title = first_key(meta, ["Title", "title", "scrivener_title", "PostTitle", "post_title"], default="").strip()
    display_title = base_title if (base_title and base_title.lower().startswith("week ")) else (f"Week {week}: {base_title}" if base_title else f"Week {week}")

    subtitle = first_key(meta, ["Subtitle", "subtitle"], default="").strip()

    # Normalize narrative to a clean plaintext form: CRLF -> LF, trim, and collapse
    # blank-line paragraph separators to single newlines.
    text = (narrative or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    text = re.sub(r"\n\s*\n+", "\n", text)

    lines: list[str] = [display_title]
    if subtitle:
        lines.append(subtitle)

    if text:
        # Ensure at least one line break between header block and body.
        lines.append(text)

    # Do not embed appendix in the Scrivener main chapter (current pipeline uses separate appendix).
    return "\n".join(lines).rstrip() + "\n"


def build_scrivener_appendix_markdown(
    week: int,
    summary: str,
    appendix_plaintext: str,
    meta: Dict[str, Any],
) -> str:
    """Construct a Scrivener-friendly appendix chapter (plaintext, no Markdown tokens)."""
    base_title = first_key(meta, ["Title", "title", "scrivener_title", "PostTitle", "post_title"], default="").strip()
    display_title = f"Week {week} Appendix: {base_title}" if base_title else f"Week {week} Appendix"
    week_date_range = get_week_date_range(meta)

    out: list[str] = [display_title]

    if summary and summary.strip():
        sum_text = summary.replace("\r\n", "\n").replace("\r", "\n").strip()
        sum_text = re.sub(r"\n\s*\n+", "\n", sum_text)
        if sum_text:
            out.append(sum_text)

    if appendix_plaintext and appendix_plaintext.strip():
        heading = f"Week {week} Events ({week_date_range})" if week_date_range else f"Week {week} Events"
        out.append(heading)
        out.append(appendix_plaintext.replace("\r\n", "\n").replace("\r", "\n").rstrip())

    return "\n".join(out).rstrip() + "\n"


def build_appendix_plaintext_with_provenance_from_json(
    path: Path,
    week_dir: Path,
    week: int,
    include_provenance: bool,
) -> str:
    """Convert events_appendix JSON into plaintext, optionally enriching each item with date + source URL(s)."""
    ensure_exists(path, "week appendix (events_appendix JSON)")

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.error("Failed to parse appendix JSON at %s: %s", path, exc)
        return f"<!-- Failed to parse appendix JSON at {path}: {exc} -->"

    provenance_index: dict[str, dict[str, object]] = {}
    if include_provenance:
        provenance_index = _build_event_provenance_index(week_dir=week_dir, week=week)

    def ev_field(ev: Dict[str, Any], keys: list[str], default: str = "") -> str:
        for k in keys:
            v = ev.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return default

    def ev_event_id(ev: Dict[str, Any]) -> str:
        eid = ev.get("event_id") or ev.get("id") or ev.get("event") or ev.get("eventId")
        return str(eid).strip() if eid is not None else ""

    def ev_sources_from_event(ev: Dict[str, Any]) -> list[str]:
        raw = ev.get("sources") or ev.get("source_urls") or ev.get("source") or ev.get("urls")
        if raw is None:
            return []
        if isinstance(raw, str) and raw.strip():
            return [raw.strip()]
        if isinstance(raw, list):
            return [str(x).strip() for x in raw if str(x).strip()]
        return []

    def ev_date_from_event(ev: Dict[str, Any]) -> str:
        return ev_field(ev, ["date", "Date", "event_date", "post_date"], default="")

    def lookup_provenance(ev: Dict[str, Any]) -> tuple[str | None, list[str]]:
        d = ev_date_from_event(ev) or None
        srcs = ev_sources_from_event(ev)

        if include_provenance:
            eid = ev_event_id(ev)
            if eid and eid in provenance_index:
                rec = provenance_index.get(eid, {})
                if d is None:
                    rec_date = rec.get("date")
                    if isinstance(rec_date, str) and rec_date.strip():
                        d = rec_date.strip()
                rec_sources = rec.get("sources")
                if isinstance(rec_sources, list):
                    merged = [*srcs, *[str(x).strip() for x in rec_sources if str(x).strip()]]
                    srcs = list(dict.fromkeys(merged))

        return d, srcs

    lines: list[str] = []

    if isinstance(data, dict) and isinstance(data.get("categories"), list):
        for cat in data.get("categories") or []:
            if not isinstance(cat, dict):
                continue
            cat_name = ev_field(cat, ["name", "Name", "category", "Category"], default="Week Events") or "Week Events"
            events = cat.get("events") or cat.get("Events") or []
            if not isinstance(events, list) or not events:
                continue

            lines.append(cat_name)
            lines.append("")

            for idx, ev in enumerate(events, start=1):
                if not isinstance(ev, dict):
                    continue
                actor = ev_field(ev, ["actor", "Actor"], default="")
                action = ev_field(ev, ["action", "Action"], default="")
                summary = ev_field(ev, ["summary_line", "SummaryLine", "summary", "Summary"], default="")

                label_parts: list[str] = []
                if actor and action:
                    label_parts.append(f"{actor} {action}")
                elif actor:
                    label_parts.append(actor)
                elif action:
                    label_parts.append(action)

                label = " ".join(label_parts).strip()
                first_line = f"{idx}. {label}" if label else f"{idx}."
                if summary:
                    first_line = f"{first_line}: {summary}"

                if include_provenance:
                    d, srcs = lookup_provenance(ev)
                    prov = _format_appendix_provenance(d, srcs)
                    if prov:
                        first_line = f"{first_line}{prov}"

                lines.append(first_line)
                lines.append("")

            lines.append("")

        while lines and not lines[-1].strip():
            lines.pop()
        return "\n".join(lines).rstrip()

    if isinstance(data, list):
        groups: Dict[str, list[Dict[str, Any]]] = {}
        for ev in data:
            if not isinstance(ev, dict):
                continue
            category = ev_field(ev, ["category", "Category", "domain", "Domain"], default="Other")
            groups.setdefault(category, []).append(ev)

        for category in sorted(groups.keys()):
            lines.append(category)
            lines.append("")
            for idx, ev in enumerate(groups[category], start=1):
                actor = ev_field(ev, ["actor", "Actor"], default="")
                action = ev_field(ev, ["action", "Action"], default="")
                summary = ev_field(ev, ["summary_line", "SummaryLine", "summary", "Summary"], default="")

                label_parts: list[str] = []
                if actor and action:
                    label_parts.append(f"{actor} {action}")
                elif actor:
                    label_parts.append(actor)
                elif action:
                    label_parts.append(action)

                label = " ".join(label_parts).strip()
                first_line = f"{idx}. {label}" if label else f"{idx}."
                if summary:
                    first_line = f"{first_line}: {summary}"

                if include_provenance:
                    d, srcs = lookup_provenance(ev)
                    prov = _format_appendix_provenance(d, srcs)
                    if prov:
                        first_line = f"{first_line}{prov}"

                lines.append(first_line)
                lines.append("")
            lines.append("")

        while lines and not lines[-1].strip():
            lines.pop()
        return "\n".join(lines).rstrip()

    return build_appendix_plaintext_from_json(path)



# ── Output helpers ────────────────────────────────────────────────────────────

# Helper to copy and normalize featured image names (WordPress-friendly sizing + JPG conversion)
def copy_featured_image(src: Path, dst_dir: Path, dst_name: str, week_number: int | None = None) -> Path:
    """Copy a featured image into Publish/Output/Images with normalized naming.

    Behavior:
      - Writes directly into Output/Images/ (no per-week subfolders)
      - If dst_name ends with .jpg/.jpeg, attempts to convert to JPG and resize to FEATURED_IMAGE_MAX_WIDTH
        to reduce file size.
      - If Pillow is unavailable or conversion fails, falls back to a direct copy.

    Resizing is conservative: only downscales if width exceeds FEATURED_IMAGE_MAX_WIDTH.
    """

    # All featured images are written directly into Publish/Output/Images.
    # Filenames are globally unique because they include the zero-padded week number.
    target_dir = dst_dir

    target_dir.mkdir(parents=True, exist_ok=True)
    dst = target_dir / dst_name

    if not src.exists():
        logger.warning("FEATURED IMAGE SOURCE MISSING: %s (will not write %s)", src, dst)
        return dst

    want_jpg = dst.suffix.lower() in {".jpg", ".jpeg"}
    src_suffix = src.suffix.lower()
    can_convert = (Image is not None) and (src_suffix in {".png", ".jpg", ".jpeg"})

    if want_jpg and can_convert and Image is not None:
        try:
            ImageMod = cast(Any, Image)
            with ImageMod.open(src) as im:
                if im.mode != "RGB":
                    im = im.convert("RGB")

                w, h = im.size
                if w > FEATURED_IMAGE_MAX_WIDTH:
                    new_w = FEATURED_IMAGE_MAX_WIDTH
                    new_h = int(round(h * (new_w / float(w))))
                    im = im.resize((new_w, new_h))

                im.save(dst, format="JPEG", quality=FEATURED_IMAGE_JPG_QUALITY, optimize=True)
                logger.info("FEATURED IMAGE WROTE (jpg): %s", dst)
                return dst
        except Exception as exc:
            logger.warning(
                "FEATURED IMAGE JPG CONVERSION FAILED: %s -> %s (%s); falling back to copy",
                src,
                dst,
                exc,
            )

    try:
        shutil.copy2(src, dst)
        logger.info("FEATURED IMAGE WROTE (copy): %s", dst)
    except Exception as exc:
        logger.warning("FEATURED IMAGE COPY FAILED: %s -> %s (%s)", src, dst, exc)

    return dst



# WordPress narrative HTML output helper (Elementor-safe, with title, subtitle, epigraphs)
# WordPress narrative HTML output helper (Elementor-safe, with title, subtitle, epigraphs)
def narrative_text_to_wordpress_html(
    *,
    week: int,
    narrative_text: str,
    meta: Dict[str, Any],
) -> str:
    """
    Build full WordPress-ready HTML for the weekly narrative, including:
      - H1 title
      - Optional subtitle
      - Optional epigraph block(s)
      - Body paragraphs

    Paragraph handling:
      - Prefer splitting on blank-line separators (\n\n or more).
      - If no blank lines exist (some pipelines collapse them), fall back to splitting on single newlines.
      - Lines inside a paragraph are normalized to single spaces.

    Output is Elementor-safe HTML (no blocks, no CSS).
    """

    def esc(s: str) -> str:
        return (
            (s or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    title = first_key(
        meta,
        ["Title", "title", "PostTitle", "post_title", "substack_title"],
        default=f"Week {week}",
    ).strip()

    if not title.lower().startswith("week "):
        title = f"Week {week}: {title}"

    subtitle = first_key(
        meta,
        ["Subtitle", "subtitle", "FramingTitle", "framing_title"],
        default="",
    )

    epigraphs = get_epigraphs(meta)

    text = (narrative_text or "").replace("\r\n", "\n").replace("\r", "\n").strip()

    # Prefer blank-line paragraph separation; if absent, fall back to single newlines.
    if re.search(r"\n\s*\n+", text):
        blocks = re.split(r"\n\s*\n+", text)
    else:
        # If there are no blank lines, treat each non-empty line as a paragraph.
        blocks = [ln for ln in text.split("\n") if ln.strip()]

    out: list[str] = []
    out.append("<!-- Generated by build_publish_week_v1 (WordPress narrative HTML) -->")

    # Title
    out.append(f"<h1>{esc(title)}</h1>")

    # Subtitle
    if subtitle:
        out.append(f"<p><em>{esc(subtitle)}</em></p>")

    # Epigraphs
    if epigraphs:
        for ep in epigraphs:
            out.append(f"<blockquote><p>{esc(ep)}</p></blockquote>")

    # Body paragraphs
    for b in blocks:
        b = (b or "").strip()
        if not b:
            continue

        # Normalize intra-paragraph newlines to spaces.
        b = re.sub(r"\s*\n\s*", " ", b).strip()

        # If the narrative contains markdown-style headings, render them as headings.
        # This helps readability and prevents headings being flattened into paragraphs.
        if b.startswith("### "):
            out.append(f"<h3>{esc(b[4:].strip())}</h3>")
            continue
        if b.startswith("## "):
            out.append(f"<h2>{esc(b[3:].strip())}</h2>")
            continue
        if b.startswith("# "):
            # Keep any stray H1 as H2 so we don't emit multiple H1s.
            out.append(f"<h2>{esc(b[2:].strip())}</h2>")
            continue

        out.append(f"<p>{esc(b)}</p>")

    return "\n".join(out).rstrip() + "\n"

# WordPress appendix HTML output helper (Elementor-safe, categories as H3, numbered items per category)
def appendix_json_to_wordpress_html(
    *,
    week: int,
    appendix_json_path: Path,
    meta: Dict[str, Any],
    summary: str = "",
) -> str:
    """Build WordPress-ready HTML for the appendix (event list).

    Rules:
      - H1 title
      - Optional subtitle (italic)
      - Optional summary paragraph(s)
      - Category headings are H3
      - Each category has its own ordered list (<ol>) and numbering resets per category

    Output is Elementor-safe HTML (no blocks, no CSS).
    """

    ensure_exists(appendix_json_path, "week appendix (events_appendix JSON)")

    base_title = first_key(
        meta,
        ["Title", "title", "PostTitle", "post_title", "substack_title"],
        default="",
    ).strip()
    if base_title:
        title = f"Week {week} Appendix: {base_title}"
    else:
        title = f"Week {week} Appendix"

    subtitle = first_key(
        meta,
        ["Subtitle", "subtitle", "FramingTitle", "framing_title"],
        default="",
    )

    try:
        data = json.loads(appendix_json_path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.error("Failed to parse appendix JSON at %s: %s", appendix_json_path, exc)
        return f"<!-- Failed to parse appendix JSON at {appendix_json_path}: {exc} -->\n"

    def esc(s: str) -> str:
        return (
            (s or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    def ev_field(ev: Dict[str, Any], keys: list[str], default: str = "") -> str:
        for k in keys:
            v = ev.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return default

    def _render_event_li(ev: Dict[str, Any]) -> str:
        # Prefer summary_line, fall back to summary
        summary_line = ev_field(ev, ["summary_line", "SummaryLine", "summary", "Summary"], default="").strip()
        actor = ev_field(ev, ["actor", "Actor"], default="").strip()
        action = ev_field(ev, ["action", "Action"], default="").strip()
        date = ev_field(ev, ["date", "Date"], default="").strip()

        # Build a short lead-in label when present
        lead_parts: list[str] = []
        if actor and action:
            lead_parts.append(f"{actor} {action}")
        elif actor:
            lead_parts.append(actor)
        elif action:
            lead_parts.append(action)

        lead = " ".join(lead_parts).strip()

        # Compose the visible line
        text_parts: list[str] = []
        if lead:
            text_parts.append(lead)
        if summary_line:
            if lead:
                text_parts.append(": ")
            text_parts.append(summary_line)

        text = "".join(text_parts).strip() or summary_line or lead or "(event)"

        # Append date (if present) as muted suffix
        if date:
            text = f"{text} ({date})"

        # Optional URL, placed on a new line
        url = ev_field(ev, ["url", "URL", "link", "Link"], default="").strip()
        if url:
            return f"<li>{esc(text)}<br><a href=\"{esc(url)}\" target=\"_blank\" rel=\"noopener\">{esc(url)}</a></li>"
        return f"<li>{esc(text)}</li>"

    out: list[str] = []
    out.append("<!-- Generated by build_publish_week_v1 (WordPress appendix HTML) -->")
    out.append(f"<h1>{esc(title)}</h1>")

    if subtitle:
        out.append(f"<p><em>{esc(subtitle)}</em></p>")

    if summary and summary.strip():
        # Preserve paragraph breaks in the summary
        sum_text = summary.replace("\r\n", "\n").replace("\r", "\n").strip()
        for para in re.split(r"\n\s*\n+", sum_text):
            para = re.sub(r"\s*\n\s*", " ", para.strip())
            if para:
                out.append(f"<p>{esc(para)}</p>")

    # Render categories
    if isinstance(data, dict) and isinstance(data.get("categories"), list):
        categories = data.get("categories") or []
        for cat in categories:
            if not isinstance(cat, dict):
                continue
            cat_name = ev_field(cat, ["name", "Name", "category", "Category"], default="Week Events") or "Week Events"
            events = cat.get("events") or cat.get("Events") or []
            if not isinstance(events, list) or not events:
                continue

            out.append(f"<h3>{esc(cat_name)}</h3>")
            out.append("<ol>")
            for ev in events:
                if isinstance(ev, dict):
                    out.append(_render_event_li(ev))
            out.append("</ol>")

        return "\n".join(out).rstrip() + "\n"

    # Fallback if appendix is a flat list: group by category/domain
    if isinstance(data, list):
        groups: Dict[str, list[Dict[str, Any]]] = {}
        for ev in data:
            if not isinstance(ev, dict):
                continue
            cat = ev_field(ev, ["category", "Category", "domain", "Domain"], default="Other") or "Other"
            groups.setdefault(cat, []).append(ev)

        for cat_name in sorted(groups.keys()):
            out.append(f"<h3>{esc(cat_name)}</h3>")
            out.append("<ol>")
            for ev in groups[cat_name]:
                out.append(_render_event_li(ev))
            out.append("</ol>")

        return "\n".join(out).rstrip() + "\n"

    # Unknown structure: emit JSON in a comment
    try:
        out.append("<!-- Unrecognized appendix structure; raw JSON follows -->")
        out.append("<!--")
        out.append(esc(json.dumps(data, ensure_ascii=False, indent=2)))
        out.append("-->")
    except Exception:
        out.append("<!-- Unrecognized appendix structure and could not serialize JSON. -->")

    return "\n".join(out).rstrip() + "\n"


def write_output(path: Path, content: str, force: bool) -> bool:
    """
    Write a file and log exactly what happened.
    Returns True if written, False if skipped (exists and no --force).
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists() and not force:
        logger.warning("SKIP (exists, no --force): %s", path)
        return False

    path.write_text(content, encoding="utf-8")
    logger.info("WROTE: %s (%d chars)", path, len(content))
    return True


def build_substack_outputs_for_week(
    *,
    week_number: int,
    substack_dir: Path,
    narrative_text: str,
    appendix_plain: str,
    paths: WeekPaths,
) -> list[Path]:
    """
    Write Substack output files for the week: plaintext narrative, plaintext appendix, and images.
    Returns a list of created Path objects.
    """
    substack_dir.mkdir(parents=True, exist_ok=True)

    narrative_txt_path = substack_dir / f"week{week_number:02d}_substack.txt"
    appendix_txt_path = substack_dir / f"week{week_number:02d}_appendix_substack.txt"

    # Plaintext exports are derived artifacts; always overwrite them for repeatability.
    wrote_narr = write_output(narrative_txt_path, narrative_text.rstrip() + "\n", force=True)
    wrote_app = write_output(appendix_txt_path, appendix_plain.rstrip() + "\n", force=True)

    created: list[Path] = [narrative_txt_path, appendix_txt_path]

    def _copy(src: Path, dst: Path) -> None:
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copy(src, dst)
            logger.info("COPIED: %s -> %s", src, dst)
        except Exception as exc:
            logger.warning("COPY FAILED: %s -> %s (%s)", src, dst, exc)

    if paths.image_wide is not None:
        substack_image_path = substack_dir / paths.image_wide.name
        _copy(paths.image_wide, substack_image_path)
        created.append(substack_image_path)

    if paths.appendix_image_wide is not None:
        substack_appendix_image_path = substack_dir / paths.appendix_image_wide.name
        _copy(paths.appendix_image_wide, substack_appendix_image_path)
        created.append(substack_appendix_image_path)

    logger.info(
        "Built Substack plaintext outputs for Week %s (narr_txt=%s, appendix_txt=%s)",
        week_number,
        "wrote" if wrote_narr else "skipped",
        "wrote" if wrote_app else "skipped",
    )
    return created


def build_publish_week(
    week_number: int,
    force: bool = False,
    use_publish: bool = False,
    include_appendix: bool = True,
    include_appendix_source: bool = True,
    build_substack: bool = True,
    build_scrivener: bool = True,
    build_wordpress: bool = False,
) -> None:
    paths = discover_week_paths(week_number)
    logger.info("Building publish outputs for Week %s from %s", week_number, paths.week_dir)

    # B) DEFINE WORDPRESS OUTPUT PATHS
    # Create per-week directory for WordPress output if enabled
    if build_wordpress:
        # Use zero-padded week folders so the filesystem sorts correctly.
        wordpress_week_dir = WORDPRESS_OUTPUT_DIR / f"Week {week_number:02d}"
        wordpress_week_dir.mkdir(parents=True, exist_ok=True)
        logger.info("WordPress week output dir: %s", wordpress_week_dir)
    else:
        wordpress_week_dir = None
    # Copy resolved week metadata into the WordPress output folder (sealed publish artifact)
    # The WordPress import generator should read metadata from Publish outputs, not Step 3.
    if build_wordpress and wordpress_week_dir is not None:
        try:
            wp_meta_dst = wordpress_week_dir / f"metadata_week{week_number:02d}.json"
            shutil.copy2(paths.metadata_json, wp_meta_dst)
            logger.info("COPIED: %s -> %s", paths.metadata_json, wp_meta_dst)
        except Exception as exc:
            logger.warning("COPY FAILED (metadata): %s -> %s (%s)", paths.metadata_json, wordpress_week_dir, exc)

        # Copy additional WordPress-relevant artifacts (sealed publish artifacts)
        def _copy_optional(src: Optional[Path], dst_name: str, label: str) -> None:
            if src is None:
                logger.info("WP artifact missing (%s): (none)", label)
                return
            if not src.exists():
                logger.info("WP artifact missing (%s): %s", label, src)
                return
            try:
                dst = wordpress_week_dir / dst_name
                shutil.copy2(src, dst)
                logger.info("COPIED: %s -> %s", src, dst)
            except Exception as exc:
                logger.warning("COPY FAILED (%s): %s -> %s (%s)", label, src, wordpress_week_dir, exc)

        _copy_optional(
            paths.weekly_brief,
            f"weekly_analytic_brief_week{week_number}.json",
            "weekly_analytic_brief",
        )
        _copy_optional(
            paths.carry_forward,
            f"carry_forward_week{week_number:02d}.json",
            "carry_forward",
        )
        _copy_optional(
            paths.moral_floor,
            f"moral_floor_week{week_number:02d}.json",
            "moral_floor",
        )
        _copy_optional(
            paths.appendix,
            f"events_appendix_week{week_number:02d}.json",
            "events_appendix",
        )

    ensure_exists(paths.week_dir, "week directory")

    # Narrative selection
    if use_publish:
        if paths.narrative_publish.exists():
            narrative_path = paths.narrative_publish
            logger.info("Using PUBLISH narrative for Week %s (explicit --use-publish): %s", week_number, narrative_path)
        elif paths.narrative_final.exists():
            narrative_path = paths.narrative_final
            logger.warning("Requested --use-publish for Week %s but no _publish file; falling back to FINAL: %s", week_number, narrative_path)
        elif paths.narrative_draft3.exists():
            narrative_path = paths.narrative_draft3
            logger.warning("Requested --use-publish for Week %s but no _publish/_final; falling back to DRAFT3: %s", week_number, narrative_path)
        else:
            raise FileNotFoundError(f"Expected narrative (_publish, _final, or _draft3) for Week {week_number} in {paths.week_dir}")
    else:
        if paths.narrative_final.exists():
            narrative_path = paths.narrative_final
            logger.info("Using FINAL narrative for Week %s (default): %s", week_number, narrative_path)
        elif paths.narrative_publish.exists():
            narrative_path = paths.narrative_publish
            logger.info("Using PUBLISH narrative for Week %s (no _final found; falling back): %s", week_number, narrative_path)
        elif paths.narrative_draft3.exists():
            narrative_path = paths.narrative_draft3
            logger.info("Using legacy DRAFT3 narrative for Week %s (no _final/_publish found): %s", week_number, narrative_path)
        else:
            raise FileNotFoundError(f"Expected narrative (_final, _publish, or _draft3) for Week {week_number} in {paths.week_dir}")

    ensure_exists(paths.metadata_json, "metadata_stack JSON")
    metadata = load_metadata(paths.metadata_json)

    # Appendix
    if include_appendix:
        appendix_text_md = build_appendix_from_json(paths.appendix)
        appendix_text_plain = build_appendix_plaintext_from_json(paths.appendix)
    else:
        appendix_text_md = ""
        appendix_text_plain = ""
        logger.info("Appendix disabled for Week %s via --no-appendix; skipping events_appendix JSON.", week_number)

    # Images
    if paths.image_wide is None:
        raise FileNotFoundError(
            f"Expected wide image for Week {week_number} at {(paths.week_dir / f'image_wide_week{week_number}.png')}"
        )
    if paths.image_prompt_wide is None:
        logger.warning("No image_prompt_wide file found for Week %s", week_number)
    if paths.image_prompt_square is None:
        logger.warning("No image_prompt_square file found for Week %s", week_number)

    # Copy images into Publish Images folder with canonical featured names
    if build_substack or build_scrivener or build_wordpress:
        image_dest_dir = PUBLISH_IMAGES_DIR

        # Narrative featured image
        if paths.image_wide is not None:
            copy_featured_image(
                src=paths.image_wide,
                dst_dir=image_dest_dir,
                dst_name=f"week{week_number:02d}-narrative-featured.jpg",
                week_number=week_number,
            )

        # Appendix featured image (if present)
        if paths.appendix_image_wide is not None:
            copy_featured_image(
                src=paths.appendix_image_wide,
                dst_dir=image_dest_dir,
                dst_name=f"week{week_number:02d}-appendix-featured.jpg",
                week_number=week_number,
            )

    narrative_text = load_text(narrative_path)

    # Main article outputs (no embedded appendix in body)
    substack_md = ""
    if build_substack:
        substack_md = build_substack_markdown(
            week_number,
            narrative_text,
            appendix_text_md,
            metadata,
            paths,
            include_appendix=False,
        )

    scrivener_md = ""
    if build_scrivener:
        scrivener_md = build_scrivener_markdown(
            week_number,
            narrative_text,
            appendix_text_md,
            metadata,
            paths,
            include_appendix=False,
        )

    # Appendix companion outputs
    if include_appendix:
        weekly_summary = load_weekly_summary(paths)
        appendix_plain = appendix_text_plain
        appendix_plain_for_scrivener = build_appendix_plaintext_with_provenance_from_json(
            paths.appendix,
            week_dir=paths.week_dir,
            week=week_number,
            include_provenance=include_appendix_source,
        )
    else:
        weekly_summary = ""
        appendix_plain = ""
        appendix_plain_for_scrivener = ""
        logger.info("Appendix disabled for Week %s via --no-appendix; no appendix summary/events generated.", week_number)

    # Output paths
    substack_dir = None
    scrivener_dir = None
    substack_path = None
    scrivener_path = None
    scrivener_appendix_path = None
    appendix_substack_path = None

    if build_substack:
        substack_dir = SUBSTACK_OUTPUT_DIR / f"Week {week_number}"
        substack_dir.mkdir(parents=True, exist_ok=True)
        substack_path = substack_dir / f"week{week_number:02d}_substack.md"
        appendix_substack_path = substack_dir / f"week{week_number:02d}_appendix_substack.md" if include_appendix else None
    else:
        substack_md = ""
        substack_dir = None
        substack_path = None
        appendix_substack_path = None

    if build_scrivener:
        scrivener_dir = SCRIVENER_OUTPUT_DIR / f"Week {week_number}"
        scrivener_dir.mkdir(parents=True, exist_ok=True)
        scrivener_filename = make_scrivener_filename(week_number, metadata)
        scrivener_path = scrivener_dir / f"{scrivener_filename}.md"
        scrivener_appendix_path = scrivener_dir / f"week{week_number:02d}_appendix.md" if include_appendix else None
    else:
        scrivener_md = ""
        scrivener_dir = None
        scrivener_path = None
        scrivener_appendix_path = None

    # Scrivener companion files: synopsis and document notes
    synopsis_path = None
    notes_path = None
    synopsis_content = ""
    notes_content = ""
    if build_scrivener and scrivener_dir is not None:
        long_synopsis = first_key(metadata, ["Long Synopsis", "long_synopsis", "LongSynopsis"], default="")
        if not long_synopsis:
            long_synopsis = first_key(metadata, ["Short Synopsis", "short_synopsis", "ShortSynopsis"], default="")
        synopsis_path = scrivener_dir / f"week{week_number:02d}_scrivener_synopsis.txt"
        notes_path = scrivener_dir / f"week{week_number:02d}_scrivener_notes.txt"
        synopsis_content = (long_synopsis or "").rstrip() + "\n"
        notes_content = format_metadata_for_notes(week_number, metadata)

    # ── Build / append/overwrite appendix artifacts (only if enabled) ──────────────
    def _rewrite_scrivener_appendix(existing_content: str, new_appendix: str) -> str:
        """Remove any prior generated appendix block and append a new one between markers."""
        begin_marker = "<!-- BEGIN GENERATED APPENDIX -->"
        end_marker = "<!-- END GENERATED APPENDIX -->"
        pattern = re.compile(r"\n?<!-- BEGIN GENERATED APPENDIX -->(.|\n)*?<!-- END GENERATED APPENDIX -->\n?", re.MULTILINE)
        cleaned = re.sub(pattern, "", existing_content).rstrip()

        block = f"{begin_marker}\n{new_appendix.rstrip()}\n{end_marker}"
        return f"{cleaned}\n\n{block}\n" if cleaned else f"{block}\n"

    if build_scrivener and include_appendix and scrivener_appendix_path is not None:
        if scrivener_appendix_path.exists():
            if force:
                existing = load_text(scrivener_appendix_path)
                logger.info("Overwriting Scrivener appendix for Week %s (--force): rewriting generated appendix block.", week_number)
                scrivener_appendix_md = _rewrite_scrivener_appendix(existing, appendix_plain_for_scrivener)
            else:
                existing = load_text(scrivener_appendix_path).rstrip()
                logger.info("Appending plaintext appendix to existing Scrivener appendix for Week %s", week_number)
                if appendix_plain_for_scrivener.strip():
                    scrivener_appendix_md = existing + "\n\n" + appendix_plain_for_scrivener.rstrip() + "\n"
                else:
                    scrivener_appendix_md = existing + "\n"
        else:
            scrivener_appendix_md = build_scrivener_appendix_markdown(
                week_number,
                weekly_summary,
                appendix_plain_for_scrivener,
                metadata,
            )
            logger.info("Created new Scrivener appendix with summary + events for Week %s", week_number)
    else:
        scrivener_appendix_md = None

    # End of build_publish_week appendix preparation

    # Build Substack appendix markdown (separate post) if appendix is enabled
    if build_substack and include_appendix:
        appendix_substack_md = build_substack_appendix_markdown(
            week=week_number,
            summary=weekly_summary,
            appendix_text=appendix_plain,
            meta=metadata,
            paths=paths,
        )
    else:
        appendix_substack_md = None

    # Only build Substack plaintext outputs (narrative + appendix txt + images) if enabled
    if build_substack and substack_dir is not None:
        substack_created_files = build_substack_outputs_for_week(
            week_number=week_number,
            substack_dir=substack_dir,
            narrative_text=narrative_text,
            appendix_plain=appendix_plain,
            paths=paths,
        )
    else:
        substack_created_files = []

    # C) WORDPRESS OUTPUT (Narrative + Appendix HTML only)
    wordpress_created_files: list[Path] = []
    if build_wordpress and wordpress_week_dir is not None:
        # --- Carry-forward fields extraction for ACF row fields ---
        # Load carry_forward_weekNN.json if present, else default to {}
        carry_forward_data: dict[str, Any] = {}
        if paths.carry_forward is not None and paths.carry_forward.exists():
            try:
                carry_forward_data = json.loads(paths.carry_forward.read_text(encoding="utf-8"))
            except Exception as exc:
                logger.warning("Failed to load carry_forward JSON at %s: %s", paths.carry_forward, exc)
                carry_forward_data = {}
        cf_fields = _extract_carry_forward_fields(carry_forward_data)

        # Narrative HTML
        wp_narrative_html = narrative_text_to_wordpress_html(
            week=week_number,
            narrative_text=narrative_text,
            meta=metadata,
        )
        wp_narrative_path = wordpress_week_dir / f"week{week_number:02d}-narrative.html"
        if write_output(wp_narrative_path, wp_narrative_html, force=force):
            wordpress_created_files.append(wp_narrative_path)

        # Appendix HTML (if enabled)
        if include_appendix:
            wp_appendix_summary = weekly_summary
            wp_appendix_html = appendix_json_to_wordpress_html(
                week=week_number,
                appendix_json_path=paths.appendix,
                meta=metadata,
                summary=wp_appendix_summary,
            )
            wp_appendix_path = wordpress_week_dir / f"week{week_number:02d}-appendix.html"
            if write_output(wp_appendix_path, wp_appendix_html, force=force):
                wordpress_created_files.append(wp_appendix_path)

        # --- Example WordPress row-building for ACF fields (for import XLSX, not shown here) ---
        # If you build a row for import, ensure:
        # row["week_start_date"] = str(cf_fields.get("week_start_date") or "")
        # row["week_end_date"] = str(cf_fields.get("week_end_date") or "")
        # row["week_start_minutes"] = cf_fields.get("week_start_minutes") if cf_fields.get("week_start_minutes") is not None else ""
        # row["week_end_minutes"] = cf_fields.get("week_end_minutes") if cf_fields.get("week_end_minutes") is not None else ""
        # row["week_start_time"] = str(cf_fields.get("week_start_time") or "")
        # row["week_end_time"] = str(cf_fields.get("week_end_time") or "")
        # row["sources"] = _normalize_sources(metadata.get("sources"))

    # Write enabled outputs only (so --only wordpress never touches Substack/Scrivener artifacts)
    wrote_substack_main = False
    wrote_scrivener_main = False
    wrote_synopsis = False
    wrote_notes = False
    wrote_substack_appendix = False
    wrote_scrivener_appendix = False

    if build_substack and substack_path is not None:
        wrote_substack_main = write_output(substack_path, substack_md, force=force)
    if build_scrivener and scrivener_path is not None:
        wrote_scrivener_main = write_output(scrivener_path, scrivener_md, force=force)
    if build_scrivener and synopsis_path is not None:
        wrote_synopsis = write_output(synopsis_path, synopsis_content, force=force)
    if build_scrivener and notes_path is not None:
        wrote_notes = write_output(notes_path, notes_content, force=force)

    if build_substack and include_appendix and appendix_substack_path is not None and appendix_substack_md is not None:
        wrote_substack_appendix = write_output(appendix_substack_path, appendix_substack_md, force=force)
    if build_scrivener and include_appendix and scrivener_appendix_path is not None and scrivener_appendix_md is not None:
        wrote_scrivener_appendix = write_output(scrivener_appendix_path, scrivener_appendix_md, force=force)


    logger.info(
        "WRITE SUMMARY (Week %s) substack_main=%s scrivener_main=%s synopsis=%s notes=%s substack_appendix=%s scrivener_appendix=%s",
        week_number,
        "wrote" if wrote_substack_main else "skipped",
        "wrote" if wrote_scrivener_main else "skipped",
        "wrote" if wrote_synopsis else "skipped",
        "wrote" if wrote_notes else "skipped",
        "wrote" if wrote_substack_appendix else "skipped",
        "wrote" if wrote_scrivener_appendix else "skipped",
    )

    # Log outputs
    if build_scrivener:
        scrivener_files = [scrivener_path, synopsis_path, notes_path]
        if include_appendix and scrivener_appendix_path is not None and scrivener_appendix_md is not None:
            scrivener_files.append(scrivener_appendix_path)
        logger.info("Scrivener outputs for Week %s: %s", week_number, ", ".join(str(p) for p in scrivener_files if p is not None))
    else:
        logger.info("Scrivener outputs for Week %s: (disabled)", week_number)

    if build_substack:
        substack_files = [substack_path] if substack_path is not None else []
        if include_appendix and appendix_substack_path is not None and appendix_substack_md is not None:
            substack_files.append(appendix_substack_path)
        substack_files.extend(substack_created_files)
        logger.info("Substack outputs for Week %s: %s", week_number, ", ".join(str(p) for p in substack_files if p is not None))
    else:
        logger.info("Substack outputs for Week %s: (disabled)", week_number)

    if build_wordpress:
        if wordpress_created_files:
            logger.info(
                "WordPress outputs for Week %s: %s",
                week_number,
                ", ".join(str(p) for p in wordpress_created_files),
            )
        else:
            logger.info("WordPress outputs for Week %s: (enabled, nothing written)", week_number)
    else:
        logger.info("WordPress outputs for Week %s: (disabled)", week_number)

    logger.info("Done building Week %s publish outputs.", week_number)


# ── CLI ───────────────────────────────────────────────────────────────────────


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build Substack / Scrivener / Wordpress publish artifacts for a Democracy Clock week.",
    )
    parser.add_argument("--week", type=int, required=True, help="Week number (e.g. 1, 11, 43).")
    parser.add_argument("--weeks", type=int, default=1, help="Number of consecutive weeks to build, starting from --week.")
    parser.add_argument("--force", action="store_true", help="Overwrite existing output files if they already exist.")
    parser.add_argument(
        "--use-publish",
        action="store_true",
        help=(
            "Use step5_narrative_weekNN_publish.txt instead of "
            "step5_narrative_weekNN_draft3.txt when selecting the narrative."
        ),
    )
    parser.add_argument(
        "--no-appendix",
        action="store_true",
        help=(
            "Build outputs without an events appendix. "
            "By default, an appendix is included using events_appendix_weekNN.json."
        ),
    )
    parser.add_argument(
        "--appendix-source",
        choices=["on", "off"],
        default="on",
        help="Scrivener appendix only: include each appendix item’s event date + source URL(s). Default: on.",
    )
    parser.add_argument(
        "--only",
        type=str,
        default="",
        help="Comma-separated targets to build (substack,scrivener,wordpress). If set, only these targets are built.",
    )
    parser.add_argument(
        "--skip",
        type=str,
        default="",
        help="Comma-separated targets to skip (substack,scrivener,wordpress). Skipped targets are not written.",
    )
    parser.add_argument(
        "--level",
        choices=["debug", "info", "warning", "error"],
        default="info",
        help="Logging level (default: info).",
    )
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> None:
    args = parse_args(argv)

    # Apply requested log level to the existing logger and its handlers.
    level_map = {
        "debug": logging.DEBUG,
        "info": logging.INFO,
        "warning": logging.WARNING,
        "error": logging.ERROR,
    }
    chosen = level_map.get(str(getattr(args, "level", "info")).lower(), logging.INFO)
    logger.setLevel(chosen)
    for h in getattr(logger, "handlers", []):
        try:
            h.setLevel(chosen)
        except Exception:
            pass
    logger.info("Log level set to %s", str(getattr(args, "level", "info")).lower())

    include_appendix_source: bool = (str(args.appendix_source).lower() == "on")
    logger.info("Scrivener appendix provenance (date+sources): %s", "on" if include_appendix_source else "off")

    start = args.week
    count = args.weeks
    include_appendix = not args.no_appendix

    # A) Target selection: --only and --skip
    valid_targets = {"substack", "scrivener", "wordpress"}

    def _parse_targets(raw: str) -> set[str]:
        raw = (raw or "").strip()
        if not raw:
            return set()
        parts = [p.strip().lower() for p in raw.split(",") if p.strip()]
        unknown = [p for p in parts if p not in valid_targets]
        if unknown:
            raise SystemExit(
                f"Unknown target(s) in {raw!r}: {', '.join(unknown)}. Valid: substack,scrivener,wordpress"
            )
        return set(parts)

    only_set = _parse_targets(getattr(args, "only", ""))
    skip_set = _parse_targets(getattr(args, "skip", ""))

    # If --only is provided, build ONLY those targets.
    # If --only is empty, build everything except what --skip excludes.
    build_substack = (not only_set or "substack" in only_set) and ("substack" not in skip_set)
    build_scrivener = (not only_set or "scrivener" in only_set) and ("scrivener" not in skip_set)
    build_wordpress = (not only_set or "wordpress" in only_set) and ("wordpress" not in skip_set)

    if only_set and skip_set:
        overlap = sorted(list(only_set.intersection(skip_set)))
        if overlap:
            logger.warning(
                "Targets appear in both --only and --skip (%s). --skip wins for these.",
                ", ".join(overlap),
            )

    current_week: Optional[int] = None
    try:
        for w in range(start, start + count):
            current_week = w
            build_publish_week(
                w,
                force=args.force,
                use_publish=args.use_publish,
                include_appendix=include_appendix,
                include_appendix_source=include_appendix_source,
                build_substack=build_substack,
                build_scrivener=build_scrivener,
                build_wordpress=build_wordpress,
            )
    except Exception as exc:
        failed = current_week if current_week is not None else start
        logger.error("Failed to build publish outputs for week %s: %s", failed, exc)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()