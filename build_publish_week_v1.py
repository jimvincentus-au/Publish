# Build Publish Week v1

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

import argparse
import json
import logging
import shutil
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional

from publish_config_v1 import (
    STEP3_WEEKS_DIR,
    SUBSTACK_OUTPUT_DIR,
    SCRIVENER_OUTPUT_DIR,
    PUBLISH_LOGS_DIR,
    PUBLISH_IMAGES_DIR,
)


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
    narrative_draft3: Path
    metadata_json: Path
    appendix: Path
    image_wide: Optional[Path]
    image_prompt_wide: Optional[Path]
    image_prompt_square: Optional[Path]


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

    return WeekPaths(
        week_number=week_number,
        week_dir=week_dir,
        narrative_publish=narrative_publish,
        narrative_draft3=narrative_draft3,
        metadata_json=metadata_json,
        appendix=appendix,
        image_wide=image_wide,
        image_prompt_wide=image_prompt_wide,
        image_prompt_square=image_prompt_square,
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

# Helper: Make a Scrivener-friendly filename based on week and metadata
def make_scrivener_filename(week: int, meta: Dict[str, Any]) -> str:
    """Return a Scrivener-friendly filename like 'Week 01 - Title'.

    Falls back gracefully if metadata is missing or malformed.
    """
    raw_title = first_key(
        meta,
        [
            "Title",
            "title",
            "scrivener_title",
            "PostTitle",
            "post_title",
            "substack_title",
        ],
        default="",
    ).strip()

    # If the title already starts with "Week" (any case), keep it verbatim.
    if raw_title and raw_title.lower().startswith("week "):
        base = raw_title
    elif raw_title:
        base = f"Week {week:02d} - {raw_title}"
    else:
        base = f"Week {week:02d}"

    # Scrivener is fine with spaces and punctuation, but we defensively
    # clean out truly problematic filesystem characters.
    base = re.sub(r"[\\/:*?\"<>|]", "-", base)
    base = re.sub(r"\s+", " ", base).strip()

    return base


# Helper: Compute a week date range string from metadata
def get_week_date_range(meta: Dict[str, Any]) -> str:
    """
    Return a human-readable week date range string like
    '2025-01-20 – 2025-01-24' (or whatever format is stored),
    based on metadata fields.
    """
    start = first_key(
        meta,
        ["WeekStartDate", "week_start", "Week Start", "week_start_date"],
        default="",
    )
    end = first_key(
        meta,
        ["WeekEndDate", "week_end", "Week End", "week_end_date"],
        default="",
    )

    if start and end:
        return f"{start} - {end}"
    if start:
        return start
    if end:
        return end
    return ""

# Helper: Extract epigraphs as a list of strings
def get_epigraphs(meta: Dict[str, Any]) -> list[str]:
    """
    Return a list of epigraph strings from the metadata.

    Supports:
      * "Epigraphs":
          - a list of strings, or
          - a list of dicts with text/quote + optional source/author/attribution
      * "Epigraph1", "Epigraph2", ... style keys
      * A single "Epigraph" string as a fallback
      * A single multiline string under "Epigraphs"/"epigraphs" (split on newlines)
    """
    epigraphs: list[str] = []

    def _combine_text_and_source(item: Dict[str, Any]) -> str:
        """Build a human-readable epigraph line from a dict item."""
        # Prefer these keys for the main text of the epigraph.
        text_keys = ["text", "Text", "quote", "Quote", "line", "Line", "EpigraphText", "epigraph_text"]
        # Prefer these keys for the source/attribution.
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

    # Try plural epigraphs first.
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
        # Single multiline string: split on newlines and treat non-empty lines as epigraphs.
        for line in raw_eps.splitlines():
            if line.strip():
                epigraphs.append(line.strip())

    # Fallback: numbered epigraph keys.
    for i in range(1, 6):
        key_candidates = [f"Epigraph{i}", f"epigraph{i}"]
        for key in key_candidates:
            val = meta.get(key)
            if isinstance(val, str) and val.strip():
                epigraphs.append(val.strip())
                break  # don't double-add if both casings exist

    # Final fallback: single Epigraph string if nothing else present.
    if not epigraphs:
        single = meta.get("Epigraph") or meta.get("epigraph")
        if isinstance(single, str) and single.strip():
            epigraphs.append(single.strip())

    return epigraphs


# Helper: Format metadata for Scrivener notes
def format_metadata_for_notes(week: int, meta: Dict[str, Any]) -> str:
    """
    Format the metadata JSON into a human-readable text block suitable
    for Scrivener's document notes.

    We avoid raw JSON and instead surface the most important fields in a
    structured, prose-friendly way. Remaining fields are listed in a
    compact "Other metadata" section.
    """
    lines: list[str] = []

    # Core fields
    title = first_key(
        meta,
        ["Title", "title", "PostTitle", "post_title", "substack_title", "scrivener_title"],
        default="",
    )
    subtitle = first_key(
        meta,
        ["Subtitle", "subtitle", "FramingTitle", "framing_title"],
        default="",
    )
    tagline = first_key(
        meta,
        ["Tagline", "tagline", "Hook", "hook", "Tagline/Hook", "tagline_hook"],
        default="",
    )
    long_synopsis = first_key(
        meta,
        ["Long Synopsis", "long_synopsis", "LongSynopsis"],
        default="",
    )
    short_synopsis = first_key(
        meta,
        ["Short Synopsis", "short_synopsis", "ShortSynopsis"],
        default="",
    )
    seo_title = first_key(
        meta,
        ["SEO Title", "seo_title", "SEOTitle"],
        default="",
    )
    seo_description = first_key(
        meta,
        ["SEO Description", "seo_description", "SEODescription"],
        default="",
    )
    clock_time = first_key(
        meta,
        ["ClockTime", "clock_time", "Clock Time Reference", "Clock Time"],
        default="",
    )
    week_date_range = get_week_date_range(meta)

    # Epigraphs
    epigraphs = get_epigraphs(meta)

    # Helper to normalize list-or-string fields into a bullet list
    def normalize_list_field(value: Any) -> list[str]:
        if isinstance(value, list):
            out: list[str] = []
            for v in value:
                if isinstance(v, str) and v.strip():
                    out.append(v.strip())
            return out
        if isinstance(value, str) and value.strip():
            # Try to split on newlines or semicolons; if that fails, just return single-item list
            parts = [p.strip() for p in value.replace(";", "\n").splitlines() if p.strip()]
            return parts or [value.strip()]
        return []

    category_flags = normalize_list_field(
        meta.get("Category Flags")
        or meta.get("category_flags")
        or meta.get("Categories")
        or meta.get("categories")
    )
    key_traits = normalize_list_field(
        meta.get("Key Traits Referenced")
        or meta.get("key_traits_referenced")
        or meta.get("KeyTraits")
        or meta.get("key_traits")
    )
    actionable_outcomes = normalize_list_field(
        meta.get("Actionable Outcomes")
        or meta.get("actionable_outcomes")
        or meta.get("Actions")
    )
    quote_sources = normalize_list_field(
        meta.get("Quote Sources")
        or meta.get("quote_sources")
        or meta.get("Epigraph Sources")
        or meta.get("epigraph_sources")
    )
    internal_tags = normalize_list_field(
        meta.get("InternalTags")
        or meta.get("Internal Tags")
        or meta.get("Tags")
        or meta.get("tags")
    )
    publishing_status = first_key(
        meta,
        ["Publishing Status", "publishing_status", "Status", "status"],
        default="",
    )
    delta_summary = first_key(
        meta,
        ["Delta Summary", "delta_summary", "DeltaSummary"],
        default="",
    )

    # Header
    lines.append(f"Metadata for Week {week}")
    lines.append("=" * len(lines[-1]))
    lines.append("")

    # Basic identification
    if title:
        lines.append("Title")
        lines.append("-----")
        lines.append(title)
        lines.append("")
    if subtitle:
        lines.append("Subtitle")
        lines.append("--------")
        lines.append(subtitle)
        lines.append("")
    if tagline:
        lines.append("Tagline / Hook")
        lines.append("-------------")
        lines.append(tagline)
        lines.append("")

    if clock_time or week_date_range:
        lines.append("Clock & Dates")
        lines.append("------------")
        if clock_time:
            lines.append(f"Clock Time Reference: {clock_time}")
        if week_date_range:
            lines.append(f"Week Date Range: {week_date_range}")
        lines.append("")

    # Synopses
    if long_synopsis or short_synopsis:
        lines.append("Synopses")
        lines.append("--------")
        if long_synopsis:
            lines.append("Long Synopsis:")
            lines.append(long_synopsis)
            lines.append("")
        if short_synopsis:
            lines.append("Short Synopsis:")
            lines.append(short_synopsis)
            lines.append("")
    # SEO
    if seo_title or seo_description:
        lines.append("SEO Metadata")
        lines.append("-----------")
        if seo_title:
            lines.append(f"SEO Title: {seo_title}")
        if seo_description:
            lines.append("SEO Description:")
            lines.append(seo_description)
        lines.append("")

    # Epigraphs
    if epigraphs:
        lines.append("Epigraphs")
        lines.append("---------")
        for idx, ep in enumerate(epigraphs, start=1):
            lines.append(f"{idx}. {ep}")
        lines.append("")

    # Category flags, traits, actions, quotes, tags
    if category_flags:
        lines.append("Category Flags")
        lines.append("-------------")
        for flag in category_flags:
            lines.append(f"- {flag}")
        lines.append("")
    if key_traits:
        lines.append("Key Traits Referenced")
        lines.append("---------------------")
        for trait in key_traits:
            lines.append(f"- {trait}")
        lines.append("")
    if actionable_outcomes:
        lines.append("Actionable Outcomes")
        lines.append("-------------------")
        for act in actionable_outcomes:
            lines.append(f"- {act}")
        lines.append("")
    if quote_sources:
        lines.append("Quote / Epigraph Sources")
        lines.append("------------------------")
        for src in quote_sources:
            lines.append(f"- {src}")
        lines.append("")
    if internal_tags:
        lines.append("Internal Tags / Keywords")
        lines.append("------------------------")
        for tag in internal_tags:
            lines.append(f"- {tag}")
        lines.append("")

    if publishing_status or delta_summary:
        lines.append("Publishing Notes")
        lines.append("----------------")
        if publishing_status:
            lines.append(f"Status: {publishing_status}")
        if delta_summary:
            lines.append("Delta Summary:")
            lines.append(delta_summary)
        lines.append("")

    # Capture any remaining keys that we haven't explicitly rendered,
    # so that nothing is silently lost.
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
        if k in known_keys:
            continue
        if v is None:
            continue
        if isinstance(v, (str, int, float, bool)):
            other_items.append(f"- {k}: {v}")
        else:
            # Compact JSON for complex structures
            try:
                v_str = json.dumps(v, ensure_ascii=False)
            except Exception:
                v_str = str(v)
            other_items.append(f"- {k}: {v_str}")

    if other_items:
        lines.append("Other Metadata")
        lines.append("--------------")
        lines.extend(other_items)
        lines.append("")

    return "\n".join(lines)


# Helper: Build appendix markdown from events_appendix JSON
def build_appendix_from_json(path: Path) -> str:
    """
    Convert the events_appendix_weekNN.json structure into a categorized,
    human-readable Markdown appendix for the "Week Events" section.

    Primary expected schema (v1):
        {
          "categories": [
            {
              "name": "Category Title",
              "events": [
                {
                  "date": "...",
                  "actor": "...",
                  "action": "...",
                  "summary_line": "...",
                  "source": "...",
                  "url": "..."
                },
                ...
              ]
            },
            ...
          ]
        }

    Fallbacks:
      * If top-level is a list of events, group by a category/domain field.
      * If schema is completely unknown, embed JSON for debugging.
    """
    ensure_exists(path, "week appendix (events_appendix JSON)")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # pragma: no cover - defensive
        logger.error("Failed to parse appendix JSON at %s: %s", path, exc)
        return f"<!-- Failed to parse appendix JSON at {path}: {exc} -->"

    # Utility to pull a field value from an event with multiple key options.
    def ev_field(ev: Dict[str, Any], keys: list[str], default: str = "") -> str:
        for k in keys:
            v = ev.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return default

    lines: list[str] = []

    # ── Primary case: top-level dict with "categories" ────────────────────────
    if isinstance(data, dict) and isinstance(data.get("categories"), list):
        categories = data.get("categories") or []
        for cat in categories:
            if not isinstance(cat, dict):
                continue
            cat_name = (
                ev_field(cat, ["name", "Name", "category", "Category"], default="Week Events")
                or "Week Events"
            )
            events = cat.get("events") or cat.get("Events") or []
            if not isinstance(events, list) or not events:
                continue

            # Category heading
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

                # First line: "1. Actor — action (date) (Source)"
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
                if label:
                    first_line = f"{idx}. {label}"
                else:
                    first_line = f"{idx}."

                # Attach summary to first line if present
                if summary:
                    first_line = f"{first_line}: {summary}"
                lines.append(first_line)

                # Optional URL on its own indented line
                if url:
                    lines.append(f"   {url}")

                lines.append("")  # blank line between events

            # Extra blank line between categories (will be trimmed later)
            lines.append("")

        # Trim trailing blank lines
        while lines and not lines[-1].strip():
            lines.pop()

        return "\n".join(lines).rstrip()

    # ── Fallback: top-level list of events, group by category/domain field ───
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
                if label:
                    first_line = f"{idx}. {label}"
                else:
                    first_line = f"{idx}."

                # Attach summary to first line if present
                if summary:
                    first_line = f"{first_line}: {summary}"
                lines.append(first_line)
                if url:
                    lines.append(f"   {url}")
                lines.append("")

        while lines and not lines[-1].strip():
            lines.pop()

        return "\n".join(lines).rstrip()

    # ── Last resort: unknown structure, embed JSON for inspection ────────────
    lines.append("```json")
    lines.append(json.dumps(data, indent=2, ensure_ascii=False))
    lines.append("```")
    return "\n".join(lines).rstrip()


# ── Content assembly ──────────────────────────────────────────────────────────



def build_substack_markdown(
    week: int,
    narrative: str,
    appendix: str,
    meta: Dict[str, Any],
    paths: WeekPaths,
) -> str:
    """
    Construct Substack-ready Markdown.

    We keep a YAML-style header for Substack metadata, and then render
    a visible title, subtitle, epigraph blockquotes, the narrative body,
    and finally the categorized Week Events appendix.
    """

    # Base title from metadata
    base_title = first_key(
        meta,
        ["Title", "title", "substack_title", "PostTitle", "post_title"],
        default="",
    ).strip()

    # Display title: "Week N: {Title}" unless the title already starts with "Week"
    if base_title:
        if base_title.lower().startswith("week "):
            display_title = base_title
        else:
            display_title = f"Week {week}: {base_title}"
    else:
        display_title = f"Week {week}"

    subtitle = first_key(
        meta,
        ["Subtitle", "subtitle", "FramingTitle", "framing_title"],
        default="",
    )
    clock_time = first_key(
        meta,
        ["ClockTime", "clock_time", "Clock Time Reference"],
        default="",
    )

    tags = meta.get("InternalTags") or meta.get("Tags") or []
    if isinstance(tags, str):
        tags = [t.strip() for t in tags.split(",") if t.strip()]
    tags_str = ", ".join(str(t) for t in tags)

    header_image_name = paths.image_wide.name if paths.image_wide is not None else ""

    long_synopsis = first_key(
        meta,
        ["Long Synopsis", "long_synopsis", "LongSynopsis"],
        default="",
    )
    short_synopsis = first_key(
        meta,
        ["Short Synopsis", "short_synopsis", "ShortSynopsis"],
        default="",
    )
    seo_description = first_key(
        meta,
        ["SEO Description", "seo_description", "SEODescription"],
        default="",
    )

    week_date_range = get_week_date_range(meta)
    epigraphs = get_epigraphs(meta)

    # YAML-style metadata header for Substack
    header_lines = [
        "---",
        f'title: "{display_title}"',
    ]
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

    parts: list[str] = []
    parts.append(header)

    # Visible title and subtitle
    parts.append(f"# {display_title}")
    if subtitle:
        parts.append("")
        parts.append(f"*{subtitle}*")

    # Hero image for Substack body, placed between subtitle and epigraphs
    if header_image_name:
        parts.append("")
        parts.append(f"![Header image]({header_image_name})")

    # Epigraphs as blockquotes
    if epigraphs:
        parts.append("")
        for ep in epigraphs:
            parts.append(f"> {ep}")
        parts.append("")

    # Narrative body
    parts.append(narrative.rstrip())
    parts.append("")

    # Week events heading with optional date range
    if week_date_range:
        events_heading = f"## Week {week} Events ({week_date_range})"
    else:
        events_heading = f"## Week {week} Events"

    parts.append(events_heading)
    parts.append("")
    parts.append(appendix.rstrip())
    parts.append("")

    # Synopses for cross-posting, retained as comments
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
) -> str:
    """
    Construct a Scrivener-friendly version.

    Scrivener can ingest Markdown just fine; here we use headings instead of YAML
    and mirror the visible structure: Week N title, subtitle, epigraphs,
    narrative, then Week Events.
    """

    base_title = first_key(
        meta,
        ["Title", "title", "scrivener_title", "PostTitle", "post_title"],
        default="",
    ).strip()

    if base_title:
        if base_title.lower().startswith("week "):
            display_title = base_title
        else:
            display_title = f"Week {week}: {base_title}"
    else:
        display_title = f"Week {week}"

    subtitle = first_key(meta, ["Subtitle", "subtitle"], default="")
    week_date_range = get_week_date_range(meta)
    epigraphs = get_epigraphs(meta)

    lines: list[str] = []
    lines.append(f"# {display_title}")
    if subtitle:
        lines.append("")
        lines.append(f"*{subtitle}*")
    if epigraphs:
        lines.append("")
        for ep in epigraphs:
            lines.append(f"> {ep}")
    lines.append("")
    lines.append("<!-- Generated by build_publish_week_v1 -->")
    lines.append("")
    lines.append(narrative.rstrip())
    lines.append("")
    if week_date_range:
        lines.append(f"## Week {week} Events ({week_date_range})")
    else:
        lines.append(f"## Week {week} Events")
    lines.append("")
    lines.append(appendix.rstrip())
    lines.append("")

    return "\n".join(lines)


# ── Output helpers ────────────────────────────────────────────────────────────


def write_output(path: Path, content: str, force: bool) -> None:
    if path.exists() and not force:
        raise FileExistsError(f"Refusing to overwrite existing file without --force: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    logger.info("Wrote %s", path)


def build_publish_week(week_number: int, force: bool = False, use_publish: bool = False) -> None:
    # Resolve Step 3 inputs
    paths = discover_week_paths(week_number)
    logger.info("Building publish outputs for Week %s from %s", week_number, paths.week_dir)

    ensure_exists(paths.week_dir, "week directory")

    # Default: use _draft3 narrative.
    # If --use-publish is explicitly requested, prefer _publish when present.
    if paths.narrative_draft3.exists() and not use_publish:
        narrative_path = paths.narrative_draft3
        logger.info(
            "Using DRAFT3 narrative for Week %s (default; _publish ignored even if present): %s",
            week_number,
            narrative_path,
        )
    elif use_publish and paths.narrative_publish.exists():
        narrative_path = paths.narrative_publish
        logger.info(
            "Using PUBLISH narrative for Week %s (explicit --use-publish): %s",
            week_number,
            narrative_path,
        )
    elif paths.narrative_draft3.exists():
        # Fallback: --use-publish was requested but _publish is missing; use _draft3.
        narrative_path = paths.narrative_draft3
        logger.warning(
            "Requested --use-publish for Week %s but no _publish file found; falling back to DRAFT3: %s",
            week_number,
            narrative_path,
        )
    else:
        raise FileNotFoundError(
            f"Expected narrative (_publish or _draft3) for Week {week_number} in {paths.week_dir}"
        )

    # Metadata is now required for publish
    ensure_exists(paths.metadata_json, "metadata_stack JSON")
    metadata = load_metadata(paths.metadata_json)

    # Appendix is required and now comes from events_appendix_weekNN.json
    appendix_text = build_appendix_from_json(paths.appendix)

    # Wide image is required; prompts are optional but warned on
    if paths.image_wide is None:
        raise FileNotFoundError(
            f"Expected wide image for Week {week_number} at "
            f"{(paths.week_dir / f'image_wide_week{week_number}.png')}"
        )
    if paths.image_prompt_wide is None:
        logger.warning("No image_prompt_wide file found for Week %s", week_number)
    if paths.image_prompt_square is None:
        logger.warning("No image_prompt_square file found for Week %s", week_number)

    # Copy the wide image into the Publish Images folder for this week.
    # This gives Substack/Scrivener a stable place to pull from.
    image_dest_dir = PUBLISH_IMAGES_DIR / f"Week {week_number}"
    image_dest_dir.mkdir(parents=True, exist_ok=True)
    image_dest_path = image_dest_dir / paths.image_wide.name

    try:
        shutil.copy2(paths.image_wide, image_dest_path)
        logger.info(
            "Copied wide image for Week %s to %s",
            week_number,
            image_dest_path,
        )
    except Exception as exc:  # pragma: no cover - defensive
        logger.warning(
            "Failed to copy wide image for Week %s to %s: %s",
            week_number,
            image_dest_path,
            exc,
        )

    narrative_text = load_text(narrative_path)
    # Construct outputs
    substack_md = build_substack_markdown(week_number, narrative_text, appendix_text, metadata, paths)
    scrivener_md = build_scrivener_markdown(week_number, narrative_text, appendix_text, metadata, paths)

    # Determine output paths
    substack_dir = SUBSTACK_OUTPUT_DIR / f"Week {week_number}"
    scrivener_dir = SCRIVENER_OUTPUT_DIR / f"Week {week_number}"

    substack_path = substack_dir / f"week{week_number:02d}_substack.md"
    scrivener_filename = make_scrivener_filename(week_number, metadata)
    scrivener_path = scrivener_dir / f"{scrivener_filename}.md"

    # Scrivener companion files: synopsis and document notes
    long_synopsis = first_key(
        metadata,
        ["Long Synopsis", "long_synopsis", "LongSynopsis"],
        default="",
    )
    if not long_synopsis:
        # Fallback to short synopsis if a long one is not present
        long_synopsis = first_key(
            metadata,
            ["Short Synopsis", "short_synopsis", "ShortSynopsis"],
            default="",
        )

    synopsis_path = scrivener_dir / f"week{week_number:02d}_scrivener_synopsis.txt"
    notes_path = scrivener_dir / f"week{week_number:02d}_scrivener_notes.txt"

    synopsis_content = (long_synopsis or "").rstrip() + "\n"
    notes_content = format_metadata_for_notes(week_number, metadata)

    write_output(substack_path, substack_md, force=force)
    write_output(scrivener_path, scrivener_md, force=force)
    write_output(synopsis_path, synopsis_content, force=force)
    write_output(notes_path, notes_content, force=force)

    logger.info("Done building Week %s publish outputs.", week_number)


# ── CLI ───────────────────────────────────────────────────────────────────────


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build Substack / Scrivener publish artifacts for a Democracy Clock week.",
    )
    parser.add_argument(
        "--week",
        type=int,
        required=True,
        help="Week number (e.g. 1, 11, 43).",
    )
    parser.add_argument(
        "--weeks",
        type=int,
        default=1,
        help="Number of consecutive weeks to build, starting from --week.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing output files if they already exist.",
    )

    parser.add_argument(
        "--use-publish",
        action="store_true",
        help=(
            "Use step5_narrative_weekNN_publish.txt instead of "
            "step5_narrative_weekNN_draft3.txt when selecting the narrative."
        ),
    )

    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> None:
    args = parse_args(argv)
    start = args.week
    count = args.weeks

    current_week: Optional[int] = None
    try:
        for w in range(start, start + count):
            current_week = w
            build_publish_week(w, force=args.force, use_publish=args.use_publish)
    except Exception as exc:  # pragma: no cover - defensive main guard
        # Log the actual week that failed, not just the starting week
        failed = current_week if current_week is not None else start
        logger.error("Failed to build publish outputs for week %s: %s", failed, exc)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()