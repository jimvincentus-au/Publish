#!/usr/bin/env python3
"""publish_controller_v1.py — orchestrate the weekly Democracy Clock publish flow.

Runs the Publish routines for one or more weeks in dependency order. `publish`
produces the artifacts (narratives, week{NN}-appendix.html, digest) that the
other routines consume, so it runs first.

  Per week (looped):
    1. publish   build_publish_week_v2.py         Substack + Scrivener + WordPress artifacts
    2. assets    build_publish_site_assets_v1.py  clock chart / week images
    3. vellum    build_vellum_import_v1.py        Vellum book import (week-by-week)

  Once for the whole week range (after the loop):
    4. archive   build_wordpress_import_v1.py         WordPress Archive import
    5. digest    build_wordpress_digest_import_v1.py  WordPress digest import

Fail-fast: any routine failure aborts the run (no partial publishes). Mirrors the
Step 3 controller_v5 conventions — subprocess line-streaming and
propagate=False logging so sub-step output shows live and does not double.

Examples:
    # Publish the whole Vol-4 backfill (contiguous), all five routines:
    python publish_controller_v1.py --week 34 --weeks 12 --force
    # Only the Substack + WordPress artifacts for one week:
    python publish_controller_v1.py --week 34 --only publish --formats substack,wordpress
    # Re-run just the range imports over a span:
    python publish_controller_v1.py --week 34 --weeks 12 --only archive,digest
"""

from __future__ import annotations

import argparse
import logging
import subprocess
import sys
from pathlib import Path
from typing import List, Optional

PUBLISH_DIR = Path(__file__).resolve().parent
LOGS_DIR = PUBLISH_DIR / "Logs"

# Routine names, split by scope and listed in execution order.
WEEK_ORDER = ["publish", "assets", "vellum"]   # run per week, in this order
RANGE_ORDER = ["archive", "digest"]            # run once for the whole range
ALL_ROUTINES = WEEK_ORDER + RANGE_ORDER

SCRIPTS = {
    "publish": "build_publish_week_v2.py",
    "assets":  "build_publish_site_assets_v1.py",
    "vellum":  "build_vellum_import_v1.py",
    "archive": "build_wordpress_import_v1.py",
    "digest":  "build_wordpress_digest_import_v1.py",
}


def setup_logger(level: str) -> logging.Logger:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("publish_controller")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    # Do NOT propagate: sub-steps call logging.basicConfig(), which would install
    # a root handler and double every relayed line on the console.
    logger.propagate = False
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s")
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    fh = logging.FileHandler(LOGS_DIR / "publish_controller.log", encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(ch)
    logger.addHandler(fh)
    logger.info("Publish controller log → %s", LOGS_DIR / "publish_controller.log")
    return logger


def routine_argv(name: str, *, week: int, start_week: int, span: int,
                 force: bool, formats: str, level: str) -> List[str]:
    """Build argv for a routine, passing ONLY the flags that script accepts.

    - assets has no --level.
    - vellum/archive/digest accept only info|debug for --level (map others -> info).
    - vellum/archive/digest require both --week and --weeks.
    - week-scoped routines get (week, 1); range-scoped get (start_week, span).
    """
    restricted_level = level if level in ("info", "debug") else "info"
    if name == "publish":
        a = ["--week", str(week), "--weeks", "1", "--level", level]
        if force:
            a.append("--force")
        if formats:
            a += ["--only", formats]
        return a
    if name == "assets":  # no --level
        a = ["--week", str(week), "--weeks", "1"]
        if force:
            a.append("--force")
        return a
    if name == "vellum":  # per-week; requires --week/--weeks
        return ["--week", str(week), "--weeks", "1", "--level", restricted_level]
    if name == "archive":  # range, once
        return ["--week", str(start_week), "--weeks", str(span), "--level", restricted_level]
    if name == "digest":  # range, once
        return ["--week", str(start_week), "--weeks", str(span), "--level", restricted_level]
    raise ValueError(f"unknown routine {name!r}")


def run_routine(logger: logging.Logger, name: str, argv: List[str]) -> bool:
    """Run a routine as a subprocess, streaming its output live. True on success."""
    script = PUBLISH_DIR / SCRIPTS[name]
    if not script.exists():
        logger.error("Routine %s: script not found: %s", name, script)
        return False
    cmd = [sys.executable, "-u", str(script)] + argv
    logger.info("▶ Running %s %s", name, " ".join(argv))
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1,
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            logger.info("[%s] %s", name, line.rstrip())
        proc.wait()
        ok = (proc.returncode == 0)
        if not ok:
            logger.error("Routine %s exited with %s", name, proc.returncode)
            logger.error("Failed command: %s", " ".join(cmd))
        else:
            logger.info("✓ Finished %s", name)
        return ok
    except Exception as e:  # pragma: no cover - defensive
        logger.exception("Failed to launch %s: %s", name, e)
        return False


def _parse_targets(raw: str, kind: str) -> set:
    raw = (raw or "").strip()
    if not raw:
        return set()
    parts = [p.strip().lower() for p in raw.split(",") if p.strip()]
    unknown = [p for p in parts if p not in ALL_ROUTINES]
    if unknown:
        raise SystemExit(
            f"Unknown routine(s) in --{kind} {raw!r}: {', '.join(unknown)}. "
            f"Valid: {','.join(ALL_ROUTINES)}"
        )
    return set(parts)


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Orchestrate the weekly Democracy Clock publish routines.",
    )
    ap.add_argument("--week", type=int, help="Starting week number (use with --weeks).")
    ap.add_argument("--weeks", type=int, default=1, help="Number of consecutive weeks (default 1).")
    ap.add_argument("--week-list", help="Comma-separated week numbers (e.g. 34,35,45). Overrides --week/--weeks.")
    ap.add_argument("--force", action="store_true", help="Overwrite existing outputs.")
    ap.add_argument("--only", default="", help=f"Comma list of routines to run ({','.join(ALL_ROUTINES)}).")
    ap.add_argument("--skip", default="", help="Comma list of routines to skip.")
    ap.add_argument("--formats", default="", help="Sub-select `publish` formats (substack,scrivener,wordpress).")
    ap.add_argument("--level", choices=["debug", "info", "warning", "error"], default="info")
    args = ap.parse_args(argv)

    logger = setup_logger(args.level)

    # Resolve the week set.
    if args.week_list:
        try:
            weeks = [int(x) for x in args.week_list.split(",") if x.strip()]
        except ValueError:
            raise SystemExit("--week-list must be comma-separated integers")
        if not weeks:
            raise SystemExit("--week-list is empty")
    elif args.week is not None:
        weeks = list(range(args.week, args.week + args.weeks))
    else:
        raise SystemExit("Provide --week (with optional --weeks) or --week-list")

    only_set = _parse_targets(args.only, "only")
    skip_set = _parse_targets(args.skip, "skip")

    def want(r: str) -> bool:
        return (not only_set or r in only_set) and (r not in skip_set)

    week_routines = [r for r in WEEK_ORDER if want(r)]
    range_routines = [r for r in RANGE_ORDER if want(r)]

    start_week = min(weeks)
    span = max(weeks) - min(weeks) + 1  # contiguous span the range imports cover

    if range_routines and span != len(weeks):
        logger.warning(
            "Week list is non-contiguous (%s); range routines %s will cover the full "
            "span weeks %s–%s (including any gaps).",
            weeks, range_routines, start_week, max(weeks),
        )

    logger.info(
        "Publish controller start: weeks=%s | per-week=%s | range=%s | force=%s | formats=%s | level=%s",
        weeks, week_routines or "(none)", range_routines or "(none)",
        args.force, args.formats or "(all)", args.level,
    )

    # 1) Per-week routines.
    for w in weeks:
        if week_routines:
            logger.info("──────── Week %s publish start ────────", w)
        for r in week_routines:
            rc_argv = routine_argv(
                r, week=w, start_week=start_week, span=span,
                force=args.force, formats=args.formats, level=args.level,
            )
            if not run_routine(logger, r, rc_argv):
                logger.error("Routine %s failed for week %s; aborting.", r, w)
                return 1
        if week_routines:
            logger.info("──────── Week %s publish done ────────", w)

    # 2) Range routines (once for the whole span).
    for r in range_routines:
        rc_argv = routine_argv(
            r, week=start_week, start_week=start_week, span=span,
            force=args.force, formats=args.formats, level=args.level,
        )
        if not run_routine(logger, r, rc_argv):
            logger.error("Range routine %s failed; aborting.", r)
            return 1

    logger.info("Publish controller complete: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
