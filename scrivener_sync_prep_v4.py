#!/usr/bin/env python3
"""
scrivener_sync_prep_v4.py

Copy Democracy Clock automation outputs from the Step 5 output tree
into the Scrivener External Folder Sync directory, in a SAFE way.

Defaults:
  --source-root = /Volumes/PRINTIFY24/Democracy Clock Automation/Publish/Output/Scrivener
  --dest-root   = /Volumes/PRINTIFY24/Democracy Clock Automation/Publish/Scrivener Sync
  --weeks       = all
"""

import argparse
import re
import shutil
from pathlib import Path
from typing import List, Set


# Default paths
DEFAULT_SOURCE_ROOT = Path("/Volumes/PRINTIFY24/Democracy Clock Automation/Publish/Output/Scrivener")
DEFAULT_DEST_ROOT   = Path("/Volumes/PRINTIFY24/Democracy Clock Automation/Publish/Scrivener Sync")

VALID_EXTS: Set[str] = {".md", ".txt"}


def parse_week_selection(weeks_arg: str, available_weeks: List[int]) -> List[int]:
    """Parse 'all', '1', '1,3-5', etc."""
    if weeks_arg.strip().lower() == "all":
        return sorted(available_weeks)

    selected: Set[int] = set()
    for part in weeks_arg.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            try:
                start = int(start_s)
                end = int(end_s)
            except ValueError:
                raise ValueError(f"Invalid week range: {part!r}")
            if start > end:
                start, end = end, start
            for w in range(start, end + 1):
                selected.add(w)
        else:
            try:
                selected.add(int(part))
            except ValueError:
                raise ValueError(f"Invalid week number: {part!r}")

    return sorted(w for w in selected if w in available_weeks)


def discover_weeks(source_root: Path) -> List[int]:
    """Look for Week N folders."""
    weeks: Set[int] = set()
    pattern = re.compile(r"^Week\s+(\d+)$", re.IGNORECASE)

    for child in source_root.iterdir():
        if child.is_dir():
            m = pattern.match(child.name)
            if m:
                weeks.add(int(m.group(1)))

    return sorted(weeks)


def copy_week(week_number: int, src_root: Path, dest_root: Path,
              incoming_dir_name: str, overwrite: bool = False) -> None:
    """Copy *.md and *.txt safely."""
    src_week_dir = src_root / f"Week {week_number}"
    if not src_week_dir.exists():
        print(f"[WARN] Week {week_number}: source folder missing: {src_week_dir}")
        return

    dest_week_dir = dest_root / incoming_dir_name / f"Week {week_number}"
    dest_week_dir.mkdir(parents=True, exist_ok=True)

    copied = 0
    skipped = 0

    for item in src_week_dir.iterdir():
        if item.is_file() and item.suffix.lower() in VALID_EXTS:
            dest_file = dest_week_dir / item.name
            if dest_file.exists() and not overwrite:
                skipped += 1
                continue
            shutil.copy2(item, dest_file)
            copied += 1

    print(f"[INFO] Week {week_number}: copied {copied}, skipped {skipped} → {dest_week_dir}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Prepare Scrivener Sync folder with Democracy Clock automation output."
    )
    parser.add_argument(
        "--source-root",
        default=str(DEFAULT_SOURCE_ROOT),
        help="Root with Week N folders (default: DC Automation/Publish/Output/Scrivener).",
    )
    parser.add_argument(
        "--dest-root",
        default=str(DEFAULT_DEST_ROOT),
        help="Scrivener Sync root (default: DC Automation/Publish/Scrivener Sync).",
    )
    parser.add_argument(
        "--weeks",
        default="all",
        help="Weeks to copy: 'all' or '1,3-5'. Default = all."
    )
    parser.add_argument(
        "--incoming-dir",
        default="_incoming_from_automation",
        help="Incoming subfolder under dest-root (default: _incoming_from_automation).",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite files already present."
    )

    args = parser.parse_args()

    src_root = Path(args.source_root).expanduser().resolve()
    dest_root = Path(args.dest_root).expanduser().resolve()

    if not src_root.exists():
        raise SystemExit(f"[ERROR] Source root does not exist: {src_root}")
    if not dest_root.exists():
        raise SystemExit(f"[ERROR] Destination root does not exist: {dest_root}")

    draft_dir = dest_root / "Draft"
    if not draft_dir.exists():
        print(f"[WARN] Destination does not contain Draft/. Is this your Scrivener Sync folder?\n{dest_root}")

    available_weeks = discover_weeks(src_root)
    if not available_weeks:
        raise SystemExit(f"[ERROR] No week folders found in {src_root}")

    try:
        selected_weeks = parse_week_selection(args.weeks, available_weeks)
    except ValueError as e:
        raise SystemExit(f"[ERROR] {e}")

    print(f"[INFO] Using source: {src_root}")
    print(f"[INFO] Using destination: {dest_root}")
    print(f"[INFO] Weeks to copy: {selected_weeks}")
    print(f"[INFO] Incoming dir: {args.incoming_dir}")
    print(f"[INFO] Overwrite: {args.overwrite}")

    for week in selected_weeks:
        copy_week(week, src_root, dest_root, args.incoming_dir, args.overwrite)

    print("[INFO] Completed. Open Scrivener → Sync to import new files into Research.")


if __name__ == "__main__":
    main()