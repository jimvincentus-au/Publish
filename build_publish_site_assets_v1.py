#!/usr/bin/env python3
"""
build_publish_site_assets_v1.py

Generates project-level (site-wide) publish assets for The Democracy Clock.
Currently emits the canonical anchor chart (SVG + PNG).

Reads from Step 3. Writes only to Publish/Output.
"""

from pathlib import Path
import json
import argparse
import shutil
import sys

import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter


# ------------------------
# Helpers
# ------------------------

def minutes_to_clock_label(minutes: float) -> str:
    """Convert minutes-after-noon to clock time label (e.g., 491.2 -> 8:11 p.m.)."""
    total_minutes = int(round(minutes))

    # Minutes-after-noon: 0 = 12:00 p.m.
    clock_minutes = (12 * 60 + total_minutes) % (24 * 60)

    hour24 = clock_minutes // 60
    mins = clock_minutes % 60

    period = "a.m." if hour24 < 12 else "p.m."
    hour12 = hour24 % 12
    hour12 = 12 if hour12 == 0 else hour12

    return f"{hour12}:{mins:02d} {period}"


def load_timeline(path: Path):
    payload = json.loads(path.read_text(encoding="utf-8"))
    weeks = payload.get("weeks", [])
    rows = []
    for w in weeks:
        if "week" in w and "clock_after_minutes" in w:
            rows.append(
                {
                    "week": int(w["week"]),
                    "minutes": float(w["clock_after_minutes"]),
                    "big": int(w.get("big_moves_count", 0)),
                }
            )
    rows.sort(key=lambda r: r["week"])
    if len(rows) < 1:
        raise ValueError("Timeline must contain at least one week.")
    return rows


# ------------------------
# Chart generation
# ------------------------

def generate_anchor_chart(rows, out_svg: Path, out_png: Path, y_mode: str):
    weeks = [r["week"] for r in rows]
    minutes = [r["minutes"] for r in rows]

    fig = plt.figure(figsize=(12, 5), dpi=150)
    ax = fig.add_subplot(111)

    # Anchor y-axis at Week 0 (prevent implied negative weeks)
    ax.spines["left"].set_position(("data", 0))
    ax.spines["right"].set_color("none")
    ax.spines["top"].set_color("none")

    ax.yaxis.set_ticks_position("left")
    ax.xaxis.set_ticks_position("bottom")

    ax.plot(weeks, minutes, linewidth=2)

    # Baseline (Week 1)
    ax.axhline(minutes[0], linewidth=0.8, alpha=0.35)

    # Highlight latest
    ax.scatter(weeks[-1], minutes[-1], s=50)

    # Big move markers
    for r in rows:
        if r["big"] > 0:
            ax.scatter(r["week"], r["minutes"], s=30)

    ax.set_xlabel("Week")

    if y_mode == "clock":
        ax.set_ylabel("Democracy Clock Time")
        ax.yaxis.set_major_formatter(
            FuncFormatter(lambda y, _: minutes_to_clock_label(y))
        )
    else:
        ax.set_ylabel("Minutes from democratic noon")

    ax.set_title(
        "The Democracy Clock — Project Timeline\n"
        "Week-by-week measurement of democratic condition"
    )

    ax.grid(True, linewidth=0.5, alpha=0.3)

    fig.text(
        0.99,
        0.01,
        "Source: The Democracy Clock · Weekly reports · Methodology available",
        ha="right",
        va="bottom",
        fontsize=8,
        alpha=0.8,
    )

    fig.tight_layout()
    fig.savefig(out_svg, format="svg")
    fig.savefig(out_png, format="png")
    plt.close(fig)


# ------------------------
# Main
# ------------------------

def main():
    ap = argparse.ArgumentParser()
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--timeline", help="Path to global timeline.json, or a week number shorthand")
    group.add_argument("--week", type=int, help="Start week number for batch generation")
    ap.add_argument("--weeks", type=int, default=1, help="Number of weeks to generate (used with --week)")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--no-images-mirror", action="store_true")
    ap.add_argument(
        "--format",
        choices=["clock", "minutes"],
        default="clock",
        help="Y-axis format (default: clock)",
    )
    args = ap.parse_args()

    def resolve_timeline_for_week(week: int) -> Path:
        script_dir = Path(__file__).resolve().parent
        return (
            script_dir.parent
            / "Step 3"
            / "Weeks"
            / f"Week {week}"
            / f"timeline_week{week}.json"
        )

    publish_root = Path(__file__).resolve().parent

    wp_site_dir = (
        publish_root
        / "Output"
        / "Wordpress"
        / "_site"
    )
    img_dir = (
        publish_root
        / "Output"
        / "Images"
    )

    wp_site_dir.mkdir(parents=True, exist_ok=True)
    img_dir.mkdir(parents=True, exist_ok=True)

    def publish_one(rows_local, svg_out: Path, png_out: Path):
        if not args.force and (svg_out.exists() or png_out.exists()):
            print(f"Anchor chart exists; skipping (use --force to overwrite): {svg_out.name}")
            return
        generate_anchor_chart(rows_local, svg_out, png_out, args.format)
        if not args.no_images_mirror:
            shutil.copy2(svg_out, img_dir / svg_out.name)
            shutil.copy2(png_out, img_dir / png_out.name)
        print("Anchor chart published:")
        print(f"  {svg_out}")
        print(f"  {png_out}")

    if args.timeline is not None:
        timeline_arg = args.timeline
        if timeline_arg.isdigit():
            week = int(timeline_arg)
            if week < 1:
                sys.exit("--timeline week number must be >= 1")
            timeline_path = resolve_timeline_for_week(week)
        else:
            timeline_path = Path(timeline_arg)

        if not timeline_path.exists():
            sys.exit(f"Missing timeline file: {timeline_path}")

        print(f"Using timeline: {timeline_path}")
        rows = load_timeline(timeline_path)

        svg_path = wp_site_dir / "democracy-clock-anchor.svg"
        png_path = wp_site_dir / "democracy-clock-anchor.png"
        publish_one(rows, svg_path, png_path)
        return

    # Batch mode: --week/--weeks
    start_week = args.week
    if start_week is None:
        sys.exit("Internal error: --week missing in batch mode")
    if args.weeks < 1:
        sys.exit("--weeks must be >= 1")

    end_week = start_week + args.weeks - 1
    for wk in range(start_week, end_week + 1):
        timeline_path = resolve_timeline_for_week(wk)
        if not timeline_path.exists():
            sys.exit(f"Missing timeline file: {timeline_path}")
        rows_local = load_timeline(timeline_path)

        svg_name = f"democracy_clock_anchor_week{wk:02d}.svg"
        png_name = f"democracy_clock_anchor_week{wk:02d}.png"
        svg_path = wp_site_dir / svg_name
        png_path = wp_site_dir / png_name

        print(f"Using timeline (week {wk}): {timeline_path}")
        publish_one(rows_local, svg_path, png_path)


if __name__ == "__main__":
    main()
