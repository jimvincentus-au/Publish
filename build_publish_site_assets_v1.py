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
    if len(rows) < 2:
        raise ValueError("Timeline must contain at least two weeks.")
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
    ap.add_argument("--timeline", required=True, help="Path to global timeline.json")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--no-images-mirror", action="store_true")
    ap.add_argument(
        "--format",
        choices=["clock", "minutes"],
        default="clock",
        help="Y-axis format (default: clock)",
    )
    args = ap.parse_args()

    # Resolve --timeline: either explicit path or week number
    timeline_arg = args.timeline
    timeline_path: Path

    # Week-number shorthand
    if timeline_arg.isdigit():
        week = int(timeline_arg)
        if week < 2:
            sys.exit("--timeline week number must be >= 2")

        script_dir = Path(__file__).resolve().parent
        timeline_path = (
            script_dir.parent
            / "Step 3"
            / "Weeks"
            / f"Week {week}"
            / f"timeline_week{week}.json"
        )

    else:
        timeline_path = Path(timeline_arg)

    if not timeline_path.exists():
        sys.exit(f"Missing timeline file: {timeline_path}")

    print(f"Using timeline: {timeline_path}")

    rows = load_timeline(timeline_path)

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

    svg_path = wp_site_dir / "democracy-clock-anchor.svg"
    png_path = wp_site_dir / "democracy-clock-anchor.png"

    if not args.force and (svg_path.exists() or png_path.exists()):
        print("Anchor chart exists; use --force to overwrite.")
        return

    generate_anchor_chart(rows, svg_path, png_path, args.format)

    if not args.no_images_mirror:
        shutil.copy2(svg_path, img_dir / svg_path.name)
        shutil.copy2(png_path, img_dir / png_path.name)

    print("Anchor chart published:")
    print(f"  {svg_path}")
    print(f"  {png_path}")


if __name__ == "__main__":
    main()
