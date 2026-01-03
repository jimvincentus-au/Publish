import json
import math
from pathlib import Path
import os

import matplotlib.pyplot as plt

def load_weeks(path: str):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    weeks = data["weeks"]

    # Filter out week 0 if it’s a baseline with no window, keep if you want.
    # Here: keep all weeks that have a window end date.
    rows = []
    for w in weeks:
        window = w.get("window") or {}
        end = window.get("end")
        if not end:
            continue
        rows.append({
            "week": int(w["week"]),
            "end": end,
            "minutes": float(w["clock_after_minutes"]),
        })
    rows.sort(key=lambda r: r["week"])
    return rows

def nice_limits(values, pad=0.5):
    vmin = min(values)
    vmax = max(values)
    # Add padding in minutes (small but visible)
    lo = math.floor(vmin - pad)
    hi = math.ceil(vmax + pad)
    return lo, hi

def resolve_week_paths(week: int) -> tuple[Path, Path]:
    """Resolve (input_json_path, out_dir) for a given week number.

    Expected canonical layout:
      ../Step 3/Weeks/Week N/timeline_weekN.json
    Output is written into the same Week N folder.

    Also tolerates common folder naming variants (Week_N, Week_01, Week 01).
    """
    script_dir = Path(__file__).resolve().parent
    base_weeks_dir = (script_dir.parent / "Step 3" / "Weeks").resolve()

    # Candidate week folder names (try in order)
    candidates = [
        f"Week {week}",
        f"Week_{week}",
        f"Week {week:02d}",
        f"Week_{week:02d}",
    ]

    week_dir: Path | None = None
    for name in candidates:
        p = base_weeks_dir / name
        if p.exists() and p.is_dir():
            week_dir = p
            break

    if week_dir is None:
        raise FileNotFoundError(
            f"Could not find week directory for week={week} under {base_weeks_dir}. Tried: {', '.join(candidates)}"
        )

    input_json = week_dir / f"timeline_week{week}.json"
    if not input_json.exists():
        raise FileNotFoundError(f"Missing timeline file: {input_json}")

    return input_json, week_dir

def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="Path to timeline.json (manual mode)")
    ap.add_argument("--outdir", help="Output directory (manual mode)")

    ap.add_argument(
        "--week",
        type=int,
        help="Week number. Week mode reads ../Step 3/Weeks/Week N/timeline_weekN.json and writes charts to the same folder.",
    )
    ap.add_argument(
        "--weeks",
        type=int,
        default=1,
        help="Number of consecutive weeks to process starting at --week (default: 1).",
    )
    ap.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing chart files if present. Without --force, weeks with existing charts are skipped.",
    )
    ap.add_argument("--title", default="The Democracy Clock — Weekly Close")
    ap.add_argument("--subtitle", default="Minutes from democratic noon (higher values indicate greater democratic degradation)")
    args = ap.parse_args()

    def generate_one(input_path: Path, outdir: Path) -> None:
        rows = load_weeks(str(input_path))
        if len(rows) < 2:
            raise SystemExit(f"Not enough weeks with window.end to plot in {input_path}.")

        x = [r["week"] for r in rows]
        y = [r["minutes"] for r in rows]

        ylo, yhi = nice_limits(y, pad=1.0)

        outdir.mkdir(parents=True, exist_ok=True)

        svg_path = outdir / "democracy-clock-weekly-close.svg"
        png_path = outdir / "democracy-clock-weekly-close.png"

        if not args.force and (svg_path.exists() or png_path.exists()):
            print(f"Skip (exists; use --force): {outdir}")
            return

        # Plot
        fig = plt.figure(figsize=(12, 4.8), dpi=150)
        ax = fig.add_subplot(111)

        ax.plot(x, y, linewidth=2)  # default color; keep neutral at render-time if your style mandates
        ax.set_xlim(min(x), max(x))
        ax.set_ylim(ylo, yhi)

        ax.set_xlabel("Week")
        ax.set_ylabel("Minutes from democratic noon")

        # Title + subtitle
        ax.set_title(args.title + "\n" + args.subtitle, fontsize=12)

        # Sparse week-number tick labeling: every ~4 weeks
        first_week = rows[0]["week"]
        tick_weeks = []
        tick_labels = []
        for r in rows:
            wk = r["week"]
            if wk == first_week or (wk - first_week) % 4 == 0 or wk == rows[-1]["week"]:
                tick_weeks.append(wk)
                tick_labels.append(str(wk))
        ax.set_xticks(tick_weeks)
        ax.set_xticklabels(tick_labels, rotation=0)

        # Optional reference line at the first plotted value (baseline for visual anchoring)
        ax.axhline(y[0], linewidth=0.8, alpha=0.35)

        ax.grid(True, linewidth=0.5, alpha=0.3)

        # Small source/method line (unobtrusive)
        fig.text(
            0.99,
            0.01,
            "Source: The Democracy Clock weekly reports · Methodology available",
            ha="right",
            va="bottom",
            fontsize=8,
            alpha=0.8,
        )

        fig.tight_layout()
        fig.savefig(svg_path, format="svg")
        fig.savefig(png_path, format="png")
        plt.close(fig)

        print(f"Wrote: {svg_path}")
        print(f"Wrote: {png_path}")

    # Validate mode selection
    in_manual_mode = bool(args.input or args.outdir)
    in_week_mode = args.week is not None

    if in_manual_mode and in_week_mode:
        raise SystemExit("Choose either manual mode (--input/--outdir) OR week mode (--week/--weeks), not both.")

    if in_week_mode:
        if plt is None:
            raise SystemExit("matplotlib is required to generate charts")
        start_week = int(args.week)
        count = int(args.weeks or 1)
        if count < 1:
            raise SystemExit("--weeks must be >= 1")

        for w in range(start_week, start_week + count):
            input_path, outdir = resolve_week_paths(w)
            generate_one(input_path, outdir)
        return

    # Manual mode
    if not args.input or not args.outdir:
        raise SystemExit("Manual mode requires --input and --outdir, or use week mode with --week.")

    if plt is None:
        raise SystemExit("matplotlib is required to generate charts")

    generate_one(Path(args.input), Path(args.outdir))

if __name__ == "__main__":
    main()