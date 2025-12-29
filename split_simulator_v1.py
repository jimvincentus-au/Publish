#!/usr/bin/env python3
"""
Split a CSV into N single-row CSVs (header + one row).

Default input:  ~/Downloads/Simulator Input.csv
Default output: ~/Downloads/Simulator/

Filename pattern (per row):
  "<COL_E> V <COL_F> <ddMMMyy>.csv"  (uppercased)

Columns E and F are positional (5th and 6th columns): index 4 and 5.
"""

from __future__ import annotations

import argparse
import csv
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

try:
    # Python 3.9+
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None  # type: ignore


INVALID_FS_CHARS = r'<>:"/\\|?*'  # common Windows-invalid; safe on macOS too


def sanitize_for_filename(s: str) -> str:
    """Make a conservative, filesystem-safe token."""
    s = (s or "").strip()
    # Replace invalid characters with space
    s = re.sub(f"[{re.escape(INVALID_FS_CHARS)}]", " ", s)
    # Replace any other non-printable/control-ish chars with space
    s = re.sub(r"[\x00-\x1f\x7f]", " ", s)
    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s.upper()


def today_ddMMMyy(tz_name: str) -> str:
    """Return date like 25DEC25 in the requested timezone (default Brisbane)."""
    if ZoneInfo is not None:
        try:
            now = datetime.now(ZoneInfo(tz_name))
        except Exception:
            now = datetime.now()
    else:
        now = datetime.now()
    return now.strftime("%d%b%y").upper()


@dataclass
class SplitStats:
    rows_written: int = 0
    rows_skipped_empty: int = 0


def split_csv(input_path: Path, output_dir: Path, tz_name: str = "Australia/Brisbane") -> SplitStats:
    if not input_path.exists():
        raise FileNotFoundError(f"Input CSV not found: {input_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    date_stamp = today_ddMMMyy(tz_name)
    stats = SplitStats()

    # utf-8-sig handles BOM if present
    with input_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        try:
            header = next(reader)
        except StopIteration:
            return stats  # empty file

        if len(header) < 6:
            raise ValueError(
                f"CSV has only {len(header)} columns; need at least 6 to use columns E and F."
            )

        for i, row in enumerate(reader, start=1):
            # Skip completely blank lines
            if not row or all((c or "").strip() == "" for c in row):
                stats.rows_skipped_empty += 1
                continue

            # Pad short rows so indexes 4 and 5 exist
            if len(row) < len(header):
                row = row + [""] * (len(header) - len(row))

            col_e = sanitize_for_filename(row[4])
            col_f = sanitize_for_filename(row[5])

            if not col_e:
                col_e = f"ROW{i}_E"
            if not col_f:
                col_f = f"ROW{i}_F"

            filename = f"{col_e} V {col_f} {date_stamp}.csv"
            out_path = output_dir / filename

            with out_path.open("w", encoding="utf-8", newline="") as out_f:
                writer = csv.writer(out_f)
                writer.writerow(header)
                writer.writerow(row)

            stats.rows_written += 1

    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description="Split Simulator Input CSV into per-row CSV files.")
    parser.add_argument(
        "--input",
        default=str(Path.home() / "Downloads" / "Simulator Input.csv"),
        help="Input CSV path (default: ~/Downloads/Simulator Input.csv)",
    )
    parser.add_argument(
        "--outdir",
        default=str(Path.home() / "Downloads" / "Simulator"),
        help="Output directory (default: ~/Downloads/Simulator/)",
    )
    parser.add_argument(
        "--tz",
        default="Australia/Brisbane",
        help="Timezone for date stamp (default: Australia/Brisbane)",
    )

    args = parser.parse_args()
    stats = split_csv(Path(args.input).expanduser(), Path(args.outdir).expanduser(), tz_name=args.tz)

    print(f"Done. Wrote {stats.rows_written} file(s) to: {Path(args.outdir).expanduser()}")
    if stats.rows_skipped_empty:
        print(f"Skipped {stats.rows_skipped_empty} empty line(s).")


if __name__ == "__main__":
    main()