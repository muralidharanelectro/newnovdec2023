#!/usr/bin/env python3
"""Generate gender-wise class performance summary from overall analytics data."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from class_performance import build_class_performance_summary

DEFAULT_OVERALL_PATH = Path("outputs/analytics_student_overall.csv")
DEFAULT_OUTPUT_PATH = Path("outputs/class_performance_summary.csv")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--overall",
        type=Path,
        default=DEFAULT_OVERALL_PATH,
        help="Path to analytics_student_overall.csv (default: outputs/analytics_student_overall.csv)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help=(
            "Destination path for the summary table (default: "
            "outputs/class_performance_summary.csv). The format is inferred from the file extension."
        ),
    )
    return parser.parse_args()


def build_summary(overall_path: Path) -> pd.DataFrame:
    if not overall_path.exists():
        raise FileNotFoundError(f"Overall analytics file not found: {overall_path}")

    df = pd.read_csv(overall_path)
    summary = build_class_performance_summary(df)
    return summary


def write_output(summary: pd.DataFrame, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    suffix = destination.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        summary.to_excel(destination, index=False)
    else:
        summary.to_csv(destination, index=False)


def main() -> None:
    args = parse_args()
    summary = build_summary(args.overall)
    write_output(summary, args.output)
    print(f"Wrote class performance summary to: {args.output}")


if __name__ == "__main__":
    main()
