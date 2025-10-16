"""Utilities for deriving class performance summaries from overall analytics."""

from __future__ import annotations

from typing import Dict, Iterable

import pandas as pd

CATEGORY_ORDER: Iterable[str] = (
    "Appeared",
    "Passed with First Class",
    "Passed with Second Class",
    "Total Passed",
    "Yet to Pass",
)


def normalise_gender(value: str) -> str:
    token = str(value).strip().title()
    if token in {"M", "Male", "Boy", "Boys"}:
        return "Male"
    if token in {"F", "Female", "Girl", "Girls"}:
        return "Female"
    return "Unknown"


def _aggregate_counts(df: pd.DataFrame, mask: pd.Series) -> Dict[str, int]:
    subset = df.loc[mask].copy()
    boys = subset.loc[subset["gender_norm"] == "Male", "register_no"].nunique()
    girls = subset.loc[subset["gender_norm"] == "Female", "register_no"].nunique()
    total = subset["register_no"].nunique()
    return {"Boys": int(boys), "Girls": int(girls), "Total": int(total)}


def build_class_performance_summary(overall: pd.DataFrame) -> pd.DataFrame:
    """Return gender-wise class performance counts from the overall analytics table."""

    columns = ["Category", "Boys", "Girls", "Total"]

    if overall is None or overall.empty:
        return pd.DataFrame(columns=columns)

    working = overall.copy()
    if "register_no" not in working.columns:
        return pd.DataFrame(columns=columns)

    working = working[working["register_no"].notna()].copy()
    if working.empty:
        return pd.DataFrame(columns=columns)

    working["register_no"] = working["register_no"].astype(str).str.strip()
    working = working[working["register_no"] != ""]
    if working.empty:
        return pd.DataFrame(columns=columns)

    working["gender_norm"] = working.get("gender", "").map(normalise_gender)
    working["cgpa_numeric"] = pd.to_numeric(working.get("cgpa"), errors="coerce")
    working["all_clear_flag"] = working.get("all_clear", 0).fillna(0).astype(int)

    appeared_mask = working["register_no"].notna()
    first_class_mask = (working["all_clear_flag"] == 1) & (working["cgpa_numeric"] >= 7.0)
    second_class_mask = (working["all_clear_flag"] == 1) & (working["cgpa_numeric"] < 7.0)
    total_passed_mask = working["all_clear_flag"] == 1
    yet_to_pass_mask = working["all_clear_flag"] != 1

    rows = []
    masks = {
        "Appeared": appeared_mask,
        "Passed with First Class": first_class_mask,
        "Passed with Second Class": second_class_mask,
        "Total Passed": total_passed_mask,
        "Yet to Pass": yet_to_pass_mask,
    }

    for label in CATEGORY_ORDER:
        mask = masks.get(label)
        if mask is None:
            continue
        counts = _aggregate_counts(working, mask)
        counts["Category"] = label
        rows.append(counts)

    summary = pd.DataFrame(rows, columns=columns)
    return summary
