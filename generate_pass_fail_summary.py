#!/usr/bin/env python3
"""Generate gender/community wise pass & reappear counts from result data."""

from __future__ import annotations

import argparse
import os
from typing import Dict, Iterable, List

import pandas as pd


def normalize_token(value: str) -> str:
    """Return a stripped uppercase token for *value* (handles NaNs)."""

    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip().upper()


def normalize_register(value: str) -> str:
    token = normalize_token(value)
    if token.endswith(".0"):
        token = token[:-2]
    return token


def normalize_gender(value: str) -> str:
    token = normalize_token(value)
    if token in {"M", "MALE", "BOY", "BOYS"}:
        return "Male"
    if token in {"F", "FEMALE", "GIRL", "GIRLS"}:
        return "Female"
    return "Unknown"


def normalize_community(value: str) -> str:
    token = normalize_token(value).replace(".", "")
    if not token:
        return "Unknown"
    synonyms: Dict[str, str] = {
        "BC": "BC",
        "MBC": "MBC",
        "OBC": "OBC",
        "SC": "SC",
        "ST": "ST",
        "SCA": "SCA",
        "OC": "OC",
        "OTHERS": "Others",
        "GENERAL": "OC",
    }
    return synonyms.get(token, token)


def load_biodata(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str, engine="openpyxl")
    if df.empty:
        return pd.DataFrame(columns=["register_no", "gender", "community"])

    register_col = df.columns[0]
    gender_col = df.columns[1]
    community_col = df.columns[2]

    out = pd.DataFrame({
        "register_no": df[register_col].map(normalize_register),
        "gender": df[gender_col].map(normalize_gender),
        "community": df[community_col].map(normalize_community),
    })

    out.loc[out["register_no"] == "", ["gender", "community"]] = "Unknown"
    return out


def locate_column(columns: Iterable[str], *aliases: str) -> str:
    candidates = [str(col).strip() for col in columns]
    for alias in aliases:
        alias_norm = alias.strip().upper()
        for orig, cand in zip(columns, candidates):
            if alias_norm == cand.upper():
                return orig
    for orig, cand in zip(columns, candidates):
        if "REGISTER" in cand.upper() and "NUMBER" in cand.upper():
            return orig
    return ""


def load_results(path: str) -> pd.DataFrame:
    xl = pd.ExcelFile(path, engine="openpyxl")
    frames: List[pd.DataFrame] = []

    for sheet in xl.sheet_names:
        df = xl.parse(sheet, dtype=str)
        if df.empty:
            continue

        reg_col = locate_column(df.columns, "REGISTER NUMBER")
        result_col = locate_column(df.columns, "RESULT")
        if not reg_col or not result_col:
            continue

        frame = pd.DataFrame({
            "register_no": df[reg_col].map(normalize_register),
            "raw_result": df[result_col].map(normalize_token),
        })
        frame = frame[frame["register_no"].str.fullmatch(r"\d+")]
        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=["register_no", "status"])

    combined = pd.concat(frames, ignore_index=True)

    status_map = {
        "PASS": "All Clear",
        "REAPPEAR": "Reappear",
        "ABSENT": "Reappear",
    }
    combined["status"] = combined["raw_result"].map(status_map).fillna("Unknown")

    def reduce_status(group: pd.Series) -> str:
        if (group == "Reappear").any():
            return "Reappear"
        if (group == "All Clear").any():
            return "All Clear"
        return "Unknown"

    return (
        combined.groupby("register_no", as_index=False)["status"]
        .agg(reduce_status)
    )


def build_summary(results: pd.DataFrame, biodata: pd.DataFrame) -> pd.DataFrame:
    merged = results.merge(biodata, on="register_no", how="left")
    merged["gender"] = merged["gender"].fillna("Unknown")
    merged["community"] = merged["community"].fillna("Unknown")

    rows: List[Dict[str, str]] = []

    for status, count in merged.groupby("status").size().items():
        rows.append({
            "view": "Status totals",
            "status": status,
            "gender": "All",
            "community": "All",
            "count": int(count),
        })

    for (status, gender), count in merged.groupby(["status", "gender"]).size().items():
        rows.append({
            "view": "Status by gender",
            "status": status,
            "gender": gender,
            "community": "All",
            "count": int(count),
        })

    for (status, community), count in merged.groupby(["status", "community"]).size().items():
        rows.append({
            "view": "Status by community",
            "status": status,
            "gender": "All",
            "community": community,
            "count": int(count),
        })

    for (status, gender, community), count in (
        merged.groupby(["status", "gender", "community"]).size().items()
    ):
        rows.append({
            "view": "Status by gender & community",
            "status": status,
            "gender": gender,
            "community": community,
            "count": int(count),
        })

    return pd.DataFrame(rows)


def parse_args(argv: List[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--biodata",
        default="biodata.xlsx",
        help="Path to the biodata workbook (default: biodata.xlsx)",
    )
    parser.add_argument(
        "--results",
        default="ALL SEM RESULTS.xlsx",
        help="Path to the consolidated results workbook",
    )
    parser.add_argument(
        "--output",
        default="biodata_pass_fail_summary.csv",
        help="Destination CSV path for the summary table",
    )
    return parser.parse_args(argv)


def main(argv: List[str] | None = None) -> None:
    args = parse_args(argv)

    if not os.path.exists(args.results):
        raise SystemExit(f"Results workbook not found: {args.results}")
    if not os.path.exists(args.biodata):
        raise SystemExit(f"Biodata workbook not found: {args.biodata}")

    biodata = load_biodata(args.biodata)
    results = load_results(args.results)
    if results.empty:
        raise SystemExit("No valid register numbers found in results workbook")

    summary = build_summary(results, biodata)
    summary.to_csv(args.output, index=False)


if __name__ == "__main__":
    main()
