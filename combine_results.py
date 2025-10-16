#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse
import json
import os
import re
from typing import Dict, Iterable, List, Optional, Set, Tuple

import numpy as np
import pandas as pd

from class_performance import build_class_performance_summary

HERE = os.path.dirname(os.path.abspath(__file__))

# Grade tokens that indicate the student actually ATTENDED the arrear for that subject
PASS_TOKENS = {"o", "a+", "a", "b+", "b", "c", "s", "p", "pass", "passed"}
FAIL_TOKENS = {
    "u",
    "ua",
    "ra",
    "f",
    "fail",
    "failed",
    "np",
    "absent",
    "ab",
    "wh",
    "wh1",
    "nr",
    "w",
    "i",
}
ATTENDED_TOKENS = PASS_TOKENS | FAIL_TOKENS

NULL_TOKENS = {"", "-", "na", "n/a", "null", "none", "nan"}  # anything else is ignored

GRADE_POINT_MAP = {
    "O": 10,
    "A+": 9,
    "A": 8,
    "B+": 7,
    "B": 6,
    "C": 5,
}

PASSING_GRADE_MIN = 5


def normalize_grade_token(value: str) -> str:
    """Return an upper-cased, whitespace-free grade token."""

    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    token = str(value).strip().upper()
    token = re.sub(r"\s+", "", token)
    return token


def grade_to_points(value: str) -> int:
    token = normalize_grade_token(value)
    if not token:
        return 0
    return int(GRADE_POINT_MAP.get(token, 0))


def first_non_empty(values: Iterable) -> str:
    """Return the first non-null / non-empty string from *values*."""

    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def last_non_empty(values: Iterable) -> str:
    """Return the last non-null / non-empty string from *values*."""

    last = ""
    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if text:
            last = text
    return last

def load_config(path: str) -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def norm(s: str) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = str(s).replace("\u00A0"," ").strip().lower()  # remove NBSPs
    s = re.sub(r"\s+", " ", s)
    return s

def is_attended_grade(v) -> bool:
    t = norm(v)
    if t in NULL_TOKENS or t == "":
        return False
    # Strict: count only known grade/result tokens
    return t in ATTENDED_TOKENS

def to_result_from_grade(grade: str, fallback_result: str) -> str:
    g = norm(grade)
    r = norm(fallback_result)
    if g in PASS_TOKENS: return "Pass"
    if g in FAIL_TOKENS: return "Fail"
    if r in {"pass","p","passed"}: return "Pass"
    if r in {"fail","f","failed"}: return "Fail"
    return fallback_result

def read_all_sheets(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(xlsx_path, engine="openpyxl")
    return {sh: xl.parse(sh, dtype=str) for sh in xl.sheet_names}

def detect_semester_from_sheet(sheet_name: str) -> int:
    m = re.search(r"\bsem(?:ester)?\s*(\d+)\b", str(sheet_name).lower())
    return int(m.group(1)) if m else -1

def find_col(df: pd.DataFrame, aliases: List[str], regex: str=None) -> str:
    cols = list(df.columns)
    ncols = [norm(c) for c in cols]
    for alias in aliases:
        a = norm(alias)
        for i, c in enumerate(ncols):
            if a == c or a in c:
                return cols[i]
    if regex:
        pat = re.compile(regex, flags=re.I)
        for i, c in enumerate(cols):
            if pat.search(str(cols[i]).strip()):
                return cols[i]
    return ""

def find_subject_columns(df: pd.DataFrame, subject_header_regex: str) -> List[Tuple[str, str, float]]:
    out = []
    pat = re.compile(subject_header_regex, flags=re.I)
    for col in df.columns:
        m = pat.match(str(col).strip())
        if m:
            code = m.group(1).strip().upper()
            try:
                credit = float(m.group(2))
            except Exception:
                credit = np.nan
            out.append((col, code, credit))
    return out


def load_subject_catalog(path: str) -> Dict[str, Dict[str, object]]:
    """Load subject metadata (name, credit, semester) keyed by subject code."""

    if not path or not os.path.isfile(path):
        return {}

    try:
        catalog_df = pd.read_excel(path, dtype=str, engine="openpyxl")
    except Exception:
        return {}

    if catalog_df is None or catalog_df.empty:
        return {}

    code_col = find_col(catalog_df, ["subject code", "code"], r"code")
    if not code_col:
        return {}

    name_col = find_col(catalog_df, ["subject name", "name", "course name"], r"name")
    credit_col = find_col(catalog_df, ["credit", "credits", "credit points"], r"credit")
    sem_col = find_col(catalog_df, ["semester", "sem"], r"sem")

    catalog: Dict[str, Dict[str, object]] = {}

    for _, row in catalog_df.iterrows():
        code_raw = row.get(code_col, "")
        code = str(code_raw).strip().upper()
        if not code:
            continue

        name = str(row.get(name_col, "")).strip() if name_col else ""

        credit_val = row.get(credit_col) if credit_col else np.nan
        try:
            credit = float(credit_val)
        except (TypeError, ValueError):
            credit = np.nan

        sem_val = row.get(sem_col) if sem_col else np.nan
        try:
            semester = int(float(sem_val))
        except (TypeError, ValueError):
            semester = np.nan

        catalog[code] = {
            "name": name,
            "credits": credit,
            "semester": semester,
        }

    return catalog


def parse_vertical_student_sheet(
    workbook_path: str,
    sheet_name: str,
    current_semester: int,
    subject_catalog: Dict[str, Dict[str, object]],
    default_exam_session: str,
) -> Tuple[pd.DataFrame, List[Dict[str, str]]]:
    """Parse single-student sheets laid out vertically (semester, subject, grade)."""

    issues: List[Dict[str, str]] = []

    try:
        df = pd.read_excel(
            workbook_path,
            sheet_name=sheet_name,
            header=None,
            dtype=str,
            engine="openpyxl",
        )
    except Exception as exc:
        issues.append({"sheet": sheet_name, "issue": f"Unable to read sheet: {exc}"})
        return pd.DataFrame(), issues

    if df is None or df.empty:
        issues.append({"sheet": sheet_name, "issue": "Empty sheet"})
        return pd.DataFrame(), issues

    reg_no = first_non_empty(df.iloc[0].tolist()) if len(df.index) > 0 else ""
    if not reg_no:
        issues.append({"sheet": sheet_name, "issue": "Missing register number in A1"})
        return pd.DataFrame(), issues

    student_name = first_non_empty(df.iloc[1].tolist()) if len(df.index) > 1 else ""

    rows: List[Dict[str, object]] = []
    missing_subjects: Set[str] = set()
    missing_credits: Set[str] = set()

    for idx in range(2, len(df.index)):
        row = df.iloc[idx]
        sem_raw = first_non_empty([row.iloc[0]]) if len(row) > 0 else ""
        subj_raw = first_non_empty([row.iloc[1]]) if len(row) > 1 else ""
        grade_raw = first_non_empty([row.iloc[2]]) if len(row) > 2 else ""

        if not sem_raw and not subj_raw and not grade_raw:
            continue

        sem_text = str(sem_raw).strip()
        if not sem_text:
            if not subj_raw and not grade_raw:
                continue
        if sem_text.lower().startswith("sem"):
            continue

        try:
            semester = int(float(sem_text))
        except (TypeError, ValueError):
            issues.append({
                "sheet": sheet_name,
                "issue": f"Unrecognised semester value '{sem_raw}' at row {idx + 1}",
            })
            continue

        subject_code = str(subj_raw).strip().upper()
        if not subject_code:
            issues.append({
                "sheet": sheet_name,
                "issue": f"Missing subject code at row {idx + 1}",
            })
            continue

        if str(subj_raw).strip().lower().startswith("sub") and str(grade_raw).strip().lower().startswith("grade"):
            continue

        grade_token = normalize_grade_token(grade_raw)
        if not grade_token:
            issues.append({
                "sheet": sheet_name,
                "issue": f"Missing grade for subject {subject_code} at row {idx + 1}",
            })
            continue

        subject_meta = subject_catalog.get(subject_code, {})
        if not subject_meta:
            missing_subjects.add(subject_code)

        credit_val: Optional[float] = subject_meta.get("credits") if subject_meta else np.nan
        try:
            credit = float(credit_val)
        except (TypeError, ValueError):
            credit = np.nan
        if np.isnan(credit):
            missing_credits.add(subject_code)

        catalog_sem = subject_meta.get("semester") if subject_meta else np.nan
        try:
            catalog_semester = int(catalog_sem)
        except (TypeError, ValueError):
            catalog_semester = np.nan

        grade_point = grade_to_points(grade_token)
        result = "Pass" if grade_point >= PASSING_GRADE_MIN else "Fail"

        rows.append(
            {
                "register_no": str(reg_no).strip(),
                "student_name": student_name,
                "semester": semester,
                "subject_code": subject_code,
                "subject_name": subject_meta.get("name", "") if subject_meta else "",
                "credit": credit,
                "grade": grade_token,
                "grade_point": grade_point,
                "result": result,
                "cgpa": "",
                "no_of_subjects_reappear": "",
                "exam_session": default_exam_session,
                "is_arrear": "Y" if (current_semester and semester < current_semester) else "N",
                "source_sheet": sheet_name,
                "catalog_semester": catalog_semester,
            }
        )

    if not rows:
        issues.append({"sheet": sheet_name, "issue": "No subject attempts detected"})
        return pd.DataFrame(), issues

    if missing_subjects:
        issues.append({
            "sheet": sheet_name,
            "issue": "Missing subjects in catalog: " + ", ".join(sorted(missing_subjects)),
        })
    elif missing_credits:
        issues.append({
            "sheet": sheet_name,
            "issue": "Missing credits for subjects: " + ", ".join(sorted(missing_credits)),
        })

    student_df = pd.DataFrame(rows)
    student_df["credit"] = pd.to_numeric(student_df["credit"], errors="coerce")
    student_df["grade_point"] = pd.to_numeric(student_df["grade_point"], errors="coerce")

    gpa_mask = (
        (student_df["semester"] == current_semester)
        & student_df["credit"].notna()
        & (student_df["credit"] > 0)
        & (student_df["grade_point"] >= PASSING_GRADE_MIN)
    )

    if "catalog_semester" in student_df.columns:
        gpa_mask &= student_df["catalog_semester"].isna() | (
            student_df["catalog_semester"] == current_semester
        )

    total_credits = student_df.loc[gpa_mask, "credit"].sum()
    if total_credits > 0:
        weighted = (student_df.loc[gpa_mask, "credit"] * student_df.loc[gpa_mask, "grade_point"]).sum()
        gpa_value = round(weighted / total_credits, 2)
        gpa_str = f"{gpa_value:.2f}"
    else:
        gpa_str = ""

    student_df["gpa"] = ""
    if gpa_str:
        student_df.loc[student_df["semester"] == current_semester, "gpa"] = gpa_str

    return student_df, issues

def valid_register(s: str, value_regex: str) -> bool:
    t = str(s).strip()
    if t == "" or norm(t) in NULL_TOKENS: return False
    if "semester" in norm(t): return False
    return bool(re.fullmatch(value_regex, t))

def load_biodata(path: str) -> pd.DataFrame:
    """Load register-to-gender/community mapping from *path* if present."""

    if not path or not os.path.isfile(path):
        return pd.DataFrame(columns=["register_no", "gender", "community"])

    df_raw = pd.read_excel(path, dtype=str, engine="openpyxl")
    if df_raw.empty:
        return pd.DataFrame(columns=["register_no", "gender", "community"])

    reg_col = find_col(df_raw, ["register number", "register no", "reg no"], r"^reg(ister)?")
    gender_col = find_col(df_raw, ["gender", "sex"], r"^(gender|sex)\b")
    community_col = find_col(df_raw, ["community", "category", "caste"], r"^(community|category|caste)\b")

    cols_needed = [reg_col, gender_col, community_col]
    if any(not c for c in cols_needed):
        missing = [lbl for lbl, col in zip(["register", "gender", "community"], cols_needed) if not col]
        raise SystemExit(
            "Unable to locate required columns in biodata workbook: {}".format(
                ", ".join(missing)
            )
        )

    def normalize_gender(value: str) -> str:
        token = norm(value)
        if token in {"m", "male", "boy", "boys"}:
            return "Male"
        if token in {"f", "female", "girl", "girls"}:
            return "Female"
        return "Unknown"

    def normalize_community(value: str) -> str:
        token = norm(value).replace(".", "")
        if not token:
            return "Unknown"
        token = token.upper()
        synonyms = {
            "OBC": "OBC",
            "BC": "BC",
            "MBC": "MBC",
            "SC": "SC",
            "ST": "ST",
            "SCA": "SCA",
            "OC": "OC",
            "OTHERS": "Others",
        }
        if token in synonyms:
            return synonyms[token]
        # handle composite strings like "OBC / Others"
        for key in synonyms:
            if key in token:
                return synonyms[key]
        if "OTH" in token:
            return "Others"
        return token

    df = pd.DataFrame({
        "register_no": df_raw[reg_col].astype(str).str.strip(),
        "gender": df_raw[gender_col].map(normalize_gender),
        "community": df_raw[community_col].map(normalize_community),
    })

    df = df[df["register_no"].astype(str).str.strip() != ""].copy()
    df["register_no"] = df["register_no"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["register_no"], keep="last")
    return df


def compute_student_semester_summary(master: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        return pd.DataFrame(columns=[
            "register_no",
            "semester",
            "subjects_attempted",
            "subjects_passed",
            "is_arrear_semester",
            "gpa",
            "cgpa",
        ])

    summary = (
        master.groupby(["register_no", "semester"], dropna=False)
        .agg(
            subjects_attempted=("subject_code", "count"),
            subjects_passed=("pass_flag", "sum"),
            is_arrear_semester=("is_arrear", lambda s: "Y" if (s == "Y").any() else "N"),
            gpa=("gpa", first_non_empty),
            cgpa=("cgpa", first_non_empty),
        )
        .reset_index()
    )
    return summary


def _latest_regular_gpa_cgpa(group: pd.DataFrame) -> pd.Series:
    """Return the latest regular-semester GPA/CGPA for *group* without altering values."""

    working = group.copy()
    working["_semester_sort"] = pd.to_numeric(working.get("semester"), errors="coerce").fillna(-1)

    regular = working[working.get("is_arrear") == "N"]
    if regular.empty:
        regular = working

    regular = regular.sort_values([
        "_semester_sort",
        "is_arrear",
    ], ascending=[True, True], na_position="last")

    gpa_val = last_non_empty(regular.get("gpa", []))
    cgpa_val = last_non_empty(regular.get("cgpa", []))

    return pd.Series({"gpa": gpa_val, "cgpa": cgpa_val})


def compute_student_overall_summary(master: pd.DataFrame, current_semester: int) -> pd.DataFrame:
    if master.empty:
        return pd.DataFrame(columns=[
            "register_no",
            "total_subjects_attempted",
            "total_subjects_passed",
            "arrear_subjects_attempted",
            "current_semester_attempts",
            "all_clear",
            "gender",
            "community",
            "gpa",
            "cgpa",
        ])

    master = master.copy()

    sort_cols = []
    ascending = []

    if "register_no" in master.columns:
        sort_cols.append("register_no")
        ascending.append(True)

    if "semester" in master.columns:
        master["_semester_sort"] = pd.to_numeric(master["semester"], errors="coerce").fillna(-1)
        sort_cols.append("_semester_sort")
        ascending.append(True)

    if "is_arrear" in master.columns:
        sort_cols.append("is_arrear")
        ascending.append(False)

    if sort_cols:
        master = master.sort_values(sort_cols, ascending=ascending, na_position="first")

    if "_semester_sort" in master.columns:
        master = master.drop(columns=["_semester_sort"])

    summary = (
        master.groupby("register_no")
        .agg(
            total_subjects_attempted=("subject_code", "count"),
            total_subjects_passed=("pass_flag", "sum"),
            arrear_subjects_attempted=("is_arrear", lambda s: int((s == "Y").sum())),
            current_semester_attempts=("is_current_semester", lambda s: int((s == "Y").sum())),
            gender=("gender", first_non_empty),
            community=("community", first_non_empty),
        )
        .reset_index()
    )

    latest_values = (
        master.groupby("register_no", group_keys=False)
        .apply(_latest_regular_gpa_cgpa, include_groups=False)
        .reset_index()
    )

    summary = summary.merge(latest_values, on="register_no", how="left")

    summary["total_subjects_attempted"] = summary["total_subjects_attempted"].fillna(0).astype(int)
    summary["total_subjects_passed"] = summary["total_subjects_passed"].fillna(0).astype(int)
    summary["arrear_subjects_attempted"] = summary["arrear_subjects_attempted"].fillna(0).astype(int)
    summary["current_semester_attempts"] = summary["current_semester_attempts"].fillna(0).astype(int)
    summary["all_clear"] = (
        (summary["total_subjects_attempted"] > 0)
        & (summary["total_subjects_attempted"] == summary["total_subjects_passed"])
    ).astype(int)
    summary.loc[summary["gender"].isna() | (summary["gender"] == ""), "gender"] = "Unknown"
    summary.loc[summary["community"].isna() | (summary["community"] == ""), "community"] = "Unknown"
    return summary


def compute_gender_community_breakdown(overall: pd.DataFrame) -> pd.DataFrame:
    if overall.empty:
        return pd.DataFrame(columns=["gender", "community", "total_students", "all_clear_students", "all_clear_pct"])

    breakdown = (
        overall.groupby(["gender", "community"], dropna=False)
        .agg(
            total_students=("register_no", "nunique"),
            all_clear_students=("all_clear", "sum"),
        )
        .reset_index()
    )
    breakdown["all_clear_pct"] = np.where(
        breakdown["total_students"] > 0,
        (breakdown["all_clear_students"] / breakdown["total_students"] * 100).round(2),
        np.nan,
    )
    return breakdown


def compute_semester_overview(master: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        return pd.DataFrame(columns=[
            "semester",
            "unique_students",
            "subjects_attempted",
            "subjects_passed",
            "pass_pct",
        ])

    semester_level = (
        master.groupby("semester")
        .agg(
            unique_students=("register_no", "nunique"),
            subjects_attempted=("subject_code", "count"),
            subjects_passed=("pass_flag", "sum"),
        )
        .reset_index()
    )
    semester_level["pass_pct"] = np.where(
        semester_level["subjects_attempted"] > 0,
        (semester_level["subjects_passed"] / semester_level["subjects_attempted"] * 100).round(2),
        np.nan,
    )
    return semester_level


def compute_student_arrear_counts(master: pd.DataFrame, semesters: Iterable[int]) -> pd.DataFrame:
    """Return per-student arrear counts across the requested *semesters*."""

    semesters = [int(s) for s in semesters]
    base_columns = ["register_no"] + [f"arrear_sem_{sem}" for sem in semesters]

    if master.empty:
        return pd.DataFrame(columns=base_columns)

    register_series = master.get("register_no")
    if register_series is None:
        return pd.DataFrame(columns=base_columns)

    register_series = register_series.astype(str).str.strip()
    register_series = register_series[register_series != ""]
    register_order = register_series.drop_duplicates().tolist()

    if not register_order:
        return pd.DataFrame(columns=base_columns)

    base_df = pd.DataFrame({"register_no": register_order})

    subject_code_series = master.get("subject_code")
    if subject_code_series is None:
        for sem in semesters:
            base_df[f"arrear_sem_{sem}"] = 0
        return base_df[base_columns]

    subjects = master[subject_code_series.notna()].copy()
    if subjects.empty:
        for sem in semesters:
            base_df[f"arrear_sem_{sem}"] = 0
        return base_df[base_columns]

    subjects["semester_num"] = pd.to_numeric(subjects.get("semester"), errors="coerce").astype("Int64")
    subjects = subjects[subjects["semester_num"].notna()].copy()

    if subjects.empty:
        for sem in semesters:
            base_df[f"arrear_sem_{sem}"] = 0
        return base_df[base_columns]

    subject_status = (
        subjects.groupby(["register_no", "semester_num", "subject_code"], dropna=False)
        .agg(has_pass=("pass_flag", lambda s: bool(pd.Series(s).fillna(False).any())))
        .reset_index()
    )

    subject_status["arrear_flag"] = (~subject_status["has_pass"]).astype(int)
    counts = (
        subject_status.groupby(["register_no", "semester_num"], dropna=False)["arrear_flag"]
        .sum()
        .reset_index(name="arrear_count")
    )

    if counts.empty:
        for sem in semesters:
            base_df[f"arrear_sem_{sem}"] = 0
        return base_df[base_columns]

    pivot = counts.pivot_table(index="register_no", columns="semester_num", values="arrear_count", fill_value=0)
    pivot = pivot.reindex(register_order, fill_value=0)

    for sem in semesters:
        if sem not in pivot.columns:
            pivot[sem] = 0

    pivot = pivot[semesters]
    pivot = pivot.reset_index()
    pivot.columns = ["register_no"] + [f"arrear_sem_{sem}" for sem in semesters]

    result = base_df.merge(pivot, on="register_no", how="left")
    for sem in semesters:
        col = f"arrear_sem_{sem}"
        result[col] = result[col].fillna(0).astype(int)

    return result[base_columns]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Path to Excel file")
    ap.add_argument("--outdir", default="outputs")
    ap.add_argument("--config", default=os.path.join(HERE, "config.json"))
    ap.add_argument(
        "--biodata",
        default=os.path.join(HERE, "biodata.xlsx"),
        help="Path to biodata workbook (register vs gender/community)",
    )
    ap.add_argument(
        "--subject-catalog",
        default=os.path.join(HERE, "CGPA.xlsx"),
        help="Workbook mapping subject codes to credits and semesters",
    )
    ap.add_argument(
        "--current-semester",
        type=int,
        default=None,
        help="Semester number considered the current semester (defaults to config or 8)",
    )
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    cfg_path = args.config if os.path.isfile(args.config) else os.path.join(os.getcwd(), "config.json")
    cfg = load_config(cfg_path)

    value_regex = cfg.get("register_value_regex", r"^\d{6,}$")
    subj_hdr_rx = cfg.get("subject_header_regex", r"^\s*([A-Z]{2,}\d{3,4})\s*\(\s*([\d\.]+)\s*\)\s*$")
    arrear_keys = set(k.lower() for k in cfg.get("arrear_sheet_keywords", ["arrear","arrears","backlog","backlogs","repeat","supplementary"]))
    current_semester = args.current_semester or cfg.get("current_semester", 8)
    default_exam_session = cfg.get("default_exam_session", "Apr-May 2025")

    subject_catalog = load_subject_catalog(args.subject_catalog)

    sheets = read_all_sheets(args.input)
    master_rows = []
    dq_issues = []

    for sh_name, df_raw in sheets.items():
        if df_raw is None or df_raw.empty:
            dq_issues.append({"sheet": sh_name, "issue": "Empty sheet"})
            continue

        # Filter out non-student rows first
        reg_header_rx = cfg.get("register_header_regex", r"^reg(ister)?\s*\.?\s*(no|number)\b")
        reg_col = find_col(df_raw, [], reg_header_rx)
        if not reg_col:
            # Attempt vertical single-student layout parsing
            student_df, issues = parse_vertical_student_sheet(
                args.input,
                sh_name,
                current_semester,
                subject_catalog,
                default_exam_session,
            )
            dq_issues.extend(issues)
            if student_df.empty:
                dq_issues.append({"sheet": sh_name, "issue": "Missing REGISTER NUMBER column"})
                continue
            master_rows.append(student_df)
            continue

        df = df_raw.copy()
        df = df[df[reg_col].apply(lambda x: valid_register(x, value_regex))].reset_index(drop=True)
        if df.empty:
            dq_issues.append({"sheet": sh_name, "issue": "No valid register rows"})
            continue

        # Other columns
        res_col = find_col(df, ["result"], cfg.get("result_header_regex", r"^result\b"))
        gpa_col = find_col(df, ["gpa","sgpa","sem gpa"], cfg.get("gpa_header_regex", r"^gpa\b"))
        cgpa_col= find_col(df, ["cgpa"], cfg.get("cgpa_header_regex", r"^cgpa\b"))
        reap_col= find_col(df, cfg.get("no_of_subjects_reappear_aliases", []))

        subj_cols = find_subject_columns(df, subj_hdr_rx)
        if not subj_cols:
            dq_issues.append({"sheet": sh_name, "issue": "No subject columns detected"})
        sem_num = detect_semester_from_sheet(sh_name)
        is_arrear = "Y" if any(k in sh_name.lower() for k in arrear_keys) else "N"

        base_cols = [c for c in [reg_col, res_col, gpa_col, cgpa_col, reap_col] if c]
        base = df[base_cols].copy()
        ren = {reg_col:"register_no"}
        if res_col: ren[res_col] = "result"
        if gpa_col: ren[gpa_col] = "gpa"
        if cgpa_col: ren[cgpa_col] = "cgpa"
        if reap_col: ren[reap_col] = "no_of_subjects_reappear"
        base = base.rename(columns=ren)

        if subj_cols:
            long_rows = []
            for idx, row in base.iterrows():
                for col, code, credit in subj_cols:
                    cell = df.loc[idx, col] if col in df.columns else ""
                    # STRICT RULE: count only if cell contains a known grade/result token
                    if not is_attended_grade(cell):
                        continue
                    long_rows.append({
                        "register_no": str(row.get("register_no","")).strip(),
                        "semester": sem_num,
                        "subject_code": code,
                        "subject_name": "",
                        "credit": credit,
                        "grade": str(cell).strip(),
                        "result": to_result_from_grade(cell, row.get("result","")),
                        "gpa": row.get("gpa",""),
                        "cgpa": row.get("cgpa",""),
                        "no_of_subjects_reappear": row.get("no_of_subjects_reappear",""),
                        "exam_session": cfg.get("default_exam_session","Apr-May 2025"),
                        "is_arrear": is_arrear,
                        "source_sheet": sh_name
                    })
            if long_rows:
                master_rows.append(pd.DataFrame(long_rows))
        else:
            # GPA/CGPA only rows (rare for your structure)
            tmp = base.copy()
            tmp["semester"] = sem_num
            tmp["subject_code"] = ""
            tmp["subject_name"] = ""
            tmp["credit"] = np.nan
            tmp["grade"] = ""
            tmp["exam_session"] = cfg.get("default_exam_session","Apr-May 2025")
            tmp["is_arrear"] = is_arrear
            tmp["source_sheet"] = sh_name
            master_rows.append(tmp)

    if not master_rows:
        raise SystemExit("No usable data found.")

    master = pd.concat(master_rows, ignore_index=True)
    master["register_no"] = master["register_no"].astype(str).str.strip()
    master["subject_code"] = master["subject_code"].replace({"": np.nan, None: np.nan}).astype(object)

    biodata_df = load_biodata(args.biodata)
    if not biodata_df.empty:
        master = master.merge(biodata_df, on="register_no", how="left")
    else:
        master["gender"] = "Unknown"
        master["community"] = "Unknown"

    # Analytics computed only on attended subjects (by design)
    master["result_norm"] = master["result"].astype(str).str.strip().str.lower()
    master["pass_flag"] = master["result_norm"].isin(["pass", "p", "passed"])
    master.loc[master["subject_code"].isna(), "pass_flag"] = False
    master["is_current_semester"] = np.where(
        (master.get("semester", -1) == current_semester) & (master["is_arrear"] == "N"),
        "Y",
        "N",
    )

    subject_outcomes = (
        master.groupby(["semester", "subject_code"], dropna=False)
        .agg(total=("register_no", "nunique"), passed=("pass_flag", "sum"))
        .reset_index()
    )
    subject_outcomes["pass_pct"] = np.where(subject_outcomes["total"]>0,
                                           (subject_outcomes["passed"]/subject_outcomes["total"]*100).round(2),
                                           np.nan)

    student_summary = compute_student_semester_summary(master)
    overall_summary = compute_student_overall_summary(master, current_semester)
    class_performance_summary = build_class_performance_summary(overall_summary)
    gender_comm_breakdown = compute_gender_community_breakdown(overall_summary)
    semester_overview = compute_semester_overview(master)

    cfg_semester_numbers = cfg.get("semester_numbers")
    if cfg_semester_numbers:
        semester_numbers = sorted({int(s) for s in cfg_semester_numbers})
    else:
        semester_series = master.get("semester")
        detected = []
        if semester_series is not None:
            detected = (
                pd.to_numeric(semester_series, errors="coerce")
                .dropna()
                .astype(int)
                .tolist()
            )
        detected_max = max(detected) if detected else 0
        total_semesters = int(cfg.get("total_semesters", 0))
        if total_semesters <= 0:
            total_semesters = current_semester if current_semester else 8
        end_sem = max(total_semesters, detected_max)
        if end_sem <= 0:
            end_sem = 8
        semester_numbers = list(range(1, end_sem + 1))

    arrear_counts = compute_student_arrear_counts(master, semester_numbers)

    all_clear_students = master.merge(
        overall_summary.loc[overall_summary["all_clear"] == 1, ["register_no"]],
        on="register_no",
        how="inner",
    )["register_no"].drop_duplicates().reset_index(drop=True)

    outdir = args.outdir
    os.makedirs(outdir, exist_ok=True)
    with pd.ExcelWriter(os.path.join(outdir, "master_results.xlsx"), engine="openpyxl") as w:
        master.to_excel(w, index=False, sheet_name="master_long")
        subject_outcomes.to_excel(w, index=False, sheet_name="subject_outcomes")
        student_summary.to_excel(w, index=False, sheet_name="student_summary")
        overall_summary.to_excel(w, index=False, sheet_name="student_overall")
        class_performance_summary.to_excel(w, index=False, sheet_name="class_performance_summary")
        gender_comm_breakdown.to_excel(w, index=False, sheet_name="gender_community")
        semester_overview.to_excel(w, index=False, sheet_name="semester_overview")
        arrear_counts.to_excel(w, index=False, sheet_name="arrear_counts_by_semester")
        all_clear_students.to_frame(name="register_no").to_excel(
            w, index=False, sheet_name="all_clear_registers"
        )

    master.to_csv(os.path.join(outdir, "master_results.csv"), index=False)
    subject_outcomes.to_csv(os.path.join(outdir, "analytics_subject_outcomes.csv"), index=False)
    student_summary.to_csv(os.path.join(outdir, "analytics_student_summary.csv"), index=False)
    overall_summary.to_csv(os.path.join(outdir, "analytics_student_overall.csv"), index=False)
    class_performance_summary.to_csv(
        os.path.join(outdir, "class_performance_summary.csv"), index=False
    )
    gender_comm_breakdown.to_csv(os.path.join(outdir, "analytics_gender_community.csv"), index=False)
    semester_overview.to_csv(os.path.join(outdir, "analytics_semester_overview.csv"), index=False)
    arrear_counts.to_csv(os.path.join(outdir, "analytics_student_arrear_counts.csv"), index=False)
    all_clear_students.to_frame(name="register_no").to_csv(
        os.path.join(outdir, "all_clear_registers.csv"), index=False
    )

    # Data-quality report
    dq_df = pd.DataFrame(dq_issues, columns=["sheet", "issue"])
    dq_df.to_csv(os.path.join(outdir, "data_quality_report.csv"), index=False)

    print("Wrote outputs to:", outdir)

if __name__ == "__main__":
    main()
