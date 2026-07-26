"""
analyzer.py — Core analysis functions for clinical pharmacy data.

Each function takes a cleaned DataFrame (with standardized column names)
and returns analysis results as dictionaries or DataFrames suitable for
chart generation and report writing.
"""

import pandas as pd


def analyze_interventions(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Analyze intervention data.

    Returns
    -------
    dict with keys:
        - total_files_reviewed (int): unique patient files
        - total_interventions (int): total intervention rows
        - accepted (int): accepted interventions
        - rejected (int): rejected interventions
        - acceptance_rate (float): percentage accepted
    """
    mr_col = col_map.get("mr_no")
    status_col = col_map.get("status")

    total_interventions = len(df)
    total_files = df[mr_col].nunique() if mr_col and mr_col in df.columns else total_interventions

    accepted = 0
    rejected = 0
    if status_col and status_col in df.columns:
        status_clean = df[status_col].astype(str).str.strip().str.lower()
        accepted = status_clean.isin(["accepted", "accept", "yes", "approved", "y"]).sum()
        rejected = status_clean.isin(["rejected", "reject", "no", "denied", "n", "not accepted"]).sum()
        # Anything else is "pending" or unknown
    else:
        # If no status column, assume all are accepted
        accepted = total_interventions

    acceptance_rate = (accepted / total_interventions * 100) if total_interventions > 0 else 0

    return {
        "total_files_reviewed": int(total_files),
        "total_interventions": int(total_interventions),
        "accepted": int(accepted),
        "rejected": int(rejected),
        "acceptance_rate": round(acceptance_rate, 1),
    }


def interventions_by_ward(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    Count interventions grouped by ward.

    Returns a DataFrame with columns: Ward, Count — sorted descending.
    """
    ward_col = col_map.get("ward")
    if not ward_col or ward_col not in df.columns:
        return pd.DataFrame(columns=["Ward", "Count"])

    counts = (
        df[ward_col]
        .astype(str).str.strip().str.title()
        .value_counts()
        .reset_index()
    )
    counts.columns = ["Ward", "Count"]
    return counts


def interventions_by_consultant(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    Count interventions grouped by consultant (after name normalization).

    Expects the consultant column to already be normalized via
    `data_loader.normalize_consultant_names`.

    Returns a DataFrame with columns: Consultant, Count — sorted descending.
    """
    cons_col = col_map.get("consultant")
    if not cons_col or cons_col not in df.columns:
        return pd.DataFrame(columns=["Consultant", "Count"])

    counts = (
        df[cons_col]
        .astype(str).str.strip().str.title()
        .value_counts()
        .reset_index()
    )
    counts.columns = ["Consultant", "Count"]
    return counts


def analyze_errors(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Analyze medication errors by type.

    Returns
    -------
    dict with keys:
        - total_errors (int)
        - by_type (pd.DataFrame): columns [Error Type, Count]
        - prescription (int)
        - dispensing (int)
        - transcription (int)
        - illegible_handwriting (int)
        - incorrect_abbreviation (int)
        - other (int)
    """
    error_type_col = col_map.get("error_type")
    total_errors = len(df)

    result = {
        "total_errors": total_errors,
        "prescription": 0,
        "administration": 0,
        "transcription": 0,
        "illegible_handwriting": 0,
        "incorrect_abbreviation": 0,
        "other": 0,
        "by_type": pd.DataFrame(columns=["Error Type", "Count"]),
    }

    if not error_type_col or error_type_col not in df.columns:
        return result

    types_clean = df[error_type_col].astype(str).str.strip().str.lower()

    # Classify each error
    categories = {
        "prescription": ["prescription", "prescription error", "prescribing", "prescribing error"],
        "administration": ["administration", "administration error", "wrong route", "wrong time",
                            "wrong rate", "missed dose", "omission", "administer"],
        "transcription": ["transcription", "transcription error", "transcrib"],
        "illegible_handwriting": ["illegible", "handwriting", "illegible handwriting",
                                  "unreadable", "unclear handwriting"],
        "incorrect_abbreviation": ["abbreviation", "incorrect abbreviation", "wrong abbreviation",
                                   "abbr", "unapproved abbreviation"],
    }

    classified = []
    for val in types_clean:
        found = False
        for cat_name, keywords in categories.items():
            if any(kw in val for kw in keywords):
                classified.append(cat_name)
                found = True
                break
        if not found:
            classified.append("other")

    df_classified = pd.DataFrame({"category": classified})
    counts = df_classified["category"].value_counts()

    for cat in categories:
        result[cat] = int(counts.get(cat, 0))
    result["other"] = int(counts.get("other", 0))

    # Build by_type DataFrame
    type_counts = (
        df[error_type_col]
        .astype(str).str.strip().str.title()
        .value_counts()
        .reset_index()
    )
    type_counts.columns = ["Error Type", "Count"]
    result["by_type"] = type_counts

    return result


def errors_by_ward(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    Cross-tabulate errors by ward and error type.

    Returns a DataFrame with Ward as index and error types as columns.
    """
    ward_col = col_map.get("ward")
    error_type_col = col_map.get("error_type")

    if (not ward_col or ward_col not in df.columns or
            not error_type_col or error_type_col not in df.columns):
        return pd.DataFrame()

    df_copy = df.copy()
    df_copy["_ward"] = df_copy[ward_col].astype(str).str.strip().str.title()
    df_copy["_error_type"] = df_copy[error_type_col].astype(str).str.strip().str.title()

    cross = pd.crosstab(df_copy["_ward"], df_copy["_error_type"])
    cross.index.name = "Ward"
    return cross


def analyze_ham_lasa(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Analyze HAM/LASA consumption data.

    Returns
    -------
    dict with keys:
        - total_patients (int): unique patients who received HAM/LASA
        - total_records (int): total rows
        - by_drug (pd.DataFrame): columns [Drug, Count] — most common drugs
        - by_type (pd.DataFrame): columns [Type, Count] — HAM vs LASA split
    """
    mr_col = col_map.get("mr_no")
    drug_col = col_map.get("drug_name")
    type_col = col_map.get("ham_lasa_type")
    qty_col = col_map.get("quantity")

    total_records = len(df)
    total_patients = df[mr_col].nunique() if mr_col and mr_col in df.columns else total_records

    # By drug
    by_drug = pd.DataFrame(columns=["Drug", "Count"])
    if drug_col and drug_col in df.columns:
        if qty_col and qty_col in df.columns:
            # Sum quantities per drug
            drug_qty = df.groupby(
                df[drug_col].astype(str).str.strip().str.title()
            )[qty_col].sum().reset_index()
            drug_qty.columns = ["Drug", "Count"]
            drug_qty = drug_qty.sort_values("Count", ascending=False).reset_index(drop=True)
            by_drug = drug_qty
        else:
            drug_counts = (
                df[drug_col]
                .astype(str).str.strip().str.title()
                .value_counts()
                .reset_index()
            )
            drug_counts.columns = ["Drug", "Count"]
            by_drug = drug_counts

    # By type (HAM vs LASA)
    by_type = pd.DataFrame(columns=["Type", "Count"])
    if type_col and type_col in df.columns:
        type_counts = (
            df[type_col]
            .astype(str).str.strip().str.upper()
            .value_counts()
            .reset_index()
        )
        type_counts.columns = ["Type", "Count"]
        by_type = type_counts

    return {
        "total_patients": int(total_patients),
        "total_records": int(total_records),
        "by_drug": by_drug,
        "by_type": by_type,
    }


def ham_lasa_by_patient(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    Tally how many HAM/LASA records each patient has — i.e. how many
    times each patient received a High Alert / Look-Alike Sound-Alike
    medication during the period.

    Returns a DataFrame with columns: Patient Name, MR No, Count —
    sorted descending by Count.
    """
    mr_col = col_map.get("mr_no")
    name_col = col_map.get("patient_name")

    if not mr_col or mr_col not in df.columns:
        return pd.DataFrame(columns=["Patient Name", "MR No", "Count"])

    df_copy = df.copy()
    df_copy["_mr"] = df_copy[mr_col].astype(str).str.strip()
    if name_col and name_col in df.columns:
        df_copy["_name"] = df_copy[name_col].astype(str).str.strip().str.title()
    else:
        df_copy["_name"] = df_copy["_mr"]

    tally = (
        df_copy.groupby("_mr")
        .agg(**{"Patient Name": ("_name", "first"), "Count": ("_mr", "size")})
        .reset_index()
        .rename(columns={"_mr": "MR No"})
    )
    tally = tally[["Patient Name", "MR No", "Count"]].sort_values(
        "Count", ascending=False
    ).reset_index(drop=True)
    return tally


def analyze_adrs(df: pd.DataFrame, col_map: dict) -> dict:
    """
    Analyze Adverse Drug Reaction data.

    Returns
    -------
    dict with keys:
        - total_adrs (int)
        - adrs_from_ham_lasa (int): ADRs caused by HAM/LASA drugs
        - adrs_other (int): non-HAM/LASA ADRs
        - by_drug (pd.DataFrame): columns [Drug, Count]
    """
    ham_col = col_map.get("is_ham_lasa")
    drug_col = col_map.get("causative_drug")

    total_adrs = len(df)
    adrs_from_ham_lasa = 0

    if ham_col and ham_col in df.columns:
        ham_clean = df[ham_col].astype(str).str.strip().str.lower()
        adrs_from_ham_lasa = ham_clean.isin(["yes", "y", "true", "1", "ham", "lasa", "ham/lasa"]).sum()

    by_drug = pd.DataFrame(columns=["Drug", "Count"])
    if drug_col and drug_col in df.columns:
        drug_counts = (
            df[drug_col]
            .astype(str).str.strip().str.title()
            .value_counts()
            .reset_index()
        )
        drug_counts.columns = ["Drug", "Count"]
        by_drug = drug_counts

    return {
        "total_adrs": int(total_adrs),
        "adrs_from_ham_lasa": int(adrs_from_ham_lasa),
        "adrs_other": int(total_adrs - adrs_from_ham_lasa),
        "by_drug": by_drug,
    }
