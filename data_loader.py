"""
data_loader.py — Excel file loading, sheet detection, and column mapping utilities.

Handles reading multi-sheet Excel workbooks, auto-detecting column purposes
via fuzzy matching, and normalizing consultant names to handle typos/variants.
"""

import pandas as pd
from thefuzz import fuzz, process
from io import BytesIO


# ---------------------------------------------------------------------------
# Known column aliases — used for auto-detection
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    # Intervention sheet
    "date": ["date", "intervention date", "dated", "dt"],
    "mr_no": ["mr no", "mrn", "file no", "file number", "mr number", "mr#", "mr no.", "patient id", "uhid"],
    "patient_name": ["patient name", "patient", "name", "pt name", "pt. name"],
    "ward": ["ward", "department", "dept", "unit", "location", "ward/dept"],
    "consultant": ["consultant", "doctor", "dr", "physician", "attending", "consultant name",
                    "prescriber", "consultant/prescriber"],
    "intervention": ["intervention", "intervention type", "type of intervention", "intervention detail",
                     "description", "intervention description", "clinical intervention"],
    "status": ["status", "acceptance", "accepted", "outcome", "result", "accepted/rejected",
               "acceptance status", "response"],

    # Medication errors sheet
    "error_type": ["error type", "type of error", "type", "error category", "category",
                   "classification", "error classification"],
    "error_description": ["description", "error description", "detail", "details", "narrative"],

    # HAM/LASA sheet
    "drug_name": ["drug name", "medication", "drug", "medicine", "ham/lasa drug", "ham drug",
                  "lasa drug", "drug/medication", "medicine name"],
    "ham_lasa_type": ["type", "ham/lasa", "category", "classification", "ham or lasa",
                      "ham/lasa type", "ham/lasa category"],
    "quantity": ["quantity", "qty", "amount", "consumption", "units", "doses", "no. of doses",
                 "total quantity"],

    # ADR sheet
    "reaction": ["reaction", "adr", "adverse reaction", "adverse event", "event",
                 "adverse drug reaction", "adr description"],
    "causative_drug": ["causative drug", "drug", "suspected drug", "medication",
                       "causative agent", "offending drug"],
    "is_ham_lasa": ["is ham/lasa", "ham/lasa", "ham lasa", "is ham", "is lasa",
                    "ham/lasa related", "related to ham/lasa"],
    "severity": ["severity", "grade", "seriousness", "severity grade"],
}

# Sheet type detection keywords
SHEET_KEYWORDS = {
    "interventions": ["intervention", "clinical intervention", "pharmacist intervention"],
    "medication_errors": ["error", "medication error", "med error", "prescribing error"],
    "ham_lasa": ["ham", "lasa", "high alert", "look alike", "sound alike", "ham/lasa", "consumption"],
    "adr": ["adr", "adverse", "adverse drug", "adverse reaction"],
    "summary": ["summary", "overview", "dashboard"],
}


def _find_header_row(uploaded_bytes: BytesIO, sheet_name: str, max_rows_to_check: int = 5) -> int:
    """
    Detect which row actually contains the column headers.

    Some hospital Excel exports have a merged title row above the real
    headers (e.g. row 1 = "Pharmacist Intervention", row 2 = the real
    column names). A title row like this has only a single non-empty
    cell, while a genuine header row has several. This scans the first
    few rows and returns the index of the first row with more than one
    non-empty cell.

    Returns
    -------
    int
        0-based row index to use as the header row.
    """
    raw = pd.read_excel(uploaded_bytes, sheet_name=sheet_name, header=None, nrows=max_rows_to_check)
    for i in range(len(raw)):
        non_null_count = raw.iloc[i].notna().sum()
        if non_null_count > 1:
            return i
    return 0


def load_excel(uploaded_file) -> dict[str, pd.DataFrame]:
    """
    Read all sheets from an uploaded Excel file.

    Automatically detects and skips a merged title row above the real
    header row, if present (see `_find_header_row`).

    Parameters
    ----------
    uploaded_file : streamlit.UploadedFile or file-like
        The Excel file uploaded by the user.

    Returns
    -------
    dict[str, pd.DataFrame]
        Mapping of sheet name → DataFrame.
    """
    file_bytes = BytesIO(uploaded_file.read())
    xls = pd.ExcelFile(file_bytes, engine="openpyxl")
    sheets = {}
    for name in xls.sheet_names:
        header_row = _find_header_row(xls, name)
        df = pd.read_excel(xls, sheet_name=name, header=header_row)
        # Drop any columns that came out unnamed (stray blank columns)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed:")]
        # Drop completely empty rows/columns
        df = df.dropna(how="all").dropna(axis=1, how="all")
        if not df.empty:
            sheets[name] = df
    return sheets


def detect_sheet_type(sheet_name: str, df: pd.DataFrame) -> str | None:
    """
    Guess the purpose of a sheet based on its name and column headers.

    Returns one of: 'interventions', 'medication_errors', 'ham_lasa', 'adr', 'summary', or None.
    """
    name_lower = sheet_name.lower().strip()
    cols_lower = " ".join(str(c).lower() for c in df.columns)
    combined = f"{name_lower} {cols_lower}"

    best_score = 0
    best_type = None

    for sheet_type, keywords in SHEET_KEYWORDS.items():
        for kw in keywords:
            if kw in combined:
                score = len(kw)  # longer keyword = more specific = better match
                if score > best_score:
                    best_score = score
                    best_type = sheet_type

    return best_type


def auto_detect_sheets(sheets: dict[str, pd.DataFrame]) -> dict[str, str]:
    """
    Auto-detect which sheet maps to which type.

    Returns
    -------
    dict[str, str]
        Mapping of sheet_type → sheet_name.
    """
    mapping = {}
    for sheet_name, df in sheets.items():
        detected = detect_sheet_type(sheet_name, df)
        if detected and detected not in mapping:
            mapping[detected] = sheet_name
    return mapping


def auto_map_columns(df: pd.DataFrame, required_fields: list[str], threshold: int = 70) -> dict[str, str | None]:
    """
    Fuzzy-match DataFrame columns to expected field names.

    Parameters
    ----------
    df : pd.DataFrame
        The sheet data.
    required_fields : list[str]
        List of field keys from COLUMN_ALIASES to map.
    threshold : int
        Minimum fuzzy match score (0-100).

    Returns
    -------
    dict[str, str | None]
        Mapping of field_key → actual column name (or None if not found).
    """
    col_names = [str(c).strip() for c in df.columns]
    mapping = {}

    for field in required_fields:
        aliases = COLUMN_ALIASES.get(field, [field])
        best_match = None
        best_score = 0

        for alias in aliases:
            result = process.extractOne(alias, col_names, scorer=fuzz.token_sort_ratio)
            if result and result[1] > best_score:
                best_score = result[1]
                best_match = result[0]

        if best_score >= threshold:
            mapping[field] = best_match
        else:
            mapping[field] = None

    return mapping


def normalize_consultant_names(names: pd.Series, threshold: int = 85) -> pd.Series:
    """
    Merge similar consultant names using fuzzy matching.

    Groups names that are above the similarity threshold, keeping the most
    frequent variant as the canonical name.

    Parameters
    ----------
    names : pd.Series
        Raw consultant name column.
    threshold : int
        Minimum fuzzy match score to consider names as the same person.

    Returns
    -------
    pd.Series
        Cleaned consultant names with similar variants unified.
    """
    cleaned = names.astype(str).str.strip().str.title()
    unique_names = cleaned.dropna().unique().tolist()

    # Build groups
    name_map = {}
    groups = []

    for name in unique_names:
        if name in name_map:
            continue
        group = [name]
        for other in unique_names:
            if other != name and other not in name_map:
                score = fuzz.token_sort_ratio(name.lower(), other.lower())
                if score >= threshold:
                    group.append(other)
        # The canonical name is the most frequent one in the original data
        canonical = max(group, key=lambda n: (cleaned == n).sum())
        for member in group:
            name_map[member] = canonical
        groups.append(group)

    return cleaned.map(lambda x: name_map.get(x, x))


def get_required_fields(sheet_type: str) -> list[str]:
    """Return the list of expected column fields for a given sheet type."""
    fields = {
        "interventions": ["date", "mr_no", "patient_name", "ward", "consultant", "intervention", "status"],
        "medication_errors": ["date", "mr_no", "ward", "error_type", "error_description"],
        "ham_lasa": ["date", "mr_no", "patient_name", "ward", "drug_name", "ham_lasa_type", "quantity"],
        "adr": ["date", "mr_no", "patient_name", "causative_drug", "reaction", "is_ham_lasa", "severity"],
    }
    return fields.get(sheet_type, [])
