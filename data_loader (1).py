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
    "mr_no": ["mr no", "mrn", "file no", "file number", "mr number", "mr#", "mr no.", "patient id",
              "uhid", "hims no", "hims no/mr", "hims"],
    "patient_name": ["patient name", "patient", "pt name", "pt. name"],
    "ward": ["ward", "department", "dept", "unit", "location", "ward/dept"],
    "consultant": ["consultant", "doctor", "dr", "physician", "attending", "consultant name",
                    "prescriber", "consultant/prescriber"],
    "intervention": ["intervention", "intervention type", "type of intervention", "intervention detail",
                     "description", "intervention description", "clinical intervention"],
    "status": ["status", "acceptance", "accepted", "outcome", "result", "accepted/rejected",
               "acceptance status", "response"],

    # Medication errors sheet
    "error_type": ["kind of error", "error type", "type of error", "error category", "category",
                   "classification", "error classification", "type"],
    "error_description": ["description", "error description", "detail", "details", "narrative"],

    # HAM/LASA sheet
    "drug_name": ["drug name", "medication", "drug", "medicine", "ham/lasa drug", "ham drug",
                  "lasa drug", "drug/medication", "medicine name", "product name", "product"],
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
    "medication_errors": ["error", "medication error", "med error", "prescribing error", "prescription error"],
    "ham_lasa": ["ham", "lasa", "high alert", "look alike", "sound alike", "ham/lasa", "consumption"],
    "adr": ["adr", "adverse", "adverse drug", "adverse reaction"],
    "summary": ["summary", "overview", "dashboard"],
}


def _find_header_row(uploaded_bytes: BytesIO, sheet_name: str, max_rows_to_check: int = 20) -> int:
    """
    Detect which row actually contains the column headers.

    Some hospital Excel exports have several rows above the real header
    row — a hospital letterhead, a report title, a date range, blank
    spacer rows, etc. (e.g. headers starting on row 6 or row 9 instead
    of row 1). To find the real header row reliably regardless of how
    many preamble rows there are, this scores each of the first
    `max_rows_to_check` rows by how many of its cell values match a
    known column-name alias from `COLUMN_ALIASES` (e.g. "ward",
    "consultant", "mr no", "drug name"...). The row with the most
    matches is almost certainly the real header row, since preamble
    rows (titles, letterheads, dates) essentially never match these
    known field names.

    Falls back to "first row with more than one filled-in cell" if no
    row scores any alias matches at all (e.g. an unusual naming scheme).

    Returns
    -------
    int
        0-based row index to use as the header row.
    """
    raw = pd.read_excel(uploaded_bytes, sheet_name=sheet_name, header=None, nrows=max_rows_to_check)

    all_aliases = set()
    for aliases in COLUMN_ALIASES.values():
        all_aliases.update(aliases)

    best_row = None
    best_score = 0
    for i in range(len(raw)):
        cells = [str(v).strip().lower() for v in raw.iloc[i] if pd.notna(v)]
        score = sum(1 for cell in cells if cell in all_aliases or any(
            alias in cell or cell in alias for alias in all_aliases if len(alias) > 2
        ))
        if score > best_score:
            best_score = score
            best_row = i

    if best_row is not None and best_score >= 2:
        return best_row

    # Fallback: no strong alias match anywhere — use the old heuristic of
    # the first row that has more than one filled-in cell (skips a
    # single merged title cell, but won't skip multiple preamble rows).
    for i in range(len(raw)):
        if raw.iloc[i].notna().sum() > 1:
            return i
    return 0


# ---------------------------------------------------------------------------
# ADR summary-block labels — for hospitals that record ADR totals as a
# small labeled summary (e.g. "Total ADRs Reported: 12") rather than one
# row per incident.
# ---------------------------------------------------------------------------
ADR_SUMMARY_LABELS = {
    "total_adrs": [
        "total adrs reported", "total adr reported", "total number of adrs",
        "total adrs", "total adr",
    ],
    "adrs_from_ham_lasa": [
        "adrs due to ham/lasa", "adrs due to ham lasa", "adr due to ham/lasa",
        "adrs from ham/lasa", "adrs from ham lasa", "ham/lasa related adrs",
        "ham/lasa related adr", "due to ham/lasa", "due to ham lasa",
    ],
    "adrs_other": [
        "other adrs", "other adr", "adrs from other drugs", "adrs other than ham/lasa",
        "non ham/lasa adrs", "non-ham/lasa adrs",
    ],
}


def try_read_adr_summary(uploaded_bytes: BytesIO, sheet_name: str) -> dict | None:
    """
    Try reading the ADR sheet as three plain labeled figures instead of
    a row-per-incident table.

    Some hospitals record ADR totals as just a few label/value pairs
    somewhere on the sheet (e.g. a cell reading "Total ADRs Reported"
    next to a cell containing the number 12) rather than logging one
    row per reaction. This scans every cell on the sheet for a label
    matching `ADR_SUMMARY_LABELS`, and reads the number immediately to
    its right — three straight figures, nothing derived or hidden:

        - Total ADRs
        - HAM/LASA-related ADRs
        - Other ADRs

    Whichever of the three are actually found on the sheet are used
    as-is. Only if exactly one of the three is missing do we fill it
    in with simple addition/subtraction (since with two known figures,
    the third is fully determined) — this is the one and only place
    any arithmetic happens.

    Returns
    -------
    dict | None
        {"total_adrs": int, "adrs_from_ham_lasa": int, "adrs_other": int}
        if at least two of the three labels were found on the sheet
        (enough to pin down all three figures), otherwise None — meaning
        the caller should fall back to the normal row-per-incident
        analysis instead.
    """
    raw = pd.read_excel(uploaded_bytes, sheet_name=sheet_name, header=None)
    found: dict[str, int] = {}

    for r in range(raw.shape[0]):
        for c in range(raw.shape[1] - 1):
            cell = raw.iat[r, c]
            if not isinstance(cell, str):
                continue
            label = cell.strip().lower()
            value = raw.iat[r, c + 1]
            if not isinstance(value, (int, float)) or pd.isna(value):
                continue
            for key, aliases in ADR_SUMMARY_LABELS.items():
                if key in found:
                    continue
                if any(alias in label for alias in aliases):
                    found[key] = int(value)

    if len(found) < 2:
        return None  # not enough figures on the sheet to pin down all three

    if "total_adrs" not in found:
        found["total_adrs"] = found["adrs_from_ham_lasa"] + found["adrs_other"]
    elif "adrs_from_ham_lasa" not in found:
        found["adrs_from_ham_lasa"] = found["total_adrs"] - found["adrs_other"]
    elif "adrs_other" not in found:
        found["adrs_other"] = found["total_adrs"] - found["adrs_from_ham_lasa"]

    return {
        "total_adrs": found["total_adrs"],
        "adrs_from_ham_lasa": found["adrs_from_ham_lasa"],
        "adrs_other": found["adrs_other"],
    }


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


def _sheet_type_scores(sheet_name: str, df: pd.DataFrame) -> dict[str, tuple[int, int]]:
    """
    Score every sheet type against a single sheet.

    Returns a dict of sheet_type -> (tier, score) for every type that has
    at least one keyword match, where tier 0 = matched via the sheet's
    own name (very reliable) and tier 1 = only matched via column
    headers (less reliable, since column text can accidentally overlap
    with another sheet type's keywords). Lower tier is always preferred;
    within the same tier, higher score (longer/more specific keyword) wins.
    """
    name_lower = sheet_name.lower().strip()
    cols_lower = " ".join(str(c).lower() for c in df.columns)

    scores: dict[str, tuple[int, int]] = {}
    for sheet_type, keywords in SHEET_KEYWORDS.items():
        best_tier1 = 0
        for kw in keywords:
            if kw in name_lower:
                # Tier 0: matched on the sheet's own name.
                scores[sheet_type] = min(scores.get(sheet_type, (1, 0)), (0, -len(kw)))
            elif kw in cols_lower:
                best_tier1 = max(best_tier1, len(kw))
        if sheet_type not in scores and best_tier1 > 0:
            scores[sheet_type] = (1, -best_tier1)
    return scores


def detect_sheet_type(sheet_name: str, df: pd.DataFrame) -> str | None:
    """
    Guess the purpose of a single sheet, considered in isolation.

    Prefers a match on the sheet's own name over one from column text.
    For resolving conflicts *between* sheets (e.g. two sheets both
    plausibly matching the same type), see `auto_detect_sheets`, which
    looks at all sheets together instead of one at a time.

    Returns one of: 'interventions', 'medication_errors', 'ham_lasa', 'adr', 'summary', or None.
    """
    scores = _sheet_type_scores(sheet_name, df)
    if not scores:
        return None
    # Lowest (tier, -score) tuple wins: tier 0 beats tier 1, and within a
    # tier, the more negative (i.e. longer keyword) score wins.
    return min(scores.items(), key=lambda kv: kv[1])[0]


def auto_detect_sheets(sheets: dict[str, pd.DataFrame]) -> dict[str, str]:
    """
    Auto-detect which sheet maps to which type.

    Considers all sheets together rather than one at a time. This
    matters because a sheet can accidentally contain a keyword that
    "belongs" to a different sheet type (e.g. a Medication Errors sheet
    with its own "Intervention" column, describing the corrective action
    taken). Rather than assigning types on a first-come-first-served
    basis (which can let a weak, coincidental match steal a slot from
    the sheet that's actually the best fit), every (sheet, type) match
    is scored, and the strongest matches are assigned first.

    Returns
    -------
    dict[str, str]
        Mapping of sheet_type → sheet_name.
    """
    # Collect every (tier, score) candidate across all sheets.
    candidates = []  # list of (tier, score, sheet_name, sheet_type)
    for sheet_name, df in sheets.items():
        for sheet_type, (tier, score) in _sheet_type_scores(sheet_name, df).items():
            candidates.append((tier, score, sheet_name, sheet_type))

    # Lower tier first (name match beats column match), then more
    # negative score first (i.e. longer/more specific keyword wins).
    candidates.sort(key=lambda c: (c[0], c[1]))

    mapping: dict[str, str] = {}
    used_sheets: set[str] = set()
    for tier, score, sheet_name, sheet_type in candidates:
        if sheet_type in mapping or sheet_name in used_sheets:
            continue
        mapping[sheet_type] = sheet_name
        used_sheets.add(sheet_name)
    return mapping


def auto_map_columns(df: pd.DataFrame, required_fields: list[str], threshold: int = 85) -> dict[str, str | None]:
    """
    Fuzzy-match DataFrame columns to expected field names.

    Uses `token_set_ratio` rather than `token_sort_ratio`, since real
    hospital exports often have headers that are a superset of the
    expected name (e.g. "Issued Qty" for "quantity", or a padded header
    like "STATUS                    (ACCEPTED/REJECTED)") — token_sort
    penalizes the length difference heavily and misses these, while
    token_set correctly recognizes the alias as effectively contained
    within the longer header.

    Also ensures each column is only assigned to a single field. Every
    (field, column) match is scored, then the strongest matches are
    assigned first — this prevents a column that happens to weakly
    match several fields (e.g. a lone "Patient" column that vaguely
    resembles several aliases) from being claimed by the wrong one, or
    from being reused for a field it doesn't really belong to.

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
    # Match against stripped names (so whitespace doesn't throw off the
    # fuzzy score), but always report back the *original* column name —
    # real hospital exports often have headers with stray leading or
    # trailing spaces (e.g. "WARD " or "   STATUS   (...)"), and
    # returning a stripped name that doesn't actually exist as a column
    # would make every downstream `col in df.columns` check silently
    # (and misleadingly) fail.
    original_by_stripped: dict[str, str] = {str(c).strip(): c for c in df.columns}
    col_names = list(original_by_stripped.keys())

    # For each field, walk its aliases in priority order and take the
    # first one that has a strong match — this ensures a specific,
    # deliberately-ordered alias (e.g. "kind of error", listed first)
    # takes precedence over a more generic one (e.g. "type", listed
    # last) even if the generic one happens to also score well against
    # a different, wrong column.
    field_candidates = []  # list of (score, field_priority, field, column)
    for field_priority, field in enumerate(required_fields):
        aliases = COLUMN_ALIASES.get(field, [field])
        for alias in aliases:
            result = process.extractOne(alias, col_names, scorer=fuzz.token_set_ratio)
            if result and result[1] >= threshold:
                field_candidates.append((result[1], field_priority, field, result[0]))
                break  # earlier aliases take precedence; stop at first hit

    # Resolve any remaining collisions where two different fields' best
    # candidates land on the same column: highest score wins, ties
    # broken by whichever field was listed first.
    field_candidates.sort(key=lambda c: (-c[0], c[1]))

    mapping: dict[str, str | None] = {field: None for field in required_fields}
    used_columns: set[str] = set()
    for score, _, field, col in field_candidates:
        original_col = original_by_stripped[col]
        if mapping[field] is not None or original_col in used_columns:
            continue
        mapping[field] = original_col
        used_columns.add(original_col)

    return mapping

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
