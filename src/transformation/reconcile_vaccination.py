"""
============================================================
AFP VACCINATION STATUS RECONCILIATION PIPELINE
============================================================

Project:
    AFP_Data_Engineering

Purpose:
    Reconcile AFP vaccination-status information between
    Microsoft Access surveillance data and an Excel workbook.

Matching key:
    EpidNumber

Calculation:
    Total polio doses =
        IPV Routine + OPV Routine + OPV AVS

AFP surveillance coding:
    0          = Zero reported doses (valid)
    1, 2, 3... = Known number of doses (valid)
    99         = Unknown number of doses
                 (information unavailable)

Standardized unknown value:
    Unknown

Rules:
    1. Skip completely empty Excel rows.
    2. Preserve existing valid numeric vaccination values.
    3. Treat 99 as Unknown, never as 99 doses.
    4. Fill unknown Excel values using valid Access values.
    5. Standardize unresolved unknown values to "Unknown".
    6. Never replace a known numeric value with an unknown.
    7. Never overwrite conflicting valid numeric dose values.
    8. Calculate total doses using all three components.
    9. If any component is unknown, total = "Unknown".
    10. Preserve original values in the audit trail.
    11. Flag unmatched, duplicate, conflicting and invalid records.
    12. Add explanatory comments to the Excel workbook.
    13. Never overwrite original source files.

Dependencies:
    pyodbc
    openpyxl

============================================================
"""

from pathlib import Path
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import datetime

import pyodbc

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 1. PROJECT CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parents[2]

RAW_DIR = ROOT / "data" / "raw"

PROCESSED_DIR = ROOT / "data" / "processed"

OUTPUT_DIR = ROOT / "outputs"


MDB_FILE = (
    RAW_DIR
    / "AAFR2512&2610Week40-11Oct26Data.mdb"
)

EXCEL_FILE = (
    RAW_DIR
    / "Statut vaccinal 08 Oct 2026.xlsx"
)

UPDATED_FILE = (
    PROCESSED_DIR
    / "AFP_Vaccination_Status_Updated.xlsx"
)

AUDIT_FILE = (
    OUTPUT_DIR
    / "AFP_Vaccination_Reconciliation_Audit.xlsx"
)


ACCESS_TABLE = "AFPcases"

EXCEL_SHEET = "Folha1"


PROCESSED_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# 2. FIELD MAPPING
# ============================================================

FIELD_MAPPING = {

    "TotalIPVdosesRI":
        "TotalIPVdosesRI",

    "Total doses VPO Routine":
        "TotalOPVdosesRI",

    "Total doses VPO AVS":
        "TotalOPVdosesSIA",

}


EPID_COLUMN = "EpidNumber"

IPV_COLUMN = "TotalIPVdosesRI"

ROUTINE_COLUMN = "Total doses VPO Routine"

SIA_COLUMN = "Total doses VPO AVS"

TOTAL_COLUMN = "Total doses polio (Routine +AVS)"


DOSE_COLUMNS = [
    IPV_COLUMN,
    ROUTINE_COLUMN,
    SIA_COLUMN,
]


# ============================================================
# 3. SURVEILLANCE CODING CONFIGURATION
# ============================================================

UNKNOWN_LABEL = "Unknown"

UNKNOWN_NUMERIC_CODES = {
    99
}

UNKNOWN_TEXT_CODES = {

    "",
    "inconnu",
    "unknown",
    "missing",
    "none",
    "null",
    "nan",
    "n/a",
    "na",

}


def is_unknown(value):
    """
    Identify missing or unknown vaccination information.

    Important:
        99 = Unknown number of doses.
        0  = Valid zero-dose count.
    """

    if value is None:
        return True

    if isinstance(value, str):

        normalized = value.strip().casefold()

        if normalized in UNKNOWN_TEXT_CODES:
            return True

    return False


# ============================================================
# 4. EPIDNUMBER NORMALIZATION
# ============================================================

def normalize_epid(value):
    """
    Standardize EpidNumber for matching.

    Original identifiers remain unchanged in Excel.
    """

    if value is None:
        return None

    epid = str(value).strip().upper()

    return epid if epid else None


# ============================================================
# 5. DOSE PARSING AND VALIDATION
# ============================================================

def parse_doses(value):
    """
    Parse AFP vaccination dose counts.

    Returns:
        int:
            Valid nonnegative integer dose count.

        None:
            Missing or unknown dose count.

    Raises:
        ValueError:
            Invalid numeric value.

    Special coding:
        99 = Unknown / information unavailable.
    """

    if is_unknown(value):
        return None

    try:

        number = Decimal(str(value).strip())

    except (InvalidOperation, ValueError, TypeError):

        raise ValueError(
            f"Invalid vaccination dose count: {value!r}"
        )

    if not number.is_finite():

        raise ValueError(
            f"Non-finite vaccination dose count: {value!r}"
        )

    if number < 0:

        raise ValueError(
            f"Negative vaccination dose count: {value!r}"
        )

    if number != number.to_integral_value():

        raise ValueError(
            f"Non-integer vaccination dose count: {value!r}"
        )

    doses = int(number)

    # AFP surveillance coding:
    # 99 = Information unavailable / Unknown.
    if doses in UNKNOWN_NUMERIC_CODES:
        return None

    return doses


def safe_parse(value):
    """
    Parse dose counts without stopping the pipeline.

    Returns:
        (parsed_value, error_message)
    """

    try:

        return parse_doses(value), None

    except ValueError as error:

        return None, str(error)


def is_special_unknown_code(value):
    """
    Detect the explicit surveillance code 99.
    """

    if value is None:
        return False

    try:

        number = Decimal(str(value).strip())

        return (
            number.is_finite()
            and number == 99
        )

    except (InvalidOperation, ValueError, TypeError):

        return False


# ============================================================
# 6. AUDIT LOGGING
# ============================================================

audit_rows = []

stats = Counter()


def add_audit(
    excel_row,
    epid,
    status,
    variable="",
    original_value=None,
    reference_value=None,
    details=""
):
    """
    Record every significant reconciliation decision.
    """

    audit_rows.append([

        excel_row,
        epid,
        status,
        variable,
        original_value,
        reference_value,
        details,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ])


# ============================================================
# 7. VERIFY INPUT FILES
# ============================================================

print("\n" + "=" * 60)

print("AFP VACCINATION STATUS RECONCILIATION")

print("=" * 60)


if not MDB_FILE.is_file():

    raise FileNotFoundError(
        f"Access database not found: {MDB_FILE}"
    )


if not EXCEL_FILE.is_file():

    raise FileNotFoundError(
        f"Excel workbook not found: {EXCEL_FILE}"
    )


print("\nInput files verified.")


# ============================================================
# 8. CONNECT TO MICROSOFT ACCESS
# ============================================================

print("\nConnecting to AFP Access database...")


connection_string = (

    "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={MDB_FILE};"
    "READONLY=1;"

)


connection = pyodbc.connect(
    connection_string
)

cursor = connection.cursor()


# ============================================================
# 9. EXTRACT ACCESS VACCINATION DATA
# ============================================================

query = f"""

SELECT
    EpidNumber,
    TotalIPVdosesRI,
    TotalOPVdosesRI,
    TotalOPVdosesSIA
FROM [{ACCESS_TABLE}]

"""


cursor.execute(query)


access_records = {}

access_rows = 0

access_missing_epid = 0


for row in cursor:

    access_rows += 1

    epid = normalize_epid(
        row.EpidNumber
    )

    if epid is None:

        access_missing_epid += 1

        continue

    record = {

        "TotalIPVdosesRI":
            row.TotalIPVdosesRI,

        "TotalOPVdosesRI":
            row.TotalOPVdosesRI,

        "TotalOPVdosesSIA":
            row.TotalOPVdosesSIA,

    }

    access_records.setdefault(
        epid, []
    ).append(record)


connection.close()


print(
    f"Access records read: {access_rows:,}"
)

print(
    f"Distinct EpidNumber: {len(access_records):,}"
)

print(
    f"Access records without EpidNumber: "
    f"{access_missing_epid:,}"
)


# ============================================================
# 10. LOAD EXCEL WORKBOOK
# ============================================================

print("\nLoading vaccination Excel workbook...")


workbook = load_workbook(
    EXCEL_FILE,
    data_only=False
)


if EXCEL_SHEET not in workbook.sheetnames:

    raise ValueError(
        f"Worksheet '{EXCEL_SHEET}' not found."
    )


worksheet = workbook[EXCEL_SHEET]


headers = {

    str(cell.value).strip(): cell.column

    for cell in worksheet[1]

    if cell.value is not None

}


required_columns = [

    EPID_COLUMN,
    *DOSE_COLUMNS,
    TOTAL_COLUMN,

]


missing_headers = [

    column

    for column in required_columns

    if column not in headers

]


if missing_headers:

    raise ValueError(
        f"Missing Excel columns: {missing_headers}"
    )


print(
    f"Worksheet: {EXCEL_SHEET}"
)

print(
    f"Excel worksheet rows: {worksheet.max_row - 1}"
)


# ============================================================
# 11. IDENTIFY EMPTY ROWS AND DUPLICATE EPIDNUMBERS
# ============================================================

def is_empty_excel_row(row_index):
    """
    Return True if every cell in the row is empty.

    Empty strings and whitespace-only strings
    are also considered empty.

    A row containing any other value is not empty.
    """

    for column_index in range(
        1,
        worksheet.max_column + 1
    ):

        value = worksheet.cell(
            row=row_index,
            column=column_index
        ).value

        if value is None:
            continue

        if isinstance(value, str) and not value.strip():
            continue

        return False

    return True


# Identify real data rows.
data_rows = []

empty_rows = []


for row_index in range(
    2,
    worksheet.max_row + 1
):

    if is_empty_excel_row(row_index):

        empty_rows.append(row_index)

    else:

        data_rows.append(row_index)


stats["empty_rows_skipped"] = len(empty_rows)


print(
    f"Actual Excel data rows: {len(data_rows)}"
)

print(
    f"Empty Excel rows skipped: {len(empty_rows)}"
)


# Count identifiers only for actual data rows.
excel_epids = Counter()


for row_index in data_rows:

    epid = normalize_epid(

        worksheet.cell(
            row=row_index,
            column=headers[EPID_COLUMN]
        ).value

    )

    if epid is not None:

        excel_epids[epid] += 1


# ============================================================
# 12. RECONCILIATION PROCESS
# ============================================================

print("\nReconciling vaccination records...")


for row_index in data_rows:

    stats["rows_processed"] += 1

    epid_cell = worksheet.cell(
        row=row_index,
        column=headers[EPID_COLUMN]
    )

    epid = normalize_epid(
        epid_cell.value
    )

    # --------------------------------------------------------
    # 12.1 IDENTIFIER VALIDATION
    # --------------------------------------------------------

    eligible_for_matching = True

    access_record = None

    if epid is None:

        stats["missing_epid"] += 1

        eligible_for_matching = False

        add_audit(
            row_index,
            None,
            "MISSING_EPID",
            EPID_COLUMN,
            epid_cell.value,
            None,
            "Missing case identifier in a nonempty Excel row"
        )

    elif excel_epids[epid] > 1:

        stats["duplicate_excel_epid"] += 1

        eligible_for_matching = False

        add_audit(
            row_index,
            epid,
            "DUPLICATE_EXCEL_EPID",
            EPID_COLUMN,
            epid,
            None,
            "Duplicate identifier in Excel"
        )

    else:

        matches = access_records.get(
            epid,
            []
        )

        if len(matches) == 0:

            stats["unmatched"] += 1

            eligible_for_matching = False

            add_audit(
                row_index,
                epid,
                "UNMATCHED",
                EPID_COLUMN,
                epid,
                None,
                "No matching Access record"
            )

        elif len(matches) > 1:

            stats["duplicate_access_epid"] += 1

            eligible_for_matching = False

            add_audit(
                row_index,
                epid,
                "DUPLICATE_ACCESS_EPID",
                EPID_COLUMN,
                epid,
                len(matches),
                "Multiple Access records"
            )

        else:

            access_record = matches[0]

            stats["matched"] += 1

    # ========================================================
    # 12.2 RECONCILE AND STANDARDIZE VACCINATION VARIABLES
    # ========================================================

    for excel_field, access_field in FIELD_MAPPING.items():

        cell = worksheet.cell(
            row=row_index,
            column=headers[excel_field]
        )

        original_value = cell.value

        excel_value, excel_error = safe_parse(
            original_value
        )

        # ----------------------------------------------------
        # INVALID EXCEL VALUE
        # ----------------------------------------------------

        if excel_error:

            stats["invalid_excel"] += 1

            add_audit(
                row_index,
                epid,
                "INVALID_EXCEL_VALUE",
                excel_field,
                original_value,
                None,
                excel_error
            )

            # Do not replace invalid values automatically.
            continue

        # ----------------------------------------------------
        # EXISTING VALID NUMERIC VALUE
        # ----------------------------------------------------

        if excel_value is not None:

            # Preserve valid numeric values.
            if not eligible_for_matching:

                continue

            access_raw = access_record[access_field]

            access_value, access_error = safe_parse(
                access_raw
            )

            if access_error:

                stats["invalid_access"] += 1

                add_audit(
                    row_index,
                    epid,
                    "INVALID_ACCESS_VALUE",
                    excel_field,
                    original_value,
                    access_raw,
                    access_error
                )

                continue

            # Access 99 is unknown, not a conflict.
            if access_value is None:

                stats["access_unknown"] += 1

                continue

            if excel_value != access_value:

                stats["conflicts"] += 1

                add_audit(
                    row_index,
                    epid,
                    "CONFLICT",
                    excel_field,
                    original_value,
                    access_raw,
                    "Conflicting known numeric values; "
                    "Excel value preserved"
                )

            else:

                stats["matching_values"] += 1

            continue

        # ----------------------------------------------------
        # UNKNOWN EXCEL VALUE
        # ----------------------------------------------------

        if is_special_unknown_code(original_value):

            stats["unknown_code_99"] += 1

            add_audit(
                row_index,
                epid,
                "UNKNOWN_DOSE_CODE",
                excel_field,
                original_value,
                UNKNOWN_LABEL,
                "Original value 99 means information "
                "about the number of doses was unavailable"
            )

        # ----------------------------------------------------
        # NO SAFE ACCESS MATCH
        # ----------------------------------------------------

        if not eligible_for_matching:

            cell.value = UNKNOWN_LABEL

            stats["standardized_unknown"] += 1

            add_audit(
                row_index,
                epid,
                "UNKNOWN_STANDARDIZED",
                excel_field,
                original_value,
                UNKNOWN_LABEL,
                "No eligible Access match; "
                "unknown value standardized"
            )

            continue

        # ----------------------------------------------------
        # CHECK ACCESS VALUE
        # ----------------------------------------------------

        access_raw = access_record[access_field]

        access_value, access_error = safe_parse(
            access_raw
        )

        if access_error:

            stats["invalid_access"] += 1

            add_audit(
                row_index,
                epid,
                "INVALID_ACCESS_VALUE",
                excel_field,
                original_value,
                access_raw,
                access_error
            )

            # The Excel value remains unknown.
            cell.value = UNKNOWN_LABEL

            stats["standardized_unknown"] += 1

            continue

        # ----------------------------------------------------
        # FILL FROM VALID ACCESS VALUE
        # ----------------------------------------------------

        if access_value is not None:

            cell.value = access_value

            stats["filled"] += 1

            stats[f"filled_{excel_field}"] += 1

            add_audit(
                row_index,
                epid,
                "FILLED",
                excel_field,
                original_value,
                access_value,
                "Unknown Excel value completed using "
                "a valid Access dose count"
            )

        # ----------------------------------------------------
        # UNKNOWN IN BOTH SOURCES
        # ----------------------------------------------------

        else:

            cell.value = UNKNOWN_LABEL

            stats["still_missing"] += 1

            stats["standardized_unknown"] += 1

            add_audit(
                row_index,
                epid,
                "STILL_UNKNOWN",
                excel_field,
                original_value,
                access_raw,
                "Vaccination dose count unavailable "
                "in both sources"
            )

    # ========================================================
    # 13. CALCULATE TOTAL POLIO DOSES
    # ========================================================

    ipv_cell = worksheet.cell(
        row=row_index,
        column=headers[IPV_COLUMN]
    )

    routine_cell = worksheet.cell(
        row=row_index,
        column=headers[ROUTINE_COLUMN]
    )

    sia_cell = worksheet.cell(
        row=row_index,
        column=headers[SIA_COLUMN]
    )

    total_cell = worksheet.cell(
        row=row_index,
        column=headers[TOTAL_COLUMN]
    )

    original_total = total_cell.value

    is_formula = (
        total_cell.data_type == "f"
    )

    is_excel_error = (
        total_cell.data_type == "e"
    )

    ipv, ipv_error = safe_parse(
        ipv_cell.value
    )

    routine, routine_error = safe_parse(
        routine_cell.value
    )

    sia, sia_error = safe_parse(
        sia_cell.value
    )

    # --------------------------------------------------------
    # 13.1 INVALID VACCINATION COMPONENT
    # --------------------------------------------------------

    if ipv_error or routine_error or sia_error:

        stats["invalid_total_inputs"] += 1

        add_audit(
            row_index,
            epid,
            "INVALID_TOTAL_INPUT",
            TOTAL_COLUMN,
            original_total,
            None,
            "At least one vaccination component "
            "contains an invalid value"
        )

        # Invalid component values must not be used
        # in arithmetic. The total is unknown.
        total_cell.value = UNKNOWN_LABEL

        continue

    # --------------------------------------------------------
    # 13.2 UNKNOWN VACCINATION COMPONENT
    # --------------------------------------------------------

    if (
        ipv is None
        or routine is None
        or sia is None
    ):

        stats["incomplete_total"] += 1

        missing_components = []

        if ipv is None:

            missing_components.append(
                "IPV Routine"
            )

        if routine is None:

            missing_components.append(
                "OPV Routine"
            )

        if sia is None:

            missing_components.append(
                "OPV AVS"
            )

        # An unknown component makes the total unknown.
        total_cell.value = UNKNOWN_LABEL

        stats["totals_unknown"] += 1

        add_audit(
            row_index,
            epid,
            "TOTAL_UNKNOWN",
            TOTAL_COLUMN,
            original_total,
            UNKNOWN_LABEL,
            "Cannot calculate total because "
            "the following components are unknown: "
            + ", ".join(missing_components)
        )

        continue

    # --------------------------------------------------------
    # 13.3 ALL COMPONENTS ARE KNOWN
    # --------------------------------------------------------

    calculated_total = (
        ipv
        + routine
        + sia
    )

    # --------------------------------------------------------
    # 13.4 PARSE EXISTING TOTAL
    # --------------------------------------------------------

    if is_formula or is_excel_error:

        existing_total = None

        total_error = None

    else:

        existing_total, total_error = safe_parse(
            original_total
        )

    if total_error:

        stats["invalid_existing_total"] += 1

        add_audit(
            row_index,
            epid,
            "INVALID_EXISTING_TOTAL",
            TOTAL_COLUMN,
            original_total,
            calculated_total,
            total_error
        )

    # --------------------------------------------------------
    # 13.5 DETECT CONFLICTING EXISTING TOTALS
    # --------------------------------------------------------

    if (
        existing_total is not None
        and existing_total != calculated_total
    ):

        stats["total_conflicts"] += 1

        add_audit(
            row_index,
            epid,
            "TOTAL_CONFLICT_CORRECTED",
            TOTAL_COLUMN,
            original_total,
            calculated_total,
            "Existing total differs from the sum "
            "of the three validated components; "
            "calculated total written to output"
        )

    elif existing_total == calculated_total:

        stats["totals_verified"] += 1

    else:

        stats["totals_calculated"] += 1

        add_audit(
            row_index,
            epid,
            "TOTAL_CALCULATED",
            TOTAL_COLUMN,
            original_total,
            calculated_total,
            "IPV Routine + OPV Routine + OPV AVS"
        )

    # Write validated total as numeric value.
    total_cell.value = calculated_total


# ============================================================
# 14. ADD EXPLANATORY EXCEL COMMENTS
# ============================================================

print("\nAdding surveillance coding and calculation comments...")


# ------------------------------------------------------------
# 14.1 TOTAL POLIO DOSES COMMENT
# ------------------------------------------------------------

total_header_cell = worksheet.cell(
    row=1,
    column=headers[TOTAL_COLUMN]
)


total_header_cell.comment = Comment(

    (
        "AFP VACCINATION STATUS - CALCULATION METHODOLOGY\n\n"

        "Total polio doses = "
        "IPV Routine + OPV Routine + OPV AVS.\n\n"

        "COMPONENTS:\n"

        "1. IPV Routine:\n"
        "   TotalIPVdosesRI\n\n"

        "2. OPV Routine:\n"
        "   Total doses VPO Routine\n\n"

        "3. OPV AVS:\n"
        "   Total doses VPO AVS\n\n"

        "AFP SURVEILLANCE CODING:\n"

        "- 99 = Unknown.\n"

        "- 99 means information about the number "
        "of doses was unavailable.\n"

        "- 99 does not mean 99 administered doses.\n"

        "- 0 means zero reported doses and is valid.\n\n"

        "RECONCILIATION RULES:\n"

        "- Unknown values are completed from Access "
        "when valid numeric information is available.\n"

        "- Unresolved unknown values are standardized "
        "to 'Unknown'.\n"

        "- Total doses are calculated only when all "
        "three vaccination components are known.\n"

        "- If any component is unknown, "
        "the total is 'Unknown'.\n"

        "- Existing numeric discrepancies are "
        "documented in the audit report.\n\n"

        "SOURCE:\n"
        "AFP surveillance Access database.\n\n"

        "PROCESSING:\n"
        "AFP Data Engineering reconciliation pipeline."
    ),

    "AFP Data Engineering"

)


total_header_cell.comment.width = 450

total_header_cell.comment.height = 450


# ------------------------------------------------------------
# 14.2 COMMENTS ON VACCINATION COMPONENT HEADERS
# ------------------------------------------------------------

for column_name in DOSE_COLUMNS:

    header_cell = worksheet.cell(
        row=1,
        column=headers[column_name]
    )

    header_cell.comment = Comment(

        (
            "AFP SURVEILLANCE VACCINATION CODING\n\n"

            "0 = Zero reported doses (valid).\n"

            "1, 2, 3... = Known number of doses.\n"

            "99 = Unknown: information about the "
            "number of doses was unavailable.\n\n"

            "During reconciliation, unknown values "
            "are completed using valid Access values "
            "when available.\n\n"

            "Unresolved unknown values are represented "
            "as 'Unknown' in the processed workbook.\n\n"

            "Original values remain documented "
            "in the reconciliation audit."
        ),

        "AFP Data Engineering"

    )

    header_cell.comment.width = 420

    header_cell.comment.height = 300


# ============================================================
# 15. SAVE UPDATED WORKBOOK
# ============================================================

print("\nSaving updated vaccination workbook...")


workbook.save(
    UPDATED_FILE
)


print(
    f"Updated workbook: {UPDATED_FILE}"
)


# ============================================================
# 16. GENERATE AUDIT WORKBOOK
# ============================================================

print("\nGenerating reconciliation audit...")


audit_workbook = Workbook()


# ------------------------------------------------------------
# 16.1 DETAILED AUDIT
# ------------------------------------------------------------

audit_sheet = audit_workbook.active

audit_sheet.title = "Audit"


audit_sheet.append([

    "Excel Row",
    "EpidNumber",
    "Status",
    "Variable",
    "Original Excel Value",
    "Access or Calculated Value",
    "Details",
    "Timestamp",

])


for record in audit_rows:

    audit_sheet.append(record)


# ------------------------------------------------------------
# 16.2 SUMMARY
# ------------------------------------------------------------

summary_sheet = audit_workbook.create_sheet(
    "Summary"
)


summary_sheet.append([
    "Metric",
    "Count"
])


summary_sheet.append([
    "Access records read",
    access_rows
])


summary_sheet.append([
    "Distinct Access EpidNumber",
    len(access_records)
])


summary_sheet.append([
    "Access records without EpidNumber",
    access_missing_epid
])


summary_sheet.append([
    "Excel worksheet rows",
    worksheet.max_row - 1
])


summary_sheet.append([
    "Actual Excel data rows",
    len(data_rows)
])


summary_sheet.append([
    "Empty Excel rows skipped",
    len(empty_rows)
])


for metric, count in sorted(stats.items()):

    summary_sheet.append([
        metric,
        count
    ])


# ------------------------------------------------------------
# 16.3 MISSING VALUES SUMMARY
# ------------------------------------------------------------

missing_sheet = audit_workbook.create_sheet(
    "Missing_Values"
)


missing_sheet.append([
    "Variable",
    "Values Filled",
    "Remaining Unknown"
])


for excel_field in DOSE_COLUMNS:

    filled_count = stats.get(
        f"filled_{excel_field}",
        0
    )

    unknown_count = sum(

        1

        for row_index in data_rows

        if worksheet.cell(
            row=row_index,
            column=headers[excel_field]
        ).value == UNKNOWN_LABEL

    )

    missing_sheet.append([

        excel_field,
        filled_count,
        unknown_count,

    ])


# ------------------------------------------------------------
# 16.4 DATA QUALITY ISSUES
# ------------------------------------------------------------

quality_sheet = audit_workbook.create_sheet(
    "Data_Quality_Issues"
)


quality_sheet.append([

    "Excel Row",
    "EpidNumber",
    "Status",
    "Variable",
    "Original Value",
    "Reference Value",
    "Details",

])


QUALITY_STATUSES = {

    "MISSING_EPID",
    "UNMATCHED",
    "DUPLICATE_EXCEL_EPID",
    "DUPLICATE_ACCESS_EPID",
    "INVALID_EXCEL_VALUE",
    "INVALID_ACCESS_VALUE",
    "CONFLICT",
    "TOTAL_CONFLICT_CORRECTED",
    "INVALID_TOTAL_INPUT",
    "INVALID_EXISTING_TOTAL",
    "TOTAL_UNKNOWN",
    "STILL_UNKNOWN",

}


for record in audit_rows:

    if record[2] in QUALITY_STATUSES:

        quality_sheet.append(
            record[:7]
        )


# ============================================================
# 17. FORMAT AUDIT WORKBOOK
# ============================================================

HEADER_FILL = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

HEADER_FONT = Font(
    color="FFFFFF",
    bold=True
)


for sheet in audit_workbook.worksheets:

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = sheet.dimensions

    for cell in sheet[1]:

        cell.fill = HEADER_FILL

        cell.font = HEADER_FONT

        cell.alignment = Alignment(
            horizontal="center"
        )

    for column in sheet.columns:

        letter = get_column_letter(
            column[0].column
        )

        max_length = max(

            len(str(cell.value))
            if cell.value is not None
            else 0

            for cell in column

        )

        sheet.column_dimensions[
            letter
        ].width = min(
            max(max_length + 2, 14),
            60
        )


audit_workbook.save(
    AUDIT_FILE
)


print(
    f"Audit workbook: {AUDIT_FILE}"
)


# ============================================================
# 18. PRINT RECONCILIATION SUMMARY
# ============================================================

print("\n" + "=" * 60)

print("AFP VACCINATION RECONCILIATION SUMMARY")

print("=" * 60)


summary_metrics = [

    "rows_processed",
    "empty_rows_skipped",
    "matched",
    "unmatched",
    "missing_epid",
    "duplicate_excel_epid",
    "duplicate_access_epid",
    "filled",
    "still_missing",
    "standardized_unknown",
    "unknown_code_99",
    "access_unknown",
    "conflicts",
    "invalid_excel",
    "invalid_access",
    "totals_calculated",
    "totals_verified",
    "total_conflicts",
    "incomplete_total",
    "totals_unknown",
    "invalid_total_inputs",
    "invalid_existing_total",

]


for metric in summary_metrics:

    print(
        f"{metric:<30}: "
        f"{stats.get(metric, 0)}"
    )


print("\nFILLED VACCINATION VARIABLES")

print("-" * 60)


for excel_field in DOSE_COLUMNS:

    count = stats.get(
        f"filled_{excel_field}",
        0
    )

    print(
        f"{excel_field:<30}: {count}"
    )


print("\nREMAINING UNKNOWN VALUES")

print("-" * 60)


for excel_field in DOSE_COLUMNS:

    unknown_count = sum(

        1

        for row_index in data_rows

        if worksheet.cell(
            row=row_index,
            column=headers[excel_field]
        ).value == UNKNOWN_LABEL

    )

    print(
        f"{excel_field:<30}: {unknown_count}"
    )


print("\nOUTPUT FILES")

print("-" * 60)

print(UPDATED_FILE)

print(AUDIT_FILE)


print("\n" + "=" * 60)

print("AFP RECONCILIATION COMPLETED")

print("=" * 60)