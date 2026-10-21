"""excel_exporter.py — Export audited VaultSheet transactions to Excel.

Dependencies:
    pip install pandas openpyxl

The executable example expects:
    bank_parser.parse_bank_pdf(pdf_path) -> pandas.DataFrame
    audit_engine.enrich_and_audit_ledger(df, vendor_master_path)
        -> tuple[pandas.DataFrame, dict]

Excel 2021 or Microsoft 365 is required for XLOOKUP. openpyxl writes
formulas but does not calculate them; Excel recalculates on opening.
"""

from __future__ import annotations

import math
import os
import re
import tempfile
from datetime import date, datetime
from numbers import Integral, Real
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.table import Table, TableStyleInfo


__all__ = ["export_to_excel"]

COLUMNS = [
    "tx_id",
    "date",
    "description",
    "ref_no",
    "debit",
    "credit",
    "balance",
    "math_verified",
    "extraction_method",
    "category",
    "tally_ledger",
    "tally_voucher_type",
    "confidence",
    "anomaly_flag",
    "anomaly_reason",
]

METADATA_KEYS = [
    "total_rows",
    "opening_balance",
    "closing_balance",
    "total_debit",
    "total_credit",
    "net_cash_flow",
    "daily_burn_rate",
    "estimated_runway_days",
    "verification_rate_pct",
    "unverified_row_ids",
    "ocr_pages_count",
    "hidden_bank_fees_total",
    "flagged_anomalies_count",
    "vendor_mismatches_count",
]

# The shared contract does not name its eleven standard categories.
# These defaults may be replaced with metadata["standard_categories"]:
# exactly eleven distinct, nonempty category names. Observed categories
# outside that taxonomy are also included, so no spending is omitted.
STANDARD_CATEGORIES = (
    "Revenue",
    "Vendor Payments",
    "Salaries & Wages",
    "Rent",
    "Utilities",
    "Travel & Entertainment",
    "Marketing & Advertising",
    "Taxes & Statutory Payments",
    "Bank Charges & Fees",
    "Transfers",
    "Other",
)

NAVY = "1B365D"
SLATE = "52657A"
PALE_SLATE = "EAF0F6"
WHITE = "FFFFFF"
GREY = "D9D9D9"
RED = "FADBD8"
AMBER = "FCF3CF"
YELLOW = "FEF9E7"
GREEN = "E2F0D9"

MONEY_FORMAT = '#,##,##0.00;[Red](#,##,##0.00);"-"'
NUMBER_FORMAT = "#,##,##0.00"
INTEGER_FORMAT = "#,##,##0"
PERCENT_FORMAT = "0.00%"
DATE_FORMAT = "dd-mm-yyyy"

THIN_BORDER = Border(
    left=Side(style="thin", color=GREY),
    right=Side(style="thin", color=GREY),
    top=Side(style="thin", color=GREY),
    bottom=Side(style="thin", color=GREY),
)

HEADER_ROW = 6
FIRST_DATA_ROW = 7
MAX_EXCEL_ROWS = 1_048_576


def _missing(value: Any) -> bool:
    if value is None:
        return True
    result = pd.isna(value)
    return bool(result) if isinstance(result, (bool, type(pd.isna(0)))) else False


def _text(value: Any, *, clean_whitespace: bool = False) -> str:
    if _missing(value):
        return ""
    result = ILLEGAL_CHARACTERS_RE.sub("", str(value))
    if clean_whitespace:
        result = re.sub(r"\s+", " ", result).strip()
    if len(result) > 32_767:
        raise ValueError("A text value exceeds Excel's 32,767-character limit.")
    return result


def _number(value: Any, label: str, *, nullable: bool = False) -> float | None:
    if _missing(value):
        if nullable:
            return None
        raise ValueError(f"{label} must contain a finite number.")
    if isinstance(value, bool):
        raise ValueError(f"{label} must be numeric, not Boolean.")
    try:
        result = float(value)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(f"{label} must contain a finite number.") from exc
    if not math.isfinite(result):
        raise ValueError(f"{label} must contain a finite number.")
    return result


def _boolean(value: Any, label: str) -> bool:
    if _missing(value):
        raise ValueError(f"{label} cannot be missing.")
    if isinstance(value, str):
        token = value.strip().casefold()
        if token in {"true", "1", "yes"}:
            return True
        if token in {"false", "0", "no"}:
            return False
    elif value == True:
        return True
    elif value == False:
        return False
    raise ValueError(f"{label} must be True or False.")


def _date_value(value: Any, label: str) -> datetime:
    if _missing(value):
        raise ValueError(f"{label} cannot be missing.")
    if isinstance(value, Real) and not isinstance(value, (date, datetime)):
        raise ValueError(f"{label}: numeric date serials are not accepted.")
    try:
        parsed = pd.to_datetime(value, dayfirst=True, errors="raise")
        if not isinstance(parsed, pd.Timestamp) or pd.isna(parsed):
            raise ValueError("Invalid scalar date.")
        if parsed.tzinfo is not None:
            parsed = parsed.tz_localize(None)
        result = datetime(parsed.year, parsed.month, parsed.day)
    except (ValueError, TypeError, OverflowError) as exc:
        raise ValueError(f"{label} contains an invalid date: {value!r}.") from exc
    if result.year < 1900:
        raise ValueError(f"{label} precedes Excel's supported date range.")
    return result


def _put(ws, row: int, column: int, value: Any, *, formula: bool = False):
    """Write imported text literally, including strings starting with '='."""
    cell = ws.cell(row=row, column=column)
    cell.value = value
    if isinstance(value, str) and not formula:
        cell.data_type = "s"
    cell.font = Font(name="Calibri", size=11, color=NAVY)
    cell.border = THIN_BORDER
    cell.alignment = Alignment(vertical="center")
    return cell


def _formula(ws, row: int, column: int, expression: str, fmt: str | None = None):
    cell = _put(ws, row, column, expression, formula=True)
    if fmt:
        cell.number_format = fmt
    return cell


def _band(ws, row: int, start: int, end: int, text: str, color: str = NAVY):
    for column in range(start, end + 1):
        cell = _put(ws, row, column, None)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    ws.merge_cells(
        start_row=row, start_column=start, end_row=row, end_column=end
    )
    cell = ws.cell(row, start)
    cell.value = text
    cell.data_type = "s"
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 27
    return cell


def _setup_sheet(ws, title: str, subtitle: str, columns: int):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = NAVY
    ws.freeze_panes = "C7"
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.sheet_format.defaultRowHeight = 20
    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:6"
    ws.oddFooter.center.text = "VaultSheet | Page &P of &N"
    ws.oddFooter.center.size = 9
    ws.oddFooter.right.text = "Confidential"
    _band(ws, 1, 1, columns, title)
    ws.cell(1, 1).font = Font(
        name="Calibri", size=17, bold=True, color=WHITE
    )
    ws.row_dimensions[1].height = 36
    _band(ws, 2, 1, columns, subtitle, SLATE)


def _headers(ws, labels: list[str], row: int = HEADER_ROW, start: int = 1):
    for column, label in enumerate(labels, start):
        cell = _put(ws, row, column, label)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 34


def _add_table(ws, name: str, columns: int, last_row: int):
    ref = f"A{HEADER_ROW}:{get_column_letter(columns)}{last_row}"
    ws.auto_filter.ref = ref
    if last_row > HEADER_ROW:
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)


def _risk_formatting(ws, last_row: int, last_column: int):
    if last_row < FIRST_DATA_ROW:
        return
    area = f"A{FIRST_DATA_ROW}:{get_column_letter(last_column)}{last_row}"
    ws.conditional_formatting.add(
        area,
        FormulaRule(
            formula=[f"$N{FIRST_DATA_ROW}=TRUE"],
            fill=PatternFill("solid", fgColor=RED),
            stopIfTrue=True,
        ),
    )
    ws.conditional_formatting.add(
        area,
        FormulaRule(
            formula=[f"AND($A{FIRST_DATA_ROW}<>\"\",$H{FIRST_DATA_ROW}=FALSE)"],
            fill=PatternFill("solid", fgColor=AMBER),
            stopIfTrue=True,
        ),
    )


def _write_ledger_row(ws, row_number: int, record: list[Any]):
    for column, value in enumerate(record, 1):
        cell = _put(ws, row_number, column, value)
        if column == 2:
            cell.number_format = DATE_FORMAT
        elif column in (5, 6, 7, 13):
            cell.number_format = NUMBER_FORMAT
        elif column in (1, 4):
            cell.number_format = "@"
        if column in (3, 15):
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[row_number].height = 32


def _autofit(ws, *, start_row: int = HEADER_ROW):
    """Approximate autofit; openpyxl cannot invoke Excel's layout engine."""
    for column in range(1, ws.max_column + 1):
        width = 12
        for row in ws.iter_rows(
            min_row=start_row, min_col=column, max_col=column
        ):
            cell = row[0]
            if cell.value is None:
                continue
            if cell.data_type == "f":
                length = 20
            elif isinstance(cell.value, (date, datetime)):
                length = 12
            elif isinstance(cell.value, (int, float)):
                length = len(f"{cell.value:,.2f}") + 2
            else:
                length = max(
                    (len(line) for line in str(cell.value).splitlines()),
                    default=0,
                ) + 2
            width = max(width, min(length, 55))
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"



# ==============================================================================
# EXCEL 2024 OPENXML SANITIZER & ANTI-OVERLAP PATCH
# ==============================================================================
import math as _math
import re as _re

_ILLEGAL_XML_RE = _re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")

def _clean_excel_str(val) -> str:
    if val is None:
        return ""
    s = str(val)
    if s.lower() in ("nan", "none", "nat", "<na>"):
        return ""
    # 1. Strip illegal OpenXML control characters that trigger Excel 2024 repair popup
    s = _ILLEGAL_XML_RE.sub(" ", s)
    # 2. Collapse embedded newlines/tabs into a single space to prevent overlapping lines
    s = _re.sub(r"[\r\n\t]+", " | ", s)
    s = _re.sub(r"\s{2,}", " ", s).strip()
    # 3. Prevent Excel 2024 from misinterpreting text starting with =, +, -, @ as broken formulas
    if s.startswith(("=", "+", "-", "@")):
        s = "'" + s
    return s[:32000]

def _clean_excel_num(val) -> float:
    try:
        f = float(val)
        return 0.0 if (_math.isnan(f) or _math.isinf(f)) else round(f, 2)
    except Exception:
        return 0.0


def _validate_and_prepare(df: pd.DataFrame, metadata: dict):
    if not isinstance(df, pd.DataFrame):
        raise TypeError("df must be a pandas DataFrame.")
    if not isinstance(metadata, dict):
        raise TypeError("metadata must be a dictionary.")
    if list(df.columns) != COLUMNS:
        raise ValueError(f"df columns must match this exact order: {COLUMNS!r}")
    missing_keys = [key for key in METADATA_KEYS if key not in metadata]
    if missing_keys:
        raise ValueError(f"Missing metadata keys: {missing_keys!r}")
    if len(df) > MAX_EXCEL_ROWS - HEADER_ROW:
        raise ValueError("The ledger exceeds Excel's worksheet row limit.")

    normalized = dict(metadata)
    numeric_keys = [
        key for key in METADATA_KEYS
        if key not in {"unverified_row_ids", "estimated_runway_days"}
    ]
    for key in numeric_keys:
        normalized[key] = _number(metadata[key], f"metadata[{key!r}]")

    for key in (
        "total_rows", "ocr_pages_count",
        "flagged_anomalies_count", "vendor_mismatches_count",
    ):
        value = normalized[key]
        if value < 0 or not value.is_integer():
            raise ValueError(f"metadata[{key!r}] must be a nonnegative integer.")

    if normalized["total_rows"] != len(df):
        raise ValueError("metadata['total_rows'] does not match len(df).")
    if not 0 <= normalized["verification_rate_pct"] <= 100:
        raise ValueError("verification_rate_pct must be between 0 and 100.")

    runway = metadata["estimated_runway_days"]
    if _missing(runway):
        normalized["estimated_runway_days"] = None
    else:
        try:
            runway_number = float(runway)
        except (ValueError, TypeError, OverflowError) as exc:
            raise ValueError("estimated_runway_days must be numeric or null.") from exc
        if runway_number < 0 or math.isnan(runway_number):
            raise ValueError("estimated_runway_days cannot be negative or NaN.")
        normalized["estimated_runway_days"] = (
            runway_number if math.isfinite(runway_number) else None
        )

    if not isinstance(metadata["unverified_row_ids"], (list, tuple, set)):
        raise ValueError("unverified_row_ids must be a list, tuple, or set.")

    categories = metadata.get("standard_categories", STANDARD_CATEGORIES)
    if not isinstance(categories, (list, tuple)) or len(categories) != 11:
        raise ValueError("standard_categories must contain exactly eleven names.")
    categories = [_text(value).strip() for value in categories]
    if any(not value for value in categories):
        raise ValueError("Standard category names cannot be empty.")
    if len({value.casefold() for value in categories}) != 11:
        raise ValueError("Standard category names must be distinct.")

    records = []
    seen_ids = set()
    for ordinal, raw in enumerate(df.itertuples(index=False, name=None), 1):
        record = list(raw)
        label = f"Transaction {ordinal}"

        record[0] = _text(record[0]).strip()
        if not record[0]:
            raise ValueError(f"{label}: tx_id cannot be empty.")
        if record[0] in seen_ids:
            raise ValueError(f"{label}: duplicate tx_id {record[0]!r}.")
        seen_ids.add(record[0])
        record[1] = _date_value(record[1], f"{label}: date")

        for index in (2, 3, 8, 9, 10, 11, 14):
            record[index] = _text(record[index])
        if not record[9].strip():
            raise ValueError(f"{label}: category cannot be empty.")

        for index in (4, 5, 6):
            record[index] = _number(record[index], f"{label}: {COLUMNS[index]}")
        record[12] = _number(
            record[12], f"{label}: confidence", nullable=True
        )
        record[7] = _boolean(record[7], f"{label}: math_verified")
        record[13] = _boolean(record[13], f"{label}: anomaly_flag")
        records.append(record)

    observed = sorted({record[9] for record in records}, key=str.casefold)
    category_keys = {value.casefold() for value in categories}
    for category in observed:
        if category.casefold() not in category_keys:
            categories.append(category)
            category_keys.add(category.casefold())

    return records, normalized, categories


def export_to_excel(df: pd.DataFrame, metadata: dict, output_path: str) -> str:
    """Create the four-tab workbook and return its absolute file path.

    Imported values are written literally to prevent formula injection.
    Source transaction order is preserved. Metadata is treated as the
    upstream audit snapshot; summary calculations remain live in Excel.
    The destination is replaced atomically after a successful save.
    """
    records, audit, categories = _validate_and_prepare(df, metadata)
    destination = Path(output_path).expanduser().resolve()
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("output_path must end in '.xlsx'.")
    destination.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ledger = wb.active
    ledger.title = "Clean_Ledger"
    summary = wb.create_sheet("Executive_Summary")
    risks = wb.create_sheet("Risk_And_Anomalies")
    tally = wb.create_sheet("Tally_Prime_Import")
    wb.properties.creator = "VaultSheet"
    wb.properties.title = "VaultSheet Audited Financial Ledger"
    wb.properties.subject = "Executive summary, audit exceptions and voucher mapping"
    wb.calculation = CalcProperties(
        calcId=0,
        calcMode="auto",
        fullCalcOnLoad=True,
        forceFullCalc=True,
    )

    actual_last = HEADER_ROW + len(records)
    # A single blank source row makes empty-ledger formulas valid.
    formula_last = max(FIRST_DATA_ROW, actual_last)

    def source(column: str) -> str:
        return (
            f"'Clean_Ledger'!${column}${FIRST_DATA_ROW}:"
            f"${column}${formula_last}"
        )

    ids = source("A")
    dates = source("B")
    descriptions = source("C")
    debits = source("E")
    credits = source("F")
    verified = source("H")
    category_range = source("J")
    anomalies = source("N")

    _setup_sheet(
        ledger,
        "VAULTSHEET — ZERO-KNOWLEDGE AUDIT LEDGER",
        "Audited source ledger | INR | Red: anomaly | Amber: unverified arithmetic",
        15,
    )

    kpis = [
        ("Opening Balance", audit["opening_balance"], False, MONEY_FORMAT),
        ("Total Debits", f"=SUM(E7:E{formula_last})", True, MONEY_FORMAT),
        ("Total Credits", f"=SUM(F7:F{formula_last})", True, MONEY_FORMAT),
        ("Closing Balance", audit["closing_balance"], False, MONEY_FORMAT),
        (
            "Math Verification %",
            f'=IFERROR(COUNTIFS(H7:H{formula_last},TRUE,'
            f'A7:A{formula_last},"<>")/COUNTA(A7:A{formula_last}),0)',
            True,
            PERCENT_FORMAT,
        ),
        ("Hidden Bank Fees", audit["hidden_bank_fees_total"], False, MONEY_FORMAT),
        ("Daily Burn Rate", audit["daily_burn_rate"], False, MONEY_FORMAT),
    ]
    for index, (label, value, is_formula, fmt) in enumerate(kpis):
        column = 1 + index * 2
        _band(ledger, 3, column, column + 1, label, SLATE)
        for offset in (0, 1):
            cell = _put(ledger, 4, column + offset, None)
            cell.fill = PatternFill("solid", fgColor=PALE_SLATE)
        ledger.merge_cells(
            start_row=4, start_column=column,
            end_row=4, end_column=column + 1,
        )
        cell = _put(ledger, 4, column, value, formula=is_formula)
        cell.number_format = fmt
        cell.font = Font(name="Calibri", size=15, bold=True, color=NAVY)
    ledger.row_dimensions[4].height = 32

    _headers(ledger, COLUMNS)
    for row_number, record in enumerate(records, FIRST_DATA_ROW):
        _write_ledger_row(ledger, row_number, record)
    _add_table(ledger, "VaultSheetLedger", 15, actual_last)
    _risk_formatting(ledger, actual_last, 15)

    _setup_sheet(
        summary,
        "VAULTSHEET — EXECUTIVE SUMMARY",
        "Live Excel formulas | INR | Audit metrics reference the supplied audit snapshot",
        7,
    )
    _band(summary, 5, 2, 6, "A  |  CATEGORY CASH FLOW")
    _headers(
        summary,
        ["Category", "Tx Count", "Total Spend (Debit)",
         "Total Inflow (Credit)", "% of Spend"],
        start=2,
    )

    for row_number, category in enumerate(categories, FIRST_DATA_ROW):
        _put(summary, row_number, 2, category)
        # Escape Excel criteria wildcards so category names match literally.
        criteria = (
            f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B{row_number},'
            f'"~","~~"),"*","~*"),"?","~?")'
        )
        _formula(
            summary, row_number, 3,
            f"=COUNTIFS({category_range},{criteria})",
            INTEGER_FORMAT,
        )
        _formula(
            summary, row_number, 4,
            f"=SUMIFS({debits},{category_range},{criteria})",
            MONEY_FORMAT,
        )
        _formula(
            summary, row_number, 5,
            f"=SUMIFS({credits},{category_range},{criteria})",
            MONEY_FORMAT,
        )
        _formula(
            summary, row_number, 6,
            f"=IFERROR(D{row_number}/SUM({debits}),0)",
            PERCENT_FORMAT,
        )

    total_row = FIRST_DATA_ROW + len(categories)
    _put(summary, total_row, 2, "TOTAL")
    for column in (3, 4, 5, 6):
        letter = get_column_letter(column)
        _formula(
            summary, total_row, column,
            f"=SUM({letter}{FIRST_DATA_ROW}:{letter}{total_row - 1})",
            INTEGER_FORMAT if column == 3 else
            PERCENT_FORMAT if column == 6 else MONEY_FORMAT,
        )
    for cell in summary[total_row][1:6]:
        cell.fill = PatternFill("solid", fgColor=PALE_SLATE)
        cell.font = Font(name="Calibri", size=11, bold=True, color=NAVY)

    section_b = total_row + 3
    _band(summary, section_b, 2, 6, "B  |  TOP EXPENSE INSPECTOR")
    max_row = section_b + 1
    narration_row = section_b + 2
    date_row = section_b + 3

    _put(summary, max_row, 2, "Largest Single Debit Outflow")
    _formula(summary, max_row, 3, f"=MAX({debits})", MONEY_FORMAT)
    _put(summary, narration_row, 2, "Exact Narration")
    summary.merge_cells(
        start_row=narration_row, start_column=3,
        end_row=narration_row, end_column=6,
    )
    _formula(
        summary, narration_row, 3,
        f'=IF(C{max_row}<=0,"None",'
        f'_xlfn.XLOOKUP(C{max_row},{debits},{descriptions},"None",0))',
    )
    summary.cell(narration_row, 3).alignment = Alignment(
        wrap_text=True, vertical="top"
    )
    summary.row_dimensions[narration_row].height = 48
    _put(summary, date_row, 2, "Transaction Date")
    _formula(
        summary, date_row, 3,
        f'=IF(C{max_row}<=0,"None",'
        f'_xlfn.XLOOKUP(C{max_row},{debits},{dates},"None",0))',
        DATE_FORMAT,
    )
    _band(
        summary, section_b + 4, 2, 6,
        "Equal largest debits: the first transaction in source order is shown.",
        SLATE,
    )

    section_c = section_b + 7
    _band(
        summary, section_c, 2, 6,
        "C  |  BURN RATE, HIDDEN FEES & MATHEMATICAL RECONCILIATION",
    )
    metric_rows = {}
    metrics = [
        ("Daily Burn Rate", "='Clean_Ledger'!M4", MONEY_FORMAT),
        (
            "Estimated Cash Runway (Days)",
            f'=IF(C{section_c + 1}>0,'
            f'MAX(0,\'Clean_Ledger\'!G4)/C{section_c + 1},"N/A")',
            NUMBER_FORMAT,
        ),
        ("Total Hidden Bank Fees", "='Clean_Ledger'!K4", MONEY_FORMAT),
        ("Opening Balance", "='Clean_Ledger'!A4", MONEY_FORMAT),
        ("Total Credits", f"=SUM({credits})", MONEY_FORMAT),
        ("Total Debits", f"=SUM({debits})", MONEY_FORMAT),
        ("Closing Balance", "='Clean_Ledger'!G4", MONEY_FORMAT),
        ("Net Cash Flow", f"=SUM({credits})-SUM({debits})", MONEY_FORMAT),
    ]
    for offset, (label, expression, fmt) in enumerate(metrics, 1):
        row = section_c + offset
        metric_rows[label] = row
        _put(summary, row, 2, label)
        _formula(summary, row, 3, expression, fmt)

    opening_row = metric_rows["Opening Balance"]
    credit_row = metric_rows["Total Credits"]
    debit_row = metric_rows["Total Debits"]
    closing_row = metric_rows["Closing Balance"]
    difference_row = section_c + 9
    reconciliation_row = section_c + 10

    _put(summary, difference_row, 2, "Reconciliation Difference")
    _formula(
        summary, difference_row, 3,
        f"=(C{opening_row}+C{credit_row}-C{debit_row})-C{closing_row}",
        MONEY_FORMAT,
    )
    _put(summary, reconciliation_row, 2, "Mathematical Reconciliation")
    summary.merge_cells(
        start_row=reconciliation_row, start_column=3,
        end_row=reconciliation_row, end_column=6,
    )
    _formula(
        summary, reconciliation_row, 3,
        f'=IF(ABS((C{opening_row}+C{credit_row}-C{debit_row})'
        f'-C{closing_row})<0.01,'
        '"RECONCILED (0.00 DIFF)","DISCREPANCY DETECTED")',
    )
    summary.conditional_formatting.add(
        f"C{reconciliation_row}:F{reconciliation_row}",
        FormulaRule(
            formula=[f"ABS($C${difference_row})<0.01"],
            fill=PatternFill("solid", fgColor=GREEN),
        ),
    )
    summary.conditional_formatting.add(
        f"C{reconciliation_row}:F{reconciliation_row}",
        FormulaRule(
            formula=[f"ABS($C${difference_row})>=0.01"],
            fill=PatternFill("solid", fgColor=RED),
        ),
    )

    for offset, label, expression, fmt in [
        (12, "Transaction Count", f'=COUNTIFS({ids},"<>")', INTEGER_FORMAT),
        (
            13, "Math Verification %",
            f'=IFERROR(COUNTIFS({verified},TRUE,{ids},"<>")/'
            f'COUNTIFS({ids},"<>"),0)',
            PERCENT_FORMAT,
        ),
        (
            14, "Unverified Transactions",
            f'=COUNTIFS({verified},FALSE,{ids},"<>")',
            INTEGER_FORMAT,
        ),
        (
            15, "Flagged Anomalies",
            f'=COUNTIFS({anomalies},TRUE,{ids},"<>")',
            INTEGER_FORMAT,
        ),
    ]:
        _put(summary, section_c + offset, 2, label)
        _formula(summary, section_c + offset, 3, expression, fmt)

    _band(
        summary, section_c + 17, 2, 6,
        "Runway = nonnegative closing cash / supplied daily burn; "
        "N/A when burn is zero or negative. Hidden fees and burn are audit inputs.",
        SLATE,
    )
    summary.row_dimensions[section_c + 17].height = 38

    _setup_sheet(
        risks,
        "VAULTSHEET — RISK & ANOMALIES",
        "Audit exceptions snapshot | Enter review decisions and sign-off in the yellow column",
        16,
    )
    risk_records = [record for record in records if record[13] or not record[7]]
    _band(
        risks, 4, 1, 16,
        f"{len(risk_records):,} transactions require review | "
        "Regenerate this tab after changing source audit flags.",
        SLATE,
    )
    _headers(
        risks,
        COLUMNS + ["Chartered Accountant / Auditor Remarks"],
    )
    for row_number, record in enumerate(risk_records, FIRST_DATA_ROW):
        _write_ledger_row(risks, row_number, record)
        remarks = _put(risks, row_number, 16, "")
        remarks.fill = PatternFill("solid", fgColor=YELLOW)
        remarks.alignment = Alignment(wrap_text=True, vertical="top")
        risks.row_dimensions[row_number].height = 44
    risk_last = HEADER_ROW + len(risk_records)
    _add_table(risks, "VaultSheetRisks", 16, risk_last)
    # Limit conditional formatting to A:O so remarks remain yellow.
    _risk_formatting(risks, risk_last, 15)

    _setup_sheet(
        tally,
        "VAULTSHEET — TALLYPRIME VOUCHER IMPORT",
        "Map row 6 as headers and row 7 onward as data | "
        "Use ledger names configured in the destination company",
        8,
    )
    tally_headers = [
        "Date (DD-MM-YYYY)",
        "Voucher Type",
        "Party / Ledger Name",
        "Bank Ledger Name",
        "Instrument / UTR No",
        "Debit Amount",
        "Credit Amount",
        "Narration",
    ]
    _headers(tally, tally_headers)
    for row_number, record in enumerate(records, FIRST_DATA_ROW):
        values = [
            record[1].strftime("%d-%m-%Y"),
            record[11],
            record[10],
            "Primary Bank A/c",
            record[3],
            record[4],
            record[5],
            _text(record[2], clean_whitespace=True),
        ]
        for column, value in enumerate(values, 1):
            cell = _put(tally, row_number, column, value)
            if column in (6, 7):
                cell.number_format = NUMBER_FORMAT
            elif column in (1, 5):
                cell.number_format = "@"
            elif column == 8:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        tally.row_dimensions[row_number].height = 32
    _add_table(tally, "VaultSheetTallyImport", 8, actual_last)

    for ws in wb:
        _autofit(ws)

    ledger.column_dimensions["C"].width = 55
    ledger.column_dimensions["O"].width = 55
    risks.column_dimensions["C"].width = 55
    risks.column_dimensions["O"].width = 55
    risks.column_dimensions["P"].width = 55
    tally.column_dimensions["H"].width = 65
    summary.column_dimensions["A"].width = 3
    summary.column_dimensions["B"].width = 38
    for letter in ("C", "D", "E", "F"):
        summary.column_dimensions[letter].width = 24
    summary.column_dimensions["G"].width = 3

    ledger.sheet_properties.tabColor = NAVY
    summary.sheet_properties.tabColor = SLATE
    risks.sheet_properties.tabColor = "B9770E"
    tally.sheet_properties.tabColor = "2874A6"
    wb.active = 1

    temporary_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{destination.stem}_",
            suffix=".xlsx",
            dir=str(destination.parent),
            delete=False,
        ) as temporary:
            temporary_path = temporary.name
        wb.save(temporary_path)
        wb.close()
        os.replace(temporary_path, destination)
        temporary_path = None
    finally:
        wb.close()
        if temporary_path and os.path.exists(temporary_path):
            os.unlink(temporary_path)

    return str(destination)


if __name__ == "__main__":
    from pathlib import Path
    from parser_engine import parse_bank_pdf
    from categorizer_and_auditor import enrich_and_audit_ledger

    statement_path = Path("./test_statements/bank_4_messy_headers.pdf")
    vendor_master_path = Path("./test_statements/vendor_master.csv")
    workbook_path = Path("./test_statements/audit_output.xlsx")

    for required_path in (statement_path, vendor_master_path):
        if not required_path.is_file():
            raise FileNotFoundError(f"Required input file not found: {required_path}")

    parsed_df, parse_metadata = parse_bank_pdf(str(statement_path))
    audited_df, audit_metadata = enrich_and_audit_ledger(
        parsed_df,
        parse_metadata,
        vendor_csv_path=str(vendor_master_path),
    )

    saved_path = export_to_excel(
        audited_df,
        audit_metadata,
        str(workbook_path),
    )
    print(f"Created audited workbook: {saved_path}")