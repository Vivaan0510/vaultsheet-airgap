#!/usr/bin/env python3
"""
parser_engine.py
----------------
Production-grade, 100% offline Indian Bank Statement Parser Engine.
Supports RAM-only decryption, 3-mode extraction (Native Grid, Coordinate Stream,
Offline RapidOCR), Indian numbering system normalization, multi-line narration
stitching, and self-healing running-balance mathematical verification.

Dependencies:
    pip install pdfplumber pypdf pypdfium2 rapidocr_onnxruntime pandas
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
from datetime import datetime
from typing import Any

import numpy as np
import pandas as pd
import pdfplumber
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
from rapidocr_onnxruntime import RapidOCR


# ==============================================================================
# 1. REGEX PATTERNS & NORMALIZATION UTILITIES
# ==============================================================================

DATE_PATTERNS = [
    (re.compile(r"\b(\d{2}/\d{2}/\d{4})\b"), "%d/%m/%Y"),
    (re.compile(r"\b(\d{2}-\d{2}-\d{4})\b"), "%d-%m-%Y"),
    (re.compile(r"\b(\d{2}-[A-Za-z]{3}-\d{4})\b"), "%d-%b-%Y"),
    (re.compile(r"\b(\d{2}\s+[A-Za-z]{3}\s+\d{4})\b"), "%d %b %Y"),
    (re.compile(r"\b(\d{2}\s+[A-Za-z]{3}\s+\d{2})\b"), "%d %b %y"),
    (re.compile(r"\b(\d{4}-\d{2}-\d{2})\b"), "%Y-%m-%d"),
]

# Matches Indian & International formatted decimals with optional Dr/Cr suffix
AMOUNT_TOKEN_RE = re.compile(
    r"(?<![A-Za-z0-9/])"
    r"(-?\d{1,3}(?:,\d{2,3})*\.\d{2}|-?\d+\.\d{2})"
    r"(?:\s*\(?(Dr|Cr|DR|CR|dr|cr)\)?)?"
    r"(?![A-Za-z0-9/])"
)

REF_TOKEN_RE = re.compile(r"\b([A-Z]{1,4}\d{6,14}|\d{6,14})\b")

NON_TX_LINE_PATTERNS = [
    re.compile(r"page\s+\d+\s+of\s+\d+", re.IGNORECASE),
    re.compile(r"system-generated|does not require|physical signature", re.IGNORECASE),
    re.compile(r"important regulatory|statutory information", re.IGNORECASE),
    re.compile(r"discrepancies in this statement|dicgc insures", re.IGNORECASE),
    re.compile(r"abbreviations:\s*upi|legend:\s*\(dr\)", re.IGNORECASE),
    re.compile(r"certified true scanned copy|archival flatbed", re.IGNORECASE),
    re.compile(r"executive account relationship summary", re.IGNORECASE),
    re.compile(r"sanctioned od limit|uncleared cheques|registered address", re.IGNORECASE),
    re.compile(r"scroll down to mid-page|statement period:", re.IGNORECASE),
    re.compile(r"branch:.*ifsc:|customer name:|account holder:|account name:", re.IGNORECASE),
    re.compile(r"corporate current account|detailed account ledger|statement of account", re.IGNORECASE),
    re.compile(r"corporate financial summary|privy corporate statement|scanned voucher statement", re.IGNORECASE),
    re.compile(r"^date\s+narration|^txn date\s+description|^date\s+transaction particulars", re.IGNORECASE),
    re.compile(r"opening balance\s*\(inr\)\s+total debits", re.IGNORECASE),
]

OPENING_BAL_RE = re.compile(
    r"\b(b/f\s+opening\s+balance|opening\s+balance\s+brought\s+forward|opening\s+balance)\b",
    re.IGNORECASE,
)


def parse_date_str(text: str) -> tuple[str | None, tuple[int, int] | None]:
    """Extract and normalize a date string to YYYY-MM-DD along with its character span."""
    if not text:
        return None, None
    # Common OCR character substitutions inside date tokens
    cleaned = text.strip()
    for regex, fmt in DATE_PATTERNS:
        m = regex.search(cleaned)
        if m:
            raw_dt = m.group(1)
            try:
                dt = datetime.strptime(raw_dt, fmt)
                if 1990 <= dt.year <= 2100:
                    return dt.strftime("%Y-%m-%d"), m.span()
            except ValueError:
                continue
    return None, None


def parse_indian_amount(val_str: str | None) -> tuple[float | None, str | None]:
    """
    Parse an Indian-formatted currency string into (float_amount, 'DR'|'CR'|None).
    Examples:
        '1,25,450.75'      -> (125450.75, None)
        '14,500.00 (Dr)'   -> (14500.0, 'DR')
        '85,000.00 Cr'     -> (85000.0, 'CR')
        '-' or ''          -> (None, None)
    """
    if val_str is None:
        return None, None
    s = str(val_str).strip()
    if not s or s in ("-", "--", "N/A", "nil", "NIL"):
        return None, None

    # Normalize common OCR artifacts in numbers (e.g. O -> 0 when adjacent to digits)
    s = re.sub(r"(?<=\d)[Oo](?=\d|\.)", "0", s)
    s = re.sub(r"(?<=\.)[Oo](?=\d)", "0", s)

    drcr = None
    if re.search(r"\b(dr)\b|\(dr\)", s, re.IGNORECASE):
        drcr = "DR"
    elif re.search(r"\b(cr)\b|\(cr\)", s, re.IGNORECASE):
        drcr = "CR"

    m = AMOUNT_TOKEN_RE.search(s)
    if not m:
        return None, drcr

    num_clean = m.group(1).replace(",", "")
    if m.group(2) and not drcr:
        drcr = m.group(2).upper()

    try:
        val = round(float(num_clean), 2)
        return abs(val), drcr
    except ValueError:
        return None, drcr


def is_noise_or_header_line(text: str) -> bool:
    """Detect repeating bank headers, footers, legends, or summary banners."""
    s = " ".join(text.split()).strip()
    if not s:
        return True
    for pat in NON_TX_LINE_PATTERNS:
        if pat.search(s):
            return True
    return False


def infer_direction_from_narration(description: str, ref_no: str = "") -> str:
    """Heuristic fallback to infer DR vs CR from narration/reference tokens."""
    u = f"{description} {ref_no}".upper()
    if any(
        tok in u
        for tok in (
            "/CR/",
            "NEFT/CR",
            "IMPS/CR",
            "ACH/CR",
            "RTGS/CR",
            "CREDIT INTEREST",
            "INT.PD",
            "SETTLEMENT",
            "PAYOUT",
        )
    ):
        if "/DR/" not in u:
            return "CR"
    if u.strip().startswith("CR") and not u.strip().startswith("CREDIT CARD"):
        return "CR"
    return "DR"


def extract_ref_from_description(desc: str) -> str:
    """Extract UTR / NEFT / IMPS / Fee reference code from narration if column was merged."""
    if not desc:
        return ""
    # Check structured slash tokens like UPI/DR/4129837192/... or NEFT/DR/N261998210/...
    slash_match = re.search(
        r"\b(?:UPI|NEFT|RTGS|IMPS)/(?:DR|CR)/([A-Za-z0-9]{6,16})\b", desc, re.IGNORECASE
    )
    if slash_match:
        return slash_match.group(1)

    m = REF_TOKEN_RE.search(desc)
    if m:
        candidate = m.group(1)
        # Avoid matching IFSC codes (4 letters + 0 + 6 alphanum) or years
        if not re.match(r"^[A-Z]{4}0[A-Z0-9]{6}$", candidate):
            return candidate
    return ""


# ==============================================================================
# 2. RAM-ONLY PDF DECRYPTION
# ==============================================================================

def load_decrypted_pdf_bytes(pdf_path: str | bytes, password: str | None = None) -> io.BytesIO:
    """
    Accepts a file path (str) or raw bytes, decrypts in volatile RAM if locked,
    and returns an in-memory io.BytesIO stream without touching disk.
    """
    if isinstance(pdf_path, (bytes, bytearray)):
        raw_stream = io.BytesIO(pdf_path)
    elif isinstance(pdf_path, str):
        with open(pdf_path, "rb") as f:
            raw_stream = io.BytesIO(f.read())
    else:
        raise TypeError("pdf_path must be a file path (str) or raw bytes.")

    raw_stream.seek(0)
    reader = PdfReader(raw_stream)

    if not reader.is_encrypted:
        raw_stream.seek(0)
        return raw_stream

    # Try user-supplied password or empty password
    candidates = [password] if password is not None else [""]
    decrypted = False
    for pw in candidates:
        if pw is None:
            continue
        try:
            res = reader.decrypt(pw)
            if res > 0:
                decrypted = True
                break
        except Exception:
            continue

    if not decrypted:
        raise ValueError("PDF is password-protected and decryption failed with the provided password.")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    out_stream = io.BytesIO()
    writer.write(out_stream)
    out_stream.seek(0)
    return out_stream


# ==============================================================================
# 3. THREE-MODE PAGE EXTRACTION PIPELINE
# ==============================================================================

_OCR_ENGINE: RapidOCR | None = None


def get_ocr_engine() -> RapidOCR:
    """Singleton initializer for offline RapidOCR engine."""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        _OCR_ENGINE = RapidOCR()
    return _OCR_ENGINE


def _parse_grid_table(table: list[list[str | None]]) -> tuple[list[dict[str, Any]], float | None]:
    """
    Parse a structured pdfplumber table (Mode A: native_grid).
    Handles multi-line wrapped cells inside the table and header mapping.
    """
    records: list[dict[str, Any]] = []
    explicit_opening_bal: float | None = None

    if not table or len(table) < 2:
        return records, explicit_opening_bal

    # Locate the table header row
    header_idx = -1
    col_map: dict[str, int] = {}

    for idx, row in enumerate(table):
        cleaned_cells = [re.sub(r"\s+", " ", (c or "")).strip().lower() for c in row]
        row_joined = " ".join(cleaned_cells)
        if ("date" in row_joined) and ("balance" in row_joined) and (
            "narration" in row_joined or "description" in row_joined or "particulars" in row_joined
        ):
            header_idx = idx
            for c_i, cell_txt in enumerate(cleaned_cells):
                if "date" in cell_txt and "value" not in cell_txt and "date" not in col_map:
                    col_map["date"] = c_i
                elif any(k in cell_txt for k in ("narration", "description", "particulars")):
                    col_map["desc"] = c_i
                elif any(k in cell_txt for k in ("ref", "chq", "utr")):
                    col_map["ref"] = c_i
                elif any(k in cell_txt for k in ("debit", "withdrawal", "(dr)")):
                    col_map["debit"] = c_i
                elif any(k in cell_txt for k in ("credit", "deposit", "(cr)")):
                    col_map["credit"] = c_i
                elif "amount" in cell_txt:
                    col_map["amount"] = c_i
                elif "balance" in cell_txt:
                    col_map["balance"] = c_i
            break

    if header_idx == -1 or "date" not in col_map or "balance" not in col_map:
        return records, explicit_opening_bal

    for row in table[header_idx + 1 :]:
        if not row or all(not (c or "").strip() for c in row):
            continue

        raw_date_cell = (row[col_map["date"]] or "").strip() if col_map["date"] < len(row) else ""
        raw_desc_cell = (row[col_map.get("desc", 1)] or "").strip() if col_map.get("desc", 1) < len(row) else ""
        raw_ref_cell = (
            (row[col_map["ref"]] or "").strip()
            if "ref" in col_map and col_map["ref"] < len(row)
            else ""
        )
        raw_bal_cell = (row[col_map["balance"]] or "").strip() if col_map["balance"] < len(row) else ""

        desc_clean = " ".join(line.strip() for line in raw_desc_cell.splitlines() if line.strip())
        ref_clean = " ".join(line.strip() for line in raw_ref_cell.splitlines() if line.strip())
        if ref_clean == "-":
            ref_clean = ""

        # Check for Opening Balance row
        if OPENING_BAL_RE.search(desc_clean):
            bal_val, _ = parse_indian_amount(raw_bal_cell)
            if bal_val is not None:
                explicit_opening_bal = bal_val
            continue

        dt_norm, _ = parse_date_str(raw_date_cell)
        bal_val, _ = parse_indian_amount(raw_bal_cell)

        debit_val = 0.0
        credit_val = 0.0

        if "amount" in col_map and col_map["amount"] < len(row):
            amt_val, drcr = parse_indian_amount(row[col_map["amount"]])
            if amt_val is not None:
                if drcr == "CR":
                    credit_val = amt_val
                elif drcr == "DR":
                    debit_val = amt_val
                else:
                    if infer_direction_from_narration(desc_clean, ref_clean) == "CR":
                        credit_val = amt_val
                    else:
                        debit_val = amt_val
        else:
            if "debit" in col_map and col_map["debit"] < len(row):
                d_val, _ = parse_indian_amount(row[col_map["debit"]])
                if d_val is not None:
                    debit_val = d_val
            if "credit" in col_map and col_map["credit"] < len(row):
                c_val, _ = parse_indian_amount(row[col_map["credit"]])
                if c_val is not None:
                    credit_val = c_val

        # If this is a valid parent transaction row
        if dt_norm and bal_val is not None:
            if not ref_clean:
                ref_clean = extract_ref_from_description(desc_clean)
            records.append({
                "is_parent": True,
                "date": dt_norm,
                "description": desc_clean,
                "ref_no": ref_clean,
                "debit": round(debit_val, 2),
                "credit": round(credit_val, 2),
                "balance": round(bal_val, 2),
                "extraction_method": "native_grid",
            })
        elif desc_clean and not dt_norm and bal_val is None and not is_noise_or_header_line(desc_clean):
            # Wrapped continuation row if grid split across horizontal lines
            records.append({
                "is_parent": False,
                "continuation": desc_clean,
                "extraction_method": "native_grid",
            })

    return records, explicit_opening_bal


def _cluster_items_into_lines(items: list[dict[str, Any]], y_tolerance: float = 4.0) -> list[list[dict[str, Any]]]:
    """
    Cluster words or OCR bounding boxes into horizontal lines by their vertical `top` / `y_center`
    coordinate, then sort items left-to-right by `x0`.
    """
    if not items:
        return []
    sorted_items = sorted(items, key=lambda w: (w["y"], w["x0"]))
    lines: list[list[dict[str, Any]]] = []
    current_line: list[dict[str, Any]] = [sorted_items[0]]
    current_y = sorted_items[0]["y"]

    for item in sorted_items[1:]:
        if abs(item["y"] - current_y) <= y_tolerance:
            current_line.append(item)
            current_y = sum(x["y"] for x in current_line) / len(current_line)
        else:
            lines.append(sorted(current_line, key=lambda w: w["x0"]))
            current_line = [item]
            current_y = item["y"]

    if current_line:
        lines.append(sorted(current_line, key=lambda w: w["x0"]))
    return lines


def _parse_clustered_lines(
    lines: list[list[dict[str, Any]]],
    page_width: float,
    method_name: str,
) -> tuple[list[dict[str, Any]], float | None]:
    """
    Shared coordinate-aware row parser for Mode B (`native_stream`) and Mode C (`offline_ocr`).
    Detects column header X-anchors dynamically and parses transaction lines + wrapped lines.
    """
    records: list[dict[str, Any]] = []
    explicit_opening_bal: float | None = None

    # Default normalized X boundaries (as fraction of page width) if header not yet encountered
    # Standard layout: Date (~0.05) | Narration (~0.15-0.46) | Ref (~0.46-0.58) | ValDate | Debit (~0.70-0.78) | Credit (~0.82-0.88) | Balance (~0.92-0.98)
    dr_x_center = page_width * 0.72
    cr_x_center = page_width * 0.84
    single_amount_mode = False
    in_ledger_section = False

    for line_items in lines:
        line_text = " ".join(w["text"] for w in line_items).strip()
        if not line_text:
            continue

        low_line = line_text.lower()

        # Detect table header row & calibrate Debit / Credit X-coordinates dynamically
        if ("date" in low_line) and ("balance" in low_line) and any(
            k in low_line for k in ("narration", "description", "particulars")
        ):
            in_ledger_section = True
            if "amount" in low_line and "withdrawal" not in low_line and "debit" not in low_line:
                single_amount_mode = True
            for w in line_items:
                wt = w["text"].lower()
                wx = (w["x0"] + w["x1"]) / 2.0
                if wt in ("withdrawal", "debit", "(dr)"):
                    dr_x_center = wx
                elif wt in ("deposit", "credit", "(cr)"):
                    cr_x_center = wx
            continue

        # Skip executive summary tables before the main transaction header
        if "opening balance (inr)" in low_line and "total debits" in low_line:
            continue

        if is_noise_or_header_line(line_text):
            continue

        # Check for Opening Balance row
        if OPENING_BAL_RE.search(line_text):
            in_ledger_section = True
            amt_matches = list(AMOUNT_TOKEN_RE.finditer(line_text))
            if amt_matches:
                bal_val, _ = parse_indian_amount(amt_matches[-1].group(0))
                if bal_val is not None:
                    explicit_opening_bal = bal_val
            continue

        # Check if line starts with a valid transaction date on the left side of the page
        first_word_x = line_items[0]["x0"]
        left_segment = " ".join(w["text"] for w in line_items if w["x0"] < page_width * 0.25)
        dt_norm, _ = parse_date_str(left_segment)

        # Gather all amount tokens along with their horizontal X position
        amt_tokens: list[dict[str, Any]] = []
        for w in line_items:
            if w["x0"] < page_width * 0.52:
                continue
            val, drcr = parse_indian_amount(w["text"])
            if val is not None:
                amt_tokens.append({
                    "val": val,
                    "drcr": drcr,
                    "x_center": (w["x0"] + w["x1"]) / 2.0,
                    "x0": w["x0"],
                    "raw": w["text"],
                })

        # Also check combined adjacent tokens (e.g., "14,500.00" followed by "(Dr)")
        if not amt_tokens:
            right_text = " ".join(w["text"] for w in line_items if w["x0"] >= page_width * 0.52)
            for m in AMOUNT_TOKEN_RE.finditer(right_text):
                val, drcr = parse_indian_amount(m.group(0))
                if val is not None:
                    amt_tokens.append({
                        "val": val,
                        "drcr": drcr,
                        "x_center": page_width * 0.80,
                        "x0": page_width * 0.65,
                        "raw": m.group(0),
                    })

        if dt_norm and len(amt_tokens) >= 2:
            in_ledger_section = True
            # Sort amount tokens left-to-right; rightmost is always Closing Balance
            amt_tokens.sort(key=lambda a: a["x_center"])
            bal_token = amt_tokens[-1]
            tx_amt_tokens = amt_tokens[:-1]

            debit_val = 0.0
            credit_val = 0.0

            if len(tx_amt_tokens) >= 2 and tx_amt_tokens[0]["val"] > 0 and tx_amt_tokens[1]["val"] > 0:
                debit_val = tx_amt_tokens[0]["val"]
                credit_val = tx_amt_tokens[1]["val"]
            else:
                # Pick non-zero transaction amount token (or the only one)
                active_tok = next((t for t in tx_amt_tokens if t["val"] > 0), tx_amt_tokens[0])
                amt_val = active_tok["val"]
                drcr = active_tok["drcr"]

                # Check if explicit (Dr)/(Cr) token appears right after the number
                if not drcr:
                    right_str = " ".join(w["text"] for w in line_items if w["x0"] >= active_tok["x0"])
                    if re.search(r"\(dr\)|\bdr\b", right_str, re.IGNORECASE):
                        drcr = "DR"
                    elif re.search(r"\(cr\)|\bcr\b", right_str, re.IGNORECASE):
                        drcr = "CR"

                if single_amount_mode or drcr in ("DR", "CR"):
                    if drcr == "CR":
                        credit_val = amt_val
                    else:
                        debit_val = amt_val
                else:
                    # Use horizontal X-coordinate proximity to Debit vs Credit column anchors
                    dist_dr = abs(active_tok["x_center"] - dr_x_center)
                    dist_cr = abs(active_tok["x_center"] - cr_x_center)
                    if dist_cr < dist_dr:
                        credit_val = amt_val
                    else:
                        debit_val = amt_val

            # Extract Narration & Reference Number from middle region
            first_amt_x0 = amt_tokens[0]["x0"]
            middle_words = [
                w for w in line_items if page_width * 0.13 <= w["x0"] < (first_amt_x0 - 5)
            ]

            desc_parts: list[str] = []
            ref_val = ""
            for w in middle_words:
                wt = w["text"].strip()
                # Skip secondary Value Date column on borderless statements (~0.52 - 0.65 width)
                if w["x0"] >= page_width * 0.50 and parse_date_str(wt)[0] is not None:
                    continue
                # Identify Reference / UTR column (~0.43 - 0.56 width)
                if page_width * 0.43 <= w["x0"] <= page_width * 0.57 and not ref_val:
                    if re.match(r"^[A-Za-z0-9-]{5,16}$", wt) and wt != "-":
                        ref_val = wt
                        continue
                desc_parts.append(wt)

            desc_clean = " ".join(desc_parts).strip()
            if not ref_val:
                ref_val = extract_ref_from_description(desc_clean)

            records.append({
                "is_parent": True,
                "date": dt_norm,
                "description": desc_clean,
                "ref_no": ref_val,
                "debit": round(debit_val, 2),
                "credit": round(credit_val, 2),
                "balance": round(bal_token["val"], 2),
                "extraction_method": method_name,
            })

        elif in_ledger_section and not dt_norm and len(amt_tokens) == 0:
            # Wrapped multi-line UPI / NEFT narration line (no date, no balance)
            if first_word_x < page_width * 0.55:
                cont_text = " ".join(
                    w["text"].strip() for w in line_items if w["x0"] < page_width * 0.62
                ).strip()
                if cont_text and not is_noise_or_header_line(cont_text):
                    records.append({
                        "is_parent": False,
                        "continuation": cont_text,
                        "extraction_method": method_name,
                    })

    return records, explicit_opening_bal


def _extract_page_native(page: pdfplumber.page.Page) -> tuple[list[dict[str, Any]], float | None]:
    """
    Mode A (`native_grid`) & Mode B (`native_stream`) extraction for pages with embedded text.
    Attempts explicit grid line table extraction first, falling back to coordinate stream clustering.
    """
    # Check if the page has sufficient explicit vector lines for a grid table
    has_grid_lines = len(page.lines) >= 10 or len(page.rects) >= 6
    if has_grid_lines:
        tables = page.extract_tables(
            table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 4,
                "join_tolerance": 4,
            }
        )
        grid_records: list[dict[str, Any]] = []
        opening_bal: float | None = None
        for tbl in tables:
            recs, op_bal = _parse_grid_table(tbl)
            if op_bal is not None and opening_bal is None:
                opening_bal = op_bal
            grid_records.extend(recs)

        if any(r.get("is_parent") for r in grid_records):
            return grid_records, opening_bal

    # Fallback to Mode B (`native_stream`): horizontal word clustering by `top` y-coordinate
    words = page.extract_words(x_tolerance=2.5, y_tolerance=3.0, keep_blank_chars=False)
    items = [
        {
            "text": w["text"],
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "y": (float(w["top"]) + float(w["bottom"])) / 2.0,
        }
        for w in words
    ]
    clustered_lines = _cluster_items_into_lines(items, y_tolerance=4.0)
    return _parse_clustered_lines(clustered_lines, float(page.width), "native_stream")


def _extract_page_ocr(pdf_bytes: bytes, page_index: int) -> tuple[list[dict[str, Any]], float | None]:
    """
    Mode C (`offline_ocr`): Renders a scanned page to a numpy array in RAM at 200 DPI using
    `pypdfium2`, runs `RapidOCR`, groups bounding boxes into horizontal lines by Y-center,
    and parses columns via the coordinate-aware normalizer.
    """
    pdf_doc = pdfium.PdfDocument(pdf_bytes)
    page = pdf_doc[page_index]
    # 200 DPI rendering (200 / 72.0 = 2.7778 scale)
    scale = 200.0 / 72.0
    bitmap = page.render(scale=scale)
    pil_img = bitmap.to_pil().convert("RGB")
    img_width, _ = pil_img.size
    img_np = np.array(pil_img)
    page.close()
    pdf_doc.close()

    ocr_engine = get_ocr_engine()
    ocr_result, _ = ocr_engine(img_np)

    if not ocr_result:
        return [], None

    items: list[dict[str, Any]] = []
    for box, text, _score in ocr_result:
        # box is [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        x0, x1 = float(min(xs)), float(max(xs))
        y_center = float(sum(ys)) / 4.0

        # Split wide OCR text boxes that contain multiple space-separated columns
        # (e.g., Date + Narration or Amount + Balance merged in one OCR box)
        date_m, date_span = parse_date_str(text)
        if date_m and date_span and date_span[0] == 0 and len(text.strip()) > (date_span[1] + 2):
            dt_str = text[: date_span[1]].strip()
            rem_str = text[date_span[1] :].strip()
            dt_ratio = len(dt_str) / max(1, len(text))
            split_x = x0 + (x1 - x0) * dt_ratio
            items.append({"text": dt_str, "x0": x0, "x1": split_x, "y": y_center})
            items.append({"text": rem_str, "x0": split_x + 2.0, "x1": x1, "y": y_center})
        else:
            items.append({"text": text.strip(), "x0": x0, "x1": x1, "y": y_center})

    # Y-tolerance scaled for 200 DPI (~12 pixels)
    clustered_lines = _cluster_items_into_lines(items, y_tolerance=12.0)
    return _parse_clustered_lines(clustered_lines, float(img_width), "offline_ocr")


# ==============================================================================
# 4. STITCHING, SELF-HEALING BALANCE PASS & METADATA ENGINE
# ==============================================================================

def parse_bank_pdf(
    pdf_path: str | bytes,
    password: str | None = None,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Primary shared data contract entry point.
    Parses any native or scanned Indian bank statement PDF into a standardized
    9-column DataFrame and an 11-key financial summary & verification metadata dict.
    """
    decrypted_stream = load_decrypted_pdf_bytes(pdf_path, password=password)
    pdf_raw_bytes = decrypted_stream.getvalue()

    all_raw_records: list[dict[str, Any]] = []
    explicit_opening_balance: float | None = None
    ocr_pages_count = 0

    with pdfplumber.open(io.BytesIO(pdf_raw_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            if len(page.chars) > 20:
                page_records, op_bal = _extract_page_native(page)
            else:
                ocr_pages_count += 1
                page_records, op_bal = _extract_page_ocr(pdf_raw_bytes, page_idx)

            if op_bal is not None and explicit_opening_balance is None:
                explicit_opening_balance = op_bal
            all_raw_records.extend(page_records)

    # Multi-line Stitching Pass: Attach wrapped lines (no date/balance) onto parent row
    stitched_rows: list[dict[str, Any]] = []
    for rec in all_raw_records:
        if rec.get("is_parent"):
            stitched_rows.append({
                "date": rec["date"],
                "description": rec["description"],
                "ref_no": rec["ref_no"],
                "debit": round(float(rec["debit"]), 2),
                "credit": round(float(rec["credit"]), 2),
                "balance": round(float(rec["balance"]), 2),
                "extraction_method": rec["extraction_method"],
            })
        else:
            if stitched_rows:
                cont = rec.get("continuation", "").strip()
                if cont:
                    stitched_rows[-1]["description"] = (
                        f"{stitched_rows[-1]['description']} {cont}".strip()
                    )
                    if not stitched_rows[-1]["ref_no"]:
                        stitched_rows[-1]["ref_no"] = extract_ref_from_description(
                            stitched_rows[-1]["description"]
                        )

    # Handle empty statement edge case cleanly
    columns = [
        "tx_id",
        "date",
        "description",
        "ref_no",
        "debit",
        "credit",
        "balance",
        "math_verified",
        "extraction_method",
    ]
    if not stitched_rows:
        empty_df = pd.DataFrame(columns=columns)
        empty_meta = {
            "total_rows": 0,
            "opening_balance": 0.0,
            "closing_balance": 0.0,
            "total_debit": 0.0,
            "total_credit": 0.0,
            "net_cash_flow": 0.0,
            "daily_burn_rate": 0.0,
            "estimated_runway_days": 999.0,
            "verification_rate_pct": 0.0,
            "unverified_row_ids": [],
            "ocr_pages_count": ocr_pages_count,
        }
        return empty_df, empty_meta

    # Self-Healing Running-Balance Pass
    # Step 1: Resolve Row 1 & Opening Balance
    first = stitched_rows[0]
    if explicit_opening_balance is not None:
        # Check if Row 1 satisfies explicit_opening_balance; if not, test debit/credit swap
        expected_r1 = round(explicit_opening_balance + first["credit"] - first["debit"], 2)
        if expected_r1 != round(first["balance"], 2):
            swapped_r1 = round(explicit_opening_balance + first["debit"] - first["credit"], 2)
            if swapped_r1 == round(first["balance"], 2):
                first["debit"], first["credit"] = first["credit"], first["debit"]

    # Always enforce data contract formula: opening_balance = round(first_row.balance - first_row.credit + first_row.debit, 2)
    opening_balance = round(first["balance"] - first["credit"] + first["debit"], 2)

    prev_balance = opening_balance
    unverified_row_ids: list[int] = []
    verified_count = 0

    for idx, row in enumerate(stitched_rows, start=1):
        row["tx_id"] = idx
        row["description"] = re.sub(r"\s+", " ", row["description"]).strip()
        row["ref_no"] = str(row["ref_no"] or "").strip()

        expected_bal = round(prev_balance + row["credit"] - row["debit"], 2)
        actual_bal = round(row["balance"], 2)

        if expected_bal == actual_bal:
            row["math_verified"] = True
        else:
            # Test if swapping debit and credit satisfies the running balance equation
            swapped_bal = round(prev_balance + row["debit"] - row["credit"], 2)
            if swapped_bal == actual_bal:
                row["debit"], row["credit"] = row["credit"], row["debit"]
                row["math_verified"] = True
            else:
                row["math_verified"] = False
                unverified_row_ids.append(idx)

        if row["math_verified"]:
            verified_count += 1

        prev_balance = actual_bal

    df = pd.DataFrame(stitched_rows)[columns]
    df = df.astype({
        "tx_id": "int64",
        "date": "object",
        "description": "object",
        "ref_no": "object",
        "debit": "float64",
        "credit": "float64",
        "balance": "float64",
        "math_verified": "bool",
        "extraction_method": "object",
    })

    total_rows = len(df)
    closing_balance = round(float(df["balance"].iloc[-1]), 2)
    total_debit = round(float(df["debit"].sum()), 2)
    total_credit = round(float(df["credit"].sum()), 2)
    net_cash_flow = round(total_credit - total_debit, 2)

    # Calculate Date Span & Daily Burn Rate
    parsed_dates = pd.to_datetime(df["date"], format="%Y-%m-%d", errors="coerce").dropna()
    if len(parsed_dates) >= 2:
        span_days = max(1, (parsed_dates.max() - parsed_dates.min()).days + 1)
    else:
        span_days = 1

    daily_burn_rate = round(total_debit / span_days, 2)
    estimated_runway_days = (
        round(closing_balance / daily_burn_rate, 1) if daily_burn_rate > 0 else 999.0
    )
    verification_rate_pct = round((verified_count / total_rows) * 100.0, 2)

    metadata: dict[str, Any] = {
        "total_rows": int(total_rows),
        "opening_balance": float(opening_balance),
        "closing_balance": float(closing_balance),
        "total_debit": float(total_debit),
        "total_credit": float(total_credit),
        "net_cash_flow": float(net_cash_flow),
        "daily_burn_rate": float(daily_burn_rate),
        "estimated_runway_days": float(estimated_runway_days),
        "verification_rate_pct": float(verification_rate_pct),
        "unverified_row_ids": unverified_row_ids,
        "ocr_pages_count": int(ocr_pages_count),
    }

    return df, metadata


# ==============================================================================
# 5. CLI & N8N EXECUTE COMMAND INTERFACE
# ==============================================================================

def _run_test_suite():
    """Default verification suite across all 6 synthetic PDFs in ./test_statements/."""
    test_dir = os.path.join(".", "test_statements")
    test_files = [
        ("bank_1_borderless.pdf", None),
        ("bank_2_multiline_heavy.pdf", None),
        ("bank_3_single_amount_drcr.pdf", None),
        ("bank_4_messy_headers.pdf", None),
        ("bank_5_password_locked.pdf", "VAULT123"),
        ("bank_6_scanned_image.pdf", None),
    ]

    summary_records = []
    for fname, pw in test_files:
        fpath = os.path.join(test_dir, fname)
        if not os.path.exists(fpath):
            print(f"[WARN] File not found: {fpath}. Run generate_test_pdfs.py first.", file=sys.stderr)
            continue

        df, meta = parse_bank_pdf(fpath, password=pw)
        methods = ",".join(sorted(df["extraction_method"].unique())) if not df.empty else "none"
        summary_records.append({
            "Statement PDF": fname,
            "Method": methods,
            "Rows": meta["total_rows"],
            "Opening Bal": f"{meta['opening_balance']:,.2f}",
            "Total Debit": f"{meta['total_debit']:,.2f}",
            "Total Credit": f"{meta['total_credit']:,.2f}",
            "Closing Bal": f"{meta['closing_balance']:,.2f}",
            "Verified %": f"{meta['verification_rate_pct']:.1f}%",
            "OCR Pages": meta["ocr_pages_count"],
            "Runway (Days)": meta["estimated_runway_days"],
        })

    if summary_records:
        summary_df = pd.DataFrame(summary_records)
        print("\n" + "=" * 120)
        print("PARSER ENGINE VERIFICATION SUMMARY (ALL 3 EXTRACTION MODES)")
        print("=" * 120)
        print(summary_df.to_string(index=False))
        print("=" * 120 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Offline Indian Bank Statement Parser Engine (n8n CLI & Python Module)"
    )
    parser.add_argument("--pdf", type=str, help="Path to the bank statement PDF file")
    parser.add_argument("--password", type=str, default=None, help="Optional PDF password")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output standardized JSON {metadata, rows} to stdout for n8n integration",
    )
    args = parser.parse_args()

    if args.pdf:
        df, metadata = parse_bank_pdf(args.pdf, password=args.password)
        if args.json:
            payload = {
                "metadata": metadata,
                "rows": df.to_dict(orient="records"),
            }
            print(json.dumps(payload, indent=2))
        else:
            print("=== METADATA ===")
            print(json.dumps(metadata, indent=2))
            print("\n=== EXTRACTED TRANSACTIONS ===")
            print(df.to_string(index=False))
    else:
        _run_test_suite()


if __name__ == "__main__":
    main()