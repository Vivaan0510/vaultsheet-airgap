#!/usr/bin/env python3
"""
generate_test_pdfs.py
---------------------
Standalone synthetic test suite generator for Indian bank statement parsing,
OCR fallback, password decryption, mathematical balance verification, and B2B
vendor reconciliation engines.

Dependencies:
    pip install reportlab pypdf pypdfium2 Pillow pandas
"""

import io
import os
import random
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import pypdfium2 as pdfium
from PIL import Image, ImageFilter
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


# ==============================================================================
# 1. INDIAN NUMBER FORMATTING & DECIMAL ARITHMETIC ENGINE
# ==============================================================================

def to_dec(val) -> Decimal:
    """Convert float/int/str to exact 2-decimal-place Decimal."""
    return Decimal(str(val)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def format_inr(val) -> str:
    """Format a number into the Indian numbering system (e.g., 12,05,450.75)."""
    if val is None or val == "":
        return ""
    d = to_dec(val)
    sign = "-" if d < 0 else ""
    d = abs(d)
    s = f"{d:.2f}"
    int_part, dec_part = s.split(".")
    if len(int_part) <= 3:
        res = int_part
    else:
        last3 = int_part[-3:]
        rest = int_part[:-3]
        groups = []
        while rest:
            groups.append(rest[-2:])
            rest = rest[:-2]
        groups.reverse()
        res = ",".join(groups) + "," + last3
    return f"{sign}{res}.{dec_part}"


# ==============================================================================
# 2. SYNTHETIC TRANSACTION LEDGER GENERATOR
# ==============================================================================

class StatementDataGenerator:
    """Generates mathematically exact Indian bank transaction ledgers."""

    UPI_MERCHANTS = [
        ("SWIGGY INSTAMART", "HDFC0000123", "Order 99812"),
        ("ZOMATO LIMITED", "ICIC0000045", "Food Order 77410"),
        ("BLINKIT GROCERY", "UTIB0000891", "Instant Delivery 3321"),
        ("AMAZON PAY INDIA", "SBIN0004512", "Marketplace Ref 8812"),
        ("UBER INDIA SYST", "KKBK0000912", "Trip Ride 10942"),
        ("RELIANCE RETAIL", "HDFC0000551", "JioMart Bill 6631"),
        ("AIRTEL PAYMENTS", "AIRP0000001", "Broadband Fiber 4491"),
        ("TATA STARBUCKS", "ICIC0000229", "POS Beverage 2219"),
        ("PHARMEASY HEALTH", "YESB0000118", "Med Order 55210"),
        ("BIGBASKET DAILY", "HDFC0000991", "Wallet Recharge 712"),
    ]

    HIDDEN_FEES = [
        ("SMS ALERT CHG Q3", Decimal("17.70")),
        ("MIN BAL PENALTY", Decimal("590.00")),
        ("IMPS CONVENIENCE FEE", Decimal("5.90")),
        ("DEBIT CARD ANNUAL FEE", Decimal("590.00")),
    ]

    CREDITS = [
        ("NEFT/CR/N289100441/RAZORPAY SOFTWARE PVT LTD/SETTLEMENT", Decimal("145200.00")),
        ("NEFT/CR/N294188210/STRIPE INDIA PRIVATE LTD/PAYOUT", Decimal("88450.50")),
        ("IMPS/CR/428519902112/CLIENT RETENTION MILESTONE/ICIC", Decimal("65000.00")),
        ("ACH/CR/TECHFLOW CLIENT INVOICE 2026-091/HDFC", Decimal("210000.00")),
        ("INT.PD:01-07-2026 TO 30-09-2026 CREDIT INTEREST", Decimal("4820.00")),
    ]

    @classmethod
    def build_ledger(
        cls,
        seed: int,
        num_txns: int,
        opening_balance: str = "1250000.00",
        multiline_ratio: float = 0.35,
        include_duplicate: bool = False,
    ):
        rng = random.Random(seed)
        bal = to_dec(opening_balance)
        curr_date = datetime(2026, 9, 1)
        txns = []

        # Mandatory special auditing & reconciliation test transactions:
        # 1. Exact B2B Contract matches, 1 Overcharge, 1 GST Calculation Mismatch
        b2b_specials = [
            {
                "lines": [
                    "NEFT/DR/N261998210/APEX LOGISTICS PVT LTD",
                    "SBIN0004129/INV-2026-881 FREIGHT CHARGES",
                ],
                "ref": "N261998210",
                "debit": to_dec("59000.00"),  # Contract: 50,000 + 18% GST = 59,000.00 (EXACT)
                "credit": Decimal("0.00"),
                "category": "B2B_EXACT",
            },
            {
                "lines": [
                    "NEFT/DR/N262884129/CLOUDHOST INDIA SERVICES",
                    "HDFC0001209/BILL-OCT-992 DEDICATED SERVER",
                    "UNAUTHORIZED BANDWIDTH OVERAGE SURCHARGE",
                ],
                "ref": "N262884129",
                "debit": to_dec("112100.00"),  # Contract expected: 94,400.00 -> OVERCHARGE (+17,700)
                "credit": Decimal("0.00"),
                "category": "B2B_OVERCHARGE",
            },
            {
                "lines": [
                    "NEFT/DR/N263771029/SHARMA PACKAGING INDUSTRIES",
                    "ICIC0000812/PO-4410 CORRUGATED BOX SUPPLY",
                    "TAX BILLED AT 18 PCT INSTEAD OF CONTRACT 12 PCT",
                ],
                "ref": "N263771029",
                "debit": to_dec("41300.00"),  # Contract: 35,000 + 12% GST = 39,200.00 -> GST MISMATCH (18% billed)
                "credit": Decimal("0.00"),
                "category": "B2B_GST_MISMATCH",
            },
            {
                "lines": [
                    "RTGS/DR/R264991201/METRO POWER INFRASTRUCTURE",
                    "UTIB0000104/MTHLY DG SET AMNT CONTRACT",
                ],
                "ref": "R264991201",
                "debit": to_dec("141600.00"),  # Contract: 1,20,000 + 18% GST = 1,41,600.00 (EXACT)
                "credit": Decimal("0.00"),
                "category": "B2B_EXACT",
            },
        ]

        # Mandatory hidden bank fee rows
        fee_specials = [
            {
                "lines": [fee_name],
                "ref": f"FEE{rng.randint(100000, 999999)}",
                "debit": fee_amt,
                "credit": Decimal("0.00"),
                "category": "HIDDEN_FEE",
            }
            for fee_name, fee_amt in cls.HIDDEN_FEES
        ]

        # Combine specials and distribute across the timeline
        special_pool = b2b_specials + fee_specials
        special_indices = sorted(
            rng.sample(range(1, max(len(special_pool) + 2, num_txns - 2)), len(special_pool))
        )
        special_map = dict(zip(special_indices, special_pool))

        dup_inserted = False

        for idx in range(num_txns):
            if idx > 0 and rng.random() < 0.65:
                curr_date += timedelta(days=rng.randint(0, 2))
            if curr_date > datetime(2026, 10, 14):
                curr_date = datetime(2026, 10, 14)

            if idx in special_map:
                spec = special_map[idx]
                lines = list(spec["lines"])
                # Pad to 3 lines if multiline_ratio demands heavy wrapping
                if rng.random() < multiline_ratio and len(lines) < 3:
                    lines.append("BRANCH AUTO-SETTLEMENT BATCH MUMBAI MAIN")
                debit = spec["debit"]
                credit = spec["credit"]
                ref = spec["ref"]
                cat = spec["category"]
            elif include_duplicate and not dup_inserted and idx == num_txns // 2:
                # Deliberate duplicate debit on the exact same day
                dup_inserted = True
                utr = "428719928311"
                lines = [
                    f"UPI/DR/{utr}/SWIGGY INSTAMART",
                    "HDFC0000123/Order 99812/Instant",
                ]
                debit = to_dec("2450.00")
                credit = Decimal("0.00")
                ref = utr
                cat = "DUPLICATE_ORIGINAL"

                bal = bal + credit - debit
                txns.append({
                    "date": curr_date,
                    "val_date": curr_date,
                    "lines": lines,
                    "narration": " ".join(lines),
                    "ref": ref,
                    "debit": debit,
                    "credit": credit,
                    "balance": bal,
                    "category": cat,
                })

                # Immediate identical duplicate debit on the same date
                bal = bal + credit - debit
                txns.append({
                    "date": curr_date,
                    "val_date": curr_date,
                    "lines": list(lines),
                    "narration": " ".join(lines),
                    "ref": ref,
                    "debit": debit,
                    "credit": credit,
                    "balance": bal,
                    "category": "DUPLICATE_DEBIT",
                })
                continue
            else:
                # Generate realistic UPI or Corporate Credit
                is_credit = rng.random() < 0.24
                is_multiline = rng.random() < multiline_ratio

                if is_credit:
                    raw_desc, base_amt = rng.choice(cls.CREDITS)
                    jitter = to_dec(rng.randint(-2500, 5000))
                    credit = max(to_dec("5000.00"), base_amt + jitter)
                    debit = Decimal("0.00")
                    ref = f"CR{rng.randint(100000000, 999999999)}"
                    if is_multiline:
                        parts = raw_desc.split("/")
                        lines = [
                            "/".join(parts[:3]),
                            "/".join(parts[3:]) if len(parts) > 3 else "CORP PAYOUT",
                            f"REMARKS: INV SETTLEMENT BATCH {rng.randint(100, 999)}",
                        ]
                    else:
                        lines = [raw_desc]
                    cat = "CREDIT"
                else:
                    merch, ifsc, order_tag = rng.choice(cls.UPI_MERCHANTS)
                    utr = f"4129{rng.randint(100000, 999999)}"
                    dir_code = "DR"
                    debit = to_dec(f"{rng.randint(180, 14500)}.{rng.choice(['00', '25', '50', '75'])}")
                    credit = Decimal("0.00")
                    ref = utr
                    if is_multiline:
                        lines = [
                            f"UPI/{dir_code}/{utr}/{merch}/",
                            f"{ifsc}/{order_tag}/MUMBAI",
                            f"AUTH-REF-{rng.randint(10000, 99999)} VERIFIED",
                        ]
                    else:
                        lines = [
                            f"UPI/{dir_code}/{utr}/{merch}/{ifsc}/",
                            f"{order_tag}",
                        ]
                    cat = "UPI_DEBIT"

            bal = bal + credit - debit
            txns.append({
                "date": curr_date,
                "val_date": curr_date,
                "lines": lines,
                "narration": " ".join(lines),
                "ref": ref,
                "debit": debit,
                "credit": credit,
                "balance": bal,
                "category": cat,
            })

        return {
            "opening_balance": to_dec(opening_balance),
            "closing_balance": bal,
            "transactions": txns,
        }


# ==============================================================================
# 3. PDF GENERATORS FOR ALL 6 VARIATIONS
# ==============================================================================

def create_vendor_master_csv(output_dir: str) -> str:
    """Creates vendor_master.csv with 4 Indian B2B vendors."""
    csv_path = os.path.join(output_dir, "vendor_master.csv")
    rows = [
        {
            "vendor_keyword": "APEX LOGISTICS",
            "contract_amount": 50000.00,
            "gst_rate_pct": 18,
            "expected_total": 59000.00,
        },
        {
            "vendor_keyword": "CLOUDHOST INDIA",
            "contract_amount": 80000.00,
            "gst_rate_pct": 18,
            "expected_total": 94400.00,
        },
        {
            "vendor_keyword": "SHARMA PACKAGING",
            "contract_amount": 35000.00,
            "gst_rate_pct": 12,
            "expected_total": 39200.00,
        },
        {
            "vendor_keyword": "METRO POWER",
            "contract_amount": 120000.00,
            "gst_rate_pct": 18,
            "expected_total": 141600.00,
        },
    ]
    df = pd.DataFrame(rows)
    df.to_csv(csv_path, index=False)
    return csv_path


def generate_bank_1_borderless(output_dir: str) -> dict:
    """
    bank_1_borderless.pdf:
    - Exact 3 pages
    - Borderless table (whitespace alignment only, zero cell grid lines)
    - Repeating bank header/footer on every page
    - Columns: Date | Narration | Chq/Ref No | Value Date | Withdrawal (Dr) | Deposit (Cr) | Closing Balance
    """
    path = os.path.join(output_dir, "bank_1_borderless.pdf")
    ledger = StatementDataGenerator.build_ledger(
        seed=101, num_txns=48, opening_balance="1425000.00", multiline_ratio=0.45
    )

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    # Split 48 transactions across exactly 3 pages (16 per page)
    txns = ledger["transactions"]
    pages_data = [txns[0:16], txns[16:32], txns[32:48]]

    # Column X coordinates (whitespace aligned)
    X_DATE = 28
    X_NARR = 88
    X_REF = 265
    X_VAL = 335
    X_DR_R = 442
    X_CR_R = 508
    X_BAL_R = 572

    for page_idx, page_txns in enumerate(pages_data, start=1):
        # Repeating Bank Header
        c.setFont("Helvetica-Bold", 13)
        c.drawString(28, height - 38, "AXIS COMMERCIAL BANK LTD - CORPORATE CURRENT ACCOUNT")
        c.setFont("Helvetica", 8.5)
        c.drawString(28, height - 52, "Branch: Nariman Point, Mumbai - 400021 | IFSC: UTIB0000004 | GSTIN: 27AAACU1234F1Z5")
        c.drawString(28, height - 64, "Customer: NEXTGEN FINTECH LABS PVT LTD | A/C No: 92602004419821 | Currency: INR")
        c.drawRightString(572, height - 64, f"Statement Period: 01-Sep-2026 to 14-Oct-2026")

        # Column Headers (Whitespace only - no grid borders)
        y = height - 88
        c.setFont("Helvetica-Bold", 8)
        c.drawString(X_DATE, y, "Date")
        c.drawString(X_NARR, y, "Narration")
        c.drawString(X_REF, y, "Chq/Ref No")
        c.drawString(X_VAL, y, "Value Date")
        c.drawRightString(X_DR_R, y, "Withdrawal (Dr)")
        c.drawRightString(X_CR_R, y, "Deposit (Cr)")
        c.drawRightString(X_BAL_R, y, "Closing Balance")

        y -= 18
        c.setFont("Helvetica", 7.8)

        # Opening Balance Row on Page 1
        if page_idx == 1:
            c.drawString(X_DATE, y, "01/09/2026")
            c.drawString(X_NARR, y, "B/F OPENING BALANCE")
            c.drawString(X_REF, y, "-")
            c.drawString(X_VAL, y, "01/09/2026")
            c.drawRightString(X_DR_R, y, "")
            c.drawRightString(X_CR_R, y, "")
            c.drawRightString(X_BAL_R, y, format_inr(ledger["opening_balance"]))
            y -= 16

        for t in page_txns:
            d_str = t["date"].strftime("%d/%m/%Y")
            v_str = t["val_date"].strftime("%d/%m/%Y")
            dr_str = format_inr(t["debit"]) if t["debit"] > 0 else ""
            cr_str = format_inr(t["credit"]) if t["credit"] > 0 else ""
            bal_str = format_inr(t["balance"])

            # First line of transaction (has Date, Ref, Value Date, Amounts)
            c.drawString(X_DATE, y, d_str)
            c.drawString(X_NARR, y, t["lines"][0][:34])
            c.drawString(X_REF, y, t["ref"])
            c.drawString(X_VAL, y, v_str)
            c.drawRightString(X_DR_R, y, dr_str)
            c.drawRightString(X_CR_R, y, cr_str)
            c.drawRightString(X_BAL_R, y, bal_str)
            y -= 11

            # Wrapped multi-line UPI / NEFT strings (no date on wrapped lines)
            for extra_line in t["lines"][1:]:
                c.drawString(X_NARR, y, extra_line[:36])
                y -= 10

            y -= 4  # subtle row spacing

        # Repeating Footer on every page
        c.setFont("Helvetica-Oblique", 7.5)
        c.drawString(28, 36, "This is a system-generated borderless statement and does not require a physical signature.")
        c.drawRightString(572, 36, f"Page {page_idx} of 3")
        c.showPage()

    c.save()
    return {"file": "bank_1_borderless.pdf", "pages": 3, "ledger": ledger}


def generate_bank_2_multiline_heavy(output_dir: str) -> dict:
    """
    bank_2_multiline_heavy.pdf:
    - Exact 3 pages
    - Grid-lined table where >=60% of transactions wrap across 3 vertical lines inside Description cell
    - Informational footer blocks on every page
    """
    path = os.path.join(output_dir, "bank_2_multiline_heavy.pdf")
    ledger = StatementDataGenerator.build_ledger(
        seed=202, num_txns=36, opening_balance="980500.00", multiline_ratio=0.75
    )

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    txns = ledger["transactions"]
    pages_data = [txns[0:12], txns[12:24], txns[24:36]]

    # Grid X positions
    COLS = [28, 86, 280, 350, 425, 498, 570]

    for page_idx, page_txns in enumerate(pages_data, start=1):
        # Header
        c.setFont("Helvetica-Bold", 13)
        c.drawString(28, height - 36, "HDFC CORPORATE BANKING - DETAILED ACCOUNT LEDGER")
        c.setFont("Helvetica", 8.5)
        c.drawString(28, height - 50, "Account Holder: NEXTGEN FINTECH LABS PVT LTD | A/C: 50200088192011 | IFSC: HDFC0000123")

        top_y = height - 68
        # Draw header row box
        c.setFillColor(colors.HexColor("#EAECEE"))
        c.rect(COLS[0], top_y - 18, COLS[-1] - COLS[0], 18, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 8)

        headers = ["Txn Date", "Description / Narration", "Ref / UTR", "Debit (INR)", "Credit (INR)", "Balance (INR)"]
        for idx_h, h_text in enumerate(headers):
            if idx_h >= 3:
                c.drawRightString(COLS[idx_h + 1] - 4, top_y - 12, h_text)
            else:
                c.drawString(COLS[idx_h] + 4, top_y - 12, h_text)

        for x in COLS:
            c.line(x, top_y, x, top_y - 18)

        y = top_y - 18
        c.setFont("Helvetica", 7.8)

        if page_idx == 1:
            row_h = 18
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)
            c.drawString(COLS[0] + 4, y - 12, "01/09/2026")
            c.drawString(COLS[1] + 4, y - 12, "OPENING BALANCE BROUGHT FORWARD")
            c.drawString(COLS[2] + 4, y - 12, "-")
            c.drawRightString(COLS[6] - 4, y - 12, format_inr(ledger["opening_balance"]))
            y -= row_h

        for t in page_txns:
            num_lines = len(t["lines"])
            row_h = max(18, num_lines * 11 + 6)

            # Draw full grid cell borders
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)

            c.drawString(COLS[0] + 4, y - 12, t["date"].strftime("%d/%m/%Y"))
            c.drawString(COLS[2] + 4, y - 12, t["ref"])

            if t["debit"] > 0:
                c.drawRightString(COLS[4] - 4, y - 12, format_inr(t["debit"]))
            if t["credit"] > 0:
                c.drawRightString(COLS[5] - 4, y - 12, format_inr(t["credit"]))
            c.drawRightString(COLS[6] - 4, y - 12, format_inr(t["balance"]))

            # Multi-line wrapped description inside cell
            text_y = y - 12
            for line_str in t["lines"]:
                c.drawString(COLS[1] + 4, text_y, line_str[:38])
                text_y -= 10.5

            y -= row_h

        # Informational Footer Block
        footer_y = 95
        c.setStrokeColor(colors.HexColor("#888888"))
        c.rect(28, 30, 542, footer_y - 25, fill=0, stroke=1)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(34, footer_y - 38, "IMPORTANT REGULATORY & STATUTORY INFORMATION:")
        c.setFont("Helvetica", 7)
        c.drawString(34, footer_y - 49, "1. All discrepancies in this statement must be reported within 15 days of receipt to corporate.care@hdfcbank.example.in.")
        c.drawString(34, footer_y - 59, "2. DICGC insures principal and accrued interest up to INR 5,00,000.00 per depositor across all branches.")
        c.drawString(34, footer_y - 69, "3. Abbreviations: UPI = Unified Payments Interface | NEFT = National Electronic Funds Transfer | CHG = Service Charge.")
        c.drawRightString(564, footer_y - 69, f"Page {page_idx} of 3")
        c.showPage()

    c.save()
    return {"file": "bank_2_multiline_heavy.pdf", "pages": 3, "ledger": ledger}


def generate_bank_3_single_amount_drcr(output_dir: str) -> dict:
    """
    bank_3_single_amount_drcr.pdf:
    - Exact 2 pages
    - Single `Amount (INR)` column with `Dr/Cr` suffixes (e.g., `14,500.00 (Dr)` or `85,000.00 (Cr)`)
    - Dates in `14-Oct-2026` (%d-%b-%Y) format
    """
    path = os.path.join(output_dir, "bank_3_single_amount_drcr.pdf")
    ledger = StatementDataGenerator.build_ledger(
        seed=303, num_txns=28, opening_balance="850000.00", multiline_ratio=0.40
    )

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    txns = ledger["transactions"]
    pages_data = [txns[0:14], txns[14:28]]
    COLS = [28, 96, 315, 385, 480, 570]

    for page_idx, page_txns in enumerate(pages_data, start=1):
        c.setFont("Helvetica-Bold", 13)
        c.drawString(28, height - 38, "ICICI BUSINESS BANKING - STATEMENT OF ACCOUNT")
        c.setFont("Helvetica", 8.5)
        c.drawString(28, height - 52, "Account Name: NEXTGEN FINTECH LABS PVT LTD | Account No: 000405118921 | Branch: BKC Mumbai")

        top_y = height - 70
        c.setFillColor(colors.HexColor("#F2F4F4"))
        c.rect(COLS[0], top_y - 18, COLS[-1] - COLS[0], 18, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 8)

        c.drawString(COLS[0] + 4, top_y - 12, "Date")
        c.drawString(COLS[1] + 4, top_y - 12, "Transaction Particulars")
        c.drawString(COLS[2] + 4, top_y - 12, "Ref No")
        c.drawRightString(COLS[4] - 4, top_y - 12, "Amount (INR)")
        c.drawRightString(COLS[5] - 4, top_y - 12, "Balance (INR)")

        for x in COLS:
            c.line(x, top_y, x, top_y - 18)

        y = top_y - 18
        c.setFont("Helvetica", 7.8)

        if page_idx == 1:
            row_h = 18
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)
            c.drawString(COLS[0] + 4, y - 12, "01-Sep-2026")
            c.drawString(COLS[1] + 4, y - 12, "OPENING BALANCE")
            c.drawString(COLS[2] + 4, y - 12, "-")
            c.drawRightString(COLS[4] - 4, y - 12, "0.00 (Cr)")
            c.drawRightString(COLS[5] - 4, y - 12, f"{format_inr(ledger['opening_balance'])} (Cr)")
            y -= row_h

        for t in page_txns:
            row_h = max(18, len(t["lines"]) * 11 + 6)
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)

            c.drawString(COLS[0] + 4, y - 12, t["date"].strftime("%d-%b-%Y"))
            c.drawString(COLS[2] + 4, y - 12, t["ref"])

            if t["debit"] > 0:
                amt_str = f"{format_inr(t['debit'])} (Dr)"
            else:
                amt_str = f"{format_inr(t['credit'])} (Cr)"

            c.drawRightString(COLS[4] - 4, y - 12, amt_str)
            c.drawRightString(COLS[5] - 4, y - 12, f"{format_inr(t['balance'])} (Cr)")

            text_y = y - 12
            for line_str in t["lines"]:
                c.drawString(COLS[1] + 4, text_y, line_str[:42])
                text_y -= 10.5

            y -= row_h

        c.setFont("Helvetica", 7.5)
        c.drawString(28, 35, "Legend: (Dr) = Debit / Withdrawal | (Cr) = Credit / Deposit")
        c.drawRightString(570, 35, f"Page {page_idx} of 2")
        c.showPage()

    c.save()
    return {"file": "bank_3_single_amount_drcr.pdf", "pages": 2, "ledger": ledger}


def generate_bank_4_messy_headers(output_dir: str) -> dict:
    """
    bank_4_messy_headers.pdf:
    - Exact 2 pages
    - Account summary block at the top of Page 1 before the transaction ledger starts mid-page
    - Dates in `14/10/2026` (%d/%m/%Y) format
    - Deliberate duplicate debit transaction on the same day
    """
    path = os.path.join(output_dir, "bank_4_messy_headers.pdf")
    ledger = StatementDataGenerator.build_ledger(
        seed=404,
        num_txns=24,
        opening_balance="1640000.00",
        multiline_ratio=0.35,
        include_duplicate=True,
    )

    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    txns = ledger["transactions"]
    # Fewer rows on Page 1 because the messy Account Summary block takes up top half of Page 1
    pages_data = [txns[0:9], txns[9:]]
    COLS = [28, 86, 285, 355, 428, 500, 570]

    total_dr = sum(t["debit"] for t in txns)
    total_cr = sum(t["credit"] for t in txns)

    for page_idx, page_txns in enumerate(pages_data, start=1):
        c.setFont("Helvetica-Bold", 13)
        c.drawString(28, height - 36, "STATE BANK OF INDIA - CORPORATE FINANCIAL SUMMARY & LEDGER")
        c.setFont("Helvetica", 8.5)
        c.drawString(28, height - 50, "Branch: Fort Main Branch (00300) | IFSC: SBIN0000300 | CIF: 8891204412")

        if page_idx == 1:
            # Messy Account Summary & Mini Table at top of Page 1 before ledger starts mid-page
            c.setFillColor(colors.HexColor("#F8F9F9"))
            c.rect(28, height - 210, 542, 145, fill=1, stroke=1)
            c.setFillColor(colors.black)

            c.setFont("Helvetica-Bold", 9.5)
            c.drawString(38, height - 80, "EXECUTIVE ACCOUNT RELATIONSHIP SUMMARY (AS OF 14/10/2026)")
            c.setFont("Helvetica", 8)
            c.drawString(38, height - 96, "Customer Name: NEXTGEN FINTECH LABS PVT LTD")
            c.drawString(38, height - 110, "Registered Address: Unit 402, Peninsula Corporate Park, Lower Parel, Mumbai 400013")
            c.drawString(38, height - 124, "Sanctioned OD Limit: INR 0.00 | Drawing Power: INR 0.00 | Nominee Registered: NO")
            c.drawString(38, height - 138, "Uncleared Cheques: INR 0.00 | Lien Amount: INR 0.00 | Hold Balance: INR 0.00")

            # Embedded summary table that can confuse naive table parsers
            sum_y = height - 156
            c.setFont("Helvetica-Bold", 8)
            c.drawString(38, sum_y, "Opening Balance (INR)")
            c.drawString(175, sum_y, "Total Debits (INR)")
            c.drawString(310, sum_y, "Total Credits (INR)")
            c.drawString(445, sum_y, "Closing Balance (INR)")
            c.line(38, sum_y - 4, 555, sum_y - 4)

            c.setFont("Helvetica", 8.5)
            c.drawString(38, sum_y - 18, format_inr(ledger["opening_balance"]))
            c.drawString(175, sum_y - 18, format_inr(total_dr))
            c.drawString(310, sum_y - 18, format_inr(total_cr))
            c.drawString(445, sum_y - 18, format_inr(ledger["closing_balance"]))

            c.setFont("Helvetica-Oblique", 7.5)
            c.drawString(38, height - 202, "Note: Scroll down to mid-page for itemized chronological transaction ledger.")

            top_y = height - 235
        else:
            top_y = height - 70

        # Transaction Table Header
        c.setFillColor(colors.HexColor("#D6EAF8"))
        c.rect(COLS[0], top_y - 18, COLS[-1] - COLS[0], 18, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 8)

        headers = ["Date", "Narration", "Ref / Chq", "Debit (Dr)", "Credit (Cr)", "Balance"]
        for idx_h, h_text in enumerate(headers):
            if idx_h >= 3:
                c.drawRightString(COLS[idx_h + 1] - 4, top_y - 12, h_text)
            else:
                c.drawString(COLS[idx_h] + 4, top_y - 12, h_text)

        for x in COLS:
            c.line(x, top_y, x, top_y - 18)

        y = top_y - 18
        c.setFont("Helvetica", 7.8)

        if page_idx == 1:
            row_h = 18
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)
            c.drawString(COLS[0] + 4, y - 12, "01/09/2026")
            c.drawString(COLS[1] + 4, y - 12, "B/F OPENING BALANCE")
            c.drawString(COLS[2] + 4, y - 12, "-")
            c.drawRightString(COLS[6] - 4, y - 12, format_inr(ledger["opening_balance"]))
            y -= row_h

        for t in page_txns:
            row_h = max(18, len(t["lines"]) * 11 + 6)
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)

            c.drawString(COLS[0] + 4, y - 12, t["date"].strftime("%d/%m/%Y"))
            c.drawString(COLS[2] + 4, y - 12, t["ref"])

            if t["debit"] > 0:
                c.drawRightString(COLS[4] - 4, y - 12, format_inr(t["debit"]))
            if t["credit"] > 0:
                c.drawRightString(COLS[5] - 4, y - 12, format_inr(t["credit"]))
            c.drawRightString(COLS[6] - 4, y - 12, format_inr(t["balance"]))

            text_y = y - 12
            for line_str in t["lines"]:
                c.drawString(COLS[1] + 4, text_y, line_str[:38])
                text_y -= 10.5

            y -= row_h

        c.setFont("Helvetica", 7.5)
        c.drawRightString(570, 32, f"Page {page_idx} of 2")
        c.showPage()

    c.save()
    return {"file": "bank_4_messy_headers.pdf", "pages": 2, "ledger": ledger}


def generate_bank_5_password_locked(output_dir: str) -> dict:
    """
    bank_5_password_locked.pdf:
    - Clean 2-page statement encrypted in-memory using `pypdf` (`PdfWriter.encrypt("VAULT123")`)
    """
    path = os.path.join(output_dir, "bank_5_password_locked.pdf")
    ledger = StatementDataGenerator.build_ledger(
        seed=505, num_txns=28, opening_balance="2150000.00", multiline_ratio=0.35
    )

    # Render clean 2-page PDF into RAM buffer
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    txns = ledger["transactions"]
    pages_data = [txns[0:14], txns[14:28]]
    COLS = [28, 86, 285, 355, 428, 500, 570]

    for page_idx, page_txns in enumerate(pages_data, start=1):
        c.setFont("Helvetica-Bold", 13)
        c.drawString(28, height - 38, "KOTAK MAHINDRA BANK - PRIVY CORPORATE STATEMENT (ENCRYPTED)")
        c.setFont("Helvetica", 8.5)
        c.drawString(28, height - 52, "Client: NEXTGEN FINTECH LABS PVT LTD | A/C: 4412098812 | IFSC: KKBK0000958")

        top_y = height - 70
        c.setFillColor(colors.HexColor("#EAECEE"))
        c.rect(COLS[0], top_y - 18, COLS[-1] - COLS[0], 18, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 8)

        headers = ["Date", "Narration", "Ref No", "Withdrawal (Dr)", "Deposit (Cr)", "Balance"]
        for idx_h, h_text in enumerate(headers):
            if idx_h >= 3:
                c.drawRightString(COLS[idx_h + 1] - 4, top_y - 12, h_text)
            else:
                c.drawString(COLS[idx_h] + 4, top_y - 12, h_text)

        for x in COLS:
            c.line(x, top_y, x, top_y - 18)

        y = top_y - 18
        c.setFont("Helvetica", 7.8)

        if page_idx == 1:
            row_h = 18
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)
            c.drawString(COLS[0] + 4, y - 12, "01/09/2026")
            c.drawString(COLS[1] + 4, y - 12, "B/F OPENING BALANCE")
            c.drawString(COLS[2] + 4, y - 12, "-")
            c.drawRightString(COLS[6] - 4, y - 12, format_inr(ledger["opening_balance"]))
            y -= row_h

        for t in page_txns:
            row_h = max(18, len(t["lines"]) * 11 + 6)
            c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
            for x in COLS:
                c.line(x, y, x, y - row_h)

            c.drawString(COLS[0] + 4, y - 12, t["date"].strftime("%d/%m/%Y"))
            c.drawString(COLS[2] + 4, y - 12, t["ref"])

            if t["debit"] > 0:
                c.drawRightString(COLS[4] - 4, y - 12, format_inr(t["debit"]))
            if t["credit"] > 0:
                c.drawRightString(COLS[5] - 4, y - 12, format_inr(t["credit"]))
            c.drawRightString(COLS[6] - 4, y - 12, format_inr(t["balance"]))

            text_y = y - 12
            for line_str in t["lines"]:
                c.drawString(COLS[1] + 4, text_y, line_str[:38])
                text_y -= 10.5

            y -= row_h

        c.setFont("Helvetica", 7.5)
        c.drawRightString(570, 35, f"Page {page_idx} of 2")
        c.showPage()

    c.save()
    buf.seek(0)

    # Encrypt with pypdf using password "VAULT123"
    reader = PdfReader(buf)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt("VAULT123")

    with open(path, "wb") as f_out:
        writer.write(f_out)

    return {"file": "bank_5_password_locked.pdf", "pages": 2, "ledger": ledger}


def generate_bank_6_scanned_image(output_dir: str) -> dict:
    """
    bank_6_scanned_image.pdf:
    - Render a 1-page statement to a high-DPI PIL Image using `pypdfium2` (>= 200 DPI)
    - Save the rasterized image directly as an image-only PDF with ZERO embedded selectable text layer
    """
    path = os.path.join(output_dir, "bank_6_scanned_image.pdf")
    ledger = StatementDataGenerator.build_ledger(
        seed=606, num_txns=14, opening_balance="740000.00", multiline_ratio=0.35
    )

    # Step 1: Render crisp vector PDF into RAM buffer
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    COLS = [28, 86, 285, 355, 428, 500, 570]

    c.setFont("Helvetica-Bold", 13)
    c.drawString(28, height - 38, "PUNJAB NATIONAL BANK - SCANNED VOUCHER STATEMENT")
    c.setFont("Helvetica", 8.5)
    c.drawString(28, height - 52, "Account: NEXTGEN FINTECH LABS PVT LTD | A/C: 019200210004812 | IFSC: PUNB0019200")

    top_y = height - 70
    c.setFillColor(colors.HexColor("#EAECEE"))
    c.rect(COLS[0], top_y - 18, COLS[-1] - COLS[0], 18, fill=1, stroke=1)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 8.5)

    headers = ["Date", "Narration", "Ref No", "Debit (Dr)", "Credit (Cr)", "Balance"]
    for idx_h, h_text in enumerate(headers):
        if idx_h >= 3:
            c.drawRightString(COLS[idx_h + 1] - 4, top_y - 12, h_text)
        else:
            c.drawString(COLS[idx_h] + 4, top_y - 12, h_text)

    for x in COLS:
        c.line(x, top_y, x, top_y - 18)

    y = top_y - 18
    c.setFont("Helvetica", 8)

    row_h = 18
    c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
    for x in COLS:
        c.line(x, y, x, y - row_h)
    c.drawString(COLS[0] + 4, y - 12, "01/09/2026")
    c.drawString(COLS[1] + 4, y - 12, "B/F OPENING BALANCE")
    c.drawString(COLS[2] + 4, y - 12, "-")
    c.drawRightString(COLS[6] - 4, y - 12, format_inr(ledger["opening_balance"]))
    y -= row_h

    for t in ledger["transactions"]:
        row_h = max(18, len(t["lines"]) * 11 + 6)
        c.rect(COLS[0], y - row_h, COLS[-1] - COLS[0], row_h, fill=0, stroke=1)
        for x in COLS:
            c.line(x, y, x, y - row_h)

        c.drawString(COLS[0] + 4, y - 12, t["date"].strftime("%d/%m/%Y"))
        c.drawString(COLS[2] + 4, y - 12, t["ref"])

        if t["debit"] > 0:
            c.drawRightString(COLS[4] - 4, y - 12, format_inr(t["debit"]))
        if t["credit"] > 0:
            c.drawRightString(COLS[5] - 4, y - 12, format_inr(t["credit"]))
        c.drawRightString(COLS[6] - 4, y - 12, format_inr(t["balance"]))

        text_y = y - 12
        for line_str in t["lines"]:
            c.drawString(COLS[1] + 4, text_y, line_str[:38])
            text_y -= 10.5

        y -= row_h

    c.setFont("Helvetica", 8)
    c.drawString(28, 36, "CERTIFIED TRUE SCANNED COPY - ARCHIVAL FLATBED SCAN @ 216 DPI")
    c.drawRightString(570, 36, "Page 1 of 1")
    c.showPage()
    c.save()

    # Step 2: Rasterize using pypdfium2 at 3.0x scale (~216 DPI)
    pdf_bytes = buf.getvalue()
    pdf_doc = pdfium.PdfDocument(pdf_bytes)
    page = pdf_doc[0]
    bitmap = page.render(scale=3.0)
    pil_img = bitmap.to_pil().convert("RGB")
    page.close()
    pdf_doc.close()

    # Apply subtle flatbed scanner visual realism while preserving high OCR legibility
    pil_img = pil_img.filter(ImageFilter.UnsharpMask(radius=1.0, percent=120, threshold=3))

    # Step 3: Save directly via Pillow as a pure raster PDF (ZERO embedded selectable text)
    pil_img.save(path, "PDF", resolution=216.0)

    return {"file": "bank_6_scanned_image.pdf", "pages": 1, "ledger": ledger}


# ==============================================================================
# 4. QA SELF-VERIFICATION & SUMMARY TABLE ENGINE
# ==============================================================================

def verify_and_summarize(output_dir: str, results: list):
    """Verifies mathematical accuracy, page counts, encryption, and zero-text OCR properties."""
    summary_rows = []

    for item in results:
        fname = item["file"]
        fpath = os.path.join(output_dir, fname)
        ledger = item["ledger"]
        txns = ledger["transactions"]

        # 1. Verify Running Balance Math: New Balance == Previous Balance + Credit - Debit
        running = ledger["opening_balance"]
        math_ok = True
        for t in txns:
            expected = running + t["credit"] - t["debit"]
            if expected != t["balance"]:
                math_ok = False
                break
            running = expected
        if running != ledger["closing_balance"]:
            math_ok = False

        # 2. Inspect generated PDF with pypdf
        reader = PdfReader(fpath)
        is_encrypted = reader.is_encrypted
        if is_encrypted:
            reader.decrypt("VAULT123")

        actual_pages = len(reader.pages)
        extracted_text = "".join((p.extract_text() or "") for p in reader.pages).strip()
        selectable_chars = len(extracted_text)

        summary_rows.append({
            "File Name": fname,
            "Pages": f"{actual_pages}/{item['pages']}",
            "Txns": len(txns),
            "Opening Bal (INR)": format_inr(ledger["opening_balance"]),
            "Closing Bal (INR)": format_inr(ledger["closing_balance"]),
            "Math Check": "PASS" if math_ok else "FAIL",
            "Encrypted": "YES (VAULT123)" if is_encrypted else "NO",
            "Text Layer Chars": selectable_chars,
        })

    summary_df = pd.DataFrame(summary_rows)
    print("\n" + "=" * 112)
    print("SYNTHETIC INDIAN BANK STATEMENT QA SUITE - VERIFICATION SUMMARY")
    print("=" * 112)
    print(summary_df.to_string(index=False))
    print("=" * 112)

    # Also verify vendor_master.csv
    v_df = pd.read_csv(os.path.join(output_dir, "vendor_master.csv"))
    print("\nVENDOR MASTER CSV (./test_statements/vendor_master.csv):")
    print("-" * 72)
    print(v_df.to_string(index=False))
    print("-" * 72)
    print(
        "Audit Triggers Embedded in Statements:\n"
        "  [1] Overcharge Trigger   : CLOUDHOST INDIA billed 1,12,100.00 vs Expected 94,400.00\n"
        "  [2] GST Mismatch Trigger : SHARMA PACKAGING billed 41,300.00 (18% GST) vs Expected 39,200.00 (12% GST)\n"
        "  [3] Duplicate Debit      : bank_4_messy_headers.pdf contains identical same-day SWIGGY INSTAMART debits\n"
        "  [4] Zero-Text OCR Test   : bank_6_scanned_image.pdf has 0 selectable characters (216 DPI raster PDF)\n"
    )


def main():
    output_dir = os.path.join(".", "test_statements")
    os.makedirs(output_dir, exist_ok=True)

    create_vendor_master_csv(output_dir)

    results = [
        generate_bank_1_borderless(output_dir),
        generate_bank_2_multiline_heavy(output_dir),
        generate_bank_3_single_amount_drcr(output_dir),
        generate_bank_4_messy_headers(output_dir),
        generate_bank_5_password_locked(output_dir),
        generate_bank_6_scanned_image(output_dir),
    ]

    verify_and_summarize(output_dir, results)


if __name__ == "__main__":
    main()