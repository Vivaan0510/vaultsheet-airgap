#!/usr/bin/env python3
"""
categorizer_and_auditor.py
--------------------------
100% Offline 3-Tier Transaction Categorization & 6-Rule Forensic Audit Engine
for Indian Bank Statements. Integrates seamlessly with `parser_engine.py` and
exports TallyPrime-ready accounting ledgers and voucher classifications.

Dependencies:
    pip install pandas numpy scikit-learn
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.error
import urllib.request
from typing import Any

import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression


# ==============================================================================
# 1. TAXONOMY & TALLYPRIME LEDGER MAPPING CONSTANTS
# ==============================================================================

ALLOWED_CATEGORIES: list[str] = [
    "Food & Dining",
    "Groceries & Quick-Commerce",
    "Travel & Commute",
    "Utilities & Telecom",
    "Vendor & B2B Transfer",
    "Salary & Income",
    "Tax & GST",
    "ATM & Cash",
    "Bank Charges & Hidden Fees",
    "Shopping & E-Commerce",
    "Uncategorized",
]

DEFAULT_TALLY_LEDGER_MAP: dict[str, str] = {
    "Food & Dining": "Staff Welfare & Food Exp",
    "Groceries & Quick-Commerce": "Office Pantry & Consumables A/c",
    "Travel & Commute": "Travelling & Conveyance Exp",
    "Utilities & Telecom": "Electricity, Internet & Utilities A/c",
    "Vendor & B2B Transfer": "Sundry Creditors / B2B Vendor A/c",
    "Salary & Income": "Sales & Operating Revenue A/c",
    "Tax & GST": "GST Output/Input Tax A/c",
    "ATM & Cash": "Petty Cash A/c",
    "Bank Charges & Hidden Fees": "Bank Charges A/c",
    "Shopping & E-Commerce": "Office Maintenance & Supplies Exp",
    "Uncategorized": "Suspense & Miscellaneous A/c",
}


# ==============================================================================
# 2. TIER 1: DETERMINISTIC INDIAN BANKING REGEX ENGINE (Confidence = 0.99)
# ==============================================================================

# Ordered by forensic specificity (e.g., Bank Fees & Tax checked before generic UPI/NEFT tokens)
TIER1_REGEX_RULES: list[tuple[re.Pattern, str]] = [
    (
        re.compile(
            r"\b(SMS\s+ALERT\s+CHG|MIN\s+BAL\s+PENALTY|CONVENIENCE\s+FEE|ANNUAL\s+FEE|"
            r"IMPS\s+CONVENIENCE\s+FEE|DEBIT\s+CARD\s+ANNUAL\s+FEE|PENAL\s+CHARGES?|"
            r"FOLIO\s+CHG|CHEQUE\s+RETURN\s+CHG|NACH\s+RETURN\s+CHARGES?)\b",
            re.IGNORECASE,
        ),
        "Bank Charges & Hidden Fees",
    ),
    (
        re.compile(
            r"\b(ATM\s+WDL|ATM\s+CASH|CASH\s+WDL|ATM-WDL|SELF\s+CHQ|CASH\s+WITHDRAWAL|"
            r"CARDLESS\s+CASH|NFS\s+CASH\s+WDL)\b",
            re.IGNORECASE,
        ),
        "ATM & Cash",
    ),
    (
        re.compile(
            r"\b(CBDT|GSTIN|CGST|SGST|IGST|TDS\s+PAY|INCOME\s+TAX|ADVANCE\s+TAX|"
            r"CHALLAN\s+NO|EPFO|ESIC\s+CONTRIBUTION|PROF\s+TAX)\b",
            re.IGNORECASE,
        ),
        "Tax & GST",
    ),
    (
        re.compile(
            r"\b(SWIGGY\s+INSTAMART|BLINKIT|ZEPTO|BIGBASKET|DUNZO\s+DAILY|"
            r"JIOMART\s+GROCERY|RELIANCE\s+SMART|DMART\s+READY)\b",
            re.IGNORECASE,
        ),
        "Groceries & Quick-Commerce",
    ),
    (
        re.compile(
            r"\b(SWIGGY|ZOMATO|STARBUCKS|TATA\s+STARBUCKS|DOMINOS|MCDONALDS|"
            r"KFC\s+INDIA|BURGER\s+KING|HALDIRAM|BARBEQUE\s+NATION|CAFE\s+COFFEE)\b",
            re.IGNORECASE,
        ),
        "Food & Dining",
    ),
    (
        re.compile(
            r"\b(UBER\s+INDIA|UBER\s+TRIP|UBER|OLA\s+CABS|OLACABS|RAPIDO|IRCTC|"
            r"DELHI\s+METRO|MUMBAI\s+METRO|BMTC|INDIGO\s+AIRLINES|AIR\s+INDIA|"
            r"MAKEMYTRIP|CLEARTRIP|REDBUS|FASTAG|NHAI\s+TOLL)\b",
            re.IGNORECASE,
        ),
        "Travel & Commute",
    ),
    (
        re.compile(
            r"\b(AIRTEL\s+PAYMENTS|AIRTEL\s+BROADBAND|JIO\s+FIBER|RELIANCE\s+JIO|"
            r"VODAFONE\s+IDEA|BSNL|TATA\s+POWER|ADANI\s+ELECTRICITY|BESCOM|MSEB|"
            r"MAHANAGAR\s+GAS|INDRAPRASTHA\s+GAS|ACT\s+FIBERNET)\b",
            re.IGNORECASE,
        ),
        "Utilities & Telecom",
    ),
    (
        re.compile(
            r"\b(AMAZON\s+PAY|AMAZON\s+INDIA|FLIPKART|MYNTRA|AJIO|NYKAA|MEESHO|"
            r"RELIANCE\s+RETAIL|CROMA\s+RETAIL|PHARMEASY|TATA\s+1MG|APOLLO\s+PHARMACY|"
            r"DECATHLON|IKEA\s+INDIA)\b",
            re.IGNORECASE,
        ),
        "Shopping & E-Commerce",
    ),
    (
        re.compile(
            r"\b(SALARY\s+CREDIT|PAYROLL\s+DISBURSEMENT|RAZORPAY\s+SOFTWARE.*SETTLEMENT|"
            r"STRIPE\s+INDIA.*PAYOUT|CLIENT\s+RETENTION\s+MILESTONE|CLIENT\s+INVOICE|"
            r"CREDIT\s+INTEREST|INT\.PD:|DIVIDEND\s+WARRANT)\b",
            re.IGNORECASE,
        ),
        "Salary & Income",
    ),
    (
        re.compile(
            r"\b(APEX\s+LOGISTICS|CLOUDHOST\s+INDIA|SHARMA\s+PACKAGING|METRO\s+POWER|"
            r"NEFT/DR/|RTGS/DR/|IMPS/DR/|B2B\s+INVOICE|FREIGHT\s+CHARGES|"
            r"DEDICATED\s+SERVER|CORRUGATED\s+BOX\s+SUPPLY|DG\s+SET\s+AMNT)\b",
            re.IGNORECASE,
        ),
        "Vendor & B2B Transfer",
    ),
]


def _match_tier1_regex(description: str, debit: float, credit: float) -> tuple[str | None, float]:
    """Evaluate Tier 1 deterministic regex patterns with credit/debit context."""
    text = (description or "").strip()
    if not text:
        return None, 0.0

    # If money came in via NEFT/IMPS/ACH/RTGS credit, prioritize Salary & Income
    if credit > 0 and debit == 0:
        if re.search(
            r"\b(NEFT/CR|IMPS/CR|ACH/CR|RTGS/CR|SETTLEMENT|PAYOUT|INVOICE|INT\.PD|SALARY)\b",
            text,
            re.IGNORECASE,
        ):
            return "Salary & Income", 0.99

    for pattern, category in TIER1_REGEX_RULES:
        if pattern.search(text):
            # Guard against classifying GST remarks inside a B2B vendor invoice as direct Tax & GST
            if category == "Tax & GST" and re.search(
                r"\b(NEFT/DR|RTGS/DR|PACKAGING|LOGISTICS|CLOUDHOST|METRO\s+POWER)\b",
                text,
                re.IGNORECASE,
            ):
                return "Vendor & B2B Transfer", 0.99
            return category, 0.99

    # Secondary check for generic NEFT/RTGS debit transfers to corporate entities
    if debit > 0 and re.search(r"\b(NEFT|RTGS|PVT\s+LTD|PRIVATE\s+LIMITED|INDUSTRIES|SERVICES)\b", text, re.IGNORECASE):
        return "Vendor & B2B Transfer", 0.99

    return None, 0.0


# ==============================================================================
# 3. TIER 3: IN-MEMORY SCIKIT-LEARN TF-IDF + LOGISTIC REGRESSION (90 SAMPLES)
# ==============================================================================

EMBEDDED_90_SAMPLE_CORPUS: list[tuple[str, str]] = [
    # 1. Food & Dining (9 samples)
    ("UPI/DR/SWIGGY ORDER FOOD DELIVERY BANGALORE", "Food & Dining"),
    ("UPI/DR/ZOMATO LIMITED ONLINE RESTAURANT BILL", "Food & Dining"),
    ("POS TATA STARBUCKS COFFEE MUMBAI BEVERAGE", "Food & Dining"),
    ("UPI DOMINOS PIZZA JUBILANT FOODWORKS ORDER", "Food & Dining"),
    ("MCDONALDS CONNAUGHT PLAZA RESTAURANT POS", "Food & Dining"),
    ("UPI KFC INDIA YUM RESTAURANTS DINING", "Food & Dining"),
    ("HALDIRAM SNACKS AND SWEETS FOOD COURT", "Food & Dining"),
    ("BARBEQUE NATION HOSPITALITY TEAM LUNCH", "Food & Dining"),
    ("CAFE COFFEE DAY ESPRESSO REFRESHMENTS", "Food & Dining"),
    # 2. Groceries & Quick-Commerce (9 samples)
    ("UPI/DR/SWIGGY INSTAMART GROCERY DELIVERY", "Groceries & Quick-Commerce"),
    ("UPI/DR/BLINKIT GROCERY INSTANT DELIVERY", "Groceries & Quick-Commerce"),
    ("ZEPTO KIRANA KART TECHNOLOGIES QUICK COMMERCE", "Groceries & Quick-Commerce"),
    ("BIGBASKET DAILY SUPERMARKET WALLET RECHARGE", "Groceries & Quick-Commerce"),
    ("RELIANCE SMART BAZAAR JIOMART GROCERY BILL", "Groceries & Quick-Commerce"),
    ("DMART READY AVENUE SUPERMARTS HOUSEHOLD SUPPLIES", "Groceries & Quick-Commerce"),
    ("DUNZO DAILY FRESH FRUITS AND VEGETABLES", "Groceries & Quick-Commerce"),
    ("NATURE BASKET GOURMET GROCERY STORE MUMBAI", "Groceries & Quick-Commerce"),
    ("COUNTRY DELIGHT MILK AND DAIRY SUBSCRIPTION", "Groceries & Quick-Commerce"),
    # 3. Travel & Commute (9 samples)
    ("UPI/DR/UBER INDIA SYSTEMS TRIP RIDE MUMBAI", "Travel & Commute"),
    ("OLA CABS ANI TECHNOLOGIES AIRPORT TAXI", "Travel & Commute"),
    ("IRCTC RAIL TICKET BOOKING EXPRESS TRAIN", "Travel & Commute"),
    ("INDIGO AIRLINES INTERGLOBE AVIATION FLIGHT TICKET", "Travel & Commute"),
    ("MUMBAI METRO ONE SMART CARD RECHARGE COMMUTE", "Travel & Commute"),
    ("RAPIDO BIKE TAXI CITY COMMUTE RIDE", "Travel & Commute"),
    ("MAKEMYTRIP INDIA HOTEL AND FLIGHT BOOKING", "Travel & Commute"),
    ("ICICI FASTAG NHAI HIGHWAY TOLL PLAZA DEDUCTION", "Travel & Commute"),
    ("INDIAN OIL PETROL PUMP FUEL STATION POS", "Travel & Commute"),
    # 4. Utilities & Telecom (9 samples)
    ("UPI/DR/AIRTEL PAYMENTS BROADBAND FIBER BILL", "Utilities & Telecom"),
    ("RELIANCE JIO INFOCOMM POSTPAID MOBILE RECHARGE", "Utilities & Telecom"),
    ("TATA POWER ELECTRICITY BILL PAYMENT MUMBAI", "Utilities & Telecom"),
    ("ADANI ELECTRICITY SUBURBAN POWER SUPPLY", "Utilities & Telecom"),
    ("VODAFONE IDEA CORPORATE POSTPAID PLAN", "Utilities & Telecom"),
    ("ACT FIBERNET ATRIA CONVERGENCE BROADBAND INTERNET", "Utilities & Telecom"),
    ("MAHANAGAR GAS LIMITED PIPED GAS UTILITY BILL", "Utilities & Telecom"),
    ("BESCOM BANGALORE ELECTRICITY SUPPLY BILL", "Utilities & Telecom"),
    ("BSNL LANDLINE AND LEASED LINE TELECOM BILL", "Utilities & Telecom"),
    # 5. Vendor & B2B Transfer (9 samples)
    ("NEFT/DR/APEX LOGISTICS PVT LTD FREIGHT CHARGES", "Vendor & B2B Transfer"),
    ("NEFT/DR/CLOUDHOST INDIA SERVICES DEDICATED SERVER", "Vendor & B2B Transfer"),
    ("NEFT/DR/SHARMA PACKAGING INDUSTRIES CORRUGATED BOX", "Vendor & B2B Transfer"),
    ("RTGS/DR/METRO POWER INFRASTRUCTURE DG SET AMNT", "Vendor & B2B Transfer"),
    ("NEFT B2B VENDOR INVOICE SETTLEMENT RAW MATERIALS", "Vendor & B2B Transfer"),
    ("IMPS/DR/CONSULTANCY RETAINER PROFESSIONAL FEES", "Vendor & B2B Transfer"),
    ("RTGS WAREHOUSE LEASE RENT COMMERCIAL PROPERTY", "Vendor & B2B Transfer"),
    ("NEFT SOFTWARE LICENSE ANNUAL AMC CONTRACT", "Vendor & B2B Transfer"),
    ("NEFT SECURITY AGENCY MANPOWER OUTSOURCING BILL", "Vendor & B2B Transfer"),
    # 6. Salary & Income (9 samples)
    ("NEFT/CR/RAZORPAY SOFTWARE PVT LTD SETTLEMENT", "Salary & Income"),
    ("NEFT/CR/STRIPE INDIA PRIVATE LTD PAYOUT", "Salary & Income"),
    ("IMPS/CR/CLIENT RETENTION MILESTONE PAYMENT", "Salary & Income"),
    ("ACH/CR/TECHFLOW CLIENT INVOICE SETTLEMENT", "Salary & Income"),
    ("INT.PD CREDIT INTEREST QUARTERLY SAVINGS", "Salary & Income"),
    ("NEFT/CR/CUSTOMER SUBSCRIPTION REVENUE COLLECTION", "Salary & Income"),
    ("SALARY CREDIT MONTHLY PAYROLL REIMBURSEMENT", "Salary & Income"),
    ("RTGS/CR/ENTERPRISE ANNUAL CONTRACT ADVANCE", "Salary & Income"),
    ("ACH/CR/DIVIDEND PAYOUT CORPORATE TREASURY", "Salary & Income"),
    # 7. Tax & GST (9 samples)
    ("NEFT CBDT INCOME TAX ADVANCE TAX CHALLAN", "Tax & GST"),
    ("GST PORTAL CGST SGST IGST MONTHLY TAX PAYMENT", "Tax & GST"),
    ("TDS DEDUCTION SECTION 194J PROFESSIONAL TAX", "Tax & GST"),
    ("EPFO EMPLOYER PROVIDENT FUND CONTRIBUTION CHALLAN", "Tax & GST"),
    ("ESIC STATUTORY EMPLOYEE INSURANCE CONTRIBUTION", "Tax & GST"),
    ("MAHARASHTRA STATE PROFESSION TAX RETURN", "Tax & GST"),
    ("CUSTOMS DUTY ICEGATE IMPORT CLEARANCE TAX", "Tax & GST"),
    ("PROPERTY TAX MUNICIPAL CORPORATION ASSESSMENT", "Tax & GST"),
    ("STAMP DUTY AND REGISTRATION FEE GOVT TREASURY", "Tax & GST"),
    # 8. ATM & Cash (9 samples)
    ("ATM WDL HDFC BANK BRANCH CASH WITHDRAWAL", "ATM & Cash"),
    ("NFS CASH WDL SBI ATM TERMINAL MUMBAI", "ATM & Cash"),
    ("SELF CHEQUE CASH WITHDRAWAL COUNTER TELLER", "ATM & Cash"),
    ("CARDLESS CASH WITHDRAWAL IMT ATM TERMINAL", "ATM & Cash"),
    ("ATM CASH DISPENSER AXIS BANK NARIMAN POINT", "ATM & Cash"),
    ("PETTY CASH REPLENISHMENT BRANCH WITHDRAWAL", "ATM & Cash"),
    ("ICICI ATM CASH WDL LOWER PAREL MUMBAI", "ATM & Cash"),
    ("KOTAK BANK ATM CASH WITHDRAWAL TXN", "ATM & Cash"),
    ("CASH DEPOSIT MACHINE CDM SELF DEPOSIT", "ATM & Cash"),
    # 9. Bank Charges & Hidden Fees (9 samples)
    ("SMS ALERT CHG Q3 QUARTERLY NOTIFICATION FEE", "Bank Charges & Hidden Fees"),
    ("MIN BAL PENALTY NON MAINTENANCE OF MAB", "Bank Charges & Hidden Fees"),
    ("IMPS CONVENIENCE FEE TRANSACTION SERVICE CHARGE", "Bank Charges & Hidden Fees"),
    ("DEBIT CARD ANNUAL FEEAMC ISSUANCE CHARGE", "Bank Charges & Hidden Fees"),
    ("NEFT RTGS PROCESSING SERVICE CHARGE GST", "Bank Charges & Hidden Fees"),
    ("CHEQUE BOUNCE RETURN PENALTY CHARGES", "Bank Charges & Hidden Fees"),
    ("NACH MANDATE DISHONOUR PENAL CHARGE", "Bank Charges & Hidden Fees"),
    ("ACCOUNT FOLIO MAINTENANCE CHARGES ANNUAL", "Bank Charges & Hidden Fees"),
    ("DUPLICATE STATEMENT ISSUANCE BANK FEE", "Bank Charges & Hidden Fees"),
    # 10. Shopping & E-Commerce (9 samples)
    ("UPI/DR/AMAZON PAY INDIA MARKETPLACE ORDER", "Shopping & E-Commerce"),
    ("FLIPKART INTERNET PRIVATE LIMITED ELECTRONICS", "Shopping & E-Commerce"),
    ("RELIANCE RETAIL DIGITAL STORE APPLIANCE", "Shopping & E-Commerce"),
    ("PHARMEASY HEALTH MEDICAL PHARMACY ORDER", "Shopping & E-Commerce"),
    ("MYNTRA DESIGNS APPAREL AND FOOTWEAR", "Shopping & E-Commerce"),
    ("NYKAA E RETAIL PERSONAL CARE COSMETICS", "Shopping & E-Commerce"),
    ("CROMA INFINITI RETAIL LAPTOP ACCESSORIES", "Shopping & E-Commerce"),
    ("TATA 1MG HEALTHCARE MEDICINE DELIVERY", "Shopping & E-Commerce"),
    ("IKEA INDIA OFFICE FURNITURE AND DESK", "Shopping & E-Commerce"),
]

# Fit the in-memory TF-IDF + LogisticRegression pipeline on module import (<50ms)
_CORPUS_TEXTS = [item[0] for item in EMBEDDED_90_SAMPLE_CORPUS]
_CORPUS_LABELS = [item[1] for item in EMBEDDED_90_SAMPLE_CORPUS]

_TFIDF_VECTORIZER = TfidfVectorizer(
    ngram_range=(1, 2),
    lowercase=True,
    token_pattern=r"(?u)\b[A-Za-z0-9]{2,}\b",
    sublinear_tf=True,
)
_X_TRAIN = _TFIDF_VECTORIZER.fit_transform(_CORPUS_TEXTS)

_CLF_MODEL = LogisticRegression(C=5.0, max_iter=300, random_state=42)
_CLF_MODEL.fit(_X_TRAIN, _CORPUS_LABELS)


def _predict_tier3_sklearn(description: str) -> tuple[str, float]:
    """Classify transaction narration using the embedded TF-IDF + LogisticRegression model."""
    cleaned = re.sub(r"[^A-Za-z0-9\s]", " ", description or "").strip()
    if not cleaned:
        return "Uncategorized", 0.0

    vec = _TFIDF_VECTORIZER.transform([cleaned])
    if vec.nnz == 0:
        return "Uncategorized", 0.25

    probs = _CLF_MODEL.predict_proba(vec)[0]
    best_idx = int(np.argmax(probs))
    best_cat = str(_CLF_MODEL.classes_[best_idx])
    confidence = round(float(probs[best_idx]), 2)

    if confidence < 0.22:
        return "Uncategorized", confidence
    return best_cat, max(0.55, min(0.94, confidence))


# ==============================================================================
# 4. TIER 2: OPTIONAL LOCAL OLLAMA LLM (`http://127.0.0.1:11434/api/generate`)
# ==============================================================================

def _query_local_ollama(description: str, debit: float, credit: float) -> tuple[str | None, float]:
    """
    Attempt 100% local classification via Ollama on 127.0.0.1:11434 with a strict
    1.5-second timeout. Silently returns (None, 0.0) if Ollama is unreachable.
    """
    url = "http://127.0.0.1:11434/api/generate"
    direction = "CREDIT (Deposit)" if credit > 0 else "DEBIT (Withdrawal)"
    prompt = (
        "Classify this Indian bank transaction into EXACTLY ONE of these categories: "
        f"{', '.join(ALLOWED_CATEGORIES)}.\n"
        f"Transaction ({direction}): {description}\n"
        "Respond with JSON only: {\"category\": \"<exact_category_name>\"}"
    )

    for model_name in ("llama3.2:1b", "qwen2.5:1.5b"):
        payload = json.dumps({
            "model": model_name,
            "prompt": prompt,
            "stream": False,
            "format": "json",
        }).encode("utf-8")

        req = urllib.request.Request(
            url,
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=1.5) as resp:
                body = json.loads(resp.read().decode("utf-8"))
                raw_resp = body.get("response", "{}")
                parsed = json.loads(raw_resp)
                cat = parsed.get("category", "").strip()
                if cat in ALLOWED_CATEGORIES and cat != "Uncategorized":
                    return cat, 0.92
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError, ValueError):
            break

    return None, 0.0


# ==============================================================================
# 5. TALLYPRIME LEDGER & VOUCHER TYPE MAPPER
# ==============================================================================

def _derive_tally_fields(
    description: str,
    category: str,
    debit: float,
    credit: float,
    matched_vendor_name: str | None = None,
) -> tuple[str, str]:
    """
    Map transaction into (`tally_ledger`, `tally_voucher_type`) for TallyPrime XML/CSV import.
    - `tally_voucher_type`: 'Contra' for ATM/Cash withdrawals or self-transfers,
      'Receipt' if credit > 0, else 'Payment'.
    """
    desc_up = (description or "").upper()

    # Determine Voucher Type
    if category == "ATM & Cash" or any(
        k in desc_up for k in ("ATM WDL", "CASH WDL", "SELF TRF", "CONTRA", "SELF CHQ")
    ):
        voucher_type = "Contra"
    elif credit > 0:
        voucher_type = "Receipt"
    else:
        voucher_type = "Payment"

    # Determine Tally Ledger Name
    if matched_vendor_name:
        return f"{matched_vendor_name.title()} A/c", voucher_type

    if category == "Vendor & B2B Transfer":
        # Try extracting the corporate entity name from NEFT/RTGS slash-delimited string
        m = re.search(
            r"(?:NEFT|RTGS|IMPS)/(?:DR|CR)/[A-Z0-9]+/([A-Z0-9\s&\.-]{4,35}?)(?:/|\bSBIN|\bHDFC|\bICIC|\bUTIB|\bKKBK|$)",
            desc_up,
        )
        if m:
            party = re.sub(r"\s+", " ", m.group(1)).strip().title()
            return f"{party} A/c", voucher_type

    return DEFAULT_TALLY_LEDGER_MAP.get(category, "Suspense & Miscellaneous A/c"), voucher_type


def _extract_merchant_signature(description: str) -> str:
    """
    Extract a canonical merchant/beneficiary signature for duplicate charge detection.
    Strips volatile UTR numbers, dates, and reference codes.
    """
    s = (description or "").upper()
    # Check structured UPI/NEFT format: UPI/DR/<UTR>/<MERCHANT>/...
    m = re.search(r"(?:UPI|NEFT|RTGS|IMPS)/(?:DR|CR)/[A-Z0-9]+/([A-Z0-9\s]{3,25})", s)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()

    # Fallback: strip digits and slash tokens
    cleaned = re.sub(r"\b[A-Z0-9]*\d+[A-Z0-9]*\b", " ", s)
    cleaned = re.sub(r"[^A-Z\s]", " ", cleaned)
    tokens = [t for t in cleaned.split() if len(t) >= 3 and t not in ("UPI", "NEFT", "IMPS", "RTGS", "MUMBAI")]
    return " ".join(tokens[:3]) if tokens else s[:20]


# ==============================================================================
# 6. PRIMARY SHARED DATA CONTRACT FUNCTION
# ==============================================================================

def enrich_and_audit_ledger(
    df: pd.DataFrame,
    metadata: dict[str, Any],
    vendor_csv_path: str | None = None,
    use_local_llm: bool = False,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Enriches the 9-column parsed bank statement DataFrame with 6 accounting & forensic
    audit columns (15 columns total) and updates `metadata` with 3 audit summary metrics.
    """
    required_input_cols = [
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
    missing = [c for c in required_input_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Input DataFrame is missing required columns: {missing}")

    out_df = df[required_input_cols].copy()

# Load B2B Vendor Master CSV if provided
    vendor_rules: list[dict[str, Any]] = []
    if vendor_csv_path and os.path.exists(vendor_csv_path):
        v_df = pd.read_csv(vendor_csv_path, encoding="utf-8-sig")
        v_df.columns = v_df.columns.str.strip().str.lower()
        for _, v_row in v_df.iterrows():
            kw = str(v_row.get("vendor_keyword", "")).strip()
            if kw:
                vendor_rules.append({
                    "vendor_keyword": kw.upper(),
                    "contract_amount": round(float(v_row["contract_amount"]), 2),
                    "gst_rate_pct": float(v_row["gst_rate_pct"]),
                    "expected_total": round(float(v_row["expected_total"]), 2),
                })    # Pre-compute statistical outlier threshold for Rule 4 (minimum 5 debit rows)
    positive_debits = out_df.loc[out_df["debit"] > 0, "debit"].astype(float)
    if len(positive_debits) >= 5:
        debit_mean = float(positive_debits.mean())
        debit_std = float(positive_debits.std(ddof=1))
        outlier_threshold = debit_mean + 2.5 * debit_std
    else:
        outlier_threshold = float("inf")

    # Pre-compute Duplicate Charge pairs for Rule 2 (same debit > 0 and merchant within 1 calendar day)
    parsed_dates = pd.to_datetime(out_df["date"], format="%Y-%m-%d", errors="coerce")
    merch_sigs = [
        _extract_merchant_signature(str(desc)) for desc in out_df["description"]
    ]
    duplicate_indices: set[int] = set()

    for i in range(len(out_df)):
        d_i = float(out_df["debit"].iloc[i])
        if d_i <= 0 or pd.isna(parsed_dates.iloc[i]):
            continue
        for j in range(i + 1, len(out_df)):
            d_j = float(out_df["debit"].iloc[j])
            if d_i != d_j or pd.isna(parsed_dates.iloc[j]):
                continue
            day_diff = abs((parsed_dates.iloc[j] - parsed_dates.iloc[i]).days)
            if day_diff <= 1 and merch_sigs[i] and merch_sigs[i] == merch_sigs[j]:
                duplicate_indices.add(j)

    # Row-by-row 3-Tier Categorization & 6-Rule Forensic Audit
    categories: list[str] = []
    tally_ledgers: list[str] = []
    tally_vouchers: list[str] = []
    confidences: list[float] = []
    anomaly_flags: list[bool] = []
    anomaly_reasons: list[str] = []

    ollama_available = True
    vendor_mismatches_count = 0

    for idx, row in out_df.iterrows():
        desc = str(row["description"] or "")
        desc_upper = desc.upper()
        debit = round(float(row["debit"]), 2)
        credit = round(float(row["credit"]), 2)
        math_ok = bool(row["math_verified"])

        # Check if row matches a B2B Vendor in vendor_master.csv
        matched_vendor = None
        for v_rule in vendor_rules:
            if v_rule["vendor_keyword"] in desc_upper:
                matched_vendor = v_rule
                break

        # --- 3-Tier Categorization ---
        if matched_vendor is not None:
            cat, conf = "Vendor & B2B Transfer", 0.99
        else:
            cat, conf = _match_tier1_regex(desc, debit, credit)
            if cat is None and use_local_llm and ollama_available:
                llm_cat, llm_conf = _query_local_ollama(desc, debit, credit)
                if llm_cat is not None:
                    cat, conf = llm_cat, llm_conf
                else:
                    ollama_available = False  # Avoid repeated timeouts if Ollama is offline
            if cat is None:
                cat, conf = _predict_tier3_sklearn(desc)

        if cat not in ALLOWED_CATEGORIES:
            cat = "Uncategorized"

        # Derive TallyPrime fields
        t_ledger, t_voucher = _derive_tally_fields(
            desc,
            cat,
            debit,
            credit,
            matched_vendor_name=matched_vendor["vendor_keyword"] if matched_vendor else None,
        )

        # --- 6-Rule Forensic & B2B Vendor Audit Engine ---
        reasons: list[str] = []

        # Rule 1: Math Discrepancy
        if not math_ok:
            reasons.append("Running-balance math mismatch")

        # Rule 2: Duplicate Charge
        if idx in duplicate_indices:
            reasons.append("Duplicate debit charge detected")

        # Rule 3: Hidden Bank Fee Trap
        if cat == "Bank Charges & Hidden Fees":
            reasons.append("Hidden bank fee / penalty charge")

        # Rule 4: Statistical Outlier
        if debit > 0 and debit > outlier_threshold:
            reasons.append("Statistical high-value outflow outlier")

        # Rule 5: Round-Sum Cash/Transfer
        if debit >= 25000.0 and round(debit % 5000.0, 2) == 0.0:
            reasons.append("Large round-sum outflow")

        # Rule 6: B2B Vendor Master & GST Reconciliation
        if matched_vendor is not None and debit > 0:
            v_mismatch_triggered = False
            expected_total = matched_vendor["expected_total"]
            contract_amt = matched_vendor["contract_amount"]
            gst_rate = matched_vendor["gst_rate_pct"]

            if debit > expected_total:
                reasons.append("Vendor overcharge vs contract master")
                v_mismatch_triggered = True

            implied_base = round(debit / (1.0 + gst_rate / 100.0), 2)
            if implied_base != contract_amt:
                reasons.append("Vendor GST rate/math mismatch")
                v_mismatch_triggered = True

            if v_mismatch_triggered:
                vendor_mismatches_count += 1

        categories.append(cat)
        tally_ledgers.append(t_ledger)
        tally_vouchers.append(t_voucher)
        confidences.append(round(float(conf), 2))
        anomaly_flags.append(len(reasons) > 0)
        anomaly_reasons.append("; ".join(reasons))

    # Append exact 6 required columns (15 columns total)
    out_df["category"] = pd.Series(categories, index=out_df.index, dtype="object")
    out_df["tally_ledger"] = pd.Series(tally_ledgers, index=out_df.index, dtype="object")
    out_df["tally_voucher_type"] = pd.Series(tally_vouchers, index=out_df.index, dtype="object")
    out_df["confidence"] = pd.Series(confidences, index=out_df.index, dtype="float64")
    out_df["anomaly_flag"] = pd.Series(anomaly_flags, index=out_df.index, dtype="bool")
    out_df["anomaly_reason"] = pd.Series(anomaly_reasons, index=out_df.index, dtype="object")

    # Compute the 3 additional metadata keys
    hidden_bank_fees_total = round(
        float(out_df.loc[out_df["category"] == "Bank Charges & Hidden Fees", "debit"].sum()),
        2,
    )
    flagged_anomalies_count = int(out_df["anomaly_flag"].sum())

    updated_metadata = dict(metadata)
    updated_metadata["hidden_bank_fees_total"] = float(hidden_bank_fees_total)
    updated_metadata["flagged_anomalies_count"] = int(flagged_anomalies_count)
    updated_metadata["vendor_mismatches_count"] = int(vendor_mismatches_count)

    return out_df, updated_metadata


# ==============================================================================
# 7. CLI & VERIFICATION SUITE ACROSS ALL 6 SYNTHETIC PDFS
# ==============================================================================

def _run_full_audit_suite(use_local_llm: bool = False):
    """Runs parser_engine + categorizer_and_auditor across all 6 synthetic test PDFs."""
    try:
        from parser_engine import parse_bank_pdf
    except ImportError as exc:
        print(f"[ERROR] Could not import parser_engine.py: {exc}", file=sys.stderr)
        sys.exit(1)

    test_dir = os.path.join(".", "test_statements")
    vendor_csv = os.path.join(test_dir, "vendor_master.csv")
    test_files = [
        ("bank_1_borderless.pdf", None),
        ("bank_2_multiline_heavy.pdf", None),
        ("bank_3_single_amount_drcr.pdf", None),
        ("bank_4_messy_headers.pdf", None),
        ("bank_5_password_locked.pdf", "VAULT123"),
        ("bank_6_scanned_image.pdf", None),
    ]

    summary_rows = []
    sample_anomalies = []

    for fname, pw in test_files:
        fpath = os.path.join(test_dir, fname)
        if not os.path.exists(fpath):
            print(f"[WARN] Missing test file: {fpath}. Run generate_test_pdfs.py first.", file=sys.stderr)
            continue

        raw_df, raw_meta = parse_bank_pdf(fpath, password=pw)
        enriched_df, enriched_meta = enrich_and_audit_ledger(
            raw_df,
            raw_meta,
            vendor_csv_path=vendor_csv if os.path.exists(vendor_csv) else None,
            use_local_llm=use_local_llm,
        )

        unclassified_cnt = int((enriched_df["category"] == "Uncategorized").sum())
        avg_conf = round(float(enriched_df["confidence"].mean()), 2) if not enriched_df.empty else 0.0

        summary_rows.append({
            "Statement PDF": fname,
            "Cols": len(enriched_df.columns),
            "Rows": enriched_meta["total_rows"],
            "Avg Conf": f"{avg_conf:.2f}",
            "Uncat": unclassified_cnt,
            "Hidden Fees (INR)": f"{enriched_meta['hidden_bank_fees_total']:,.2f}",
            "Anomalies": enriched_meta["flagged_anomalies_count"],
            "Vendor Mismatch": enriched_meta["vendor_mismatches_count"],
        })

        # Collect sample flagged rows from bank_4_messy_headers.pdf to display audit reasons
        if fname == "bank_4_messy_headers.pdf":
            flagged = enriched_df[enriched_df["anomaly_flag"]]
            for _, r in flagged.iterrows():
                sample_anomalies.append({
                    "tx_id": int(r["tx_id"]),
                    "date": r["date"],
                    "category": r["category"],
                    "tally_ledger": r["tally_ledger"],
                    "voucher": r["tally_voucher_type"],
                    "debit": f"{r['debit']:,.2f}",
                    "anomaly_reason": r["anomaly_reason"],
                })

    if summary_rows:
        print("\n" + "=" * 118)
        print("FORENSIC AUDITOR & ML CATEGORIZER - ENRICHED SUITE SUMMARY (15-COLUMN DATA CONTRACT)")
        print("=" * 118)
        print(pd.DataFrame(summary_rows).to_string(index=False))
        print("=" * 118)

    if sample_anomalies:
        print("\nSAMPLE FLAGGED FORENSIC ANOMALIES & B2B RECONCILIATION HITS (bank_4_messy_headers.pdf):")
        print("-" * 118)
        print(pd.DataFrame(sample_anomalies).to_string(index=False))
        print("-" * 118 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="100% Offline 3-Tier ML Categorizer & 6-Rule Forensic Auditor"
    )
    parser.add_argument("--pdf", type=str, help="Path to a single PDF file to parse and audit")
    parser.add_argument("--password", type=str, default=None, help="Optional PDF password")
    parser.add_argument(
        "--vendor-csv",
        type=str,
        default=os.path.join(".", "test_statements", "vendor_master.csv"),
        help="Path to B2B vendor_master.csv",
    )
    parser.add_argument(
        "--use-local-llm",
        action="store_true",
        help="Enable optional Tier-2 local Ollama LLM via http://127.0.0.1:11434",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output enriched JSON {metadata, rows} to stdout for n8n workflows",
    )
    args = parser.parse_args()

    if args.pdf:
        from parser_engine import parse_bank_pdf

        raw_df, raw_meta = parse_bank_pdf(args.pdf, password=args.password)
        v_path = args.vendor_csv if (args.vendor_csv and os.path.exists(args.vendor_csv)) else None
        enriched_df, enriched_meta = enrich_and_audit_ledger(
            raw_df,
            raw_meta,
            vendor_csv_path=v_path,
            use_local_llm=args.use_local_llm,
        )
        if args.json:
            print(
                json.dumps(
                    {
                        "metadata": enriched_meta,
                        "rows": enriched_df.to_dict(orient="records"),
                    },
                    indent=2,
                )
            )
        else:
            print("=== ENRICHED METADATA ===")
            print(json.dumps(enriched_meta, indent=2))
            print("\n=== ENRICHED & AUDITED LEDGER (15 COLUMNS) ===")
            print(enriched_df.to_string(index=False))
    else:
        _run_full_audit_suite(use_local_llm=args.use_local_llm)


if __name__ == "__main__":
    main()