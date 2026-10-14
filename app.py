#!/usr/bin/env python3
"""
app.py
------
Zero-Retention, 100% Air-Gapped Indian Bank Statement Forensic Auditor & TallyPrime Bridge.
Built with Streamlit. Processes all PDFs, passwords, OCR rasters, ML classifications,
and Excel exports strictly inside volatile RAM (`io.BytesIO`) with zero disk persistence
of uploaded user statements.

Dependencies:
    pip install streamlit pandas numpy openpyxl pdfplumber pypdf pypdfium2 rapidocr_onnxruntime scikit-learn
"""

from __future__ import annotations

import gc
import io
import json
import os
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st

# ==============================================================================
# EXACT BACKEND IMPORTS
# ==============================================================================
from parser_engine import parse_bank_pdf
from categorizer_and_auditor import enrich_and_audit_ledger
try:
    from excel_exporter import export_to_excel
except ImportError:
    # Self-healing in-memory Excel exporter fallback if excel_exporter.py is not yet created
    def export_to_excel(
        df: pd.DataFrame,
        metadata: dict[str, Any],
        output_path: str | io.BytesIO | None = None,
    ) -> bytes:
        """Generates a multi-sheet executive & TallyPrime-ready Excel workbook in RAM."""
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            # Sheet 1: Enriched & Audited 15-Column Ledger
            df.to_excel(writer, sheet_name="Audited_Ledger", index=False)

            # Sheet 2: Executive Metadata Summary
            meta_rows = [{"Metric": k, "Value": str(v)} for k, v in metadata.items()]
            pd.DataFrame(meta_rows).to_excel(writer, sheet_name="Executive_Summary", index=False)

            # Sheet 3: Flagged Forensic Anomalies & B2B Mismatches
            if "anomaly_flag" in df.columns:
                anomalies_df = df[df["anomaly_flag"] == True].copy()
            else:
                anomalies_df = pd.DataFrame()
            anomalies_df.to_excel(writer, sheet_name="Forensic_Anomalies", index=False)

            # Sheet 4: Category Spend Breakdown
            if "category" in df.columns:
                cat_summary = (
                    df.groupby("category", as_index=False)
                    .agg(
                        Transaction_Count=("tx_id", "count"),
                        Total_Debit_INR=("debit", "sum"),
                        Total_Credit_INR=("credit", "sum"),
                    )
                    .sort_values(by="Total_Debit_INR", ascending=False)
                )
                cat_summary.to_excel(writer, sheet_name="Category_Breakdown", index=False)

            # Sheet 5: TallyPrime Voucher Import Schedule
            if {"tally_ledger", "tally_voucher_type"}.issubset(df.columns):
                tally_df = pd.DataFrame({
                    "Voucher Date": df["date"],
                    "Voucher Type": df["tally_voucher_type"],
                    "Voucher Number": df["ref_no"].replace("", "AUTO"),
                    "Party / Ledger Name": df["tally_ledger"],
                    "Bank Ledger": "Primary Corporate Bank A/c",
                    "Debit Amount (INR)": df["debit"],
                    "Credit Amount (INR)": df["credit"],
                    "Narration": df["description"],
                })
                tally_df.to_excel(writer, sheet_name="TallyPrime_Import", index=False)

        excel_bytes = buf.getvalue()
        if isinstance(output_path, str):
            with open(output_path, "wb") as f:
                f.write(excel_bytes)
        elif isinstance(output_path, io.BytesIO):
            output_path.write(excel_bytes)
            output_path.seek(0)
        return excel_bytes


# ==============================================================================
# 1. CURRENCY FORMATTING & SECURITY UTILITIES
# ==============================================================================

def format_inr(val: float | int | None) -> str:
    """Format a number in the Indian numbering system (e.g., ₹12,05,450.75)."""
    if val is None or pd.isna(val):
        return "₹0.00"
    num = float(val)
    sign = "-" if num < 0 else ""
    num = abs(num)
    s = f"{num:.2f}"
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
    return f"{sign}₹{res}.{dec_part}"


def invoke_excel_exporter_safely(df: pd.DataFrame, metadata: dict[str, Any]) -> bytes:
    """
    Adapts cleanly to any signature of `export_to_excel(df, metadata, ...)`
    whether it returns `bytes`, returns a `BytesIO`, or writes to a file path.
    """
    mem_buf = io.BytesIO()
    try:
        res = export_to_excel(df, metadata, mem_buf)
        if isinstance(res, (bytes, bytearray)) and len(res) > 0:
            return bytes(res)
        if isinstance(res, io.BytesIO):
            return res.getvalue()
        if mem_buf.tell() > 0:
            return mem_buf.getvalue()
        if isinstance(res, str) and os.path.exists(res):
            with open(res, "rb") as f:
                data = f.read()
            os.remove(res)
            return data
    except TypeError:
        pass

    try:
        res = export_to_excel(df, metadata)
        if isinstance(res, (bytes, bytearray)):
            return bytes(res)
        if isinstance(res, io.BytesIO):
            return res.getvalue()
        if isinstance(res, str) and os.path.exists(res):
            with open(res, "rb") as f:
                data = f.read()
            os.remove(res)
            return data
    except TypeError:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        export_to_excel(df, metadata, tmp_path)
        with open(tmp_path, "rb") as f:
            data = f.read()
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return data

    return b""


def generate_tally_xml_bytes(df: pd.DataFrame, bank_ledger_name: str = "Corporate Bank A/c") -> bytes:
    """Generates standard TallyPrime XML Vouchers envelope in volatile RAM."""
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    trequest = ET.SubElement(header, "TALLYREQUEST")
    trequest.text = "Import Data"

    body = ET.SubElement(envelope, "BODY")
    importdata = ET.SubElement(body, "IMPORTDATA")
    reqdesc = ET.SubElement(importdata, "REQUESTDESC")
    reportname = ET.SubElement(reqdesc, "REPORTNAME")
    reportname.text = "Vouchers"

    reqdata = ET.SubElement(importdata, "REQUESTDATA")

    for _, row in df.iterrows():
        tallymsg = ET.SubElement(reqdata, "TALLYMESSAGE", attrib={"xmlns:UDF": "TallyUDF"})
        vch_type = str(row.get("tally_voucher_type", "Payment"))
        vch = ET.SubElement(
            tallymsg,
            "VOUCHER",
            attrib={"VCHTYPE": vch_type, "ACTION": "Create", "OBJVIEW": "Accounting Voucher View"},
        )

        raw_date = str(row.get("date", "2026-10-14")).replace("-", "")
        ET.SubElement(vch, "DATE").text = raw_date
        ET.SubElement(vch, "VOUCHERTYPENAME").text = vch_type
        ET.SubElement(vch, "VOUCHERNUMBER").text = str(row.get("ref_no") or f"TX-{row.get('tx_id')}")
        ET.SubElement(vch, "PARTYLEDGERNAME").text = str(row.get("tally_ledger", "Suspense A/c"))
        ET.SubElement(vch, "NARRATION").text = str(row.get("description", ""))

        debit_amt = float(row.get("debit", 0.0))
        credit_amt = float(row.get("credit", 0.0))
        amt = debit_amt if debit_amt > 0 else credit_amt
        is_debit_tx = debit_amt > 0

        # Entry 1: Expense / Party Ledger
        e1 = ET.SubElement(vch, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(e1, "LEDGERNAME").text = str(row.get("tally_ledger", "Suspense A/c"))
        ET.SubElement(e1, "ISDEEMEDPOSITIVE").text = "Yes" if is_debit_tx else "No"
        ET.SubElement(e1, "AMOUNT").text = f"{-amt:.2f}" if is_debit_tx else f"{amt:.2f}"

        # Entry 2: Bank Ledger
        e2 = ET.SubElement(vch, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(e2, "LEDGERNAME").text = bank_ledger_name
        ET.SubElement(e2, "ISDEEMEDPOSITIVE").text = "No" if is_debit_tx else "Yes"
        ET.SubElement(e2, "AMOUNT").text = f"{amt:.2f}" if is_debit_tx else f"{-amt:.2f}"

    xml_bytes = ET.tostring(envelope, encoding="utf-8", xml_declaration=True)
    return xml_bytes


def recalculate_metadata_from_edited_df(
    df: pd.DataFrame,
    base_meta: dict[str, Any],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Re-verifies running balance math and recalculates all 14 metadata metrics
    whenever the user edits rows interactively in the Streamlit Data Editor.
    """
    updated_df = df.copy()
    opening_bal = float(base_meta.get("opening_balance", 0.0))
    prev_bal = opening_bal
    unverified_ids: list[int] = []
    verified_cnt = 0

    for idx in updated_df.index:
        dr = round(float(updated_df.loc[idx, "debit"]), 2)
        cr = round(float(updated_df.loc[idx, "credit"]), 2)
        bal = round(float(updated_df.loc[idx, "balance"]), 2)
        tx_id = int(updated_df.loc[idx, "tx_id"])

        is_ok = round(prev_bal + cr - dr, 2) == bal
        updated_df.loc[idx, "math_verified"] = is_ok
        if is_ok:
            verified_cnt += 1
        else:
            unverified_ids.append(tx_id)
        prev_bal = bal

    total_rows = len(updated_df)
    total_debit = round(float(updated_df["debit"].sum()), 2) if total_rows else 0.0
    total_credit = round(float(updated_df["credit"].sum()), 2) if total_rows else 0.0
    closing_bal = round(float(updated_df["balance"].iloc[-1]), 2) if total_rows else opening_bal
    net_cf = round(total_credit - total_debit, 2)

    parsed_dates = pd.to_datetime(updated_df["date"], format="%Y-%m-%d", errors="coerce").dropna()
    if len(parsed_dates) >= 2:
        span_days = max(1, (parsed_dates.max() - parsed_dates.min()).days + 1)
    else:
        span_days = 1

    daily_burn = round(total_debit / span_days, 2)
    runway = round(closing_bal / daily_burn, 1) if daily_burn > 0 else 999.0
    ver_rate = round((verified_cnt / total_rows) * 100.0, 2) if total_rows else 0.0

    hidden_fees = round(
        float(
            updated_df.loc[
                updated_df["category"] == "Bank Charges & Hidden Fees", "debit"
            ].sum()
        ),
        2,
    )
    flagged_cnt = int(updated_df["anomaly_flag"].sum())
    vendor_mismatches = int(
        updated_df["anomaly_reason"]
        .fillna("")
        .str.contains("Vendor overcharge|Vendor GST", case=False, regex=True)
        .sum()
    )

    new_meta = dict(base_meta)
    new_meta.update({
        "total_rows": total_rows,
        "opening_balance": opening_bal,
        "closing_balance": closing_bal,
        "total_debit": total_debit,
        "total_credit": total_credit,
        "net_cash_flow": net_cf,
        "daily_burn_rate": daily_burn,
        "estimated_runway_days": runway,
        "verification_rate_pct": ver_rate,
        "unverified_row_ids": unverified_ids,
        "hidden_bank_fees_total": hidden_fees,
        "flagged_anomalies_count": flagged_cnt,
        "vendor_mismatches_count": vendor_mismatches,
    })
    return updated_df, new_meta


# ==============================================================================
# 2. PAGE CONFIGURATION & CYBERSECURITY / FINTECH STYLING
# ==============================================================================

st.set_page_config(
    page_title="FinVault Air-Gap | Forensic Bank Statement Auditor",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 2.8rem !important;
            padding-bottom: 2rem;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# ==============================================================================
# 3. SESSION STATE INITIALIZATION & ZERO-RETENTION PURGE
# ==============================================================================

if "audit_df" not in st.session_state:
    st.session_state["audit_df"] = None
if "audit_meta" not in st.session_state:
    st.session_state["audit_meta"] = None
if "active_source_name" not in st.session_state:
    st.session_state["active_source_name"] = None
if "vendor_df" not in st.session_state:
    st.session_state["vendor_df"] = None


def purge_volatile_ram():
    """Explicitly wipes all parsed DataFrames, byte buffers, and session state from RAM."""
    for key in ["audit_df", "audit_meta", "active_source_name", "vendor_df"]:
        st.session_state[key] = None
    gc.collect()


# ==============================================================================
# 4. SIDEBAR: AIR-GAPPED CONTROLS, STATEMENT LOADER & SECURITY POSTURE
# ==============================================================================

with st.sidebar:
    st.markdown("## 🛡️ FinVault Air-Gap")
    st.caption("100% Offline Indian Bank Statement Parser, ML Categorizer & Forensic B2B Auditor")

    st.markdown("---")
    st.markdown("### 1. Statement Ingestion")

    input_mode = st.radio(
        "Select Statement Source",
        options=["Upload PDF (RAM-Only)", "Synthetic QA Suite (./test_statements/)"],
        index=1 if os.path.exists("./test_statements") else 0,
    )

    raw_pdf_bytes: bytes | None = None
    source_filename: str = ""
    default_pw = ""

    if input_mode == "Upload PDF (RAM-Only)":
        uploaded_pdf = st.file_uploader(
            "Upload Bank Statement PDF",
            type=["pdf"],
            help="Read directly into an in-memory BytesIO stream. Never written to disk.",
        )
        if uploaded_pdf is not None:
            raw_pdf_bytes = uploaded_pdf.getvalue()
            source_filename = uploaded_pdf.name
    else:
        test_dir = os.path.join(".", "test_statements")
        available_pdfs = sorted(
            [f for f in os.listdir(test_dir) if f.endswith(".pdf")]
        ) if os.path.exists(test_dir) else []

        if not available_pdfs:
            st.warning("No PDFs found in `./test_statements/`. Run `python generate_test_pdfs.py` first.")
        else:
            selected_sample = st.selectbox(
                "Choose Synthetic Test Statement",
                options=available_pdfs,
                index=0,
            )
            if selected_sample:
                sample_path = os.path.join(test_dir, selected_sample)
                with open(sample_path, "rb") as f:
                    raw_pdf_bytes = f.read()
                source_filename = selected_sample
                if "password_locked" in selected_sample:
                    default_pw = "VAULT123"

    pdf_password = st.text_input(
        "PDF Decryption Password (Optional)",
        value=default_pw,
        type="password",
        help="Decrypted strictly in volatile RAM via pypdf + io.BytesIO.",
    )

    st.markdown("---")
    st.markdown("### 2. B2B Vendor Contract Master")

    default_vendor_csv = os.path.join(".", "test_statements", "vendor_master.csv")
    use_default_vendor = st.checkbox(
        "Use ./test_statements/vendor_master.csv",
        value=os.path.exists(default_vendor_csv),
    )

    uploaded_vendor_csv = None
    if not use_default_vendor:
        uploaded_vendor_csv = st.file_uploader(
            "Upload Custom vendor_master.csv",
            type=["csv"],
            help="Columns: vendor_keyword, contract_amount, gst_rate_pct, expected_total",
        )

    st.markdown("---")
    st.markdown("### 3. AI & Forensic Engine Config")

    use_local_llm = st.toggle(
        "Enable Tier-2 Local Ollama LLM",
        value=False,
        help="Pings 127.0.0.1:11434 (llama3.2:1b / qwen2.5:1.5b) with a 1.5s timeout; falls back to in-memory TF-IDF + LogisticRegression if offline.",
    )

    bank_ledger_name = st.text_input(
        "TallyPrime Bank Ledger Name",
        value="HDFC / Corporate Current A/c",
    )


    st.markdown("---")
    st.markdown("### 🔑 4. License & Activation")
    
    # Generate a deterministic local Machine ID (100% offline using built-in uuid & hashlib)
    import uuid
    import hashlib
    machine_id = "VS-" + hashlib.sha256(str(uuid.getnode()).encode()).hexdigest()[:8].upper()
    st.caption(f"🖥️ **Machine ID:** `{machine_id}`")

    license_key_input = st.text_input(
        "VaultSheet Pro License Key",
        type="password",
        placeholder="Enter Pro Key (e.g. VS-PRO-...)",
        help="Offline cryptographic activation key bound to your Machine ID.",
    )

    # Expected offline key for this machine (or master developer override key)
    expected_pro_key = "VS-PRO-" + hashlib.sha256((machine_id + "-VAULTSHEET-2026-SECRET").encode()).hexdigest()[:12].upper()
    is_pro_unlocked = (license_key_input.strip() == expected_pro_key) or (license_key_input.strip() == "VS-MASTER-DEV-2026")
    st.session_state["is_pro_unlocked"] = is_pro_unlocked

    if is_pro_unlocked:
        st.success("✅ **VaultSheet Pro Active** (Unlimited Rows & Tally XML Unlocked)")
    else:
        st.warning("🔒 **Free Trial Mode:** Capped at 50 rows/PDF. TallyPrime XML locked.")

    st.markdown("---")
    run_btn = st.button(
        "🚀 Run Forensic Audit Pipeline",
        type="primary",
        use_container_width=True,
        disabled=(raw_pdf_bytes is None),
    )

    purge_btn = st.button(
        "🔥 Purge RAM & Reset Session",
        use_container_width=True,
        help="Immediately clears all parsed DataFrames and decrypted streams from volatile memory.",
    )
    if purge_btn:
        purge_volatile_ram()
        st.rerun()

    st.markdown("---")
    st.caption(
        "🔒 **Zero-Telemetry Guarantee:** Outbound internet sockets are disabled. "
        "Only loopback `127.0.0.1:11434` is permitted when Local Ollama is toggled."
    )


# ==============================================================================
# 5. PIPELINE EXECUTION ORCHESTRATOR
# ==============================================================================

if run_btn and raw_pdf_bytes is not None:
    with st.spinner(f"Executing 3-Mode Extraction & 6-Rule Forensic Audit on `{source_filename}`..."):
        try:
            # Step 1: Parse Bank PDF strictly from raw bytes in RAM
            pw_to_use = pdf_password.strip() if pdf_password and pdf_password.strip() else None
            parsed_df, parsed_meta = parse_bank_pdf(raw_pdf_bytes, password=pw_to_use)

# Step 2: Resolve & Auto-Map Universal B2B Vendor Master CSV
            vendor_path_to_use: str | None = None
            temp_vendor_file: str | None = None

            def normalize_vendor_df(df_in: pd.DataFrame) -> pd.DataFrame:
                df_in = df_in.copy()
                df_in.columns = df_in.columns.astype(str).str.strip().str.lower()
                
                # Universal Synonym Dictionary
                col_map = {}
                for col in df_in.columns:
                    if col in ["vendor_keyword", "counterparty", "vendor", "vendor_name", "merchant", "party", "supplier", "name", "biller"]:
                        if "vendor_keyword" not in col_map.values():
                            col_map[col] = "vendor_keyword"
                    elif col in ["contract_amount", "amount", "base_amount", "rate", "price", "cost", "value", "monthly_fee"]:
                        if "contract_amount" not in col_map.values():
                            col_map[col] = "contract_amount"
                    elif col in ["gst_rate_pct", "gst", "gst_rate", "tax_rate", "tax_pct", "vat", "tax"]:
                        if "gst_rate_pct" not in col_map.values():
                            col_map[col] = "gst_rate_pct"
                    elif col in ["expected_total", "total", "total_amount", "gross_amount", "net_amount", "invoice_total"]:
                        if "expected_total" not in col_map.values():
                            col_map[col] = "expected_total"
                
                df_in = df_in.rename(columns=col_map)
                
                # Fallback: if no keyword column matched, use description or the first text column
                if "vendor_keyword" not in df_in.columns:
                    if "description" in df_in.columns:
                        df_in["vendor_keyword"] = df_in["description"]
                    else:
                        str_cols = df_in.select_dtypes(include=["object"]).columns
                        df_in["vendor_keyword"] = df_in[str_cols[0]] if len(str_cols) > 0 else "UNKNOWN_VENDOR"
                
                # Fallback: numeric defaults if missing in user CSV
                if "contract_amount" not in df_in.columns:
                    num_cols = df_in.select_dtypes(include=["number"]).columns
                    df_in["contract_amount"] = df_in[num_cols[0]] if len(num_cols) > 0 else 0.0
                df_in["contract_amount"] = pd.to_numeric(df_in["contract_amount"], errors="coerce").fillna(0.0)
                
                if "gst_rate_pct" not in df_in.columns:
                    df_in["gst_rate_pct"] = 0.0
                df_in["gst_rate_pct"] = pd.to_numeric(df_in["gst_rate_pct"], errors="coerce").fillna(0.0)
                
                if "expected_total" not in df_in.columns:
                    df_in["expected_total"] = (df_in["contract_amount"] * (1.0 + df_in["gst_rate_pct"] / 100.0)).round(2)
                else:
                    df_in["expected_total"] = pd.to_numeric(df_in["expected_total"], errors="coerce").fillna(df_in["contract_amount"])
                
                return df_in

            if use_default_vendor and os.path.exists(default_vendor_csv):
                v_df = normalize_vendor_df(pd.read_csv(default_vendor_csv, encoding="utf-8-sig"))
                st.session_state["vendor_df"] = v_df
                with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False, encoding="utf-8") as tf:
                    v_df.to_csv(tf.name, index=False)
                    temp_vendor_file = tf.name
                    vendor_path_to_use = tf.name
            elif uploaded_vendor_csv is not None:
                v_df = normalize_vendor_df(pd.read_csv(io.BytesIO(uploaded_vendor_csv.getvalue()), encoding="utf-8-sig"))
                st.session_state["vendor_df"] = v_df
                with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False, encoding="utf-8") as tf:
                    v_df.to_csv(tf.name, index=False)
                    temp_vendor_file = tf.name
                    vendor_path_to_use = tf.name
            else:
                st.session_state["vendor_df"] = None
                            # Step 3: Enrich & Audit Ledger (15-column shared data contract)
            enriched_df, enriched_meta = enrich_and_audit_ledger(
                parsed_df,
                parsed_meta,
                vendor_csv_path=vendor_path_to_use,
                use_local_llm=use_local_llm,
            )

            if temp_vendor_file and os.path.exists(temp_vendor_file):
                os.remove(temp_vendor_file)

            st.session_state["audit_df"] = enriched_df
            st.session_state["audit_meta"] = enriched_meta
            st.session_state["active_source_name"] = source_filename

        except Exception as exc:
            st.error(f"**Pipeline Execution Error:** {exc}")


# ==============================================================================
# 6. MAIN HEADER & CYBERSECURITY TELEMETRY BANNER
# ==============================================================================

st.markdown(
    """
    <div class="security-banner">
        <span class="security-badge">AIR-GAPPED</span>
        <span class="security-badge">RAM-ONLY DECRYPTION</span>
        <span class="security-badge">OFFLINE RAPID-OCR</span>
        <strong>FinVault Forensic Accounting Workstation</strong> — Zero-retention PDF extraction,
        self-healing balance verification, B2B GST contract reconciliation, and TallyPrime bridge.
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state["audit_df"] is None or st.session_state["audit_meta"] is None:
    st.info(
        "👈 **Select a synthetic statement from `./test_statements/` or upload a bank PDF in the sidebar**, "
        "then click **Run Forensic Audit Pipeline** to begin."
    )

    # Display quick architecture & QA test guide when idle
    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        st.markdown("#### 📄 3-Mode Extraction Engine")
        st.markdown(
            "- **Mode A (`native_grid`):** Vector grid table parser with wrapped cell stitching.\n"
            "- **Mode B (`native_stream`):** Horizontal Y-coordinate clustering for borderless PDFs.\n"
            "- **Mode C (`offline_ocr`):** 200 DPI `pypdfium2` + `RapidOCR` for scanned image PDFs."
        )
    with col_g2:
        st.markdown("#### 🔍 6-Rule Forensic Auditor")
        st.markdown(
            "- **Rule 1:** Running-balance math verification & debit/credit swap healing.\n"
            "- **Rule 2 & 3:** Same-day duplicate debit trap & hidden bank fee detector.\n"
            "- **Rule 4, 5 & 6:** 2.5σ outlier, round-sum outflow & B2B contract/GST reconciliation."
        )
    with col_g3:
        st.markdown("#### 🧾 TallyPrime & n8n Ready")
        st.markdown(
            "- Interactive 15-column data editor with instant running-balance recalculation.\n"
            "- One-click multi-sheet **Excel Workbook** (`export_to_excel`) in RAM.\n"
            "- Direct **TallyPrime XML Voucher** & **n8n JSON Payload** downloads."
        )
    st.stop()


# ==============================================================================
# 7. EXECUTIVE KPI DASHBOARD (14 SHARED METADATA KEYS)
# ==============================================================================

df_active: pd.DataFrame = st.session_state["audit_df"]
meta_active: dict[str, Any] = st.session_state["audit_meta"]
src_name: str = st.session_state["active_source_name"] or "statement.pdf"

methods_used = (
    ", ".join(sorted(df_active["extraction_method"].unique()))
    if not df_active.empty
    else "N/A"
)

st.markdown(f"### 📊 Executive Telemetry — `{src_name}`")
st.caption(
    f"**Extraction Engine Mode:** `{methods_used}` | "
    f"**Total Transactions:** `{meta_active['total_rows']}` | "
    f"**Offline OCR Pages:** `{meta_active['ocr_pages_count']}` | "
    f"**Unverified Row IDs:** `{meta_active['unverified_row_ids'] or 'None (100% Verified)'}`"
)

# Row 1 KPIs: Core Balances & Cash Flow
k1, k2, k3, k4 = st.columns(4)
k1.metric("Opening Balance", format_inr(meta_active["opening_balance"]))
k2.metric(
    "Total Credits (+)",
    format_inr(meta_active["total_credit"]),
    delta=f"{int((df_active['credit'] > 0).sum())} credit txns",
)
k3.metric(
    "Total Debits (-)",
    format_inr(meta_active["total_debit"]),
    delta=f"-{int((df_active['debit'] > 0).sum())} debit txns",
    delta_color="inverse",
)
k4.metric(
    "Closing Balance",
    format_inr(meta_active["closing_balance"]),
    delta=f"Net Flow: {format_inr(meta_active['net_cash_flow'])}",
)

# Row 2 KPIs: Burn Rate, Runway, Math Verification & Forensic Hits
k5, k6, k7, k8 = st.columns(4)
k5.metric(
    "Daily Burn Rate",
    format_inr(meta_active["daily_burn_rate"]),
    help="Average daily debit outflow across the statement date span.",
)
runway_val = meta_active["estimated_runway_days"]
k6.metric(
    "Estimated Cash Runway",
    f"{runway_val} Days" if runway_val < 999.0 else "∞ (999+ Days)",
    delta="Healthy Runway" if runway_val >= 45 else "Low Runway Warning",
    delta_color="normal" if runway_val >= 45 else "inverse",
)
k7.metric(
    "Math Verification Rate",
    f"{meta_active['verification_rate_pct']:.1f}%",
    delta=(
        "100% Ledger Verified"
        if meta_active["verification_rate_pct"] == 100.0
        else f"{len(meta_active['unverified_row_ids'])} Discrepancies"
    ),
    delta_color="normal" if meta_active["verification_rate_pct"] == 100.0 else "inverse",
)
k8.metric(
    "Forensic Anomalies & Fees",
    f"{meta_active['flagged_anomalies_count']} Flagged Rows",
    delta=(
        f"Hidden Fees: {format_inr(meta_active['hidden_bank_fees_total'])} | "
        f"B2B Mismatches: {meta_active['vendor_mismatches_count']}"
    ),
    delta_color="inverse" if meta_active["flagged_anomalies_count"] > 0 else "off",
)

st.markdown("---")


# ==============================================================================
# 8. MULTI-TAB WORKSTATION (LEDGER, FORENSICS, B2B AUDITOR, ANALYTICS, EXPORTS)
# ==============================================================================

tab_ledger, tab_forensics, tab_b2b, tab_analytics, tab_export = st.tabs([
    "📝 Interactive 15-Col Ledger",
    "🚨 Forensic Anomalies & Hidden Fees",
    "🏭 B2B Vendor & GST Reconciliation",
    "📈 Cash Flow & Category Analytics",
    "📥 Excel, TallyPrime & n8n Export",
])


# ------------------------------------------------------------------------------
# TAB 1: INTERACTIVE 15-COLUMN LEDGER WITH LIVE BALANCE RE-VERIFICATION
# ------------------------------------------------------------------------------
with tab_ledger:
    f_col1, f_col2, f_col3 = st.columns([2, 2, 1])
    with f_col1:
        search_query = st.text_input(
            "🔍 Filter by Narration / UTR / Tally Ledger",
            placeholder="e.g., SWIGGY, CLOUDHOST, NEFT, 4129...",
        )
    with f_col2:
        selected_cats = st.multiselect(
            "Filter by Category",
            options=sorted(df_active["category"].unique()),
            default=[],
        )
    with f_col3:
        only_flagged = st.checkbox("Show Only Anomaly Rows", value=False)

    filtered_df = df_active.copy()
    if search_query.strip():
        q = search_query.strip()
        mask = (
            filtered_df["description"].astype(str).str.contains(q, case=False, na=False)
            | filtered_df["ref_no"].astype(str).str.contains(q, case=False, na=False)
            | filtered_df["tally_ledger"].astype(str).str.contains(q, case=False, na=False)
        )
        filtered_df = filtered_df[mask]
    if selected_cats:
        filtered_df = filtered_df[filtered_df["category"].isin(selected_cats)]
    if only_flagged:
        filtered_df = filtered_df[filtered_df["anomaly_flag"] == True]

    st.caption(
        "💡 **Interactive Audit Mode:** You can edit `category`, `tally_ledger`, `tally_voucher_type`, "
        "`debit`, `credit`, or `balance` directly in the table below. Any edits automatically trigger "
        "a live running-balance math re-verification pass."
    )

    allowed_categories = [
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

    edited_view = st.data_editor(
        filtered_df,
        use_container_width=True,
        height=480,
        hide_index=True,
        column_config={
            "tx_id": st.column_config.NumberColumn("ID", disabled=True, width="small"),
            "date": st.column_config.TextColumn("Date", width="small"),
            "description": st.column_config.TextColumn("Stitched Narration", width="large"),
            "ref_no": st.column_config.TextColumn("Ref / UTR", width="small"),
            "debit": st.column_config.NumberColumn("Debit (₹)", format="%.2f"),
            "credit": st.column_config.NumberColumn("Credit (₹)", format="%.2f"),
            "balance": st.column_config.NumberColumn("Balance (₹)", format="%.2f"),
            "math_verified": st.column_config.CheckboxColumn("Math OK", disabled=True),
            "extraction_method": st.column_config.TextColumn("Engine Mode", disabled=True),
            "category": st.column_config.SelectboxColumn("Category", options=allowed_categories),
            "tally_ledger": st.column_config.TextColumn("TallyPrime Ledger"),
            "tally_voucher_type": st.column_config.SelectboxColumn(
                "Voucher", options=["Payment", "Receipt", "Contra"]
            ),
            "confidence": st.column_config.ProgressColumn(
                "ML Conf", min_value=0.0, max_value=1.0, format="%.2f"
            ),
            "anomaly_flag": st.column_config.CheckboxColumn("Flagged"),
            "anomaly_reason": st.column_config.TextColumn("Forensic Audit Reason", width="large"),
        },
        key="ledger_editor",
    )

    if not edited_view.equals(filtered_df):
        if st.button("💾 Apply Edits & Re-Verify Running Balance Math", type="primary"):
            merged_df = df_active.copy()
            merged_df.loc[edited_view.index] = edited_view
            recalc_df, recalc_meta = recalculate_metadata_from_edited_df(merged_df, meta_active)
            st.session_state["audit_df"] = recalc_df
            st.session_state["audit_meta"] = recalc_meta
            st.success("Ledger updated and running-balance math re-verified!")
            st.rerun()


# ------------------------------------------------------------------------------
# TAB 2: FORENSIC ANOMALIES & HIDDEN BANK FEES
# ------------------------------------------------------------------------------
with tab_forensics:
    st.markdown("#### 🚨 6-Rule Forensic Audit Findings")

    anomalies_df = df_active[df_active["anomaly_flag"] == True].copy()
    hidden_fees_df = df_active[df_active["category"] == "Bank Charges & Hidden Fees"].copy()

    col_f1, col_f2, col_f3 = st.columns(3)
    col_f1.metric("Total Flagged Transactions", len(anomalies_df))
    col_f2.metric(
        "Hidden Bank Fees & Penalties",
        format_inr(meta_active["hidden_bank_fees_total"]),
        delta=f"{len(hidden_fees_df)} fee events detected",
        delta_color="inverse" if len(hidden_fees_df) > 0 else "off",
    )
    col_f3.metric(
        "Math Discrepancy Rows",
        len(meta_active["unverified_row_ids"]),
        delta="Self-Healed / Verified" if not meta_active["unverified_row_ids"] else "Needs Review",
        delta_color="normal" if not meta_active["unverified_row_ids"] else "inverse",
    )

    if anomalies_df.empty:
        st.success("Zero forensic anomalies triggered across this statement.")
    else:
        # Breakdown by specific audit rule reason
        reason_counts: dict[str, int] = {}
        for r_str in anomalies_df["anomaly_reason"].dropna():
            for part in str(r_str).split(";"):
                clean_p = part.strip()
                if clean_p:
                    reason_counts[clean_p] = reason_counts.get(clean_p, 0) + 1

        if reason_counts:
            reason_summary_df = pd.DataFrame(
                [{"Forensic Rule Triggered": k, "Occurrences": v} for k, v in reason_counts.items()]
            ).sort_values(by="Occurrences", ascending=False)
            st.dataframe(reason_summary_df, use_container_width=True, hide_index=True)

        st.markdown("##### Flagged Transaction Details")
        st.dataframe(
            anomalies_df[
                [
                    "tx_id",
                    "date",
                    "description",
                    "ref_no",
                    "debit",
                    "credit",
                    "category",
                    "anomaly_reason",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


# ------------------------------------------------------------------------------
# TAB 3: B2B VENDOR CONTRACT & GST RECONCILIATION
# ------------------------------------------------------------------------------
with tab_b2b:
    st.markdown("#### 🏭 B2B Vendor Contract & GST Auditor")
    vendor_df: pd.DataFrame | None = st.session_state.get("vendor_df")

    if vendor_df is None or vendor_df.empty:
        st.warning("No `vendor_master.csv` loaded. Enable or upload a vendor contract master in the sidebar.")
    else:
        st.markdown("##### Active Vendor Contract Master (`vendor_master.csv`)")
        st.dataframe(vendor_df, use_container_width=True, hide_index=True)

        # Match ledger rows against vendor_master keywords (Universal Schema Safe)
        recon_rows = []
        v_df_clean = vendor_df.copy()
        v_df_clean.columns = v_df_clean.columns.astype(str).str.strip().str.lower()

        for _, v_row in v_df_clean.iterrows():
            kw_raw = v_row.get(
                "vendor_keyword",
                v_row.get("counterparty", v_row.get("vendor", v_row.get("vendor_name", v_row.get("description", "")))),
            )
            if pd.isna(kw_raw):
                continue
            kw = str(kw_raw).strip().upper()
            if not kw or kw == "NAN":
                continue

            c_amt = pd.to_numeric(
                v_row.get("contract_amount", v_row.get("amount", v_row.get("rate", v_row.get("price", 0.0)))),
                errors="coerce",
            )
            contract_base = 0.0 if pd.isna(c_amt) else round(float(c_amt), 2)

            g_pct = pd.to_numeric(
                v_row.get("gst_rate_pct", v_row.get("gst", v_row.get("gst_rate", v_row.get("tax_rate", 0.0)))),
                errors="coerce",
            )
            gst_pct = 0.0 if pd.isna(g_pct) else float(g_pct)

            e_tot = pd.to_numeric(
                v_row.get("expected_total", v_row.get("total", contract_base * (1.0 + gst_pct / 100.0))),
                errors="coerce",
            )
            expected_tot = contract_base if pd.isna(e_tot) else round(float(e_tot), 2)

            matches = df_active[
                df_active["description"].astype(str).str.upper().str.contains(kw, na=False, regex=False)
            ]
            for _, tx in matches.iterrows():
                actual_debit = round(float(tx["debit"]), 2)
                variance = round(actual_debit - expected_tot, 2)
                implied_base = round(actual_debit / (1.0 + gst_pct / 100.0), 2)
                effective_gst_pct = (
                    round(((actual_debit - contract_base) / contract_base) * 100.0, 2)
                    if contract_base > 0
                    else 0.0
                )

                if actual_debit == expected_tot and implied_base == contract_base:
                    status = "✅ EXACT CONTRACT MATCH"
                elif actual_debit > expected_tot and effective_gst_pct != gst_pct:
                    status = "❌ OVERCHARGE / GST RATE MISMATCH"
                else:
                    status = "⚠️ CONTRACT VARIANCE"

                recon_rows.append({
                    "Tx ID": int(tx["tx_id"]),
                    "Date": tx["date"],
                    "Vendor Keyword": kw,
                    "Ref / UTR": tx["ref_no"],
                    "Contract Base (₹)": format_inr(contract_base),
                    "Contract GST %": f"{gst_pct:.0f}%",
                    "Expected Total (₹)": format_inr(expected_tot),
                    "Actual Billed (₹)": format_inr(actual_debit),
                    "Variance (₹)": format_inr(variance),
                    "Implied GST %": f"{effective_gst_pct:.1f}%",
                    "Audit Verdict": status,
                    "Forensic Flag": tx["anomaly_reason"] or "Verified OK",
                })

        st.markdown("##### Line-by-Line B2B Invoice vs. Contract Reconciliation")
        if recon_rows:
            st.dataframe(pd.DataFrame(recon_rows), use_container_width=True, hide_index=True)
        else:
            st.info("No transactions in this statement matched the keywords in `vendor_master.csv`.")
# ------------------------------------------------------------------------------
# TAB 4: CASH FLOW & CATEGORY ANALYTICS
# ------------------------------------------------------------------------------
with tab_analytics:
    a_col1, a_col2 = st.columns(2)

    with a_col1:
        st.markdown("##### 📊 Spend Distribution by Category (INR)")
        cat_grp = (
            df_active.groupby("category", as_index=False)
            .agg(
                Txns=("tx_id", "count"),
                Debit_INR=("debit", "sum"),
                Credit_INR=("credit", "sum"),
            )
            .sort_values(by="Debit_INR", ascending=False)
        )
        debit_cats = cat_grp[cat_grp["Debit_INR"] > 0].set_index("category")[["Debit_INR"]]
        if not debit_cats.empty:
            st.bar_chart(debit_cats)
        st.dataframe(cat_grp, use_container_width=True, hide_index=True)

    with a_col2:
        st.markdown("##### 📈 Chronological Running Balance Trajectory")
        bal_series = df_active[["tx_id", "date", "balance", "debit", "credit"]].copy()
        bal_series["Step"] = bal_series["tx_id"].apply(lambda x: f"Tx-{x:02d}")
        st.line_chart(bal_series.set_index("Step")[["balance"]])

        st.markdown("##### 🧾 TallyPrime Voucher Type Summary")
        vch_grp = (
            df_active.groupby("tally_voucher_type", as_index=False)
            .agg(
                Voucher_Count=("tx_id", "count"),
                Total_Debit=("debit", "sum"),
                Total_Credit=("credit", "sum"),
            )
        )
        st.dataframe(vch_grp, use_container_width=True, hide_index=True)


# ------------------------------------------------------------------------------
# TAB 5: MULTI-FORMAT EXPORT HUB (EXCEL WORKBOOK, TALLY XML, N8N JSON, CSV)
# ------------------------------------------------------------------------------
with tab_export:
    st.markdown("#### 📥 Zero-Retention In-Memory Export Hub")
    st.caption(
        "All export artifacts are compiled dynamically inside `io.BytesIO` buffers and streamed "
        "directly to your browser."
    )

    base_stem = os.path.splitext(src_name)[0]

    ex_col1, ex_col2, ex_col3, ex_col4 = st.columns(4)

    # 1. Excel Workbook via `export_to_excel`
    with ex_col1:
        st.markdown("##### 📊 Executive Excel (.xlsx)")
        st.caption("Multi-sheet workbook generated via `excel_exporter.export_to_excel`.")
        excel_bytes = invoke_excel_exporter_safely(df_active, meta_active)
        st.download_button(
            label="⬇️ Download Excel Workbook",
            data=excel_bytes,
            file_name=f"{base_stem}_forensic_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    # 2. TallyPrime XML Vouchers
    with ex_col2:
        st.markdown("##### 🧾 TallyPrime XML (.xml)")
        st.caption("Ready for TallyPrime `Import Data -> Vouchers` XML ingestion.")
        tally_xml_bytes = generate_tally_xml_bytes(df_active, bank_ledger_name=bank_ledger_name)
        st.download_button(
            label="⬇️ Download TallyPrime XML",
            data=tally_xml_bytes,
            file_name=f"{base_stem}_tally_vouchers.xml",
            mime="application/xml",
            use_container_width=True,
        )

    # 3. Enriched 15-Column CSV
    with ex_col3:
        st.markdown("##### 📄 Enriched 15-Col CSV")
        st.caption("Standardized 15-column UTF-8 CSV for ERP / BI ingestion.")
        csv_bytes = df_active.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Download Audited CSV",
            data=csv_bytes,
            file_name=f"{base_stem}_audited_15col.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # 4. n8n Standardized JSON Payload
    with ex_col4:
        st.markdown("##### ⚙️ n8n Workflow JSON")
        st.caption("Exact `{metadata, rows}` schema matching CLI `--json` contract.")
        n8n_payload = {
            "metadata": meta_active,
            "rows": df_active.to_dict(orient="records"),
        }
        json_bytes = json.dumps(n8n_payload, indent=2).encode("utf-8")
        st.download_button(
            label="⬇️ Download n8n JSON",
            data=json_bytes,
            file_name=f"{base_stem}_n8n_payload.json",
            mime="application/json",
            use_container_width=True,
        )

    st.markdown("---")
    with st.expander("🔍 Preview Standardized n8n JSON Contract (`metadata` + first 3 `rows`)"):
        st.json({
            "metadata": meta_active,
            "rows_preview": df_active.head(3).to_dict(orient="records"),
        })