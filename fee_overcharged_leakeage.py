import streamlit as st
import pandas as pd
import io

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Fees Overcharged Leakage",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Fees Overcharged Leakage")
st.markdown(
    "Upload your **Unified Transaction** report and **Fee Estimate (FREE)** CSV "
    "to detect commission and weight-handling overcharges."
)

# ──────────────────────────────────────────────
# Helper: numeric cleaner
# ──────────────────────────────────────────────
def to_numeric(series: pd.Series) -> pd.Series:
    return (
        pd.to_numeric(
            series.astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce",
        ).fillna(0)
    )


# ──────────────────────────────────────────────
# Step 1 – File upload
# ──────────────────────────────────────────────
st.header("Step 1 — Upload Files")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Unified Transaction Report")
    txn_file = st.file_uploader(
        "CSV exported from Seller Central (header starts at row 14)",
        type=["csv"],
        key="txn",
    )

with col2:
    st.subheader("📄 Fee Estimate Report (FREE)")
    free_file = st.file_uploader(
        "SKU-level fee estimate CSV from Amazon",
        type=["csv"],
        key="free",
    )

if not txn_file or not free_file:
    st.info("⬆️ Please upload both files to continue.")
    st.stop()

# ──────────────────────────────────────────────
# Step 2 – Load & preview Transaction data
# ──────────────────────────────────────────────
st.header("Step 2 — Transaction Data")

with st.spinner("Reading transaction file…"):
    try:
        transaction_raw = pd.read_csv(txn_file, header=13)
    except Exception as e:
        st.error(f"Could not read transaction file: {e}")
        st.stop()

with st.expander("Raw transaction preview (first 5 rows)", expanded=False):
    st.dataframe(transaction_raw.head(), use_container_width=True)

# Clean numeric columns
numeric_cols = [
    "Total sales tax liable(GST before adjusting TCS)",
    "other transaction fees",
    "product sales",
    "selling fees",
]

missing = [c for c in numeric_cols if c not in transaction_raw.columns]
if missing:
    st.error(f"Transaction file is missing columns: {missing}")
    st.stop()

for col in numeric_cols:
    transaction_raw[col] = to_numeric(transaction_raw[col])

# Filter: only Order type, non-zero product sales
transaction = transaction_raw[
    transaction_raw["type"].astype(str).str.strip() == "Order"
].copy()
transaction = transaction[transaction["product sales"] != 0].reset_index(drop=True)

st.success(f"✅ Loaded **{len(transaction):,}** valid Order rows")

# ──────────────────────────────────────────────
# Step 3 – Load & preview FREE data
# ──────────────────────────────────────────────
st.header("Step 3 — Fee Estimate Data")

with st.spinner("Reading fee estimate file…"):
    try:
        free = pd.read_csv(free_file)
    except Exception as e:
        st.error(f"Could not read fee estimate file: {e}")
        st.stop()

with st.expander("Raw fee estimate preview (first 5 rows)", expanded=False):
    st.dataframe(free.head(), use_container_width=True)

# --- Compute derived fee columns ---
for col in ["sales-price", "estimated-referral-fee-per-unit"]:
    free[col] = to_numeric(free[col])

free["Commission %"] = (
    free["estimated-referral-fee-per-unit"] / free["sales-price"] * 100
).round(2)

free["Commission With GST"] = (
    free["estimated-referral-fee-per-unit"] * 1.18
).round(2)

free["estimated-fixed-closing-fee"] = to_numeric(free["estimated-fixed-closing-fee"])
free["Closing With GST"] = (free["estimated-fixed-closing-fee"] * 1.18).round(2)

for col in [
    "estimated-pick-pack-fee-per-unit",
    "estimated-weight-handling-fee-per-unit",
]:
    free[col] = to_numeric(free[col])

free["Pick& Pack With GST"] = (free["estimated-pick-pack-fee-per-unit"] * 1.18).round(2)
free["Weight GST"] = (free["estimated-weight-handling-fee-per-unit"] * 1.18).round(2)
free["Total Weight Handling"] = (free["Pick& Pack With GST"] + free["Weight GST"]).round(2)
free["Total Weight Handling With Gst"] = (free["Total Weight Handling"] * 1.18).round(2)

st.success(f"✅ Loaded **{len(free):,}** SKU rows from fee estimate file")

# ──────────────────────────────────────────────
# Step 4 – Build pivot table
# ──────────────────────────────────────────────
st.header("Step 4 — Pivot Table (Order × SKU)")

required_pivot_cols = [
    "order id", "Sku", "quantity", "product sales",
    "Total sales tax liable(GST before adjusting TCS)",
    "selling fees", "fba fees", "other transaction fees",
]
missing_pivot = [c for c in required_pivot_cols if c not in transaction.columns]
if missing_pivot:
    st.error(f"Transaction file missing columns for pivot: {missing_pivot}")
    st.stop()

pivot_table = pd.pivot_table(
    transaction,
    index=["order id", "Sku"],
    values=[
        "quantity",
        "product sales",
        "Total sales tax liable(GST before adjusting TCS)",
        "selling fees",
        "fba fees",
        "other transaction fees",
    ],
    aggfunc="sum",
    fill_value=0,
).reset_index()

pivot_table["Sales Amount"] = (
    pivot_table["product sales"]
    + pivot_table["Total sales tax liable(GST before adjusting TCS)"]
)

# Merge commission & weight lookups from free
pivot_table["Sku"] = pivot_table["Sku"].astype(str).str.strip()
free["sku"] = free["sku"].astype(str).str.strip()

commission_lookup = free.set_index("sku")["Commission %"].to_dict()
pivot_table["Commission %"] = pivot_table["Sku"].map(commission_lookup).fillna(0)

# Commission calculations
pivot_table["Commission Base"] = (
    pivot_table["Sales Amount"] * pivot_table["Commission %"] / 100
).round(2)

pivot_table["Commission With GST"] = (pivot_table["Commission Base"] * 1.18).round(2)

pivot_table["With GST Commission Qty"] = (
    pivot_table["Commission With GST"] * pivot_table["quantity"]
).round(2)

pivot_table["Diff With Charge & Present Commission"] = (
    pivot_table["selling fees"] + pivot_table["With GST Commission Qty"]
).round(2)

# Weight handling
weight_lookup = free.set_index("sku")["Total Weight Handling With Gst"].to_dict()
pivot_table["Weight Handling With GST"] = pivot_table["Sku"].map(weight_lookup)

pivot_table["Weight Handling With GST Qty"] = (
    pivot_table["Weight Handling With GST"] * pivot_table["quantity"]
).round(2)

pivot_table["Weight Handling Different"] = (
    pivot_table["fba fees"] + pivot_table["Weight Handling With GST Qty"]
).round(2)

st.subheader(f"📋 Full Pivot Table — {len(pivot_table):,} rows")
st.dataframe(pivot_table, use_container_width=True, height=400)

pivot_buf = io.BytesIO()
pivot_table.to_excel(pivot_buf, index=False, engine="openpyxl")
st.download_button(
    "⬇️ Download Full Pivot Table",
    data=pivot_buf.getvalue(),
    file_name="pivot_table_full.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ──────────────────────────────────────────────
# Step 5 – Overcharge Reports
# ──────────────────────────────────────────────
st.header("Step 5 — Overcharge Reports")

tab1, tab2 = st.tabs(
    ["💰 Commission Overcharge", "⚖️ Weight Handling Overcharge"]
)

# ── Commission Overcharge ──
with tab1:
    negative_diff_report = (
        pivot_table[pivot_table["Diff With Charge & Present Commission"] < 0]
        .copy()
        .sort_values("Diff With Charge & Present Commission")
        .reset_index(drop=True)
    )

    total_commission_leakage = negative_diff_report[
        "Diff With Charge & Present Commission"
    ].sum()

    m1, m2 = st.columns(2)
    m1.metric("Overcharged Orders", f"{len(negative_diff_report):,}")
    m2.metric(
        "Total Commission Leakage",
        f"₹{abs(total_commission_leakage):,.2f}",
        delta=f"-₹{abs(total_commission_leakage):,.2f}",
        delta_color="inverse",
    )

    if negative_diff_report.empty:
        st.success("🎉 No commission overcharges detected!")
    else:
        st.dataframe(negative_diff_report, use_container_width=True)

        buf = io.BytesIO()
        negative_diff_report.to_excel(buf, index=False, engine="openpyxl")
        st.download_button(
            "⬇️ Download Commission Overcharge Report",
            data=buf.getvalue(),
            file_name="commission_overcharge_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── Weight Handling Overcharge ──
with tab2:
    negative_weight_report = (
        pivot_table[pivot_table["Weight Handling Different"] < 0]
        .copy()
        .sort_values("Weight Handling Different")
        .reset_index(drop=True)
    )

    total_weight_leakage = negative_weight_report["Weight Handling Different"].sum()

    m3, m4 = st.columns(2)
    m3.metric("Overcharged Orders", f"{len(negative_weight_report):,}")
    m4.metric(
        "Total Weight Leakage",
        f"₹{abs(total_weight_leakage):,.2f}",
        delta=f"-₹{abs(total_weight_leakage):,.2f}",
        delta_color="inverse",
    )

    if negative_weight_report.empty:
        st.success("🎉 No weight handling overcharges detected!")
    else:
        st.dataframe(negative_weight_report, use_container_width=True)

        buf2 = io.BytesIO()
        negative_weight_report.to_excel(buf2, index=False, engine="openpyxl")
        st.download_button(
            "⬇️ Download Weight Handling Overcharge Report",
            data=buf2.getvalue(),
            file_name="weight_overcharge_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ──────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────
st.divider()
st.caption(
    "Free Overcharged Leakage • Amazon Seller Analytics • "
    "Commission & FBA Fee Discrepancy Detector"
)