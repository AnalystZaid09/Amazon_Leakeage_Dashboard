# Helper to get cached ASIN-to-Brand mapping from PM.xlsx
@st.cache_data
def get_pm_brand_map():
    import pandas as pd
    import os
    if not os.path.exists("PM.xlsx"):
        return {}
    try:
        df_header = pd.read_excel("PM.xlsx", nrows=1)
        cols = list(df_header.columns)
        
        asin_col = None
        for c in cols:
            if str(c).strip().lower() == "asin":
                asin_col = c
                break
        if not asin_col:
            for c in cols:
                if "asin" in str(c).lower():
                    asin_col = c
                    break
                    
        brand_col = None
        for c in cols:
            if str(c).strip().lower() == "brand":
                brand_col = c
                break
        if not brand_col:
            for c in cols:
                if "brand" in str(c).lower():
                    brand_col = c
                    break
                    
        if asin_col and brand_col:
            df_pm = pd.read_excel("PM.xlsx", usecols=[asin_col, brand_col])
            df_pm[asin_col] = df_pm[asin_col].astype(str).str.strip().str.upper()
            df_pm[brand_col] = df_pm[brand_col].astype(str).str.strip()
            brand_map = df_pm.dropna(subset=[asin_col]).drop_duplicates(subset=asin_col).set_index(asin_col)[brand_col].to_dict()
            return brand_map
    except Exception as e:
        print(f"Error loading PM.xlsx: {e}")
    return {}

# Helper to insert mapped Brand column at the starting of a DataFrame
def map_brand_column_from_pm(df):
    import pandas as pd
    if df is None or len(df) == 0:
        return df
    try:
        asin_target = None
        for c in df.columns:
            if str(c).strip().lower() == "asin":
                asin_target = c
                break
        if not asin_target:
            for c in df.columns:
                if "asin" in str(c).lower():
                    asin_target = c
                    break
                    
        if asin_target:
            brand_map = get_pm_brand_map()
            if brand_map:
                asins_clean = df[asin_target].astype(str).str.strip().str.upper()
                mapped_brands = asins_clean.map(brand_map)
                
                # Drop existing brand columns (case-insensitive) to avoid duplication
                brand_cols = [c for c in df.columns if str(c).strip().lower() == "brand"]
                if brand_cols:
                    df = df.drop(columns=brand_cols)
                    
                df.insert(0, "Brand", mapped_brands)
    except Exception as e:
        print(f"Error mapping brand: {e}")
    return df

st.set_page_config(page_title="Amazon Refund Cross Check", page_icon="📊", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f2937;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #6b7280;
        margin-bottom: 2rem;
    }
    .stat-card {
        padding: 1.5rem;
        border-radius: 0.5rem;
        border: 2px solid #e5e7eb;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

def read_any_file(file, sheet_name=None):
    """Read CSV or Excel safely, with optional sheet_name for Excel"""
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    else:
        if sheet_name:
            return pd.read_excel(file, sheet_name=sheet_name)
        return pd.read_excel(file)

def process_refund_data(refund_file, qwt_file, returns_file, bulk_rto_file, safe_t_file, reim_file,
                        door_tat_min, door_tat_max, fba_tat_min):
    """Process all uploaded files and perform the analysis"""
    try:
        # Load Refund Data
        Refund_data = read_any_file(refund_file)
        
        # Filter only Refund type
        TYPE_COL = [c for c in Refund_data.columns if c.strip().lower() == "type"][0]
        Refund_data = Refund_data[Refund_data[TYPE_COL].astype(str).str.lower() == "refund"].copy()
        
        # Drop unwanted extra columns that might be present in the raw file
        unwanted_cols = ["date/time.1", "Todays", "Diff", "OrdersReturns", "FBA Re", "Reim", "Safe", "FBA", "Rep"]
        cols_to_drop = [c for c in Refund_data.columns if str(c).strip() in unwanted_cols]
        if cols_to_drop:
            Refund_data.drop(columns=cols_to_drop, inplace=True, errors="ignore")
        
        # Remove Product Sales = 0
        product_sales_col = "product sales"
        Refund_data[product_sales_col] = pd.to_numeric(Refund_data[product_sales_col], errors="coerce")
        Refund_data = Refund_data[Refund_data[product_sales_col] != 0].copy()
        
        # Convert date and calculate Date_Diff
        Refund_data['Date1'] = pd.to_datetime(
            Refund_data['date/time'],
            errors="coerce",
            dayfirst=True,
            format="mixed"
        ).dt.date

        Refund_data['Today'] = datetime.today().date()
        Refund_data['Date_Diff'] = (pd.to_datetime(Refund_data['Today']) - pd.to_datetime(Refund_data['Date1'])).dt.days
        
        # Load QWT and perform Door Ship lookup
        qwt = read_any_file(qwt_file)
        order_id_cols = [c for c in Refund_data.columns if "order" in c.lower() and "id" in c.lower()]
        if not order_id_cols:
            raise KeyError("Could not find order ID column in Refund data. Required columns: order id")
        order_id_col = order_id_cols[0]
        Refund_data["__key"] = Refund_data[order_id_col].astype(str).str.strip().str.upper()
        
        qwt_order_cols = [c for c in qwt.columns if "customer" in c.lower() and "order" in c.lower()]
        qwt_order_col = qwt_order_cols[0] if qwt_order_cols else None
        if qwt_order_col and qwt_order_col in qwt.columns:
            qwt_map = (
                qwt.assign(__key = qwt[qwt_order_col].astype(str).str.strip().str.upper())
                   .drop_duplicates("__key", keep="first")
                   .set_index("__key")[qwt_order_col]
            )
            Refund_data["Door Ship (Seller Flex)"] = Refund_data["__key"].map(qwt_map)
        else:
            Refund_data["Door Ship (Seller Flex)"] = pd.NA
        Refund_data.drop(columns="__key", inplace=True)
        
        # Load Returns and perform FBA Return lookup
        returns = read_any_file(returns_file)
        Refund_data.loc[:, "__key"] = Refund_data[order_id_col].astype(str).str.strip().str.upper()
        
        ret_order_cols = [c for c in returns.columns if "order" in c.lower() and "id" in c.lower()]
        ret_order_col = ret_order_cols[0] if ret_order_cols else None
        if ret_order_col and ret_order_col in returns.columns:
            returns.loc[:, "__key"] = returns[ret_order_col].astype(str).str.strip().str.upper()
            ret_map = (
                returns[["__key", ret_order_col]]
                  .drop_duplicates("__key", keep="first")
                  .set_index("__key")[ret_order_col]
            )
            Refund_data["FBA Return"] = Refund_data["__key"].map(ret_map)
        else:
            Refund_data["FBA Return"] = pd.NA
        Refund_data.drop(columns="__key", inplace=True)
        
        # Load Bulk RTO and perform Seller Flex Return lookup
        bulk_rto = read_any_file(bulk_rto_file, sheet_name="All" if not bulk_rto_file.name.lower().endswith(".csv") else None)
        Refund_data["__key"] = Refund_data["Door Ship (Seller Flex)"].fillna("").astype(str).str.strip().str.upper()
        Refund_data.loc[Refund_data["__key"] == "", "__key"] = pd.NA
        
        rto_order_cols = [c for c in bulk_rto.columns if "order" in c.lower() and "id" in c.lower()]
        rto_order_col = rto_order_cols[0] if rto_order_cols else None
        if rto_order_col and rto_order_col in bulk_rto.columns:
            bulk_rto["__key"] = bulk_rto[rto_order_col].astype(str).str.strip().str.upper()
            right_key = bulk_rto[["__key", rto_order_col]].drop_duplicates()
            right_key = right_key[right_key["__key"].astype(str) != "NAN"]
            rto_map = right_key.set_index("__key")[rto_order_col]
            Refund_data["Seller Flex Return"] = Refund_data["__key"].map(rto_map)
        else:
            Refund_data["Seller Flex Return"] = pd.NA
        Refund_data.drop(columns="__key", inplace=True, errors="ignore")
        
        # Load Safe-T Claim and perform lookup
        safeT = read_any_file(safe_t_file, sheet_name="Sheet1" if not safe_t_file.name.lower().endswith(".csv") else None)
        Refund_data.loc[:, "__key"] = Refund_data["Door Ship (Seller Flex)"].fillna("").astype(str).str.strip().str.upper()
        Refund_data.loc[Refund_data["__key"] == "", "__key"] = pd.NA
        
        if len(safeT.columns) > 0:
            lookup_col = safeT.columns[3] if len(safeT.columns) > 3 else safeT.columns[0]
            safeT.loc[:, "__key"] = safeT[lookup_col].astype(str).str.strip().str.upper()
            safeT_small = safeT[["__key", lookup_col]].drop_duplicates()
            safeT_small = safeT_small[safeT_small["__key"].astype(str) != "NAN"]
            
            temp_col = lookup_col
            if temp_col in Refund_data.columns:
                temp_col = temp_col + "_from_safet"
                safeT_small = safeT_small.rename(columns={lookup_col: temp_col})
                
            Refund_data = Refund_data.merge(safeT_small, on="__key", how="left")
            Refund_data.rename(columns={temp_col: "Safe T Claim"}, inplace=True, errors="ignore")
        else:
            Refund_data["Safe T Claim"] = pd.NA
        Refund_data.drop(columns="__key", inplace=True, errors="ignore")
        
        # Load Reimbursement and perform FBA Reimbursement lookup
        reim = read_any_file(reim_file)
        reim_order_cols = [c for c in reim.columns if "amazon" in c.lower() and "order" in c.lower()]
        reim_order_col = reim_order_cols[0] if reim_order_cols else None
        
        order_id_col_currents = [c for c in Refund_data.columns if "order id" in c]
        order_id_col_current = order_id_col_currents[0] if order_id_col_currents else order_id_col
        Refund_data.loc[:, "__key"] = Refund_data[order_id_col_current].astype(str).str.strip().str.upper()
        
        if reim_order_col and reim_order_col in reim.columns and "reason" in reim.columns:
            filtered_reim = reim[reim["reason"].isin(["CustomerReturn", "CustomerServiceIssue"])].copy()
            filtered_reim.loc[:, "__key"] = filtered_reim[reim_order_col].astype(str).str.strip().str.upper()
            filtered_reim_small = filtered_reim[["__key", reim_order_col]].drop_duplicates()
            
            temp_col = reim_order_col
            if temp_col in Refund_data.columns:
                temp_col = temp_col + "_from_reim"
                filtered_reim_small = filtered_reim_small.rename(columns={reim_order_col: temp_col})
                
            Refund_data = Refund_data.merge(filtered_reim_small, on="__key", how="left")
            Refund_data.rename(columns={temp_col: "FBA Reimbursement"}, inplace=True)
        else:
            Refund_data["FBA Reimbursement"] = pd.NA
        Refund_data.drop(columns="__key", inplace=True, errors="ignore")
        
        # Create filtered dataframes
        filtered_doorship = Refund_data[
            Refund_data["Door Ship (Seller Flex)"].notna() &
            Refund_data["FBA Return"].isna() &
            Refund_data["Seller Flex Return"].isna() &
            Refund_data["Safe T Claim"].isna()
        ].copy()
        
        fba_return_df = Refund_data[
            (Refund_data["Door Ship (Seller Flex)"].isna()) &
            (Refund_data["FBA Return"].isna()) &
            (Refund_data["Seller Flex Return"].isna()) &
            (
                Refund_data["FBA Reimbursement"].isna() |
                (Refund_data["FBA Reimbursement"].astype(str).str.strip() == "")
            )
        ].copy()
        
        # 🔹 Use dynamic TAT filters instead of fixed values
        filtered_df_TAT = filtered_doorship[
            filtered_doorship["Date_Diff"].between(door_tat_min, door_tat_max, inclusive="both")
        ].copy()
        
        fba_return_TAT = fba_return_df[fba_return_df["Date_Diff"] >= fba_tat_min].copy()
        
        # Map brand column using PM file
        Refund_data = map_brand_column_from_pm(Refund_data)
        filtered_doorship = map_brand_column_from_pm(filtered_doorship)
        fba_return_df = map_brand_column_from_pm(fba_return_df)
        filtered_df_TAT = map_brand_column_from_pm(filtered_df_TAT)
        fba_return_TAT = map_brand_column_from_pm(fba_return_TAT)
        
        return {
            'main': Refund_data,
            'filtered_doorship': filtered_doorship,
            'fba_return': fba_return_df,
            'doorship_tat': filtered_df_TAT,
            'fba_return_tat': fba_return_TAT
        }
        
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None

# Main App
st.markdown('<div class="main-header">📊 Amazon Refund Cross Check</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your files to analyze refund and return data</div>', unsafe_allow_html=True)

# File Upload Section
st.markdown("### 📁 Upload Required Files")

col1, col2 = st.columns(2)

with col1:
    refund_file = st.file_uploader("Refund Data (Excel/CSV)", type=['xlsx', 'xls','csv'], key="refund")
    qwt_file = st.file_uploader("QWT Customer Shipments (Excel/CSV)", type=['xlsx','csv'], key="qwt")
    returns_file = st.file_uploader("Returns (Excel/CSV)", type=['xlsx','csv'], key="returns")

with col2:
    bulk_rto_file = st.file_uploader("Bulk RTO Returns (Excel/CSV)", type=['xlsx', 'xls','csv'], key="bulk")
    safe_t_file = st.file_uploader("Safe-T Claim (Excel/CSV)", type=['xlsx', 'xls','csv'], key="safe")
    reim_file = st.file_uploader("FBA Reimbursement (Excel/CSV)", type=['xlsx','csv'], key="reim")

# 🔹 TAT inputs for days
st.markdown("### ⏱️ TAT Day Filters")

tat_col1, tat_col2 = st.columns(2)

with tat_col1:
    door_tat_min = st.number_input(
        "Door Ship TAT start (days):",
        min_value=0,
        max_value=365,
        value=50,
        step=1,
        help="Starting day value for Door Ship TAT range."
    )
    door_tat_max = st.number_input(
        "Door Ship TAT end (days):",
        min_value=door_tat_min,
        max_value=365,
        value=75,
        step=1,
        help="Ending day value for Door Ship TAT range."
    )

with tat_col2:
    fba_tat_min = st.number_input(
        "FBA Return TAT minimum days:",
        min_value=0,
        max_value=365,
        value=40,
        step=1,
        help="Filter FBA Return records with Date_Diff ≥ this number of days."
    )

# Process Button
all_files = [refund_file, qwt_file, returns_file, bulk_rto_file, safe_t_file, reim_file]
if all(all_files):
    if st.button("🔍 Analyze Refund Data", type="primary", use_container_width=True):
        with st.spinner("Processing data..."):
            results = process_refund_data(*all_files, door_tat_min, door_tat_max, fba_tat_min)
            
            if results:
                st.success("✅ Analysis completed successfully!")
                
                # Save generated report to local folder for reports dashboard
                try:
                    with pd.ExcelWriter("refund_leakage.xlsx", engine="openpyxl") as writer:
                        results['main'].to_excel(writer, sheet_name="Full Data", index=False)
                        results['doorship_tat'].to_excel(writer, sheet_name="Door Ship TAT", index=False)
                        results['fba_return_tat'].to_excel(writer, sheet_name="FBA Return TAT", index=False)
                    st.info("💾 Saved report automatically as 'refund_leakage.xlsx' for the dashboard.")
                except Exception as e:
                    st.warning(f"⚠️ Could not auto-save reports: {str(e)}")
                
                # Display Statistics
                st.markdown("### 📈 Analysis Results")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total Refunds", f"{len(results['main']):,}")
                    st.metric("Door Ship Returns", f"{len(results['filtered_doorship']):,}")
                
                with col2:
                    st.metric("FBA Returns (Missing)", f"{len(results['fba_return']):,}")
                    st.metric(
                        f"Door Ship ({door_tat_min}-{door_tat_max} days TAT)",
                        f"{len(results['doorship_tat']):,}"
                    )
                
                with col3:
                    st.metric("Safe-T Claims", f"{results['main']['Safe T Claim'].notna().sum():,}")
                    st.metric(
                        f"FBA Return (≥{int(fba_tat_min)} days TAT)",
                        f"{len(results['fba_return_tat']):,}"
                    )
                
                # Download Buttons
                st.markdown("### 💾 Download Reports")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        results['main'].to_excel(writer, index=False, sheet_name='All Refunds')
                    st.download_button(
                        "⬇️ Download Full Report",
                        buffer.getvalue(),
                        "full_refund_report.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    buffer2 = io.BytesIO()
                    with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
                        results['doorship_tat'].to_excel(writer, index=False, sheet_name='Door Ship TAT')
                    st.download_button(
                        f"⬇️ Download Door Ship TAT ({door_tat_min}-{door_tat_max}d)",
                        buffer2.getvalue(),
                        "door_ship_tat.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col3:
                    buffer3 = io.BytesIO()
                    with pd.ExcelWriter(buffer3, engine='openpyxl') as writer:
                        results['fba_return_tat'].to_excel(writer, index=False, sheet_name='FBA Return TAT')
                    st.download_button(
                        f"⬇️ Download FBA Return TAT (≥{int(fba_tat_min)}d)",
                        buffer3.getvalue(),
                        "fba_return_tat.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
else:
    st.info("👆 Please upload all required files to begin analysis")

# Footer
st.markdown("---")
st.markdown("*Developed for Amazon Seller Refund Analysis By IBI*")




