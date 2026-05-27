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

st.set_page_config(page_title="Amazon Replacement Without Reimbursement Analyzer", page_icon="🔄", layout="wide")

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
    .metric-card {
        padding: 1rem;
        border-radius: 0.5rem;
        border: 2px solid #e5e7eb;
        background-color: #f9fafb;
    }
</style>
""", unsafe_allow_html=True)

def read_any_file(file, sheet_name=None):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    else:
        if sheet_name:
            return pd.read_excel(file, sheet_name=sheet_name)
        return pd.read_excel(file)
        
def process_replacement_data(replace_file, return_file, refund_file, bulk_rto_file, reim_file, days_threshold: int):
    """Process replacement data with all lookups and filters"""
    try:
        # Load Replace.csv
        Replace = read_any_file(replace_file)
        
        # Convert date and calculate difference
        Replace['Date'] = pd.to_datetime(
            Replace['shipment-date'],
            errors='coerce',
            utc=True
        ).dt.tz_convert(None)
        
        today = pd.Timestamp.now().normalize()
        
        Replace['Date_Difference'] = (
            today - Replace['Date']
        ).dt.days
        # Replace['Today_Date'] = date.today()
        # Replace['Date_Difference'] = (pd.to_datetime(Replace['Today_Date']) - pd.to_datetime(Replace['Date'])).dt.days
        # Convert date and calculate difference
        # Replace['Date'] = pd.to_datetime(
        #     Replace['shipment-date'],
        #     errors='coerce',
        #     utc=True
        # ).dt.date
        
        # Replace['Today_Date'] = pd.Timestamp.utcnow().date()
        
        # Replace['Date_Difference'] = (
        #     pd.Timestamp.utcnow().normalize() -
        #     pd.to_datetime(Replace['Date'])
        # ).dt.days
        
        # Load Return.csv
        Return = read_any_file(return_file)
        
        # FBA Original Return Lookup (Column I → Column B → Column 8)
        lookup_value_col = Replace.columns[8]
        lookup_key_col = Return.columns[1]
        return_value_col = Return.columns[8]
        
        Replace = Replace.merge(
            Return[[lookup_key_col, return_value_col]],
            how="left",
            left_on=lookup_value_col,
            right_on=lookup_key_col
        )
        Replace.rename(columns={return_value_col: "FBA Original Return"}, inplace=True)
        Replace.drop(columns=[lookup_key_col], inplace=True, errors="ignore")
        
        # FBA Replacement Return Lookup (Column H → Column B → Column 8)
        lookup_value_col = Replace.columns[7]
        lookup_key_col = Return.columns[1]
        return_value_col = Return.columns[8]
        
        Replace = Replace.merge(
            Return[[lookup_key_col, return_value_col]],
            how="left",
            left_on=lookup_value_col,
            right_on=lookup_key_col
        )
        Replace.rename(columns={return_value_col: "FBA Replacement Return"}, inplace=True)
        Replace.drop(columns=[lookup_key_col], inplace=True, errors="ignore")
        
        # Filter 1: Damaged Returns
        filtered_df = Replace[
            (
                Replace["FBA Original Return"].isin(["CARRIER_DAMAGED", "CUSTOMER_DAMAGED"]) &
                Replace["FBA Replacement Return"].isin(["CARRIER_DAMAGED", "CUSTOMER_DAMAGED"])
            )
        ].copy()
        
        # Load Reimbursement
        Reimbursement =read_any_file(reim_file, sheet_name='Sheet1')
        
        # Filter reimbursement data
        filtered_reimb = Reimbursement[
            Reimbursement["reason"].isin(["CustomerReturn", "CustomerServiceIssue"])
        ].copy()
        
        # Add CountIF
        filtered_reimb.loc[:, "CountIF"] = (
            filtered_reimb.groupby("amazon-order-id")["amazon-order-id"].transform("count")
        )
        
        # Merge CountIF into filtered_df
        lookup_value_col = filtered_df.columns[8]
        lookup_key_col = filtered_reimb.columns[3]
        return_value_col = filtered_reimb.columns[18]
        
        filtered_df = filtered_df.merge(
            filtered_reimb[[lookup_key_col, return_value_col]],
            how="left",
            left_on=lookup_value_col,
            right_on=lookup_key_col
        )
        filtered_df.rename(columns={return_value_col: "CountIF"}, inplace=True)
        filtered_df.drop(columns=[lookup_key_col], inplace=True, errors="ignore")
        
        # 🔍 Filter: CountIF = 1 and Date_Difference >= days_threshold
        filtered_df_final = filtered_df[
            (filtered_df["CountIF"] == 1.0) &
            (filtered_df["Date_Difference"] >= days_threshold)
        ].copy()
        
        # Load Refund Only file
        Refund = read_any_file(refund_file, sheet_name='Sheet1')
        
        # Refund Check Lookup
        lookup_value_col = Replace.columns[8]
        lookup_key_col = Refund.columns[4]
        
        refund_map = Refund[[lookup_key_col]].drop_duplicates().set_index(lookup_key_col)
        refund_series = pd.Series(refund_map.index, index=refund_map.index)
        Replace["Refund Check"] = Replace[lookup_value_col].map(refund_series)
        
        # Filter 2: Refund without Returns
        filtered_df_step2 = Replace[
            (
                (Replace["FBA Original Return"].isna()) |
                (Replace["FBA Original Return"].astype(str).str.upper().eq("NA"))
            ) &
            (
                (Replace["FBA Replacement Return"].isna()) |
                (Replace["FBA Replacement Return"].astype(str).str.upper().eq("NA"))
            ) &
            (
                (~Replace["Refund Check"].astype(str).str.upper().eq("NA")) &
                (~Replace["Refund Check"].isna())
            ) &
            (
                (Replace["Date_Difference"] >= days_threshold)
            )
        ].copy()
        
        # Load Bulk RTO
        BulkRTO = read_any_file(bulk_rto_file, sheet_name="All")
        
        # Door Step Return Lookup
        lookup_value_col = filtered_df_step2.columns[8]
        lookup_key_col = BulkRTO.columns[0]
        
        blk_map = BulkRTO[[lookup_key_col]].drop_duplicates().set_index(lookup_key_col)
        blk_series = pd.Series(blk_map.index, index=blk_map.index)
        filtered_df_step2["Door Step Return"] = filtered_df_step2[lookup_value_col].map(blk_series)
        
        # Filter 3: No Door Step Return
        filtered_df_step3 = filtered_df_step2[
            (filtered_df_step2["Door Step Return"].isna()) |
            (filtered_df_step2["Door Step Return"].astype(str).str.strip().str.upper().eq("NA"))
        ].copy()
        
        # Map brand column using PM file
        Replace = map_brand_column_from_pm(Replace)
        filtered_df_final = map_brand_column_from_pm(filtered_df_final)
        filtered_df_step3 = map_brand_column_from_pm(filtered_df_step3)
        
        return {
            'main': Replace,
            'damaged_returns': filtered_df_final,
            'refund_without_return': filtered_df_step3,
            'damaged_count': len(filtered_df_final),
            'refund_count': len(filtered_df_step3)
        }
        
    except Exception as e:
        st.error(f"Error processing files: {str(e)}")
        return None

# Main App
st.markdown('<div class="main-header">🔄 Amazon Replacement Without Reimbursement Data Analyzer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your files to analyze replacement and return data</div>', unsafe_allow_html=True)

# File Upload Section
st.markdown("### 📁 Upload Required Files")

col1, col2 = st.columns(2)

with col1:
    replace_file = st.file_uploader("Replace File", type=['csv','xlsx'], key="replace")
    return_file = st.file_uploader("Return File", type=['csv','xlsx'], key="return")
    refund_file = st.file_uploader("Refund Only File", type=['xlsx','csv'], key="refund")

with col2:
    bulk_rto_file = st.file_uploader("Bulk RTO Returns File", type=['xlsx','csv'], key="bulk")
    reim_file = st.file_uploader("Reimbursement File", type=['xlsx','csv'], key="reim")

# ⏱️ Days slicer / input
st.markdown("### ⏱️ Days Filter")
days_threshold = st.number_input(
    "Consider records older than (days):",
    min_value=0,
    max_value=365,
    value=40,       # default 40 days
    step=1,
    help="Example: 10, 30, 40, 45, 60..."
)

# Process Button
all_files = [replace_file, return_file, refund_file, bulk_rto_file, reim_file]
if all(all_files):
    if st.button("🔍 Analyze Replacement Data", type="primary", use_container_width=True):
        with st.spinner("Processing data... This may take a moment."):
            results = process_replacement_data(*all_files, days_threshold)
            
            if results:
                st.success("✅ Analysis completed successfully!")
                
                # Save generated report to local folder for reports dashboard
                try:
                    with pd.ExcelWriter("replacement_leakage.xlsx", engine="openpyxl") as writer:
                        results['main'].to_excel(writer, sheet_name="Full Data", index=False)
                        results['damaged_returns'].to_excel(writer, sheet_name="Damaged Returns", index=False)
                        results['refund_without_return'].to_excel(writer, sheet_name="Refund Without Return", index=False)
                    st.info("💾 Saved report automatically as 'replacement_leakage.xlsx' for the dashboard.")
                except Exception as e:
                    st.warning(f"⚠️ Could not auto-save reports: {str(e)}")
                
                # Display Statistics
                st.markdown("### 📈 Analysis Results")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total Replacements", f"{len(results['main']):,}")
                
                with col2:
                    st.metric(
                        f"Damaged Returns (≥{int(days_threshold)} days)",
                        f"{results['damaged_count']:,}", 
                        help=f"Replacements with CARRIER_DAMAGED or CUSTOMER_DAMAGED status and age ≥ {int(days_threshold)} days"
                    )
                
                with col3:
                    st.metric(
                        f"Refund without Return (≥{int(days_threshold)} days)",
                        f"{results['refund_count']:,}",
                        help=f"Refunds processed but no return record found, age ≥ {int(days_threshold)} days"
                    )
                
                # Data Preview
                st.markdown("### 📊 Data Preview")
                
                tab1, tab2, tab3 = st.tabs(["Damaged Returns", "Refund Without Return", "Full Data"])
                
                with tab1:
                    st.markdown(f"**Replacements with damaged items (≥{int(days_threshold)} days old)**")
                    st.dataframe(results['damaged_returns'], use_container_width=True)
                
                with tab2:
                    st.markdown(f"**Replacements with refund but no return record (≥{int(days_threshold)} days old)**")
                    st.dataframe(results['refund_without_return'], use_container_width=True)
                
                with tab3:
                    st.markdown("**All processed replacement data**")
                    st.dataframe(results['main'], use_container_width=True)
                
                # Download Buttons
                st.markdown("### 💾 Download Reports")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        results['damaged_returns'].to_excel(writer, index=False, sheet_name='Damaged Returns')
                    st.download_button(
                        "⬇️ Download Damaged Returns",
                        buffer.getvalue(),
                        "damaged_returns_report.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    buffer2 = io.BytesIO()
                    with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
                        results['refund_without_return'].to_excel(writer, index=False, sheet_name='Refund Without Return')
                    st.download_button(
                        "⬇️ Download Refund Without Return",
                        buffer2.getvalue(),
                        "refund_without_return_report.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col3:
                    buffer3 = io.BytesIO()
                    with pd.ExcelWriter(buffer3, engine='openpyxl') as writer:
                        results['main'].to_excel(writer, index=False, sheet_name='All Replacements')
                    st.download_button(
                        "⬇️ Download Full Report",
                        buffer3.getvalue(),
                        "full_replacement_report.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
else:
    st.info("👆 Please upload all required files to begin analysis")

# Footer
st.markdown("---")
st.markdown("*Amazon Seller Replacement & Return Analysis Tool*")

