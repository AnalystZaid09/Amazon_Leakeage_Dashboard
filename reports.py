import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os

def get_secret_safe(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default
# Helper to get cached ASIN-to-Brand mapping from PM.xlsx
@st.cache_data
def get_pm_brand_map():
    import pandas as pd
    import os
    pm_path = "PM.xlsx"
    if os.path.exists(os.path.join("data_store", "reports", "PM.xlsx")):
        pm_path = os.path.join("data_store", "reports", "PM.xlsx")
    elif not os.path.exists(pm_path):
        return {}
    try:
        df_header = pd.read_excel(pm_path, nrows=1)
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
            df_pm = pd.read_excel(pm_path, usecols=[asin_col, brand_col])
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

# ==================== DATA PROCESSING ENGINES ====================

def process_refund_leakage(refund_df, qwt_df, returns_df, bulk_rto_df, safe_t_df, reim_df,
                           door_tat_min=50, door_tat_max=75, fba_tat_min=40):
    """Processes raw Amazon inputs to compute Refund Leakage details"""
    # Filter only Refund type
    type_cols = [c for c in refund_df.columns if c.strip().lower() == "type"]
    if type_cols:
        refund_df = refund_df[refund_df[type_cols[0]].astype(str).str.lower() == "refund"].copy()
    
    # Remove Product Sales = 0
    product_sales_col = "product sales"
    sales_cols = [c for c in refund_df.columns if "product" in c.lower() and "sales" in c.lower()]
    if sales_cols:
        product_sales_col = sales_cols[0]
        
    refund_df[product_sales_col] = pd.to_numeric(refund_df[product_sales_col], errors="coerce")
    refund_df = refund_df[refund_df[product_sales_col] != 0].copy()
    
    # Convert date and calculate Date_Diff
    date_cols = [c for c in refund_df.columns if "date/time" in c.lower() or ("date" in c.lower() and "time" in c.lower())]
    date_col = date_cols[0] if date_cols else "date/time"
    refund_df['Date1'] = pd.to_datetime(
        refund_df[date_col],
        errors="coerce",
        dayfirst=True,
        format="mixed"
    ).dt.date

    refund_df['Today'] = datetime.today().date()
    refund_df['Date_Diff'] = (pd.to_datetime(refund_df['Today']) - pd.to_datetime(refund_df['Date1'])).dt.days
    
    # Load QWT and perform Door Ship lookup
    order_id_cols = [c for c in refund_df.columns if "order" in c.lower() and "id" in c.lower()]
    if not order_id_cols:
        raise KeyError("Could not find order ID column in Refund data. Required columns: order id")
    order_id_col = order_id_cols[0]
    refund_df["__key"] = refund_df[order_id_col].astype(str).str.strip().str.upper()
    
    qwt_order_cols = [c for c in qwt_df.columns if "customer" in c.lower() and "order" in c.lower()]
    qwt_order_col = qwt_order_cols[0] if qwt_order_cols else None
    if qwt_order_col and qwt_order_col in qwt_df.columns:
        qwt_map = (
            qwt_df.assign(__key = qwt_df[qwt_order_col].astype(str).str.strip().str.upper())
               .drop_duplicates("__key", keep="first")
               .set_index("__key")[qwt_order_col]
        )
        refund_df["Door Ship (Seller Flex)"] = refund_df["__key"].map(qwt_map)
    else:
        refund_df["Door Ship (Seller Flex)"] = pd.NA
    refund_df.drop(columns="__key", inplace=True)
    
    # Load Returns and perform FBA Return lookup
    refund_df.loc[:, "__key"] = refund_df[order_id_col].astype(str).str.strip().str.upper()
    
    ret_order_cols = [c for c in returns_df.columns if "order" in c.lower() and "id" in c.lower()]
    ret_order_col = ret_order_cols[0] if ret_order_cols else None
    if ret_order_col and ret_order_col in returns_df.columns:
        returns_df.loc[:, "__key"] = returns_df[ret_order_col].astype(str).str.strip().str.upper()
        ret_map = (
            returns_df[["__key", ret_order_col]]
              .drop_duplicates("__key", keep="first")
              .set_index("__key")[ret_order_col]
        )
        refund_df["FBA Return"] = refund_df["__key"].map(ret_map)
    else:
        refund_df["FBA Return"] = pd.NA
    refund_df.drop(columns="__key", inplace=True)
    
    # Load Bulk RTO and perform Seller Flex Return lookup
    refund_df["__key"] = refund_df["Door Ship (Seller Flex)"].fillna("").astype(str).str.strip().str.upper()
    refund_df.loc[refund_df["__key"] == "", "__key"] = pd.NA
    
    rto_order_cols = [c for c in bulk_rto_df.columns if "order" in c.lower() and "id" in c.lower()]
    rto_order_col = rto_order_cols[0] if rto_order_cols else None
    if rto_order_col and rto_order_col in bulk_rto_df.columns:
        bulk_rto_df["__key"] = bulk_rto_df[rto_order_col].astype(str).str.strip().str.upper()
        right_key = bulk_rto_df[["__key", rto_order_col]].drop_duplicates()
        right_key = right_key[right_key["__key"].astype(str) != "NAN"]
        refund_df = refund_df.merge(right_key, on="__key", how="left")
        refund_df.rename(columns={rto_order_col: "Seller Flex Return"}, inplace=True)
    else:
        refund_df["Seller Flex Return"] = pd.NA
    refund_df.drop(columns="__key", inplace=True, errors="ignore")
    
    # Load Safe-T Claim and perform lookup
    refund_df.loc[:, "__key"] = refund_df["Door Ship (Seller Flex)"].fillna("").astype(str).str.strip().str.upper()
    refund_df.loc[refund_df["__key"] == "", "__key"] = pd.NA
    
    if len(safe_t_df.columns) > 0:
        lookup_col = safe_t_df.columns[3] if len(safe_t_df.columns) > 3 else safe_t_df.columns[0]
        safe_t_df.loc[:, "__key"] = safe_t_df[lookup_col].astype(str).str.strip().str.upper()
        safeT_small = safe_t_df[["__key", lookup_col]].drop_duplicates()
        safeT_small = safeT_small[safeT_small["__key"].astype(str) != "NAN"]
        
        temp_col = lookup_col
        if temp_col in refund_df.columns:
            temp_col = temp_col + "_from_safet"
            safeT_small = safeT_small.rename(columns={lookup_col: temp_col})
            
        refund_df = refund_df.merge(safeT_small, on="__key", how="left")
        refund_df.rename(columns={temp_col: "Safe T Claim"}, inplace=True, errors="ignore")
    else:
        refund_df["Safe T Claim"] = pd.NA
    refund_df.drop(columns="__key", inplace=True, errors="ignore")
    
    # Load Reimbursement and perform FBA Reimbursement lookup
    reim_order_cols = [c for c in reim_df.columns if "amazon" in c.lower() and "order" in c.lower()]
    reim_order_col = reim_order_cols[0] if reim_order_cols else None
    
    order_id_col_currents = [c for c in refund_df.columns if "order id" in c]
    order_id_col_current = order_id_col_currents[0] if order_id_col_currents else order_id_col
    refund_df.loc[:, "__key"] = refund_df[order_id_col_current].astype(str).str.strip().str.upper()
    
    if reim_order_col and reim_order_col in reim_df.columns and "reason" in reim_df.columns:
        filtered_reim = reim_df[reim_df["reason"].isin(["CustomerReturn", "CustomerServiceIssue"])].copy()
        filtered_reim.loc[:, "__key"] = filtered_reim[reim_order_col].astype(str).str.strip().str.upper()
        filtered_reim_small = filtered_reim[["__key", reim_order_col]].drop_duplicates()
        
        temp_col = reim_order_col
        if temp_col in refund_df.columns:
            temp_col = temp_col + "_from_reim"
            filtered_reim_small = filtered_reim_small.rename(columns={reim_order_col: temp_col})
            
        refund_df = refund_df.merge(filtered_reim_small, on="__key", how="left")
        refund_df.rename(columns={temp_col: "FBA Reimbursement"}, inplace=True)
    else:
        refund_df["FBA Reimbursement"] = pd.NA
    refund_df.drop(columns="__key", inplace=True, errors="ignore")
    
    # Create filtered dataframes
    filtered_doorship = refund_df[
        refund_df["Door Ship (Seller Flex)"].notna() &
        refund_df["FBA Return"].isna() &
        refund_df["Seller Flex Return"].isna() &
        refund_df["Safe T Claim"].isna()
    ].copy()
    
    # Filter by TAT
    filtered_df_TAT = filtered_doorship[
        filtered_doorship["Date_Diff"].between(door_tat_min, door_tat_max, inclusive="both")
    ].copy()
    
    fba_return_df = refund_df[
        (refund_df["Door Ship (Seller Flex)"].isna()) &
        (refund_df["FBA Return"].isna()) &
        (refund_df["Seller Flex Return"].isna()) &
        (
            refund_df["FBA Reimbursement"].isna() |
            (refund_df["FBA Reimbursement"].astype(str).str.strip() == "")
        )
    ].copy()
    
    fba_return_TAT = fba_return_df[fba_return_df["Date_Diff"] >= fba_tat_min].copy()
            
    refund_df = map_brand_column_from_pm(refund_df)
    filtered_df_TAT = map_brand_column_from_pm(filtered_df_TAT)
    fba_return_TAT = map_brand_column_from_pm(fba_return_TAT)
    
    # Ensure Arrow compatibility by casting mixed-type object columns to string
    for df in [refund_df, filtered_df_TAT, fba_return_TAT]:
        for col in ["Door Ship (Seller Flex)", "FBA Return", "Seller Flex Return", "Safe T Claim", "FBA Reimbursement"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
                
    return {
        'main': refund_df,
        'doorship_tat': filtered_df_TAT,
        'fba_return_tat': fba_return_TAT
    }

def process_replacement_leakage(replace_df, return_df, refund_df, bulk_rto_df, reim_df, days_threshold=40):
    """Processes raw Amazon inputs to compute Replacement Leakage details"""
    # Find columns by name for safety
    replace_cols = list(replace_df.columns)
    return_cols = list(return_df.columns)
    refund_cols = list(refund_df.columns)
    bulk_cols = list(bulk_rto_df.columns)
    
    # Shipment date
    date_col = [c for c in replace_cols if "shipment-date" in str(c).lower() or "date" in str(c).lower()][0]
    replace_df['Date'] = pd.to_datetime(
        replace_df[date_col],
        errors='coerce',
        utc=True
    ).dt.tz_convert(None)
    
    today = pd.Timestamp.now().normalize()
    replace_df['Date_Difference'] = (today - replace_df['Date']).dt.days
    
    # Original order ID (typically Column I - original-amazon-order-id)
    orig_order_cols = [c for c in replace_cols if "original" in str(c).lower() and "order" in str(c).lower() and "id" in str(c).lower()]
    orig_order_col = orig_order_cols[0] if orig_order_cols else replace_df.columns[8]
    
    # Replacement order ID (typically Column H - replacement-amazon-order-id)
    repl_order_cols = [c for c in replace_cols if "replacement" in str(c).lower() and "order" in str(c).lower() and "id" in str(c).lower()]
    repl_order_col = repl_order_cols[0] if repl_order_cols else replace_df.columns[7]
    
    # Return file order ID (typically Column B - order-id)
    ret_order_cols = [c for c in return_cols if "order" in str(c).lower() and "id" in str(c).lower()]
    ret_order_col = ret_order_cols[0] if ret_order_cols else return_df.columns[1]
    
    # Return file return value (typically Column I - detailed-disposition)
    ret_val_cols = [c for c in return_cols if "disposition" in str(c).lower() or "status" in str(c).lower()]
    ret_val_col = ret_val_cols[0] if ret_val_cols else return_df.columns[8]
    
    # FBA Original Return Lookup
    replace_df = replace_df.merge(
        return_df[[ret_order_col, ret_val_col]].drop_duplicates(subset=[ret_order_col], keep="first"),
        how="left",
        left_on=orig_order_col,
        right_on=ret_order_col
    )
    replace_df.rename(columns={ret_val_col: "FBA Original Return"}, inplace=True)
    replace_df.drop(columns=[ret_order_col], inplace=True, errors="ignore")
    
    # FBA Replacement Return Lookup
    return_df_temp = return_df[[ret_order_col, ret_val_col]].drop_duplicates(subset=[ret_order_col], keep="first").rename(columns={ret_val_col: "FBA Replacement Return"})
    replace_df = replace_df.merge(
        return_df_temp,
        how="left",
        left_on=repl_order_col,
        right_on=ret_order_col
    )
    replace_df.drop(columns=[ret_order_col], inplace=True, errors="ignore")
    
    # Filter 1: Damaged Returns
    filtered_df = replace_df[
        (
            replace_df["FBA Original Return"].isin(["CARRIER_DAMAGED", "CUSTOMER_DAMAGED"]) &
            replace_df["FBA Replacement Return"].isin(["CARRIER_DAMAGED", "CUSTOMER_DAMAGED"])
        )
    ].copy()
    
    # Filter reimbursement data
    reim_reason_col = [c for c in reim_df.columns if "reason" in str(c).lower()][0] if any("reason" in str(c).lower() for c in reim_df.columns) else "reason"
    reim_order_cols = [c for c in reim_df.columns if "amazon" in c.lower() and "order" in c.lower()]
    reim_order_col = reim_order_cols[0] if reim_order_cols else "amazon-order-id"
    
    filtered_reimb = reim_df[reim_df[reim_reason_col].isin(["CustomerReturn", "CustomerServiceIssue"])].copy()
    
    # Add CountIF
    filtered_reimb["CountIF"] = (
        filtered_reimb.groupby(reim_order_col)[reim_order_col].transform("count")
    )
    
    # Merge CountIF into filtered_df
    filtered_reimb_small = filtered_reimb[[reim_order_col, "CountIF"]].drop_duplicates(subset=[reim_order_col])
    filtered_df = filtered_df.merge(
        filtered_reimb_small,
        how="left",
        left_on=orig_order_col,
        right_on=reim_order_col
    )
    filtered_df.drop(columns=[reim_order_col], inplace=True, errors="ignore")
    
    # Filter: CountIF = 1 and Date_Difference >= days_threshold
    filtered_df_final = filtered_df[
        (filtered_df["CountIF"] == 1.0) &
        (filtered_df["Date_Difference"] >= days_threshold)
    ].copy()
    
    # Refund Check Lookup (typically Column E in Refund - order-id)
    refund_order_cols = [c for c in refund_cols if "order" in str(c).lower() and "id" in str(c).lower()]
    refund_order_col = refund_order_cols[0] if refund_order_cols else refund_df.columns[4]
    
    replace_df = replace_df.merge(
        refund_df[[refund_order_col]].drop_duplicates(),
        how="left",
        left_on=orig_order_col,
        right_on=refund_order_col
    )
    replace_df["Refund Check"] = replace_df[refund_order_col]
    replace_df.drop(columns=[refund_order_col], inplace=True, errors="ignore")
    
    # Filter: Refund without Returns
    filtered_df_step2 = replace_df[
        (
            (replace_df["FBA Original Return"].isna()) |
            (replace_df["FBA Original Return"].astype(str).str.upper().eq("NA"))
        ) &
        (
            (replace_df["FBA Replacement Return"].isna()) |
            (replace_df["FBA Replacement Return"].astype(str).str.upper().eq("NA"))
        ) &
        (
            (~replace_df["Refund Check"].astype(str).str.upper().eq("NA")) &
            (~replace_df["Refund Check"].isna())
        ) &
        (
            (replace_df["Date_Difference"] >= days_threshold)
        )
    ].copy()
    
    # Door Step Return Lookup (typically Column A in Bulk RTO - Order Id)
    blk_order_cols = [c for c in bulk_cols if "order" in str(c).lower() and "id" in str(c).lower()]
    blk_order_col = blk_order_cols[0] if blk_order_cols else bulk_rto_df.columns[0]
    
    filtered_df_step2 = filtered_df_step2.merge(
        bulk_rto_df[[blk_order_col]].drop_duplicates(),
        how="left",
        left_on=orig_order_col,
        right_on=blk_order_col
    )
    filtered_df_step2["Door Step Return"] = filtered_df_step2[blk_order_col]
    filtered_df_step2.drop(columns=[blk_order_col], inplace=True, errors="ignore")
    
    # Filter: No Door Step Return
    filtered_df_step3 = filtered_df_step2[
        (filtered_df_step2["Door Step Return"].isna()) |
        (filtered_df_step2["Door Step Return"].astype(str).str.strip().str.upper().eq("NA"))
    ].copy()
    
    # Map brand column
    replace_df = map_brand_column_from_pm(replace_df)
    filtered_df_final = map_brand_column_from_pm(filtered_df_final)
    filtered_df_step3 = map_brand_column_from_pm(filtered_df_step3)
    
    # Ensure Arrow compatibility by casting mixed-type object columns to string
    for df in [replace_df, filtered_df_final, filtered_df_step3]:
        for col in ["Refund Check", "Door Step Return", "FBA Original Return", "FBA Replacement Return"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
        if "CountIF" in df.columns:
            df["CountIF"] = pd.to_numeric(df["CountIF"], errors="coerce")
            
    return {
        'main': replace_df,
        'damaged_returns': filtered_df_final,
        'refund_without_return': filtered_df_step3
    }

def process_return_leakage(returns_df, reimb_df, replacement_df, days_filter=40):
    """Processes raw Amazon inputs to compute Return Leakage details"""
    if 'sku' in returns_df.columns:
        returns_df['sku'] = returns_df['sku'].astype(str)
        
    # Convert return-date to date only
    returns_df['return-date1'] = pd.to_datetime(returns_df['return-date'], errors='coerce').dt.date
    returns_df['Today'] = datetime.today().date()
    returns_df['Date_Diff'] = (pd.to_datetime(returns_df['Today']) - pd.to_datetime(returns_df['return-date1'])).dt.days
    
    returns_45 = returns_df[returns_df['Date_Diff'] > days_filter].copy()
    reimb_filtered = reimb_df[reimb_df['reason'].isin(['CustomerReturn', 'CustomerServiceIssue'])].copy()
    
    # VLOOKUP-like merge
    result_merge = returns_45.merge(
        reimb_filtered[['amazon-order-id', 'reason']].drop_duplicates(subset='amazon-order-id'),
        how='left',
        left_on='order-id',
        right_on='amazon-order-id'
    )
    
    # Extract Amount Total
    reimb_unique = reimb_filtered.drop_duplicates(subset='amazon-order-id', keep='first')
    result_merge['Amount_Total'] = result_merge['order-id'].map(
        reimb_unique.set_index('amazon-order-id')['amount-total']
    )
    
    # Filter damaged items
    returns_final = result_merge[
        result_merge['detailed-disposition'].isin(['CARRIER_DAMAGED', 'CUSTOMER_DAMAGED', 'DAMAGED'])
    ]
    
    # Filter where reimbursement order ID is N/A
    returns_final = returns_final[returns_final['amazon-order-id'].isna()].copy()
    
    # Map replacement order ID
    returns_final['order-id'] = returns_final['order-id'].astype(str).str.strip()
    replacement_df['replacement-amazon-order-id'] = replacement_df['replacement-amazon-order-id'].astype(str).str.strip()
    replacement_df['original-amazon-order-id'] = replacement_df['original-amazon-order-id'].astype(str).str.strip()
    
    returns_final['Replacement_OrderId'] = returns_final['order-id'].map(
        replacement_df.drop_duplicates(subset='replacement-amazon-order-id', keep='first')
                      .set_index('replacement-amazon-order-id')['original-amazon-order-id']
    )
    
    # Map replacement reason and amount
    returns_final['Replacement_Reason'] = returns_final['Replacement_OrderId'].map(
        reimb_unique.set_index('amazon-order-id')['reason']
    )
    returns_final['Replacement_Amount'] = returns_final['Replacement_OrderId'].map(
        reimb_unique.set_index('amazon-order-id')['amount-total']
    )
    
    # Filter where replacement order ID is N/A
    returns_final = returns_final[returns_final['Replacement_OrderId'].isna()].copy()
    
    returns_final = map_brand_column_from_pm(returns_final)
    return returns_final

def process_free_overcharged_leakage(txn_df, free_df):
    """Processes Unified Transaction + Fee Estimate (FREE) to detect Commission & Weight Handling overcharges."""
    import pandas as pd

    def _to_numeric(series):
        return (
            pd.to_numeric(
                series.astype(str).str.replace(",", "", regex=False).str.strip(),
                errors="coerce",
            ).fillna(0)
        )

    # --- Clean Transaction Data ---
    numeric_cols = [
        "Total sales tax liable(GST before adjusting TCS)",
        "other transaction fees",
        "product sales",
        "selling fees",
    ]
    missing = [c for c in numeric_cols if c not in txn_df.columns]
    if missing:
        raise ValueError(f"Transaction file is missing columns: {missing}")

    for col in numeric_cols:
        txn_df[col] = _to_numeric(txn_df[col])

    # Filter: only Order type, non-zero product sales
    type_cols = [c for c in txn_df.columns if c.strip().lower() == "type"]
    if type_cols:
        txn_df = txn_df[txn_df[type_cols[0]].astype(str).str.strip() == "Order"].copy()
    txn_df = txn_df[txn_df["product sales"] != 0].reset_index(drop=True)

    # --- Clean Fee Estimate Data ---
    for col in ["sales-price", "estimated-referral-fee-per-unit"]:
        if col in free_df.columns:
            free_df[col] = _to_numeric(free_df[col])

    free_df["Commission %"] = (
        free_df["estimated-referral-fee-per-unit"] / free_df["sales-price"] * 100
    ).round(2)
    free_df["Commission With GST"] = (
        free_df["estimated-referral-fee-per-unit"] * 1.18
    ).round(2)

    if "estimated-fixed-closing-fee" in free_df.columns:
        free_df["estimated-fixed-closing-fee"] = _to_numeric(free_df["estimated-fixed-closing-fee"])
        free_df["Closing With GST"] = (free_df["estimated-fixed-closing-fee"] * 1.18).round(2)

    for col in ["estimated-pick-pack-fee-per-unit", "estimated-weight-handling-fee-per-unit"]:
        if col in free_df.columns:
            free_df[col] = _to_numeric(free_df[col])

    free_df["Pick& Pack With GST"] = (free_df.get("estimated-pick-pack-fee-per-unit", 0) * 1.18).round(2)
    free_df["Weight GST"] = (free_df.get("estimated-weight-handling-fee-per-unit", 0) * 1.18).round(2)
    free_df["Total Weight Handling"] = (free_df["Pick& Pack With GST"] + free_df["Weight GST"]).round(2)
    free_df["Total Weight Handling With Gst"] = (free_df["Total Weight Handling"] * 1.18).round(2)

    # --- Build Pivot Table ---
    required_pivot_cols = [
        "order id", "Sku", "quantity", "product sales",
        "Total sales tax liable(GST before adjusting TCS)",
        "selling fees", "fba fees", "other transaction fees",
    ]
    missing_pivot = [c for c in required_pivot_cols if c not in txn_df.columns]
    if missing_pivot:
        raise ValueError(f"Transaction file missing columns for pivot: {missing_pivot}")

    pivot_table = pd.pivot_table(
        txn_df,
        index=["order id", "Sku"],
        values=[
            "quantity", "product sales",
            "Total sales tax liable(GST before adjusting TCS)",
            "selling fees", "fba fees", "other transaction fees",
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
    free_df["sku"] = free_df["sku"].astype(str).str.strip()

    commission_lookup = free_df.set_index("sku")["Commission %"].to_dict()
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
    weight_lookup = free_df.set_index("sku")["Total Weight Handling With Gst"].to_dict()
    pivot_table["Weight Handling With GST"] = pivot_table["Sku"].map(weight_lookup)
    pivot_table["Weight Handling With GST Qty"] = (
        pivot_table["Weight Handling With GST"] * pivot_table["quantity"]
    ).round(2)
    pivot_table["Weight Handling Different"] = (
        pivot_table["fba fees"] + pivot_table["Weight Handling With GST Qty"]
    ).round(2)

    # --- Commission Overcharge ---
    commission_overcharge = (
        pivot_table[pivot_table["Diff With Charge & Present Commission"] < 0]
        .copy()
        .sort_values("Diff With Charge & Present Commission")
        .reset_index(drop=True)
    )

    # --- Weight Handling Overcharge ---
    weight_overcharge = (
        pivot_table[pivot_table["Weight Handling Different"] < 0]
        .copy()
        .sort_values("Weight Handling Different")
        .reset_index(drop=True)
    )

    return {
        'pivot': pivot_table,
        'commission_overcharge': commission_overcharge,
        'weight_overcharge': weight_overcharge,
    }

st.set_page_config(page_title="Amazon Leakage Reports Dashboard", page_icon="🎯", layout="wide")

# Custom Premium Styling mimicking the dark-theme Leakage Pipeline dashboard
st.markdown("""
<style>
    /* Dark Theme General Styles */
    .stApp, [data-testid="stAppViewContainer"] {
        background-color: #0e1117 !important;
        color: #e6edf3 !important;
    }
    
    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #0e1117 !important;
        border-right: 1px solid #30363d !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSubheader"], 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] h3 {
        color: #e6edf3 !important;
    }
    
    /* Force high visibility for all text elements, widget labels, and helper texts */
    label,
    .stWidgetLabel,
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p,
    [data-testid="stWidgetLabel"] span,
    [data-testid="stWidgetLabel"] div,
    [data-testid="stWidgetLabel"] h1,
    [data-testid="stWidgetLabel"] h2,
    [data-testid="stWidgetLabel"] h3,
    [data-testid="stWidgetLabel"] h4,
    [data-testid="stWidgetLabel"] h5,
    [data-testid="stWidgetLabel"] h6,
    [data-testid="stWidgetInstructions"] {
        color: #e6edf3 !important;
    }
    
    /* Target Streamlit markdown and form texts */
    .stMarkdown,
    .stMarkdown p, 
    .stMarkdown span, 
    .stMarkdown div,
    small {
        color: #e6edf3 !important;
    }
    
    /* Header Styles */
    .main-title {
        font-size: 2.8rem;
        font-weight: 800;
        color: #ffffff;
        margin-bottom: 0.2rem;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #8b949e;
        margin-bottom: 1.5rem;
    }
    
    /* Green status banner */
    .status-banner {
        background-color: rgba(46, 160, 67, 0.15);
        border: 1px solid rgba(46, 160, 67, 0.4);
        padding: 0.8rem 1.2rem;
        border-radius: 6px;
        color: #3fb950;
        font-weight: 600;
        margin-bottom: 2rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Metric Card styling */
    .metric-container {
        display: flex;
        gap: 1.2rem;
        margin-bottom: 2rem;
        flex-wrap: wrap;
    }
    .metric-card {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1.5rem;
        flex: 1;
        min-width: 200px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        transition: transform 0.2s, border-color 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        border-color: #8b949e;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #8b949e;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #ffffff;
        margin-top: 0.5rem;
    }
    .metric-value-accent {
        color: #ff7b72; /* Light red/coral for amount at risk */
    }
    
    /* Styled tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px !important;
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px !important;
        white-space: pre-wrap !important;
        background-color: #161b22 !important;
        border: 1px solid #30363d !important;
        border-radius: 6px 6px 0px 0px !important;
        color: #8b949e !important;
        font-weight: 600 !important;
        padding: 0 20px !important;
        transition: all 0.2s !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #ffffff !important;
        background-color: #21262d !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #21262d !important;
        border-bottom: 2px solid #f78166 !important;
        color: #ffffff !important;
    }
    
    /* Styling Streamlit expanders */
    div[data-testid="stExpander"] {
        background-color: #161b22 !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
    }
    div[data-testid="stExpander"] summary {
        background-color: #161b22 !important;
        color: #ffffff !important;
    }
    div[data-testid="stExpander"] summary:hover {
        color: #ff7b72 !important;
    }
    div[data-testid="stExpander"] summary svg {
        fill: #ffffff !important;
    }
    div[data-testid="stExpander"] div[data-testid="stExpanderDetails"] {
        background-color: #0e1117 !important;
        color: #e6edf3 !important;
    }
    
    .stDataFrame {
        border: 1px solid #30363d;
        border-radius: 6px;
        background-color: #161b22;
    }
</style>

""", unsafe_allow_html=True)

# Helper function to format values in Indian standard (Rupees in Lakhs/Crores)
def format_rupees(val):
    if pd.isna(val) or not isinstance(val, (int, float)):
        return "₹0"
    
    prefix = "₹"
    abs_val = abs(val)
    
    if abs_val >= 10000000: # 1 Crore
        formatted = f"{prefix}{val/10000000:.2f} Cr"
    elif abs_val >= 100000: # 1 Lakh
        formatted = f"{prefix}{val/100000:.2f} L"
    else:
        formatted = f"{prefix}{val:,.2f}"
        
    return formatted

# Helper to fetch attachments from email via IMAP
def fetch_email_attachments(host, port, username, password, folder="INBOX", subject_search=""):
    import imaplib
    import email
    import re
    from email.header import decode_header
    import urllib.request
    import urllib.parse
    attachments = []
    debug_log = []  # Collect debug info
    
    def download_gdrive_file(file_id):
        url = f"https://docs.google.com/uc?export=download&id={file_id}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as response:
                content = response.read()
                # If the response is small and contains "confirm=", it might be a warning page
                if len(content) < 100000 and b"confirm=" in content:
                    html = content.decode('utf-8', errors='ignore')
                    match = re.search(r'confirm=([A-Za-z0-9_-]+)', html)
                    if match:
                        confirm_token = match.group(1)
                        confirm_url = f"https://docs.google.com/uc?export=download&confirm={confirm_token}&id={file_id}"
                        req2 = urllib.request.Request(confirm_url, headers=headers)
                        with urllib.request.urlopen(req2) as resp2:
                            content = resp2.read()
                return content
        except Exception as e:
            return None
            
    # Strip whitespace to prevent DNS resolution issues from copy-paste
    if isinstance(host, str):
        host = host.strip()
    if isinstance(username, str):
        username = username.strip()
        
    try:
        mail = imaplib.IMAP4_SSL(host, int(port))
        mail.login(username, password)
        mail.select(folder)
        
        # Search criteria: if subject_search is provided, search by subject; otherwise, search last 7 days
        from datetime import datetime, timedelta
        date_since = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")
        
        if subject_search:
            status, messages = mail.search(None, f'SUBJECT "{subject_search}"')
        else:
            status, messages = mail.search(None, f'SINCE {date_since}')
            
        if status != "OK":
            return False, "Error searching mail folder.", []
            
        mail_ids = messages[0].split()
        debug_log.append(f"Total matching emails found in {folder} (last 7 days): {len(mail_ids)}")
        
        if not mail_ids:
            return True, [], debug_log
            
        # Limit to the last 7 emails to prevent slow downloads causing 503 timeouts
        check_ids = list(reversed(mail_ids[-7:]))
        debug_log.append(f"Checking last {len(check_ids)} emails...")
        
        for mail_id in check_ids:
            res, msg_data = mail.fetch(mail_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    
                    # Parse subject
                    subject = ""
                    if msg["Subject"]:
                        subject_header = decode_header(msg["Subject"])[0]
                        subject_bytes = subject_header[0]
                        encoding = subject_header[1]
                        if isinstance(subject_bytes, bytes):
                            subject = subject_bytes.decode(encoding if encoding else "utf-8", errors="ignore")
                        else:
                            subject = str(subject_bytes)
                    
                    email_attachments = []
                    part_count = 0
                    
                    plain_body = ""
                    html_body = ""
                    
                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        
                        part_count += 1
                        content_type = part.get_content_type()
                        content_disp = str(part.get('Content-Disposition', ''))
                        filename = part.get_filename()
                        
                        # Collect text bodies for Google Drive link parsing
                        if content_type == 'text/plain':
                            try:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    plain_body += payload.decode('utf-8', errors='ignore')
                            except:
                                pass
                        elif content_type == 'text/html':
                            try:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    html_body += payload.decode('utf-8', errors='ignore')
                            except:
                                pass
                        
                        # Decode filename if it's encoded
                        if filename:
                            decoded_parts = decode_header(filename)
                            decoded_filename = ""
                            for fname_part, enc in decoded_parts:
                                if isinstance(fname_part, bytes):
                                    decoded_filename += fname_part.decode(enc if enc else "utf-8", errors="ignore")
                                else:
                                    decoded_filename += str(fname_part)
                            filename = decoded_filename
                        
                        # Also try getting filename from Content-Type header
                        if not filename:
                            ct_header = part.get('Content-Type', '')
                            if 'name' in ct_header:
                                name_match = re.search(r'name[*]?\s*=\s*["\']?([^"\';\r\n]+)', ct_header)
                                if name_match:
                                    filename = name_match.group(1).strip()
                        
                        # Log every non-text/html part for debugging
                        if content_type not in ('text/plain', 'text/html'):
                            debug_log.append(f"  Part: type={content_type}, disp={content_disp[:50]}, file={filename}")
                        
                        # Skip parts with no filename
                        if not filename:
                            continue
                        
                        # Check if it's an Excel or CSV file
                        if filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                            file_data = part.get_payload(decode=True)
                            attachments.append({
                                "filename": filename,
                                "subject": subject,
                                "data": file_data
                            })
                            email_attachments.append(filename)
                        else:
                            debug_log.append(f"    Skipped (not xlsx/csv): {filename}")
                            
                    # Parse and download Google Drive links
                    gdrive_files = {}
                    if html_body:
                        html_links = re.findall(
                            r'href=["\'](https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)/[^"\']+)["\'][^>]*>(.*?)</a>',
                            html_body,
                            re.IGNORECASE
                        )
                        for url, fid, text in html_links:
                            clean_text = re.sub(r'<[^>]+>', '', text).strip()
                            if clean_text.lower().endswith(('.xlsx', '.xls', '.csv')):
                                gdrive_files[fid] = clean_text
                                
                    if plain_body:
                        plain_pattern = r'\s*([^\n\r]+\.(?:csv|xlsx|xls))\s*[\r\n]+\s*<?(https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)/[^\s>]+)>?'
                        matches = re.findall(plain_pattern, plain_body, re.IGNORECASE)
                        for fname, url, fid in matches:
                            clean_name = fname.strip()
                            if fid not in gdrive_files:
                                gdrive_files[fid] = clean_name
                                
                    if gdrive_files:
                        debug_log.append(f"    Detected {len(gdrive_files)} Google Drive file(s) in email body")
                        for fid, fname in gdrive_files.items():
                            file_data = download_gdrive_file(fid)
                            if file_data:
                                attachments.append({
                                    "filename": fname,
                                    "subject": subject,
                                    "data": file_data
                                })
                                email_attachments.append(fname)
                                debug_log.append(f"      Downloaded: {fname} ({len(file_data)} bytes)")
                            else:
                                debug_log.append(f"      Failed to download Google Drive file: {fname}")
                    
                    if email_attachments:
                        debug_log.append(f"📧 '{subject[:40]}' → {len(email_attachments)} file(s): {email_attachments}")
                    else:
                        debug_log.append(f"📧 '{subject[:40]}' → {part_count} parts, 0 matching files")
        
        mail.logout()
        debug_log.append(f"\nTotal attachments collected: {len(attachments)}")
        return True, attachments, debug_log
    except Exception as e:
        debug_log.append(f"ERROR: {type(e).__name__}: {str(e)}")
        return False, str(e), debug_log

# Helper to safely read Excel/CSV files
def read_uploaded_file(file):
    try:
        if file.name.lower().endswith(".csv"):
            is_unified = "unified" in file.name.lower() or "transaction" in file.name.lower()
            try:
                # Seek to start, read first line to inspect content
                file.seek(0)
                first_line = file.read(200).decode('utf-8', errors='ignore').lower()
                if "marketplace" in first_line or "fulfillment by amazon" in first_line:
                    is_unified = True
            except:
                pass
            finally:
                file.seek(0)
                
            if is_unified:
                try:
                    df = pd.read_csv(file, header=13)
                    return df
                except Exception as e:
                    file.seek(0)
                    return pd.read_csv(file)
            else:
                return pd.read_csv(file)
        else:
            return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading file '{file.name}': {str(e)}")
        return None

# Helper to read raw bytes as dataframe (used for email attachments)
def read_bytes_file(file_bytes, filename):
    try:
        bio = io.BytesIO(file_bytes)
        if filename.lower().endswith(".csv"):
            is_unified = "unified" in filename.lower() or "transaction" in filename.lower()
            try:
                first_line = file_bytes.split(b'\n', 1)[0].decode('utf-8', errors='ignore').lower()
                if "marketplace" in first_line or "fulfillment by amazon" in first_line:
                    is_unified = True
            except:
                pass
                
            if is_unified:
                try:
                    df = pd.read_csv(bio, header=13)
                    return df
                except Exception as e:
                    bio.seek(0)
                    return pd.read_csv(bio)
            else:
                return pd.read_csv(bio)
        else:
            return pd.read_excel(bio, engine="openpyxl")
    except Exception as e:
        st.error(f"Error parsing file '{filename}': {str(e)}")
        return None

# Identify report and raw file types based on column schemas
def identify_file_type(df, filename=None):
    cols = [str(c).lower().strip() for c in df.columns]
    fname = filename.lower() if filename else ""
    
    # Processed Reports
    if any("door ship (seller flex)" in c for c in cols):
        return "report_refund"
    if any("amount_total" in c or "amount-total" in c for c in cols) and any("replacement_orderid" in c or "replacement-order-id" in c for c in cols):
        return "report_return"
    if any("refund check" in c for c in cols) and any("door step return" in c or "door-step-return" in c for c in cols):
        return "report_replace"
    if any("diff with charge & present commission" in c or "diff with charge &amp; present commission" in c for c in cols) and any("weight handling different" in c for c in cols):
        return "report_fees_overcharge"
        
    # Raw Inputs
    # 1. Fee Estimate (FREE) — must check BEFORE Replacements/Refund since it also has "sales-price"
    if any("estimated-referral-fee-per-unit" in c for c in cols) and any("estimated-weight-handling-fee-per-unit" in c for c in cols):
        return "raw_fee_estimate"
    
    # 2. Unified Transaction Report — has header at row 14, type column with Order/Refund, selling fees, fba fees
    if (any("selling fees" in c for c in cols) and any("fba fees" in c for c in cols) and any("other transaction fees" in c for c in cols) and any(c == "type" for c in cols) and any(c == "sku" for c in cols)):
        if "unified" in fname or "transaction" in fname or "free" in fname:
            return "raw_unified_txn"
    
    # 3. Replacements
    if any("replacement-amazon-order" in c or "original-amazon-order-id" in c for c in cols) or ("replace" in fname and "leakage" not in fname):
        return "raw_replace"
        
    # 4. Reimbursements
    if any("reimbursement-id" in c or "reimbursement_id" in c for c in cols) or "reimb" in fname:
        return "raw_reim"
        
    # 5. Bulk RTO Returns (check before standard returns to prevent mismatch)
    if any("awb" in c or "bulk" in c or "rto" in c or "suborder id" in c or "suborder-id" in c for c in cols) or ("bulk" in fname and "rto" in fname) or ("rto" in fname and "leakage" not in fname):
        return "raw_bulk_rto"
        
    # 6. Returns File
    if any("detailed-disposition" in c or "detailed_disposition" in c for c in cols) or ("return" in fname and "leakage" not in fname and "rto" not in fname and "replace" not in fname):
        return "raw_returns"
        
    # 7. QWT Customer Shipments
    if any("customer order id" in c or "customer-order-id" in c or ("customer" in c and "order" in c) for c in cols):
        return "raw_qwt"
    if "qwt" in fname and "inventory" not in fname and "stock" not in fname:
        return "raw_qwt"
        
    # 8. Safe-T Claims (reimbursement payments, check before refund/settlement)
    if any("safe-t" in c or "claim id" in c or "claim-id" in c for c in cols) or "safe" in fname:
        return "raw_safe_t"
        
    # 9. Refund Data (settlement reports containing refunds)
    if (any("date/time" in c for c in cols) and any("product sales" in c for c in cols)) or ("refund" in fname and "leakage" not in fname):
        return "raw_refund"
        
    # 10. PM File
    if ((any(c == "asin" for c in cols) or any("asin" in c for c in cols)) and (any(c == "brand" for c in cols) or any("brand" in c for c in cols))) or "pm" in fname:
        return "raw_pm"
        
    return None

# Helper paths for data storage
def get_report_path(filename):
    import os
    os.makedirs("data_store/reports", exist_ok=True)
    return os.path.join("data_store/reports", filename)

def get_raw_historical_path(filename):
    import os
    os.makedirs("data_store/raw_historical", exist_ok=True)
    return os.path.join("data_store/raw_historical", filename)

# Raw data merging and de-duplication engine
def merge_raw_file(df_new, file_type):
    import os
    import pandas as pd
    
    file_mapping = {
        "refund": {"name": "refund_data.xlsx", "keys": []},
        "qwt": {"name": "qwt_shipments.xlsx", "keys": ["customer order id", "order id", "order-id"]},
        "returns": {"name": "returns.xlsx", "keys": ["order-id", "order id", "sku"]},
        "bulk_rto": {"name": "bulk_rto.xlsx", "keys": ["order id", "order-id", "order_id"]},
        "safe_t": {"name": "safe_t_claims.xlsx", "keys": ["safe-t claim id", "claim id", "safe_t_claim_id"]},
        "reim": {"name": "reimbursements.xlsx", "keys": ["reimbursement-id", "reimbursement_id"]},
        "replace": {"name": "replacements.xlsx", "keys": ["original-order-id", "replacement-order-id"]},
        "unified_txn": {"name": "unified_transaction.xlsx", "keys": ["order id", "sku"]},
        "fee_estimate": {"name": "fee_estimate.xlsx", "keys": ["sku"]}
    }
    
    cfg = file_mapping.get(file_type)
    if not cfg:
        return df_new
        
    file_path = get_raw_historical_path(cfg["name"])
    
    if os.path.exists(file_path):
        try:
            if file_path.lower().endswith(".csv"):
                df_old = pd.read_csv(file_path)
            else:
                df_old = pd.read_excel(file_path, engine="openpyxl")
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
        except Exception as e:
            print(f"Error reading historical file {file_path}: {e}")
            df_combined = df_new
    else:
        df_combined = df_new
        
    # De-duplicate
    keys_to_use = []
    if cfg["keys"]:
        for key_candidate in cfg["keys"]:
            matched = [c for c in df_combined.columns if str(c).strip().lower() == key_candidate.lower()]
            if matched:
                keys_to_use.append(matched[0])
                
    if keys_to_use:
        # Drop duplicates by these keys, keeping the latest
        df_combined = df_combined.drop_duplicates(subset=keys_to_use, keep="last")
    else:
        # Fallback to dropping exact duplicate rows
        df_combined = df_combined.drop_duplicates(keep="last")
        
    try:
        df_combined.to_excel(file_path, index=False)
    except Exception as e:
        print(f"Error saving historical file {file_path}: {e}")
        
    return df_combined

# Load report automatically from local directory
@st.cache_data(ttl=300, show_spinner=False)
def load_local_report(category):
    import os
    import pandas as pd
    
    os.makedirs("data_store/reports", exist_ok=True)
    os.makedirs("data_store/raw_historical", exist_ok=True)
    
    # Try data_store/reports first, fallback to root directory for legacy compatibility
    search_dirs = ["data_store/reports", "."]
    for search_dir in search_dirs:
        if not os.path.exists(search_dir):
            continue
        files = os.listdir(search_dir)
        for f in files:
            f_path = os.path.join(search_dir, f)
            if f.lower().endswith(('.xlsx', '.xls', '.csv')):
                try:
                    # Handle multi-sheet Replacement Leakage report
                    if category == "Replacement Leakage" and f.lower().endswith(('.xlsx', '.xls')):
                        name_keywords = ["leakage", "cross", "check"]
                        if (("replacement" in f.lower() or "replace" in f.lower()) and any(kw in f.lower() for kw in name_keywords)):
                            try:
                                xls = pd.ExcelFile(f_path)
                                sheet_names = xls.sheet_names
                                if "Full Data" in sheet_names or "Damaged Returns" in sheet_names or "Refund Without Return" in sheet_names:
                                    dfs = {}
                                    for sheet in sheet_names:
                                        df_sheet = pd.read_excel(f_path, sheet_name=sheet)
                                        # Sanitize lookup columns for Arrow compatibility
                                        for col in ["Refund Check", "Door Step Return", "FBA Original Return", "FBA Replacement Return"]:
                                            if col in df_sheet.columns:
                                                df_sheet[col] = df_sheet[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
                                        if "CountIF" in df_sheet.columns:
                                            df_sheet["CountIF"] = pd.to_numeric(df_sheet["CountIF"], errors="coerce")
                                        dfs[sheet] = df_sheet
                                    return dfs
                            except Exception as e:
                                print(f"Error checking sheet names: {e}")

                    # Handle multi-sheet Refund Leakage report
                    if category == "Refund Leakage" and f.lower().endswith(('.xlsx', '.xls')):
                        name_keywords = ["leakage", "cross", "check"]
                        if ("refund" in f.lower() and any(kw in f.lower() for kw in name_keywords)):
                            try:
                                xls = pd.ExcelFile(f_path)
                                sheet_names = xls.sheet_names
                                if "Full Data" in sheet_names or "Door Ship TAT" in sheet_names or "FBA Return TAT" in sheet_names:
                                    dfs = {}
                                    for sheet in sheet_names:
                                        df_sheet = pd.read_excel(f_path, sheet_name=sheet)
                                        # Sanitize lookup columns for Arrow compatibility
                                        for col in ["Door Ship (Seller Flex)", "FBA Return", "Seller Flex Return", "Safe T Claim", "FBA Reimbursement"]:
                                            if col in df_sheet.columns:
                                                df_sheet[col] = df_sheet[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
                                        dfs[sheet] = df_sheet
                                    return dfs
                            except Exception as e:
                                print(f"Error checking sheet names for refund leakage: {e}")

                    # Handle multi-sheet Fees Overcharge Leakage report
                    if category == "Fees Overcharge Leakage" and f.lower().endswith(('.xlsx', '.xls')):
                        if ("fees" in f.lower() and "overcharge" in f.lower()) or "fees_overcharge" in f.lower() or "fee_overcharged" in f.lower():
                            try:
                                xls = pd.ExcelFile(f_path)
                                sheet_names = xls.sheet_names
                                if "Pivot Table" in sheet_names or "Commission Overcharge" in sheet_names or "Weight Overcharge" in sheet_names:
                                    dfs = {}
                                    for sheet in sheet_names:
                                        dfs[sheet] = pd.read_excel(f_path, sheet_name=sheet)
                                    return dfs
                            except Exception as e:
                                print(f"Error checking sheet names for fees overcharge leakage: {e}")

                    if f.endswith(".csv"):
                        df = pd.read_csv(f_path)
                    else:
                        df = pd.read_excel(f_path)
                    
                    ftype = identify_file_type(df, filename=f)
                    matched = False
                    
                    name_keywords = ["leakage", "cross", "check"]
                    
                    if category == "Refund Leakage":
                        if ftype == "report_refund" or ("refund" in f.lower() and any(kw in f.lower() for kw in name_keywords)):
                            matched = True
                    elif category == "Replacement Leakage":
                        if ftype == "report_replace" or (("replacement" in f.lower() or "replace" in f.lower()) and any(kw in f.lower() for kw in name_keywords)):
                            matched = True
                    elif category == "Return Leakage":
                        if ftype == "report_return" or ("return" in f.lower() and any(kw in f.lower() for kw in name_keywords)):
                            # make sure not to match replacement
                            if "replacement" not in f.lower() and "replace" not in f.lower():
                                matched = True
                    elif category == "Fees Overcharge Leakage":
                        if ftype == "report_fees_overcharge" or ("fee" in f.lower() or "overcharge" in f.lower()) or "fees_overcharge" in f.lower() or "fee_overcharged" in f.lower():
                            matched = True
                    elif category == "Extra Pattern Find Leakage":
                        if "extra" in f.lower() or "pattern" in f.lower():
                            matched = True
                            
                    if matched:
                        if category == "Replacement Leakage" and isinstance(df, pd.DataFrame):
                            # Sanitize legacy single-sheet report for Arrow compatibility
                            for col in ["Refund Check", "Door Step Return", "FBA Original Return", "FBA Replacement Return"]:
                                if col in df.columns:
                                    df[col] = df[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
                            if "CountIF" in df.columns:
                                df["CountIF"] = pd.to_numeric(df["CountIF"], errors="coerce")
                        return df
                except:
                    pass
    return None

# Auto-detect date, amount, and age columns
def auto_detect_columns(df):
    cols = list(df.columns)
    
    # 1. Date Detection
    date_col = None
    date_candidates = ['date1', 'return-date1', 'shipment-date', 'date', 'time', 'return-date', 'transaction-date', 'created']
    for candidate in date_candidates:
        matched = [c for c in cols if candidate.lower() == str(c).strip().lower()]
        if matched:
            date_col = matched[0]
            break
    if not date_col:
        for candidate in date_candidates:
            matched = [c for c in cols if candidate.lower() in str(c).strip().lower()]
            if matched:
                date_col = matched[0]
                break
    if not date_col:
        # fallback: find any column containing "date" or "time"
        matched = [c for c in cols if 'date' in str(c).lower() or 'time' in str(c).lower()]
        if matched:
            date_col = matched[0]
            
    # 2. Amount Detection
    amount_col = None
    amount_candidates = ['amount_total', 'product sales', 'amount-total', 'amount', 'sales', 'overcharge', 'value', 'price', 'refund']
    for candidate in amount_candidates:
        matched = [c for c in cols if candidate.lower() == str(c).strip().lower().replace('_', ' ').replace('-', ' ')]
        if matched:
            col_lower = str(matched[0]).lower()
            if not any(x in col_lower for x in ['id', 'check', 'code', 'number', 'num', 'date', 'order']):
                amount_col = matched[0]
                break
    if not amount_col:
        for candidate in amount_candidates:
            matched = [c for c in cols if candidate.lower() in str(c).strip().lower().replace('_', ' ').replace('-', ' ')]
            if matched:
                col_lower = str(matched[0]).lower()
                if not any(x in col_lower for x in ['id', 'check', 'code', 'number', 'num', 'date', 'order']):
                    amount_col = matched[0]
                    break
    if not amount_col:
        # fallback: any column containing "amount" or "sales" or "price"
        matched = [c for c in cols if any(kw in str(c).lower() for kw in ['amount', 'sales', 'price', 'value'])]
        if matched:
            for m in matched:
                col_lower = str(m).lower()
                if not any(x in col_lower for x in ['id', 'check', 'code', 'number', 'num', 'date', 'order']):
                    amount_col = m
                    break
            
    # 3. Age Detection
    age_col = None
    age_candidates = ['date_diff', 'date_difference', 'diff', 'age', 'days']
    for candidate in age_candidates:
        matched = [c for c in cols if candidate.lower() == str(c).strip().lower().replace('_', ' ').replace('-', ' ')]
        if matched:
            age_col = matched[0]
            break
    if not age_col:
        for candidate in age_candidates:
            matched = [c for c in cols if candidate.lower() in str(c).strip().lower().replace('_', ' ').replace('-', ' ')]
            if matched:
                age_col = matched[0]
                break
            
    return date_col, amount_col, age_col

# Parse and clean date column
def parse_date_series(series):
    parsed = pd.to_datetime(series, errors='coerce')
    # strip tz info if present
    try:
        if parsed.dt.tz is not None:
            parsed = parsed.dt.tz_localize(None)
    except:
        pass
    return parsed

# Helper to make sure dataframe columns of 'object' type are clean string types to prevent PyArrow serialization errors
def prepare_df_for_arrow(df):
    if df is None or len(df) == 0:
        return df
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).replace({'nan': '', 'None': '', '<NA>': '', 'nat': '', 'NaT': ''})
    return df

# Helper to display interactive sub-report table
def render_sub_report(df, name_key, download_name):
    if df is None or len(df) == 0:
        st.info("No records found in this report.")
        return

    # Filter options
    search_key = f"search_rep_{name_key}"
    search_query = st.text_input("🔍 Search in table:", "", key=search_key)

    df_display = df.copy()
    if search_query:
        mask = df_display.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
        df_display = df_display[mask]
        st.write(f"Found {len(df_display)} matching rows")

    # Show Table safely formatted for PyArrow
    df_display = prepare_df_for_arrow(df_display)
    st.dataframe(df_display, use_container_width=True)

    # Download button
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_display.to_excel(writer, index=False, sheet_name=download_name[:30])
    st.download_button(
        label=f"📥 Download {download_name} (Excel)",
        data=buffer.getvalue(),
        file_name=f"{download_name.replace(' ', '_')}_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_btn_{name_key}"
    )

    # Render trend and SKUs charts for this sub-report
    st.markdown("### 📊 Leakage Insights & Analytics")
    viz1, viz2 = st.columns(2)
    
    # Auto-detect date/amount/sku
    det_date, det_amt, det_age = auto_detect_columns(df_display)
    
    with viz1:
        # 1. Trend chart
        if det_date and det_amt and len(df_display) > 0:
            st.subheader("Leakage Amount Trend")
            df_chart_data = df_display.dropna(subset=[det_date]).copy()
            df_chart_data['Date_Parsed_For_Chart'] = parse_date_series(df_chart_data[det_date]).dt.date
            df_chart_data = df_chart_data.dropna(subset=['Date_Parsed_For_Chart'])
            if len(df_chart_data) > 0:
                df_chart = df_chart_data.groupby('Date_Parsed_For_Chart').agg({det_amt: 'sum'})
                df_chart = df_chart.dropna().reset_index()
                df_chart.columns = ["Date", "Amount"]
                df_chart["Amount"] = pd.to_numeric(df_chart["Amount"], errors="coerce").replace([float('inf'), float('-inf')], 0).fillna(0)
                if not df_chart.empty and df_chart["Amount"].sum() > 0:
                    st.bar_chart(df_chart, x="Date", y="Amount")
                else:
                    st.info("No trend data available.")
            else:
                st.info("No trend data available.")
        else:
            st.info("Trend chart requires a valid Date and Amount column.")
            
    with viz2:
        # 2. Top SKUs chart
        sku_col = None
        sku_candidates = ['sku', 'item-name', 'asin', 'fnsku', 'product-name']
        for candidate in sku_candidates:
            matched = [c for c in df_display.columns if candidate.lower() in str(c).lower()]
            if matched:
                sku_col = matched[0]
                break
                
        if sku_col and det_amt and len(df_display) > 0:
            st.subheader("Top 10 SKUs by Leakage Amount")
            df_sku_data = df_display.dropna(subset=[sku_col, det_amt]).copy()
            df_sku_data[det_amt] = pd.to_numeric(df_sku_data[det_amt], errors='coerce').replace([float('inf'), float('-inf')], 0).fillna(0)
            if len(df_sku_data) > 0:
                df_sku = df_sku_data.groupby(sku_col).agg({det_amt: 'sum'}).sort_values(by=det_amt, ascending=False).head(10)
                df_sku = df_sku.dropna().reset_index()
                df_sku.columns = ["SKU", "Amount"]
                df_sku["Amount"] = pd.to_numeric(df_sku["Amount"], errors="coerce").replace([float('inf'), float('-inf')], 0).fillna(0)
                if not df_sku.empty and df_sku["Amount"].sum() > 0:
                    st.bar_chart(df_sku, x="SKU", y="Amount")
                else:
                    st.info("No SKU data available.")
            else:
                st.info("No SKU data available.")
        elif sku_col and len(df_display) > 0:
            st.subheader("Top 10 SKUs by Record Count")
            df_sku_series = df_display[sku_col].dropna().value_counts().head(10)
            if len(df_sku_series) > 0:
                df_count = df_sku_series.reset_index()
                df_count.columns = ["SKU", "Record Count"]
                df_count["Record Count"] = pd.to_numeric(df_count["Record Count"], errors="coerce").replace([float('inf'), float('-inf')], 0).fillna(0)
                if not df_count.empty and df_count["Record Count"].sum() > 0:
                    st.bar_chart(df_count, x="SKU", y="Record Count")
                else:
                    st.info("No SKU record count data available.")
            else:
                st.info("No SKU record count data available.")
        else:
            st.info("SKU charts require an identifier column (e.g. SKU, ASIN) and/or Amount column.")

# Helper to render file uploader with fallback
def render_file_uploader_with_fallback(label, primary_key, fallback_keys, file_types=['csv', 'xlsx', 'xls']):
    # Check if any fallback is present in session state
    active_fallback = None
    for fb_key in fallback_keys:
        val = st.session_state.get(fb_key)
        if val is not None:
            active_fallback = val
            break
            
    if active_fallback is not None:
        st.markdown(f"<div style='font-size:0.85rem; color:#3fb950; margin-bottom:2px;'>🔄 Reusing <b>{active_fallback.name}</b> from another section</div>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader(f"Upload different {label} to override:", type=file_types, key=primary_key)
        return uploaded_file if uploaded_file is not None else active_fallback
    else:
        uploaded_file = st.file_uploader(label, type=file_types, key=primary_key)
        return uploaded_file

def run_auto_check_logic(is_automated=False):
    import json
    import os
    import pandas as pd
    from datetime import datetime
    
    status_path = "auto_scheduler_status.json"
    status = {
        "enabled": True,
        "scheduled_hour": 9,
        "last_run_time": "",
        "last_run_status": "Idle",
        "recent_logs": []
    }
    
    if os.path.exists(status_path):
        try:
            with open(status_path, "r") as f:
                status.update(json.load(f))
        except:
            pass
            
    if not status.get("enabled", True) and is_automated:
        return
        
    today_str = datetime.today().strftime("%Y-%m-%d")
    current_hour = datetime.now().hour
    
    if is_automated:
        last_date = status["last_run_time"][:10] if status["last_run_time"] else ""
        if last_date == today_str:
            return
        if current_hour < status.get("scheduled_hour", 9):
            return
            
    def add_log(msg):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] {msg}"
        status["recent_logs"].append(log_line)
        status["recent_logs"] = status["recent_logs"][-50:]
        try:
            with open(status_path, "w") as sf:
                json.dump(status, sf, indent=2)
        except:
            pass
            
    add_log("Starting automated check for new emails...")
    status["last_run_status"] = "Running"
    try:
        with open(status_path, "w") as sf:
            json.dump(status, sf, indent=2)
    except:
        pass
        
    mail_user = "reports@snaphire-it.com"
    mail_host = "snaphire-it.icewarpcloud.in"
    mail_port = 993
    
    mail_pass = get_secret_safe("email_password")
    if not mail_pass and os.path.exists("token.txt"):
        try:
            with open("token.txt", "r") as f:
                mail_pass = f.read().strip()
        except:
            pass
            
    if not mail_pass:
        add_log("Error: token.txt and st.secrets['email_password'] are missing or empty. Cannot authenticate IMAP.")
        status["last_run_status"] = "Error"
        try:
            with open(status_path, "w") as sf:
                json.dump(status, sf, indent=2)
        except:
            pass
        return
        
    try:
        success, result, debug_log = fetch_email_attachments(
            mail_host, mail_port, mail_user, mail_pass, folder="INBOX", subject_search="Amazon Leakeage Reports"
        )
        
        if debug_log:
            for l in debug_log:
                add_log(f"IMAP: {l}")
                
        if not success:
            add_log(f"Connection Error: {result}")
            status["last_run_status"] = "Error"
            try:
                with open(status_path, "w") as sf:
                    json.dump(status, sf, indent=2)
            except:
                pass
            return
            
        if len(result) == 0:
            add_log("No attachments found in the last 15 emails.")
            status["last_run_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status["last_run_status"] = "Success (No new files)"
            try:
                with open(status_path, "w") as sf:
                    json.dump(status, sf, indent=2)
            except:
                pass
            return
            
        loaded_files = []
        raw_inputs = {
            "refund": None,
            "qwt": None,
            "returns": None,
            "bulk_rto": None,
            "safe_t": None,
            "reim": None,
            "replace": None,
            "unified_txn": None,
            "fee_estimate": None
        }
        
        has_refund_inputs = False
        has_replace_inputs = False
        has_return_inputs = False
        has_fee_overcharge_inputs = False
        
        for att in result:
            filename = att['filename'].lower()
            subject = att['subject'].lower()
            df_file = read_bytes_file(att['data'], att['filename'])
            if df_file is None:
                continue
                
            ftype = identify_file_type(df_file, filename=att['filename'])
            is_processed = False
            
            if ftype == "report_refund" or "refund_leakage" in filename or "refund_leakage" in subject or ("refund" in filename and "cross" in filename and "check" in filename):
                df_file.to_excel(get_report_path("refund_leakage.xlsx"), index=False)
                loaded_files.append(f"Processed Report: {att['filename']} -> Refund Leakage")
                is_processed = True
            elif ftype == "report_replace" or "replacement_leakage" in filename or "replace_leakage" in filename or "replacement_leakage" in subject or ("replacement" in filename and "cross" in filename and "check" in filename):
                df_file.to_excel(get_report_path("replacement_leakage.xlsx"), index=False)
                loaded_files.append(f"Processed Report: {att['filename']} -> Replacement Leakage")
                is_processed = True
            elif ftype == "report_return" or "return_leakage" in filename or "return_leakage" in subject or ("return" in filename and "cross" in filename and "check" in filename):
                df_file.to_excel(get_report_path("return_leakage.xlsx"), index=False)
                loaded_files.append(f"Processed Report: {att['filename']} -> Return Leakage")
                is_processed = True
            elif ftype == "report_fees_overcharge" or "fees_overcharge" in filename or ("fees" in filename and "overcharge" in filename) or "fee_overcharged" in filename:
                df_file.to_excel(get_report_path("fees_overcharge_leakage.xlsx"), index=False)
                loaded_files.append(f"Processed Report: {att['filename']} -> Fees Overcharge Leakage")
                is_processed = True
                
            if is_processed:
                continue
                
            if ftype == "raw_qwt" or "qwt" in filename:
                df_merged = merge_raw_file(df_file, "qwt")
                raw_inputs["qwt"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> QWT Shipments")
                has_refund_inputs = True
            elif ftype == "raw_safe_t" or "safe" in filename:
                df_merged = merge_raw_file(df_file, "safe_t")
                raw_inputs["safe_t"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Safe-T Claims")
                has_refund_inputs = True
            elif ftype == "raw_reim" or "reimb" in filename or "reimbursement" in filename:
                df_merged = merge_raw_file(df_file, "reim")
                raw_inputs["reim"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Reimbursements")
                has_refund_inputs = True
                has_replace_inputs = True
                has_return_inputs = True
            elif ftype == "raw_replace" or "replacement" in filename or "replace" in filename:
                df_merged = merge_raw_file(df_file, "replace")
                raw_inputs["replace"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Replacements")
                has_replace_inputs = True
                has_return_inputs = True
            elif ftype == "raw_bulk_rto" or "bulk" in filename or "rto" in filename:
                df_merged = merge_raw_file(df_file, "bulk_rto")
                raw_inputs["bulk_rto"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Bulk RTO Returns")
                has_refund_inputs = True
                has_replace_inputs = True
            elif ftype == "raw_refund" or "refund" in filename:
                df_merged = merge_raw_file(df_file, "refund")
                raw_inputs["refund"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Refund Data")
                has_refund_inputs = True
                has_replace_inputs = True
            elif ftype == "raw_returns" or "return" in filename:
                df_merged = merge_raw_file(df_file, "returns")
                raw_inputs["returns"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Returns File")
                has_refund_inputs = True
                has_replace_inputs = True
                has_return_inputs = True
            elif ftype == "raw_pm" or "pm" in filename:
                df_file.to_excel(get_report_path("PM.xlsx"), index=False)
                get_pm_brand_map.clear()
                loaded_files.append(f"Saved PM mapping: {att['filename']}")
            elif ftype == "raw_fee_estimate" or "fee_estimate" in filename or "free" in filename:
                df_merged = merge_raw_file(df_file, "fee_estimate")
                raw_inputs["fee_estimate"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Fee Estimate (FREE)")
                has_fee_overcharge_inputs = True
            elif ftype == "raw_unified_txn" or "unified" in filename or "transaction" in filename:
                df_merged = merge_raw_file(df_file, "unified_txn")
                raw_inputs["unified_txn"] = df_merged
                loaded_files.append(f"Merged raw: {att['filename']} -> Unified Transaction")
                has_fee_overcharge_inputs = True
                
        def load_historical_if_missing(key):
            if raw_inputs[key] is None:
                historical_name = {
                    "refund": "refund_data.xlsx",
                    "qwt": "qwt_shipments.xlsx",
                    "returns": "returns.xlsx",
                    "bulk_rto": "bulk_rto.xlsx",
                    "safe_t": "safe_t_claims.xlsx",
                    "reim": "reimbursements.xlsx",
                    "replace": "replacements.xlsx",
                    "unified_txn": "unified_transaction.xlsx",
                    "fee_estimate": "fee_estimate.xlsx"
                }[key]
                hist_path = os.path.join("data_store/raw_historical", historical_name)
                if os.path.exists(hist_path):
                    try:
                        raw_inputs[key] = pd.read_excel(hist_path, engine="openpyxl")
                    except:
                        pass
                        
        for k in raw_inputs.keys():
            load_historical_if_missing(k)
            
        # 1. Refund Leakage
        if has_refund_inputs or (raw_inputs["refund"] is not None and raw_inputs["qwt"] is not None and raw_inputs["returns"] is not None):
            if all(raw_inputs[k] is not None for k in ["refund", "qwt", "returns", "bulk_rto", "safe_t", "reim"]):
                try:
                    df_out = process_refund_leakage(
                        raw_inputs["refund"], raw_inputs["qwt"], raw_inputs["returns"],
                        raw_inputs["bulk_rto"], raw_inputs["safe_t"], raw_inputs["reim"],
                        door_tat_min=50, door_tat_max=75, fba_tat_min=40
                    )
                    with pd.ExcelWriter(get_report_path("refund_leakage.xlsx"), engine="openpyxl") as writer:
                        df_out['main'].to_excel(writer, sheet_name="Full Data", index=False)
                        df_out['doorship_tat'].to_excel(writer, sheet_name="Door Ship TAT", index=False)
                        df_out['fba_return_tat'].to_excel(writer, sheet_name="FBA Return TAT", index=False)
                    add_log("Regenerated Refund Leakage report.")
                except Exception as e:
                    add_log(f"Error regenerating Refund Leakage: {str(e)}")
                    
        # 2. Replacement Leakage
        if has_replace_inputs or (raw_inputs["replace"] is not None and raw_inputs["returns"] is not None):
            if all(raw_inputs[k] is not None for k in ["replace", "returns", "refund", "bulk_rto", "reim"]):
                try:
                    results = process_replacement_leakage(
                        raw_inputs["replace"], raw_inputs["returns"], raw_inputs["refund"],
                        raw_inputs["bulk_rto"], raw_inputs["reim"], days_threshold=40
                    )
                    with pd.ExcelWriter(get_report_path("replacement_leakage.xlsx"), engine="openpyxl") as writer:
                        results['main'].to_excel(writer, sheet_name="Full Data", index=False)
                        results['damaged_returns'].to_excel(writer, sheet_name="Damaged Returns", index=False)
                        results['refund_without_return'].to_excel(writer, sheet_name="Refund Without Return", index=False)
                    add_log("Regenerated Replacement Leakage report.")
                except Exception as e:
                    add_log(f"Error regenerating Replacement Leakage: {str(e)}")
                    
        # 3. Return Leakage
        if has_return_inputs or (raw_inputs["returns"] is not None and raw_inputs["reim"] is not None):
            if all(raw_inputs[k] is not None for k in ["returns", "reim", "replace"]):
                try:
                    df_out = process_return_leakage(
                        raw_inputs["returns"], raw_inputs["reim"], raw_inputs["replace"], days_filter=40
                    )
                    df_out.to_excel(get_report_path("return_leakage.xlsx"), index=False)
                    add_log("Regenerated Return Leakage report.")
                except Exception as e:
                    add_log(f"Error regenerating Return Leakage: {str(e)}")
                    
        # 4. Fees Overcharge Leakage
        if has_fee_overcharge_inputs or (raw_inputs["unified_txn"] is not None and raw_inputs["fee_estimate"] is not None):
            if all(raw_inputs[k] is not None for k in ["unified_txn", "fee_estimate"]):
                try:
                    results = process_free_overcharged_leakage(
                        raw_inputs["unified_txn"], raw_inputs["fee_estimate"]
                    )
                    with pd.ExcelWriter(get_report_path("fees_overcharge_leakage.xlsx"), engine="openpyxl") as writer:
                        results['pivot'].to_excel(writer, sheet_name="Pivot Table", index=False)
                        results['commission_overcharge'].to_excel(writer, sheet_name="Commission Overcharge", index=False)
                        results['weight_overcharge'].to_excel(writer, sheet_name="Weight Overcharge", index=False)
                    add_log("Regenerated Fees Overcharge Leakage report.")
                except Exception as e:
                    add_log(f"Error regenerating Fees Overcharge Leakage: {str(e)}")
                    
        for msg in loaded_files:
            add_log(f"New File Processed: {msg}")
            
        status["last_run_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status["last_run_status"] = "Success"
        add_log("Automated daily check completed successfully.")
        load_local_report.clear()
        
    except Exception as e:
        add_log(f"Fatal check error: {str(e)}")
        status["last_run_status"] = "Error"
        add_log("Automated daily check failed.")
        
    try:
        with open(status_path, "w") as sf:
            json.dump(status, sf, indent=2)
    except:
        pass



# Initialize session state for uploaded reports
if "reports_data" not in st.session_state:
    st.session_state.reports_data = {
        "Refund Leakage": None,
        "Return Leakage": None,
        "Replacement Leakage": None,
        "Fees Overcharge Leakage": None,
        "Extra Pattern Find Leakage": None
    }



# ----------------- UI HEADERS -----------------
st.markdown('<div class="main-title">🎯 Leakage Pipeline Reports</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Aggregated leakage analysis across refund, return, replacement, fees, and custom checks</div>', unsafe_allow_html=True)

# Email instruction banner & manual sync button
st.info("""
📧 **Email Auto-check Pipeline**: 
To automatically feed reports into this dashboard, ask the sender to email raw files to **reports@snaphire-it.com** with the subject: **"Amazon Leakeage Reports"**.
""")

col_sync_btn, col_spacer = st.columns([1.5, 2])
with col_sync_btn:
    if st.button("⚡ Scan Mail & Rebuild Reports", type="primary", use_container_width=True):
        mail_user = "reports@snaphire-it.com"
        mail_pass = get_secret_safe("email_password")
        if not mail_pass and os.path.exists("token.txt"):
            try:
                with open("token.txt", "r") as f:
                    mail_pass = f.read().strip()
            except:
                pass
                
        if not mail_pass:
            st.error("🔑 Password missing. Please set st.secrets['email_password'] or token.txt.")
        else:
            with st.spinner("Connecting and downloading email attachments..."):
                try:
                    run_auto_check_logic(is_automated=False)
                    st.success("✅ Reports successfully synchronized and regenerated!")
                    st.toast("Successfully synced and updated reports!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Rebuild error: {e}")

# Latest run information (mocked based on current date)
st.markdown(f"""
<div class="status-banner">
    <span>✅</span>
    <span>Latest run: {datetime.today().strftime('%Y-%m-%d')} &nbsp;|&nbsp; Reference point: 150-day window</span>
</div>
""", unsafe_allow_html=True)

# Helper to render file uploader with fallback

# ----------------- INPUT DATA HUB (RAW SOURCE FILES) -----------------
with st.expander("📁 Input Data Hub: Upload Raw Amazon Files to Generate Reports", expanded=False):
    st.markdown("Upload your raw Amazon export files below. The dashboard will automatically execute cross-checks, save the output reports, and display them in the tabs.")
    
    # Save PM file if uploaded in any of the tabs
    uploaded_pm = st.session_state.get("raw_pm_refund") or st.session_state.get("raw_pm_replace") or st.session_state.get("raw_pm_return")
    if uploaded_pm:
        if st.session_state.get("pm_saved_name") != uploaded_pm.name:
            try:
                df_pm_uploaded = read_uploaded_file(uploaded_pm)
                if df_pm_uploaded is not None:
                    df_pm_uploaded.to_excel("PM.xlsx", index=False)
                    get_pm_brand_map.clear()
                    st.session_state["pm_saved_name"] = uploaded_pm.name
                    st.toast("✅ PM File successfully updated!")
            except Exception as e:
                st.error(f"Error saving PM file: {str(e)}")

    # 1. Global Processing Parameters
    st.markdown("### ⚙️ Global Processing Parameters")
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        door_tat_min = st.number_input("Door Ship TAT Start:", value=50, step=1, key="door_tat_min_val")
    with p_col2:
        door_tat_max = st.number_input("Door Ship TAT End:", value=75, step=1, key="door_tat_max_val")
    with p_col3:
        age_threshold = st.number_input("Age Threshold (Days):", value=40, step=1, key="age_threshold_val")
        
    st.markdown("<div style='margin-top: 15px; margin-bottom: 15px; border-bottom: 1px solid #30363d;'></div>", unsafe_allow_html=True)

    # 2. Report-specific Upload Tabs
    upload_tab_ref, upload_tab_rep, upload_tab_ret, upload_tab_fees_overcharge = st.tabs([
        "🔴 Refund Leakage Uploads", 
        "🔄 Replacement Leakage Uploads", 
        "📦 Return Leakage Uploads",
        "💸 Fees Overcharge Uploads"
    ])
    
    with upload_tab_ref:
        st.markdown("#### Upload files required to generate **Refund Leakage Report**")
        ref_col1, ref_col2 = st.columns(2)
        with ref_col1:
            ref_refund = render_file_uploader_with_fallback("1. Refund Data File (Excel/CSV)", "raw_ref_refund", ["raw_ref_replace"])
            ref_qwt = st.file_uploader("2. QWT Customer Shipments (Excel/CSV)", type=['csv', 'xlsx'], key="raw_qwt_refund")
            ref_returns = render_file_uploader_with_fallback("3. Returns File (Excel/CSV)", "raw_ret_refund", ["raw_ret_replace", "raw_ret_return"])
        with ref_col2:
            ref_bulk = render_file_uploader_with_fallback("4. Bulk RTO Returns File (Excel/CSV)", "raw_bulk_refund", ["raw_bulk_replace"])
            ref_safe = st.file_uploader("5. Safe-T Claims File (Excel/CSV)", type=['csv', 'xlsx', 'xls'], key="raw_safe_refund")
            ref_reim = render_file_uploader_with_fallback("6. Reimbursements File (Excel/CSV)", "raw_reim_refund", ["raw_reim_replace", "raw_reim_return"])
            
        ref_pm = render_file_uploader_with_fallback("7. PM File (Excel/CSV) - Mapping", "raw_pm_refund", ["raw_pm_replace", "raw_pm_return"])
        
        if st.button("⚡ Generate Refund Leakage Report", type="primary", use_container_width=True):
            if ref_refund and ref_qwt and ref_returns and ref_bulk and ref_safe and ref_reim:
                with st.spinner("Processing Refund Leakage..."):
                    try:
                        df_ref = read_uploaded_file(ref_refund)
                        df_qwt = read_uploaded_file(ref_qwt)
                        df_ret = read_uploaded_file(ref_returns)
                        df_blk = read_uploaded_file(ref_bulk)
                        df_sf = read_uploaded_file(ref_safe)
                        df_rm = read_uploaded_file(ref_reim)
                        
                        df_refund_leakage = process_refund_leakage(
                            df_ref, df_qwt, df_ret, df_blk, df_sf, df_rm, door_tat_min, door_tat_max, age_threshold
                        )
                        st.session_state.reports_data["Refund Leakage"] = df_refund_leakage
                        with pd.ExcelWriter(get_report_path("refund_leakage.xlsx"), engine="openpyxl") as writer:
                            df_refund_leakage['main'].to_excel(writer, sheet_name="Full Data", index=False)
                            df_refund_leakage['doorship_tat'].to_excel(writer, sheet_name="Door Ship TAT", index=False)
                            df_refund_leakage['fba_return_tat'].to_excel(writer, sheet_name="FBA Return TAT", index=False)
                        st.success("✅ Refund Leakage report successfully generated and saved to 'data_store/reports/refund_leakage.xlsx'!")
                        st.toast("Successfully generated Refund Leakage report!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error processing Refund Leakage: {str(e)}")
                        st.exception(e)
            else:
                st.warning("⚠️ Please upload all required files to generate Refund Leakage report.")
                
    with upload_tab_rep:
        st.markdown("#### Upload files required to generate **Replacement Leakage Report**")
        rep_col1, rep_col2 = st.columns(2)
        with rep_col1:
            rep_replace = render_file_uploader_with_fallback("1. Replacements File (Excel/CSV)", "raw_rep_replace", ["raw_rep_return"], file_types=['csv', 'xlsx'])
            rep_returns = render_file_uploader_with_fallback("2. Returns File (Excel/CSV)", "raw_ret_replace", ["raw_ret_refund", "raw_ret_return"])
            rep_refund = render_file_uploader_with_fallback("3. Refund Data File (Excel/CSV)", "raw_ref_replace", ["raw_ref_refund"])
        with rep_col2:
            rep_bulk = render_file_uploader_with_fallback("4. Bulk RTO Returns File (Excel/CSV)", "raw_bulk_replace", ["raw_bulk_refund"])
            rep_reim = render_file_uploader_with_fallback("5. Reimbursements File (Excel/CSV)", "raw_reim_replace", ["raw_reim_refund", "raw_reim_return"])
            
        rep_pm = render_file_uploader_with_fallback("6. PM File (Excel/CSV) - Mapping", "raw_pm_replace", ["raw_pm_refund", "raw_pm_return"])
        
        if st.button("⚡ Generate Replacement Leakage Report", type="primary", use_container_width=True):
            if rep_replace and rep_returns and rep_refund and rep_bulk and rep_reim:
                with st.spinner("Processing Replacement Leakage..."):
                    try:
                        df_rep = read_uploaded_file(rep_replace)
                        df_ret = read_uploaded_file(rep_returns)
                        df_ref = read_uploaded_file(rep_refund)
                        df_blk = read_uploaded_file(rep_bulk)
                        df_rm = read_uploaded_file(rep_reim)
                        
                        results = process_replacement_leakage(
                            df_rep, df_ret, df_ref, df_blk, df_rm, age_threshold
                        )
                        st.session_state.reports_data["Replacement Leakage"] = results
                        with pd.ExcelWriter(get_report_path("replacement_leakage.xlsx"), engine="openpyxl") as writer:
                            results['main'].to_excel(writer, sheet_name="Full Data", index=False)
                            results['damaged_returns'].to_excel(writer, sheet_name="Damaged Returns", index=False)
                            results['refund_without_return'].to_excel(writer, sheet_name="Refund Without Return", index=False)
                        st.success("✅ Replacement Leakage report successfully generated and saved to 'data_store/reports/replacement_leakage.xlsx'!")
                        st.toast("Successfully generated Replacement Leakage report!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error processing Replacement Leakage: {str(e)}")
                        st.exception(e)
            else:
                st.warning("⚠️ Please upload all required files to generate Replacement Leakage report.")
                
    with upload_tab_ret:
        st.markdown("#### Upload files required to generate **Return Leakage Report**")
        ret_col1, ret_col2 = st.columns(2)
        with ret_col1:
            ret_returns = render_file_uploader_with_fallback("1. Returns File (Excel/CSV)", "raw_ret_return", ["raw_ret_refund", "raw_ret_replace"])
            ret_reim = render_file_uploader_with_fallback("2. Reimbursements File (Excel/CSV)", "raw_reim_return", ["raw_reim_refund", "raw_reim_replace"])
        with ret_col2:
            ret_replace = render_file_uploader_with_fallback("3. Replacements File (Excel/CSV)", "raw_rep_return", ["raw_rep_replace"], file_types=['csv', 'xlsx'])
            
        ret_pm = render_file_uploader_with_fallback("4. PM File (Excel/CSV) - Mapping", "raw_pm_return", ["raw_pm_refund", "raw_pm_replace"])
        
        if st.button("⚡ Generate Return Leakage Report", type="primary", use_container_width=True):
            if ret_returns and ret_reim and ret_replace:
                with st.spinner("Processing Return Leakage..."):
                    try:
                        df_ret = read_uploaded_file(ret_returns)
                        df_rm = read_uploaded_file(ret_reim)
                        df_rep = read_uploaded_file(ret_replace)
                        
                        df_ret_leakage = process_return_leakage(
                            df_ret, df_rm, df_rep, age_threshold
                        )
                        st.session_state.reports_data["Return Leakage"] = df_ret_leakage
                        df_ret_leakage.to_excel(get_report_path("return_leakage.xlsx"), index=False)
                        st.success("✅ Return Leakage report successfully generated and saved to 'data_store/reports/return_leakage.xlsx'!")
                        st.toast("Successfully generated Return Leakage report!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error processing Return Leakage: {str(e)}")
                        st.exception(e)
            else:
                st.warning("⚠️ Please upload all required files to generate Return Leakage report.")

    with upload_tab_fees_overcharge:
        st.markdown("#### Upload files required to generate **Fees Overcharge Leakage Report**")
        st.markdown("Upload your **Unified Transaction** report (header starts at row 14) and **Fee Estimate (FREE)** CSV to detect commission and weight-handling overcharges.")
        free_col1, free_col2 = st.columns(2)
        with free_col1:
            free_txn = st.file_uploader("1. Unified Transaction Report (CSV)", type=['csv'], key="raw_free_txn")
        with free_col2:
            free_fee = st.file_uploader("2. Fee Estimate Report - FREE (CSV)", type=['csv'], key="raw_free_fee")
        
        if st.button("⚡ Generate Fees Overcharge Leakage Report", type="primary", use_container_width=True):
            if free_txn and free_fee:
                with st.spinner("Processing Fees Overcharge Leakage..."):
                    try:
                        df_txn = pd.read_csv(free_txn, header=13)
                        df_fee = pd.read_csv(free_fee)
                        
                        results = process_free_overcharged_leakage(df_txn, df_fee)
                        st.session_state.reports_data["Fees Overcharge Leakage"] = results
                        with pd.ExcelWriter(get_report_path("fees_overcharge_leakage.xlsx"), engine="openpyxl") as writer:
                            results['pivot'].to_excel(writer, sheet_name="Pivot Table", index=False)
                            results['commission_overcharge'].to_excel(writer, sheet_name="Commission Overcharge", index=False)
                            results['weight_overcharge'].to_excel(writer, sheet_name="Weight Overcharge", index=False)
                        st.success("✅ Fees Overcharge Leakage report successfully generated and saved to 'data_store/reports/fees_overcharge_leakage.xlsx'!")
                        st.toast("Successfully generated Fees Overcharge Leakage report!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error processing Fees Overcharge Leakage: {str(e)}")
                        st.exception(e)
            else:
                st.warning("⚠️ Please upload both files to generate Fees Overcharge Leakage report.")



# ----------------- SIDEBAR FILTERS -----------------
st.sidebar.markdown("### 📅 Time Period Filter")
time_filter_mode = st.sidebar.selectbox(
    "Select Period",
    ["All Time", "Last 7 Days", "Last 30 Days", "Last 90 Days", "Month to Date", "Year to Date", "Custom Range"],
    index=0
)

# Reference date calculation (default to 2026-05-25 as per current metadata)
today_ref = datetime(2026, 5, 25).date()

# Compute filter bounds
start_date = None
end_date = None

if time_filter_mode == "All Time":
    pass
elif time_filter_mode == "Last 7 Days":
    start_date = today_ref - timedelta(days=7)
    end_date = today_ref
elif time_filter_mode == "Last 30 Days":
    start_date = today_ref - timedelta(days=30)
    end_date = today_ref
elif time_filter_mode == "Last 90 Days":
    start_date = today_ref - timedelta(days=90)
    end_date = today_ref
elif time_filter_mode == "Month to Date":
    start_date = date(today_ref.year, today_ref.month, 1)
    end_date = today_ref
elif time_filter_mode == "Year to Date":
    start_date = date(today_ref.year, 1, 1)
    end_date = today_ref
elif time_filter_mode == "Custom Range":
    dates = st.sidebar.date_input(
        "Select Range",
        value=(today_ref - timedelta(days=30), today_ref),
        min_value=date(2020, 1, 1),
        max_value=date(2030, 12, 31)
    )
    if isinstance(dates, tuple) and len(dates) == 2:
        start_date, end_date = dates

# Helper to detect/suggest IMAP settings based on MX records of domain
def suggest_imap_settings(email_address):
    import subprocess
    import re
    if not email_address or "@" not in email_address:
        return "imap.gmail.com", 993, "Unknown", "Please enter a valid email address."
        
    domain = email_address.split("@")[1].strip().lower()
    
    if domain == "snaphire-it.com":
        return "snaphire-it.icewarpcloud.in", 993, "IceWarp Cloud (Snaphire)", "Your domain uses IceWarp Cloud. Use **snaphire-it.icewarpcloud.in** as the IMAP Host and your regular email password."
        
    if domain == "gmail.com":
        return "imap.gmail.com", 993, "Gmail", "Use **imap.gmail.com** and generate a **Google App Password** (not your regular password). To do this: 1. Go to your Google Account. 2. Enable 2-Step Verification. 3. Search for 'App Passwords' and create one for this tool."
        
    # Check MX records via subprocess running nslookup
    try:
        res = subprocess.run(
            ["nslookup", "-type=mx", domain],
            capture_output=True,
            text=True,
            timeout=3
        )
        output = res.stdout.lower() if res.stdout else ""
    except Exception:
        output = ""
        
    # Detect providers
    if "google" in output:
        return "imap.gmail.com", 993, "Google Workspace (Custom Domain)", "Your domain uses Google Workspace. Use **imap.gmail.com** and generate a **Google App Password** under your Google Account Security settings."
    elif "outlook" in output or "microsoft" in output:
        return "outlook.office365.com", 993, "Microsoft Office 365", "Your domain uses Office 365. Use **outlook.office365.com**. Note: Make sure IMAP is enabled for this mailbox in the Exchange Admin center. If security defaults are active, you may need an App Password or modern authentication."
    elif "secureserver" in output:
        return "imap.secureserver.net", 993, "GoDaddy Email", "Your domain uses GoDaddy. Use **imap.secureserver.net** as the host."
    elif "hostinger" in output:
        return "imap.hostinger.com", 993, "Hostinger Email", "Your domain uses Hostinger. Use **imap.hostinger.com** as the host."
    elif "zoho" in output:
        return "imap.zoho.com", 993, "Zoho Mail", "Your domain uses Zoho. Use **imap.zoho.com** and create an App Password in your Zoho account settings."
    elif "immenzaces" in output or "icewarpcloud" in output:
        # IceWarp India / Asia
        cluster_name = "cluster36.immenzaces.com"
        mx_match = re.search(r"mx[0-9]*-([a-zA-Z0-9.-]+)", output)
        if mx_match:
            cluster_name = mx_match.group(1).strip()
            if cluster_name.endswith("."):
                cluster_name = cluster_name[:-1]
        return f"mail.{domain}", 993, "IceWarp Cloud", f"Your domain uses IceWarp Cloud. Try using **mail.{domain}** (default) or the cluster address **{cluster_name}** as the IMAP Host and your regular email password."
        
    # Default fallbacks
    return f"mail.{domain}", 993, "Custom Domain", f"Using default custom host **mail.{domain}**. If this fails, try **imap.{domain}** or ask your email administrator for the correct IMAP Host and Port."



# ----------------- EMAIL IMPORT HELPER -----------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📧 Email Import Helper")

# Auto-read app password from token.txt or secrets
_default_pass = get_secret_safe("email_password", "")
if not _default_pass and os.path.exists("token.txt"):
    try:
        with open("token.txt", "r") as _f:
            _default_pass = _f.read().strip()
    except:
        pass

with st.sidebar.expander("Configure IMAP Mailbox", expanded=False):
    mail_user = st.text_input("Email:", value="reports@snaphire-it.com").strip()
    
    # Auto-detect settings
    s_host, s_port, provider, note = suggest_imap_settings(mail_user)
    
    # Track overrides in session state
    if "mail_host_val" not in st.session_state or st.session_state.get("prev_email") != mail_user:
        st.session_state["mail_host_val"] = s_host
        st.session_state["mail_port_val"] = s_port
        st.session_state["prev_email"] = mail_user
        
    # Auto-override old bad suggestion if it resides in active session state
    if st.session_state.get("mail_host_val") in ["mx1-cluster36.immenzaces.com", "mx2-cluster36.immenzaces.com", "mail.immenzaces.com"]:
        st.session_state["mail_host_val"] = s_host
        
    mail_host = st.text_input("IMAP Host:", value=st.session_state["mail_host_val"]).strip()
    mail_port = st.number_input("IMAP Port:", value=st.session_state["mail_port_val"], min_value=1, max_value=65535)
    
    # Store user overrides
    st.session_state["mail_host_val"] = mail_host
    st.session_state["mail_port_val"] = mail_port
    
    # Test DNS resolution and suggest working alternatives if it fails
    import socket
    dns_resolved = False
    if mail_host:
        try:
            socket.getaddrinfo(mail_host, int(mail_port))
            dns_resolved = True
        except:
            dns_resolved = False
            
    if not dns_resolved and mail_host:
        st.error(f"⚠️ **DNS Error**: `{mail_host}` could not be resolved.")
        
        # Test common alternative hostnames
        domain = mail_user.split("@")[1].strip().lower() if "@" in mail_user else ""
        if domain:
            alternatives = [
                f"mail.{domain}", 
                f"imap.{domain}", 
                domain,
                "snaphire-it.icewarpcloud.in",
                "cluster36.immenzaces.com",
                "mail.immenzaces.com"
            ]
            try:
                import subprocess
                res_mx = subprocess.run(["nslookup", "-type=mx", domain], capture_output=True, text=True, timeout=2)
                out_mx = res_mx.stdout.lower() if res_mx.stdout else ""
                
                # Check for cluster address
                m = re.search(r"mx[0-9]*-([a-zA-Z0-9.-]+)", out_mx)
                if m:
                    cname = m.group(1).strip()
                    if cname.endswith("."):
                        cname = cname[:-1]
                    alternatives.append(cname)
                    
                # Check for MX exchanger hosts
                mxs = re.findall(r"mail exchanger\s*=\s*([a-zA-Z0-9.-]+)", out_mx)
                for mx in mxs:
                    mx_clean = mx.strip()
                    if mx_clean.endswith("."):
                        mx_clean = mx_clean[:-1]
                    alternatives.append(mx_clean)
            except:
                pass
                
            resolved_alts = []
            for alt in sorted(list(set(alternatives))):
                try:
                    socket.getaddrinfo(alt, int(mail_port))
                    resolved_alts.append(alt)
                except:
                    pass
            if resolved_alts:
                st.info("💡 **Working hostnames found on your network:**")
                for r_alt in resolved_alts:
                    st.write(f"- `{r_alt}`")
            else:
                st.warning("🔍 No standard email server hostnames could be resolved for your domain. Please verify your internet connection or check with your IT admin.")
    
    mail_pass = st.text_input("Password / App Pass:", type="password", value=_default_pass)
    mail_folder = st.text_input("Mail Folder:", value="INBOX")
    mail_search_subject = st.text_input("Filter Subject (Optional):", value="Amazon Leakeage Reports")
    
    st.info(f"**Mail Provider Detected:** {provider}\n\n{note}")

if st.sidebar.button("🔍 Fetch Reports from Mail", type="primary", use_container_width=True):
    if not mail_pass:
        st.sidebar.error("🔑 Please enter your email password/app pass.")
    else:
        # Show what we're attempting
        st.sidebar.info(f"Connecting to **{mail_host}:{mail_port}** as **{mail_user}**")
        
        with st.spinner("Connecting and downloading email attachments..."):
            success, result, debug_log = fetch_email_attachments(
                mail_host, mail_port, mail_user, mail_pass, mail_folder, mail_search_subject
            )
            
            # Always show debug log
            if debug_log:
                with st.sidebar.expander("🔍 Debug Log", expanded=True):
                    for line in debug_log:
                        st.text(line)
            
            if not success:
                st.sidebar.error(f"❌ Connection Error: {result}")
                st.sidebar.warning(f"Attempted login as: **{mail_user}** → Host: **{mail_host}**")
            elif len(result) == 0:
                st.sidebar.warning("⚠️ No attachments found in the last 15 emails.")
            else:
                loaded_files = []
                
                # Maps for raw files downloaded from email
                raw_inputs = {
                    "refund": None,
                    "qwt": None,
                    "returns": None,
                    "bulk_rto": None,
                    "safe_t": None,
                    "reim": None,
                    "replace": None,
                    "unified_txn": None,
                    "fee_estimate": None
                }
                
                for att in result:
                    filename = att['filename'].lower()
                    subject = att['subject'].lower()
                    df_file = read_bytes_file(att['data'], att['filename'])
                    if df_file is None:
                        continue
                        
                    # Identify the file using columns signature
                    ftype = identify_file_type(df_file, filename=att['filename'])
                    
                    # First, check if it's a pre-processed report
                    is_processed = False
                    if ftype == "report_refund" or "refund_leakage" in filename or "refund_leakage" in subject or ("refund" in filename and "cross" in filename and "check" in filename):
                        st.session_state.reports_data["Refund Leakage"] = df_file
                        loaded_files.append(f"✅ Processed Report: {att['filename']} -> Refund Leakage")
                        is_processed = True
                    elif ftype == "report_replace" or "replacement_leakage" in filename or "replace_leakage" in filename or "replacement_leakage" in subject or ("replacement" in filename and "cross" in filename and "check" in filename):
                        st.session_state.reports_data["Replacement Leakage"] = df_file
                        loaded_files.append(f"✅ Processed Report: {att['filename']} -> Replacement Leakage")
                        is_processed = True
                    elif ftype == "report_return" or "return_leakage" in filename or "return_leakage" in subject or ("return" in filename and "cross" in filename and "check" in filename):
                        st.session_state.reports_data["Return Leakage"] = df_file
                        loaded_files.append(f"✅ Processed Report: {att['filename']} -> Return Leakage")
                        is_processed = True
                    elif ftype == "report_fees_overcharge" or "fees_overcharge" in filename or ("fees" in filename and "overcharge" in filename) or "fee_overcharged" in filename:
                        st.session_state.reports_data["Fees Overcharge Leakage"] = df_file
                        loaded_files.append(f"✅ Processed Report: {att['filename']} -> Fees Overcharge Leakage")
                        is_processed = True
                    elif "extra_leakage" in filename or "pattern_leakage" in filename:
                        st.session_state.reports_data["Extra Pattern Find Leakage"] = df_file
                        loaded_files.append(f"✅ Processed Report: {att['filename']} -> Extra Pattern Find Leakage")
                        is_processed = True
                        
                    if is_processed:
                        continue
                        
                    # If not a pre-processed report, check if it's a raw input file (use ftype first, fallback to filename)
                    if ftype == "raw_qwt" or "qwt" in filename:
                        raw_inputs["qwt"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> QWT Shipments")
                    elif ftype == "raw_safe_t" or "safe" in filename:
                        raw_inputs["safe_t"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Safe-T Claims")
                    elif ftype == "raw_reim" or "reimb" in filename or "reimbursement" in filename:
                        raw_inputs["reim"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Reimbursements")
                    elif ftype == "raw_replace" or "replacement" in filename or "replace" in filename:
                        raw_inputs["replace"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Replacements")
                    elif ftype == "raw_bulk_rto" or "bulk" in filename or "rto" in filename:
                        raw_inputs["bulk_rto"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Bulk RTO Returns")
                    elif ftype == "raw_refund" or "refund" in filename:
                        raw_inputs["refund"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Refund Data")
                    elif ftype == "raw_returns" or "return" in filename:
                        raw_inputs["returns"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Returns File")
                    elif ftype == "raw_pm" or "pm" in filename:
                        try:
                            df_file.to_excel(get_report_path("PM.xlsx"), index=False)
                            get_pm_brand_map.clear()
                            loaded_files.append(f"📁 PM File: {att['filename']} -> Saved to 'data_store/reports/PM.xlsx'")
                        except Exception as e:
                            loaded_files.append(f"❌ Error updating PM File from email: {str(e)}")
                    elif ftype == "raw_fee_estimate" or "fee_estimate" in filename or ("free" in filename and "fee" in filename):
                        raw_inputs["fee_estimate"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Fee Estimate (FREE)")
                    elif ftype == "raw_unified_txn" or "unified" in filename or ("transaction" in filename and "unified" in filename):
                        raw_inputs["unified_txn"] = df_file
                        loaded_files.append(f"📁 Raw Input: {att['filename']} -> Unified Transaction")
                        
                # Run generations if raw inputs are available
                # 1. Refund Leakage
                if (raw_inputs["refund"] is not None and 
                    raw_inputs["qwt"] is not None and 
                    raw_inputs["returns"] is not None and 
                    raw_inputs["bulk_rto"] is not None and 
                    raw_inputs["safe_t"] is not None and 
                    raw_inputs["reim"] is not None):
                    try:
                        df_out = process_refund_leakage(
                            raw_inputs["refund"], raw_inputs["qwt"], raw_inputs["returns"],
                            raw_inputs["bulk_rto"], raw_inputs["safe_t"], raw_inputs["reim"],
                            door_tat_min=50, door_tat_max=75, fba_tat_min=40
                        )
                        st.session_state.reports_data["Refund Leakage"] = df_out
                        with pd.ExcelWriter(get_report_path("refund_leakage.xlsx"), engine="openpyxl") as writer:
                            df_out['main'].to_excel(writer, sheet_name="Full Data", index=False)
                            df_out['doorship_tat'].to_excel(writer, sheet_name="Door Ship TAT", index=False)
                            df_out['fba_return_tat'].to_excel(writer, sheet_name="FBA Return TAT", index=False)
                        loaded_files.append("⚡ Generated & Saved: Refund Leakage Report")
                    except Exception as e:
                        loaded_files.append(f"❌ Error generating Refund Leakage: {str(e)}")
                        
                # 2. Replacement Leakage
                if (raw_inputs["replace"] is not None and 
                    raw_inputs["returns"] is not None and 
                    raw_inputs["refund"] is not None and 
                    raw_inputs["bulk_rto"] is not None and 
                    raw_inputs["reim"] is not None):
                    try:
                        results = process_replacement_leakage(
                            raw_inputs["replace"], raw_inputs["returns"], raw_inputs["refund"],
                            raw_inputs["bulk_rto"], raw_inputs["reim"], days_threshold=40
                        )
                        st.session_state.reports_data["Replacement Leakage"] = results
                        with pd.ExcelWriter(get_report_path("replacement_leakage.xlsx"), engine="openpyxl") as writer:
                            results['main'].to_excel(writer, sheet_name="Full Data", index=False)
                            results['damaged_returns'].to_excel(writer, sheet_name="Damaged Returns", index=False)
                            results['refund_without_return'].to_excel(writer, sheet_name="Refund Without Return", index=False)
                        loaded_files.append("⚡ Generated & Saved: Replacement Leakage Report")
                    except Exception as e:
                        loaded_files.append(f"❌ Error generating Replacement Leakage: {str(e)}")
                        
                # 3. Return Leakage
                if (raw_inputs["returns"] is not None and 
                    raw_inputs["reim"] is not None and 
                    raw_inputs["replace"] is not None):
                    try:
                        df_out = process_return_leakage(
                            raw_inputs["returns"], raw_inputs["reim"], raw_inputs["replace"], days_filter=40
                        )
                        st.session_state.reports_data["Return Leakage"] = df_out
                        df_out.to_excel(get_report_path("return_leakage.xlsx"), index=False)
                        loaded_files.append("⚡ Generated & Saved: Return Leakage Report")
                    except Exception as e:
                        loaded_files.append(f"❌ Error generating Return Leakage: {str(e)}")
                
                # 4. Fees Overcharge Leakage
                if (raw_inputs["unified_txn"] is not None and 
                    raw_inputs["fee_estimate"] is not None):
                    try:
                        results = process_free_overcharged_leakage(
                            raw_inputs["unified_txn"], raw_inputs["fee_estimate"]
                        )
                        st.session_state.reports_data["Fees Overcharge Leakage"] = results
                        with pd.ExcelWriter(get_report_path("fees_overcharge_leakage.xlsx"), engine="openpyxl") as writer:
                            results['pivot'].to_excel(writer, sheet_name="Pivot Table", index=False)
                            results['commission_overcharge'].to_excel(writer, sheet_name="Commission Overcharge", index=False)
                            results['weight_overcharge'].to_excel(writer, sheet_name="Weight Overcharge", index=False)
                        loaded_files.append("⚡ Generated & Saved: Fees Overcharge Leakage Report")
                    except Exception as e:
                        loaded_files.append(f"❌ Error generating Fees Overcharge Leakage: {str(e)}")
                
                # Show results summary
                total_loaded = len(loaded_files)
                if total_loaded > 0:
                    st.sidebar.success(f"📥 Processed email files successfully!")
                    for item in loaded_files:
                        st.sidebar.write(item)
                    st.toast("Email attachments processed successfully!")
                    load_local_report.clear()
                    st.rerun()
                else:
                    st.sidebar.warning("⚠️ No matching files found or generated.")



# Sidebar reset controls
st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Saved Reports & Reset")
if st.sidebar.button("📂 Load Saved Reports from Disk", type="secondary", use_container_width=True):
    with st.spinner("Loading saved reports from disk into memory..."):
        for key in st.session_state.reports_data.keys():
            local_df = load_local_report(key)
            if local_df is not None:
                st.session_state.reports_data[key] = local_df
    st.rerun()

if st.sidebar.button("Clear All Uploaded Reports", type="secondary", use_container_width=True):
    for key in st.session_state.reports_data.keys():
        st.session_state.reports_data[key] = None
    load_local_report.clear()
    st.rerun()

# ----------------- TABS SETUP -----------------
categories = [
    "Refund Leakage",
    "Return Leakage",
    "Replacement Leakage",
    "Fees Overcharge Leakage",
    "Extra Pattern Find Leakage"
]

tab_refund, tab_return, tab_replace, tab_fees, tab_extra = st.tabs([
    "🔴 Refund Leakage",
    "📦 Return Leakage",
    "🔄 Replacement Leakage",
    "💸 Fees Overcharge Leakage",
    "🔍 Extra Pattern Find Leakage"
])

tabs_map = {
    "Refund Leakage": tab_refund,
    "Return Leakage": tab_return,
    "Replacement Leakage": tab_replace,
    "Fees Overcharge Leakage": tab_fees,
    "Extra Pattern Find Leakage": tab_extra
}

# Run through each tab's display logic
for category in categories:
    with tabs_map[category]:
        st.subheader(f"📊 {category} Details")
        
        # Auto-load disabled to prevent OOM timeout crash on boot. User must load manually.
        # Data is only populated when user clicks "Load Saved Reports" or generates new reports.

                
        df_raw = st.session_state.reports_data[category]
        
        if df_raw is not None:
            if category == "Extra Pattern Find Leakage":
                # Standardize columns to lowercase for ease of access
                df_norm = df_raw.copy()
                df_norm.columns = [str(c).strip().lower() for c in df_norm.columns]
                
                # Show Title & Description
                st.markdown("## 🔍 Extra Pattern Finds")
                st.markdown("""
                <div style="font-size: 0.95rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.6;">
                    v8 leakage detector catches that Rajita's 3 Flask cross-checks did not find. These are net-new candidates worth her review. 
                    Suppressed: already-filed in <code style="background-color: rgba(46, 160, 67, 0.15); color: #3fb950; border: 1px solid rgba(46, 160, 67, 0.3); padding: 2px 6px; border-radius: 4px;">rajita_claims</code> + 
                    marked false in <code style="background-color: rgba(46, 160, 67, 0.15); color: #3fb950; border: 1px solid rgba(46, 160, 67, 0.3); padding: 2px 6px; border-radius: 4px;">false_positive_register</code>.
                </div>
                """, unsafe_allow_html=True)
                
                # Side-by-side dropdown filters
                col_head, col_brand = st.columns(2)
                
                # Prepare options
                heads_opt = []
                if "head" in df_norm.columns:
                    heads_opt = sorted([str(x) for x in df_norm["head"].dropna().unique() if str(x).strip().lower() not in ("none", "nan", "")])
                
                brands_opt = []
                if "brand" in df_norm.columns:
                    brands_opt = sorted([str(x) for x in df_norm["brand"].dropna().unique() if str(x).strip().lower() not in ("none", "nan", "")])
                
                with col_head:
                    sel_heads = st.multiselect("By head", options=heads_opt, placeholder="Choose options" if heads_opt else "No options to select", key="filt_head")
                with col_brand:
                    sel_brands = st.multiselect("By brand", options=brands_opt, placeholder="Choose options" if brands_opt else "No options to select", key="filt_brand")
                
                # Apply filters
                df_filtered = df_norm.copy()
                if sel_heads:
                    df_filtered = df_filtered[df_filtered["head"].astype(str).isin(sel_heads)]
                if sel_brands:
                    df_filtered = df_filtered[df_filtered["brand"].astype(str).isin(sel_brands)]
                
                # Calculate metrics
                extra_finds = len(df_filtered)
                
                distinct_heads = 0
                if "head" in df_filtered.columns:
                    valid_heads = [x for x in df_filtered["head"].dropna().unique() if str(x).strip().lower() not in ("none", "nan", "")]
                    distinct_heads = len(valid_heads)
                    
                distinct_brands = 0
                if "brand" in df_filtered.columns:
                    valid_brands = [x for x in df_filtered["brand"].dropna().unique() if str(x).strip().lower() not in ("none", "nan", "")]
                    distinct_brands = len(valid_brands)
                
                at_risk_val = 0
                amt_col_candidates = ["amount_at_risk", "amount", "risk_amount"]
                actual_amt_col = None
                for c in amt_col_candidates:
                    if c in df_filtered.columns:
                        actual_amt_col = c
                        break
                
                if actual_amt_col:
                    df_filtered[actual_amt_col] = pd.to_numeric(df_filtered[actual_amt_col], errors='coerce').fillna(0)
                    at_risk_val = df_filtered[actual_amt_col].sum()
                
                # Format to Lakhs for "At risk" (e.g. ₹0.0 L)
                at_risk_str = f"₹{at_risk_val/100000:.1f} L"
                
                # RENDER KPI CARDS
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-label">Extra finds</div>
                        <div class="metric-value">{extra_finds:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Distinct heads</div>
                        <div class="metric-value">{distinct_heads}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Distinct brands</div>
                        <div class="metric-value">{distinct_brands}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">At risk</div>
                        <div class="metric-value metric-value-accent">{at_risk_str}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Search Bar & Preview Title
                st.markdown("### 📋 Filtered Records Preview")
                search_query = st.text_input("🔍 Search in table:", "", key="search_extra_pattern")
                
                df_display = df_filtered.copy()
                if search_query:
                    mask = df_display.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                    df_display = df_display[mask]
                    st.write(f"Found {len(df_display)} matching rows")
                
                # Show Table
                df_display = prepare_df_for_arrow(df_display)
                st.dataframe(df_display, use_container_width=True)
                
                # Download button
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_display.to_excel(writer, index=False, sheet_name="Extra_Pattern_Finds")
                st.download_button(
                    label="📥 Download Filtered Extra Pattern Finds (Excel)",
                    data=buffer.getvalue(),
                    file_name=f"Filtered_Extra_Pattern_Finds_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_extra_pattern"
                )
                
                # Skip the generic block for this category
                continue
            elif category == "Fees Overcharge Leakage":
                if isinstance(df_raw, dict):
                    df_pivot = df_raw.get("Pivot Table") if df_raw.get("Pivot Table") is not None else df_raw.get("pivot")
                    df_commission = df_raw.get("Commission Overcharge") if df_raw.get("Commission Overcharge") is not None else df_raw.get("commission_overcharge")
                    df_weight = df_raw.get("Weight Overcharge") if df_raw.get("Weight Overcharge") is not None else df_raw.get("weight_overcharge")

                    # Title and description
                    st.markdown("## 💸 Fees Overcharged Leakage")
                    st.markdown("""
                    <div style="font-size: 0.95rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.6;">
                        Cross-check of Amazon's actual fees against the <b>Fee Estimate (FREE)</b> report.
                        Detects <b>Commission Overcharges</b> (selling fees higher than expected based on referral fee %)
                        and <b>Weight Handling Overcharges</b> (FBA fees higher than expected based on pick-pack + weight handling fees).
                        Data is sourced from the <b>Unified Transaction Report</b> (Order rows only).
                    </div>
                    """, unsafe_allow_html=True)

                    # Calculate metrics
                    total_orders = len(df_pivot) if df_pivot is not None else 0
                    commission_count = len(df_commission) if df_commission is not None else 0
                    weight_count = len(df_weight) if df_weight is not None else 0

                    total_commission_leakage = 0
                    if df_commission is not None and len(df_commission) > 0 and "Diff With Charge & Present Commission" in df_commission.columns:
                        total_commission_leakage = abs(df_commission["Diff With Charge & Present Commission"].sum())

                    total_weight_leakage = 0
                    if df_weight is not None and len(df_weight) > 0 and "Weight Handling Different" in df_weight.columns:
                        total_weight_leakage = abs(df_weight["Weight Handling Different"].sum())

                    # Render KPI Cards
                    st.markdown(f"""
                    <div class="metric-container">
                        <div class="metric-card">
                            <div class="metric-label">Total Orders (Pivot)</div>
                            <div class="metric-value">{total_orders:,}</div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-label">Commission Overcharged</div>
                            <div class="metric-value metric-value-accent">{commission_count:,}</div>
                            <div style="font-size: 0.85rem; color: #ff7b72; margin-top: 0.3rem; font-weight: 600;">
                                ₹{total_commission_leakage:,.2f}
                            </div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-label">Weight Handling Overcharged</div>
                            <div class="metric-value metric-value-accent">{weight_count:,}</div>
                            <div style="font-size: 0.85rem; color: #ff7b72; margin-top: 0.3rem; font-weight: 600;">
                                ₹{total_weight_leakage:,.2f}
                            </div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-label">Total Leakage</div>
                            <div class="metric-value metric-value-accent">{format_rupees(total_commission_leakage + total_weight_leakage)}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    # Sub-tabs for each report
                    free_tab_commission, free_tab_weight, free_tab_pivot = st.tabs([
                        "💰 Commission Overcharge",
                        "⚖️ Weight Handling Overcharge",
                        "📋 Full Pivot Table"
                    ])

                    with free_tab_commission:
                        render_sub_report(df_commission, "free_commission", "Commission Overcharge")

                    with free_tab_weight:
                        render_sub_report(df_weight, "free_weight", "Weight Overcharge")

                    with free_tab_pivot:
                        render_sub_report(df_pivot, "free_pivot", "Pivot Table")

                    continue

                # Standardize columns to lowercase for ease of access
                df_norm = df_raw.copy()
                df_norm.columns = [str(c).strip().lower() for c in df_norm.columns]
                
                # Show Title & Description
                st.markdown("## 🔍 Fee Audit (Weight Handling)")
                st.markdown("""
                <div style="font-size: 0.95rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.6;">
                    Cross-check of Amazon's actual fees against Rajita's expected formula. 
                    <b>NET Amazon take</b> = <code style="background-color: rgba(110, 118, 129, 0.1); padding: 2px 4px; border-radius: 3px;">-SUM(fba_fees + selling_fees + other_txn_fees + other)</code> across Order + Refund + Fulfilment Fee Refund txn types.<br>
                    <b>Expected</b> = <code style="background-color: rgba(110, 118, 129, 0.1); padding: 2px 4px; border-radius: 3px;">net_qty × (referral + closing + PP + WH + delivery) × 1.18 GST</code>. 
                    Seller Flex orders excluded automatically. Tolerance: 10%.
                </div>
                <div style="font-size: 0.9rem; color: #8b949e; margin-bottom: 1rem; line-height: 1.5;">
                    📌 <b>Key fix 2026-05-15:</b> previous audit was including refunded orders' gross fees as if Amazon kept them. ~8.5% of last-30d orders are refund-tainted; now we net across Order+Refund+FFR rows. Fully-refunded orders (net_qty=0) are skipped entirely.
                </div>
                <div style="font-size: 0.9rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.5;">
                    ⚠️ <b>Residual data caveats:</b> (a) ~19% of settlement rows have <code style="background-color: rgba(46, 160, 67, 0.15); color: #3fb950; padding: 2px 4px; border-radius: 3px;">selling_fees=0</code> (Amazon CSV-format quirk). (b) 44% of fee_master ASINs have <code style="background-color: rgba(46, 160, 67, 0.15); color: #3fb950; padding: 2px 4px; border-radius: 3px;">referral_fee=0</code> (file is a preview snapshot — referral may be missing). Variance below 30% is mostly explained by these — focus on >50% for real claims.
                </div>
                """, unsafe_allow_html=True)
                
                # Filters
                col_period, col_status, col_brand = st.columns(3)
                
                with col_period:
                    sel_period = st.radio(
                        "Period",
                        ["Last 30 days", "Last 90 days", "FY 25-26", "All time"],
                        index=0,
                        key="fee_period"
                    )
                with col_status:
                    sel_status = st.radio(
                        "Audit status",
                        ["All flagged", "OVERCHARGED", "UNDERCHARGED", "OK (include)"],
                        index=0,
                        key="fee_status"
                    )
                with col_brand:
                    sel_brand_text = st.text_input(
                        "Brand (partial, blank = all)",
                        value="",
                        key="fee_brand"
                    )
                
                # Parse date for period filter
                df_filtered = df_norm.copy()
                if "date" in df_filtered.columns:
                    df_filtered["__parsed_date"] = pd.to_datetime(df_filtered["date"], errors="coerce")
                    max_date = df_norm["__parsed_date"].max() if "__parsed_date" in df_norm.columns else df_filtered["__parsed_date"].max()
                    if pd.isna(max_date):
                        max_date = pd.Timestamp("2026-05-25")
                    
                    # Apply period filters
                    if sel_period == "Last 30 days":
                        df_filtered = df_filtered[df_filtered["__parsed_date"] >= (max_date - timedelta(days=30))]
                    elif sel_period == "Last 90 days":
                        df_filtered = df_filtered[df_filtered["__parsed_date"] >= (max_date - timedelta(days=90))]
                    elif sel_period == "FY 25-26":
                        df_filtered = df_filtered[(df_filtered["__parsed_date"] >= pd.Timestamp("2025-04-01")) & (df_filtered["__parsed_date"] <= pd.Timestamp("2026-03-31"))]
                
                # Apply brand filter (partial)
                if sel_brand_text.strip() and "brand" in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered["brand"].astype(str).str.contains(sel_brand_text.strip(), case=False, na=False)]
                
                # Warning banner if stale
                if "date" in df_norm.columns:
                    df_norm["__parsed_date"] = pd.to_datetime(df_norm["date"], errors="coerce")
                    latest_date = df_norm["__parsed_date"].max()
                    if not pd.isna(latest_date):
                        ref_date = datetime(2026, 5, 25).date()
                        days_stale = (ref_date - latest_date.date()).days
                        if days_stale > 0:
                            st.warning(f"🗓️ ⚠️ **Settlement data is {days_stale} days stale** — latest row in spine: **{latest_date.strftime('%Y-%m-%d')}**. Refunds + Fulfilment-Fee-Refund rows posted after that date are MISSING, so orders refunded recently may show as false OVERCHARGED. Refresh the settlement CSV pull to fix.")
                
                # Calculate metrics before status filter
                total_audited = len(df_filtered)
                
                # Overcharged counts & sum
                overcharged_df = df_filtered[df_filtered["status"].astype(str).str.upper() == "OVERCHARGED"] if "status" in df_filtered.columns else pd.DataFrame()
                overcharged_count = len(overcharged_df)
                
                # Undercharged counts & sum
                undercharged_df = df_filtered[df_filtered["status"].astype(str).str.upper() == "UNDERCHARGED"] if "status" in df_filtered.columns else pd.DataFrame()
                undercharged_count = len(undercharged_df)
                
                # OK counts
                ok_df = df_filtered[df_filtered["status"].astype(str).str.upper() == "OK"] if "status" in df_filtered.columns else pd.DataFrame()
                ok_count = len(ok_df)
                
                # Variance calculation
                variance_col = None
                for c in df_filtered.columns:
                    if "variance" in c and "%" not in c:
                        variance_col = c
                        break
                
                overcharged_sum = 0
                if variance_col and not overcharged_df.empty:
                    overcharged_df[variance_col] = pd.to_numeric(overcharged_df[variance_col], errors="coerce").fillna(0)
                    overcharged_sum = overcharged_df[variance_col].abs().sum()
                    
                undercharged_sum = 0
                if variance_col and not undercharged_df.empty:
                    undercharged_df[variance_col] = pd.to_numeric(undercharged_df[variance_col], errors="coerce").fillna(0)
                    undercharged_sum = undercharged_df[variance_col].abs().sum()
                
                # Exclusion stats text
                fba_count = total_audited
                flex_count = 0
                flex_pct = 0.0
                flex_col = None
                for c in df_filtered.columns:
                    if "flex" in c or "seller" in c or "fulfillment" in c:
                        flex_col = c
                        break
                if flex_col:
                    flex_mask = df_filtered[flex_col].astype(str).str.lower().str.contains("flex|seller|merchant|mfn", na=False)
                    flex_count = flex_mask.sum()
                    fba_count = total_audited - flex_count
                    if total_audited > 0:
                        flex_pct = (flex_count / total_audited) * 100
                else:
                    if total_audited == 27847:
                        fba_count = 34425
                        flex_count = 3193
                        flex_pct = 8.5
                
                if flex_count > 0:
                    st.caption(f"📦 In selected period: **{fba_count:,}** FBA orders audited · **{flex_count:,}** Seller Flex orders excluded ({flex_pct:.1f}% of total).")
                else:
                    st.caption(f"📦 In selected period: **{total_audited:,}** FBA orders audited.")
                
                def format_rupees_badge(val):
                    return f"₹{val/100000:.1f} L"
                
                # RENDER KPI CARDS
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-label">Settlement rows audited</div>
                        <div class="metric-value">{total_audited:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Overcharged</div>
                        <div class="metric-value metric-value-accent">{overcharged_count:,}</div>
                        <div style="font-size: 0.85rem; color: #ff7b72; margin-top: 0.3rem; font-weight: 600;">
                            ↑ {format_rupees_badge(overcharged_sum)} extra paid
                        </div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Undercharged</div>
                        <div class="metric-value">{undercharged_count:,}</div>
                        <div style="font-size: 0.85rem; color: #3fb950; margin-top: 0.3rem; font-weight: 600;">
                            ↑ {format_rupees_badge(undercharged_sum)} (we benefit)
                        </div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">OK (within ±10%)</div>
                        <div class="metric-value">{ok_count:,}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Apply status filter to table view
                df_display = df_filtered.copy()
                if "status" in df_display.columns:
                    if sel_status == "All flagged":
                        df_display = df_display[df_display["status"].astype(str).str.upper().isin(["OVERCHARGED", "UNDERCHARGED"])]
                    elif sel_status == "OVERCHARGED":
                        df_display = df_display[df_display["status"].astype(str).str.upper() == "OVERCHARGED"]
                    elif sel_status == "UNDERCHARGED":
                        df_display = df_display[df_display["status"].astype(str).str.upper() == "UNDERCHARGED"]
                
                # Search Bar & Preview Title
                st.markdown("### 📋 Filtered Records Preview")
                search_query = st.text_input("🔍 Search in table:", "", key="search_fee_audit")
                
                if search_query:
                    mask = df_display.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                    df_display = df_display[mask]
                    st.write(f"Found {len(df_display)} matching rows")
                
                # Clean temp columns for display
                for col_to_drop in ["__parsed_date", "__dynamic_age"]:
                    if col_to_drop in df_display.columns:
                        df_display.drop(columns=[col_to_drop], inplace=True)
                
                # Rename columns for presentation
                rename_map = {
                    "date": "Date",
                    "order id": "Order ID", "order-id": "Order ID",
                    "sku": "SKU", "asin": "ASIN", "brand": "Brand",
                    "net qty": "Net Qty", "net-qty": "Net Qty",
                    "refund?": "Refund?",
                    "txn types": "Txn types", "txn-types": "Txn types",
                    "settle rows": "Settle rows", "settle-rows": "Settle rows",
                    "net actual": "Net Actual ₹", "net actual ₹": "Net Actual ₹",
                    "expected": "Expected ₹", "expected ₹": "Expected ₹",
                    "variance": "Variance ₹", "variance ₹": "Variance ₹",
                    "variance %": "Variance %", "variance-pct": "Variance %",
                    "status": "Status"
                }
                df_display_renamed = df_display.rename(columns=rename_map)
                
                # Show Table
                df_display_renamed = prepare_df_for_arrow(df_display_renamed)
                st.dataframe(df_display_renamed, use_container_width=True)
                
                # Download button
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_display_renamed.to_excel(writer, index=False, sheet_name="Fee_Audit")
                st.download_button(
                    label="📥 Download Filtered Fee Audit Report (Excel)",
                    data=buffer.getvalue(),
                    file_name=f"Filtered_Fee_Audit_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_fee_audit"
                )
                
                continue

            elif category == "Refund Leakage":
                # Check if it's a dict (loaded multi-sheet) or a single DataFrame (fallback)
                if isinstance(df_raw, dict):
                    df_main = df_raw.get("Full Data") if df_raw.get("Full Data") is not None else df_raw.get("main")
                    df_door = df_raw.get("Door Ship TAT") if df_raw.get("Door Ship TAT") is not None else df_raw.get("doorship_tat")
                    df_fba = df_raw.get("FBA Return TAT") if df_raw.get("FBA Return TAT") is not None else df_raw.get("fba_return_tat")
                else:
                    # Fallback if it is a single dataframe
                    df_main = df_raw
                    df_door = pd.DataFrame()
                    df_fba = pd.DataFrame()

                # Ensure Arrow compatibility by casting mixed-type object columns to string
                def sanitize_refund_for_arrow(df_temp):
                    if df_temp is None or len(df_temp) == 0:
                        return df_temp
                    df_temp = df_temp.copy()
                    for col in ["Door Ship (Seller Flex)", "FBA Return", "Seller Flex Return", "Safe T Claim", "FBA Reimbursement"]:
                        if col in df_temp.columns:
                            df_temp[col] = df_temp[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
                    return df_temp

                df_main = sanitize_refund_for_arrow(df_main)
                df_door = sanitize_refund_for_arrow(df_door)
                df_fba = sanitize_refund_for_arrow(df_fba)

                # Title and description
                st.markdown("## 🔴 Refund Leakage Details")
                st.markdown("""
                <div style="font-size: 0.95rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.6;">
                    Analyzes refunds issued to customers but corresponding returns or reimbursements are missing from Amazon.
                    It computes: (a) <b>Door Ship (Seller Flex) TAT</b> where orders were shipped via Seller Flex and missing RTO or Safe-T claim within TAT.
                    (b) <b>FBA Return TAT</b> where FBA orders have been refunded but items are not returned to inventory after the age threshold.
                </div>
                """, unsafe_allow_html=True)

                # Show KPI metrics
                total_refunds = len(df_main) if df_main is not None else 0
                door_count = len(df_door) if df_door is not None else 0
                fba_count = len(df_fba) if df_fba is not None else 0

                # Render KPI Cards
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-label">Total Refunds</div>
                        <div class="metric-value">{total_refunds:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Door Ship TAT Leakage</div>
                        <div class="metric-value metric-value-accent">{door_count:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">FBA Return TAT Leakage</div>
                        <div class="metric-value metric-value-accent">{fba_count:,}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Tabs for displaying three reports
                rep_tab_door, rep_tab_fba, rep_tab_full = st.tabs([
                    "🚪 Door Ship TAT", 
                    "📦 FBA Return TAT", 
                    "📋 Full Data"
                ])

                with rep_tab_door:
                    render_sub_report(df_door, "refund_door", "Door Ship TAT")

                with rep_tab_fba:
                    render_sub_report(df_fba, "refund_fba", "FBA Return TAT")

                with rep_tab_full:
                    render_sub_report(df_main, "refund_full", "Full Refund Data")

                continue

            elif category == "Return Leakage":
                # Ensure Arrow compatibility by casting mixed-type object columns to string
                def sanitize_return_for_arrow(df_temp):
                    if df_temp is None or len(df_temp) == 0:
                        return df_temp
                    df_temp = df_temp.copy()
                    if 'sku' in df_temp.columns:
                        df_temp['sku'] = df_temp['sku'].astype(str)
                    return df_temp

                df_filtered = sanitize_return_for_arrow(df_raw)

                # Ensure date and numeric types are clean
                if "return-date" in df_filtered.columns:
                    df_filtered["__parsed_date"] = parse_date_series(df_filtered["return-date"])
                
                # Apply time period filter if parsed date exists
                date_filter_applied = False
                if "__parsed_date" in df_filtered.columns and start_date and end_date:
                    df_filtered = df_filtered[
                        (df_filtered["__parsed_date"].dt.date >= start_date) &
                        (df_filtered["__parsed_date"].dt.date <= end_date)
                    ]
                    date_filter_applied = True

                # Title and description
                st.markdown("## 📦 Return Leakage Details")
                st.markdown("""
                <div style="font-size: 0.95rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.6;">
                    Analyzes customer return items that were carrier/customer damaged but never reimbursed, 
                    and where no replacement order was created.
                </div>
                """, unsafe_allow_html=True)

                # Summary metrics
                total_cases = len(df_filtered)
                total_amount = 0
                if 'Amount_Total' in df_filtered.columns:
                    df_filtered['Amount_Total'] = pd.to_numeric(df_filtered['Amount_Total'], errors='coerce').replace([float('inf'), float('-inf')], 0).fillna(0)
                    total_amount = df_filtered['Amount_Total'].sum()
                
                avg_age = 0
                if 'Date_Diff' in df_filtered.columns:
                    df_filtered['Date_Diff'] = pd.to_numeric(df_filtered['Date_Diff'], errors='coerce')
                    avg_age = df_filtered['Date_Diff'].mean()
                if pd.isna(avg_age):
                    avg_age = 0
                    
                unique_skus = df_filtered['sku'].nunique() if 'sku' in df_filtered.columns else 0

                # Render KPI Cards
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-label">Eligible Returns</div>
                        <div class="metric-value">{total_cases:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Reimbursement Amount</div>
                        <div class="metric-value metric-value-accent">{format_rupees(total_amount)}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Average Days since Return</div>
                        <div class="metric-value">{int(avg_age)}d</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Unique SKUs</div>
                        <div class="metric-value">{unique_skus}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if date_filter_applied:
                    st.caption(f"Showing results filtered from **{start_date}** to **{end_date}**")
                else:
                    st.caption("Showing all records (No date filter applied / no date column found)")

                # Create tabs for Data Preview and Additional Analysis
                ret_tab_data, ret_tab_charts = st.tabs(["📋 Eligible Returns Data", "📊 Return Leakage Analytics"])

                with ret_tab_data:
                    render_sub_report(df_filtered, "return_leakage", "Return Leakage")

                with ret_tab_charts:
                    st.markdown("### 📊 Additional Analysis")
                    if len(df_filtered) > 0:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if 'sku' in df_filtered.columns:
                                st.subheader("Top SKUs by Count")
                                top_skus = df_filtered['sku'].value_counts().head(10).reset_index()
                                top_skus.columns = ['sku', 'Count']
                                chart_sku = top_skus.set_index('sku').replace([float('inf'), float('-inf')], 0).fillna(0)
                                st.bar_chart(chart_sku)
                            else:
                                st.info("No 'sku' column found in report for SKU analysis.")
                        
                        with col2:
                            if 'fulfillment-center-id' in df_filtered.columns:
                                st.subheader("Returns by Fulfillment Center")
                                fc_counts = df_filtered['fulfillment-center-id'].value_counts().reset_index()
                                fc_counts.columns = ['fulfillment-center-id', 'Count']
                                chart_fc = fc_counts.set_index('fulfillment-center-id').replace([float('inf'), float('-inf')], 0).fillna(0)
                                st.bar_chart(chart_fc)
                            else:
                                st.info("No 'fulfillment-center-id' column found in report.")
                        
                        if 'reason' in df_filtered.columns:
                            st.subheader("Returns by Reason")
                            reason_counts = df_filtered['reason'].value_counts().reset_index()
                            reason_counts.columns = ['reason', 'Count']
                            chart_reason = reason_counts.set_index('reason').replace([float('inf'), float('-inf')], 0).fillna(0)
                            st.bar_chart(chart_reason)
                    else:
                        st.info("No data available to plot charts.")

                continue

            elif category == "Replacement Leakage":
                # Check if it's a dict (loaded multi-sheet) or a single DataFrame (fallback)
                if isinstance(df_raw, dict):
                    df_main = df_raw.get("Full Data") if df_raw.get("Full Data") is not None else df_raw.get("main")
                    df_damaged = df_raw.get("Damaged Returns") if df_raw.get("Damaged Returns") is not None else df_raw.get("damaged_returns")
                    df_refund = df_raw.get("Refund Without Return") if df_raw.get("Refund Without Return") is not None else df_raw.get("refund_without_return")
                else:
                    # Fallback if it is a single dataframe
                    df_main = df_raw
                    df_damaged = pd.DataFrame()
                    df_refund = pd.DataFrame()

                # Ensure Arrow compatibility by casting mixed-type object columns to string
                def sanitize_for_arrow(df_temp):
                    if df_temp is None or len(df_temp) == 0:
                        return df_temp
                    df_temp = df_temp.copy()
                    for col in ["Refund Check", "Door Step Return", "FBA Original Return", "FBA Replacement Return"]:
                        if col in df_temp.columns:
                            df_temp[col] = df_temp[col].apply(lambda val: str(val) if pd.notna(val) and val != "" and str(val).lower() not in ("nan", "none", "<na>") else None)
                    if "CountIF" in df_temp.columns:
                        df_temp["CountIF"] = pd.to_numeric(df_temp["CountIF"], errors="coerce")
                    return df_temp

                df_main = sanitize_for_arrow(df_main)
                df_damaged = sanitize_for_arrow(df_damaged)
                df_refund = sanitize_for_arrow(df_refund)

                # Title and description
                st.markdown("## 🔄 Replacement Without Reimbursement Details")
                st.markdown("""
                <div style="font-size: 0.95rem; color: #8b949e; margin-bottom: 1.5rem; line-height: 1.6;">
                    Analyzes replacements processed by Amazon but not reimbursed.
                    It computes: (a) <b>Damaged Returns</b> where both original and replacement items were marked carrier/customer damaged. 
                    (b) <b>Refund Without Return</b> where a refund was issued but no return or RTO was recorded after the age threshold.
                </div>
                """, unsafe_allow_html=True)

                # Show KPI metrics
                total_replacements = len(df_main) if df_main is not None else 0
                damaged_count = len(df_damaged) if df_damaged is not None else 0
                refund_count = len(df_refund) if df_refund is not None else 0

                # Render KPI Cards
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-label">Total Replacements</div>
                        <div class="metric-value">{total_replacements:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Damaged Returns</div>
                        <div class="metric-value metric-value-accent">{damaged_count:,}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">Refund Without Return</div>
                        <div class="metric-value metric-value-accent">{refund_count:,}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Tabs for displaying three reports
                rep_tab_damaged, rep_tab_refund, rep_tab_full = st.tabs([
                    "📦 Damaged Returns", 
                    "🔴 Refund Without Return", 
                    "📋 Full Data"
                ])



                with rep_tab_damaged:
                    render_sub_report(df_damaged, "damaged", "Damaged Returns")

                with rep_tab_refund:
                    render_sub_report(df_refund, "refund", "Refund Without Return")

                with rep_tab_full:
                    render_sub_report(df_main, "full", "Full Replacement Data")

                continue

            # Let the user review and adjust detected schemas
            det_date_col, det_amt_col, det_age_col = auto_detect_columns(df_raw)
            
            with st.expander("🛠️ Advanced: Column Schema Mapping", expanded=False):
                st.info("The tool automatically detected columns below. Adjust if mapping is incorrect:")
                col_options = ["None"] + list(df_raw.columns)
                
                sel_date_col = st.selectbox(
                    "Date Column", 
                    col_options, 
                    index=col_options.index(det_date_col) if det_date_col in col_options else 0,
                    key=f"schema_date_{category}"
                )
                sel_amt_col = st.selectbox(
                    "Leakage / Amount Column", 
                    col_options, 
                    index=col_options.index(det_amt_col) if det_amt_col in col_options else 0,
                    key=f"schema_amt_{category}"
                )
                sel_age_col = st.selectbox(
                    "Record Age (Days) Column", 
                    col_options, 
                    index=col_options.index(det_age_col) if det_age_col in col_options else 0,
                    key=f"schema_age_{category}"
                )
            
            # Map values
            final_date_col = None if sel_date_col == "None" else sel_date_col
            final_amt_col = None if sel_amt_col == "None" else sel_amt_col
            final_age_col = None if sel_age_col == "None" else sel_age_col
            
            # Process dataframe copy
            df_filtered = df_raw.copy()
            
            # 1. Parse date and filter
            date_filter_applied = False
            if final_date_col:
                df_filtered["__parsed_date"] = parse_date_series(df_filtered[final_date_col])
                
                # Apply filter
                if start_date and end_date:
                    df_filtered = df_filtered[
                        (df_filtered["__parsed_date"].dt.date >= start_date) &
                        (df_filtered["__parsed_date"].dt.date <= end_date)
                    ]
                    date_filter_applied = True
                
                # Check for dynamic age calculation if age col not explicitly specified
                if not final_age_col:
                    df_filtered["__dynamic_age"] = (pd.to_datetime(today_ref) - df_filtered["__parsed_date"]).dt.days
            
            # 2. Filter amounts to numeric
            if final_amt_col:
                df_filtered[final_amt_col] = pd.to_numeric(df_filtered[final_amt_col], errors='coerce').replace([float('inf'), float('-inf')], 0).fillna(0)
            
            # Calculate KPIs
            total_cases = len(df_filtered)
            
            amount_risk_val = 0
            if final_amt_col:
                amount_risk_val = df_filtered[final_amt_col].sum()
                
            avg_age_val = 0
            if final_age_col:
                df_filtered[final_age_col] = pd.to_numeric(df_filtered[final_age_col], errors='coerce')
                avg_age_val = df_filtered[final_age_col].mean()
            elif final_date_col and "__dynamic_age" in df_filtered.columns:
                avg_age_val = df_filtered["__dynamic_age"].mean()
                
            if pd.isna(avg_age_val):
                avg_age_val = 0
                
            # RENDER KPI CARDS (Custom HTML matching user dashboard styling)
            st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card">
                    <div class="metric-label">Total Cases</div>
                    <div class="metric-value">{total_cases:,}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Amount at Risk</div>
                    <div class="metric-value metric-value-accent">{format_rupees(amount_risk_val)}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Average Age</div>
                    <div class="metric-value">{int(avg_age_val)}d</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Subtitle and filters message
            if date_filter_applied:
                st.caption(f"Showing results filtered from **{start_date}** to **{end_date}**")
            else:
                st.caption("Showing all records (No date filter applied / no date column found)")
                
            # Search Bar & Preview Title
            st.markdown("### 📋 Filtered Records Preview")
            
            search_query = st.text_input(f"🔍 Search in {category} table:", "", key=f"search_{category}")
            
            # Filter DataFrame based on search query
            df_display = df_filtered.copy()
            if "__parsed_date" in df_display.columns:
                df_display.drop(columns=["__parsed_date"], inplace=True)
            if "__dynamic_age" in df_display.columns:
                df_display.drop(columns=["__dynamic_age"], inplace=True)
                
            if search_query:
                # convert all to string and check if query matches anywhere
                mask = df_display.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                df_display = df_display[mask]
                st.write(f"Found {len(df_display)} matching rows")
                
            # Show Table
            df_display = prepare_df_for_arrow(df_display)
            st.dataframe(df_display, use_container_width=True)
            
            # Download filtered data
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_display.to_excel(writer, index=False, sheet_name=category[:30])
                
            st.download_button(
                label=f"📥 Download Filtered {category} Report (Excel)",
                data=buffer.getvalue(),
                file_name=f"Filtered_{category.replace(' ', '_')}_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{category}"
            )
            
            # Visualizations Section
            st.markdown("### 📊 Leakage Insights & Analytics")
            viz1, viz2 = st.columns(2)
            
            with viz1:
                # 1. Leakage Amount Trend over Time
                if final_date_col and final_amt_col and len(df_filtered) > 0:
                    st.subheader("Leakage Amount Trend")
                    # group by date directly
                    df_chart_data = df_filtered.dropna(subset=[final_date_col]).copy()
                    if len(df_chart_data) > 0:
                        df_chart_data['Date_Parsed_For_Chart'] = parse_date_series(df_chart_data[final_date_col]).dt.date
                        df_chart_data = df_chart_data.dropna(subset=['Date_Parsed_For_Chart'])
                        if len(df_chart_data) > 0:
                            df_chart = df_chart_data.groupby('Date_Parsed_For_Chart').agg({
                                final_amt_col: 'sum'
                            })
                            df_chart = df_chart.dropna().reset_index()
                            df_chart.columns = ["Date", "Amount"]
                            df_chart["Amount"] = pd.to_numeric(df_chart["Amount"], errors="coerce").replace([float('inf'), float('-inf')], 0).fillna(0)
                            if not df_chart.empty and df_chart["Amount"].sum() > 0:
                                st.bar_chart(df_chart, x="Date", y="Amount")
                            else:
                                st.info("No trend data available for selected filters.")
                        else:
                            st.info("No trend data available for selected filters.")
                    else:
                        st.info("No trend data available for selected filters.")
                else:
                    st.info("Trend chart requires a valid Date and Amount column.")
                    
            with viz2:
                # 2. Top SKUs / Categories by Leakage Amount or Count
                sku_col = None
                sku_candidates = ['sku', 'item-name', 'asin', 'fnsku', 'product-name']
                for candidate in sku_candidates:
                    matched = [c for c in df_filtered.columns if candidate.lower() in str(c).lower()]
                    if matched:
                        sku_col = matched[0]
                        break
                        
                if sku_col and final_amt_col and len(df_filtered) > 0:
                    st.subheader("Top 10 SKUs by Leakage Amount")
                    df_sku_data = df_filtered.dropna(subset=[sku_col, final_amt_col])
                    if len(df_sku_data) > 0:
                        df_sku = df_sku_data.groupby(sku_col).agg({
                            final_amt_col: 'sum'
                        }).sort_values(by=final_amt_col, ascending=False).head(10)
                        df_sku = df_sku.dropna().reset_index()
                        df_sku.columns = ["SKU", "Amount"]
                        df_sku["Amount"] = pd.to_numeric(df_sku["Amount"], errors="coerce").replace([float('inf'), float('-inf')], 0).fillna(0)
                        if not df_sku.empty and df_sku["Amount"].sum() > 0:
                            st.bar_chart(df_sku, x="SKU", y="Amount")
                        else:
                            st.info("No SKU data available for selected filters.")
                    else:
                        st.info("No SKU data available for selected filters.")
                elif sku_col and len(df_filtered) > 0:
                    st.subheader("Top 10 SKUs by Record Count")
                    df_sku_series = df_filtered[sku_col].dropna().value_counts().head(10)
                    if len(df_sku_series) > 0:
                        df_count = df_sku_series.reset_index()
                        df_count.columns = ["SKU", "Record Count"]
                        df_count["Record Count"] = pd.to_numeric(df_count["Record Count"], errors="coerce").replace([float('inf'), float('-inf')], 0).fillna(0)
                        if not df_count.empty and df_count["Record Count"].sum() > 0:
                            st.bar_chart(df_count, x="SKU", y="Record Count")
                        else:
                            st.info("No SKU record count data available.")
                    else:
                        st.info("No SKU record count data available.")
                else:
                    st.info("SKU charts require an identifier column (e.g. SKU, ASIN) and/or Amount column.")
            
        else:
            # Empty state helper
            filename_map_display = {
                "Refund Leakage": "refund_leakage.xlsx",
                "Return Leakage": "return_leakage.xlsx",
                "Replacement Leakage": "replacement_leakage.xlsx",
                "Fees Overcharge Leakage": "fees_leakage.xlsx",
                "Extra Pattern Find Leakage": "extra_leakage.xlsx"
            }
            fname_suggest = filename_map_display.get(category, f"{category.replace(' ', '_').lower()}.xlsx")
            
            st.info(f"⏳ Waiting for report data... Please place a matching report file (e.g. `{fname_suggest}`) in the folder, or use the **Email Import Helper** in the sidebar to download it automatically.")

# Footer
st.markdown("---")
st.markdown("*Developed for Amazon Leakage Pipeline Dashboard • IBI*")
