import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import os

# --- 設定網頁標題 ---
st.set_page_config(page_title="IPQC CPK & Yield 自動生成器", layout="wide")
st.title("📊 IPQC CPK & Yield 報表生成器")
st.markdown("自動掃描 Excel，並依照指定順序輸出 19 個站點的良率與 CPK 報表。")

# --- 1. 定義 19 個站點的標準順序 ---
TARGET_ORDER = [
    "MLA assy installation",
    "Mirror attachment",
    "Barrel attachment",
    "Condenser lens attach",
    "LED Module  attachment",
    "ILLU Module cover attachment",
    "Relay lens attachment",
    "LED FLEX GRAPHITE-1",
    "reflector attach",
    "singlet attach",
    "HWP Mylar attach",
    "PBS attachment",
    "Doublet attachment",
    "Top cover installation",
    "PANEL PRECISION AA（LAA）",
    "POST DAA INSPECTION",
    "PANEL FLEX ASSY",
    "LCOS GRAPHITE ATTACH",
    "DE OQC"
]

# --- 2. 輔助函式：名稱正規化 ---
def normalize_name(name):
    """移除空格、括號、特殊符號並轉小寫，用於模糊比對"""
    return name.lower().replace(" ", "").replace("　", "").replace("(", "").replace(")", "").replace("（", "").replace("）", "").replace("-", "").replace("_", "")

# 建立對照表
TARGET_MAP = {}
for name in TARGET_ORDER:
    key = normalize_name(name)
    TARGET_MAP[key] = name

# --- 3. 核心計算函式 (CPK) ---
def calculate_cpk_value(data, usl, lsl):
    try:
        clean_data = pd.to_numeric(data, errors='coerce').dropna()
        if len(clean_data) < 2: return np.nan
        
        mean = np.mean(clean_data)
        std = np.std(clean_data, ddof=1)
        if std == 0: return np.nan
        
        cpu = np.nan
        cpl = np.nan
        has_usl = False
        has_lsl = False
        
        if not pd.isna(usl):
            cpu = (usl - mean) / (3 * std)
            has_usl = True
        if not pd.isna(lsl):
            cpl = (mean - lsl) / (3 * std)
            has_lsl = True
            
        if has_usl and has_lsl: return min(cpu, cpl)
        elif has_usl: return cpu
        elif has_lsl: return cpl
        else: return np.nan
    except:
        return np.nan

# --- 4. 尋找 Header 列索引 ---
def find_header_row(df, keywords):
    for i in range(min(60, len(df))):
        row_str = " ".join(df.iloc[i].astype(str).fillna("").str.lower())
        for kw in keywords:
            if kw in row_str:
                return i
    return -1

# --- 5. 處理單一 Sheet (Yield) ---
def process_yield(station_display_name, df):
    best_col = -1
    max_count = 0
    cols_to_scan = min(df.shape[1], 30)
    
    for c in range(cols_to_scan):
        col_data = df.iloc[:, c].astype(str).str.upper()
        ok_count = (col_data == "OK").sum()
        ng_count = (col_data == "NG").sum()
        total = ok_count + ng_count
        
        if total > max_count:
            max_count = total
            best_col = c
            
    if best_col != -1 and max_count > 0:
        col_data = df.iloc[:, best_col].astype(str).str.upper()
        ok_qty = (col_data == "OK").sum()
        ng_qty = (col_data == "NG").sum()
        total_qty = ok_qty + ng_qty
        yield_rate = ok_qty / total_qty if total_qty > 0 else 0
        
        return {
            "Station": station_display_name,
            "Total Qty": total_qty,
            "OK Qty": ok_qty,
            "NG Qty": ng_qty,
            "Yield": yield_rate
        }
    return None

# --- 6. 處理單一 Sheet (CPK) ---
def process_cpk(station_display_name, df):
    # 1. 定位標題列
    dim_row_idx = find_header_row(df, ["dim. no", "dim no", "dim.no"])
    usl_row_idx = find_header_row(df, ["usl"])
    lsl_row_idx = find_header_row(df, ["lsl"])
    
    # 嘗試定位 Config 相關標題列 (若 Dim No 同列則無需額外定位)
    # 這裡假設 Config 可能在 Dim No 同一列，或者前 10 列的 metadata 區域
    config_col_idx = -1
    
    if dim_row_idx == -1: return []

    # 2. 解析欄位 (Dim No) 和尋找 Config 欄位
    headers = df.iloc[dim_row_idx].astype(str).fillna("").tolist()
    dim_cols = {}
    
    # 關鍵字黑名單
    ignore_list = ["date", "time", "no.", "remark", "judge", "note", "supplier", "station", "model", "lot", "cavity", "nan", "", "config", "configuration", "type"]
    
    for idx, name in enumerate(headers):
        clean_name = name.strip()
        lower_name = clean_name.lower()
        
        # 偵測 Config 欄位 (如果表頭有 'config', 'model', 'type' 等字眼)
        if config_col_idx == -1 and any(k in lower_name for k in ["config", "model", "type", "description"]):
             config_col_idx = idx
        
        if len(clean_name) > 1 and lower_name not in ignore_list:
            dim_cols[idx] = clean_name

    # 3. 取得規格限 (USL/LSL)
    usls = {}
    lsls = {}
    
    if usl_row_idx != -1:
        row_vals = df.iloc[usl_row_idx].tolist()
        for idx, val in enumerate(row_vals):
            try: usls[idx] = float(val)
            except: pass
            
    if lsl_row_idx != -1:
        row_vals = df.iloc[lsl_row_idx].tolist()
        for idx, val in enumerate(row_vals):
            try: lsls[idx] = float(val)
            except: pass

    # 4. 提取數據並計算
    results = []
    start_row = max(dim_row_idx, usl_row_idx, lsl_row_idx) + 1
    
    # 提取需要的資料區塊
    data_block = df.iloc[start_row:].copy()
    
    # 尋找日期欄位 (假設在前 15 欄內)
    date_col_idx = -1
    for c in range(min(15, data_block.shape[1])):
        sample = data_block.iloc[:, c].astype(str)
        if sample.str.contains(r'202\d-\d{2}-\d{2}', regex=True).any():
            date_col_idx = c
            break
            
    if date_col_idx != -1:
        # 統一日期格式
        data_block['Date_Clean'] = data_block.iloc[:, date_col_idx].astype(str).str.extract(r'(202\d-\d{2}-\d{2})')[0]
        data_block = data_block.dropna(subset=['Date_Clean'])
        
        # 處理 Config 值
        # 如果有找到 Config 欄位，就取值；否則設為空字串或預設值
        if config_col_idx != -1:
            data_block['Config_Val'] = data_block.iloc[:, config_col_idx].astype(str).replace('nan', '')
        else:
            data_block['Config_Val'] = "" # 預設為空，若需要預設值可改這裡，如 "Default"

        grouped = data_block.groupby(['Date_Clean', 'Config_Val'])
        
        for (date, config_val), group in grouped:
            for col_idx, dim_name in dim_cols.items():
                vals = group.iloc[:, col_idx]
                
                u = usls.get(col_idx, np.nan)
                l = lsls.get(col_idx, np.nan)
                
                cpk = calculate_cpk_value(vals, u, l)
                
                clean_vals = pd.to_numeric(vals, errors='coerce').dropna()
                sample_size = len(clean_vals)
                
                if sample_size > 0:
                    results.append({
                        "Station": station_display_name,
                        "Dim No": dim_name,
                        "config": config_val,  # 新增 config 欄位
                        "Date": date,
                        "Sample Size": sample_size,
                        "USL": u if not pd.isna(u) else "",
                        "LSL": l if not pd.isna(l) else "",
                        "CPK": round(cpk, 3) if not pd.isna(cpk) else ""
                    })
                    
    return results

# --- 主程式介面 ---

uploaded_file = st.file_uploader("📂 請上傳 Excel 檔案 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.info("正在讀取並分析所有分頁，請稍候...")
    
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheet_names = xls.sheet_names
        
        yield_list = []
        cpk_list = []
        
        progress_bar = st.progress(0)
        
        for i, sheet_name in enumerate(all_sheet_names):
            norm_sheet = normalize_name(sheet_name)
            
            display_name = None
            for key, val in TARGET_MAP.items():
                if key in norm_sheet or norm_sheet in key:
                    display_name = val
                    break
            
            # 特殊修正
            if "postdaa" in norm_sheet: display_name = "POST DAA INSPECTION"
            if "ledmoduleattachment" in norm_sheet: display_name = "LED Module  attachment"
            
            if not display_name or any(x in norm_sheet for x in ["summary", "slice", "template", "inline", "history"]):
                continue

            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            y_res = process_yield(display_name, df)
            if y_res: yield_list.append(y_res)
            
            c_res = process_cpk(display_name, df)
            if c_res: cpk_list.extend(c_res)
            
            progress_bar.progress((i + 1) / len(all_sheet_names))

        # --- 整理與排序 (Yield) ---
        df_yield = pd.DataFrame(yield_list)
        if not df_yield.empty:
            df_yield['Station'] = pd.Categorical(df_yield['Station'], categories=TARGET_ORDER, ordered=True)
            df_yield = df_yield.sort_values('Station')
            df_yield["Yield"] = df_yield["Yield"].apply(lambda x: f"{x*100:.2f}%")

        # --- 整理與排序 (CPK) ---
        df_cpk = pd.DataFrame(cpk_list)
        if not df_cpk.empty:
            df_cpk['Station'] = pd.Categorical(df_cpk['Station'], categories=TARGET_ORDER, ordered=True)
            
            # 指定欄位順序：加入 config
            cols = ["Station", "Dim No", "config", "Date", "Sample Size", "USL", "LSL", "CPK"]
            df_cpk = df_cpk[cols]
            
            df_cpk = df_cpk.sort_values(by=['Station', 'Dim No', 'Date'])

        st.success("✅ 計算完成！")
        
        tab1, tab2 = st.tabs(["良率總表 (Yield)", "CPK 詳細報表 (含 Config)"])
        
        with tab1:
            st.dataframe(df_yield, use_container_width=True)
            
        with tab2:
            st.dataframe(df_cpk, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            if not df_yield.empty:
                df_yield.to_excel(writer, sheet_name='Yield Summary', index=False)
            if not df_cpk.empty:
                df_cpk.to_excel(writer, sheet_name='CPK Detail', index=False)
                
        output.seek(0)
        
        st.download_button(
            label="📥 下載完整 Excel 報表",
            data=output,
            file_name="IPQC_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"發生錯誤：{e}")