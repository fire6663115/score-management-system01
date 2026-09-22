import os
import sys
import subprocess
import re
import json
from io import BytesIO

# Auto-install dependencies
def install_dependencies():
    packages = []
    try:
        import streamlit
    except ImportError:
        packages.append("streamlit")
    try:
        import openpyxl
    except ImportError:
        packages.append("openpyxl")
    try:
        import pandas
    except ImportError:
        packages.append("pandas")
    
    if packages:
        print(f"Installing missing dependencies: {', '.join(packages)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + packages)

install_dependencies()

import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment

st.set_page_config(page_title="比賽成績整合系統", layout="wide", page_icon="🏆")

# --- Session State Initialization ---
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "mode_b_mapping" not in st.session_state:
    st.session_state.mode_b_mapping = {}
if "judge_header_row" not in st.session_state:
    st.session_state.judge_header_row = None
if "sum_header_row" not in st.session_state:
    st.session_state.sum_header_row = None
if "uploaded_summary" not in st.session_state:
    st.session_state.uploaded_summary = None
if "uploaded_judges" not in st.session_state:
    st.session_state.uploaded_judges = []
if "config_data" not in st.session_state:
    st.session_state.config_data = {}

def clear_cache():
    st.session_state.uploader_key += 1
    st.session_state.mode_b_mapping = {}
    st.session_state.judge_header_row = None
    st.session_state.sum_header_row = None
    st.session_state.uploaded_summary = None
    st.session_state.uploaded_judges = []
    st.session_state.config_data = {}
    st.rerun()

# --- Sidebar Navigation & i18n ---
lang = st.sidebar.radio("Language / 語言", ["English", "中文"])

def t(en_text, zh_text):
    return en_text if lang == "English" else zh_text

st.sidebar.title(t("🛠️ Navigation", "🛠️ 系統導航"))
app_mode = st.sidebar.radio(t("Select Mode", "選擇操作模式"), [
    t("🚀 Mode A: Standard Template", "🚀 模式 A：標準範本一鍵模式"), 
    t("⚙️ Mode B: Custom Data Mapping", "⚙️ 模式 B：自訂表格萬用模式")
])

st.sidebar.divider()
st.sidebar.markdown(t("### Tools", "### 系統工具"))
if st.sidebar.button(t("🗑️ Clear Cache & Reset", "🗑️ 清空重置所有設定與檔案")):
    clear_cache()

st.sidebar.divider()
st.sidebar.caption(t("🔒 Privacy Promise: This system processes your data purely in local memory. All files are immediately discarded after processing, and no sensitive event data is ever stored or logged.", "🔒 隱私承諾：本系統僅於本地記憶體進行暫存運算，處理完畢檔案即刻銷毀，絕不留存任何賽事敏感數據。"))

# --- Helpers ---
def is_participant_code(text):
    if not isinstance(text, str): return False
    return bool(re.match(r"^[A-Z0-9]+-\d{4}-\d+", text.strip()))

def safe_float(val):
    if val is None:
        return None
    try:
        s = str(val).strip()
        if s in ["", "缺考", "-", "Null", "None", "NaN"]:
            return 0.0
        return round(float(s), 1)
    except (ValueError, TypeError):
        return 0.0

def extract_judge_feature(text):
    if not isinstance(text, str): return ""
    match = re.search(r'Judge\s*\d+', text, re.IGNORECASE)
    return match.group(0).lower().replace(" ", "") if match else text.strip().lower()

def style_header(ws, row, cols, bg_color="4F81BD", font_color="FFFFFF"):
    fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    font = Font(color=font_color, bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for col in cols:
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

def generate_judge_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Score Sheet"
    headers = ["No.", "Entry Code", "Artwork Link", "Criteria 1", "Criteria 2", "Score"]
    ws.append(headers)
    style_header(ws, 1, range(1, 7))
    ws.append([1, "USER-0000-00000000-001", "http://example.com", 20, 30, 50])
    for col in ["A", "B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 20
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def generate_summary_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    headers = ["Entry Code"] + [f"Judge {i:02d}" for i in range(1, 13)] + ["TOTAL"]
    ws.append(headers)
    style_header(ws, 1, range(1, 15), bg_color="C0504D")
    ws.append(["USER-0000-00000000-001"] + [50]*12 + [600])
    ws.column_dimensions["A"].width = 25
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def clean_columns(cols):
    res = []
    for i, c in enumerate(cols):
        cs = str(c)
        if pd.isna(c) or "Unnamed" in cs:
            res.append(t(f"Empty Column (Col {i+1})", f"空白欄位 (Column {i+1})"))
        else:
            res.append(cs)
    return res

def interactive_header_selector(df_preview, key_prefix):
    st.info(t("👆 Please click the row in the table below that contains the column headers (e.g., Entry Code, Score).", "👆 請直接點擊下方表格中，包含『參賽編號 / 分數』標題的那一行。"))
    try:
        event = st.dataframe(df_preview, on_select="rerun", selection_mode="single-row", use_container_width=True, key=f"df_{key_prefix}_{st.session_state.uploader_key}")
        if hasattr(event, "selection"):
            if event.selection.rows:
                return event.selection.rows[0]
            else:
                return None
    except TypeError:
        # Fallback to radio
        options = [t("None (Deselect)", "無 (取消選取)")]
        for i, row in df_preview.iterrows():
            preview_str = " | ".join([str(x) for x in row.values[:5]])
            options.append(f"Row {i}: {preview_str} ...")
        selected = st.radio(t("Select Header Row:", "請選擇包含標題的行 (Row Index)："), options, key=f"radio_{key_prefix}_{st.session_state.uploader_key}")
        if selected and selected != t("None (Deselect)", "無 (取消選取)"):
            return int(selected.split(":")[0].replace("Row ", ""))
    return None

def get_idx(options, val):
    try:
        return options.index(val) if val in options else 0
    except ValueError:
        return 0

def filter_unmatched(unmatched_set):
    clean = set()
    noise_kws = ["SIGN", "NAME", "DATE", "TOTAL", "REMARK", "JUDGE"]
    for item in unmatched_set:
        if pd.isna(item) or str(item).strip() == "" or str(item).strip().lower() in ["none", "nan"]:
            continue
        item_str = str(item).upper()
        if any(kw in item_str for kw in noise_kws):
            continue
        clean.add(item)
    return clean

# --- Main App ---
if "A" in app_mode:
    st.title(t("🚀 Mode A: Standard Template", "🚀 標準範本一鍵模式"))
    st.markdown(t("Download standard templates and upload filled versions for auto-processing.", "下載空白標準範本，讓評審直接填寫後上傳。本模式將全自動辨識範本欄位進行核對與整合。"))
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(t("📥 Download Judge Template", "📥 下載通用評審打分表範本"), generate_judge_template(), "Judge_Template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with col2:
        st.download_button(t("📥 Download Summary Template", "📥 下載通用成績匯總表範本"), generate_summary_template(), "Summary_Template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
    st.divider()
    
    st.header(t("Step 1: Upload Standard Templates", "步驟一：上傳標準範本檔案"))
    up_col1, up_col2 = st.columns(2)
    with up_col1:
        if st.session_state.uploaded_judges:
            st.success(t(f"✅ Loaded {len(st.session_state.uploaded_judges)} judge files", f"✅ 已載入 {len(st.session_state.uploaded_judges)} 份評審表"))
            if st.button(t("🔄 Re-upload Judge Sheets", "🔄 重新上傳評審表"), key="reset_judges_A"):
                st.session_state.uploaded_judges = []
                st.rerun()
        else:
            judge_files = st.file_uploader(t("1️⃣ Upload Judge Sheets (Left)", "1️⃣ 批量上傳「評審表」(左側)"), type=["xlsx"], accept_multiple_files=True, key=f"judges_A_{st.session_state.uploader_key}")
            if judge_files: 
                st.session_state.uploaded_judges = judge_files
                st.rerun()
            
    with up_col2:
        if st.session_state.uploaded_summary:
            st.success(t(f"✅ Loaded: {st.session_state.uploaded_summary.name}", f"✅ 已載入：{st.session_state.uploaded_summary.name}"))
            if st.button(t("🔄 Re-upload Summary Sheet", "🔄 重新上傳匯總表"), key="reset_sum_A"):
                st.session_state.uploaded_summary = None
                st.rerun()
        else:
            sum_file = st.file_uploader(t("2️⃣ Upload Summary Sheet (Right)", "2️⃣ 上傳「匯總表」(右側)"), type=["xlsx"], key=f"sum_A_{st.session_state.uploader_key}")
            if sum_file: 
                st.session_state.uploaded_summary = sum_file
                st.rerun()
            
    if st.session_state.uploaded_summary and st.session_state.uploaded_judges:
        st.header(t("Step 2: Execution", "步驟二：執行功能"))
        action = st.radio(t("Select Action", "選擇操作"), [t("🔍 Discrepancy Check", "🔍 差異化成績排查"), t("✨ Auto Integration", "✨ 自動整合匯出總表")], horizontal=True, key="action_A")
        
        if st.button(t("▶️ Execute", "▶️ 開始執行"), type="primary"):
            with st.spinner(t("Processing...", "處理中...")):
                judge_data = {}
                for j_file in st.session_state.uploaded_judges:
                    filename = j_file.name
                    if "匯總" in filename or "總表" in filename or "Summary" in filename:
                        continue
                    wb = load_workbook(BytesIO(j_file.getvalue()), data_only=True)
                    
                    j_dict = {}
                    for ws in wb:
                        entry_col_idx = None
                        score_col_idx = None
                        for c in range(1, ws.max_column + 1):
                            val = str(ws.cell(row=1, column=c).value).upper()
                            if "ENTRY CODE" in val: entry_col_idx = c
                            if "SCORE" in val: score_col_idx = c
                            
                        if entry_col_idx and score_col_idx:
                            for r in range(2, ws.max_row + 1):
                                code = ws.cell(row=r, column=entry_col_idx).value
                                score = ws.cell(row=r, column=score_col_idx).value
                                if code:
                                    j_dict[str(code).strip()] = safe_float(score)
                    judge_data[filename] = j_dict
                    
                wb_sum = load_workbook(BytesIO(st.session_state.uploaded_summary.getvalue()), data_only=("🔍" in action))
                
                # Collect sum entry codes for edge case warning
                sum_entry_codes = set()
                for ws_sum in wb_sum:
                    sum_entry_col = None
                    for c in range(1, ws_sum.max_column + 1):
                        val = str(ws_sum.cell(row=1, column=c).value).upper()
                        if "ENTRY CODE" in val: sum_entry_col = c
                    if sum_entry_col:
                        for r in range(2, ws_sum.max_row + 1):
                            code = ws_sum.cell(row=r, column=sum_entry_col).value
                            if code: sum_entry_codes.add(str(code).strip())
                            
                for fname, j_dict in judge_data.items():
                    unmatched = set(j_dict.keys()) - sum_entry_codes
                    clean_unmatched = filter_unmatched(unmatched)
                    if len(clean_unmatched) > 0:
                        st.warning(t(f"⚠️ Warning: Found entry codes in '{fname}' that don't exist in the Summary sheet, they will be ignored: {clean_unmatched}", f"⚠️ 警告：在 {fname} 中發現以下參賽號碼不在總表中，已被忽略：{clean_unmatched}"))
                
                checked_count = 0
                errors = []
                filled_count = 0
                
                for ws_sum in wb_sum:
                    sum_entry_col = None
                    sum_total_col = None
                    judge_cols = {}
                    
                    for c in range(1, ws_sum.max_column + 1):
                        val = str(ws_sum.cell(row=1, column=c).value)
                        val_upper = val.upper()
                        if "ENTRY CODE" in val_upper: sum_entry_col = c
                        elif "TOTAL" in val_upper: sum_total_col = c
                        elif "JUDGE" in val_upper:
                            judge_cols[val] = c
                            
                    if sum_entry_col:
                        for r in range(2, ws_sum.max_row + 1):
                            code = ws_sum.cell(row=r, column=sum_entry_col).value
                            if not code: continue
                            code_str = str(code).strip()
                            row_total = 0.0
                            
                            for j_name, c_idx in judge_cols.items():
                                j_feature = extract_judge_feature(j_name)
                                matched_f = next((f for f in judge_data if extract_judge_feature(f) == j_feature), None)
                                
                                if matched_f and code_str in judge_data[matched_f]:
                                    correct_score = judge_data[matched_f][code_str]
                                    
                                    if "🔍" in action:
                                        sum_score = safe_float(ws_sum.cell(row=r, column=c_idx).value)
                                        if sum_score is not None and correct_score is not None:
                                            checked_count += 1
                                            if abs(sum_score - correct_score) > 0.01:
                                                errors.append({t("Sheet", "分頁"): ws_sum.title, t("Entry Code", "參賽號"): code_str, t("Judge File", "評審"): matched_f, t("Summary Value", "總表數值"): sum_score, t("Correct Value", "正確數值"): correct_score})
                                                ws_sum.cell(row=r, column=c_idx).fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                                        else:
                                            checked_count += 1
                                            errors.append({t("Sheet", "分頁"): ws_sum.title, t("Entry Code", "參賽號"): code_str, t("Judge File", "評審"): matched_f, t("Summary Value", "總表數值"): sum_score, t("Correct Value", "正確數值"): correct_score})
                                            ws_sum.cell(row=r, column=c_idx).fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                                            
                                    elif "✨" in action:
                                        if correct_score is not None:
                                            ws_sum.cell(row=r, column=c_idx).value = correct_score
                                            filled_count += 1
                                
                                s_val = safe_float(ws_sum.cell(row=r, column=c_idx).value)
                                if s_val is not None: row_total += s_val
                                    
                            if "✨" in action and sum_total_col:
                                ws_sum.cell(row=r, column=sum_total_col).value = row_total
                
                if "🔍" in action:
                    if checked_count == 0:
                        st.error(t("❌ Error: 0 records checked! Please ensure you use standard templates.", "❌ 嚴重錯誤：成功核對筆數為 0！請確認使用的是標準範本。"))
                    else:
                        st.success(t(f"Check completed! Checked {checked_count} records. Found {len(errors)} errors.", f"比對完成！成功核對 {checked_count} 筆。發現 {len(errors)} 筆錯誤。"))
                        if errors:
                            st.dataframe(pd.DataFrame(errors), use_container_width=True)
                            out = BytesIO()
                            wb_sum.save(out)
                            st.download_button(t("📥 Download Error Highlight Report", "📥 下載錯誤標記報表 (.xlsx)"), out.getvalue(), "Error_Highlight_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
                            st.info(t("✅ Perfect Match!", "✅ 恭喜！完全一致！"))
                            st.balloons()
                else:
                    st.success(t(f"✅ Integration complete! Inserted {filled_count} scores.", f"✅ 整合成功！共填入了 {filled_count} 筆分數。"))
                    out = BytesIO()
                    wb_sum.save(out)
                    st.download_button(t("📥 Download Integrated Summary", "📥 下載整合結果"), out.getvalue(), "Integrated_Summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Mode B: Custom Mapping ---
elif "B" in app_mode:
    st.title(t("⚙️ Mode B: Custom Data Mapping UI", "⚙️ 自訂表格萬用模式 (Data Mapping UI)"))
    st.markdown(t("Upload your custom tables and map columns visually to resolve compatibility issues!", "上傳您的專屬表格，並透過 UI 視覺化關聯欄位與檔案對應，無痛相容所有格式！"))
    
    with st.expander(t("📂 Custom Template Config (Import/Export)", "📂 專屬範本設定 (載入/儲存)")):
        uploaded_json = st.file_uploader(t("Upload your custom config (.json)", "上傳專屬範本 (.json)"), type=["json"], key="config_uploader")
        if uploaded_json:
            try:
                config = json.load(uploaded_json)
                st.session_state.config_data = config
                if "judge_header_row" in config and st.session_state.judge_header_row is None:
                    st.session_state.judge_header_row = config["judge_header_row"]
                if "sum_header_row" in config and st.session_state.sum_header_row is None:
                    st.session_state.sum_header_row = config["sum_header_row"]
                st.success(t("✅ Config loaded! Fields have been set.", "✅ 專屬範本載入成功！欄位已自動設定。"))
            except Exception as e:
                st.error(t(f"Invalid JSON: {e}", f"無效的 JSON 檔案：{e}"))
    
    st.header(t("Step 1: Upload Files", "步驟一：檔案上傳"))
    col1, col2 = st.columns(2)
    with col1:
        if st.session_state.uploaded_judges:
            st.success(t(f"✅ Loaded {len(st.session_state.uploaded_judges)} judge files", f"✅ 已載入 {len(st.session_state.uploaded_judges)} 份評審表"))
            if st.button(t("🔄 Re-upload Judge Sheets", "🔄 重新上傳評審表"), key="reset_judges_B"):
                st.session_state.uploaded_judges = []
                st.rerun()
        else:
            judge_files = st.file_uploader(t("1️⃣ Upload Judge Sheets (Left)", "1️⃣ 批量上傳「評審表」(左側)"), type=["xlsx"], accept_multiple_files=True, key=f"judges_B_{st.session_state.uploader_key}")
            if judge_files: 
                st.session_state.uploaded_judges = judge_files
                st.rerun()
            
    with col2:
        if st.session_state.uploaded_summary:
            st.success(t(f"✅ Loaded: {st.session_state.uploaded_summary.name}", f"✅ 已載入：{st.session_state.uploaded_summary.name}"))
            if st.button(t("🔄 Re-upload Summary Sheet", "🔄 重新上傳匯總表"), key="reset_sum_B"):
                st.session_state.uploaded_summary = None
                st.rerun()
        else:
            sum_file = st.file_uploader(t("2️⃣ Upload Summary Sheet (Right)", "2️⃣ 上傳「匯總表」(右側)"), type=["xlsx"], key=f"sum_B_{st.session_state.uploader_key}")
            if sum_file: 
                st.session_state.uploaded_summary = sum_file
                st.rerun()
        
    if not st.session_state.uploaded_summary or not st.session_state.uploaded_judges:
        st.warning(t("⚠️ Please upload both Judge sheets and a Summary sheet to enable mapping analysis.", "⚠️ 請先完整上傳「評審表」與「匯總表」以進行欄位分析與對應。"))
    else:
        st.header(t("Step 2: Interactive Header Selection", "步驟二：互動式標題行選擇 (資料視覺化對應)"))
        
        preview_col1, preview_col2 = st.columns(2)
        
        with preview_col1:
            st.markdown(t("### 📝 Judge Sheet Preview", "### 📝 評審表預覽與設定"))
            try:
                xls_j = pd.ExcelFile(st.session_state.uploaded_judges[0])
                df_judge_preview = xls_j.parse(0, header=None, nrows=10)
                st.session_state.judge_header_row = interactive_header_selector(df_judge_preview, "judge")
                
                if st.session_state.judge_header_row is not None:
                    st.success(t(f"✅ Judge header row locked: Row {st.session_state.judge_header_row}", f"✅ 已鎖定評審表標題行：第 {st.session_state.judge_header_row} 行"))
            except Exception as e:
                st.error(f"Preview Error: {e}")
                
        with preview_col2:
            st.markdown(t("### 📝 Summary Sheet Preview", "### 📝 匯總表預覽與設定"))
            try:
                xls_s = pd.ExcelFile(st.session_state.uploaded_summary)
                df_sum_preview = xls_s.parse(0, header=None, nrows=10)
                st.session_state.sum_header_row = interactive_header_selector(df_sum_preview, "sum")
                    
                if st.session_state.sum_header_row is not None:
                    st.success(t(f"✅ Summary header row locked: Row {st.session_state.sum_header_row}", f"✅ 已鎖定匯總表標題行：第 {st.session_state.sum_header_row} 行"))
            except Exception as e:
                st.error(f"Preview Error: {e}")
                
        st.divider()
        
        # Progressive Disclosure: Only show the rest if both headers are locked
        if st.session_state.judge_header_row is not None and st.session_state.sum_header_row is not None:
            with st.container():
                try:
                    df_judge = pd.read_excel(st.session_state.uploaded_judges[0], header=st.session_state.judge_header_row)
                    raw_judge_cols = df_judge.columns.tolist()
                    judge_cols = clean_columns(raw_judge_cols)
                    
                    df_sum = pd.read_excel(st.session_state.uploaded_summary, header=st.session_state.sum_header_row)
                    raw_sum_cols = df_sum.columns.tolist()
                    sum_cols = clean_columns(raw_sum_cols)
                    
                    conf = st.session_state.config_data
                    
                    st.markdown(t("#### 🅰️ Basic Column Binding", "#### 🅰️ 基本欄位綁定"))
                    bind_col1, bind_col2 = st.columns(2)
                    with bind_col1:
                        st.markdown(t("**For Judge Sheets:**", "**針對「評審表」**"))
                        j_entry = st.selectbox(t("Which column is the Entry Code?", "參賽編號是哪一欄？"), judge_cols, index=get_idx(judge_cols, conf.get("j_entry")), key="j_entry")
                        j_score = st.selectbox(t("Which column is the Score?", "總分是哪一欄？"), judge_cols, index=get_idx(judge_cols, conf.get("j_score")), key="j_score")
                    with bind_col2:
                        st.markdown(t("**For Summary Sheet:**", "**針對「匯總表」**"))
                        s_entry = st.selectbox(t("Which column is the Entry Code?", "參賽編號是哪一欄？"), sum_cols, index=get_idx(sum_cols, conf.get("s_entry")), key="s_entry")
                        s_total = st.selectbox(t("Which column is the TOTAL score?", "最終總計 (TOTAL) 是哪一欄？"), sum_cols + [t("(None)", "(無)")], index=get_idx(sum_cols + [t("(None)", "(無)")], conf.get("s_total")), key="s_total")
                        
                    st.markdown(t("#### 🅱️ File to Judge Column Mapping", "#### 🅱️ 檔案與 Judge 欄位映射"))
                    st.info(t("System auto-matched files to columns based on names. Please review and adjust below if needed.", "系統已嘗試自動為您將檔案設置至總表的欄位進行判斷。若有錯誤，請展開下方區塊手動調整！"))
                    
                    with st.expander(t("Expand Mapping Settings", "展開檔案與欄位對應清單"), expanded=False):
                        potential_judges = [c for c in sum_cols if c not in [s_entry, s_total, t("(None)", "(無)")]]
                        potential_judges_options = [t("(Do not import)", "(不匯入)")] + potential_judges
                        
                        for idx, jf in enumerate(st.session_state.uploaded_judges):
                            fname = jf.name
                            if "匯總" in fname or "Semi Final" in fname: continue
                            
                            f_feature = extract_judge_feature(fname)
                            best_match = potential_judges_options[0]
                            for pj in potential_judges:
                                if extract_judge_feature(pj) == f_feature:
                                    best_match = pj
                                    break
                                    
                            best_match = conf.get("mapping", {}).get(fname, best_match)
                                    
                            st.session_state.mode_b_mapping[fname] = st.selectbox(
                                t(f"📁 Match `{fname}` to Summary column:", f"📁 `{fname}` 應寫入/比對至總表的："),
                                potential_judges_options,
                                index=get_idx(potential_judges_options, best_match),
                                key=f"map_{idx}_{st.session_state.uploader_key}"
                            )
                            
                    # Export JSON Config
                    mapping_to_save = {}
                    for jf in st.session_state.uploaded_judges:
                        fn = jf.name
                        if fn in st.session_state.mode_b_mapping:
                            mapping_to_save[fn] = st.session_state.mode_b_mapping[fn]
                            
                    current_config = {
                        "judge_header_row": st.session_state.judge_header_row,
                        "sum_header_row": st.session_state.sum_header_row,
                        "j_entry": j_entry,
                        "j_score": j_score,
                        "s_entry": s_entry,
                        "s_total": s_total,
                        "mapping": mapping_to_save
                    }
                    config_json = json.dumps(current_config, ensure_ascii=False, indent=2)
                    st.download_button(t("💾 Download Current Config", "💾 下載當前設定為專屬範本 (.json)"), config_json, "template_config.json", "application/json")
                    st.info(t("💡 Hint: Save this .json file. Next time, upload it at the top to auto-fill these settings!", "💡 提示：下載此 .json 檔案妥善保存。下次辦比賽時，只要在上方的『專屬範本設定』上傳此檔案，所有下拉選單就會瞬間為您自動填好！"))
                            
                    st.header(t("Step 3: Execution", "步驟三：執行功能"))
                    action = st.radio(t("Select Action", "選擇操作"), [t("🔍 Discrepancy Check", "🔍 差異化成績排查"), t("✨ Auto Integration", "✨ 自動整合匯出總表")], horizontal=True, key="action_B")
                    
                    if st.button(t("🚀 Run Mapping Task", "🚀 執行綁定任務"), type="primary"):
                        with st.spinner(t("Processing...", "處理中...")):
                            
                            judge_data = {}
                            for jf in st.session_state.uploaded_judges:
                                fname = jf.name
                                if st.session_state.mode_b_mapping.get(fname, t("(Do not import)", "(不匯入)")) in ["(Do not import)", "(不匯入)"]:
                                    continue
                                    
                                wb_j = load_workbook(BytesIO(jf.getvalue()), data_only=True)
                                j_dict = {}
                                
                                for ws_j in wb_j:
                                    j_entry_idx = None
                                    j_score_idx = None
                                    for c in range(1, ws_j.max_column + 1):
                                        val = str(ws_j.cell(row=st.session_state.judge_header_row + 1, column=c).value)
                                        clean_val = val
                                        if val == "None" or "Unnamed" in val:
                                            clean_val = t(f"Empty Column (Col {c})", f"空白欄位 (Column {c})")
                                            
                                        if clean_val == j_entry or val == j_entry: j_entry_idx = c
                                        if clean_val == j_score or val == j_score: j_score_idx = c
                                        
                                    if j_entry_idx and j_score_idx:
                                        for r in range(st.session_state.judge_header_row + 2, ws_j.max_row + 1):
                                            code = ws_j.cell(row=r, column=j_entry_idx).value
                                            score = ws_j.cell(row=r, column=j_score_idx).value
                                            if code:
                                                j_dict[str(code).strip()] = safe_float(score)
                                judge_data[fname] = j_dict
                                
                            wb_sum = load_workbook(BytesIO(st.session_state.uploaded_summary.getvalue()), data_only=("🔍" in action))
                            
                            # Collect sum entry codes for edge case warning
                            sum_entry_codes = set()
                            for ws_sum in wb_sum:
                                s_entry_idx = None
                                for c in range(1, ws_sum.max_column + 1):
                                    val = str(ws_sum.cell(row=st.session_state.sum_header_row + 1, column=c).value)
                                    clean_val = val
                                    if val == "None" or "Unnamed" in val:
                                        clean_val = t(f"Empty Column (Col {c})", f"空白欄位 (Column {c})")
                                    if clean_val == s_entry or val == s_entry: s_entry_idx = c
                                if s_entry_idx:
                                    for r in range(st.session_state.sum_header_row + 2, ws_sum.max_row + 1):
                                        code = ws_sum.cell(row=r, column=s_entry_idx).value
                                        if code: sum_entry_codes.add(str(code).strip())

                            for fname, j_dict in judge_data.items():
                                unmatched = set(j_dict.keys()) - sum_entry_codes
                                clean_unmatched = filter_unmatched(unmatched)
                                if len(clean_unmatched) > 0:
                                    st.warning(t(f"⚠️ Warning: Found entry codes in '{fname}' that don't exist in the Summary sheet, they will be ignored: {clean_unmatched}", f"⚠️ 警告：在 {fname} 中發現以下參賽號碼不在總表中，已被忽略：{clean_unmatched}"))
                            
                            checked_count = 0
                            errors = []
                            filled_count = 0
                            
                            for ws_sum in wb_sum:
                                s_entry_idx = None
                                s_total_idx = None
                                s_judge_indices = {}
                                
                                for c in range(1, ws_sum.max_column + 1):
                                    val = str(ws_sum.cell(row=st.session_state.sum_header_row + 1, column=c).value)
                                    clean_val = val
                                    if val == "None" or "Unnamed" in val:
                                        clean_val = t(f"Empty Column (Col {c})", f"空白欄位 (Column {c})")
                                        
                                    if clean_val == s_entry or val == s_entry: s_entry_idx = c
                                    if clean_val == s_total or val == s_total: s_total_idx = c
                                    if clean_val in potential_judges or val in potential_judges:
                                        s_judge_indices[clean_val] = c
                                        s_judge_indices[val] = c
                                        
                                if not s_entry_idx: continue 
                                
                                for r in range(st.session_state.sum_header_row + 2, ws_sum.max_row + 1):
                                    code = ws_sum.cell(row=r, column=s_entry_idx).value
                                    if not code: continue
                                    code_str = str(code).strip()
                                    
                                    row_total = 0.0
                                    
                                    for fname, target_sum_col in st.session_state.mode_b_mapping.items():
                                        if target_sum_col in ["(Do not import)", "(不匯入)"] or target_sum_col not in s_judge_indices: continue
                                        c_idx = s_judge_indices[target_sum_col]
                                        
                                        correct_score = judge_data.get(fname, {}).get(code_str)
                                        sum_score_raw = ws_sum.cell(row=r, column=c_idx).value
                                        sum_score = safe_float(sum_score_raw)
                                        
                                        if "🔍" in action:
                                            if sum_score_raw is None and correct_score is None: continue
                                            if str(sum_score_raw).strip() == "" and correct_score is None: continue
                                            
                                            if sum_score is not None and correct_score is not None:
                                                checked_count += 1
                                                if abs(sum_score - correct_score) > 0.01:
                                                    errors.append({t("Sheet", "分頁"): ws_sum.title, t("Entry Code", "參賽號"): code_str, t("Judge File", "評審檔案"): fname, t("Summary Value", "總表數值"): sum_score, t("Correct Value", "正確數值"): correct_score})
                                                    ws_sum.cell(row=r, column=c_idx).fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                                            else:
                                                checked_count += 1
                                                errors.append({t("Sheet", "分頁"): ws_sum.title, t("Entry Code", "參賽號"): code_str, t("Judge File", "評審檔案"): fname, t("Summary Value", "總表數值"): sum_score_raw if sum_score_raw is not None else t('Empty', '空白'), t("Correct Value", "正確數值"): correct_score if correct_score is not None else t('Invalid/Empty', '無效/空白')})
                                                ws_sum.cell(row=r, column=c_idx).fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                                        
                                        elif "✨" in action:
                                            if correct_score is not None:
                                                ws_sum.cell(row=r, column=c_idx).value = correct_score
                                                filled_count += 1
                                                
                                    if "✨" in action:
                                        for pj in potential_judges:
                                            if pj in s_judge_indices:
                                                pj_idx = s_judge_indices[pj]
                                                s_val = safe_float(ws_sum.cell(row=r, column=pj_idx).value)
                                                if s_val is not None:
                                                    row_total += s_val
                                        if s_total_idx:
                                            ws_sum.cell(row=r, column=s_total_idx).value = row_total
                                            
                            if "🔍" in action:
                                if checked_count == 0:
                                    st.error(t("❌ Error: 0 records checked! Please review mapping settings and header rows.", "❌ 嚴重錯誤：成功核對筆數為 0！請確認 Mapping 選項與標題列設定是否正確。"))
                                else:
                                    st.success(t(f"Check completed! Checked {checked_count} records. Found {len(errors)} errors.", f"比對完成！成功核對 {checked_count} 筆。發現 {len(errors)} 筆錯誤。"))
                                    if errors:
                                        st.dataframe(pd.DataFrame(errors), use_container_width=True)
                                        out = BytesIO()
                                        wb_sum.save(out)
                                        st.download_button(t("📥 Download Error Highlight Report", "📥 下載錯誤標記報表 (.xlsx)"), out.getvalue(), "Error_Highlight_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                                    else:
                                        st.info(t("✅ Perfect Match!", "✅ 恭喜！完全一致！"))
                                        st.balloons()
                            else:
                                st.success(t(f"✅ Integration complete! Inserted {filled_count} scores.", f"✅ 整合成功！共精準填入了 {filled_count} 筆分數。"))
                                out = BytesIO()
                                wb_sum.save(out)
                                st.download_button(t("📥 Download Custom Integrated Summary", "📥 一鍵下載客製化整合結果"), out.getvalue(), "Custom_Integrated_Summary.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(t(f"Parsing error: {e}. Ensure rows are correct.", f"解析過程中發生錯誤：{e}。請確保所選的行數正確且檔案格式支援。"))
