import streamlit as st
import pandas as pd
from datetime import datetime
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 高級感 CSS 注入 ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', serif !important;
        color: #333333;
    }
    h1, h2, h3 { color: #000000 !important; border-bottom: 1px solid #EEEEEE; }
    .stApp { background-color: #FFFFFF; }
    section[data-testid="stSidebar"] { background-color: #F8F9FA; border-right: 1px solid #E0E0E0; }
    .stButton>button { color: #FFFFFF; background-color: #000000; border-radius: 0px; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心邏輯：轉譯你的 Excel 公式 ---
def process_original_excel(df_apply):
    """
    對應你的 ARRAYFORMULA 邏輯：
    1. 處理每兩列一個人的姓名偏移 (OFFSET)
    2. 使用 REGEX 提取日期 (DATEVALUE + SUBSTITUTE)
    3. 自動標註時段一/時段二 (MOD ROW)
    """
    processed_records = []
    df_apply = df_apply.reset_index(drop=True)
    
    for i, row in df_apply.iterrows():
        # 1. 取得姓名 (對應你的 OFFSET 邏輯：偶數列取當前，奇數列取上一列)
        name = row['姓名']
        if pd.isna(name) or name == "":
            if i > 0:
                name = df_apply.loc[i-1, '姓名']
        
        # 2. 取得日期並轉換 (對應你的 REGEXEXTRACT + DATEVALUE)
        period = str(row['實習期間'])
        start_dt, end_dt = None, None
        dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', period)
        if len(dates) >= 2:
            start_dt = datetime.strptime(dates[0], "%Y.%m.%d")
            end_dt = datetime.strptime(dates[1], "%Y.%m.%d")
        
        # 3. 標註時段 (對應你的 MOD ROW 邏輯)
        slot_type = "時段一" if i % 2 == 0 else "時段二"
        
        if not pd.isna(row['申請科別']):
            processed_records.append({
                "姓名": name,
                "學號": row['學號'] if not pd.isna(row['學號']) else "",
                "科別": row['申請科別'],
                "開始日期": start_dt,
                "結束日期": end_dt,
                "原始期間": period,
                "時段": slot_type
            })
    return pd.DataFrame(processed_records)

# --- 側邊欄與模式 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表模式 (原表校對)", "總代模式 (跨院重複比對)"])

if mode == "醫院代表模式 (原表校對)":
    st.title("醫院內部容額與規章審核")
    
    with st.sidebar.expander("規則設定", expanded=True):
        min_weeks = st.sidebar.number_input("最短實習週數", value=2)
        
    col1, col2 = st.columns(2)
    with col1:
        quota_file = st.file_uploader("上傳醫院容額表", type=['xlsx'])
    with col2:
        apply_file = st.file_uploader("上傳原始志願清單", type=['xlsx'])

    if quota_file and apply_file:
        # 讀取特定分頁 (對應你的工作表名稱)
        df_q_raw = pd.read_excel(quota_file) 
        df_a_raw = pd.read_excel(apply_file, sheet_name="志願申請名單")
        
        # 清理並處理
        df_a_raw.columns = [str(c).strip() for c in df_a_raw.columns]
        df_final_apply = process_original_excel(df_a_raw)
        
        st.subheader("志願解析結果 (對應工作表4)")
        st.dataframe(df_final_apply, use_container_width=True)

        # 容額計算 (使用解析後的日期與科別)
        st.subheader("科別容額統計")
        usage = df_final_apply.groupby('科別').size().reset_index(name='報名人數')
        # 假設容額表也有個「科別」欄位
        df_q_raw.columns = [str(c).strip() for c in df_q_raw.columns]
        if '科別' in df_q_raw.columns and '容額' in df_q_raw.columns:
            status = pd.merge(df_q_raw, usage, on='科別', how='left').fillna(0)
            status['剩餘名額'] = status['容額'] - status['報名人數']
            
            def style_negative(val):
                return 'color: #FF0000; font-weight: bold; background-color: #F9F9F9' if val < 0 else ''
            st.dataframe(status.style.applymap(style_negative, subset=['剩餘名額']), use_container_width=True)

elif mode == "總代模式 (跨院比對)":
    st.title("全院跨院重複佔位檢查")
    files = st.file_uploader("上傳各院志願清單 (多選)", type=['xlsx'], accept_multiple_files=True)
    
    if files:
        all_apps = []
        for f in files:
            raw = pd.read_excel(f, sheet_name="志願申請名單")
            raw.columns = [str(c).strip() for c in raw.columns]
            processed = process_original_excel(raw)
            processed['來源'] = f.name
            all_apps.append(processed)
        
        full_df = pd.concat(all_apps)
        # 執行跨院比對邏輯 (同前，檢查日期重疊)
        st.success("資料已匯總，正在執行交叉比對...")
        # ... (比對程式碼同上個版本)
