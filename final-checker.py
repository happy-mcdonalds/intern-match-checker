import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# --- 初始化系統記憶 ---
if "course_dur_weeks" not in st.session_state: st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state: st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state: st.session_state.require_cont = True

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- CSS 樣式 ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, .stApp { font-family: 'Noto Serif TC', serif !important; background-color: #F5F4F1; color: #5C5E5D; }
    h1, h2, h3 { color: #4A4C4B; border-bottom: 1px solid #D6D4CE; }
    [data-testid="stForm"] { border: 1px solid #D6D4CE; background-color: #FDFDFD; padding: 20px; }
    .stButton > button { background-color: #8A9A92 !important; color: white !important; }
    table { width: 100%; border-collapse: collapse; }
    th { background-color: #E3E1DB !important; padding: 10px; border-bottom: 2px solid #C0BFB8; }
    td { padding: 10px; border-bottom: 1px solid #EAE8E3; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---

def smart_read_sheet(file):
    """專門對應「申請名單.xlsx」格式的讀取函式"""
    try:
        xls = pd.ExcelFile(file)
        # 自動尋找正確的工作表
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if any(k in sn for k in ["志願", "名單", "工作表4"]):
                target_sheet = sn
                break
        
        # 讀取數據
        df = pd.read_excel(file, sheet_name=target_sheet)
        
        # 欄位清理邏輯
        clean_cols = {}
        for c in df.columns:
            s_c = str(c).strip()
            if "姓名" in s_c: clean_cols[c] = "姓名"
            elif "科別" in s_c: clean_cols[c] = "科別"
            elif "期間" in s_c or "時段" in s_c: clean_cols[c] = "實習期間"
            else: clean_cols[c] = s_c
        
        df = df.rename(columns=clean_cols)
        
        # 【關鍵優化】處理合併儲存格：如果姓名是空的，就拿上面的來填
        if "姓名" in df.columns:
            df["姓名"] = df["姓名"].ffill()
            
        # 移除姓名欄位完全沒東西的廢列
        df = df.dropna(subset=["姓名"])
        
        return df
    except Exception as e:
        st.error(f"讀取失敗: {e}")
        return None

def parse_period_dates(p_str):
    """解析如 2026.05.04-2026.05.15 的字串"""
    if pd.isna(p_str): return None, None, 0
    try:
        # 處理換行與空格
        clean_str = str(p_str).replace('\n', ' ').replace('\r', '').strip()
        # 尋找所有日期格式 (YYYY.MM.DD 或 YYYY-MM-DD)
        dates_found = re.findall(r'\d{4}[./-]\d{2}[./-]\d{2}', clean_str)
        
        if len(dates_found) >= 2:
            s_dt = datetime.strptime(dates_found[0].replace('.', '-'), '%Y-%m-%d')
            e_dt = datetime.strptime(dates_found[1].replace('.', '-'), '%Y-%m-%d')
            days = (e_dt - s_dt).days + 1
            return s_dt, e_dt, days
    except:
        pass
    return None, None, 0

# --- UI 介面 ---

st.title("醫學系實習選配管理系統")

mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])

if mode == "醫院代表":
    st.subheader("🏥 醫院端資料核對")
    
    col1, col2 = st.columns(2)
    with col1: q_file = st.file_uploader("1. 上傳醫院容額表 (xlsx)", type=['xlsx'])
    with col2: a_file = st.file_uploader("2. 上傳學生志願表 (xlsx)", type=['xlsx'])
    
    if st.button("開始核對"):
        if q_file and a_file:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)
            
            if df_a is not None and df_q is not None:
                # 檢查必要欄位
                required = ["姓名", "科別", "實習期間"]
                missing = [r for r in required if r not in df_a.columns]
                
                if missing:
                    st.error(f"志願表缺少欄位：{missing}")
                    st.write("目前偵測到的欄位有：", list(df_a.columns))
                else:
                    # 執行邏輯比對 (此處接續你原本的業務邏輯)
                    st.success("檔案讀取成功！已自動處理合併儲存格。")
                    st.dataframe(df_a[["姓名", "科別", "實習期間"]].head(10))
        else:
            st.warning("請上傳檔案")

elif mode == "系秘":
    st.subheader("🎓 跨院重複佔位檢查")
    multi_files = st.file_uploader("上傳多家醫院清單", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("執行比對") and multi_files:
        # 跨院比對邏輯...
        st.info("比對功能執行中...")
