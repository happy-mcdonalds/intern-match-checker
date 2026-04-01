import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 高級感 CSS (宋體 + 黑白灰) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', serif !important;
        color: #000000;
    }
    h1, h2, h3 { color: #000000 !important; border-bottom: 1px solid #000000; padding-bottom: 5px; }
    .stApp { background-color: #FFFFFF; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #DDDDDD; }
    .stButton>button { color: #FFFFFF !important; background-color: #000000 !important; border-radius: 0px; width: 100%; }
    </style>
    """, unsafe_allow_html=True)

# --- 聰明讀取函式 ---
def smart_read_excel(file):
    """ 自動讀取第一個分頁，並嘗試找到包含『姓名』或『科別』的標題列 """
    try:
        # 讀取整個 Excel 檔的所有分頁名稱
        xls = pd.ExcelFile(file)
        # 優先找包含「志願」或「名單」的分頁，若無則抓第一個
        sheet_names = xls.sheet_names
        target_sheet = sheet_names[0]
        for sn in sheet_names:
            if "志願" in sn or "名單" in sn:
                target_sheet = sn
                break
        
        # 讀取該分頁
        df = pd.read_excel(file, sheet_name=target_sheet)
        
        # 偵測真正的標題列 (有些 Excel 前幾行是空白或公告)
        for i in range(len(df)):
            row_values = [str(x) for x in df.iloc[i].values]
            if "姓名" in row_values or "申請科別" in row_values or "科別" in row_values:
                df = pd.read_excel(file, sheet_name=target_sheet, header=i+1)
                break
        return df
    except Exception as e:
        st.error(f"讀取 {file.name} 失敗: {e}")
        return None

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表模式 (容額校對)", "總代模式 (跨院比對)"])
st.sidebar.divider()
course_duration_weeks = st.sidebar.number_input("一個 Course 多久 (週)", min_value=1, value=2)
min_weeks_req = st.sidebar.number_input("最短實習週數要求 (週)", min_value=1, value=4)
require_cont = st.sidebar.checkbox("要求必須連續實習", value=True)

# --- 模式：醫院代表 ---
if mode == "醫院代表 (容額校對)":
    st.title("醫院內部容額與規章審核")
    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("1. 上傳醫院容額表", type=['xlsx'])
    with c2: a_file = st.file_uploader("2. 上傳學生志願表", type=['xlsx'])

    if q_file and a_file:
        df_q = smart_read_excel(q_file)
        df_a = smart_read_excel(a_file)
        
        if df_q is not None and df_a is not None:
            # (此處接續之前的容額判定與週數判定邏輯...)
            st.success("檔案讀取成功，正在分析內容...")
            # 這裡就不重複貼上長長的邏輯代碼，請保留你原本的運算部分

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    st.markdown("此模式專供系秘比對不同醫院間是否有學生重複佔位。")
    
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    
    if multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None:
                df.columns = [str(c).strip() for c in df.columns]
                if '姓名' in df.columns:
                    df['姓名'] = df['姓名'].ffill()
                    df['來源醫院'] = f.name
                    all_data.append(df[df['申請科別'].notna()])
        
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            # 跨檔案日期重疊判斷
                            d1_s, d1_e, _ = parse_dates(s_apps[i]['實習期間'])
                            d2_s, d2_e, _ = parse_dates(s_apps[j]['實習期間'])
                            if d1_s and d2_s and (d1_s <= d2_e and d2_s <= d1_e):
                                conflicts.append({
                                    "姓名": name,
                                    "醫院 A": s_apps[i]['來源醫院'], "時段 A": s_apps[i]['實習期間'],
                                    "醫院 B": s_apps[j]['來源醫院'], "時段 B": s_apps[j]['實習期間']
                                })
            
            if conflicts:
                st.subheader("偵測到跨院衝突名單")
                st.table(pd.DataFrame(conflicts).drop_duplicates())
            else:
                st.success("交叉比對完成，無重複佔位情況。")
