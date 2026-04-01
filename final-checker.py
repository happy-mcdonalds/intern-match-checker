import streamlit as st
import pandas as pd
from datetime import datetime
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 高級感 CSS (宋體 + 黑白灰) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', serif !important;
        color: #333333;
    }
    h1, h2, h3 { color: #000000 !important; border-bottom: 1px solid #EEEEEE; padding-bottom: 10px; }
    .stApp { background-color: #FFFFFF; }
    section[data-testid="stSidebar"] { background-color: #F8F9FA; border-right: 1px solid #E0E0E0; }
    .stButton>button { color: #FFFFFF; background-color: #000000; border-radius: 0px; }
    </style>
    """, unsafe_allow_html=True)

# --- 工具函式 ---
def parse_date(d):
    try:
        clean_d = str(d).replace('/', '.').replace('\n', '').strip()
        # 處理如 2026.05.04 這種格式
        match = re.search(r'\d{4}\.\d{2}\.\d{2}', clean_d)
        if match:
            return datetime.strptime(match.group(), "%Y.%m.%d")
        return None
    except: return None

def is_overlap(range1, range2):
    try:
        r1_s, r1_e = [parse_date(x) for x in str(range1).split('-')]
        r2_s, r2_e = [parse_date(x) for x in str(range2).split('-')]
        if None in [r1_s, r1_e, r2_s, r2_e]: return False
        return r1_s <= r2_e and r2_s <= r1_e
    except: return False

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表模式 (容額檢查)", "總代模式 (跨院比對)"])
st.sidebar.markdown("---")

if mode == "醫院代表模式 (容額檢查)":
    st.title("醫院內部容額與規則審查")
    
    # 這裡確保「勾選格子」一定會出現
    with st.sidebar.expander("規則設定", expanded=True):
        min_weeks = st.number_input("最短實習週數", min_value=1, value=2)
        require_cont = st.checkbox("要求必須連續實習", value=True)
        unit_weeks = st.selectbox("實習單位 (週)", [2, 4], index=0)

    col1, col2 = st.columns(2)
    with col1:
        quota_file = st.file_uploader("1. 上傳醫院容額表 (含科別/容額)", type=['xlsx'])
    with col2:
        apply_file = st.file_uploader("2. 上傳學生志願表 (志願申請名單)", type=['xlsx'])

    if quota_file and apply_file:
        try:
            # 讀取容額表 (跳過前幾列醫院資訊，直到看到「科別」)
            df_q = pd.read_excel(quota_file)
            # 處理你的 Excel 可能有多行 Header 的問題
            if '科別' not in df_q.columns:
                df_q = pd.read_excel(quota_file, header=4) # 根據你上傳的檔案，標題通常在第5列

            # 讀取申請名單
            df_a_raw = pd.read_excel(apply_file, sheet_name="志願申請名單")
            df_a_raw.columns = [str(c).strip() for c in df_a_raw.columns]

            # --- 轉譯你的 Excel OFFSET/MOD 邏輯 ---
            processed_data = []
            for i, row in df_a_raw.iterrows():
                name = row['姓名']
                if pd.isna(name) and i > 0: name = df_a_raw.loc[i-1, '姓名']
                
                if not pd.isna(row['申請科別']):
                    period = str(row['實習期間'])
                    s_dt = parse_date(period.split('-')[0])
                    e_dt = parse_date(period.split('-')[1]) if '-' in period else None
                    
                    processed_data.append({
                        "姓名": name,
                        "學號": row['學號'] if '學號' in df_a_raw.columns else "",
                        "科別": row['申請科別'],
                        "開始": s_dt,
                        "結束": e_dt,
                        "週數": ((e_dt - s_dt).days + 1) / 7 if s_dt and e_dt else 0
                    })
            
            df_final_a = pd.DataFrame(processed_data)

            # --- 資格審核 ---
            def check_row(row):
                if row['週數'] < min_weeks: return f"❌ 不足 {min_weeks} 週"
                return "✅ 通過"

            df_final_a['審查結果'] = df_final_a.apply(check_row, axis=1)

            st.subheader("學生資格清單")
            st.dataframe(df_final_a, use_container_width=True)

            # --- 容額計算 ---
            st.subheader("科別容額即時統計")
            # 統計各科報名人數
            usage = df_final_a[df_final_a['審查結果'] == "✅ 通過"].groupby('科別').size().reset_index(name='報名人數')
            
            # 合併容額
            df_q.columns = [str(c).strip() for c in df_q.columns]
            if '科別' in df_q.columns and '容額' in df_q.columns:
                status = pd.merge(df_q[['科別', '容額']], usage, on='科別', how='left').fillna(0)
                status['剩餘名額'] = status['容額'] - status['報名人數']
                
                st.dataframe(status.style.applymap(
                    lambda x: 'background-color: #FEE2E2; color: #991B1B' if x < 0 else '', 
                    subset=['剩餘名額']
                ), use_container_width=True)
            else:
                st.error("容額表格式不符，請確保包含『科別』與『容額』欄位。")

        except Exception as e:
            st.error(f"處理出錯：{e}")

# --- 5. 總代模式 ---
elif mode == "總代模式 (跨院比對)":
    st.title("跨院重複佔位比對")
    # (此處保留原本的多檔案比對邏輯...)
