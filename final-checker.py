import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 注入高級感 CSS (宋體 + 黑白灰) ---
st.markdown("""
    <style>
    /* 載入宋體並設定全站字體 */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', 'Source Han Serif TC', 'STSong', 'SimSun', serif !important;
        color: #333333;
    }
    
    /* 背景與標題顏色 */
    .stApp {
        background-color: #FFFFFF;
    }
    
    h1, h2, h3 {
        color: #000000 !important;
        font-weight: 700 !important;
        border-bottom: 1px solid #EEEEEE;
        padding-bottom: 10px;
    }

    /* 側邊欄樣式優化 */
    section[data-testid="stSidebar"] {
        background-color: #F8F9FA;
        border-right: 1px solid #E0E0E0;
    }
    
    /* 按鈕樣式：黑白極簡 */
    .stButton>button {
        color: #FFFFFF;
        background-color: #000000;
        border-radius: 0px;
        border: 1px solid #000000;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #333333;
        border-color: #333333;
        color: #FFFFFF;
    }

    /* 表格樣式優化 */
    .stDataFrame {
        border: 1px solid #EEEEEE;
    }

    /* 隱藏預設的 Emoji 或裝飾 */
    </style>
    """, unsafe_allow_html=True)

# --- 1. 工具函式庫 ---
def parse_date(d):
    try:
        clean_d = str(d).replace('/', '.').replace('\n', '').strip()
        if len(clean_d.split('.')) == 2: clean_d = "2026." + clean_d
        return datetime.strptime(clean_d, "%Y.%m.%d")
    except: return None

def is_overlap(range1, range2):
    try:
        r1_s, r1_e = [parse_date(x) for x in str(range1).split('-')]
        r2_s, r2_e = [parse_date(x) for x in str(range2).split('-')]
        if None in [r1_s, r1_e, r2_s, r2_e]: return False
        return r1_s <= r2_e and r2_s <= r1_e
    except: return False

# --- 2. 側邊欄模式切換 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表模式 (容額檢查)", "總代模式 (跨院比對)"])

st.sidebar.markdown("---")

# --- 3. 醫院代表模式 ---
if mode == "醫院代表模式 (容額檢查)":
    st.title("醫院內部容額與規則審查")
    st.markdown("針對單一醫院之申請名單進行規則校對與容額統計。")
    
    with st.sidebar.expander("規則設定", expanded=True):
        target_hosp = st.text_input("醫院名稱", value="國泰醫院")
        min_weeks = st.number_input("最短實習週數", min_value=1, value=2)
        require_cont = st.checkbox("要求連續實習", value=True)
        
    col1, col2 = st.columns(2)
    with col1:
        quota_file = st.file_uploader("上傳醫院容額表", type=['xlsx', 'csv'], key="q1")
    with col2:
        apply_file = st.file_uploader("上傳學生申請名單", type=['xlsx', 'csv'], key="a1")

    if quota_file and apply_file:
        df_q = pd.read_excel(quota_file) if quota_file.name.endswith('.xlsx') else pd.read_csv(quota_file)
        df_a = pd.read_excel(apply_file) if apply_file.name.endswith('.xlsx') else pd.read_csv(apply_file)
        df_q.columns = [c.strip() for c in df_q.columns]
        df_a.columns = [c.strip() for c in df_a.columns]

        def validate_student(row):
            try:
                s, e = [parse_date(x) for x in str(row['實習期間']).split('-')]
                weeks = ((e - s).days + 1) / 7
                if weeks < min_weeks: return f"不符: 週數不足({int(weeks)}週)"
                return "通過"
            except: return "格式錯誤"

        df_a['審查結果'] = df_a.apply(validate_student, axis=1)
        
        st.subheader("學生資格審查")
        st.dataframe(df_a[['姓名', '學號', '申請科別', '實習期間', '審查結果']], use_container_width=True)

        st.subheader("科別容額統計")
        usage = df_a[df_a['審查結果'] == "通過"].groupby('申請科別').size().reset_index(name='報名人數')
        status = pd.merge(df_q, usage, left_on='科別', right_on='申請科別', how='left').fillna(0)
        status['剩餘名額'] = status['容額'] - status['報名人數']
        
        # 使用灰階標色法
        def style_overflow(val):
            return 'background-color: #F0F0F0; color: #FF0000; font-weight: bold' if val < 0 else ''
        
        st.dataframe(status.style.applymap(style_overflow, subset=['剩餘名額']), use_container_width=True)

# --- 4. 總代模式 ---
elif mode == "總代模式 (跨院比對)":
    st.title("全院跨院重複佔位檢查")
    st.markdown("收集各院確定名單後進行交叉比對，找出時段重疊之申請。")
    
    files = st.file_uploader("上傳各院確定名單 (支援多選)", type=['xlsx', 'csv'], accept_multiple_files=True)
    
    if files:
        all_data = []
        for f in files:
            df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
            df.columns = [str(c).strip() for c in df.columns]
            df['來源醫院'] = f.name
            all_data.append(df)
        
        full_df = pd.concat(all_data, ignore_index=True)
        conflicts = []
        unique_ids = full_df['學號'].unique()
        
        for s_id in unique_ids:
            s_apps = full_df[full_df['學號'] == s_id].to_dict('records')
            if len(s_apps) > 1:
                for i in range(len(s_apps)):
                    for j in range(i + 1, len(s_apps)):
                        if is_overlap(s_apps[i]['實習期間'], s_apps[j]['實習期間']):
                            conflicts.append({
                                "姓名": s_apps[i]['姓名'], "學號": s_id,
                                "醫院A": s_apps[i]['來源醫院'], "時間A": s_apps[i]['實習期間'],
                                "醫院B": s_apps[j]['來源醫院'], "時間B": s_apps[j]['實習期間']
                            })
        
        if conflicts:
            st.markdown("### 衝突偵測結果")
            st.table(pd.DataFrame(conflicts))
        else:
            st.markdown("---")
            st.markdown("經交叉比對，目前未發現重複佔位情況。")
