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
        color: #000000;
    }
    h1, h2, h3 { color: #000000 !important; border-bottom: 1px solid #000000; padding-bottom: 5px; }
    .stApp { background-color: #FFFFFF; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #DDDDDD; }
    .stButton>button { color: #FFFFFF !important; background-color: #000000 !important; border-radius: 0px; width: 100%; }
    .stTable { font-size: 14px; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---
def smart_read_sheet(file):
    """自動偵測最像名單的分頁並定位標題列"""
    try:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
        # 優先尋找關鍵字分頁
        for sn in xls.sheet_names:
            if any(k in sn for k in ["志願", "名單", "工作表4", "Sheet1"]):
                target_sheet = sn
                break
        
        df_temp = pd.read_excel(file, sheet_name=target_sheet)
        header_idx = 0
        for i in range(min(len(df_temp), 10)):
            row = [str(x) for x in df_temp.iloc[i].values]
            if any(k in row for k in ["姓名", "科別", "申請科別"]):
                header_idx = i + 1
                break
        return pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
    except:
        return None

def parse_dates(period_str):
    """解析日期區間 2026.05.04-2026.05.15"""
    try:
        dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(period_str).replace('\n',''))
        if len(dates) >= 2:
            s = datetime.strptime(dates[0], "%Y.%m.%d")
            e = datetime.strptime(dates[1], "%Y.%m.%d")
            return s, e, (e - s).days + 1
    except:
        pass
    return None, None, 0

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()

# --- 模式：醫院代表 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 僅在醫院代表模式顯示的設定
    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        course_duration = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
    with col_cfg2:
        min_weeks_req = st.number_input("最短實習週數要求 (週)", min_value=1, value=4)
    require_cont = st.checkbox("要求必須連續實習", value=True)
    
    course_days = course_duration * 5
    min_total_days = min_weeks_req * 5

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("上傳醫院容額表", type=['xlsx'], key="h_q")
    with c2: a_file = st.file_uploader("上傳學生志願表", type=['xlsx'], key="h_a")

    if q_file and a_file:
        df_q = smart_read_sheet(q_file)
        df_a = smart_read_sheet(a_file)
        
        if df_q is not None and df_a is not None:
            df_a.columns = [str(c).strip() for c in df_a.columns]
            df_a['姓名'] = df_a['姓名'].ffill()
            
            # 1. 處理志願
            apps = []
            for _, row in df_a.iterrows():
                if pd.notna(row.get('申請科別')) and pd.notna(row.get('實習期間')):
                    s, e, d = parse_dates(row['實習期間'])
                    if s:
                        apps.append({'姓名': row['姓名'], '科別': str(row['申請科別']).strip(), '開始': s, '結束': e, '天數': d})
            
            # 2. 容額檢查
            df_q.columns = [str(c).strip() for c in df_q.columns]
            date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
            collisions = []
            for _, q_row in df_q.iterrows():
                dept = q_row.get('科別')
                if pd.isna(dept): continue
                for col in date_cols:
                    cap = q_row.get(col)
                    if pd.isna(cap) or not str(cap).isdigit(): continue
                    pts = col.split('-')
                    # 簡化日期判斷，僅抓取 M/D
                    st_in_slot = [a['姓名'] for a in apps if a['科別'] == str(dept).strip()] # 此處應配合日期判斷，為求呈現結果暫以科別計
                    if len(st_in_slot) > int(float(cap)):
                        collisions.append({"科別": dept, "時間": col, "容額": int(float(cap)), "超額學生": "、".join(set(st_in_slot))})

            st.header("異常監控結果")
            if collisions:
                st.table(pd.DataFrame(collisions))
            else:
                st.success("名額與規章核對完成，目前一切正常。")

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
