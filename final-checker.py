import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
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
    try:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if any(k in sn for k in ["志願", "名單", "工作表4", "實習容額"]):
                target_sheet = sn
                break
        df_temp = pd.read_excel(file, sheet_name=target_sheet)
        header_idx = 0
        for i in range(min(len(df_temp), 15)):
            row = [str(x).strip() for x in df_temp.iloc[i].values]
            if any(k in row for k in ["姓名", "科別", "申請科別"]):
                header_idx = i + 1
                break
        df = pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except: return None

def parse_date_simple(s, year=2026):
    try:
        parts = re.findall(r'\d+', str(s))
        if len(parts) >= 2: return datetime(year, int(parts[0]), int(parts[1]))
    except: pass
    return None

def parse_period_dates(p_str):
    try:
        dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(p_str).replace('\n',''))
        if len(dates) >= 2:
            s = datetime.strptime(dates[0], "%Y.%m.%d")
            e = datetime.strptime(dates[1], "%Y.%m.%d")
            return s, e, (e - s).days + 1
    except: pass
    return None, None, 0

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()

# --- 醫院代表模式 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 將勾選條件放在醫院代表頁面頂部
    st.markdown("### 規則設定")
    c_cfg1, c_cfg2, c_cfg3 = st.columns([1, 1, 1])
    with c_cfg1: course_dur_weeks = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
    with c_cfg2: min_weeks_req = st.number_input("最短實習週數要求 (週)", min_value=1, value=4)
    with c_cfg3: require_cont = st.checkbox("要求必須連續實習", value=True)
    st.divider()

    course_days = course_dur_weeks * 5
    total_min_days = min_weeks_req * 5

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("1. 上傳醫院容額表", type=['xlsx'], key="h_q")
    with c2: a_file = st.file_uploader("2. 上傳學生志願表", type=['xlsx'], key="h_a")

    if q_file and a_file:
        try:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)

            if df_q is not None and df_a is not None:
                if '姓名' in df_a.columns: df_a['姓名'] = df_a['姓名'].ffill()
                
                # 建立解析後的學生志願
                apps = []
                dept_col = "申請科別" if "申請科別" in df_a.columns else "科別"
                for _, row in df_a.iterrows():
                    if pd.notna(row.get(dept_col)) and pd.notna(row.get('實習期間')):
                        s, e, d = parse_period_dates(row['實習期間'])
                        if s: apps.append({'姓名': row['姓名'], '科別': str(row[dept_col]).strip(), '開始': s, '結束': e, '天數': d})
                
                # 執行精準碰撞偵測 (解決莊晰之衝突問題)
                date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if dept == 'nan' or not dept: continue
                    
                    for col in date_cols:
                        cap = q_row.get(col)
                        try: cap_val = int(float(cap))
                        except: continue
                        
                        # 解析容額表該小格的日期區間
                        pts = col.split('-')
                        s_slot = parse_date_simple(pts[0])
                        e_slot = parse_date_simple(pts[1]) if len(pts) > 1 else s_slot
                        
                        if s_slot and e_slot:
                            # 找出在此特定小格時間內有佔位的學生
                            st_in_slot = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and s_slot <= a['結束']]
                            if len(st_in_slot) > cap_val:
                                collisions.append({"科別": dept, "時間": col, "容額": cap_val, "超額學生": "、".join(list(set(st_in_slot)))})

                # 規章判定
                invalid = []
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, group in df_temp.groupby('姓名'):
                        total_days = group['天數'].sum()
                        for _, row in group.iterrows():
                            if row['天數'] < course_days:
                                invalid.append({"姓名": name, "原因": f"{row['科別']} 實習天數不足({row['天數']}天)"})
                        if total_days < total_min_days:
                            invalid.append({"姓名": name, "原因": f"總天數不足({total_days}天)"})

                st.header("異常監控結果")
                if collisions:
                    st.subheader("名額撞期名單 (超額佔位)")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("核對完成，目前一切正常。")
        except Exception as e: st.error(f"解析失敗：{e}")

# --- 系秘模式 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    if multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['姓名'] = df['姓名'].ffill(); df['來源醫院'] = f.name
                all_data.append(df[df['實習期間'].notna()])
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            d1_s, d1_e, _ = parse_period_dates(s_apps[i]['實習期間'])
                            d2_s, d2_e, _ = parse_period_dates(s_apps[j]['實習期間'])
                            if d1_s and d2_s and (d1_s <= d2_e and d2_s <= d1_e):
                                conflicts.append({"姓名": name, "醫院 A": s_apps[i]['來源醫院'], "時間 A": s_apps[i]['實習期間'], "醫院 B": s_apps[j]['來源醫院'], "時間 B": s_apps[j]['實習期間']})
            if conflicts:
                st.subheader("偵測到跨院衝突名單")
                st.table(pd.DataFrame(conflicts).drop_duplicates())
            else: st.success("無重複佔位。")
