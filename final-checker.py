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

# --- 工具函式 ---
def parse_date_simple(s, year=2026):
    try:
        parts = re.findall(r'\d+', str(s))
        if len(parts) >= 2: return datetime(year, int(parts[0]), int(parts[1]))
    except: pass
    return None

def count_workdays(start, end):
    """計算實習天數（含頭尾）"""
    if not start or not end: return 0
    return (end - start).days + 1

def parse_dates(period_str):
    """供系秘模式使用的區間解析"""
    try:
        dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(period_str).replace('\n',''))
        if len(dates) >= 2:
            s = datetime.strptime(dates[0], "%Y.%m.%d")
            e = datetime.strptime(dates[1], "%Y.%m.%d")
            return s, e, (e - s).days + 1
    except: pass
    return None, None, 0

def smart_read_sheet(file):
    """專供系秘模式使用的自動定位讀取 (容錯率高)"""
    try:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
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
        df = pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except: return None

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()

# --- 模式：醫院代表 (採用你驗證過的完美邏輯) ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 規則設定 (僅在醫院代表主畫面顯示)
    st.markdown("### 規則設定")
    c_cfg1, c_cfg2, c_cfg3 = st.columns(3)
    with c_cfg1: course_duration_weeks = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
    with c_cfg2: min_weeks_req = st.number_input("最短實習週數要求 (週)", min_value=1, value=4)
    with c_cfg3: require_cont = st.checkbox("要求必須連續實習", value=True)
    st.divider()

    course_days = course_duration_weeks * 5
    total_min_days = min_weeks_req * 5

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("1. 上傳醫院容額表", type=['xlsx'], key="h_q")
    with c2: a_file = st.file_uploader("2. 上傳學生志願表", type=['xlsx'], key="h_a")

    if q_file and a_file:
        try:
            # 1. 精準讀取容額表 (鎖定 header=4)
            xls_q = pd.ExcelFile(q_file)
            try: sn_q = [s for s in xls_q.sheet_names if "容額" in s or "時段" in s][0]
            except: sn_q = xls_q.sheet_names[0]
            df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]

            # 2. 精準讀取申請表
            xls_a = pd.ExcelFile(a_file)
            try: sn_a = [s for s in xls_a.sheet_names if "志願" in s or "名單" in s][0]
            except: sn_a = xls_a.sheet_names[0]
            df_a = pd.read_excel(a_file, sheet_name=sn_a)
            df_a.columns = [str(c).strip() for c in df_a.columns]
            
            if '姓名' in df_a.columns: df_a['姓名'] = df_a['姓名'].ffill()

            # 解析志願
            apps = []
            dept_col = "申請科別" if "申請科別" in df_a.columns else "科別"
            for _, row in df_a.iterrows():
                if pd.notna(row.get(dept_col)) and pd.notna(row.get('實習期間')):
                    dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(row['實習期間']))
                    if len(dates) >= 2:
                        s = datetime.strptime(dates[0], "%Y.%m.%d")
                        e = datetime.strptime(dates[1], "%Y.%m.%d")
                        apps.append({'姓名': row['姓名'], '科別': str(row[dept_col]).strip(), '開始': s, '結束': e, '天數': count_workdays(s, e)})

            # 3. 執行碰撞偵測 (你提供的強大區間比對)
            date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
            collisions = []
            for _, q_row in df_q.iterrows():
                dept = q_row.get('科別')
                if pd.isna(dept): continue

                for col in date_cols:
                    cap = q_row.get(col)
                    if pd.isna(cap) or not str(cap).isdigit(): continue
                    cap_val = int(float(cap))

                    pts = col.split('-')
                    s_slot = parse_date_simple(pts[0])
                    e_slot = parse_date_simple(pts[1]) if len(pts) > 1 else s_slot

                    # 核心：精準找出重疊日期的人
                    st_in_slot = [a['姓名'] for a in apps if a['科別'] == str(dept).strip() and a['開始'] <= e_slot and s_slot <= a['結束']]

                    if len(st_in_slot) > cap_val:
                        collisions.append({"科別": dept, "時間": col.replace('\n', ''), "容額": cap_val, "超額學生": "、".join(st_in_slot)})

            # 4. 資格審核
            invalid_students = []
            df_temp = pd.DataFrame(apps)
            if not df_temp.empty:
                for name, group in df_temp.groupby('姓名'):
                    total_days = group['天數'].sum()
                    for _, row in group.iterrows():
                        if row['天數'] < course_days:
                            invalid_students.append({"姓名": name, "原因": f"{row['科別']} 實習天數不足({row['天數']}天)，未達 Course 要求"})
                    if total_days < total_min_days:
                        invalid_students.append({"姓名": name, "原因": f"總實習天數不足({total_days}天)，未達最低要求"})

            # 顯示結果
            st.header("異常監控結果")
            if collisions:
                st.subheader("名額撞期名單 (超額佔位)")
                st.table(pd.DataFrame(collisions))
            else:
                st.success("名額分配正常。")

            if invalid_students:
                st.subheader("規章不符名單 (天數不足)")
                st.table(pd.DataFrame(invalid_students).drop_duplicates())

        except Exception as e:
            st.error(f"解析失敗：{e}")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    st.markdown("此模式專供系秘比對不同醫院間是否有學生重複佔位。")
    
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    
    if multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['姓名'] = df['姓名'].ffill()
                df['來源醫院'] = f.name
                all_data.append(df[df['申請科別'].notna()] if '申請科別' in df.columns else df)
        
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            d1_s, d1_e, _ = parse_dates(s_apps[i].get('實習期間'))
                            d2_s, d2_e, _ = parse_dates(s_apps[j].get('實習期間'))
                            if d1_s and d2_s and (d1_s <= d2_e and d2_s <= d1_e):
                                conflicts.append({
                                    "姓名": name,
                                    "醫院 A": s_apps[i]['來源醫院'], "時間 A": s_apps[i].get('實習期間'),
                                    "醫院 B": s_apps[j]['來源醫院'], "時間 B": s_apps[j].get('實習期間')
                                })
            
            if conflicts:
                st.subheader("偵測到跨院衝突名單")
                st.table(pd.DataFrame(conflicts).drop_duplicates())
            else:
                st.success("交叉比對完成，無重複佔位情況。")
