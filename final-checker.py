import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 高級感 CSS (宋體 + 黑白灰 + 支援條列式換行) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', serif !important;
        color: #000000;
    }
    h1, h2, h3 { 
        color: #000000 !important; 
        border-bottom: 1px solid #000000; 
        padding-bottom: 5px; 
    }
    .stApp { background-color: #FFFFFF; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #DDDDDD; }
    .stButton>button { color: #FFFFFF !important; background-color: #000000 !important; border-radius: 0px; width: 100%; }
    
    /* 表格設定：保留換行符號 (\n)，並讓文字向上對齊，適合條列式閱讀 */
    .stTable { font-size: 14px; }
    th, td {
        white-space: pre-wrap !important; 
        word-break: keep-all !important;
        vertical-align: top !important; 
        line-height: 1.8 !important;
    }
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
    """解析日期並計算精準的工作天數 (排除六日)"""
    try:
        dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(p_str).replace('\n',''))
        if len(dates) >= 2:
            s = datetime.strptime(dates[0], "%Y.%m.%d")
            e = datetime.strptime(dates[1], "%Y.%m.%d")
            # 利用 pandas 計算實際工作天 (Business days)
            workdays = len(pd.bdate_range(s, e))
            return s, e, workdays
    except: pass
    return None, None, 0

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()

# --- 醫院代表模式 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    st.markdown("### 📋 規則設定")
    c_cfg1, c_cfg2, c_cfg3 = st.columns([1, 1, 1])
    with c_cfg1: course_dur_weeks = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
    with c_cfg2: min_weeks_req = st.number_input("最短實習週數要求 (週)", min_value=1, value=4)
    with c_cfg3: require_cont = st.checkbox("要求必須連續實習", value=True)
    st.divider()

    # 換算為嚴格工作天數 (1週 = 5個工作天)
    course_workdays = course_dur_weeks * 5
    total_min_workdays = min_weeks_req * 5

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("1. 上傳醫院容額表", type=['xlsx'], key="h_q")
    with c2: a_file = st.file_uploader("2. 上傳學生志願表", type=['xlsx'], key="h_a")

    if q_file and a_file:
        try:
            # 1. 嚴格鎖定容額表 Header=4
            xls_q = pd.ExcelFile(q_file)
            try: sn_q = [s for s in xls_q.sheet_names if "容額" in s or "時段" in s][0]
            except: sn_q = xls_q.sheet_names[0]
            df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]

            # 2. 讀取志願表
            xls_a = pd.ExcelFile(a_file)
            try: sn_a = [s for s in xls_a.sheet_names if "志願" in s][0]
            except: sn_a = xls_a.sheet_names[0]
            df_temp = pd.read_excel(a_file, sheet_name=sn_a)
            header_idx = 0
            for i in range(min(len(df_temp), 15)):
                row = [str(x).strip() for x in df_temp.iloc[i].values]
                if any(k in row for k in ["姓名", "科別", "申請科別"]):
                    header_idx = i + 1
                    break
            df_a = pd.read_excel(a_file, sheet_name=sn_a, header=header_idx)
            df_a.columns = [str(c).strip() for c in df_a.columns]

            if df_q is not None and df_a is not None:
                if '姓名' in df_a.columns: df_a['姓名'] = df_a['姓名'].ffill()
                
                apps = []
                dept_col = "申請科別" if "申請科別" in df_a.columns else "科別"
                for _, row in df_a.iterrows():
                    if pd.notna(row.get(dept_col)) and pd.notna(row.get('實習期間')):
                        s, e, d = parse_period_dates(row['實習期間'])
                        if s: apps.append({'姓名': row['姓名'], '科別': str(row[dept_col]).strip(), '開始': s, '結束': e, '天數': d})
                
                # --- 碰撞偵測 ---
                date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if dept == 'nan' or not dept: continue
                    
                    for col in date_cols:
                        cap = q_row.get(col)
                        try: cap_val = int(float(re.sub(r'[^0-9.]', '', str(cap))))
                        except: continue
                        
                        pts = col.split('-')
                        s_slot = parse_date_simple(pts[0].strip())
                        e_slot = parse_date_simple(pts[1].strip()) if len(pts) > 1 else s_slot
                        
                        if s_slot and e_slot:
                            st_in_slot = []
                            for a in apps:
                                if a['科別'] == dept:
                                    if a['開始'] <= e_slot and a['結束'] >= s_slot:
                                        st_in_slot.append(a['姓名'])
                            
                            if len(st_in_slot) > cap_val:
                                collisions.append({
                                    "科別": dept, 
                                    "時間": col.replace('\n', ''), 
                                    "容額": cap_val, 
                                    "超額學生": "、".join(list(set(st_in_slot)))
                                })

                # --- 規章嚴格審核 ---
                invalid = []
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, group in df_temp.groupby('姓名'):
                        # 依照開始時間排序，確保連續性檢查正確
                        group = group.sort_values('開始')
                        total_workdays = group['天數'].sum()
                        
                        # 1. 檢查單一 Course 天數
                        for _, row in group.iterrows():
                            if row['天數'] < course_workdays:
                                invalid.append({"姓名": name, "原因": f"【Course 天數不足】 {row['科別']} 僅 {row['天數']} 個工作天 (規定需 {course_workdays} 天)"})
                        
                        # 2. 檢查總實習天數
                        if total_workdays < total_min_workdays:
                            invalid.append({"姓名": name, "原因": f"【總時長不足】 僅 {total_workdays} 個工作天 (規定需 {total_min_workdays} 天)"})
                        
                        # 3. 檢查是否連續實習
                        if require_cont and len(group) > 1:
                            courses = group.to_dict('records')
                            for i in range(len(courses) - 1):
                                prev_end = courses[i]['結束']
                                next_start = courses[i+1]['開始']
                                # 若下個開始時間晚於前一個結束時間超過 3 天 (代表中間非合理週末換科)
                                if (next_start - prev_end).days > 3:
                                    invalid.append({"姓名": name, "原因": f"【未連續實習】 {courses[i]['科別']} 與 {courses[i+1]['科別']} 之間出現中斷"})
                                    break # 報一次錯誤即可

                # --- 顯示結果 ---
                st.header("異常監控結果")
                if collisions:
                    st.subheader("⚠️ 名額撞期名單 (超額佔位)")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("📝 規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("名額分配與規章核對完全符合規定。")
        except Exception as e: st.error(f"解析失敗：{e}")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    st.markdown("請上傳各院檔案，系統將自動比對。")
    
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    if multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['姓名'] = df['姓名'].ffill()
                clean_hosp_name = f.name.replace('.xlsx', '').replace('.csv', '')
                df['來源醫院'] = clean_hosp_name
                all_data.append(df[df['實習期間'].notna()])
                
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    conflict_set = set()
                    
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            d1_s, d1_e, _ = parse_period_dates(s_apps[i]['實習期間'])
                            d2_s, d2_e, _ = parse_period_dates(s_apps[j]['實習期間'])
                            if d1_s and d2_s and (d1_s <= d2_e and d2_s <= d1_e):
                                conflict_set.add(i)
                                conflict_set.add(j)
                    
                    if conflict_set:
                        details = []
                        for idx in sorted(list(conflict_set)):
                            hosp = s_apps[idx]['來源醫院']
                            period = str(s_apps[idx]['實習期間']).replace('\n', '')
                            details.append(f"- {hosp} ({period})")
                        
                        conflicts.append({
                            "姓名": name,
                            "衝突詳情": "\n".join(details)
                        })
                        
            if conflicts:
                st.subheader("⚠️ 偵測到跨院衝突名單")
                df_conflicts = pd.DataFrame(conflicts)
                st.table(df_conflicts.set_index('姓名'))
            else: 
                st.success("無重複佔位情況。")
