import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# --- 初始化系統記憶 ---
if "course_dur_weeks" not in st.session_state:
    st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state:
    st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state:
    st.session_state.require_cont = True

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 莫蘭迪色系 + 強制純宋體 CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'PMingLiU', 'MingLiU', 'SimSun', serif !important;
        background-color: #F5F4F1 !important; 
        color: #5C5E5D !important; 
    }
    
    h1, h2, h3 { 
        color: #4A4C4B !important; 
        border-bottom: 1px solid #D6D4CE; 
        padding-bottom: 5px; 
        font-weight: 700;
    }
    
    section[data-testid="stSidebar"] { 
        background-color: #EAE8E3 !important; 
        border-right: 1px solid #D6D4CE !important; 
    }
    
    [data-testid="stForm"] {
        border: 1px solid #D6D4CE !important;
        background-color: #FDFDFD !important;
        border-radius: 4px;
        padding: 20px;
    }
    
    /* 一般按鈕 (鼠尾草綠) */
    .stButton > button, [data-testid="stFormSubmitButton"] > button { 
        background-color: #8A9A92 !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 4px !important; 
        width: 100%; 
        transition: 0.3s;
    }
    .stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #72827A !important;
    }

    /* 重新整理與儲存按鈕 (淡雅灰) */
    .btn-secondary > button {
        background-color: #C0BFB8 !important;
        color: #FFFFFF !important;
    }
    .btn-secondary > button:hover {
        background-color: #A8A7A0 !important;
    }
    
    .stTable { font-size: 14px; }
    th {
        background-color: #E3E1DB !important;
        color: #4A4C4B !important;
        border-bottom: 2px solid #C0BFB8 !important;
    }
    td { border-bottom: 1px solid #EAE8E3 !important; }
    th, td {
        white-space: pre-wrap !important; 
        vertical-align: top !important; 
        line-height: 1.8 !important;
        text-align: left !important;
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

def extract_dates_universal(text, year=2026):
    """終極日期解析引擎：強殺換行符號與缺零日期"""
    if isinstance(text, datetime): return text, text
    # 關鍵：強制把 Excel 內的換行符號 \n 轉換為破折號
    text = str(text).replace('\n', '-').replace('\r', '-').replace(' ', '').strip()
    parts = re.split(r'[-~～到至_]+', text)
    
    def extract_single_date(part):
        nums = re.findall(r'\d+', part)
        if len(nums) >= 2:
            if len(nums[0]) == 4 and len(nums) >= 3:
                return datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            return datetime(year, int(nums[-2]), int(nums[-1]))
        return None
        
    dates = [extract_single_date(p) for p in parts if extract_single_date(p) is not None]
    if len(dates) == 1: return dates[0], dates[0]
    elif len(dates) >= 2: return dates[0], dates[-1]
    return None, None

def parse_period_dates(p_str):
    try:
        s, e = extract_dates_universal(p_str)
        if s and e:
            workdays = len(pd.bdate_range(s, e))
            return s, e, workdays
    except: pass
    return None, None, 0

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()
st.sidebar.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
if st.sidebar.button("重新整理系統"): st.rerun()
st.sidebar.markdown('</div>', unsafe_allow_html=True)

# --- 醫院代表模式 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 區塊一：設定與儲存
    with st.form("settings_form"):
        st.markdown("### 規則設定")
        c_cfg1, c_cfg2, c_cfg3 = st.columns([1, 1, 1])
        with c_cfg1: cd_val = st.number_input("一個 Course 多久 (週)", min_value=1, value=st.session_state.course_dur_weeks)
        with c_cfg2: mw_val = st.number_input("最短實習週數要求 (週)", min_value=1, value=st.session_state.min_weeks_req)
        with c_cfg3: rc_val = st.checkbox("要求必須連續實習", value=st.session_state.require_cont)
        
        st.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
        save_btn = st.form_submit_button("儲存條件")
        st.markdown('</div>', unsafe_allow_html=True)
        
        if save_btn:
            st.session_state.course_dur_weeks = cd_val
            st.session_state.min_weeks_req = mw_val
            st.session_state.require_cont = rc_val
            st.success("條件已儲存！請接續上傳檔案。")

    st.divider()
    
    # 區塊二：檔案上傳與比對
    st.markdown("### 檔案上傳與比對")
    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("上傳醫院容額表", type=['xlsx'])
    with c2: a_file = st.file_uploader("上傳學生志願表", type=['xlsx'])
    run_check = st.button("確認並開始比對")

    if run_check and q_file and a_file:
        course_workdays = st.session_state.course_dur_weeks * 5
        total_min_workdays = st.session_state.min_weeks_req * 5
        
        try:
            xls_q = pd.ExcelFile(q_file)
            try: sn_q = [s for s in xls_q.sheet_names if "容額" in s or "時段" in s][0]
            except: sn_q = xls_q.sheet_names[0]
            df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]
            # ⭐ 修正第一個時段欄位被吃掉（5/4-5/8 → Unnamed: 2）
            if "Unnamed: 2" in df_q.columns:
            df_q = df_q.rename(columns={"Unnamed: 2": "5/4-5/8"})

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
                
                date_cols = []
                slot_mapping = {}
                for c in df_q.columns:
                    s_slot, e_slot = extract_dates_universal(c)
                    if s_slot and e_slot:
                        date_cols.append(c)
                        slot_mapping[c] = (s_slot, e_slot)

                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if dept == 'nan' or not dept: continue
                    
                    for col in date_cols:
                        cap = q_row.get(col)
                        try: cap_val = int(float(re.sub(r'[^0-9.]', '', str(cap))))
                        except: continue
                        
                        s_slot, e_slot = slot_mapping[col]
                        st_in_slot = []
                        for a in apps:
                            if a['科別'] == dept:
                                if a['開始'] <= e_slot and a['結束'] >= s_slot:
                                    st_in_slot.append(a['姓名'])
                        
                        if len(st_in_slot) > cap_val:
                            collisions.append({
                                "科別": dept, 
                                "時間": str(col).replace('\n', ''), 
                                "容額": cap_val, 
                                "超額學生": "、".join(list(set(st_in_slot)))
                            })

                invalid = []
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, group in df_temp.groupby('姓名'):
                        group = group.sort_values('開始')
                        total_workdays = group['天數'].sum()
                        
                        for _, row in group.iterrows():
                            if row['天數'] < course_workdays:
                                invalid.append({"姓名": name, "原因": f"Course 天數不足：{row['科別']} 僅 {row['天數']} 個工作天 (需 {course_workdays} 天)"})
                        
                        if total_workdays < total_min_workdays:
                            invalid.append({"姓名": name, "原因": f"總時長不足：僅 {total_workdays} 個工作天 (需 {total_min_workdays} 天)"})
                        
                        if st.session_state.require_cont and len(group) > 1:
                            courses = group.to_dict('records')
                            for i in range(len(courses) - 1):
                                prev_end = courses[i]['結束']
                                next_start = courses[i+1]['開始']
                                if (next_start - prev_end).days > 3:
                                    invalid.append({"姓名": name, "原因": f"未連續實習：{courses[i]['科別']} 與 {courses[i+1]['科別']} 中斷"})
                                    break 

                st.header("異常監控結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("名額分配與規章核對完全符合規定。")
        except Exception as e: st.error(f"解析失敗：{e}")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    
    st.markdown("### 檔案上傳")
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    run_check_sec = st.button("確認並開始比對")
        
    if run_check_sec and multi_files:
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
                st.subheader("偵測到跨院衝突名單")
                df_conflicts = pd.DataFrame(conflicts)
                st.table(df_conflicts.set_index('姓名'))
            else: 
                st.success("無重複佔位情況。")
