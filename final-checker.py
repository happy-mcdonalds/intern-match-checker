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

# --- 高級感行政極簡風 + 現代黑體 CSS ---
st.markdown("""
    <style>
    /* 引入思源黑體 (Noto Sans TC)，提供 400(一般), 500(中等), 700(粗體) */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
    
    /* 全局字體與背景設定 */
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'PingFang TC', 'Noto Sans TC', 'Microsoft JhengHei', 'Helvetica Neue', sans-serif !important;
        background-color: #F8FAFC !important; /* 極淺的灰藍底色，高階 SaaS 常用 */
        color: #1E293B !important; /* 深石板灰，比純黑更有質感 */
        -webkit-font-smoothing: antialiased; /* 讓字體在 Mac/iOS 上更平滑 */
    }
    
    /* 標題設計 */
    h1, h2, h3 { 
        color: #0F172A !important; /* 更深的午夜藍 */
        border-bottom: 2px solid #E2E8F0; 
        padding-bottom: 8px;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    
    /* 側邊欄設計 */
    section[data-testid="stSidebar"] { 
        background-color: #FFFFFF !important; /* 側邊欄純白，凸顯層次 */
        border-right: 1px solid #E2E8F0 !important; 
    }
    
    /* 表單區塊設計 */
    [data-testid="stForm"] {
        border: 1px solid #E2E8F0 !important;
        background-color: #FFFFFF !important;
        border-radius: 8px;
        padding: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02); /* 加入極淡的陰影增加立體感 */
    }
    
    /* 一般按鈕 (深藏青色) */
    .stButton > button, [data-testid="stFormSubmitButton"] > button { 
        background-color: #1E293B !important;
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 6px !important; 
        width: 100%; 
        font-weight: 500 !important;
        letter-spacing: 1px;
        transition: all 0.2s ease-in-out;
    }
    .stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #334155 !important;
        transform: translateY(-1px); /* 懸浮時微微上浮 */
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }

    /* 重新整理與儲存按鈕 (優雅灰) */
    .btn-secondary > button {
        background-color: #94A3B8 !important;
        color: #FFFFFF !important;
    }
    .btn-secondary > button:hover {
        background-color: #64748B !important;
    }
    
    /* 表格設計 */
    .stTable { font-size: 14.5px; }
    th {
        background-color: #F1F5F9 !important; /* 乾淨的表頭底色 */
        color: #334155 !important;
        border-bottom: 2px solid #CBD5E1 !important;
        font-weight: 700 !important;
    }
    td { border-bottom: 1px solid #F1F5F9 !important; }
    th, td {
        white-space: pre-wrap !important; 
        vertical-align: middle !important;
        line-height: 1.6 !important;
        text-align: left !important;
        padding: 10px 12px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---
def smart_read_sheet(file):
    # 這個超級引擎現在同時為「醫院代表」和「系秘」服務
    try:
        # 1. 支援 CSV 與 Excel
        if file.name.endswith('.csv'):
            df_temp = pd.read_csv(file, header=None)
            target_sheet = None
        else:
            xls = pd.ExcelFile(file)
            target_sheet = xls.sheet_names[0]
            for sn in xls.sheet_names:
                if any(k in sn for k in ["志願", "名單", "工作表4", "實習容額", "申請"]):
                    target_sheet = sn
                    break
            df_temp = pd.read_excel(file, sheet_name=target_sheet, header=None)
        
        # 2. 修正標題列判斷邏輯
        header_idx = 0
        for i in range(min(len(df_temp), 15)):
            row_str = "".join([str(x).replace(" ", "") for x in df_temp.iloc[i].values])
            if "姓名" in row_str or "科別" in row_str or "日期" in row_str:
                header_idx = i
                break
        
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=header_idx)
        else:
            df = pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
        
        # 3. 清理與統一欄位名稱 (包含對「科別」的相容)
        cols = []
        for c in df.columns:
            c_str = str(c).strip().replace('\n', '').replace(' ', '')
            if "姓名" in c_str:
                cols.append("姓名")
            elif "期間" in c_str:
                cols.append("實習期間")
            elif "科別" in c_str or "科" in c_str:
                cols.append("科別")
            else:
                cols.append(c_str)
        df.columns = cols
        
        # 4. 針對你的新格式：自動將「開始日期」與「結束日期」合併為一格「實習期間」
        start_col = next((c for c in df.columns if "開始" in str(c) and "日期" in str(c)), None)
        end_col = next((c for c in df.columns if "結束" in str(c) and "日期" in str(c)), None)
        
        if start_col and end_col and "實習期間" not in df.columns:
            def safe_format(val):
                if pd.isna(val) or str(val).strip() == '' or str(val).lower() == 'nan':
                    return ""
                try:
                    return pd.to_datetime(val).strftime('%Y/%m/%d')
                except:
                    # 確保將破折號轉為底層系統認得的斜線
                    return str(val).replace('-', '/')
                    
            df["實習期間"] = df.apply(
                lambda row: f"{safe_format(row[start_col])} ~ {safe_format(row[end_col])}" 
                if safe_format(row[start_col]) and safe_format(row[end_col]) else None, 
                axis=1
            )
            
        return df
    except Exception as e: 
        return None

def extract_dates_universal(text, year=2026):
    """終極日期解析引擎"""
    if isinstance(text, datetime): return text, text
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
    
    st.markdown("### 檔案上傳與比對")
    c1, c2 = st.columns(2)
    # 開放讓醫院代表也可以上傳 CSV
    with c1: q_file = st.file_uploader("上傳醫院容額表", type=['xlsx', 'csv'])
    with c2: a_file = st.file_uploader("上傳學生志願表", type=['xlsx', 'csv'])
    run_check = st.button("確認並開始比對")

    if run_check and q_file and a_file:
        course_workdays = st.session_state.course_dur_weeks * 5
        total_min_workdays = st.session_state.min_weeks_req * 5
        
        try:
            # 容額表讀取 (相容 CSV 或 Excel)
            if q_file.name.endswith('.csv'):
                df_q = pd.read_csv(q_file, header=4)
            else:
                xls_q = pd.ExcelFile(q_file)
                try: sn_q = [s for s in xls_q.sheet_names if "容額" in s or "時段" in s][0]
                except: sn_q = xls_q.sheet_names[0]
                df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]

            # 【關鍵修復】: 將強大的 smart_read_sheet 套用到學生志願表上
            df_a = smart_read_sheet(a_file)

            if df_q is not None and df_a is not None:
                if '姓名' in df_a.columns: df_a['姓名'] = df_a['姓名'].ffill()
                
                apps = []
                # smart_read_sheet 已經統一將欄位名稱洗成 "科別"
                dept_col = "科別" if "科別" in df_a.columns else "申請科別"
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
                    
                st.header("異常監控結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
          
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("名額分配與規章核對完全符合規定。")
        except Exception as e: 
            st.error(f"解析失敗：{e}")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    
    st.markdown("### 檔案上傳")
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx', 'csv'], accept_multiple_files=True)
    run_check_sec = st.button("確認並開始比對")
        
    if run_check_sec and multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns and '實習期間' in df.columns:
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
