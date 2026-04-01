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

# --- 莫蘭迪色系 + 強制純宋體 CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'serif' !important;
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

    .btn-secondary > button {
        background-color: #C0BFB8 !important;
        color: #FFFFFF !important;
    }
    
    /* 表格底層強化 */
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th { background-color: #E3E1DB !important; color: #4A4C4B !important; padding: 12px; text-align: left; border-bottom: 2px solid #C0BFB8; }
    td { padding: 12px; border-bottom: 1px solid #EAE8E3; vertical-align: top; line-height: 1.6; white-space: pre-wrap !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---

def smart_read_sheet(file):
    """強力偵測標題：只要欄位名稱包含關鍵字就強制歸位"""
    try:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if any(k in sn for k in ["志願", "名單", "工作表4", "實習容額"]):
                target_sheet = sn
                break
        
        # 讀取時不設 header，手動掃描
        df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)
        header_idx = 0
        for i in range(min(len(df_raw), 25)):
            row_vals = [str(x).strip() for x in df_raw.iloc[i].values]
            if any("姓名" in x or "科別" in x for x in row_vals):
                header_idx = i
                break
        
        df = pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
        
        # 欄位正名手術：解決空白、換行、各種名稱混亂
        clean_cols = {}
        for c in df.columns:
            orig_c = str(c).strip()
            # 移除所有空白後檢查
            c_simple = re.sub(r'\s+', '', orig_c)
            if "姓名" in c_simple: clean_cols[c] = "姓名"
            elif "科別" in c_simple: clean_cols[c] = "科別"
            elif "實習期間" in c_simple: clean_cols[c] = "實習期間"
            else: clean_cols[c] = c_simple
            
        df = df.rename(columns=clean_cols)
        # 確保所有欄位名稱都是乾淨的字串
        df.columns = [str(c) for c in df.columns]
        return df
    except:
        return None

def extract_dates_universal(text, year=2026):
    if isinstance(text, datetime): return text, text
    # 將所有可能的換行、多餘空格換成標準橫線
    text = re.sub(r'[\n\r\s]+', '-', str(text)).strip()
    parts = re.split(r'[-~～到至_]+', text)
    
    def parse_part(part):
        nums = re.findall(r'\d+', part)
        if len(nums) >= 2:
            y = int(nums[0]) if len(nums[0]) == 4 else year
            m, d = (int(nums[1]), int(nums[2])) if len(nums[0]) == 4 else (int(nums[-2]), int(nums[-1]))
            return datetime(y, m, d)
        return None
    
    dates = [parse_part(p) for p in parts if parse_part(p)]
    return (dates[0], dates[-1]) if len(dates) >= 2 else (dates[0], dates[0]) if dates else (None, None)

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
    with c1: q_file = st.file_uploader("上傳醫院容額表", type=['xlsx'], key="q_up")
    with c2: a_file = st.file_uploader("上傳學生志願表", type=['xlsx'], key="a_up")
    run_check = st.button("確認並開始比對")

    if run_check and q_file and a_file:
        course_workdays = st.session_state.course_dur_weeks * 5
        total_min_workdays = st.session_state.min_weeks_req * 5
        
        try:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)

            if df_q is not None and df_a is not None:
                # 最終檢查點：如果還是沒抓到姓名，直接報錯
                if '姓名' not in df_a.columns:
                    st.error("志願表中找不到「姓名」欄位。請檢查 Excel 表頭是否包含「姓名」二字。")
                    st.stop()
                
                df_a['姓名'] = df_a['姓名'].ffill()
                apps = []
                for _, row in df_a.iterrows():
                    # 改用正名後的欄位
                    d_val = row.get('科別')
                    t_val = row.get('實習期間')
                    if pd.notna(d_val) and pd.notna(t_val):
                        s, e, d = parse_period_dates(t_val)
                        if s: apps.append({'姓名': row['姓名'], '科別': str(d_val).strip(), '開始': s, '結束': e, '天數': d})
                
                # 容額比對
                date_cols = [c for c in df_q.columns if extract_dates_universal(c)[0]]
                q_dept_col = '科別' if '科別' in df_q.columns else df_q.columns[0]
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get(q_dept_col, '')).strip()
                    if dept == 'nan' or not dept: continue
                    for col in date_cols:
                        try:
                            cap_val = int(float(re.sub(r'[^0-9.]', '', str(q_row.get(col)))))
                        except: continue
                        s_slot, e_slot = extract_dates_universal(col)
                        st_in_slot = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        if len(st_in_slot) > cap_val:
                            collisions.append({"科別": dept, "時間": str(col).replace('\n', ' '), "容額": cap_val, "超額學生": "、".join(list(set(st_in_slot)))})

                # 規章檢查
                invalid = []
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, group in df_temp.groupby('姓名'):
                        group = group.sort_values('開始')
                        total_workdays = group['天數'].sum()
                        for _, row in group.iterrows():
                            if row['天數'] < course_workdays:
                                invalid.append({"姓名": name, "原因": f"Course 天數不足：{row['科別']} ({row['天數']} 天)"})
                        if total_workdays < total_min_workdays:
                            invalid.append({"姓名": name, "原因": f"總時長不足 ({total_workdays} 天)"})
                        if st.session_state.require_cont and len(group) > 1:
                            courses = group.to_dict('records')
                            for i in range(len(courses) - 1):
                                if (courses[i+1]['開始'] - courses[i]['結束']).days > 3:
                                    invalid.append({"姓名": name, "原因": "實習中斷(未連續)"})
                                    break 

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("核對完成，查無異常。")
        except Exception as e: st.error(f"分析過程錯誤：{e}")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    multi_files = st.file_uploader("上傳各院清單 (多選)", type=['xlsx'], accept_multiple_files=True)
    run_check_sec = st.button("確認並開始比對")
        
    if run_check_sec and multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['姓名'] = df['姓名'].ffill()
                df['來源醫院'] = f.name.replace('.xlsx', '')
                time_col = '實習期間' if '實習期間' in df.columns else None
                if time_col:
                    all_data.append(df[df[time_col].notna()])
                
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    hit = set()
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            s1, e1, _ = parse_period_dates(s_apps[i]['實習期間'])
                            s2, e2, _ = parse_period_dates(s_apps[j]['實習期間'])
                            if s1 and s2 and (s1 <= e2 and s2 <= e1):
                                hit.update([i, j])
                    if hit:
                        # 換行符號 \n 是條列式的核心
                        details = "<br>".join([f"• {s_apps[idx]['來源醫院']} ({str(s_apps[idx]['實習期間']).replace('\n','')})" for idx in sorted(list(hit))])
                        conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("偵測到跨院衝突名單")
                # 使用 HTML 模式顯示表格，徹底解決換行失效
                html_table = "<table style='width:100%'><tr><th>姓名</th><th>衝突詳情</th></tr>"
                for c in conflicts:
                    html_table += f"<tr><td>{c['姓名']}</td><td>{c['衝突詳情']}</td></tr>"
                html_table += "</table>"
                st.markdown(html_table, unsafe_allow_html=True)
            else: 
                st.success("無重複佔位情況。")
