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

# --- 莫蘭迪色系 + 強制全域黑體 (修正上傳字體) ---
st.markdown("""
    <style>
    /* 全域黑體設定 */
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: "Microsoft JhengHei", "Heiti TC", "Apple LiGothic Medium", sans-serif !important;
        background-color: #F5F4F1 !important; 
        color: #5C5E5D !important; 
    }
    
    /* 修正上傳器字體 */
    [data-testid="stFileUploaderLabel"], 
    [data-testid="stFileUploadDropzone"] div, 
    [data-testid="stUploadedFile"] div {
        font-family: "Microsoft JhengHei", "Heiti TC", sans-serif !important;
    }
    
    h1, h2, h3 { 
        font-family: "Microsoft JhengHei", "Heiti TC", sans-serif !important;
        color: #4A4C4B !important; 
        border-bottom: 1px solid #D6D4CE; 
        padding-bottom: 5px; 
        font-weight: bold;
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
    
    /* 按鈕樣式 (鼠尾草綠) */
    .stButton > button, [data-testid="stFormSubmitButton"] > button { 
        background-color: #8A9A92 !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 4px !important; 
        width: 100%; 
        transition: 0.3s;
        font-family: "Microsoft JhengHei", "Heiti TC", sans-serif !important;
    }
    .stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #72827A !important;
    }

    .btn-secondary > button {
        background-color: #C0BFB8 !important;
        color: #FFFFFF !important;
    }
    
    /* 表格樣式 */
    table { width: 100%; border-collapse: collapse; margin-top: 10px; background-color: white; font-family: "Microsoft JhengHei", sans-serif; }
    th { background-color: #E3E1DB !important; color: #4A4C4B !important; padding: 12px; text-align: left; border-bottom: 2px solid #C0BFB8; }
    td { padding: 12px; border-bottom: 1px solid #EAE8E3; vertical-align: top; line-height: 1.6; white-space: pre-wrap !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---

def smart_read_sheet(file):
    """改良讀取：排除範例分頁，並自動對接兩種日期格式"""
    try:
        xls = pd.ExcelFile(file)
        # 篩選掉名稱包含「範例」、「範本」、「先寫這個」、「說明」、「空白」的分頁
        valid_sheets = [sn for sn in xls.sheet_names if not any(k in sn for k in ["範例", "範本", "先寫這個", "說明", "空白"])]
        
        # 從剩下的分頁中找包含「志願」、「名單」、「容額」的分頁
        target_sheet = None
        for sn in valid_sheets:
            if any(k in sn for k in ["志願", "名單", "容額", "工作"]):
                target_sheet = sn
                break
        
        if not target_sheet:
            target_sheet = valid_sheets[0] if valid_sheets else xls.sheet_names[0]
            
        # 掃描標題列
        df_scan = pd.read_excel(file, sheet_name=target_sheet, header=None, nrows=15)
        h_idx = 0
        for i, row in df_scan.iterrows():
            row_str = "".join([str(x) for x in row.values])
            if any(k in row_str for k in ["姓名", "科別", "日期"]):
                h_idx = i
                break
                
        df = pd.read_excel(file, sheet_name=target_sheet, header=h_idx)
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
        
        # 統一欄位名稱
        rename_map = {}
        for c in df.columns:
            if "姓名" in c: rename_map[c] = "姓名"
            elif "申請科別" in c or ("科別" in c and "備選" not in c): rename_map[c] = "科別"
            elif "實習期間" in c or "日期" in c: rename_map[c] = "日期欄位"
        df = df.rename(columns=rename_map)
        
        # 處理「日期分開兩格」的情況 (確定實習名單格式)
        if "日期欄位" not in df.columns:
            start_col = next((c for c in df.columns if "開始" in c), None)
            end_col = next((c for c in df.columns if "結束" in c), None)
            if start_col and end_col:
                # 合併成統一的「日期欄位」格式
                df["日期欄位"] = df[start_col].astype(str) + " - " + df[end_col].astype(str)
                
        return df
    except:
        return None

def extract_dates_universal(text, year=2026):
    if isinstance(text, datetime): return text, text
    s = re.sub(r'[\n\r\s]+', '-', str(text)).strip()
    s = re.sub(r'-+', '-', s)
    parts = re.split(r'[-~～到至_]+', s)
    
    def parse_part(part):
        nums = re.findall(r'\d+', part)
        if len(nums) >= 2:
            if len(nums[0]) == 4 and len(nums) >= 3:
                return datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            return datetime(year, int(nums[-2]), int(nums[-1]))
        return None
    
    dates = [parse_part(p) for p in parts if parse_part(p)]
    if len(dates) >= 2: return dates[0], dates[-1]
    elif len(dates) == 1: return dates[0], dates[0]
    return None, None

def parse_period_dates(p_str):
    s, e = extract_dates_universal(p_str)
    if s and e:
        workdays = len(pd.bdate_range(s, e))
        return s, e, workdays
    return None, None, 0

# --- UI 介面 ---

st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()
st.sidebar.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
if st.sidebar.button("重新整理系統"): st.rerun()
st.sidebar.markdown('</div>', unsafe_allow_html=True)

if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    with st.form("settings_form"):
        st.markdown("### 規則設定")
        c1, c2, c3 = st.columns(3)
        with c1: cd = st.number_input("一個 Course 多久 (週)", 1, 10, st.session_state.course_dur_weeks)
        with c2: mw = st.number_input("最短實習週數要求 (週)", 1, 52, st.session_state.min_weeks_req)
        with c3: rc = st.checkbox("要求必須連續實習", st.session_state.require_cont)
        if st.form_submit_button("儲存規則"):
            st.session_state.course_dur_weeks = cd
            st.session_state.min_weeks_req = mw
            st.session_state.require_cont = rc
            st.success("規則已儲存")

    st.divider()
    col_q, col_a = st.columns(2)
    q_file = col_q.file_uploader("上傳醫院容額表", type=['xlsx'])
    a_file = col_a.file_uploader("上傳學生志願表", type=['xlsx'])
    
    if st.button("確認並開始比對"):
        if q_file and a_file:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)
            
            if df_a is not None and '姓名' in df_a.columns:
                df_a['姓名'] = df_a['姓名'].ffill()
                apps = []
                for _, row in df_a.iterrows():
                    d_val, t_val = row.get('科別'), row.get('日期欄位')
                    if pd.notna(d_val) and pd.notna(t_val):
                        s, e, d = parse_period_dates(t_val)
                        if s: apps.append({'姓名': row['姓名'], '科別': str(d_val).strip(), '開始': s, '結束': e, '天數': d})
                
                # 容額比對 (含 5/4-5/8)
                date_cols = [c for c in df_q.columns if extract_dates_universal(c)[0]]
                q_dept_col = '科別' if '科別' in df_q.columns else df_q.columns[0]
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get(q_dept_col, '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        try:
                            cap = int(float(re.sub(r'[^0-9.]', '', str(q_row.get(col)))))
                        except: continue
                        s_slot, e_slot = extract_dates_universal(col)
                        st_in = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        if len(st_in) > cap:
                            collisions.append({"科別": dept, "時間": str(col).replace('\n', ' '), "容額": cap, "超額學生": "、".join(list(set(st_in)))})

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                else: st.success("容額核對正常。")
            else: st.error("找不到姓名欄位，請檢查 Excel 表頭。")

elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    m_files = st.file_uploader("上傳各院清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    if st.button("確認並開始比對") and m_files:
        all_d = []
        for f in m_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['姓名'] = df['姓名'].ffill()
                df['來源'] = f.name.replace('.xlsx', '')
                if '日期欄位' in df.columns:
                    df = df[df['日期欄位'].notna()]
                    all_d.append(df)
        
        if all_d:
            full = pd.concat(all_d, ignore_index=True)
            conflicts = []
            for name, gp in full.groupby('姓名'):
                recs = gp.to_dict('records')
                hit = set()
                for i in range(len(recs)):
                    for j in range(i+1, len(recs)):
                        s1, e1 = extract_dates_universal(recs[i]['日期欄位'])
                        s2, e2 = extract_dates_universal(recs[j]['日期欄位'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1): hit.update([i, j])
                if hit:
                    details = "<br>".join([f"• {recs[idx]['來源']} ({str(recs[idx]['日期欄位']).replace('nan','').strip()})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("偵測到重疊佔位")
                html_table = "<table><tr><th>姓名</th><th>衝突詳情</th></tr>"
                for c in conflicts:
                    html_table += f"<tr><td>{c['姓名']}</td><td>{c['衝突詳情']}</td></tr>"
                html_table += "</table>"
                st.markdown(html_table, unsafe_allow_html=True)
            else: st.success("目前查無跨院重複佔位。")
