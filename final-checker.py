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

# --- 莫蘭迪色系 + 全域黑體 (修復 Icon 消失問題) ---
st.markdown("""
    <style>
    /* 全域使用黑體，但不干擾 Streamlit 內建 Icon */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        font-family: "Microsoft JhengHei", "Heiti TC", "Apple LiGothic Medium", sans-serif !important;
        background-color: #F5F4F1 !important; 
        color: #5C5E5D !important; 
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
    
    /* 按鈕 (鼠尾草綠) */
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
    
    /* 網頁表格樣式 (強制換行) */
    .html-table { width: 100%; border-collapse: collapse; background-color: white; border: 1px solid #D6D4CE; font-family: "Microsoft JhengHei", sans-serif; }
    .html-table th { background-color: #E3E1DB; color: #4A4C4B; padding: 12px; text-align: left; border-bottom: 2px solid #C0BFB8; }
    .html-table td { padding: 12px; border-bottom: 1px solid #EAE8E3; vertical-align: top; line-height: 1.6; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---

def smart_read_sheet(file):
    """改良讀取：精準尋找「確定名單」，並將日期欄位自動對齊"""
    try:
        xls = pd.ExcelFile(file)
        
        # 1. 優先尋找確定名單 (避開充滿甄漂亮的志願申請名單)
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if "確定" in sn or "正式" in sn:
                target_sheet = sn
                break
        else:
            # 退而求其次尋找包含名單的分頁，但避開「先寫這個」
            for sn in xls.sheet_names:
                if any(k in sn for k in ["志願", "名單", "容額"]) and "先寫" not in sn:
                    target_sheet = sn
                    break
        
        # 2. 掃描標題列
        df_scan = pd.read_excel(file, sheet_name=target_sheet, header=None, nrows=15)
        h_idx = 0
        for i, row in df_scan.iterrows():
            row_str = "".join([str(x) for x in row.values])
            if any(k in row_str for k in ["姓名", "科別", "日期"]):
                h_idx = i
                break
        
        df = pd.read_excel(file, sheet_name=target_sheet, header=h_idx)
        
        # 解決重複標籤問題
        df = df.loc[:, ~df.columns.duplicated()].copy()
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
        
        # 3. 欄位自動對接 (模糊匹配)
        rename_map = {}
        for c in df.columns:
            if "姓名" in c: rename_map[c] = "姓名"
            elif "科別" in c or "申請科別" in c: rename_map[c] = "科別"
            # 只有當沒有「開始/結束」時，才將日期欄位直接改名，避免重複
            elif ("實習期間" in c or "日期" in c) and "開始" not in c and "結束" not in c: 
                rename_map[c] = "日期欄位"
        df = df.rename(columns=rename_map)
        
        # 4. 處理日期分開兩格的情況 (開始日期, 結束日期)
        if "日期欄位" not in df.columns:
            start = next((c for c in df.columns if "開始" in c), None)
            end = next((c for c in df.columns if "結束" in c), None)
            if start and end:
                df["日期欄位"] = df[start].astype(str) + " - " + df[end].astype(str)
        
        # 5. 過濾掉「甄漂亮」或「範例」，確保只留真實學生
        if "姓名" in df.columns:
            df = df[~df['姓名'].astype(str).str.contains('甄漂亮|範例|例|說明|空白', na=False)]
            
        return df
    except Exception as e:
        return None

def extract_dates_universal(text, year=2026):
    if pd.isna(text) or str(text).strip() == 'nan': return None, None
    if isinstance(text, datetime): return text, text
    # 將所有換行、空格轉為橫線，並防止雙橫線導致解析失敗
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
        return s, e, len(pd.bdate_range(s, e))
    return None, None, 0

# --- UI 介面 ---

st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()
st.sidebar.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
if st.sidebar.button("重新整理系統"): 
    st.session_state.clear()
    st.rerun()
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
    ca, cb = st.columns(2)
    q_file = ca.file_uploader("上傳醫院容額表", type=['xlsx'], key="q")
    a_file = cb.file_uploader("上傳學生志願表", type=['xlsx'], key="a")
    
    if st.button("確認並開始比對"):
        if q_file and a_file:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)
            
            if df_a is not None and '姓名' in df_a.columns:
                df_a['姓名'] = df_a['姓名'].ffill()
                apps = []
                # 確保姓名欄位不為空
                df_a = df_a[df_a['姓名'].notna()]
                
                for _, row in df_a.iterrows():
                    d_val, t_val = row.get('科別'), row.get('日期欄位')
                    if pd.notna(d_val) and pd.notna(t_val):
                        s, e, d = parse_period_dates(t_val)
                        if s: apps.append({'姓名': row['姓名'], '科別': str(d_val).strip(), '開始': s, '結束': e, '天數': d})
                
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

                invalid = []
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, gp in df_temp.groupby('姓名'):
                        gp = gp.sort_values('開始')
                        if gp['天數'].sum() < st.session_state.min_weeks_req * 5:
                            invalid.append({"姓名": name, "原因": f"總實習天數不足 ({gp['天數'].sum()} 天)"})
                        for _, r in gp.iterrows():
                            if r['天數'] < st.session_state.course_dur_weeks * 5:
                                invalid.append({"姓名": name, "原因": f"{r['科別']} 週數不足"})
                        if st.session_state.require_cont and len(gp) > 1:
                            recs = gp.to_dict('records')
                            for i in range(len(recs)-1):
                                if (recs[i+1]['開始'] - recs[i]['結束']).days > 3:
                                    invalid.append({"姓名": name, "原因": "時段未連續實習"})
                                    break

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid: st.success("核對完成，查無異常。")
            else: st.error("檔案內找不到學生姓名，請確認上傳的表格是否有資料。")

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
                    df = df[df['日期欄位'].notna() & (df['日期欄位'].astype(str) != 'nan')]
                    all_d.append(df)
        
        if all_d:
            full = pd.concat(all_d, ignore_index=True)
            conflicts = []
            for name, gp in full.groupby('姓名'):
                recs = gp.to_dict('records')
                if len(recs) < 2: continue
                hit = set()
                for i in range(len(recs)):
                    for j in range(i+1, len(recs)):
                        s1, e1 = extract_dates_universal(recs[i]['日期欄位'])
                        s2, e2 = extract_dates_universal(recs[j]['日期欄位'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1): hit.update([i, j])
                if hit:
                    # 使用 HTML 的 <br> 確保換行
                    details = "<br>".join([f"• {recs[idx]['來源']} ({str(recs[idx]['日期欄位']).strip()})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("偵測到跨院重複佔位")
                # 使用 HTML 繪製表格
                html_code = '<table class="html-table"><tr><th>姓名</th><th>衝突詳情</th></tr>'
                for c in conflicts:
                    html_code += f'<tr><td>{c["姓名"]}</td><td>{c["衝突詳情"]}</td></tr>'
                html_code += '</table>'
                st.markdown(html_code, unsafe_allow_html=True)
            else: st.success("查無重複佔位，目前名單一切正常。")
