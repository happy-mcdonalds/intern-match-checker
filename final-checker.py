import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re

# --- 1. 初始化系統記憶 ---
if "course_dur_weeks" not in st.session_state: st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state: st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state: st.session_state.require_cont = True

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 2. 溫和版 CSS (黑體、修復 UI 圖示) ---
st.markdown("""
    <style>
    /* 只修改基本文字字體，絕不干擾 icon (避開類別選擇器) */
    html, body, p, span, h1, h2, h3, li, label {
        font-family: "Microsoft JhengHei", "Heiti TC", sans-serif !important;
    }
    
    .stApp { background-color: #F5F4F1; }
    
    h1, h2, h3 { 
        color: #4A4C4B !important; 
        border-bottom: 1px solid #D6D4CE; 
        padding-bottom: 5px; 
    }
    
    /* 按鈕樣式 (鼠尾草綠) */
    .stButton > button, [data-testid="stFormSubmitButton"] > button { 
        background-color: #8A9A92 !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 4px !important; 
        width: 100%; 
        transition: 0.3s;
    }
    .stButton > button:hover { background-color: #72827A !important; }

    /* 網頁自訂表格 (系秘用) */
    .html-table { width: 100%; border-collapse: collapse; background-color: white; border: 1px solid #D6D4CE; }
    .html-table th { background-color: #E3E1DB; color: #4A4C4B; padding: 12px; text-align: left; border-bottom: 2px solid #C0BFB8; }
    .html-table td { padding: 12px; border-bottom: 1px solid #EAE8E3; vertical-align: top; line-height: 1.8; }
    </style>
    """, unsafe_allow_html=True)

# --- 3. 核心工具函式 ---

def smart_read_sheet(file, is_capacity=False):
    """分離式讀取：精準判斷是容額表還是學生名單"""
    try:
        xls = pd.ExcelFile(file)
        sheet = xls.sheet_names[0] # 直接無腦抓第一頁
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None, nrows=20)
        
        # 尋找真正的標題列
        h_idx = 0
        for i, row in df_raw.iterrows():
            row_str = "".join([str(x) for x in row.values])
            if is_capacity:
                # 容額表要找「科別」與「時間」
                if "科別" in row_str and "時間" in row_str:
                    h_idx = i
                    break
            else:
                # 學生名單找「姓名」
                if "姓名" in row_str:
                    h_idx = i
                    break
                    
        df = pd.read_excel(file, sheet_name=sheet, header=h_idx)
        df = df.loc[:, ~df.columns.duplicated()].copy() # 防止重複欄位當機
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
        
        if not is_capacity:
            # 學生名單欄位對接
            rename_map = {}
            start_col, end_col = None, None
            for c in df.columns:
                if "姓名" in c: rename_map[c] = "姓名"
                elif "科別" in c and "備選" not in c: rename_map[c] = "科別"
                elif "期間" in c or "時間" in c: rename_map[c] = "實習期間"
                elif "開始" in c or "起" in c: start_col = c
                elif "結束" in c or "迄" in c: end_col = c
            
            df = df.rename(columns=rename_map)
            
            # 如果沒有單一「實習期間」欄位，合併開始與結束
            if "實習期間" not in df.columns and start_col and end_col:
                df["實習期間"] = df[start_col].astype(str) + " - " + df[end_col].astype(str)
                
            # 過濾範例與無效列
            if "姓名" in df.columns:
                df = df[df["姓名"].notna()]
                df = df[~df["姓名"].astype(str).str.contains("甄漂亮|範例|例|說明|空白|nan", na=False)]
                
        return df
    except Exception as e:
        st.error(f"檔案解析失敗：{e}")
        return None

def extract_dates_universal(text, year=2026):
    """神級日期正則：精準抽取 2026.05.04 或 5/4，不怕任何符號干擾"""
    if pd.isna(text) or str(text).strip() in ['nan', '']: return None, None
    if isinstance(text, datetime): return text, text
    
    s = str(text)
    # 匹配格式如：2026.05.04, 2026-05-04, 5/4, 05/04
    pattern = r'(?:20\d\d[-./])?\d{1,2}[-./]\d{1,2}'
    matches = re.findall(pattern, s)
    
    dates = []
    for m in matches:
        parts = re.split(r'[-./]', m)
        try:
            if len(parts) == 3:
                dates.append(datetime(int(parts[0]), int(parts[1]), int(parts[2])))
            elif len(parts) == 2:
                dates.append(datetime(year, int(parts[0]), int(parts[1])))
        except: pass
        
    if len(dates) >= 2: return dates[0], dates[-1]
    elif len(dates) == 1: return dates[0], dates[0]
    return None, None

def parse_period_dates(p_str):
    s, e = extract_dates_universal(p_str)
    if s and e:
        return s, e, len(pd.bdate_range(s, e))
    return None, None, 0

# --- 4. UI 介面 ---

st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()
if st.sidebar.button("重新整理系統"): 
    st.session_state.clear()
    st.rerun()

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
            # 關鍵修正：告知讀取器哪一份是容額表！
            df_q = smart_read_sheet(q_file, is_capacity=True)
            df_a = smart_read_sheet(a_file, is_capacity=False)
            
            if df_q is not None and df_a is not None and '姓名' in df_a.columns:
                apps = []
                for _, row in df_a.iterrows():
                    d_val, t_val = row.get('科別'), row.get('實習期間')
                    if pd.notna(d_val) and pd.notna(t_val):
                        s, e, d = parse_period_dates(t_val)
                        if s: apps.append({'姓名': row['姓名'], '科別': str(d_val).strip(), '開始': s, '結束': e, '天數': d})
                
                # 容額比對
                date_cols = [c for c in df_q.columns if extract_dates_universal(c)[0]]
                q_dept_col = '科別' if '科別' in df_q.columns else df_q.columns[0]
                collisions = []
                
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get(q_dept_col, '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        raw_cap = str(q_row.get(col)).strip()
                        try:
                            cap = int(float(re.sub(r'[^0-9.]', '', raw_cap))) if raw_cap != 'nan' else 0
                        except: cap = 0
                            
                        s_slot, e_slot = extract_dates_universal(col)
                        st_in = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        if len(st_in) > cap:
                            collisions.append({"科別": dept, "時間": str(col).replace('\n', ' '), "容額": cap, "超額學生": "、".join(list(set(st_in)))})

                # 規章審核
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
                    st.subheader("⚠️ 名額撞期名單")
                    st.dataframe(pd.DataFrame(collisions), use_container_width=True)
                if invalid:
                    st.subheader("📝 規章不符名單")
                    st.dataframe(pd.DataFrame(invalid).drop_duplicates(), use_container_width=True)
                if not collisions and not invalid: st.success("核對完成，查無異常。")
            else: st.error("讀取失敗：請確保容額表有「科別」，學生表有「姓名」。")

elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    m_files = st.file_uploader("上傳各院清單 (單頁版)", type=['xlsx'], accept_multiple_files=True)
    if st.button("確認並開始比對") and m_files:
        all_d = []
        for f in m_files:
            df = smart_read_sheet(f, is_capacity=False)
            if df is not None and '姓名' in df.columns and '實習期間' in df.columns:
                df['來源'] = f.name.replace('.xlsx', '')
                df = df[df['實習期間'].notna() & (df['實習期間'].astype(str) != 'nan')]
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
                        s1, e1 = extract_dates_universal(recs[i]['實習期間'])
                        s2, e2 = extract_dates_universal(recs[j]['實習期間'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1): hit.update([i, j])
                
                if hit:
                    details = "<br>".join([f"• {recs[idx]['來源']} ({str(recs[idx]['實習期間']).replace('nan','').strip()})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("⚠️ 偵測到重疊佔位")
                html_table = "<table class='html-table'><tr><th>姓名</th><th>衝突詳情</th></tr>"
                for c in conflicts:
                    html_table += f"<tr><td>{c['姓名']}</td><td>{c['衝突詳情']}</td></tr>"
                html_table += "</table>"
                st.markdown(html_table, unsafe_allow_html=True)
            else: st.success("查無重複佔位，目前名單一切正常。")
