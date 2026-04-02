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

# --- 2. 極簡莫蘭迪 CSS (絕不干擾 Icon) ---
st.markdown("""
    <style>
    /* 移除強制的 font-family，回歸原生乾淨介面並修復 Icon */
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

def smart_read_sheet(file, sheet_hints):
    """智慧讀取引擎：自動尋找正確分頁，並完美對接北榮與容額表格式"""
    try:
        xls = pd.ExcelFile(file)
        
        # 根據身份尋找正確的分頁
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if any(hint in sn for hint in sheet_hints):
                target_sheet = sn
                break
        
        # 尋找真正的標題列
        df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None, nrows=20)
        h_idx = 0
        for i, row in df_raw.iterrows():
            row_str = "".join([str(x) for x in row.values])
            if "姓名" in row_str or "科別" in row_str:
                h_idx = i
                break
                
        df = pd.read_excel(file, sheet_name=target_sheet, header=h_idx)
        
        # 清除重複欄位，避免當機
        df = df.loc[:, ~df.columns.duplicated()].copy() 
        df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
        
        # 處理北榮格式：若有「開始」和「結束」兩格，自動合併
        start_col = next((c for c in df.columns if "開始" in c or "起" in c), None)
        end_col = next((c for c in df.columns if "結束" in c or "迄" in c), None)
        if start_col and end_col:
            df["日期欄位"] = df[start_col].astype(str) + " - " + df[end_col].astype(str)
            
        # 欄位自動對接
        rename_map = {}
        for c in df.columns:
            if "姓名" in c: rename_map[c] = "姓名"
            elif "科別" in c and "備選" not in c: rename_map[c] = "科別"
            elif ("期間" in c or "時間" in c or "日期" in c) and c not in [start_col, end_col, "日期欄位"]:
                if "日期欄位" not in df.columns:
                    rename_map[c] = "日期欄位"
        
        df = df.rename(columns=rename_map)
        
        # 注意：這裡「絕對不能」把 NA 刪掉，否則志願表的第二列會遺失！
        # 只過濾掉官方範例 (如甄漂亮)
        if "姓名" in df.columns:
            mask = df["姓名"].astype(str).str.contains("甄漂亮|範例|例|說明|空白", na=False)
            df = df[~mask]
                
        return df
    except Exception as e:
        return None

def extract_dates_universal(text, year=2026):
    """不受符號干擾的神級日期萃取"""
    if pd.isna(text) or str(text).strip() in ['nan', '']: return None, None
    if isinstance(text, datetime): return text, text
    
    # 直接抓取所有「像日期的結構」
    pattern = r'(?:20\d\d[-./_])?\d{1,2}[-./_]\d{1,2}'
    matches = re.findall(pattern, str(text))
    
    dates = []
    for m in matches:
        parts = re.split(r'[-./_]', m)
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
        if e < s: s, e = e, s
        # 使用 bdate_range 精確計算工作天
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
    a_file = col_a.file_uploader("上傳學生志願表 (申請名單)", type=['xlsx'])
    
    if st.button("確認並開始比對"):
        if q_file and a_file:
            # 醫院代表專用分頁提示
            df_q = smart_read_sheet(q_file, ["容額", "時段", "空白"])
            df_a = smart_read_sheet(a_file, ["志願", "申請"])
            
            if df_q is not None and df_a is not None and '姓名' in df_a.columns:
                # 關鍵修復：向下填補姓名，挽救因為儲存格合併而空白的第二志願
                df_a['姓名'] = df_a['姓名'].ffill()
                df_a = df_a[df_a['姓名'].notna()]
                
                apps = []
                for _, row in df_a.iterrows():
                    d_val, t_val = row.get('科別'), row.get('日期欄位')
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
                        # 顯示給使用者看的詳細天數
                        if gp['天數'].sum() < st.session_state.min_weeks_req * 5:
                            invalid.append({"姓名": name, "原因": f"總實習天數不足 (累計 {gp['天數'].sum()} 個工作天)"})
                        for _, r in gp.iterrows():
                            if r['天數'] < st.session_state.course_dur_weeks * 5:
                                invalid.append({"姓名": name, "原因": f"{r['科別']} 週數不足 (僅 {r['天數']} 個工作天)"})
                        if st.session_state.require_cont and len(gp) > 1:
                            recs = gp.to_dict('records')
                            for i in range(len(recs)-1):
                                if (recs[i+1]['開始'] - recs[i]['結束']).days > 3:
                                    invalid.append({"姓名": name, "原因": "時段未連續實習"})
                                    break

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.dataframe(pd.DataFrame(collisions), use_container_width=True)
                if invalid:
                    st.subheader("規章不符名單")
                    st.dataframe(pd.DataFrame(invalid).drop_duplicates(), use_container_width=True)
                if not collisions and not invalid: st.success("核對完成，查無異常。")
            else: st.error("讀取失敗：請確保容額表有「科別」，學生表有「姓名」。")

elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    m_files = st.file_uploader("上傳各院清單 (北榮確定名單格式)", type=['xlsx'], accept_multiple_files=True)
    if st.button("確認並開始比對") and m_files:
        all_d = []
        for f in m_files:
            # 系秘專用分頁提示
            df = smart_read_sheet(f, ["確定", "名單", "正式"])
            if df is not None and '姓名' in df.columns and '日期欄位' in df.columns:
                df['姓名'] = df['姓名'].ffill()
                df['來源'] = f.name.replace('.xlsx', '')
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
                    details = "<br>".join([f"• {recs[idx]['來源']} ({str(recs[idx]['日期欄位']).replace('nan','').strip()})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("偵測到重疊佔位")
                html_table = "<table class='html-table'><tr><th>姓名</th><th>衝突詳情</th></tr>"
                for c in conflicts:
                    html_table += f"<tr><td>{c['姓名']}</td><td>{c['衝突詳情']}</td></tr>"
                html_table += "</table>"
                st.markdown(html_table, unsafe_allow_html=True)
            else: st.success("查無重複佔位，目前名單一切正常。")
