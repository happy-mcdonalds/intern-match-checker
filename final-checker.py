import streamlit as st
import pandas as pd
from datetime import datetime
import re

# --- 初始化系統記憶 (Session State) ---
if "course_dur_weeks" not in st.session_state: st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state: st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state: st.session_state.require_cont = True
if "show_results" not in st.session_state: st.session_state.show_results = False

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 莫蘭迪色系 + 強制純宋體 CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'PMingLiU', serif !important;
        background-color: #F5F4F1 !important; color: #5C5E5D !important;
    }
    h1, h2, h3 { color: #4A4C4B !important; border-bottom: 1px solid #D6D4CE; padding-bottom: 5px; }
    section[data-testid="stSidebar"] { background-color: #EAE8E3 !important; border-right: 1px solid #D6D4CE !important; }
    
    /* 按鈕樣式 */
    .stButton > button { 
        background-color: #8A9A92 !important; color: #FFFFFF !important; 
        border: none !important; border-radius: 4px !important; width: 100%; transition: 0.3s;
    }
    .stButton > button:hover { background-color: #72827A !important; }
    .btn-secondary > button { background-color: #C0BFB8 !important; }

    /* 表格樣式 */
    .stTable { font-size: 14px; background-color: #FFFFFF; }
    th { background-color: #E3E1DB !important; color: #4A4C4B !important; }
    td { white-space: pre-wrap !important; vertical-align: top !important; line-height: 1.8 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---
def smart_read_sheet(file):
    try:
        xls = pd.ExcelFile(file)
        target_sheet = next((sn for sn in xls.sheet_names if any(k in sn for k in ["志願", "名單", "工作表4", "實習容額"])), xls.sheet_names[0])
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
    """加強版日期解析：處理換行符號 \n 與不同間隔符"""
    if isinstance(text, datetime): return text, text
    # 移除換行與空格，統一用破折號取代
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
    s, e = extract_dates_universal(p_str)
    return (s, e, len(pd.bdate_range(s, e))) if s and e else (None, None, 0)

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()
if st.sidebar.button("重新整理系統", key="sys_reset"):
    for key in st.session_state.keys(): del st.session_state[key]
    st.rerun()

# --- 醫院代表模式 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 規則設定區
    with st.expander("規則設定 (點擊展開)", expanded=True):
        c_cfg1, c_cfg2, c_cfg3 = st.columns(3)
        st.session_state.course_dur_weeks = c_cfg1.number_input("一個 Course 多久 (週)", 1, 10, st.session_state.course_dur_weeks)
        st.session_state.min_weeks_req = c_cfg2.number_input("最短實習週數要求 (週)", 1, 52, st.session_state.min_weeks_req)
        st.session_state.require_cont = c_cfg3.checkbox("要求必須連續實習", st.session_state.require_cont)
        if st.button("儲存規則條件"): st.success("條件已更新")

    st.divider()
    
    # 檔案上傳區
    c1, c2 = st.columns(2)
    q_file = c1.file_uploader("上傳醫院容額表", type=['xlsx'])
    a_file = c2.file_uploader("上傳學生志願表", type=['xlsx'])
    
    if st.button("確認並開始比對"):
        if q_file and a_file: st.session_state.show_results = True
        else: st.warning("請先上傳兩個檔案")

    if st.session_state.show_results and q_file and a_file:
        try:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)
            if df_q is not None and df_a is not None:
                df_a['姓名'] = df_a['姓名'].ffill()
                dept_col = "申請科別" if "申請科別" in df_a.columns else "科別"
                apps = []
                for _, r in df_a.iterrows():
                    if pd.notna(r.get(dept_col)) and pd.notna(r.get('實習期間')):
                        s, e, d = parse_period_dates(r['實習期間'])
                        if s: apps.append({'姓名': r['姓名'], '科別': str(r[dept_col]).strip(), '開始': s, '結束': e, '天數': d})
                
                # 容額比對
                date_cols = [c for c in df_q.columns if extract_dates_universal(c)[0]]
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        try: cap = int(float(re.sub(r'[^0-9.]', '', str(q_row[col]))))
                        except: continue
                        s_slot, e_slot = extract_dates_universal(col)
                        st_in = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        if len(st_in) > cap:
                            collisions.append({"科別": dept, "時間": str(col).replace('\n',' '), "容額": cap, "超額學生": "、".join(list(set(st_in)))})

                # 規章檢查
                invalid = []
                df_temp = pd.DataFrame(apps)
                for name, gp in df_temp.groupby('姓名'):
                    gp = gp.sort_values('開始')
                    if gp['天數'].sum() < st.session_state.min_weeks_req * 5:
                        invalid.append({"姓名": name, "原因": f"總時長不足 ({gp['天數'].sum()} 天)"})
                    for _, r in gp.iterrows():
                        if r['天數'] < st.session_state.course_dur_weeks * 5:
                            invalid.append({"姓名": name, "原因": f"{r['科別']} Course 天數不足 ({r['天數']} 天)"})
                    if st.session_state.require_cont and len(gp) > 1:
                        for i in range(len(gp)-1):
                            if (gp.iloc[i+1]['開始'] - gp.iloc[i]['結束']).days > 3:
                                invalid.append({"姓名": name, "原因": "未連續實習 (中斷)"})

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid: st.success("核對完成，查無異常。")
        except Exception as e: st.error(f"分析失敗: {e}")

# --- 系秘模式 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    if st.button("執行跨院比對") and multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['姓名'] = df['姓名'].ffill()
                df['來源醫院'] = f.name.split('.')[0]
                all_data.append(df[df['實習期間'].notna()])
        if all_data:
            full = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name, gp in full.groupby('姓名'):
                recs = gp.to_dict('records')
                hit = set()
                for i in range(len(recs)):
                    for j in range(i+1, len(recs)):
                        s1, e1, _ = parse_period_dates(recs[i]['實習期間'])
                        s2, e2, _ = parse_period_dates(recs[j]['實習期間'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1): hit.update([i, j])
                if hit:
                    details = "\n".join([f"• {recs[idx]['來源醫院']} ({str(recs[idx]['實習期間']).strip()})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            if conflicts:
                st.subheader("偵測到重疊佔位")
                st.table(pd.DataFrame(conflicts).set_index('姓名'))
            else: st.success("無跨院重疊佔位。")
