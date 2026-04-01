import streamlit as st
import pandas as pd
from datetime import datetime
import re

# --- 1. 系統狀態初始化 ---
if "course_dur_weeks" not in st.session_state: st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state: st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state: st.session_state.require_cont = True

# 頁面設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 2. 莫蘭迪高級感 CSS (強制宋體、強制換行、低對比) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'serif' !important;
        background-color: #F5F4F1 !important; 
        color: #5C5E5D !important;
    }
    
    h1, h2, h3 { color: #4A4C4B !important; border-bottom: 1px solid #D6D4CE; padding-bottom: 5px; }
    
    /* 側邊欄 */
    section[data-testid="stSidebar"] { 
        background-color: #EAE8E3 !important; 
        border-right: 1px solid #D6D4CE !important; 
    }
    
    /* 按鈕樣式 (鼠尾草綠) */
    .stButton > button { 
        background-color: #8A9A92 !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 4px !important; 
        transition: 0.3s;
    }
    .stButton > button:hover { background-color: #72827A !important; }
    
    /* 表格：強制換行與條列式 */
    .stTable, [data-testid="stTable"] { font-size: 14px; background-color: #FFFFFF; }
    [data-testid="stTable"] td {
        white-space: pre-wrap !important;
        word-break: break-word !important;
        line-height: 1.8 !important;
        vertical-align: top !important;
        color: #5C5E5D !important;
    }
    [data-testid="stTable"] th {
        background-color: #E3E1DB !important;
        color: #4A4C4B !important;
        text-align: left !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 3. 核心工具函式 (解決日期與讀取問題) ---

def clean_and_extract_dates(text, year=2026):
    """解決換行符號問題，強制提取所有日期數字"""
    if isinstance(text, datetime): return text, text
    # 關鍵：將換行符 \n 換成逗號，並移除所有雜質
    s = str(text).replace('\n', ',').replace('\r', ',').replace(' ', '')
    nums = re.findall(r'\d+', s)
    
    try:
        # 格式：2026.05.04
        if len(nums) >= 3 and len(nums[0]) == 4:
            d1 = datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            d2 = datetime(int(nums[3]), int(nums[4]), int(nums[5])) if len(nums) >= 6 else d1
            return d1, d2
        # 格式：5/4
        elif len(nums) >= 2:
            d1 = datetime(year, int(nums[0]), int(nums[1]))
            d2 = datetime(year, int(nums[2]), int(nums[3])) if len(nums) >= 4 else d1
            return d1, d2
    except: pass
    return None, None

def smart_read_excel(file, keywords):
    """解決標題偏移問題，模糊偵測姓名欄位"""
    try:
        xls = pd.ExcelFile(file)
        target = next((sn for sn in xls.sheet_names if any(k in sn for k in ["志願", "名單", "容額", "工作"])), xls.sheet_names[0])
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        
        for i in range(min(len(df_raw), 25)):
            row_str = "".join([str(x) for x in df_raw.iloc[i].values])
            if any(k in row_str for k in keywords):
                df = pd.read_excel(file, sheet_name=target, header=i)
                df.columns = [str(c).strip() for c in df.columns]
                return df
        return pd.read_excel(file, sheet_name=target)
    except: return None

# --- 4. UI 介面與邏輯 ---

mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
if st.sidebar.button("系統重整"): st.rerun()

if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    with st.expander("規則設定 (點擊儲存)", expanded=True):
        c1, c2, c3 = st.columns(3)
        tmp_dur = c1.number_input("一個 Course 多久 (週)", 1, 10, st.session_state.course_dur_weeks)
        tmp_min = c2.number_input("最短實習週數要求 (週)", 1, 52, st.session_state.min_weeks_req)
        tmp_cont = c3.checkbox("要求必須連續實習", st.session_state.require_cont)
        if st.button("儲存規則設定"):
            st.session_state.course_dur_weeks = tmp_dur
            st.session_state.min_weeks_req = tmp_min
            st.session_state.require_cont = tmp_cont
            st.success("規則已更新")

    st.divider()
    ca, cb = st.columns(2)
    q_f = ca.file_uploader("1. 上傳醫院容額表", type=['xlsx'])
    a_f = cb.file_uploader("2. 上傳學生志願表", type=['xlsx'])
    
    if st.button("確認並開始比對"):
        if q_f and a_f:
            df_q = smart_read_excel(q_f, ["科別", "容額"])
            df_a = smart_read_excel(a_f, ["姓名", "科別", "期間"])
            
            if df_a is not None:
                # 模糊鎖定欄位，解決 KeyError
                name_col = next((c for c in df_a.columns if "姓名" in c), None)
                dept_col = next((c for c in df_a.columns if "科別" in c), None)
                time_col = next((c for c in df_a.columns if "期間" in c or "時間" in c), None)
                
                if not name_col:
                    st.error("志願表中找不到「姓名」欄位，請確認 Excel 第一列是否正確。")
                    st.stop()
                
                df_a['姓名_標準'] = df_a[name_col].ffill()
                apps = []
                for _, r in df_a.iterrows():
                    if pd.notna(r.get(dept_col)) and pd.notna(r.get(time_col)):
                        s, e = clean_and_extract_dates(r[time_col])
                        if s: 
                            apps.append({'姓名': r['姓名_標準'], '科別': str(r[dept_col]).strip(), '開始': s, '結束': e})
                
                # 容額比對
                date_cols = [c for c in df_q.columns if clean_and_extract_dates(c)[0]]
                q_dept_c = next((c for c in df_q.columns if "科別" in c), df_q.columns[0])
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept_name = str(q_row.get(q_dept_c, '')).strip()
                    if not dept_name or dept_name == 'nan': continue
                    for col in date_cols:
                        try:
                            cap_val = re.sub(r'[^\d]+', '', str(q_row[col]))
                            cap = int(cap_val) if cap_val else 0
                        except: continue
                        s_slot, e_slot = clean_and_extract_dates(col)
                        # 關鍵比對邏輯 (Overlapping)
                        st_in = [a['姓名'] for a in apps if a['科別'] == dept_name and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        if len(st_in) > cap:
                            collisions.append({"科別": dept_name, "時間": col.replace('\n',' '), "容額": cap, "超額學生": "、".join(list(set(st_in)))})

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                else: 
                    st.success("容額核對正常。")

elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    m_files = st.file_uploader("上傳各院清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    if st.button("執行跨院比對") and m_files:
        all_d = []
        for f in m_files:
            df = smart_read_excel(f, ["姓名", "科別"])
            if df is not None:
                n_c = next((c for c in df.columns if "姓名" in c), None)
                t_c = next((c for c in df.columns if "期間" in c or "時間" in c), None)
                if n_c and t_c:
                    df['姓名_標準'] = df[n_c].ffill()
                    df['來源醫院'] = f.name.replace('.xlsx', '')
                    df['時間_標準'] = df[t_c]
                    all_d.append(df[df['時間_標準'].notna()])
        if all_d:
            full = pd.concat(all_d, ignore_index=True)
            conflicts = []
            for name, gp in full.groupby('姓名_標準'):
                recs = gp.to_dict('records')
                hit = set()
                for i in range(len(recs)):
                    for j in range(i+1, len(recs)):
                        s1, e1 = clean_and_extract_dates(recs[i]['時間_標準'])
                        s2, e2 = clean_and_extract_dates(recs[j]['時間_標準'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1): hit.update([i, j])
                if hit:
                    details = "\n".join([f"• {recs[idx]['來源醫院']} ({str(recs[idx]['時間_標準']).replace('\n','')})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            if conflicts:
                st.subheader("偵測到重疊佔位")
                st.table(pd.DataFrame(conflicts).set_index('姓名'))
            else: st.success("查無重複佔位。")
