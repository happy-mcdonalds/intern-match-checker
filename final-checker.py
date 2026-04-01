import streamlit as st
import pandas as pd
from datetime import datetime
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 莫蘭迪色系 + 強制純宋體 CSS (無 Emoji) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'PMingLiU', serif !important;
        background-color: #F2F0EB !important; /* 燕麥灰背景 */
        color: #4A4F4D !important; /* 炭灰色字體 */
    }
    
    h1, h2, h3 { 
        color: #3B403E !important; 
        border-bottom: 1px solid #D1D1C9; 
        padding-bottom: 5px; 
    }
    
    section[data-testid="stSidebar"] { 
        background-color: #E8E6E1 !important; 
        border-right: 1px solid #D1D1C9 !important; 
    }
    
    /* 莫蘭迪按鈕樣式 (鼠尾草綠) */
    .stButton>button { 
        background-color: #7D8B84 !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 4px !important; 
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #65736C !important;
    }

    /* 重新整理與儲存按鈕 (淡灰色) */
    div[data-testid="stVerticalBlock"] > div:nth-child(2) .stButton>button {
        background-color: #B5B5AD !important;
    }

    /* 表格樣式 */
    .stTable { font-size: 14px; background-color: #FFFFFF; border-radius: 4px; }
    th { background-color: #E4E5E0 !important; color: #4A4F4D !important; text-align: left !important; }
    td { white-space: pre-wrap !important; vertical-align: top !important; line-height: 1.6 !important; border-bottom: 1px solid #F0F0F0 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---
def extract_dates_ultra(text, year=2026):
    """最寬鬆日期解析：處理換行與各種間隔符"""
    if isinstance(text, datetime): return text, text
    clean_text = re.sub(r'[^\d]+', ',', str(text)).strip(',')
    nums = clean_text.split(',')
    try:
        if len(nums) >= 3 and len(nums[0]) == 4:
            d1 = datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            d2 = datetime(int(nums[3]), int(nums[4]), int(nums[5])) if len(nums) >= 6 else d1
            return d1, d2
        elif len(nums) >= 2:
            d1 = datetime(year, int(nums[0]), int(nums[1]))
            d2 = datetime(year, int(nums[2]), int(nums[3])) if len(nums) >= 4 else d1
            return d1, d2
    except: pass
    return None, None

def smart_read_sheet(file):
    """自動偵測標題列並統一欄位名稱"""
    try:
        xls = pd.ExcelFile(file)
        target = next((sn for sn in xls.sheet_names if any(k in sn for k in ["志願", "名單", "容額"])), xls.sheet_names[0])
        df_raw = pd.read_excel(file, sheet_name=target)
        
        # 尋找真正的 Header
        header_row = 0
        name_col, dept_col, time_col = None, None, None
        
        for i in range(min(len(df_raw), 20)):
            row_vals = [str(x).strip() for x in df_raw.iloc[i].values]
            if any("姓名" in x or "名單" in x for x in row_vals):
                header_row = i + 1
                df_final = pd.read_excel(file, sheet_name=target, header=header_row)
                # 統一欄位名稱
                df_final.columns = [str(c).strip() for c in df_final.columns]
                for c in df_final.columns:
                    if "姓名" in c: name_col = c
                    if any(k in c for k in ["申請科別", "科別"]): dept_col = c
                    if "實習期間" in c: time_col = c
                return df_final, name_col, dept_col, time_col
        return df_raw, None, None, None
    except: return None, None, None, None

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
if st.sidebar.button("重新整理系統"):
    st.rerun()

# --- 模式：醫院代表 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    st.markdown("### 1. 規則設定")
    c1, c2, c3 = st.columns(3)
    c_dur = c1.number_input("一個 Course 多久 (週)", 1, 10, 2)
    m_weeks = c2.number_input("最短實習週數要求 (週)", 1, 52, 4)
    r_cont = c3.checkbox("要求必須連續實習", True)
    
    if st.button("儲存規則設定"):
        st.success("規則已儲存成功")

    st.divider()
    st.markdown("### 2. 檔案上傳與分析")
    col_a, col_b = st.columns(2)
    q_f = col_a.file_uploader("上傳醫院容額表", type=['xlsx'])
    a_f = col_b.file_uploader("上傳學生志願表", type=['xlsx'])
    
    if st.button("確認並開始比對"):
        if q_f and a_f:
            df_q, _, _, _ = smart_read_sheet(q_f)
            df_a, name_c, dept_c, time_c = smart_read_sheet(a_f)
            
            if df_a is not None and name_c:
                df_a[name_c] = df_a[name_c].ffill()
                apps = []
                for _, r in df_a.iterrows():
                    if pd.notna(r.get(dept_c)) and pd.notna(r.get(time_c)):
                        s, e = extract_dates_ultra(r[time_c])
                        if s: apps.append({'姓名': r[name_c], '科別': str(r[dept_c]).strip(), '開始': s, '結束': e})
                
                # 容額比對
                df_q.columns = [str(c).strip() for c in df_q.columns]
                date_cols = [c for c in df_q.columns if extract_dates_ultra(c)[0]]
                collisions = []
                
                # 取得容額表中的科別欄位
                q_dept_col = next((c for c in df_q.columns if "科別" in c), df_q.columns[0])
                
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get(q_dept_col, '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        try:
                            cap_val = re.sub(r'[^\d]+', '', str(q_row[col]))
                            cap = int(cap_val) if cap_val else 0
                        except: continue
                        
                        s_slot, e_slot = extract_dates_ultra(col)
                        st_in = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        
                        if len(st_in) > cap:
                            collisions.append({"科別": dept, "時間": col.replace('\n',' '), "容額": cap, "超額學生": "、".join(list(set(st_in)))})

                # 規章檢查
                invalid = []
                df_temp = pd.DataFrame(apps)
                if not df_temp.empty:
                    for name, gp in df_temp.groupby('姓名'):
                        gp = gp.sort_values('開始')
                        tot_days = sum([(r['結束'] - r['開始']).days + 1 for _, r in gp.iterrows()])
                        if tot_days < m_weeks * 5:
                            invalid.append({"姓名": name, "原因": f"總實習天數不足 ({tot_days}天)"})
                        for _, r in gp.iterrows():
                            actual = (r['結束'] - r['開始']).days + 1
                            if actual < c_dur * 5:
                                invalid.append({"姓名": name, "原因": f"{r['科別']} 實習週數不足"})
                        if r_cont and len(gp) > 1:
                            for i in range(len(gp)-1):
                                if (gp.iloc[i+1]['開始'] - gp.iloc[i]['結束']).days > 3:
                                    invalid.append({"姓名": name, "原因": "時段未連續實習"})

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("目前分配一切正常")
            else:
                st.error("無法正確偵測學生名單格式，請確認 Excel 中是否有「姓名」欄位。")
        else:
            st.warning("請先上傳檔案。")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    m_files = st.file_uploader("上傳各院清單 (多選)", type=['xlsx'], accept_multiple_files=True)
    if st.button("執行比對") and m_files:
        all_data = []
        for f in m_files:
            df, name_c, dept_c, time_c = smart_read_sheet(f)
            if df is not None and name_c:
                df[name_c] = df[name_c].ffill()
                df['來源'] = f.name.split('.')[0]
                df['_name_col'] = df[name_c]
                df['_time_col'] = df[time_c]
                all_data.append(df[df['_time_col'].notna()])
        
        if all_data:
            full = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name, gp in full.groupby('_name_col'):
                recs = gp.to_dict('records')
                hit = set()
                for i in range(len(recs)):
                    for j in range(i+1, len(recs)):
                        s1, e1 = extract_dates_ultra(recs[i]['_time_col'])
                        s2, e2 = extract_dates_ultra(recs[j]['_time_col'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1):
                            hit.update([i, j])
                if hit:
                    details = "\n".join([f"• {recs[idx]['來源']} ({str(recs[idx]['_time_col']).strip()})" for idx in sorted(list(hit))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("偵測到重疊佔位")
                st.table(pd.DataFrame(conflicts).set_index('姓名'))
            else:
                st.success("無跨院重疊佔位")
