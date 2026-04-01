import streamlit as st
import pandas as pd
from datetime import datetime
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 高級感 CSS (宋體 + 黑白灰) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', serif !important;
        color: #000000;
    }
    h1, h2, h3 { color: #000000 !important; border-bottom: 1px solid #000000; padding-bottom: 5px; }
    .stApp { background-color: #FFFFFF; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #DDDDDD; }
    .stButton>button { color: #FFFFFF !important; background-color: #000000 !important; border-radius: 0px; width: 100%; }
    .stTable { font-size: 14px; border: 1px solid #000000; }
    </style>
    """, unsafe_allow_html=True)

# --- 1. 智慧讀取與欄位標準化 (解決找不到姓名問題) ---
def smart_read_sheet(file):
    try:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if any(k in sn for k in ["志願", "名單", "時段", "工作表4"]):
                target_sheet = sn
                break
        
        # 掃描前 20 行找標題
        df_preview = pd.read_excel(file, sheet_name=target_sheet, nrows=20, header=None)
        header_idx = 0
        for i, row in df_preview.iterrows():
            row_str = [str(x).strip() for x in row.values]
            if any(k in "".join(row_str) for k in ["姓名", "科別", "期間"]):
                header_idx = i
                break
        
        df = pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
        
        # 強制更名邏輯：只要欄位名包含關鍵字就統一
        new_cols = {}
        for c in df.columns:
            c_str = str(c).replace(" ", "").replace("\n", "")
            if "姓名" in c_str: new_cols[c] = "姓名"
            elif "學號" in c_str: new_cols[c] = "學號"
            elif "申請科別" in c_str: new_cols[c] = "申請科別"
            elif "科別" in c_str: new_cols[c] = "科別"
            elif "期間" in c_str or "時間" in c_str: new_cols[c] = "實習期間"
            else: new_cols[c] = c_str
        return df.rename(columns=new_cols)
    except: return None

# --- 2. 日期解析工具 ---
def parse_date_simple(s, year=2026):
    try:
        if "." in str(s):
            m = re.search(r'\d{4}\.\d{2}\.\d{2}', str(s))
            if m: return datetime.strptime(m.group(), "%Y.%m.%d")
        pts = re.findall(r'\d+', str(s))
        if len(pts) >= 2: return datetime(year, int(pts[0]), int(pts[1]))
    except: pass
    return None

def parse_period(p_str):
    try:
        ds = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(p_str).replace('\n',''))
        if len(ds) >= 2:
            s = datetime.strptime(ds[0], "%Y.%m.%d"); e = datetime.strptime(ds[1], "%Y.%m.%d")
            return s, e, (e - s).days + 1
    except: pass
    return None, None, 0

# --- 3. 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()

# --- 醫院代表模式 (精準碰撞版) ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1: course_dur = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
    with col_cfg2: min_weeks = st.number_input("最短實習週數要求 (週)", min_value=1, value=4)
    require_cont = st.checkbox("要求必須連續實習", value=True)

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("上傳醫院容額表", type=['xlsx'], key="h_q")
    with c2: a_file = st.file_uploader("上傳學生志願表", type=['xlsx'], key="h_a")

    if q_file and a_file:
        df_q = smart_read_sheet(q_file)
        df_a = smart_read_sheet(a_file)
        
        if df_a is not None and "姓名" in df_a.columns:
            df_a['姓名'] = df_a['姓名'].ffill()
            dept_col = "申請科別" if "申請科別" in df_a.columns else "科別"
            
            # 解析志願
            apps = []
            for _, row in df_a.iterrows():
                if pd.notna(row.get(dept_col)) and pd.notna(row.get('實習期間')):
                    s, e, d = parse_period(row['實習期間'])
                    if s: apps.append({'姓名': row['姓名'], '科別': str(row[dept_col]).strip(), '開始': s, '結束': e, '天數': d})
            
            # 精準容額碰撞偵測
            collisions = []
            if df_q is not None and "科別" in df_q.columns:
                date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        try: cap = int(float(q_row[col]))
                        except: continue
                        pts = col.split('-'); s_s = parse_date_simple(pts[0]); e_s = parse_date_simple(pts[1]) if len(pts)>1 else s_s
                        if not s_s or not e_s: continue
                        # 核心比對：學生日期是否蓋到這週日期
                        sts = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_s and s_s <= a['結束']]
                        if len(sts) > cap:
                            collisions.append({"科別": dept, "時間": col, "容額": cap, "超額學生": "、".join(list(set(sts)))})

            st.header("異常監控結果")
            if collisions:
                st.subheader("名額撞期名單")
                st.table(pd.DataFrame(collisions))
            
            # 規章判定
            invalid = []
            if apps:
                df_tmp = pd.DataFrame(apps)
                for name, gp in df_tmp.groupby('姓名'):
                    total_d = gp['天數'].sum()
                    for _, r in gp.iterrows():
                        if r['天數'] < course_dur * 5:
                            invalid.append({"姓名": name, "原因": f"{r['科別']} 天數不足 ({r['天數']}天)"})
                    if total_d < min_weeks * 5:
                        invalid.append({"姓名": name, "原因": f"總天數 ({total_d}天) 未達最低要求"})
            
            if invalid:
                st.subheader("規章不符名單")
                st.table(pd.DataFrame(invalid).drop_duplicates())
            if not collisions and not invalid: st.success("核對完成，目前一切正常。")
        else: st.error("找不到『姓名』欄位，請確認志願表內容。")

# --- 系秘模式 (智慧多檔版) ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    m_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    if m_files:
        all_d = []
        for f in m_files:
            df = smart_read_sheet(f)
            if df is not None and "姓名" in df.columns:
                df['姓名'] = df['姓名'].ffill(); df['來源醫院'] = f.name
                all_d.append(df[df['實習期間'].notna()])
        
        if all_d:
            full = pd.concat(all_d, ignore_index=True)
            conflicts = []
            for name in full['姓名'].unique():
                s_aps = full[full['姓名'] == name].to_dict('records')
                if len(s_aps) > 1:
                    for i in range(len(s_aps)):
                        for j in range(i + 1, len(s_aps)):
                            s1, e1, _ = parse_period(s_aps[i]['實習期間'])
                            s2, e2, _ = parse_period(s_aps[j]['實習期間'])
                            if s1 and s2 and (s1 <= e2 and s2 <= e1):
                                conflicts.append({"姓名": name, "醫院 A": s_aps[i]['來源醫院'], "時間 A": s_aps[i]['實習期間'], "醫院 B": s_aps[j]['來源醫院'], "時間 B": s_aps[j]['實習期間']})
            if conflicts:
                st.subheader("偵測到跨院衝突名單")
                st.table(pd.DataFrame(conflicts).drop_duplicates())
            else: st.success("交叉比對完成，無重複佔位。")
