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
    h1, h2, h3 { color: #4A4C4B !important; border-bottom: 1px solid #D6D4CE; padding-bottom: 5px; font-weight: 700; }
    section[data-testid="stSidebar"] { background-color: #EAE8E3 !important; border-right: 1px solid #D6D4CE !important; }
    [data-testid="stForm"] { border: 1px solid #D6D4CE !important; background-color: #FDFDFD !important; border-radius: 4px; padding: 20px; }
    .stButton > button, [data-testid="stFormSubmitButton"] > button { 
        background-color: #8A9A92 !important; color: #FFFFFF !important; border: none !important;
        border-radius: 4px !important; width: 100%; transition: 0.3s;
    }
    .stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover { background-color: #72827A !important; }
    .btn-secondary > button { background-color: #C0BFB8 !important; color: #FFFFFF !important; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th { background-color: #E3E1DB !important; color: #4A4C4B !important; padding: 12px; text-align: left; border-bottom: 2px solid #C0BFB8; }
    td { padding: 12px; border-bottom: 1px solid #EAE8E3; vertical-align: top; line-height: 1.6; white-space: pre-wrap !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 (優化版) ---

def smart_read_sheet(file):
    """超強化版 Excel 讀取：自動掃描 Header 座標與模糊匹配欄位"""
    try:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
        for sn in xls.sheet_names:
            if any(k in sn for k in ["志願", "名單", "實習", "選配", "工作表"]):
                target_sheet = sn
                break
        
        # 1. 讀取原始數據掃描標題列
        df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)
        header_idx = 0
        for i, row in df_raw.iterrows():
            row_str = "".join([str(x) for x in row.values if pd.notna(x)])
            # 只要包含「姓」跟「科」或「期間」就認定為標題列
            if ("姓名" in row_str or "學生" in row_str) and ("科" in row_str or "期間" in row_str or "日期" in row_str):
                header_idx = i
                break
        
        # 2. 正式讀取
        df = pd.read_excel(file, sheet_name=target_sheet, header=header_idx)
        
        # 3. 欄位正名手術 (正則表達式優化)
        clean_cols = {}
        for c in df.columns:
            s_c = re.sub(r'[\s\n\r\t]+', '', str(c)) # 移除所有空白與換行
            if "姓名" in s_c or "學生" in s_c: clean_cols[c] = "姓名"
            elif "科" in s_c and ("別" in s_c or "目" in s_c or "部" in s_c): clean_cols[c] = "科別"
            elif "期間" in s_c or "時段" in s_c or "日期" in s_c: clean_cols[c] = "實習期間"
            else: clean_cols[c] = s_c
            
        df = df.rename(columns=clean_cols)
        
        # 4. 數據清理：處理合併儲存格與無效列
        if "姓名" in df.columns:
            df['姓名'] = df['姓名'].ffill()
            df = df[df['姓名'].notna()] # 濾掉姓名空白的列
        return df
    except Exception as e:
        st.error(f"檔案讀取失敗: {e}")
        return None

def extract_dates_universal(text, year=2026):
    if isinstance(text, datetime): return text, text
    text = re.sub(r'[\n\r\s]+', '-', str(text)).strip()
    parts = re.split(r'[-~～到至_]+', text)
    
    def parse_part(part):
        nums = re.findall(r'\d+', part)
        if len(nums) >= 2:
            y = int(nums[0]) if len(nums[0]) == 4 else year
            m, d = (int(nums[1]), int(nums[2])) if len(nums[0]) == 4 else (int(nums[-2]), int(nums[-1]))
            try: return datetime(y, m, d)
            except: return None
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

# --- UI 邏輯 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()
if st.sidebar.button("重新整理系統"): st.rerun()

if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    with st.form("settings_form"):
        st.markdown("### 規則設定")
        c_cfg1, c_cfg2, c_cfg3 = st.columns([1, 1, 1])
        with c_cfg1: cd_val = st.number_input("Course 週期 (週)", min_value=1, value=st.session_state.course_dur_weeks)
        with c_cfg2: mw_val = st.number_input("最少總週數", min_value=1, value=st.session_state.min_weeks_req)
        with c_cfg3: rc_val = st.checkbox("強制連續實習", value=st.session_state.require_cont)
        if st.form_submit_button("儲存並更新條件"):
            st.session_state.course_dur_weeks, st.session_state.min_weeks_req, st.session_state.require_cont = cd_val, mw_val, rc_val
            st.success("規則已更新")

    st.divider()
    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("1. 上傳醫院容額表", type=['xlsx'])
    with c2: a_file = st.file_uploader("2. 上傳學生志願表", type=['xlsx'])
    
    if st.button("開始核對資料"):
        if not q_file or not a_file:
            st.warning("請先上傳兩個 Excel 檔案")
        else:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)
            
            if df_a is not None and df_q is not None:
                # 再次驗證關鍵欄位是否存在
                missing = [c for c in ["姓名", "科別", "實習期間"] if c not in df_a.columns]
                if missing:
                    st.error(f"學生志願表缺少必要欄位：{', '.join(missing)}")
                    st.info("提示：請檢查 Excel 標題是否包含這些文字。")
                    st.stop()

                # 解析學生申請數據
                apps = []
                for _, row in df_a.iterrows():
                    s, e, d = parse_period_dates(row['實習期間'])
                    if s: apps.append({'姓名': row['姓名'], '科別': str(row['科別']).strip(), '開始': s, '結束': e, '天數': d})
                
                # --- A. 容額衝突檢查 ---
                collisions = []
                # 找出容額表中看起來像日期的欄位
                date_cols = [c for c in df_q.columns if extract_dates_universal(c)[0]]
                q_dept_col = "科別" if "科別" in df_q.columns else df_q.columns[0]
                
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get(q_dept_col, '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        try:
                            cap = int(float(re.sub(r'[^0-9.]', '', str(q_row.get(col, 0)))))
                        except: continue
                        s_slot, e_slot = extract_dates_universal(col)
                        st_in_slot = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        if len(st_in_slot) > cap:
                            collisions.append({"科別": dept, "時段": col, "容額": cap, "超額學生": "、".join(list(set(st_in_slot)))})

                # --- B. 規章符合檢查 ---
                invalid = []
                course_min_days = st.session_state.course_dur_weeks * 5
                total_min_days = st.session_state.min_weeks_req * 5
                
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, group in df_temp.groupby('姓名'):
                        group = group.sort_values('開始')
                        # 1. Course 天數檢查
                        for _, row in group.iterrows():
                            if row['天數'] < course_min_days:
                                invalid.append({"姓名": name, "原因": f"{row['科別']}天數不足({row['天數']}天)"})
                        # 2. 總天數檢查
                        if group['天數'].sum() < total_min_days:
                            invalid.append({"姓名": name, "原因": f"總實習時長不足"})
                        # 3. 連續性檢查
                        if st.session_state.require_cont and len(group) > 1:
                            courses = group.to_dict('records')
                            for i in range(len(courses)-1):
                                if (courses[i+1]['開始'] - courses[i]['結束']).days > 3:
                                    invalid.append({"姓名": name, "原因": "實習中斷(未連續)"})
                                    break

                # --- 顯示結果 ---
                st.header("分析報告")
                if collisions:
                    st.subheader("⚠️ 容額超額警告")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("❌ 規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("✅ 檢查完畢：所有安排皆符合容額與規章。")

elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    multi_files = st.file_uploader("上傳各院清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("執行跨院比對") and multi_files:
        all_data = []
        for f in multi_files:
            df = smart_read_sheet(f)
            if df is not None and '姓名' in df.columns:
                df['來源'] = f.name.split('.')[0]
                all_data.append(df)
        
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    hit = set()
                    for i in range(len(s_apps)):
                        for j in range(i+1, len(s_apps)):
                            s1, e1, _ = parse_period_dates(s_apps[i].get('實習期間'))
                            s2, e2, _ = parse_period_dates(s_apps[j].get('實習期間'))
                            if s1 and s2 and (s1 <= e2 and s2 <= e1):
                                hit.update([i, j])
                    if hit:
                        details = "<br>".join([f"• {s_apps[idx]['來源']}: {s_apps[idx].get('實習期間')}" for idx in sorted(list(hit))])
                        conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("🚨 偵測到重複佔位")
                html_table = "<table style='width:100%'><tr><th>姓名</th><th>衝突說明</th></tr>"
                for c in conflicts:
                    html_table += f"<tr><td>{c['姓名']}</td><td>{c['衝突詳情']}</td></tr>"
                st.markdown(html_table + "</table>", unsafe_allow_html=True)
            else:
                st.success("✅ 無重複佔位情況。")
