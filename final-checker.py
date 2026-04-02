import streamlit as st
import pandas as pd
from datetime import datetime
import re

# --- 初始化系統記憶 ---
if "course_dur_weeks" not in st.session_state:
    st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state:
    st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state:
    st.session_state.require_cont = True

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- CSS 視覺設定 (移除所有上傳框魔改，回歸原生穩定版) ---
st.markdown("""
    <style>
    * { font-family: "PingFang TC", "微軟正黑體", sans-serif !important; }
    html, body, .stApp { background-color: #F6F5F2 !important; color: #5C5E5D !important; }
    h1, h2, h3 { color: #3B4441 !important; border-bottom: 2px solid #D6D4CE; padding-bottom: 8px; font-weight: bold; }
    
    [data-testid="stForm"] {
        border: 1px solid #D6D4CE !important; background-color: #FFFFFF !important;
        border-radius: 8px; padding: 24px; box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
    
    /* 按鈕顏色統一 */
    .stButton > button, [data-testid="stFormSubmitButton"] > button { 
        background-color: #8A9A92 !important; color: #FFFFFF !important; 
        border: none !important; border-radius: 6px !important; width: 100%; 
        font-size: 16px !important; font-weight: bold !important; transition: 0.3s;
    }
    .stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
        background-color: #72827A !important; transform: translateY(-1px);
    }
    .btn-secondary > button { background-color: #C0BFB8 !important; color: #FFFFFF !important; }
    .btn-secondary > button:hover { background-color: #A8A7A0 !important; }

    /* 表格設計 */
    th { background-color: #E6E4DF !important; color: #4A4C4B !important; font-weight: bold !important; }
    td, th { padding: 10px 12px !important; line-height: 1.5 !important; }
    </style>
""", unsafe_allow_html=True)

# --- 核心工具函式 (全新升級：避開空白範本、自動填補姓名) ---
def read_excel_or_csv(file, mode="app"):
    try:
        # 讀取檔案內所有的 sheet
        if file.name.endswith('.csv'):
            df_dict = {"sheet1": pd.read_csv(file, header=None)}
        else:
            df_dict = pd.read_excel(file, sheet_name=None, header=None)
            
        target_df = None
        
        # 1. 篩選分頁：避開空白範本
        for sheet_name, df_temp in df_dict.items():
            if "確定" in sheet_name or "規範" in sheet_name or "基本資料" in sheet_name:
                continue # 絕對跳過這些已知無用的表單
            
            # 檢查內容是否符合特徵
            sample_text = df_temp.iloc[:15].astype(str).sum().sum()
            if mode == "app" and "姓名" in sample_text and ("科別" in sample_text or "志願" in sample_text):
                target_df = df_temp
                break
            elif mode == "cap" and "科別" in sample_text and ("時段" in sample_text or "/" in sample_text):
                target_df = df_temp
                break
                    
        # 如果沒找到特徵，就拿第一張不是「確定」的表
        if target_df is None:
            for sheet_name, df_temp in df_dict.items():
                if "確定" not in sheet_name:
                    target_df = df_temp
                    break

        if target_df is None: return None

        # 2. 定位真正的標題列在哪裡
        header_idx = 0
        for i in range(min(len(target_df), 15)):
            row_str = "".join([str(x).replace(" ", "") for x in target_df.iloc[i].values])
            if mode == "app" and "姓名" in row_str: 
                header_idx = i; break
            elif mode == "cap" and "科別" in row_str: 
                header_idx = i; break
                
        # 3. 整理 DataFrame
        df = target_df.iloc[header_idx:].reset_index(drop=True)
        # 統一去掉欄位名稱的換行跟空白
        df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.iloc[0]]
        df = df[1:].reset_index(drop=True)
        df = df.dropna(how='all') # 清除全空的列
        
        # 4. 欄位正規化與姓名補齊
        if mode == "app":
            # 統一實習期間
            period_col = next((c for c in df.columns if "期間" in c), None)
            start_col = next((c for c in df.columns if "開始" in c), None)
            end_col = next((c for c in df.columns if "結束" in c), None)
            if period_col:
                df.rename(columns={period_col: "實習期間"}, inplace=True)
            elif start_col and end_col:
                df["實習期間"] = df[start_col].astype(str) + "-" + df[end_col].astype(str)
            
            # 統一科別名稱
            dept_col = next((c for c in df.columns if "申請科" in c), None) or next((c for c in df.columns if "科別" in c), None)
            if dept_col:
                df.rename(columns={dept_col: "科別"}, inplace=True)
                
            # 填補合併儲存格造成的姓名空白
            name_col = next((c for c in df.columns if "姓名" in c), None)
            if name_col:
                df.rename(columns={name_col: "姓名"}, inplace=True)
                df['姓名'] = df['姓名'].replace('nan', pd.NA).ffill()
                df = df.dropna(subset=['姓名']) # 把真的沒有姓名的列清掉
                
        return df
    except Exception as e:
        return None

def extract_dates_universal(text, year=2026):
    if isinstance(text, datetime): return text, text
    text = str(text).replace('\n', '-').replace('\r', '-').replace(' ', '').strip()
    parts = re.split(r'[-~～到至_]+', text)
    
    def extract_single_date(part):
        nums = re.findall(r'\d+', part)
        if len(nums) >= 2:
            if len(nums[0]) == 4 and len(nums) >= 3:
                return datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            return datetime(year, int(nums[-2]), int(nums[-1]))
        return None
       
    dates = [extract_single_date(p) for p in parts if extract_single_date(p) is not None]
    if len(dates) == 1: return dates[0], dates[0]
    elif len(dates) >= 2: return dates[0], dates[-1]
    return None, None

def parse_period_dates(p_str):
    try:
        s, e = extract_dates_universal(p_str)
        if s and e: return s, e, len(pd.bdate_range(s, e))
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
    with c1: 
        st.markdown("**1. 志願申請名單**")
        a_file = st.file_uploader("上傳申請名單 (CSV/XLSX)", type=['xlsx', 'csv'], key="app_up")
    with c2: 
        st.markdown("**2. 實習容額與時段表**")
        q_file = st.file_uploader("上傳容額表 (CSV/XLSX)", type=['xlsx', 'csv'], key="cap_up")
    
    run_check = st.button("確認並開始比對")

    if run_check and q_file and a_file:
        course_workdays = st.session_state.course_dur_weeks * 5
        total_min_workdays = st.session_state.min_weeks_req * 5
        
        try:
            # 完美讀取資料，不再受到空白 sheet 影響
            df_q = read_excel_or_csv(q_file, mode="cap")
            df_a = read_excel_or_csv(a_file, mode="app")

            if df_q is not None and df_a is not None:
                apps = []
                for _, row in df_a.iterrows():
                    dept = str(row.get('科別', '')).strip()
                    period = str(row.get('實習期間', '')).strip()
                    if dept and dept != 'nan' and period and period != 'nan':
                        s, e, d = parse_period_dates(period)
                        if s: apps.append({'姓名': row['姓名'], '科別': dept, '開始': s, '結束': e, '天數': d})
                
                date_cols = []
                slot_mapping = {}
                for c in df_q.columns:
                    s_slot, e_slot = extract_dates_universal(c)
                    if s_slot and e_slot:
                        date_cols.append(c)
                        slot_mapping[c] = (s_slot, e_slot)

                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if dept == 'nan' or not dept: continue
    
                    for col in date_cols:
                        cap = q_row.get(col)
                        try: cap_val = int(float(re.sub(r'[^0-9.]', '', str(cap))))
                        except: continue
                        
                        s_slot, e_slot = slot_mapping[col]
                        st_in_slot = []
                        for a in apps:
                            if a['科別'] == dept:
                                if a['開始'] <= e_slot and a['結束'] >= s_slot:
                                    st_in_slot.append(a['姓名'])
                        
                        if len(st_in_slot) > cap_val:
                            collisions.append({
                                "科別": dept, 
                                "時間": str(col).replace('\n', ''), 
                                "容額": cap_val, 
                                "超額學生": "、".join(list(set(st_in_slot)))
                            })

                invalid = []
                if apps:
                    df_temp = pd.DataFrame(apps)
                    for name, group in df_temp.groupby('姓名'):
                        group = group.sort_values('開始')
                        total_workdays = group['天數'].sum()
                        
                        for _, row in group.iterrows():
                            if row['天數'] < course_workdays:
                                invalid.append({"姓名": name, "原因": f"Course 天數不足：{row['科別']} 僅 {row['天數']} 個工作天 (需 {course_workdays} 天)"})
                        
                        if total_workdays < total_min_workdays:
                            invalid.append({"姓名": name, "原因": f"總時長不足：僅 {total_workdays} 個工作天 (需 {total_min_workdays} 天)"})
                        
                        if st.session_state.require_cont and len(group) > 1:
                            courses = group.to_dict('records')
                            for i in range(len(courses) - 1):
                                prev_end = courses[i]['結束']
                                next_start = courses[i+1]['開始']
                                if (next_start - prev_end).days > 3:
                                    invalid.append({"姓名": name, "原因": f"未連續實習：{courses[i]['科別']} 與 {courses[i+1]['科別']} 中斷"})
                    
                st.header("異常監控結果")
                if collisions:
                    st.subheader("⚠️ 名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("⚠️ 規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("🎉 太棒了！名額分配與規章核對完全符合規定。")
            else:
                st.error("讀取檔案失敗，請確認檔案格式是否正確。")
        except Exception as e: 
            st.error(f"解析過程中發生錯誤：{e}")

# --- 模式：系秘 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    st.markdown("### 檔案上傳")
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx', 'csv'], accept_multiple_files=True)
    run_check_sec = st.button("確認並開始比對")
        
    if run_check_sec and multi_files:
        all_data = []
        for f in multi_files:
            df = read_excel_or_csv(f, mode="app")
            if df is not None and '姓名' in df.columns and '實習期間' in df.columns:
                clean_hosp_name = f.name.replace('.xlsx', '').replace('.csv', '')
                df['來源醫院'] = clean_hosp_name
                all_data.append(df[df['實習期間'].notna()])
                
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            
            for name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == name].to_dict('records')
                if len(s_apps) > 1:
                    conflict_set = set()
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            d1_s, d1_e, _ = parse_period_dates(s_apps[i]['實習期間'])
                            d2_s, d2_e, _ = parse_period_dates(s_apps[j]['實習期間'])
                            if d1_s and d2_s and (d1_s <= d2_e and d2_s <= d1_e):
                                conflict_set.add(i)
                                conflict_set.add(j)
                    
                    if conflict_set:
                        details = []
                        for idx in sorted(list(conflict_set)):
                            hosp = s_apps[idx]['來源醫院']
                            period = str(s_apps[idx]['實習期間']).replace('\n', '')
                            details.append(f"- {hosp} ({period})")
                        
                        conflicts.append({"姓名": name, "衝突詳情": "\n".join(details)})
                        
            if conflicts:
                st.subheader("⚠️ 偵測到跨院衝突名單")
                st.table(pd.DataFrame(conflicts).set_index('姓名'))
            else: 
                st.success("🎉 無任何重複佔位情況。")
