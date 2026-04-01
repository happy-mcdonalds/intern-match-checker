import streamlit as st
import pandas as pd
from datetime import datetime
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 莫蘭迪色系 + 強制純宋體 CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'PMingLiU', serif !important;
        background-color: #F2F0EB !important; color: #4A4F4D !important;
    }
    h1, h2, h3 { color: #3B403E !important; border-bottom: 1px solid #D1D1C9; padding-bottom: 5px; }
    section[data-testid="stSidebar"] { background-color: #E8E6E1 !important; border-right: 1px solid #D1D1C9 !important; }
    
    .stButton > button { 
        background-color: #7D8B84 !important; color: #FFFFFF !important; 
        border: none !important; border-radius: 4px !important; width: 100%; transition: 0.3s;
    }
    .stButton > button:hover { background-color: #65736C !important; }

    .stTable { font-size: 14px; background-color: #FFFFFF; border-radius: 5px; }
    th { background-color: #E4E5E0 !important; color: #4A4F4D !important; text-align: left !important; }
    td { white-space: pre-wrap !important; vertical-align: top !important; line-height: 1.6 !important; border-bottom: 1px solid #F0F0F0 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---
def extract_dates_ultra(text, year=2026):
    """最寬鬆日期解析：直接抓取數字組，不理會中間是什麼符號"""
    if isinstance(text, datetime): return text, text
    # 移除所有空白與換行，將非數字符號簡化為逗號
    clean_text = re.sub(r'[^\d]+', ',', str(text)).strip(',')
    nums = clean_text.split(',')
    
    try:
        # 如果有四位數開頭 (年.月.日)
        if len(nums) >= 3 and len(nums[0]) == 4:
            d1 = datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            if len(nums) >= 6:
                d2 = datetime(int(nums[3]), int(nums[4]), int(nums[5]))
            else:
                d2 = d1
            return d1, d2
        # 如果是 (月.日) 格式
        elif len(nums) >= 2:
            d1 = datetime(year, int(nums[0]), int(nums[1]))
            if len(nums) >= 4:
                d2 = datetime(year, int(nums[2]), int(nums[3]))
            else:
                d2 = d1
            return d1, d2
    except: pass
    return None, None

def smart_read_sheet(file):
    try:
        xls = pd.ExcelFile(file)
        target = next((sn for sn in xls.sheet_names if any(k in sn for k in ["志願", "名單", "容額"])), xls.sheet_names[0])
        df_temp = pd.read_excel(file, sheet_name=target)
        for i in range(min(len(df_temp), 15)):
            row = [str(x).strip() for x in df_temp.iloc[i].values]
            if any(k in row for k in ["姓名", "科別", "申請科別"]):
                return pd.read_excel(file, sheet_name=target, header=i+1)
        return df_temp
    except: return None

# --- 側邊欄與設定 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
if st.sidebar.button("系統重新整理"):
    st.rerun()

# --- 醫院代表模式 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 規則設定區 (直接顯示，不使用 Form 以免阻擋更新)
    st.markdown("### 1. 規則設定")
    c_cfg1, c_cfg2, c_cfg3 = st.columns(3)
    c_dur = c_cfg1.number_input("一個 Course 多久 (週)", 1, 10, 2)
    m_weeks = c_cfg2.number_input("最短實習週數要求 (週)", 1, 52, 4)
    r_cont = c_cfg3.checkbox("要求必須連續實習", True)
    
    if st.button("儲存並套用設定"):
        st.success("規則已儲存")

    st.divider()
    
    st.markdown("### 2. 檔案上傳")
    c1, c2 = st.columns(2)
    q_file = c1.file_uploader("上傳醫院容額表 (須包含「容額」字樣)", type=['xlsx'])
    a_file = c2.file_uploader("上傳學生志願表 (須包含「姓名」)", type=['xlsx'])
    
    if st.button("確認並開始比對"):
        if q_file and a_file:
            df_q = smart_read_sheet(q_file)
            df_a = smart_read_sheet(a_file)
            
            if df_q is not None and df_a is not None:
                df_q.columns = [str(c).strip() for c in df_q.columns]
                df_a.columns = [str(c).strip() for c in df_a.columns]
                df_a['姓名'] = df_a['姓名'].ffill()
                dept_col = "申請科別" if "申請科別" in df_a.columns else "科別"
                
                # 解析學生名單
                apps = []
                for _, r in df_a.iterrows():
                    if pd.notna(r.get(dept_col)) and pd.notna(r.get('實習期間')):
                        s, e = extract_dates_ultra(r['實習期間'])
                        if s: apps.append({'姓名': r['姓名'], '科別': str(r[dept_col]).strip(), '開始': s, '結束': e})
                
                # 比對容額
                date_cols = [c for c in df_q.columns if extract_dates_ultra(c)[0]]
                collisions = []
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get('科別', '')).strip()
                    if not dept or dept == 'nan': continue
                    for col in date_cols:
                        try:
                            # 清除容額數字裡的雜質
                            cap_text = re.sub(r'[^\d]+', '', str(q_row[col]))
                            cap = int(cap_text) if cap_text else 0
                        except: continue
                        
                        s_slot, e_slot = extract_dates_ultra(col)
                        # 核心比對邏輯
                        st_in = [a['姓名'] for a in apps if a['科別'] == dept and a['開始'] <= e_slot and a['結束'] >= s_slot]
                        
                        if len(st_in) > cap:
                            collisions.append({
                                "科別": dept, 
                                "時間": str(col).replace('\n',' '), 
                                "容額": cap, 
                                "超額學生": "、".join(list(set(st_in)))
                            })

                # 規章檢查
                invalid = []
                df_temp = pd.DataFrame(apps)
                if not df_temp.empty:
                    for name, gp in df_temp.groupby('姓名'):
                        gp = gp.sort_values('開始')
                        # 計算總天數 (簡單日期差)
                        total_days = sum([(r['結束'] - r['開始']).days + 1 for _, r in gp.iterrows()])
                        if total_days < m_weeks * 5:
                            invalid.append({"姓名": name, "原因": f"總時長不足 ({total_days} 天)"})
                        
                        if r_cont and len(gp) > 1:
                            for i in range(len(gp)-1):
                                if (gp.iloc[i+1]['開始'] - gp.iloc[i]['結束']).days > 3:
                                    invalid.append({"姓名": name, "原因": "實習時段不連續"})

                st.header("分析結果")
                if collisions:
                    st.subheader("名額撞期名單")
                    st.table(pd.DataFrame(collisions))
                if invalid:
                    st.subheader("規章不符名單")
                    st.table(pd.DataFrame(invalid).drop_duplicates())
                if not collisions and not invalid:
                    st.success("核對完成，目前一切正常。")
        else:
            st.warning("請確保兩個檔案都已上傳。")

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
                hit_idx = set()
                for i in range(len(recs)):
                    for j in range(i+1, len(recs)):
                        s1, e1 = extract_dates_ultra(recs[i]['實習期間'])
                        s2, e2 = extract_dates_ultra(recs[j]['實習期間'])
                        if s1 and s2 and (s1 <= e2 and s2 <= e1):
                            hit_idx.update([i, j])
                if hit_idx:
                    details = "\n".join([f"• {recs[idx]['來源醫院']} ({str(recs[idx]['實習期間']).strip()})" for idx in sorted(list(hit_idx))])
                    conflicts.append({"姓名": name, "衝突詳情": details})
            
            if conflicts:
                st.subheader("偵測到重疊佔位")
                st.table(pd.DataFrame(conflicts).set_index('姓名'))
            else:
                st.success("恭喜！無跨院重疊佔位。")
