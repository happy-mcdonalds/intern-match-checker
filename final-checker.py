import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
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
    /* 表格字體縮小以符合專業感 */
    .stTable { font-size: 14px; }
    </style>
    """, unsafe_allow_html=True)

# --- 工具函式 ---
def parse_date_simple(s, year=2026):
    try:
        parts = re.findall(r'\d+', str(s))
        if len(parts) >= 2:
            return datetime(year, int(parts[0]), int(parts[1]))
    except: pass
    return None

def count_workdays(start, end):
    """計算實習天數（含頭尾，但不精確扣除國定假日，僅供邏輯判定）"""
    if not start or not end: return 0
    return (end - start).days + 1

# --- 側邊欄模式切換 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表 (容額校對)", "系秘 (跨院比對)"])
st.sidebar.markdown("---")

if mode == "醫院代表 (容額校對)":
    st.title("醫院內部容額與規章審核")
    
    # 規則設定區 (確保勾選框存在)
    with st.sidebar.expander("規則設定", expanded=True):
        course_duration_weeks = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
        min_weeks_req = st.number_input("最短實習週數要求 (週)", min_value=1, value=4)
        require_cont = st.checkbox("要求必須連續實習", value=True)

    # 換算為天數 (以一週 5 天上課日計算，2 週 = 10 天，4 週 = 20 天)
    course_days = course_duration_weeks * 5
    total_min_days = min_weeks_req * 5

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("上傳醫院容額表 (實習容額與時段)", type=['xlsx'])
    with c2: a_file = st.file_uploader("上傳學生志願表 (志願申請名單)", type=['xlsx'])

    if q_file and a_file:
        try:
            # 1. 讀取容額表 (自動定位標題列)
            xls_q = pd.ExcelFile(q_file)
            sn_q = [s for s in xls_q.sheet_names if "容額" in s or "時段" in s][0]
            df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]

            # 2. 讀取申請表
            df_a = pd.read_excel(a_file, sheet_name="志願申請名單")
            df_a.columns = [str(c).strip() for c in df_a.columns]
            df_a['姓名'] = df_a['姓名'].ffill() # 處理 OFFSET 邏輯
            
            # 解析志願 (模擬工作表4)
            apps = []
            for _, row in df_a.iterrows():
                if pd.notna(row['申請科別']) and pd.notna(row['實習期間']):
                    dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(row['實習期間']))
                    if len(dates) >= 2:
                        s = datetime.strptime(dates[0], "%Y.%m.%d")
                        e = datetime.strptime(dates[1], "%Y.%m.%d")
                        apps.append({
                            '姓名': row['姓名'],
                            '科別': str(row['申請科別']).strip(),
                            '開始': s, '結束': e,
                            '天數': count_workdays(s, e)
                        })
            
            # 3. 執行碰撞偵測 (名額爆掉)
            date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
            collisions = []
            for _, q_row in df_q.iterrows():
                dept = q_row.get('科別')
                if pd.isna(dept): continue
                
                for col in date_cols:
                    cap = q_row.get(col)
                    if pd.isna(cap) or not str(cap).isdigit(): continue
                    cap_val = int(float(cap))
                    
                    # 判定該週日期
                    pts = col.split('-')
                    s_slot = parse_date_simple(pts[0])
                    e_slot = parse_date_simple(pts[1]) if len(pts) > 1 else s_slot
                    
                    # 找出佔位的學生
                    st_in_slot = [a['姓名'] for a in apps if a['科別'] == str(dept).strip() and a['開始'] <= e_slot and s_slot <= a['結束']]
                    
                    if len(st_in_slot) > cap_val:
                        collisions.append({
                            "科別": dept, "時間": col, "容額": cap_val, "超額學生": "、".join(st_in_slot)
                        })

            # 4. 執行資格審核 (Course 天數與總週數)
            invalid_students = []
            # 以姓名分組檢查總週數
            df_temp = pd.DataFrame(apps)
            for name, group in df_temp.groupby('姓名'):
                total_days = group['天數'].sum()
                
                # 檢查單一 Course 是否達標
                for _, row in group.iterrows():
                    if row['天數'] < course_days:
                        invalid_students.append({"姓名": name, "原因": f"{row['科別']} 實習天數不足({row['天數']}天)，未達 Course 要求 {course_days} 天"})
                
                # 檢查總週數是否達標
                if total_days < total_min_days:
                    invalid_students.append({"姓名": name, "原因": f"總實習天數不足({total_days}天)，未達最低要求 {total_min_days} 天"})

            # --- 顯示結果 ---
            st.header("異常監控結果")
            
            if collisions:
                st.subheader("名額撞期名單 (超額佔位)")
                st.table(pd.DataFrame(collisions))
            else:
                st.success("名額分配正常。")

            if invalid_students:
                st.subheader("規章不符名單 (天數不足)")
                st.table(pd.DataFrame(invalid_students).drop_duplicates())

        except Exception as e:
            st.error(f"解析失敗：{e}")

elif mode == "系秘 (跨院比對)":
    st.title("跨院重複佔位檢查")
    # (保留原本的跨院比對邏輯...)
