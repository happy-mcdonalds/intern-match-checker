import streamlit as st
import pandas as pd
from datetime import datetime
import re

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 高級感 CSS (宋體 + 黑白灰 + 無 Emoji) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Serif TC', 'Songti TC', 'Source Han Serif TC', serif !important;
        color: #000000;
    }
    
    /* 標題與分隔線 */
    h1, h2, h3 {
        color: #000000 !important;
        font-weight: 700 !important;
        border-bottom: 1px solid #000000;
        padding-bottom: 5px;
        margin-top: 20px;
    }

    /* 側邊欄極簡化 */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #DDDDDD;
    }
    
    /* 按鈕黑白化 */
    .stButton>button {
        color: #FFFFFF;
        background-color: #000000;
        border-radius: 0px;
        border: 1px solid #000000;
        font-size: 14px;
        width: 100%;
    }
    
    /* 表格樣式優化 */
    .stTable {
        border: 1px solid #000000;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 工具函式 ---
def parse_date_simple(s, year=2026):
    try:
        # 處理 M/D 格式
        parts = re.findall(r'\d+', str(s))
        if len(parts) >= 2:
            return datetime(year, int(parts[0]), int(parts[1]))
    except: pass
    return None

def extract_apply_data(df):
    """ 模擬 Excel OFFSET 與 REGEX 邏輯提取學生名單 """
    records = []
    # 填充合併單元格的姓名 (ffill)
    df['姓名'] = df['姓名'].ffill()
    for _, row in df.iterrows():
        if pd.notna(row['申請科別']) and pd.notna(row['實習期間']):
            dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(row['實習期間']).replace('\n',''))
            if len(dates) >= 2:
                start = datetime.strptime(dates[0], "%Y.%m.%d")
                end = datetime.strptime(dates[1], "%Y.%m.%d")
                records.append({
                    '姓名': row['姓名'],
                    '科別': str(row['申請科別']).strip(),
                    '開始': start,
                    '結束': end,
                    '週數': round(((end - start).days + 1) / 7, 1)
                })
    return records

# --- 側邊欄 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表模式 (容額校對)", "總代模式 (跨院比對)"])

if mode == "醫院代表模式 (容額校對)":
    st.title("醫院內部容額與規章審核")
    
    with st.sidebar.expander("規則設定", expanded=True):
        course_duration = st.number_input("一個 Course 多久 (週)", min_value=1, value=2)
        min_weeks_req = st.number_input("最短實習週數要求", min_value=1, value=4)

    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("上傳醫院容額表", type=['xlsx'])
    with c2: a_file = st.file_uploader("上傳學生志願表", type=['xlsx'])

    if q_file and a_file:
        try:
            # 讀取容額 (自動定位標題列)
            xls_q = pd.ExcelFile(q_file)
            sn_q = [s for s in xls_q.sheet_names if "容額" in s][0]
            df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]

            # 讀取申請 (鎖定志願申請名單分頁)
            df_a = pd.read_excel(a_file, sheet_name="志願申請名單")
            df_a.columns = [str(c).strip() for c in df_a.columns]
            
            # 提取學生志願
            student_apps = extract_apply_data(df_a)
            
            # 執行衝突偵測
            date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
            collisions = []
            
            for _, q_row in df_q.iterrows():
                dept = q_row.get('科別')
                if pd.isna(dept): continue
                dept_name = str(dept).strip()
                
                for col in date_cols:
                    cap = q_row.get(col)
                    if pd.isna(cap) or not str(cap).isdigit(): cap_val = 0
                    else: cap_val = int(float(cap))
                    
                    # 解析時段日期 (例如 5/4-5/8)
                    pts = col.split('-')
                    s_slot = parse_date_simple(pts[0])
                    e_slot = parse_date_simple(pts[1]) if len(pts) > 1 else s_slot
                    
                    if not s_slot or not e_slot: continue
                    
                    # 篩選佔用此時段的學生
                    st_in_slot = [a['姓名'] for a in student_apps if a['科別'] == dept_name and a['開始'] <= e_slot and s_slot <= a['結束']]
                    
                    if len(st_in_slot) > cap_val:
                        collisions.append({
                            "衝突科別": dept_name,
                            "衝突時間": col,
                            "容額限制": cap_val,
                            "超額名單": "、".join(st_in_slot)
                        })

            # 顯示結果
            st.header("異常監控結果")
            
            if collisions:
                st.subheader("名額撞期名單")
                st.table(pd.DataFrame(collisions))
            else:
                st.success("目前名單無名額爆掉情況。")

            # 週數檢查
            short_stay = []
            for app in student_apps:
                if app['週數'] < min_weeks_req:
                    short_stay.append({
                        "姓名": app['姓名'], "申請科別": app['科別'], 
                        "實際週數": app['週數'], "狀態": f"低於要求 {min_weeks_req} 週"
                    })
            
            if short_stay:
                st.subheader("週數不符要求名單")
                st.table(pd.DataFrame(short_stay))

        except Exception as e:
            st.error(f"解析失敗：{e}")

elif mode == "總代模式 (跨院比對)":
    st.title("跨院重複佔位檢查")
    # (保留原本的跨院比對邏輯，並同步套用宋體與無 Emoji 樣式)
