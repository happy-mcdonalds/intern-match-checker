import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import io

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

def is_overlap(range1, range2):
    try:
        d1 = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(range1))
        d2 = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(range2))
        r1_s = datetime.strptime(d1[0], "%Y.%m.%d")
        r1_e = datetime.strptime(d1[1], "%Y.%m.%d")
        r2_s = datetime.strptime(d2[0], "%Y.%m.%d")
        r2_e = datetime.strptime(d2[1], "%Y.%m.%d")
        return r1_s <= r2_e and r2_s <= r1_e
    except: return False

# --- 側邊欄控制 ---
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表 (容額校對)", "系秘 (跨院比對)"])

st.sidebar.divider()

# 即使在系秘也顯示規則設定，以便進行全域判定
st.sidebar.subheader("規則設定")
course_duration_weeks = st.sidebar.number_input("一個 Course 多久 (週)", min_value=1, value=2)
min_weeks_req = st.sidebar.number_input("最短實習週數要求 (週)", min_value=1, value=4)
require_cont = st.sidebar.checkbox("要求必須連續實習", value=True)

# 換算天數 (1週以5天計)
course_days = course_duration_weeks * 5
total_min_days = min_weeks_req * 5

# --- 模式一：醫院代表 ---
if mode == "醫院代表 (容額校對)":
    st.title("醫院內部容額與規章審核")
    
    c1, c2 = st.columns(2)
    with c1: q_file = st.file_uploader("1. 上傳醫院容額表 (實習容額與時段)", type=['xlsx'])
    with c2: a_file = st.file_uploader("2. 上傳學生志願表 (志願申請名單)", type=['xlsx'])

    if q_file and a_file:
        try:
            # 讀取容額表
            xls_q = pd.ExcelFile(q_file)
            sn_q = [s for s in xls_q.sheet_names if "容額" in s or "時段" in s][0]
            df_q = pd.read_excel(q_file, sheet_name=sn_q, header=4)
            df_q.columns = [str(c).strip() for c in df_q.columns]

            # 讀取申請表
            df_a = pd.read_excel(a_file, sheet_name="志願申請名單")
            df_a.columns = [str(c).strip() for c in df_a.columns]
            df_a['姓名'] = df_a['姓名'].ffill()
            
            apps = []
            for _, row in df_a.iterrows():
                if pd.notna(row['申請科別']) and pd.notna(row['實習期間']):
                    dates = re.findall(r'\d{4}\.\d{2}\.\d{2}', str(row['實習期間']))
                    if len(dates) >= 2:
                        s = datetime.strptime(dates[0], "%Y.%m.%d")
                        e = datetime.strptime(dates[1], "%Y.%m.%d")
                        apps.append({'姓名': row['姓名'], '科別': str(row['申請科別']).strip(), '開始': s, '結束': e, '天數': (e-s).days + 1})

            # 容額檢查 (撞期)
            date_cols = [c for c in df_q.columns if '-' in c and any(i.isdigit() for i in c)]
            collisions = []
            for _, q_row in df_q.iterrows():
                dept = q_row.get('科別')
                if pd.isna(dept): continue
                for col in date_cols:
                    cap = q_row.get(col)
                    if pd.isna(cap) or not str(cap).isdigit(): continue
                    cap_val = int(float(cap))
                    pts = col.split('-')
                    s_slot = parse_date_simple(pts[0]); e_slot = parse_date_simple(pts[1]) if len(pts)>1 else s_slot
                    st_in_slot = [a['姓名'] for a in apps if a['科別'] == str(dept).strip() and a['開始'] <= e_slot and s_slot <= a['結束']]
                    if len(st_in_slot) > cap_val:
                        collisions.append({"科別": dept, "時間": col, "容額": cap_val, "超額學生": "、".join(st_in_slot)})

            # 顯示結果
            st.header("異常監控結果")
            if collisions:
                st.subheader("名額撞期名單")
                st.table(pd.DataFrame(collisions))
            else:
                st.success("名額分配正常。")

            # 資格檢查
            invalid = []
            df_temp = pd.DataFrame(apps)
            for name, group in df_temp.groupby('姓名'):
                total_d = group['天數'].sum()
                for _, row in group.iterrows():
                    if row['天數'] < course_days:
                        invalid.append({"姓名": name, "原因": f"{row['科別']} 未達 Course 要求 {course_days} 天"})
                if total_d < total_min_days:
                    invalid.append({"姓名": name, "原因": f"總實習天數 {total_d} 天未達要求 {total_min_days} 天"})
            
            if invalid:
                st.subheader("規章不符名單")
                st.table(pd.DataFrame(invalid).drop_duplicates())
        except Exception as e:
            st.error(f"解析失敗：{e}")

# --- 模式二：系秘 (修正後：顯示上傳欄位) ---
elif mode == "系秘 (跨院比對)":
    st.title("跨院重複佔位檢查")
    st.markdown("請同時選取並上傳多個醫院的「志願申請名單」檔案。")
    
    # 這裡就是你說沒出現的欄位
    multi_files = st.file_uploader("上傳各院確定名單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    
    if multi_files:
        all_data = []
        for f in multi_files:
            try:
                # 每個檔案讀取其「志願申請名單」分頁
                df_raw = pd.read_excel(f, sheet_name="志願申請名單")
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                df_raw['姓名'] = df_raw['姓名'].ffill()
                df_raw['來源醫院'] = f.name
                all_data.append(df_raw[df_raw['申請科別'].notna()])
            except:
                st.error(f"檔案 {f.name} 讀取失敗，請確認分頁名稱為『志願申請名單』")
        
        if all_data:
            full_df = pd.concat(all_data, ignore_index=True)
            conflicts = []
            # 依姓名/學號檢查跨醫院日期重疊
            for s_name in full_df['姓名'].unique():
                s_apps = full_df[full_df['姓名'] == s_name].to_dict('records')
                if len(s_apps) > 1:
                    for i in range(len(s_apps)):
                        for j in range(i + 1, len(s_apps)):
                            if is_overlap(s_apps[i]['實習期間'], s_apps[j]['實習期間']):
                                conflicts.append({
                                    "姓名": s_name,
                                    "醫院A": s_apps[i]['來源醫院'], "時間A": s_apps[i]['實習期間'],
                                    "醫院B": s_apps[j]['來源醫院'], "時間B": s_apps[j]['實習期間']
                                })
            
            st.header("跨院衝突偵測結果")
            if conflicts:
                st.table(pd.DataFrame(conflicts).drop_duplicates())
            else:
                st.success("交叉比對完成，無重複佔位情況。")
