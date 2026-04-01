import streamlit as st
import pandas as pd
from datetime import datetime
import re

# --- 系統記憶 ---
if "course_dur_weeks" not in st.session_state: st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state: st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state: st.session_state.require_cont = True

# 頁面基本設定
st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# --- 莫蘭迪色系 + 強制純宋體 CSS (無 Emoji) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');
    html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
        font-family: 'Noto Serif TC', 'Songti TC', 'PMingLiU', serif !important;
        background-color: #F5F4F1 !important; color: #5C5E5D !important;
    }
    h1, h2, h3 { color: #4A4C4B !important; border-bottom: 1px solid #D6D4CE; padding-bottom: 5px; }
    section[data-testid="stSidebar"] { background-color: #EAE8E3 !important; border-right: 1px solid #D6D4CE !important; }
    
    /* 按鈕樣式 (鼠尾草綠) */
    .stButton > button { 
        background-color: #8A9A92 !important; color: #FFFFFF !important; 
        border: none !important; border-radius: 4px !important; width: 100%; transition: 0.3s;
    }
    .stButton > button:hover { background-color: #72827A !important; }
    .btn-secondary > button { background-color: #C0BFB8 !important; }

    /* 表格樣式 */
    .stTable { font-size: 14px; background-color: #FFFFFF; }
    th { background-color: #E3E1DB !important; color: #4A4C4B !important; text-align: left !important;}
    td { white-space: pre-wrap !important; vertical-align: top !important; line-height: 1.8 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 核心工具函式 ---
def safe_read_excel(file, target_keywords):
    """終極安全讀取：強制取消預設標題，從第 0 行開始掃描"""
    try:
        xls = pd.ExcelFile(file)
        target = next((sn for sn in xls.sheet_names if any(k in sn for k in ["志願", "名單", "工作表4", "實習容額", "時段"])), xls.sheet_names[0])
        
        # 強制 header=None，確保 Excel 第一行絕對不會被漏掉
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        
        for i in range(min(len(df_raw), 20)):
            row_vals = [str(x).strip() for x in df_raw.iloc[i].values]
            # 只要該列包含我們設定的關鍵字 (例如：姓名、科別)，就認定為標題列
            if any(kw in val for kw in target_keywords for val in row_vals):
                df = pd.read_excel(file, sheet_name=target, header=i)
                df.columns = [str(c).strip() for c in df.columns]
                return df
                
        # 萬一都沒找到，直接以預設方式讀取
        df_fallback = pd.read_excel(file, sheet_name=target)
        df_fallback.columns = [str(c).strip() for c in df_fallback.columns]
        return df_fallback
    except Exception as e:
        return None

def extract_dates_universal(text, year=2026):
    """加強版日期解析：處理換行符號與各種怪異間隔符"""
    if isinstance(text, datetime): return text, text
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
st.sidebar.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
if st.sidebar.button("重新整理系統"):
    st.rerun()
st.sidebar.markdown('</div>', unsafe_allow_html=True)

# --- 醫院代表模式 ---
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")
    
    # 規則設定區
    with st.expander("規則設定 (點擊展開修改)", expanded=False):
        c_cfg1, c_cfg2, c_cfg3 = st.columns(3)
        st.session_state.course_dur_weeks = c_cfg1.number_input("一個 Course 多久 (週)", 1, 10, st.session_state.course_dur_weeks)
        st.session_state.min_weeks_req = c_cfg2.number_input("最短實習週數要求 (週)", 1, 52, st.session_state.min_weeks_req)
        st.session_state.require_cont = c_cfg3.checkbox("要求必須連續實習", st.session_state.require_cont)
        if st.button("儲存規則條件"): 
            st.success("條件已更新")

    st.divider()
    
    # 檔案上傳區
    c1, c2 = st.columns(2)
    q_file = c1.file_uploader("上傳醫院容額表", type=['xlsx'])
    a_file = c2.file_uploader("上傳學生志願表", type=['xlsx'])
    
    # 單向觸發：按下後直接執行並顯示結果，避免畫面跳轉
    if st.button("確認並開始比對"):
        if q_file and a_file:
            # 針對容額表與志願表設定專屬關鍵字，確保精準抓取 Header
            df_q = safe_read_excel(q_file, ["科別", "容額"])
            df_a = safe_read_excel(a_file, ["姓名", "申請科別", "科別"])
            
            if df_q is not None and df_a is not None:
                # 防呆：確認是否有姓名欄位
                if '姓名' not in df_a.columns:
                    st.error("學生志願表中找不到「姓名」欄位，請檢查表頭。")
                    st.stop()
                    
                df_a['姓名'] = df_a['姓名'].ffill()
                dept_col = next((c for c in df_a.columns if "科別" in c), None)
                time_col = next((c for c in df_a.columns if "期間" in c or "時間" in c), None)
                
                if not dept_col or not time_col:
                    st.error("學生志願表中缺少「科別」或「實習期間」欄位。")
                    st.stop()

                apps = []
                for _, r in df_a.iterrows():
                    if pd.notna(r.get(dept_col)) and pd.notna(r.get(time_col)):
                        s, e, d = parse_period_dates(r[time_col])
                        if s: apps.append({'姓名': r['姓名'], '科別': str(r[dept_col]).strip(), '開始': s, '結束': e, '天數': d})
                
                # 容額比對
                date_cols = [c for c in df_q.columns if extract_dates_universal(c)[0]]
                q_dept_col = next((c for c in df_q.columns if "科別" in c), df_q.columns[0])
                collisions = []
                
                for _, q_row in df_q.iterrows():
                    dept = str(q_row.get(q_dept_col, '')).strip()
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
                if not df_temp.empty:
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
                if not collisions and not invalid: 
                    st.success("核對完成，查無異常。")
            else:
                st.error("檔案解析失敗，請確認檔案格式是否正確。")
        else:
            st.warning("請先上傳兩個檔案後再點擊比對。")

# --- 系秘模式 ---
elif mode == "系秘":
    st.title("跨院重複佔位檢查")
    multi_files = st.file_uploader("上傳各院志願清單 (可多選)", type=['xlsx'], accept_multiple_files=True)
    
    if st.button("執行跨院比對"):
        if multi_files:
            all_data = []
            for f in multi_files:
                df = safe_read_excel(f, ["姓名", "申請科別", "科別"])
                if df is not None and '姓名' in df.columns:
                    df['姓名'] = df['姓名'].ffill()
                    df['來源醫院'] = f.name.split('.')[0]
                    time_col = next((c for c in df.columns if "期間" in c or "時間" in c), None)
                    if time_col:
                        df['_time_val'] = df[time_col]
                        all_data.append(df[df['_time_val'].notna()])
            
            if all_data:
                full = pd.concat(all_data, ignore_index=True)
                conflicts = []
                for name, gp in full.groupby('姓名'):
                    recs = gp.to_dict('records')
                    hit = set()
                    for i in range(len(recs)):
                        for j in range(i+1, len(recs)):
                            s1, e1, _ = parse_period_dates(recs[i]['_time_val'])
                            s2, e2, _ = parse_period_dates(recs[j]['_time_val'])
                            if s1 and s2 and (s1 <= e2 and s2 <= e1): hit.update([i, j])
                    if hit:
                        details = "\n".join([f"• {recs[idx]['來源醫院']} ({str(recs[idx]['_time_val']).strip()})" for idx in sorted(list(hit))])
                        conflicts.append({"姓名": name, "衝突詳情": details})
                if conflicts:
                    st.subheader("偵測到重疊佔位")
                    st.table(pd.DataFrame(conflicts).set_index('姓名'))
                else: 
                    st.success("無跨院重疊佔位。")
        else:
            st.warning("請先上傳檔案。")
