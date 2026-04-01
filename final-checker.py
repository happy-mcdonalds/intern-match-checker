import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="醫學系跨院重複佔位檢查器", layout="wide")
st.title("🛡️ 醫學系自選實習：跨院重複佔位檢查系統")
st.info("請各醫院代表上傳各自的『志願申請名單』Excel 或 CSV 檔")

# --- 1. 日期工具 (支援連字號格式) ---
def parse_date(d):
    try:
        clean_d = str(d).replace('/', '.').replace('\n', '').strip()
        if len(clean_d.split('.')) == 2: clean_d = "2026." + clean_d
        return datetime.strptime(clean_d, "%Y.%m.%d")
    except: return None

def is_overlap(range1, range2):
    try:
        r1_s, r1_e = [parse_date(x) for x in str(range1).split('-')]
        r2_s, r2_e = [parse_date(x) for x in str(range2).split('-')]
        if None in [r1_s, r1_e, r2_s, r2_e]: return False
        return r1_s <= r2_e and r2_s <= r1_e
    except: return False

# --- 2. 檔案上傳區塊 ---
uploaded_files = st.file_uploader("上傳各院名單 (可一次選多個檔案)", type=['xlsx', 'csv'], accept_multiple_files=True)

all_data = []

if uploaded_files:
    for file in uploaded_files:
        try:
            # 讀取檔案 (自動辨識 Excel 或 CSV)
            if file.name.endswith('.xlsx'):
                df = pd.read_excel(file)
            else:
                df = pd.read_csv(file)
            
            # 清理標題
            df.columns = [str(c).strip() for c in df.columns]
            
            # 標註這份資料是哪家醫院的 (從檔名判斷)
            df['來源檔案'] = file.name
            all_data.append(df)
            st.success(f"成功讀取: {file.name}")
        except Exception as e:
            st.error(f"檔案 {file.name} 格式不符: {e}")

# --- 3. 執行交叉比對 ---
if all_data:
    full_df = pd.concat(all_data, ignore_index=True)
    
    # 強制確保必要欄位存在
    required_cols = ['學號', '姓名', '申請科別', '實習期間']
    missing = [c for c in required_cols if c not in full_df.columns]
    
    if missing:
        st.error(f"檔案中缺少必要欄位: {missing}")
    else:
        st.header("🚩 重複佔位偵測結果")
        
        conflicts = []
        # 以學號為基準尋找
        for s_id in full_df['學號'].unique():
            s_apps = full_df[full_df['學號'] == s_id]
            
            if len(s_apps) > 1:
                # 兩兩比對時間
                records = s_apps.to_dict('records')
                for i in range(len(records)):
                    for j in range(i + 1, len(records)):
                        if is_overlap(records[i]['實習期間'], records[j]['實習期間']):
                            conflicts.append({
                                "姓名": records[i]['姓名'],
                                "學號": s_id,
                                "醫院A": records[i]['來源檔案'],
                                "科別A": records[i]['申請科別'],
                                "時間A": records[i]['實習期間'],
                                "醫院B": records[j]['來源檔案'],
                                "科別B": records[j]['申請科別'],
                                "時間B": records[j]['實習期間'],
                            })

        if conflicts:
            conflicts_df = pd.DataFrame(conflicts)
            st.warning(f"🚨 警告！偵測到 {len(conflicts)} 筆重複佔位衝突：")
            st.dataframe(conflicts_df, use_container_width=True)
            
            # 匯出黑名單
            csv = conflicts_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載衝突名單 (協調用)", csv, "conflicts.csv", "text/csv")
        else:
            st.balloons()
            st.success("✨ 太棒了！所有醫院名單交叉比對後，無人重複佔位。")

        with st.expander("🔍 查看合併後的總表"):
            st.dataframe(full_df)

