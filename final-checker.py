import streamlit as st
import pandas as pd
from datetime import datetime
import re
import hashlib

# =========================
# 初始化系統設定
# =========================
if "course_dur_weeks" not in st.session_state:
    st.session_state.course_dur_weeks = 2
if "min_weeks_req" not in st.session_state:
    st.session_state.min_weeks_req = 4
if "require_cont" not in st.session_state:
    st.session_state.require_cont = True

st.set_page_config(page_title="醫學系實習選配管理系統", layout="wide")

# =========================
# CSS
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap');

html, body, [class*="css"], [data-testid="stAppViewContainer"], .stApp {
    font-family: 'Noto Serif TC', 'Songti TC', 'PMingLiU', 'MingLiU', 'SimSun', serif !important;
    background-color: #F5F4F1 !important;
    color: #5C5E5D !important;
}

h1, h2, h3 {
    color: #4A4C4B !important;
    border-bottom: 1px solid #D6D4CE;
    padding-bottom: 5px;
    font-weight: 700;
}

section[data-testid="stSidebar"] {
    background-color: #EAE8E3 !important;
    border-right: 1px solid #D6D4CE !important;
}

[data-testid="stForm"] {
    border: 1px solid #D6D4CE !important;
    background-color: #FDFDFD !important;
    border-radius: 4px;
    padding: 20px;
}

.stButton > button, [data-testid="stFormSubmitButton"] > button {
    background-color: #8A9A92 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 4px !important;
    width: 100%;
    transition: 0.3s;
}

.stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
    background-color: #72827A !important;
}

.btn-secondary > button {
    background-color: #C0BFB8 !important;
    color: #FFFFFF !important;
}

.btn-secondary > button:hover {
    background-color: #A8A7A0 !important;
}

th {
    background-color: #E3E1DB !important;
    color: #4A4C4B !important;
    border-bottom: 2px solid #C0BFB8 !important;
}

td {
    border-bottom: 1px solid #EAE8E3 !important;
}

th, td {
    white-space: pre-wrap !important;
    vertical-align: top !important;
    line-height: 1.8 !important;
    text-align: left !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# 工具函式
# =========================
def reset_pointer(file_obj):
    try:
        file_obj.seek(0)
    except:
        pass

def file_md5(file_obj):
    try:
        data = file_obj.getvalue()
        return hashlib.md5(data).hexdigest()
    except:
        return None

def business_days(start, end):
    try:
        if pd.isna(start) or pd.isna(end):
            return 0
        return len(pd.bdate_range(start, end))
    except:
        return 0

def extract_dates_universal(text, default_year=2026):
    """通用日期解析：支援 datetime、Timestamp、'5/4-5/15'、'2026/5/4~2026/5/15'、含換行格式"""
    if isinstance(text, (datetime, pd.Timestamp)):
        dt = pd.to_datetime(text)
        return dt, dt

    if pd.isna(text):
        return None, None

    text = str(text).strip()
    if not text or text.lower() == "nan":
        return None, None

    text = (
        text.replace("\n", "-")
            .replace("\r", "-")
            .replace("～", "-")
            .replace("~", "-")
            .replace("至", "-")
            .replace("到", "-")
            .replace("_", "-")
            .replace(" ", "")
    )

    parts = re.split(r"[-]+", text)

    def parse_one(part):
        nums = re.findall(r"\d+", part)
        try:
            if len(nums) >= 3 and len(nums[0]) == 4:
                return datetime(int(nums[0]), int(nums[1]), int(nums[2]))
            elif len(nums) >= 2:
                return datetime(default_year, int(nums[-2]), int(nums[-1]))
            return None
        except:
            return None

    dates = [parse_one(p) for p in parts if parse_one(p) is not None]

    if len(dates) == 1:
        return dates[0], dates[0]
    elif len(dates) >= 2:
        return dates[0], dates[-1]
    return None, None

def parse_period_dates(period_str):
    try:
        s, e = extract_dates_universal(period_str)
        if s is not None and e is not None:
            return s, e, business_days(s, e)
    except:
        pass
    return None, None, 0

def periods_overlap(s1, e1, s2, e2):
    if pd.isna(s1) or pd.isna(e1) or pd.isna(s2) or pd.isna(e2):
        return False
    return s1 <= e2 and s2 <= e1

def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).replace("\n", "").replace("\r", "").strip() for c in df.columns]
    return df

def pick_sheet_by_keywords(sheet_names, keyword_groups):
    for keywords in keyword_groups:
        for sn in sheet_names:
            if any(k in str(sn) for k in keywords):
                return sn
    return sheet_names[0]

def smart_read_application_sheet(file):
    """讀學生名單 / 志願表 / 確定實習名單"""
    try:
        reset_pointer(file)
        xls = pd.ExcelFile(file)

        sheet = pick_sheet_by_keywords(
            xls.sheet_names,
            [
                ["確定實習名單"],
                ["志願", "名單"],
                ["工作表4"],
            ]
        )

        reset_pointer(file)
        raw = pd.read_excel(file, sheet_name=sheet, header=None)

        header_idx = 0
        scan_rows = min(len(raw), 25)

        for i in range(scan_rows):
            row = [str(x).strip() for x in raw.iloc[i].values]

            has_name = any(k in row for k in ["姓名", "中文姓名"])
            has_dept = any(k in row for k in ["科別", "申請科別", "實習科別"])
            has_date = any(k in row for k in ["實習期間", "實習日期(開始)", "實習日期(結束)"])

            if has_name and has_dept and has_date:
                header_idx = i
                break

        reset_pointer(file)
        df = pd.read_excel(file, sheet_name=sheet, header=header_idx)
        df = clean_columns(df)
        return df

    except Exception as e:
        st.warning(f"讀取 {getattr(file, 'name', '檔案')} 失敗：{e}")
        return None

def smart_read_capacity_sheet(file):
    """讀容額表：尋找含科別 + 日期欄位的表"""
    try:
        reset_pointer(file)
        xls = pd.ExcelFile(file)

        sheet = pick_sheet_by_keywords(
            xls.sheet_names,
            [
                ["容額", "時段"],
                ["名單"],
            ]
        )

        reset_pointer(file)
        raw = pd.read_excel(file, sheet_name=sheet, header=None)

        header_idx = 0
        scan_rows = min(len(raw), 25)

        for i in range(scan_rows):
            row = [str(x).strip() for x in raw.iloc[i].values]
            has_dept = "科別" in row
            date_like_count = sum(1 for x in row if extract_dates_universal(x)[0] is not None)
            if has_dept and date_like_count >= 1:
                header_idx = i
                break

        reset_pointer(file)
        df = pd.read_excel(file, sheet_name=sheet, header=header_idx)
        df = clean_columns(df)
        df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
        return df

    except Exception as e:
        st.warning(f"讀取容額表 {getattr(file, 'name', '檔案')} 失敗：{e}")
        return None

def normalize_application_df(df, source_name):
    """把不同格式統一成：
       學號 / 姓名 / 科別 / 開始 / 結束 / 天數 / 來源醫院 / 比對鍵
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df = clean_columns(df)

    name_col = None
    for c in ["姓名", "中文姓名"]:
        if c in df.columns:
            name_col = c
            break

    dept_col = None
    for c in ["申請科別", "實習科別", "科別"]:
        if c in df.columns:
            dept_col = c
            break

    id_col = None
    for c in ["學號", "StudentID", "student_id"]:
        if c in df.columns:
            id_col = c
            break

    if name_col is None or dept_col is None:
        return pd.DataFrame()

    df[name_col] = df[name_col].ffill()

    if "實習日期(開始)" in df.columns and "實習日期(結束)" in df.columns:
        df["開始"] = pd.to_datetime(df["實習日期(開始)"], errors="coerce")
        df["結束"] = pd.to_datetime(df["實習日期(結束)"], errors="coerce")
    elif "實習期間" in df.columns:
        parsed = df["實習期間"].apply(parse_period_dates)
        df["開始"] = parsed.apply(lambda x: x[0])
        df["結束"] = parsed.apply(lambda x: x[1])
    else:
        return pd.DataFrame()

    out = pd.DataFrame({
        "學號": df[id_col].astype(str).str.strip() if id_col else "",
        "姓名": df[name_col].astype(str).str.strip(),
        "科別": df[dept_col].astype(str).str.strip(),
        "開始": df["開始"],
        "結束": df["結束"],
        "來源醫院": source_name
    })

    out = out.dropna(subset=["姓名", "開始", "結束"])
    out = out[out["姓名"] != ""]
    out["天數"] = out.apply(lambda r: business_days(r["開始"], r["結束"]), axis=1)
    out["比對鍵"] = out.apply(
        lambda r: r["學號"] if str(r["學號"]).strip() not in ["", "nan", "None"] else r["姓名"],
        axis=1
    )

    return out.reset_index(drop=True)

def extract_capacity_value(val):
    if pd.isna(val):
        return None
    nums = re.findall(r"\d+", str(val))
    if not nums:
        return None
    try:
        return int(nums[0])
    except:
        return None

def detect_same_file_uploads(files):
    md5_map = {}
    duplicates = []
    for f in files:
        h = file_md5(f)
        if h:
            if h in md5_map:
                duplicates.append((md5_map[h], f.name))
            else:
                md5_map[h] = f.name
    return duplicates

# =========================
# 側邊欄
# =========================
st.sidebar.title("系統模式")
mode = st.sidebar.radio("身份選擇", ["醫院代表", "系秘"])
st.sidebar.divider()

st.sidebar.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
if st.sidebar.button("重新整理系統"):
    st.rerun()
st.sidebar.markdown('</div>', unsafe_allow_html=True)

# =========================
# 醫院代表模式
# =========================
if mode == "醫院代表":
    st.title("醫院內部容額與規章審核")

    with st.form("settings_form"):
        st.markdown("### 規則設定")
        c1, c2, c3 = st.columns(3)
        with c1:
            cd_val = st.number_input(
                "一個 Course 多久（週）",
                min_value=1,
                value=st.session_state.course_dur_weeks
            )
        with c2:
            mw_val = st.number_input(
                "最短實習週數要求（週）",
                min_value=1,
                value=st.session_state.min_weeks_req
            )
        with c3:
            rc_val = st.checkbox(
                "要求必須連續實習",
                value=st.session_state.require_cont
            )

        st.markdown('<div class="btn-secondary">', unsafe_allow_html=True)
        save_btn = st.form_submit_button("儲存條件")
        st.markdown('</div>', unsafe_allow_html=True)

        if save_btn:
            st.session_state.course_dur_weeks = cd_val
            st.session_state.min_weeks_req = mw_val
            st.session_state.require_cont = rc_val
            st.success("條件已儲存。")

    st.divider()
    st.markdown("### 檔案上傳與比對")

    col1, col2 = st.columns(2)
    with col1:
        q_file = st.file_uploader("上傳醫院容額表", type=["xlsx"], key="quota_file")
    with col2:
        a_file = st.file_uploader("上傳學生志願／名單表", type=["xlsx"], key="app_file")

    run_check = st.button("確認並開始比對")

    if run_check:
        if not q_file or not a_file:
            st.warning("請同時上傳容額表與學生志願／名單表。")
        else:
            df_q = smart_read_capacity_sheet(q_file)
            df_a_raw = smart_read_application_sheet(a_file)
            df_a = normalize_application_df(df_a_raw, "本院申請資料")

            if df_q is None or df_q.empty:
                st.error("容額表讀取失敗，請確認表頭包含『科別』與日期欄。")
            elif df_a.empty:
                st.error("學生志願／名單表讀取失敗，請確認包含姓名、科別與日期欄位。")
            else:
                course_workdays = st.session_state.course_dur_weeks * 5
                total_min_workdays = st.session_state.min_weeks_req * 5

                # 找出容額表中的時段欄
                date_cols = []
                slot_mapping = {}
                for c in df_q.columns:
                    s_slot, e_slot = extract_dates_universal(c)
                    if s_slot is not None and e_slot is not None:
                        date_cols.append(c)
                        slot_mapping[c] = (pd.to_datetime(s_slot), pd.to_datetime(e_slot))

                if not date_cols:
                    st.error("容額表中找不到可辨識的日期欄位。")
                else:
                    collisions = []

                    for _, q_row in df_q.iterrows():
                        dept = str(q_row.get("科別", "")).strip()
                        if dept in ["", "nan", "None"]:
                            continue

                        dept_apps = df_a[df_a["科別"] == dept].copy()
                        if dept_apps.empty:
                            continue

                        for col in date_cols:
                            cap_val = extract_capacity_value(q_row.get(col))
                            if cap_val is None:
                                continue

                            s_slot, e_slot = slot_mapping[col]
                            in_slot = dept_apps[
                                (dept_apps["開始"] <= e_slot) &
                                (dept_apps["結束"] >= s_slot)
                            ].copy()

                            unique_students = sorted(in_slot["比對鍵"].astype(str).unique().tolist())
                            display_names = sorted(in_slot["姓名"].astype(str).unique().tolist())

                            if len(unique_students) > cap_val:
                                collisions.append({
                                    "科別": dept,
                                    "時間": str(col).replace("\n", " "),
                                    "容額": cap_val,
                                    "申請人數": len(unique_students),
                                    "超額學生": "、".join(display_names)
                                })

                    invalid = []

                    if not df_a.empty:
                        for key, group in df_a.groupby("比對鍵"):
                            group = group.sort_values(["開始", "結束"]).reset_index(drop=True)
                            name = group.iloc[0]["姓名"]
                            total_workdays = group["天數"].sum()

                            # 單一 course 長度不足
                            for _, row in group.iterrows():
                                if row["天數"] < course_workdays:
                                    invalid.append({
                                        "姓名": name,
                                        "原因": f"Course 天數不足：{row['科別']} 僅 {row['天數']} 個工作天（需 {course_workdays} 天）"
                                    })

                            # 總長度不足
                            if total_workdays < total_min_workdays:
                                invalid.append({
                                    "姓名": name,
                                    "原因": f"總時長不足：僅 {total_workdays} 個工作天（需 {total_min_workdays} 天）"
                                })

                            # 連續性檢查
                            if st.session_state.require_cont and len(group) > 1:
                                records = group.to_dict("records")
                                for i in range(len(records) - 1):
                                    prev_end = records[i]["結束"]
                                    next_start = records[i + 1]["開始"]
                                    gap_days = (next_start - prev_end).days
                                    if gap_days > 3:
                                        invalid.append({
                                            "姓名": name,
                                            "原因": f"未連續實習：{records[i]['科別']} 與 {records[i+1]['科別']} 中斷"
                                        })
                                        break

                    st.header("異常監控結果")

                    if collisions:
                        st.subheader("名額撞期名單")
                        st.dataframe(pd.DataFrame(collisions), use_container_width=True)
                    else:
                        st.success("未發現名額超額。")

                    if invalid:
                        st.subheader("規章不符名單")
                        st.dataframe(pd.DataFrame(invalid).drop_duplicates(), use_container_width=True)
                    else:
                        st.success("未發現規章不符。")

# =========================
# 系秘模式
# =========================
elif mode == "系秘":
    st.title("跨院重複佔位檢查")

    st.markdown("### 檔案上傳")
    multi_files = st.file_uploader(
        "上傳各院志願／名單清單（可多選）",
        type=["xlsx"],
        accept_multiple_files=True
    )

    run_check_sec = st.button("確認並開始比對")

    if run_check_sec:
        if not multi_files:
            st.warning("請先上傳至少一份 Excel。")
        else:
            # 提醒是否上傳相同內容檔案
            same_files = detect_same_file_uploads(multi_files)
            if same_files:
                dup_text = "；".join([f"{a} 與 {b}" for a, b in same_files])
                st.warning(f"偵測到內容完全相同的檔案：{dup_text}。若這不是你預期的結果，請先確認是否誤上傳相同名單。")

            all_data = []

            for f in multi_files:
                df_raw = smart_read_application_sheet(f)
                if df_raw is None or df_raw.empty:
                    continue

                hosp_name = f.name.replace(".xlsx", "").replace(".csv", "").strip()
                df_norm = normalize_application_df(df_raw, hosp_name)

                if not df_norm.empty:
                    all_data.append(df_norm)

            if not all_data:
                st.error("沒有成功讀到可用資料，請確認欄位格式。")
            else:
                full_df = pd.concat(all_data, ignore_index=True)
                full_df = full_df.sort_values(["姓名", "開始", "結束"]).reset_index(drop=True)

                conflicts = []

                for key, group in full_df.groupby("比對鍵"):
                    group = group.sort_values(["開始", "結束"]).reset_index(drop=True)
                    records = group.to_dict("records")

                    if len(records) <= 1:
                        continue

                    for i in range(len(records)):
                        for j in range(i + 1, len(records)):
                            a = records[i]
                            b = records[j]

                            if a["來源醫院"] == b["來源醫院"]:
                                continue

                            if periods_overlap(a["開始"], a["結束"], b["開始"], b["結束"]):
                                overlap_type = "完全同時段" if (
                                    a["開始"] == b["開始"] and a["結束"] == b["結束"]
                                ) else "部分重疊"

                                conflicts.append({
                                    "學號": a["學號"] if str(a["學號"]).strip() not in ["", "nan", "None"] else "",
                                    "姓名": a["姓名"],
                                    "重疊型態": overlap_type,
                                    "醫院A": a["來源醫院"],
                                    "科別A": a["科別"],
                                    "開始A": a["開始"].date(),
                                    "結束A": a["結束"].date(),
                                    "醫院B": b["來源醫院"],
                                    "科別B": b["科別"],
                                    "開始B": b["開始"].date(),
                                    "結束B": b["結束"].date(),
                                })

                if conflicts:
                    df_conflicts = pd.DataFrame(conflicts).drop_duplicates()
                    df_conflicts = df_conflicts.sort_values(["姓名", "開始A", "醫院A"]).reset_index(drop=True)

                    st.subheader("偵測到跨院重複佔位")
                    st.dataframe(df_conflicts, use_container_width=True)

                    st.markdown("### 摘要")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.metric("衝突筆數", len(df_conflicts))
                    with c2:
                        st.metric("衝突學生數", df_conflicts["姓名"].nunique())

                    # 只列每位學生一次的摘要
                    summary_rows = []
                    for name, g in df_conflicts.groupby("姓名"):
                        details = []
                        for _, r in g.iterrows():
                            details.append(
                                f"{r['醫院A']}（{r['科別A']} {r['開始A']}~{r['結束A']}）"
                                f" ↔ {r['醫院B']}（{r['科別B']} {r['開始B']}~{r['結束B']}）"
                            )
                        summary_rows.append({
                            "姓名": name,
                            "學號": g.iloc[0]["學號"],
                            "衝突詳情": "\n".join(details)
                        })

                    st.markdown("### 依學生彙整")
                    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)
                else:
                    st.success("無跨院重複佔位情況。")
