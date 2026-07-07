import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

# ==============================================================
# CosBuy — 화장품 원료 견적 비교 대시보드 (Streamlit 버전)
# ==============================================================

st.set_page_config(
    page_title="CosBuy — 원료 견적 비교 대시보드",
    page_icon="\U0001F9EA",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------
# 기본 데이터 (업로드된 "화장품원료_견적비교_30종.xlsx" 내용을 내장)
# ------------------------------------------------------------
DEFAULT_CSV = """원료,공급업체,단가,MOQ,리드타임
히알루론산,Sino Chem Trading,331.08,48,23
히알루론산,그린바이오텍,331.56,25,20
히알루론산,코스켐,347.78,45,20
히알루론산,퓨어사이언스,350.8,14,23
히알루론산,한성원료,370.89,28,24
나이아신아마이드,오션바이오켐,23.59,170,16
나이아신아마이드,신라머티리얼즈,24.85,57,19
나이아신아마이드,대한케미칼,24.91,251,11
나이아신아마이드,EuroLab Ingredients,25.48,94,13
나이아신아마이드,동양파인켐,26.53,56,17
글리세린,네이처글로벌,1.72,795,15
글리세린,세종케미컬,1.73,891,6
글리세린,미래소재,1.89,1837,7
글리세린,바이오랩코리아,1.9,1182,14
글리세린,휴먼텍바이오,1.93,4456,16
세라마이드,코리아켐텍,820.76,7,19
세라마이드,이든바이오,869.99,9,25
세라마이드,대성정밀화학,897.85,3,18
세라마이드,한빛소재,902.96,8,29
세라마이드,청담케미칼,913.15,8,23
콜라겐,Shanghai BioTech,90.96,35,12
콜라겐,Zhejiang Chem Co.,98.93,23,15
콜라겐,태평양화학원료,99.85,188,12
콜라겐,Vietnam Natural Extracts,100.15,123,17
콜라겐,Guangzhou Ingredients Ltd.,100.21,71,10
레티놀,코스켐,3046.65,1.6,20
레티놀,퓨어사이언스,3096.7,1.9,23
레티놀,Sino Chem Trading,3180.71,1,24
레티놀,한성원료,3229.89,2.7,28
레티놀,그린바이오텍,3402.37,1.6,23
스쿠알란,EuroLab Ingredients,30.92,500,18
스쿠알란,동양파인켐,30.99,144,16
스쿠알란,대한케미칼,31.74,341,12
스쿠알란,신라머티리얼즈,32.08,113,12
스쿠알란,오션바이오켐,32.51,57,13
아스코르빈산,휴먼텍바이오,7.08,471,13
아스코르빈산,네이처글로벌,7.2,704,8
아스코르빈산,미래소재,7.65,264,12
아스코르빈산,바이오랩코리아,7.67,442,9
아스코르빈산,세종케미컬,7.86,608,10
판테놀,코리아켐텍,21.07,159,10
판테놀,한빛소재,21.2,197,14
판테놀,이든바이오,21.56,275,11
판테놀,청담케미칼,22.41,206,13
판테놀,대성정밀화학,23.01,161,15
알부틴,태평양화학원료,66.26,43,18
알부틴,Shanghai BioTech,71.16,140,18
알부틴,Guangzhou Ingredients Ltd.,71.7,76,8
알부틴,Zhejiang Chem Co.,72.22,130,9
알부틴,Vietnam Natural Extracts,72.42,122,11
알란토인,코스켐,17.02,119,11
알란토인,퓨어사이언스,17.04,293,14
알란토인,한성원료,18.55,261,12
알란토인,Sino Chem Trading,18.75,415,13
알란토인,그린바이오텍,18.93,454,13
아데노신,신라머티리얼즈,1784.81,4,23
아데노신,EuroLab Ingredients,1854.34,3,31
아데노신,오션바이오켐,1857.31,4,30
아데노신,대한케미칼,1927.37,3,33
아데노신,동양파인켐,1931.85,2,27
펩타이드,바이오랩코리아,3990.91,1.3,33
펩타이드,미래소재,4291.32,2.8,31
펩타이드,네이처글로벌,4306.5,0.8,33
펩타이드,휴먼텍바이오,4401.4,0.9,26
펩타이드,세종케미컬,4451.52,1.1,38
카페인,청담케미칼,11.83,409,14
카페인,이든바이오,11.91,258,11
카페인,대성정밀화학,12.03,483,8
카페인,한빛소재,12.08,218,5
카페인,코리아켐텍,12.4,271,6
티트리오일,Guangzhou Ingredients Ltd.,42.79,97,12
티트리오일,태평양화학원료,44.3,95,10
티트리오일,Zhejiang Chem Co.,44.34,67,19
티트리오일,Vietnam Natural Extracts,47.08,265,11
티트리오일,Shanghai BioTech,48.01,66,15
살리실산,그린바이오텍,8.69,117,14
살리실산,Sino Chem Trading,8.77,184,13
살리실산,한성원료,8.89,143,11
살리실산,퓨어사이언스,8.96,293,13
살리실산,코스켐,9.01,114,9
글리콜릭애씨드,신라머티리얼즈,5.68,399,9
글리콜릭애씨드,EuroLab Ingredients,5.75,762,5
글리콜릭애씨드,대한케미칼,5.79,526,11
글리콜릭애씨드,동양파인켐,5.98,557,6
글리콜릭애씨드,오션바이오켐,6.24,750,6
마데카소사이드,세종케미컬,583.76,14,19
마데카소사이드,네이처글로벌,584.63,10,27
마데카소사이드,휴먼텍바이오,597.28,4,26
마데카소사이드,미래소재,599.72,5,16
마데카소사이드,바이오랩코리아,617.77,9,19
병풀추출물,코리아켐텍,27.15,84,14
병풀추출물,한빛소재,27.2,277,14
병풀추출물,대성정밀화학,27.83,194,11
병풀추출물,청담케미칼,29.65,221,13
병풀추출물,이든바이오,29.97,127,10
어성초추출물,Zhejiang Chem Co.,23.88,184,12
어성초추출물,Vietnam Natural Extracts,24.22,245,15
어성초추출물,Guangzhou Ingredients Ltd.,24.29,146,12
어성초추출물,태평양화학원료,24.3,111,10
어성초추출물,Shanghai BioTech,25.27,86,16
티타늄디옥사이드,그린바이오텍,4.01,1371,10
티타늄디옥사이드,코스켐,4.08,994,6
티타늄디옥사이드,퓨어사이언스,4.17,864,11
티타늄디옥사이드,한성원료,4.33,4864,7
티타늄디옥사이드,Sino Chem Trading,4.46,1963,6
징크옥사이드,신라머티리얼즈,5.45,870,8
징크옥사이드,대한케미칼,5.47,623,8
징크옥사이드,오션바이오켐,5.58,1815,12
징크옥사이드,EuroLab Ingredients,5.72,666,9
징크옥사이드,동양파인켐,5.93,1618,10
토코페롤,바이오랩코리아,15.48,146,13
토코페롤,휴먼텍바이오,15.55,307,7
토코페롤,세종케미컬,15.86,384,7
토코페롤,네이처글로벌,15.96,165,7
토코페롤,미래소재,16.23,383,11
베타글루칸,한빛소재,56.13,141,12
베타글루칸,청담케미칼,57.1,46,10
베타글루칸,코리아켐텍,57.18,42,16
베타글루칸,이든바이오,58.06,136,17
베타글루칸,대성정밀화학,58.79,86,14
세테아릴알코올,Zhejiang Chem Co.,3.42,3846,8
세테아릴알코올,Shanghai BioTech,3.45,1459,7
세테아릴알코올,태평양화학원료,3.49,2469,7
세테아릴알코올,Vietnam Natural Extracts,3.5,4325,6
세테아릴알코올,Guangzhou Ingredients Ltd.,3.58,1055,8
다이메티콘,Sino Chem Trading,6.11,1103,8
다이메티콘,한성원료,6.29,2886,10
다이메티콘,퓨어사이언스,6.29,1849,11
다이메티콘,그린바이오텍,6.34,1251,6
다이메티콘,코스켐,6.41,2814,5
부틸렌글라이콜,동양파인켐,2.28,7779,11
부틸렌글라이콜,대한케미칼,2.36,4908,5
부틸렌글라이콜,EuroLab Ingredients,2.52,2554,5
부틸렌글라이콜,신라머티리얼즈,2.56,2694,6
부틸렌글라이콜,오션바이오켐,2.57,4959,9
잔탄검,네이처글로벌,8,674,6
잔탄검,세종케미컬,8.24,1243,12
잔탄검,미래소재,8.6,755,9
잔탄검,휴먼텍바이오,8.77,895,8
잔탄검,바이오랩코리아,8.99,1194,11
폴리소르베이트20,코리아켐텍,3.01,2344,6
폴리소르베이트20,한빛소재,3.03,2284,8
폴리소르베이트20,이든바이오,3.21,3214,11
폴리소르베이트20,청담케미칼,3.22,2563,8
폴리소르베이트20,대성정밀화학,3.25,1491,7
카보머,Zhejiang Chem Co.,13.31,128,14
카보머,태평양화학원료,13.59,270,13
카보머,Guangzhou Ingredients Ltd.,13.61,493,7
카보머,Shanghai BioTech,14.71,375,14
카보머,Vietnam Natural Extracts,14.86,285,7"""

COSMAX_LOGO_B64 = "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"

# ------------------------------------------------------------
# 스타일 (원본 index.html 디자인 톤을 최대한 재현)
# ------------------------------------------------------------
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap');

html, body, [class*="css"]  {
    font-family: 'Noto Sans KR', system-ui, -apple-system, sans-serif;
}

:root {
    --page: #f1f0ea;
    --surface: #fafaf5;
    --ink-900: #1a1c2e;
    --ink-600: #565a72;
    --ink-400: #8b8e9e;
    --border: #e3e1d8;
    --accent: #283282;
    --button: #e0512b;
    --good: #0f8a4b;
    --warn: #b9770e;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(175deg, #1c2461, #12173f);
}
section[data-testid="stSidebar"] * {
    color: rgba(255,255,255,0.85) !important;
}
section[data-testid="stSidebar"] .stRadio label {
    font-weight: 600;
}
.cosmax-badge {
    display: inline-flex; align-items: center;
    background: #ffffff; border-radius: 8px;
    padding: 8px 14px; margin: 4px 0 14px;
}
.cosmax-badge img { height: 16px; width: auto; display: block; }

.sidebar-brand {
    font-size: 20px; font-weight: 900; color: #ffffff !important;
    letter-spacing: -0.01em; margin-bottom: 6px;
}

.eyebrow {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.08em; color: var(--ink-400); margin: 0 0 2px;
}
.page-title { font-size: 26px; font-weight: 900; margin: 0 0 18px; letter-spacing: -0.01em; color: var(--ink-900); }

div[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 14px 18px 10px;
    box-shadow: 0 1px 2px rgba(20,18,26,0.05);
}
div[data-testid="stMetricLabel"] { font-weight: 700; color: var(--ink-400) !important; }
div[data-testid="stMetricValue"] { font-weight: 900 !important; color: var(--ink-900) !important; }

.material-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 18px 20px;
    margin-bottom: 16px;
    box-shadow: 0 1px 2px rgba(20,18,26,0.05);
}
.material-title { font-size: 16px; font-weight: 800; color: var(--ink-900); margin: 0 0 2px; }
.material-sub { font-size: 12.5px; color: var(--ink-400); margin: 0 0 12px; }

table.cosbuy-table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
table.cosbuy-table th {
    text-align: left; font-size: 11.5px; font-weight: 700; color: var(--ink-400);
    text-transform: uppercase; letter-spacing: 0.04em;
    padding: 8px 10px; border-bottom: 1px solid var(--border);
}
table.cosbuy-table td {
    padding: 9px 10px; border-bottom: 1px solid var(--border);
    color: var(--ink-900);
}
table.cosbuy-table tr:last-child td { border-bottom: none; }
.badge-best {
    display: inline-flex; align-items: center; gap: 4px;
    background: rgba(15,138,75,0.10); color: var(--good);
    font-weight: 800; font-size: 12px; padding: 3px 9px; border-radius: 999px;
}
.badge-plain {
    color: var(--warn); font-weight: 700; font-family: ui-monospace, monospace; font-size: 12.5px;
}
.footnote { font-size: 12px; color: var(--ink-400); margin-top: 6px; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ------------------------------------------------------------
# 세션 상태 초기화
# ------------------------------------------------------------
def load_default_df():
    df = pd.read_csv(io.StringIO(DEFAULT_CSV))
    df["단가"] = df["단가"].astype(float)
    df["MOQ"] = df["MOQ"].astype(float)
    df["리드타임"] = df["리드타임"].astype(float)
    return df

if "df" not in st.session_state:
    st.session_state.df = load_default_df()
if "weekly_uploads" not in st.session_state:
    st.session_state.weekly_uploads = 3
if "search" not in st.session_state:
    st.session_state.search = ""
if "price_min" not in st.session_state:
    st.session_state.price_min = None
if "price_max" not in st.session_state:
    st.session_state.price_max = None
if "lead_filter" not in st.session_state:
    st.session_state.lead_filter = "전체"
if "sort_option" not in st.session_state:
    st.session_state.sort_option = "최저단가 낮은순"

# ------------------------------------------------------------
# 파일 파싱 유틸 (열 이름 자동 인식: 원본 JS 로직과 동일한 규칙)
# ------------------------------------------------------------
def find_col(columns, patterns):
    for col in columns:
        c = str(col).strip().lower()
        for p in patterns:
            if p in c:
                return col
    return None

def parse_uploaded_file(file):
    name = file.name
    ext = name.split(".")[-1].lower()
    try:
        if ext == "csv":
            raw = file.read()
            try:
                text = raw.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = raw.decode("cp949", errors="ignore")
            df = pd.read_csv(io.StringIO(text))
        else:
            df = pd.read_excel(file)
    except Exception:
        return None, "read_error"

    if df.empty:
        return None, "empty"

    cols = list(df.columns)
    mat_col = find_col(cols, ["원료", "material", "품목"])
    sup_col = find_col(cols, ["업체", "공급", "supplier", "vendor", "회사"])
    price_col = find_col(cols, ["단가", "가격", "price", "cost"])
    moq_col = find_col(cols, ["moq", "최소주문", "minimum"])
    lead_col = find_col(cols, ["리드타임", "lead", "납기"])

    if not mat_col or not sup_col or not price_col:
        return None, "missing_columns"

    out_rows = []
    for _, row in df.iterrows():
        material = str(row.get(mat_col, "")).strip()
        supplier = str(row.get(sup_col, "")).strip()
        try:
            price = float(str(row.get(price_col, "")).replace(",", "").strip())
        except (ValueError, TypeError):
            continue
        if not material or not supplier or material == "nan" or supplier == "nan":
            continue
        try:
            moq = float(str(row.get(moq_col, 0)).replace(",", "").strip()) if moq_col else 0.0
        except (ValueError, TypeError):
            moq = 0.0
        try:
            lead = float(str(row.get(lead_col, 0)).replace(",", "").strip()) if lead_col else 0.0
        except (ValueError, TypeError):
            lead = 0.0
        out_rows.append({"원료": material, "공급업체": supplier, "단가": price, "MOQ": moq, "리드타임": lead})

    if not out_rows:
        return None, "no_valid_rows"

    return pd.DataFrame(out_rows), "ok"

def merge_into_state(new_df):
    df = st.session_state.df.copy()
    added, updated, new_materials = 0, 0, 0
    existing_materials = set(df["원료"].unique())

    for _, r in new_df.iterrows():
        mask = (df["원료"].str.lower() == r["원료"].lower()) & (df["공급업체"].str.lower() == r["공급업체"].lower())
        if mask.any():
            df.loc[mask, ["단가", "MOQ", "리드타임"]] = [r["단가"], r["MOQ"], r["리드타임"]]
            updated += 1
        else:
            df = pd.concat([df, pd.DataFrame([r])], ignore_index=True)
            added += 1
        if r["원료"] not in existing_materials:
            new_materials += 1
            existing_materials.add(r["원료"])

    st.session_state.df = df
    st.session_state.weekly_uploads += added
    return added, updated, new_materials

# ------------------------------------------------------------
# 사이드바
# ------------------------------------------------------------
with st.sidebar:
    st.markdown(
        f'<div class="cosmax-badge"><img src="data:image/webp;base64,{COSMAX_LOGO_B64}" /></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="sidebar-brand">\U0001F4E6 CosBuy</div>', unsafe_allow_html=True)

    st.caption("메인")
    nav = st.radio(
        "메인 메뉴", ["원료 비교", "공급업체", "견적함", "리포트"],
        label_visibility="collapsed",
    )
    st.caption("관리")
    st.button("\u2699\uFE0F 설정", use_container_width=True)

    st.markdown("---")
    st.markdown("**\uD83D\uDC65 구매팀**")
    st.caption("팀원 10명")

if nav != "원료 비교":
    st.markdown(f'<p class="eyebrow">구매팀 워크스페이스</p><p class="page-title">{nav}</p>', unsafe_allow_html=True)
    st.info(f"{nav} 화면은 준비 중입니다.")
    st.stop()

# ------------------------------------------------------------
# 헤더
# ------------------------------------------------------------
st.markdown('<p class="eyebrow">구매팀 워크스페이스</p><p class="page-title">원료 비교</p>', unsafe_allow_html=True)

df = st.session_state.df

# ------------------------------------------------------------
# 통계 카드
# ------------------------------------------------------------
def compute_stats(df):
    n_materials = df["원료"].nunique()
    n_suppliers = df["공급업체"].nunique()
    savings = []
    for _, g in df.groupby("원료"):
        best, worst = g["단가"].min(), g["단가"].max()
        if best > 0:
            savings.append((worst - best) / best * 100)
    avg_saving = float(np.mean(savings)) if savings else 0.0
    return n_materials, n_suppliers, avg_saving

n_mat, n_sup, avg_saving = compute_stats(df)

c1, c2, c3, c4 = st.columns(4)
c1.metric("비교 중인 원료", f"{n_mat}종", help="전체 카탈로그 기준")
c2.metric("참여 공급업체", f"{n_sup}곳", help="이번 견적에 응답")
c3.metric("평균 절감 여지", f"{avg_saving:.1f}%", help="최저가 대비 최고가 평균 격차")
c4.metric("이번 주 신규 견적", f"{st.session_state.weekly_uploads}건", help="업로드 반영 포함")

st.write("")

# ------------------------------------------------------------
# 업로드 존
# ------------------------------------------------------------
with st.container(border=True):
    st.markdown("#### \U0001F4E4 견적서를 업로드하세요")
    st.caption("엑셀(.xlsx, .xls) 또는 CSV 견적서를 올리면 자동으로 파싱되어 아래 표에 즉시 반영됩니다.")
    uploaded_files = st.file_uploader(
        "견적서 업로드",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    sample_csv = "원료,공급업체,단가,MOQ,리드타임\n히알루론산,신규업체A,340.00,20,14\n"
    st.download_button(
        "\U0001F4C4 샘플 양식 다운로드",
        data=sample_csv.encode("utf-8-sig"),
        file_name="cosbuy_upload_template.csv",
        mime="text/csv",
    )

    if uploaded_files:
        total_added, total_updated, total_new_materials = 0, 0, 0
        errors = []
        for f in uploaded_files:
            new_df, status = parse_uploaded_file(f)
            if status == "ok":
                a, u, nmat = merge_into_state(new_df)
                total_added += a
                total_updated += u
                total_new_materials += nmat
            elif status == "missing_columns":
                errors.append(f"{f.name}: 필수 열(원료/공급업체/단가)을 찾을 수 없습니다.")
            elif status == "read_error":
                errors.append(f"{f.name}: 파일을 읽을 수 없습니다.")
            else:
                errors.append(f"{f.name}: 반영할 견적을 찾지 못했습니다.")

        if total_added + total_updated > 0:
            parts = []
            if total_new_materials > 0:
                parts.append(f"신규 원료 {total_new_materials}종")
            parts.append(f"신규 견적 {total_added}건")
            if total_updated > 0:
                parts.append(f"기존 견적 갱신 {total_updated}건")
            st.success("완료 — " + ", ".join(parts) + " 반영됨")
            df = st.session_state.df
        for e in errors:
            st.error(e)

st.write("")

# ------------------------------------------------------------
# 필터바
# ------------------------------------------------------------
with st.container(border=True):
    fc1, fc2, fc3, fc4, fc5 = st.columns([2.2, 1.6, 1.2, 1.6, 0.8])
    with fc1:
        search = st.text_input("검색", placeholder="원료명 또는 공급업체명 검색", label_visibility="visible")
    with fc2:
        pcol1, pcol2 = st.columns(2)
        price_min = pcol1.number_input("최소 단가", min_value=0.0, value=0.0, step=1.0, format="%.2f")
        price_max = pcol2.number_input("최대 단가", min_value=0.0, value=0.0, step=1.0, format="%.2f")
    with fc3:
        lead_filter = st.selectbox("리드타임", ["전체", "7일 이내", "14일 이내", "21일 이내", "30일 이내"])
    with fc4:
        sort_option = st.selectbox(
            "원료 정렬",
            ["최저단가 낮은순", "최저단가 높은순", "원료명 오름차순", "원료명 내림차순"],
        )
    with fc5:
        st.write("")
        reset = st.button("초기화", use_container_width=True)
        if reset:
            st.rerun()

# ------------------------------------------------------------
# 필터링 + 정렬 + 렌더링
# ------------------------------------------------------------
filtered = df.copy()

if search:
    s = search.strip().lower()
    mask = filtered["원료"].str.lower().str.contains(s) | filtered["공급업체"].str.lower().str.contains(s)
    filtered = filtered[mask]

if price_min and price_min > 0:
    filtered = filtered[filtered["단가"] >= price_min]
if price_max and price_max > 0:
    filtered = filtered[filtered["단가"] <= price_max]

lead_map = {"7일 이내": 7, "14일 이내": 14, "21일 이내": 21, "30일 이내": 30}
if lead_filter in lead_map:
    filtered = filtered[filtered["리드타임"] <= lead_map[lead_filter]]

groups = list(filtered.groupby("원료"))

if sort_option == "최저단가 낮은순":
    groups.sort(key=lambda kv: kv[1]["단가"].min())
elif sort_option == "최저단가 높은순":
    groups.sort(key=lambda kv: kv[1]["단가"].min(), reverse=True)
elif sort_option == "원료명 오름차순":
    groups.sort(key=lambda kv: kv[0])
elif sort_option == "원료명 내림차순":
    groups.sort(key=lambda kv: kv[0], reverse=True)

total_groups = len(groups)
total_rows = sum(len(g) for _, g in groups)

st.markdown(
    f"전체 원료 **{total_groups}종** · 견적 **{total_rows}건** 표시 중",
)

st.markdown("#### 원료별 공급업체 비교")
st.caption("단가는 최근 접수된 견적 기준이며, MOQ·리드타임은 공급업체가 제시한 조건입니다. '최저가 대비'는 현재 필터에 표시된 업체 중 최저 단가를 기준으로 계산됩니다.")

if total_groups == 0:
    st.warning("조건에 맞는 견적이 없습니다. 필터를 조정해 보세요.")
else:
    for material, g in groups:
        g_sorted = g.sort_values("단가")
        best_price = g_sorted["단가"].min()

        rows_html = ""
        for _, r in g_sorted.iterrows():
            moq_display = int(r["MOQ"]) if float(r["MOQ"]).is_integer() else round(r["MOQ"], 1)
            lead_display = int(r["리드타임"]) if float(r["리드타임"]).is_integer() else r["리드타임"]
            if r["단가"] == best_price:
                delta_html = '<span class="badge-best">\u2605 최적</span>'
            else:
                pct = (r["단가"] - best_price) / best_price * 100
                delta_html = f'<span class="badge-plain">+{pct:.1f}%</span>'
            rows_html += f"""
            <tr>
                <td>{r['공급업체']}</td>
                <td>{r['단가']:,.2f}</td>
                <td>{moq_display}</td>
                <td>{lead_display}일</td>
                <td>{delta_html}</td>
            </tr>
            """

        table_html = f"""
        <div class="material-card">
            <p class="material-title">{material}</p>
            <p class="material-sub">공급업체 {len(g_sorted)}곳 · 최저 단가 {best_price:,.2f} USD/kg</p>
            <table class="cosbuy-table">
                <thead>
                    <tr>
                        <th>공급업체</th>
                        <th>단가 (USD/kg)</th>
                        <th>MOQ (kg)</th>
                        <th>리드타임</th>
                        <th>최저가 대비</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
        </div>
        """
        st.markdown(table_html, unsafe_allow_html=True)
