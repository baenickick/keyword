import streamlit as st
import pandas as pd
import re
from io import BytesIO
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns

# --------------- 연도 추출 함수 -----------------
def extract_year_from_filename(filename):
    found = re.findall(r'(\d{2})\d{4,}', filename)  # 예: 250101 -> "25"
    year = None
    if found:
        year = int('20' + found[0])
    else:
        found = re.findall(r'20\d{2}', filename)
        year = int(found[0]) if found else None
    return year

# -------------- 엑셀 파일(sheeet별) 읽기 --------------
def load_and_label_excel(file, year):
    # 파일 바이너리 인식 위해 BytesIO 처리
    in_memory_file = BytesIO(file.read())
    xls = pd.ExcelFile(in_memory_file)
    dfs = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df['연도'] = year
        df['분석채널'] = sheet_name
        dfs.append(df)
    return dfs

def merge_and_standardize(files):
    all_dfs = []
    for upfile in files:
        upfile.seek(0)  # 여러 번 읽을 경우 포인터 복원
        year = extract_year_from_filename(upfile.name)
        dfs = load_and_label_excel(upfile, year)
        all_dfs.extend(dfs)
    df = pd.concat(all_dfs, ignore_index=True)
    # 주요 칼럼만 남기고 컬럼명 통일 (불필요시 삭제)
    rename_dict = {
        "순위": "순위", 
        "연관어": "연관어", 
        "건수": "건수",
        "카테고리 대분류": "카테고리 대분류",
        "카테고리 소분류": "카테고리 소분류"
    }
    curr_cols = df.columns
    for k in list(rename_dict):
        if k not in curr_cols:
            df[rename_dict[k]] = None   # 칼럼 없는 경우 빈값 추가
    df = df[list(rename_dict) + ['연도','분석채널'] + [c for c in df.columns if c not in list(rename_dict)+['연도','분석채널']]]
    return df

# ----------- Excel 다운로드 ------------
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ----------- Rising Keyword 계산 ----------
def rising_keywords(df, recent_n=2, threshold=3):
    all_years = sorted(df["연도"].unique())
    if len(all_years) <= recent_n: return pd.DataFrame()
    prev_years, recent_years = all_years[:-recent_n], all_years[-recent_n:]
    prev_df = df[df["연도"].isin(prev_years)]
    recent_df = df[df["연도"].isin(recent_years)]
    prev_count = prev_df.groupby("연관어")["건수"].sum()
    recent_count = recent_df.groupby("연관어")["건수"].sum()
    merged = pd.DataFrame({"과거": prev_count, "최근": recent_count}).fillna(0)
    merged["증가율"] = (merged["최근"]-merged["과거"])/(merged["과거"]+1)
    selected = merged[merged["과거"]<threshold].sort_values("증가율", ascending=False)
    return selected.reset_index()

# ----------- 라벨 칩(HTML) -------------
def label_chip(label, value, color="black", bg="#DDD"):
    return f"<span style='padding:4px 12px 4px 12px; border-radius:15px; background:{bg}; color:{color}; margin-right:8px; font-size:0.95em;'>{label}: {value}</span>"

# =============== Streamlit Layout =================

st.set_page_config(layout='wide')
st.title("연관어 빅데이터 자동 전처리·시각화 툴")
st.markdown("##### 엑셀 여러개 `drag & drop`하면 자동 연도/시트별 취합, 미리보기, 라벨, 연관어·채널·카테고리별 시각화와 다운로드까지 한 번에!")

uploaded_files = st.file_uploader(
    "엑셀 파일 여러 개 업로드 (예: 썸트렌드_여름여행_연관어_250101-250730.xlsx)", 
    type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    df = merge_and_standardize(uploaded_files)

    # ===== 미리보기 & 라벨 =====
    meta_cols = ["주제어", "동의어", "포함어", "분석채널"]
    meta_info = {col: df[col].dropna().unique()[:3] if col in df else ["-"] for col in meta_cols}
    label_html = " ".join([
        label_chip(lbl, ', '.join([str(v) for v in val if v!='-']), 
        color="white" if lbl in ('분석채널',) else "black", 
        bg="#222" if lbl in ('분석채널',) else "#eee")
        for lbl, val in meta_info.items()
    ])
    st.markdown("#### [ 미리보기 / Preview ]")
    st.markdown(label_html, unsafe_allow_html=True)
    st.dataframe(df[["연도", "분석채널", "순위", "연관어", "건수", "카테고리 대분류", "카테고리 소분류"]].head(20), use_container_width=True)
    st.download_button("엑셀 파일 다운받기", data=to_excel(df), file_name="통합_연관어_취합.xlsx", mime="application/vnd.ms-excel")

    # ===== 연도, 채널 선택 시각화 =====
    st.markdown("#### [ 연관어/카테고리별 분석 및 시각화 ]")
    year_sel = st.selectbox("연도 선택", list(sorted(df["연도"].dropna().unique())), key="year")
    ch_sel = st.selectbox("분석채널", ["전체"]+list(sorted(df["분석채널"].dropna().unique())), key="ch")
    view_df = df[(df["연도"]==year_sel) & ((df["분석채널"]==ch_sel) if ch_sel!="전체" else True)]

    # 워드클라우드
    st.markdown("**[워드클라우드]**")
    if len(view_df):
        word_freq = view_df.groupby("연관어")["건수"].sum().to_dict()
        wc = WordCloud(width=700, height=400, background_color='white', font_path=None).generate_from_frequencies(word_freq)
        plt.figure(figsize=(9, 5))
        plt.imshow(wc); plt.axis('off')
        st.pyplot(plt)
    else:
        st.info("해당 데이터 없음")

    # 버블차트
    st.markdown("**[버블차트 (순위 vs 건수, 원 크기는 건수, 색은 대분류)]**")
    if len(view_df):
        fig, ax = plt.subplots(figsize=(10,7))
        sns.scatterplot(
            data=view_df, x="순위", y="건수", size="건수", hue="카테고리 대분류", 
            legend=False, sizes=(100, 1800), alpha=0.3, ax=ax)
        for _, r in view_df.iterrows():
            try:
                ax.text(r["순위"], r["건수"], str(r["연관어"]), fontsize=9, alpha=0.8)
            except:
                pass
        st.pyplot(fig)
    else:
        st.info("해당 데이터 없음")

    # 대분류/소분류 합계 랭킹
    st.markdown("#### [올해 가장 많이 언급된 대분류/소분류]")
    st.write("대분류 Top5", view_df.groupby("카테고리 대분류")["건수"].sum().sort_values(ascending=False).head(5))
    st.write("소분류 Top5", view_df.groupby("카테고리 소분류")["건수"].sum().sort_values(ascending=False).head(5))

    # Rising keyword
    st.markdown("#### [Rising Keyword 탐색]")
    n_year = st.slider("최근 N년 기준", 1, min(3, len(df["연도"].unique())), 2)
    rising_df = rising_keywords(df, recent_n=n_year)
    if len(rising_df):
        st.dataframe(rising_df.head(10), use_container_width=True)
        st.markdown("**Rising Keyword Bubble Chart**")
        fig2, ax2 = plt.subplots(figsize=(7,5))
        sns.scatterplot(data=rising_df, x="증가율", y="최근", size="최근", hue="최근", sizes=(10, 800), ax=ax2)
        for _, r in rising_df.head(10).iterrows():
            try: ax2.text(r["증가율"], r["최근"], str(r["연관어"]), fontsize=9)
            except: pass
        st.pyplot(fig2)
    else:
        st.info("Rising Keyword 데이터 없음")

else:
    st.info("엑셀 파일을 여러개 업로드하면 자동으로 연도/채널별 취합과 전처리, 시각화가 시작됩니다.")
