import streamlit as st
import pandas as pd
import re
from io import BytesIO
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns

# ===================== 도우미 함수 =====================
def extract_year_from_filename(filename):
    found = re.findall(r'(\d{2})\d{4,}', filename)  # 예: 250101 -> '25' 추출
    year = None
    if found:
        year = int('20' + found[0])  # '25'->2025
    else:
        # 혹시 다른 방식 연도 들어감 대비
        found = re.findall(r'20\d{2}', filename)
        year = int(found[0]) if found else None
    return year

def load_and_label_excel(file, year):
    # 여러 시트를 모두 읽어서, sheet_name과 year 컬럼 추가해서 반환(DataFrame 리스트)
    xls = pd.ExcelFile(file)
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
        year = extract_year_from_filename(upfile.name)
        dfs = load_and_label_excel(upfile, year)
        all_dfs.extend(dfs)
    df = pd.concat(all_dfs, ignore_index=True)
    # 칼럼명 통일/정렬
    df = df.rename(columns={
        "순위": "순위", 
        "연관어": "연관어", 
        "건수": "건수",
        "카테고리 대분류": "카테고리 대분류",
        "카테고리 소분류": "카테고리 소분류"
    })
    return df

# Excel로 다운로드: stream에 적재
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def rising_keywords(df, recent_n=2, threshold=3):
    # rising keyword: 최근 n년 동안 급증, 이전엔 거의 없던 단어
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

# ================== Streamlit 레이아웃 ==================

st.set_page_config(layout="wide")
st.title("연관어 빅데이터 자동 전처리 · 시각화 툴")
st.markdown("##### 엑셀 여러개 `drag & drop`하면 자동 연도/시트별 취합, 미리보기, 라벨, 연관어·채널·카테고리별 시각화와 다운로드까지 한 번에!")
uploaded_files = st.file_uploader(
    "엑셀 파일 여러 개 업로드 (예시: 썸트렌드_여름여행_연관어_250101-250730.xlsx)", 
    type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # 1,2,3. 데이터 병합
    df = merge_and_standardize(uploaded_files)
    
    # 주제어, 동의어, 포함어 유추(칼럼이 있다면) // 필요한 값만 추출
    meta_cols = ["주제어", "동의어", "포함어", "분석채널"]
    meta_info = {col: df[col].dropna().unique()[:3] if col in df else ["-"] for col in meta_cols}
    
    # 귀여운 label(각이 둥근 칩 스타일)
    def label_chip(label, value, color="black", bg="#DDD"):
        return f"<span style='padding:4px 12px 4px 12px; border-radius:15px; background:{bg}; color:{color}; margin-right:8px; font-size:0.95em;'>{label}: {value}</span>"

    st.markdown("#### [ 미리보기 / Preview ]")
    label_html = " ".join([label_chip(lbl, ', '.join([str(v) for v in val if v!='-']), color="white" if lbl in ('분석채널',) else "black", bg="#222" if lbl in ('분석채널',) else "#eee") for lbl, val in meta_info.items()])
    st.markdown(label_html, unsafe_allow_html=True)
    st.dataframe(df[["연도","분석채널","순위", "연관어", "건수", "카테고리 대분류", "카테고리 소분류"]].head(20), use_container_width=True)

    # 3. 다운로드 버튼
    st.download_button("엑셀 파일 다운받기", data=to_excel(df), file_name="통합_연관어_취합.xlsx", mime="application/vnd.ms-excel")

    # 4. YEAR/채널/카테고리 필터와 시각화
    st.markdown("#### [ 연관어/카테고리별 분석 및 시각화 ]")
    year_sel = st.selectbox("연도 선택", list(sorted(df["연도"].unique())), key="year")
    ch_sel = st.selectbox("분석채널", ["전체"]+list(sorted(df["분석채널"].unique())), key="ch")
    view_df = df[(df["연도"]==year_sel) & ((df["분석채널"]==ch_sel) if ch_sel!="전체" else True)]

    # 워드클라우드
    word_freq = view_df.groupby("연관어")["건수"].sum().to_dict()
    wc = WordCloud(width=700, height=400, background_color='white', font_path=None).generate_from_frequencies(word_freq)
    st.pyplot(plt.figure(figsize=(9,5))); plt.imshow(wc); plt.axis('off')

    # 버블 차트
    st.markdown("**[버블차트] : 연관어별 건수**")
    if len(view_df):
        plt.figure(figsize=(10,7))
        sns.scatterplot(
            data=view_df, x="순위", y="건수", size="건수", hue="카테고리 대분류", legend=False, sizes=(100, 1800), alpha=0.3)
        for _, r in view_df.iterrows():
            plt.text(r["순위"], r["건수"], r["연관어"], fontsize=9, alpha=0.8)
        st.pyplot(plt)

    # 5. 대분류/소분류 count
    st.markdown("#### [올해 가장 많이 언급된 대분류/소분류]")
    st.write(view_df.groupby("카테고리 대분류")["건수"].sum().sort_values(ascending=False))
    st.write(view_df.groupby("카테고리 소분류")["건수"].sum().sort_values(ascending=False))

    # 6. Rising Keyword
    st.markdown("#### [Rising Keyword 탐색]")
    n_year = st.slider("최근 N년 기준", 1, min(3, len(df["연도"].unique())), 2)
    rising_df = rising_keywords(df, recent_n=n_year)
    st.dataframe(rising_df.head(10), use_container_width=True)
    if len(rising_df):
        st.markdown("**Rising Keyword Bubble Chart**")
        plt.figure(figsize=(7,5))
        sns.scatterplot(data=rising_df, x="증가율", y="최근", size="최근", hue="최근", sizes=(10, 800))
        for _, r in rising_df.iterrows():
            plt.text(r["증가율"], r["최근"], r["연관어"], fontsize=9)
        st.pyplot(plt)
else:
    st.info("엑셀 파일을 여러개 업로드하면 자동으로 연도/채널별 취합과 전처리, 시각화가 시작됩니다.")


