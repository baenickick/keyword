import streamlit as st
import pandas as pd
import re
from io import BytesIO
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns

def extract_year_from_filename(filename):
    found = re.findall(r'(\d{2})\d{4,}', filename)
    year = None
    if found:
        year = int('20' + found[0])
    else:
        found = re.findall(r'20\d{2}', filename)
        year = int(found[0]) if found else None
    return year

def load_and_label_excel(file, year):
    try:
        file.seek(0)
        sig = file.read(4)
        if sig != b'PK\x03\x04':
            st.error(f"{file.name}: 정상적인 엑셀(xlsx) 파일이 아닙니다.")
            return []
        file.seek(0)
        in_memory_file = BytesIO(file.read())
        xls = pd.ExcelFile(in_memory_file)
        dfs = []
        for sheet_name in xls.sheet_names:
            # 표 데이터가 15번째 줄에서 시작(헤더), 위는 메타정보!
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=14)
                df.columns = df.columns.str.strip()
                if df.shape[0] == 0 or ("연관어" not in df.columns):
                    st.warning(f"{file.name} [{sheet_name}]: 데이터가 없거나 칼럼명 불일치")
                    continue
                df['연도'] = year
                df['분석채널'] = sheet_name
                dfs.append(df)
            except Exception as e:
                st.warning(f"{file.name}의 시트 [{sheet_name}] 로딩 실패: {e}")
        return dfs
    except Exception as e:
        st.error(f"{file.name} 파일을 읽는 중 문제 발생: {e}")
        return []

def merge_and_standardize(files):
    all_dfs = []
    for upfile in files:
        upfile.seek(0)
        year = extract_year_from_filename(upfile.name)
        dfs = load_and_label_excel(upfile, year)
        if len(dfs) == 0:
            st.warning(f"⚠️ {upfile.name} 파일에서 데이터를 불러오지 못했습니다.")
        all_dfs.extend(dfs)
    if not all_dfs:
        st.error("업로드한 모든 파일에서 데이터를 추출하지 못했습니다. 엑셀 시트 구조나 파일 자체를 확인하세요.")
        return pd.DataFrame()
    df = pd.concat(all_dfs, ignore_index=True)
    df.columns = df.columns.str.strip()  # <--- 칼럼명 공백 제거 (최종합본에도 적용)
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

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

def label_chip(label, value, color="black", bg="#DDD"):
    return f"<span style='padding:4px 12px 4px 12px; border-radius:15px; background:{bg}; color:{color}; margin-right:8px; font-size:0.95em;'>{label}: {value}</span>"

st.set_page_config(layout='wide')
st.title("연관어 빅데이터 자동 전처리·시각화 툴")
st.markdown("##### 엑셀 여러개 `drag & drop`하면 자동 연도/시트별 취합, 미리보기, 라벨, 연관어·채널·카테고리별 시각화와 다운로드까지 한 번에!")

uploaded_files = st.file_uploader(
    "엑셀 파일 여러 개 업로드 (예: 썸트렌드_여름여행_연관어_250101-250730.xlsx)", 
    type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    df = merge_and_standardize(uploaded_files)
    if df.empty:
        st.stop()

    # 칼럼명 리스트 직접 표시(문제 진단 확인용)
    st.write("실제 DataFrame 칼럼:", df.columns.tolist())

    # ================== 미리보기 & 라벨 ==================
    show_cols = ["연도", "분석채널", "순위", "연관어", "건수", "카테고리 대분류", "카테고리 소분류"]
    view_cols = [col for col in show_cols if col in df.columns]
    
    meta_cols = ["주제어", "동의어", "포함어", "분석채널"]
    meta_info = {col: (df[col].dropna().unique()[:3] if col in df else ["-"]) for col in meta_cols}
    label_html = " ".join([
        label_chip(lbl, ', '.join([str(v) for v in val if v!='-']),
        color="white" if lbl in ('분석채널',) else "black",
        bg="#222" if lbl in ('분석채널',) else "#eee")
        for lbl, val in meta_info.items()
    ])
    st.markdown("#### [ 미리보기 / Preview ]")
    st.markdown(label_html, unsafe_allow_html=True)
    st.dataframe(df[view_cols].head(20), use_container_width=True)
    st.download_button("엑셀 파일 다운받기", data=to_excel(df), file_name="통합_연관어_취합.xlsx", mime="application/vnd.ms-excel")

    # ================== 연도, 채널 선택 시각화 ==================
    year_list = list(sorted(df["연도"].dropna().unique()))
    channel_list = ["전체"] + list(sorted(df["분석채널"].dropna().unique())) if "분석채널" in df.columns else ["전체"]
    st.markdown("#### [ 연관어/카테고리별 분석 및 시각화 ]")
    year_sel = st.selectbox("연도 선택", year_list, key="year")
    ch_sel = st.selectbox("분석채널", channel_list, key="ch")
    # "분석채널"이 칼럼에 없을 경우 전체 데이터 사용
    if "분석채널" in df.columns:
        view_df = df[(df["연도"]==year_sel) & ((df["분석채널"]==ch_sel) if ch_sel!="전체" else True)]
    else:
        view_df = df[df["연도"]==year_sel]

    # 워드클라우드
    st.markdown("**[워드클라우드]**")
    if "연관어" in view_df.columns and "건수" in view_df.columns and len(view_df):
        word_freq = view_df.groupby("연관어")["건수"].sum().to_dict()
        wc = WordCloud(width=700, height=400, background_color='white', font_path=None).generate_from_frequencies(word_freq)
        plt.figure(figsize=(9, 5))
        plt.imshow(wc); plt.axis('off')
        st.pyplot(plt)
    else:
        st.info("워드클라우드 생성을 위한 연관어/건수 데이터가 없습니다.")

    # 버블차트
    st.markdown("**[버블차트 (순위 vs 건수, 원 크기는 건수, 색은 대분류)]**")
    if all(x in view_df.columns for x in ["순위","건수","카테고리 대분류","연관어"]) and len(view_df):
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
        st.info("버블차트 생성을 위한 칼럼 또는 데이터가 부족합니다.")

    # 대분류/소분류 합계 랭킹
    st.markdown("#### [올해 가장 많이 언급된 대분류/소분류]")
    if "카테고리 대분류" in view_df.columns and "건수" in view_df.columns:
        st.write("대분류 Top5", view_df.groupby("카테고리 대분류")["건수"].sum().sort_values(ascending=False).head(5))
    if "카테고리 소분류" in view_df.columns and "건수" in view_df.columns:
        st.write("소분류 Top5", view_df.groupby("카테고리 소분류")["건수"].sum().sort_values(ascending=False).head(5))

    # Rising keyword
    st.markdown("#### [Rising Keyword 탐색]")
    unique_years = df["연도"].unique()
    if len(unique_years) >= 2:
        n_years = min(3, len(unique_years))
        n_year = st.slider("최근 N년 기준", 1, n_years, 2)
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
        st.info("Rising Keyword 탐색을 위한 연도별 데이터가 충분하지 않습니다.")
else:
    st.info("엑셀 파일을 여러개 업로드하면 자동으로 연도/채널별 취합과 전처리, 시각화가 시작됩니다.")
