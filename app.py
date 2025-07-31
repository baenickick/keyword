import streamlit as st
import pandas as pd
import re
from io import BytesIO
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
import os

# 1. 나눔스퀘어_ac 폰트 적용
NANUM_FONT_PATH = 'NanumSquare_acR.ttf'
def setup_nanum_font():
    font_paths = [
        './NanumSquare_acR.ttf',
        '/usr/share/fonts/truetype/nanum/NanumSquare_acR.ttf',
        'C:/Windows/Fonts/NanumSquare_acR.ttf',
        '/System/Library/Fonts/NanumSquare_acR.ttf'
    ]
    for path in font_paths:
        if os.path.exists(path):
            plt.rcParams['font.family'] = fm.FontProperties(fname=path).get_name()
            plt.rcParams['axes.unicode_minus'] = False
            return path
    st.error("❗ '나눔스퀘어_ac' 폰트 파일(NanumSquare_acR.ttf)을 프로젝트 폴더나 시스템 폰트 경로에 넣어주세요!")
    plt.rcParams['font.family'] = 'NanumSquare_acR'
    return None
FONT_PATH = setup_nanum_font()

def extract_year_from_filename(filename):
    found = re.findall(r'(\d{2})\d{4,}', filename)
    year = None
    if found:
        year = int('20' + found[0])
    else:
        found = re.findall(r'20\d{2}', filename)
        year = int(found[0]) if found else None
    return year

def find_table_start(file, sheet_name):
    try:
        df_preview = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=30)
        header_keywords = ['순위', '연관어', '건수', '카테고리', 'rank', 'keyword', 'count']
        for row_idx in range(len(df_preview)):
            row_values = df_preview.iloc[row_idx].astype(str).str.lower()
            if any(keyword.lower() in ' '.join(row_values) for keyword in header_keywords):
                return row_idx
        return 0
    except:
        return 0

def load_and_label_excel(file, year):
    try:
        file.seek(0)
        sig = file.read(4)
        if sig != b'PK\x03\x04':
            return []
        file.seek(0)
        in_memory_file = BytesIO(file.read())
        xls = pd.ExcelFile(in_memory_file)
        if not xls.sheet_names:
            return []
        dfs = []
        for sheet_name in xls.sheet_names:
            try:
                header_row = find_table_start(in_memory_file, sheet_name)
                df = pd.read_excel(in_memory_file, sheet_name=sheet_name, header=header_row)
                df.columns = df.columns.str.strip()
                if df.empty or len(df) == 0:
                    continue
                essential_cols = ['순위', '연관어', '건수']
                if not any(col in df.columns for col in essential_cols):
                    continue
                df['연도'] = year
                df['분석채널'] = sheet_name
                dfs.append(df)
            except Exception:
                continue
        return dfs
    except Exception:
        return []

def merge_and_standardize(files):
    all_dfs = []
    success_count = 0
    for upfile in files:
        upfile.seek(0)
        year = extract_year_from_filename(upfile.name)
        if year is None:
            year = 2024
        dfs = load_and_label_excel(upfile, year)
        if len(dfs) > 0:
            success_count += 1
        all_dfs.extend(dfs)
    if success_count == 0:
        st.error("❌ 처리할 수 있는 파일이 없습니다.")
        return pd.DataFrame()
    df = pd.concat(all_dfs, ignore_index=True)
    df.columns = df.columns.str.strip()
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def rising_keywords(df, recent_n=2, threshold=3):
    if '연관어' not in df.columns or '건수' not in df.columns or '연도' not in df.columns:
        return pd.DataFrame()
    all_years = sorted(df["연도"].unique())
    if len(all_years) <= recent_n:
        return pd.DataFrame()
    prev_years, recent_years = all_years[:-recent_n], all_years[-recent_n:]
    prev_df = df[df["연도"].isin(prev_years)]
    recent_df = df[df["연도"].isin(recent_years)]
    prev_count = prev_df.groupby("연관어")["건수"].sum()
    recent_count = recent_df.groupby("연관어")["건수"].sum()
    merged = pd.DataFrame({"과거": prev_count, "최근": recent_count}).fillna(0)
    merged["증가율"] = (merged["최근"]-merged["과거"])/(merged["과거"]+1)
    selected = merged[merged["과거"] < threshold].sort_values("증가율", ascending=False)
    return selected.reset_index()

def label_chip(label, value, color="black", bg="#DDD"):
    return f"<span style='padding:4px 12px 4px 12px; border-radius:15px; background:{bg}; color:{color}; margin-right:8px; font-size:0.95em;'>{label}: {value}</span>"

# 와이드 레이아웃 + 중앙 max-width 가상 컨테이너
st.set_page_config(layout='wide')
st.markdown("<style>.main-container-wrap{max-width:1100px; margin:0 auto;}</style>", unsafe_allow_html=True)
st.markdown("<div class='main-container-wrap'>", unsafe_allow_html=True)
st.markdown("<h1 style='margin-top:12px;'>연관어 빅데이터 자동 전처리·시각화 툴</h1>", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "엑셀 파일 여러 개 업로드", 
    type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    with st.spinner('📊 파일 처리 중...'):
        df = merge_and_standardize(uploaded_files)
    if df.empty:
        st.markdown("</div>", unsafe_allow_html=True)
        st.stop()

    show_cols = ["연도", "분석채널", "순위", "연관어", "건수", "카테고리 대분류", "카테고리 소분류"]
    view_cols = [col for col in show_cols if col in df.columns]
    meta_cols = ["주제어", "동의어", "포함어", "분석채널"]
    meta_info = {col: (df[col].dropna().unique()[:3] if col in df.columns else ["-"]) for col in meta_cols}

    st.markdown("#### 📋 Preview")
    label_html = " ".join([
        label_chip(lbl, ', '.join([str(v) for v in val if v != '-']),
                   color="white" if lbl == '분석채널' else "black", 
                   bg="#222" if lbl == '분석채널' else "#eee")
        for lbl, val in meta_info.items()
    ])
    # 다운로드 버튼을 데이터 표 우측 상단에 띄우는 CSS
    st.markdown("""
        <style>.stDataFrame{padding-top:0!important;}.xy-btn{position:absolute;top:60px;right:30px;z-index:1;}@media(max-width:1200px){.xy-btn{right:13vw;}}</style>
        <div class='xy-btn'>""", unsafe_allow_html=True)
    st.download_button(
        "📥 엑셀 다운로드", data=to_excel(df), file_name="통합_연관어_취합.xlsx", mime="application/vnd.ms-excel"
    )
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(label_html, unsafe_allow_html=True)
    st.dataframe(df[view_cols].head(20), use_container_width=True)

    st.markdown("#### 📊 시각화 및 분석 조건 선택")
    col_year, col_ch, col_main, col_sub = st.columns([1,1,2,2])

    with col_year:
        year_options = ["전체"] + list(map(str, sorted(df["연도"].dropna().unique())))
        year_sel = st.selectbox("📅 연도", year_options, key="year")
    with col_ch:
        channel_options = ["전체"] + list(map(str, sorted(df["분석채널"].dropna().unique()))) if "분석채널" in df.columns else ["전체"]
        ch_sel = st.selectbox("📺 분석채널", channel_options, key="ch")
    with col_main:
        major_list = sorted(df["카테고리 대분류"].dropna().unique()) if "카테고리 대분류" in df.columns else []
        if major_list:
            major_sel = st.multiselect("대분류(복수선택)", ["전체"] + major_list, ["전체"])
        else:
            major_sel = ["전체"]
    with col_sub:
        minor_list = sorted(df["카테고리 소분류"].dropna().unique()) if "카테고리 소분류" in df.columns else []
        if minor_list:
            minor_sel = st.multiselect("소분류(복수선택)", ["전체"] + minor_list, ["전체"])
        else:
            minor_sel = ["전체"]

    chart_col, table_col = st.columns([2, 1])
    with chart_col:
        chart_options = ["워드클라우드", "버블차트"]
        chart_sel = st.selectbox("💡 차트 스타일", chart_options, key="chart")

    # ====== 필터링 ======
    view_df = df.copy()
    if year_sel != "전체":
        view_df = view_df[view_df["연도"] == int(year_sel)]
    if "분석채널" in view_df.columns and ch_sel != "전체":
        view_df = view_df[view_df["분석채널"] == ch_sel]
    if "카테고리 대분류" in view_df.columns and "전체" not in major_sel:
        view_df = view_df[view_df["카테고리 대분류"].isin(major_sel)]
    if "카테고리 소분류" in view_df.columns and "전체" not in minor_sel:
        view_df = view_df[view_df["카테고리 소분류"].isin(minor_sel)]

    st.markdown("---")

    # ======= 차트 (왼쪽) + 탑 20 (오른쪽) =======
    with chart_col:
        if chart_sel == "워드클라우드":
            st.markdown("#### ☁️ 워드클라우드")
            if "연관어" in view_df.columns and "건수" in view_df.columns and len(view_df):
                word_freq = view_df.groupby("연관어")["건수"].sum().to_dict()
                if word_freq and FONT_PATH:
                    wc = WordCloud(
                        width=1000, height=500,
                        background_color='white',
                        font_path=FONT_PATH, max_words=100
                    ).generate_from_frequencies(word_freq)
                    fig, ax = plt.subplots(figsize=(10,4))
                    ax.imshow(wc, interpolation='bilinear')
                    ax.axis('off')
                    st.pyplot(fig)
                else:
                    st.info("워드클라우드: 데이터/폰트 부족")
            else:
                st.info("연관어/건수 컬럼 없음.")
        elif chart_sel == "버블차트":
            st.markdown("#### 🫧 버블차트")
            if all(x in view_df.columns for x in ["순위","건수","연관어"]) and len(view_df):
                fig, ax = plt.subplots(figsize=(10,5))
                if "카테고리 대분류" in view_df.columns:
                    sns.scatterplot(
                        data=view_df.head(30), x="순위", y="건수",
                        size="건수", hue="카테고리 대분류",
                        sizes=(100, 1200), alpha=0.7, ax=ax
                    )
                else:
                    sns.scatterplot(
                        data=view_df.head(30), x="순위", y="건수",
                        size="건수", sizes=(100, 1200), alpha=0.7, ax=ax
                    )
                for _, r in view_df.head(15).iterrows():
                    try:
                        ax.text(r["순위"], r["건수"], str(r["연관어"])[:10],
                                fontsize=11, alpha=0.85, ha='center',
                                fontproperties=fm.FontProperties(fname=FONT_PATH))
                    except:
                        pass
                ax.set_title("연관어 순위 vs 건수", fontsize=15,
                             fontproperties=fm.FontProperties(fname=FONT_PATH))
                st.pyplot(fig)
            else:
                st.info("버블차트: 데이터 부족/컬럼 없음.")

    with table_col:
        st.markdown("#### 🔝 TOP 20 키워드")
        if "연관어" in view_df.columns and "건수" in view_df.columns:
            top20 = view_df.groupby("연관어")["건수"].sum().sort_values(ascending=False).head(20)
            st.dataframe(top20.reset_index().rename(columns={"연관어":"연관어","건수":"건수"}), use_container_width=True)
        else:
            st.info("데이터 부족")

    # ==== 카테고리 분석 ====
    st.markdown("#### 📈 카테고리 분석")
    col11, col12 = st.columns(2)
    with col11:
        if "카테고리 대분류" in view_df.columns and "건수" in view_df.columns:
            st.markdown("**대분류 Top5**")
            st.dataframe(view_df.groupby("카테고리 대분류")["건수"].sum().sort_values(ascending=False).head(5), use_container_width=True)
    with col12:
        if "카테고리 소분류" in view_df.columns and "건수" in view_df.columns:
            st.markdown("**소분류 Top5**")
            st.dataframe(view_df.groupby("카테고리 소분류")["건수"].sum().sort_values(ascending=False).head(5), use_container_width=True)

    # ==== Rising Keywords ====
    st.markdown("#### 🚀 Rising Keyword")
    unique_years = df["연도"].unique()
    if len(unique_years) >= 2:
        n_years = min(3, len(unique_years))
        n_year = st.slider("최근 N년 기준", 1, n_years, 2)
        rising_df = rising_keywords(df, recent_n=n_year)
        if not rising_df.empty:
            col1, col2 = st.columns([2, 1])
            with col1:
                fig2, ax2 = plt.subplots(figsize=(10,4))
                sns.scatterplot(
                    data=rising_df.head(10), x="증가율", y="최근",
                    size="최근", sizes=(60, 400), alpha=0.7, ax=ax2
                )
                for _, r in rising_df.head(8).iterrows():
                    try:
                        ax2.text(r["증가율"], r["최근"], str(r["연관어"])[:12], fontsize=11, alpha=0.9, fontproperties=fm.FontProperties(fname=FONT_PATH))
                    except:
                        pass
                ax2.set_title("Rising Keywords", fontsize=15, fontproperties=fm.FontProperties(fname=FONT_PATH))
                st.pyplot(fig2)
            with col2:
                st.dataframe(rising_df.head(10), use_container_width=True)
        else:
            st.info("Rising Keyword 데이터 없음")
    else:
        st.info("최소 2개 연도 데이터 필요")

    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.info("🔼 엑셀 파일을 업로드하면 자동으로 처리됩니다. 예시 파일 구조/조건에 꼭 맞춰주세요.")
