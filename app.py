import streamlit as st
import pandas as pd
import re
from io import BytesIO
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm

# 한글 폰트 설정 (나눔스퀘어)
def setup_font():
    try:
        # 나눔스퀘어 폰트 경로들 시도
        font_paths = [
            '/usr/share/fonts/truetype/nanum/NanumSquareR.ttf',  # Linux
            'C:/Windows/Fonts/NanumSquareR.ttf',  # Windows
            '/System/Library/Fonts/AppleGothic.ttf',  # Mac 대체
            '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'  # 대체폰트
        ]
        
        for font_path in font_paths:
            try:
                font_name = fm.FontProperties(fname=font_path).get_name()
                plt.rcParams['font.family'] = font_name
                plt.rcParams['axes.unicode_minus'] = False
                return font_path
            except:
                continue
        
        # 폰트를 찾지 못한 경우 기본 설정
        plt.rcParams['font.family'] = 'DejaVu Sans'
        return None
    except:
        return None

# 폰트 설정 실행
FONT_PATH = setup_font()

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
    """엑셀 시트에서 실제 데이터 표가 시작하는 행을 자동으로 찾는 함수"""
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
    
    # 간단한 결과 알림만 표시
    if success_count > 0:
        st.success(f"✅ {success_count}개 파일 처리 완료!")
    else:
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
    selected = merged[merged["과거"]<threshold].sort_values("증가율", ascending=False)
    return selected.reset_index()

def label_chip(label, value, color="black", bg="#DDD"):
    return f"<span style='padding:4px 12px 4px 12px; border-radius:15px; background:{bg}; color:{color}; margin-right:8px; font-size:0.95em;'>{label}: {value}</span>"

# =============== Streamlit Layout =================

st.set_page_config(layout='wide')
st.title("🚀 연관어 빅데이터 자동 전처리·시각화 툴")
st.markdown("##### 엑셀 여러개 업로드하면 **자동 테이블 감지**로 연도/시트별 취합, 시각화까지 한 번에!")

uploaded_files = st.file_uploader(
    "엑셀 파일 여러 개 업로드", 
    type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner('📊 파일 처리 중...'):
        df = merge_and_standardize(uploaded_files)
    
    if df.empty:
        st.stop()

    # ================== 미리보기 & 라벨 ==================
    show_cols = ["연도", "분석채널", "순위", "연관어", "건수", "카테고리 대분류", "카테고리 소분류"]
    view_cols = [col for col in show_cols if col in df.columns]
    
    meta_cols = ["주제어", "동의어", "포함어", "분석채널"]
    meta_info = {col: (df[col].dropna().unique()[:3] if col in df.columns else ["-"]) for col in meta_cols}
    
    st.markdown("#### 📋 [ 미리보기 / Preview ]")
    label_html = " ".join([
        label_chip(lbl, ', '.join([str(v) for v in val if v!='-']),
        color="white" if lbl in ('분석채널',) else "black",
        bg="#222" if lbl in ('분석채널',) else "#eee")
        for lbl, val in meta_info.items()
    ])
    st.markdown(label_html, unsafe_allow_html=True)
    
    col1, col2 = st.columns([4, 1])
    with col1:
        st.dataframe(df[view_cols].head(20), use_container_width=True)
    with col2:
        st.download_button(
            "📥 엑셀 다운로드", 
            data=to_excel(df), 
            file_name="통합_연관어_취합.xlsx", 
            mime="application/vnd.ms-excel"
        )

    # ================== 연도, 채널 선택 시각화 ==================
    st.markdown("#### 📊 [ 연관어/카테고리별 분석 및 시각화 ]")
    
    col1, col2 = st.columns(2)
    with col1:
        year_list = list(sorted(df["연도"].dropna().unique()))
        year_sel = st.selectbox("📅 연도 선택", year_list, key="year")
    
    with col2:
        channel_list = ["전체"] + list(sorted(df["분석채널"].dropna().unique())) if "분석채널" in df.columns else ["전체"]
        ch_sel = st.selectbox("📺 분석채널", channel_list, key="ch")
    
    # 데이터 필터링
    if "분석채널" in df.columns:
        view_df = df[(df["연도"]==year_sel) & ((df["분석채널"]==ch_sel) if ch_sel!="전체" else True)]
    else:
        view_df = df[df["연도"]==year_sel]

    if not view_df.empty:
        # 워드클라우드와 버블차트를 나란히 배치
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**☁️ [워드클라우드]**")
            if "연관어" in view_df.columns and "건수" in view_df.columns:
                try:
                    word_freq = view_df.groupby("연관어")["건수"].sum().to_dict()
                    if word_freq:
                        wc = WordCloud(
                            width=500, height=300, 
                            background_color='white', 
                            font_path=FONT_PATH,
                            max_words=30
                        ).generate_from_frequencies(word_freq)
                        
                        fig, ax = plt.subplots(figsize=(6, 4))
                        ax.imshow(wc, interpolation='bilinear')
                        ax.axis('off')
                        st.pyplot(fig, clear_figure=True)
                    else:
                        st.info("데이터 부족")
                except Exception as e:
                    st.error(f"워드클라우드 생성 오류: {e}")
            else:
                st.info("연관어/건수 컬럼 없음")

        with col2:
            st.markdown("**🫧 [버블차트]**")
            if all(x in view_df.columns for x in ["순위","건수","연관어"]):
                try:
                    fig, ax = plt.subplots(figsize=(6, 4))
                    
                    if "카테고리 대분류" in view_df.columns:
                        sns.scatterplot(
                            data=view_df.head(15), x="순위", y="건수", 
                            size="건수", hue="카테고리 대분류", 
                            sizes=(50, 500), alpha=0.7, ax=ax
                        )
                    else:
                        sns.scatterplot(
                            data=view_df.head(15), x="순위", y="건수", 
                            size="건수", sizes=(50, 500), alpha=0.7, ax=ax
                        )
                    
                    # 상위 5개만 텍스트 표시
                    for _, r in view_df.head(5).iterrows():
                        try:
                            ax.text(r["순위"], r["건수"], str(r["연관어"])[:8], 
                                   fontsize=8, alpha=0.8, ha='center')
                        except:
                            pass
                    
                    ax.set_title(f"{year_sel}년 {ch_sel} 연관어 분석")
                    st.pyplot(fig, clear_figure=True)
                except Exception as e:
                    st.error(f"버블차트 생성 오류: {e}")
            else:
                st.info("필수 컬럼 부족")

        # 카테고리 분석
        st.markdown("#### 📈 [카테고리 분석]")
        col1, col2 = st.columns(2)
        
        with col1:
            if "카테고리 대분류" in view_df.columns and "건수" in view_df.columns:
                st.markdown("**대분류 Top5**")
                top_major = view_df.groupby("카테고리 대분류")["건수"].sum().sort_values(ascending=False).head(5)
                st.dataframe(top_major, use_container_width=True)
        
        with col2:
            if "카테고리 소분류" in view_df.columns and "건수" in view_df.columns:
                st.markdown("**소분류 Top5**")
                top_minor = view_df.groupby("카테고리 소분류")["건수"].sum().sort_values(ascending=False).head(5)
                st.dataframe(top_minor, use_container_width=True)

    # Rising keyword
    st.markdown("#### 🚀 [Rising Keyword 탐색]")
    unique_years = df["연도"].unique()
    if len(unique_years) >= 2:
        n_years = min(3, len(unique_years))
        n_year = st.slider("최근 N년 기준", 1, n_years, 2)
        
        rising_df = rising_keywords(df, recent_n=n_year)
        if not rising_df.empty:
            col1, col2 = st.columns([2, 1])
            with col1:
                try:
                    fig2, ax2 = plt.subplots(figsize=(8, 4))
                    sns.scatterplot(
                        data=rising_df.head(10), x="증가율", y="최근", 
                        size="최근", sizes=(30, 300), alpha=0.7, ax=ax2
                    )
                    
                    for _, r in rising_df.head(8).iterrows():
                        try: 
                            ax2.text(r["증가율"], r["최근"], str(r["연관어"])[:6], 
                                    fontsize=8, alpha=0.8)
                        except: 
                            pass
                    
                    ax2.set_title("Rising Keywords")
                    st.pyplot(fig2, clear_figure=True)
                except Exception as e:
                    st.error(f"Rising Keyword 차트 오류: {e}")
            
            with col2:
                st.dataframe(rising_df.head(10), use_container_width=True)
        else:
            st.info("Rising Keyword 데이터 없음")
    else:
        st.info("최소 2개 연도 데이터 필요")

else:
    st.info("🔼 엑셀 파일을 업로드하면 자동으로 처리됩니다.")
    
    with st.expander("💡 사용법"):
        st.markdown("""
        **지원 파일:** .xlsx 엑셀 (여러 시트 지원)  
        **자동 감지:** 순위/연관어/건수 헤더 자동 찾기  
        **한글 지원:** 나눔스퀘어 폰트 적용  
        **기능:** 워드클라우드, 버블차트, Rising Keyword  
        """)
