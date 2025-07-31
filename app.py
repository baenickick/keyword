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

# -------------- 표 시작 행 자동 감지 --------------
def find_table_start(file, sheet_name):
    """엑셀 시트에서 실제 데이터 표가 시작하는 행을 자동으로 찾는 함수"""
    try:
        # 처음 30행 정도만 읽어서 헤더 위치 찾기
        df_preview = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=30)
        
        # 찾을 키워드들 (실제 데이터 헤더로 예상되는 것들)
        header_keywords = ['순위', '연관어', '건수', '카테고리', 'rank', 'keyword', 'count']
        
        for row_idx in range(len(df_preview)):
            row_values = df_preview.iloc[row_idx].astype(str).str.lower()
            # 키워드 중 하나라도 포함된 행을 찾으면 그게 헤더
            if any(keyword.lower() in ' '.join(row_values) for keyword in header_keywords):
                return row_idx
        
        # 키워드를 못 찾으면 기본값 0 반환
        return 0
    except:
        return 0

# -------------- 엑셀 파일(sheet별) 읽기 (자동감지 적용) --------------
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
        
        if not xls.sheet_names:
            st.error(f"{file.name} 내에 읽을 수 있는 시트가 없습니다.")
            return []
        
        dfs = []
        for sheet_name in xls.sheet_names:
            try:
                # 표 시작 행 자동 감지
                header_row = find_table_start(in_memory_file, sheet_name)
                st.info(f"📊 {file.name} [{sheet_name}]: {header_row+1}번째 행에서 데이터 표 시작 감지")
                
                # 감지된 행부터 데이터 읽기
                df = pd.read_excel(in_memory_file, sheet_name=sheet_name, header=header_row)
                df.columns = df.columns.str.strip()
                
                # 데이터가 비어있는지 확인
                if df.empty or len(df) == 0:
                    st.warning(f"{file.name} [{sheet_name}]: 데이터가 없습니다.")
                    continue
                
                # 필수 컬럼 중 하나라도 있는지 확인
                essential_cols = ['순위', '연관어', '건수']
                if not any(col in df.columns for col in essential_cols):
                    st.warning(f"{file.name} [{sheet_name}]: 필수 컬럼(순위/연관어/건수)을 찾을 수 없습니다.")
                    continue
                
                df['연도'] = year
                df['분석채널'] = sheet_name
                dfs.append(df)
                st.success(f"✅ {file.name} [{sheet_name}]: {len(df)}행 데이터 로드 완료")
                
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
        if year is None:
            st.warning(f"⚠️ {upfile.name}: 파일명에서 연도를 추출할 수 없습니다.")
            year = 2024  # 기본값
        
        dfs = load_and_label_excel(upfile, year)
        if len(dfs) == 0:
            st.warning(f"⚠️ {upfile.name} 파일에서 데이터를 불러오지 못했습니다.")
        all_dfs.extend(dfs)
    
    if not all_dfs:
        st.error("업로드한 모든 파일에서 데이터를 추출하지 못했습니다. 엑셀 시트 구조나 파일 자체를 확인하세요.")
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
st.markdown("##### 엑셀 여러개 `drag & drop`하면 **자동 테이블 감지**로 연도/시트별 취합, 미리보기, 시각화까지 한 번에!")

uploaded_files = st.file_uploader(
    "엑셀 파일 여러 개 업로드 (예: 썸트렌드_여름여행_연관어_250101-250730.xlsx)", 
    type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner('📊 파일 업로드 중... 자동으로 데이터 테이블을 감지하고 있습니다.'):
        df = merge_and_standardize(uploaded_files)
    
    if df.empty:
        st.error("❌ 처리할 수 있는 데이터가 없습니다.")
        st.stop()

    # 성공 메시지
    st.success(f"🎉 총 {len(df)}행의 데이터가 성공적으로 병합되었습니다!")
    
    # 칼럼명 리스트 직접 표시(문제 진단 확인용)
    with st.expander("🔍 데이터 구조 확인"):
        st.write("**실제 DataFrame 칼럼:**", df.columns.tolist())
        st.write("**데이터 형태:**", df.shape)

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
    st.dataframe(df[view_cols].head(20), use_container_width=True)
    
    # 다운로드 버튼
    st.download_button(
        "📥 엑셀 파일 다운받기", 
        data=to_excel(df), 
        file_name="통합_연관어_취합.xlsx", 
        mime="application/vnd.ms-excel"
    )

    # ================== 연도, 채널 선택 시각화 ==================
    st.markdown("#### 📊 [ 연관어/카테고리별 분석 및 시각화 ]")
    
    # 필터 설정
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

    if view_df.empty:
        st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    else:
        # 워드클라우드
        st.markdown("**☁️ [워드클라우드]**")
        if "연관어" in view_df.columns and "건수" in view_df.columns:
            try:
                word_freq = view_df.groupby("연관어")["건수"].sum().to_dict()
                if word_freq:
                    wc = WordCloud(
                        width=700, height=400, 
                        background_color='white', 
                        font_path=None,
                        max_words=50
                    ).generate_from_frequencies(word_freq)
                    
                    fig, ax = plt.subplots(figsize=(10, 6))
                    ax.imshow(wc, interpolation='bilinear')
                    ax.axis('off')
                    st.pyplot(fig)
                else:
                    st.info("워드클라우드 생성을 위한 데이터가 부족합니다.")
            except Exception as e:
                st.error(f"워드클라우드 생성 중 오류: {e}")
        else:
            st.info("워드클라우드 생성을 위한 연관어/건수 컬럼이 없습니다.")

        # 버블차트
        st.markdown("**🫧 [버블차트 (순위 vs 건수)]**")
        if all(x in view_df.columns for x in ["순위","건수","연관어"]):
            try:
                fig, ax = plt.subplots(figsize=(12,8))
                
                # 카테고리 대분류가 있으면 색상으로 구분
                if "카테고리 대분류" in view_df.columns:
                    sns.scatterplot(
                        data=view_df.head(20), x="순위", y="건수", 
                        size="건수", hue="카테고리 대분류", 
                        sizes=(100, 1500), alpha=0.7, ax=ax
                    )
                else:
                    sns.scatterplot(
                        data=view_df.head(20), x="순위", y="건수", 
                        size="건수", sizes=(100, 1500), alpha=0.7, ax=ax
                    )
                
                # 연관어 텍스트 추가
                for _, r in view_df.head(15).iterrows():
                    try:
                        ax.text(r["순위"], r["건수"], str(r["연관어"])[:10], 
                               fontsize=8, alpha=0.8, ha='center')
                    except:
                        pass
                
                ax.set_title(f"{year_sel}년 {ch_sel} 연관어 분석")
                st.pyplot(fig)
            except Exception as e:
                st.error(f"버블차트 생성 중 오류: {e}")
        else:
            st.info("버블차트 생성을 위한 필수 컬럼이 부족합니다.")

        # 대분류/소분류 합계 랭킹
        st.markdown("#### 📈 [가장 많이 언급된 카테고리]")
        col1, col2 = st.columns(2)
        
        with col1:
            if "카테고리 대분류" in view_df.columns and "건수" in view_df.columns:
                st.markdown("**대분류 Top5**")
                top_major = view_df.groupby("카테고리 대분류")["건수"].sum().sort_values(ascending=False).head(5)
                st.dataframe(top_major)
        
        with col2:
            if "카테고리 소분류" in view_df.columns and "건수" in view_df.columns:
                st.markdown("**소분류 Top5**")
                top_minor = view_df.groupby("카테고리 소분류")["건수"].sum().sort_values(ascending=False).head(5)
                st.dataframe(top_minor)

    # Rising keyword
    st.markdown("#### 🚀 [Rising Keyword 탐색]")
    unique_years = df["연도"].unique()
    if len(unique_years) >= 2:
        n_years = min(3, len(unique_years))
        n_year = st.slider("최근 N년 기준", 1, n_years, 2)
        
        rising_df = rising_keywords(df, recent_n=n_year)
        if not rising_df.empty:
            st.dataframe(rising_df.head(10), use_container_width=True)
            
            st.markdown("**🚀 Rising Keyword Bubble Chart**")
            try:
                fig2, ax2 = plt.subplots(figsize=(10,6))
                sns.scatterplot(
                    data=rising_df.head(15), x="증가율", y="최근", 
                    size="최근", sizes=(50, 800), alpha=0.7, ax=ax2
                )
                
                for _, r in rising_df.head(10).iterrows():
                    try: 
                        ax2.text(r["증가율"], r["최근"], str(r["연관어"])[:8], 
                                fontsize=9, alpha=0.8)
                    except: 
                        pass
                
                ax2.set_title("Rising Keywords (증가율 vs 최근 언급량)")
                st.pyplot(fig2)
            except Exception as e:
                st.error(f"Rising Keyword 차트 생성 중 오류: {e}")
        else:
            st.info("Rising Keyword 데이터가 없거나 조건에 맞는 키워드가 없습니다.")
    else:
        st.info("Rising Keyword 탐색을 위해서는 최소 2개 연도의 데이터가 필요합니다.")

else:
    st.info("🔼 엑셀 파일을 여러개 업로드하면 자동으로 테이블을 감지하고 연도/채널별 취합과 전처리, 시각화가 시작됩니다.")
    
    # 사용법 안내
    with st.expander("💡 사용법 안내"):
        st.markdown("""
        **지원하는 파일 형식:**
        - .xlsx 엑셀 파일 (여러 시트 지원)
        - 파일명에 연도 정보 포함 (예: 250101, 240315 등)
        
        **자동 감지 기능:**
        - 📊 엑셀 시트에서 '순위', '연관어', '건수' 등 헤더 자동 감지
        - 🔍 메타 정보와 실제 데이터 테이블 구분
        - 📅 파일명에서 연도 자동 추출
        
        **제공 기능:**
        - 🔄 여러 파일/시트 자동 병합
        - ☁️ 워드클라우드 시각화
        - 🫧 버블차트 분석
        - 🚀 Rising Keyword 탐지
        - 📥 결과 엑셀 다운로드
        """)
