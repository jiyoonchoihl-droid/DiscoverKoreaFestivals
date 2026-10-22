# ============================================================
# Discover Korea Through Festivals
# 한국 지역 축제 큐레이션 서비스
#
# 실행 파일: festival_app.py
# 데이터 파일: festival_data_4개.xlsx
#
# 주의: dashboard.py는 수정하지 않습니다.
# ============================================================

import streamlit as st
import pandas as pd
import pydeck as pdk
from pathlib import Path
from urllib.parse import quote
import base64
import textwrap
import re
import os
import requests
from datetime import date, datetime, timedelta
from dotenv import load_dotenv
from html import escape


def get_poster_image(festival_name, fallback):
    from pathlib import Path
    import base64
    import mimetypes

    poster_names = {
        "강릉커피축제": "gangneung",
        "수원화성문화제": "suwon",
        "춘천마임축제": "chuncheon",
        "전주비빔밥축제": "jeonju",
    }

    name = poster_names.get(festival_name)
    if not name:
        return fallback

    folder = Path(__file__).resolve().parent / "festival_posters"

    for ext in [".jpg", ".jpeg", ".png", ".webp"]:
        path = folder / (name + ext)
        if path.exists():
            mime = mimetypes.guess_type(path.name)[0] or "image/jpeg"
            encoded = base64.b64encode(path.read_bytes()).decode("ascii")
            return f"data:{mime};base64,{encoded}"

    return fallback




def compact_html(content):
    """Prevent Markdown from treating indented HTML as code."""
    return re.sub(r"\s+", " ", content).strip()


# ------------------------------------------------------------
# 1. 페이지 기본 설정
# ------------------------------------------------------------

st.set_page_config(
    page_title="Discover Korea Through Festivals",
    page_icon="✳️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ------------------------------------------------------------
# 2. 디자인 CSS
# ------------------------------------------------------------

st.markdown(
    textwrap.dedent("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@300;400;500;600&family=Noto+Sans+KR:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', 'Noto Sans KR', sans-serif;
    }

    .stApp {
        background-color: #F7F6F2;
        color: #242820;
    }

    .block-container {
        max-width: 1440px;
        padding-top: 1.5rem;
        padding-bottom: 4rem;
        padding-left: 4%;
        padding-right: 4%;
    }

    #MainMenu, footer {
        visibility: hidden;
    }

    header[data-testid="stHeader"] {
        background-color: #F7F6F2;
    }

    .topline {
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #DADBD2;
        padding-bottom: 17px;
        margin-bottom: 42px;
        color: #62675C;
        font-size: 12px;
        letter-spacing: 2px;
    }

    .brand {
        color: #283B30;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 1.6px;
    }

    .eyebrow {
        color: #9A806A;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 3px;
        margin-bottom: 14px;
    }

    .hero-title {
        font-family: 'DM Sans', 'Noto Sans KR', sans-serif;
        color: #25382C;
        font-size: clamp(38px, 5.4vw, 76px);
        font-weight: 600;
        letter-spacing: -3.5px;
        line-height: 1.08;
        margin: 0 0 22px 0;
    }

    .hero-subtitle {
        font-size: 15px;
        line-height: 1.9;
        color: #676B60;
        max-width: 590px;
        margin-bottom: 24px;
    }

    .hero-note {
        display: inline-block;
        color: #3D5945;
        border: 1px solid #B9C4B4;
        border-radius: 2px;
        padding: 9px 13px;
        font-size: 11px;
        letter-spacing: 1px;
    }

    .hero-image {
        width: 100%;
        height: 330px;
        object-fit: cover;
        border-radius: 3px;
        display: block;
    }

    .section-kicker {
        color:  #9A806A;
        font-size: 11px;
        letter-spacing: 3px;
        font-weight: 700;
        margin-bottom: 8px;
    }

    .section-title {
        color: #45382F;
        font-size: 32px;
        font-weight: 600;
        letter-spacing: -1px;
        line-height: 1.3;
        margin: 0 0 10px 0;
    }

    .section-desc {
        color: #796F65;
        font-size: 17px;
        line-height: 1.8;
        margin-bottom: 22px;
    }

    .stat-box {
        border-top: 1px solid #C9CEC3;
        border-bottom: 1px solid #C9CEC3;
        padding: 15px 0;
        margin: 22px 0 28px 0;
    }

    .stat-number {
        color: #45382F;
        font-size: 25px;
        font-weight: 600;
    }

    .stat-label {
        color: #797D72;
        font-size: 11px;
        letter-spacing: 1px;
    }

    .festival-card {
        background: #FFFFFF;
        border: 1px solid #E8E7E0;
        border-radius: 3px;
        overflow: hidden;
        margin-bottom: 12px;
    }

    .festival-card img {
        width: 100%;
        height: 185px;
        object-fit: cover;
        display: block;
    }

    .festival-card-body {
        padding: 18px 18px 12px 18px;
    }

    .card-region {
        color: #9A806A;
        font-size: 14px;
        font-weight: 700;
        letter-spacing: 2px;
        margin-bottom: 8px;
    }

    .card-title {
        color: #45382F;
        font-size: 20px;
        font-weight: 600;
        line-height: 1.45;
        margin-bottom: 9px;
    }

    .card-description {
        color: #796F65;
        font-size: 17px;
        line-height: 1.6;
        min-height: 0;
        margin-bottom: 12px;
    }

    .detail-panel {
        background: #FFFFFF;
        border: 1px solid #E5E5DD;
        border-radius: 3px;
        padding: 25px;
        margin-top: 15px;
    }

    .detail-name {
        color: #45382F;
        font-size: 26px;
        font-weight: 600;
        margin-bottom: 10px;
    }


    /* Where to Next: 설명 간격 미세 조정 */
    .st-key-map .detail-text {
        line-height: 1.65;
    }

    .detail-text {
        color: #796F65;
        font-size: 18px;
        line-height: 1.9;
    }

    .map-caption {
        color: #796F65;
        font-size: 11px;
        line-height: 1.8;
        margin-top: 8px;
    }

    .footer-line {
        border-top: 1px solid #DADBD2;
        margin-top: 55px;
        padding-top: 22px;
        color: #85877D;
        font-size: 11px;
        letter-spacing: 1px;
    }

    div.stButton > button {
        background-color: #45382F;
        color: #FFFFFF;
        border: 1px solid #45382F;
        border-radius: 2px;
        min-height: 43px;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 0.5px;
        transition: all 0.2s ease;
    }

    div.stButton > button:hover {
        background-color: #705A49;
        color: #FFFFFF;
        border-color: #705A49;
    }

    div[data-baseweb="select"] > div {
        background-color: #FFFFFF;
        border-color: #D8DBD1;
        border-radius: 2px;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-left: 5%;
            padding-right: 5%;
        }

        .hero-title {
            letter-spacing: -1.5px;
        }

        .hero-image {
            height: 220px;
        }
    }
/* HOME: full-screen background and hidden Streamlit header */
.stApp:has(.st-key-home_hero) {
    background-color: #F7F6F2;
}

.stApp:has(.st-key-home_hero) header[data-testid="stHeader"] {
    display: none !important;
}

.stApp:has(.st-key-home_hero) .block-container {
    max-width: 100% !important;
    width: 100% !important;
    padding: 0 !important;
    margin: 0 !important;
}

.stApp:has(.st-key-home_hero) [data-testid="stMainBlockContainer"] {
    padding: 0 !important;
    margin: 0 !important;
}

.stApp:has(.st-key-home_hero) [data-testid="stVerticalBlock"] {
    gap: 0 !important;
}

.stApp:has(.st-key-home_hero) [data-testid="stMain"] {
    padding: 0 !important;
}

.stApp:has(.st-key-home_hero) .st-key-home_hero {
    width: 100% !important;
    min-height: 100vh !important;
    min-height: 100dvh !important;
    margin: 0 !important;
    border-radius: 0 !important;
}
    /* HOME only: isolated editorial full-screen hero */
    .st-key-home_hero {
        position: relative !important;
        box-sizing: border-box !important;
        width: 100% !important;
        min-height: 100vh !important;
        min-height: 100dvh !important;
        padding: clamp(42px, 5vw, 84px) !important;
        background-size: cover !important;
        background-position: center center !important;
        background-repeat: no-repeat !important;
        overflow: hidden !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }
    .st-key-home_hero .home-hero-kicker {
        font-family: 'Manrope', Arial, sans-serif !important;
        font-size: clamp(10px, 0.9vw, 14px) !important;
        font-weight: 400 !important;
        letter-spacing: 0.25em !important;
        line-height: 1.5 !important;
        color: white !important;
        margin: 0 0 clamp(48px, 9vh, 110px) 0 !important;
    }
    .st-key-home_hero h1.home-hero-title,
    .st-key-home_hero h1.home-hero-title span {
        font-family: 'Manrope', 'Helvetica Neue', Arial, sans-serif !important;
        font-size: clamp(48px, 6.6vw, 112px) !important;
        font-weight: 300 !important;
        font-style: normal !important;
        line-height: 1.04 !important;
        letter-spacing: -0.065em !important;
        color: #FFFFFF !important;
        text-transform: uppercase !important;
        text-shadow: 0 2px 8px rgba(0,0,0,0.16) !important;
    }
    .st-key-home_hero h1.home-hero-title {
        margin: 0 !important;
        padding: 0 !important;
        max-width: 100% !important;
    }
    .st-key-home_hero h1.home-hero-title span { display: block !important; }
    .st-key-home_hero .home-hero-subtitle {
        font-family: 'Manrope', Arial, sans-serif !important;
        font-size: clamp(16px, 1.7vw, 27px) !important;
        font-weight: 300 !important;
        line-height: 1.4 !important;
        letter-spacing: -0.02em !important;
        color: white !important;
        margin: 20px 0 0 0 !important;
    }
    .st-key-home_hero .st-key-home_cta {
        position: absolute !important;
        right: clamp(24px, 5vw, 84px) !important;
        bottom: clamp(24px, 4vw, 58px) !important;
        left: auto !important;
        top: auto !important;
        width: auto !important;
        z-index: 10 !important;
    }
    .st-key-home_hero .st-key-home_cta button {
        background: #F8F5EF !important;
        border: 0 !important;
        border-radius: 100px !important;
        color: #332A26 !important;
        padding: 16px 30px !important;
        min-height: 55px !important;
        width: auto !important;
        font-family: 'Manrope', Arial, sans-serif !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        letter-spacing: .13em !important;
        white-space: nowrap !important;
    }
    .st-key-home_hero .st-key-home_cta button:hover {
        background: #E8DDD1 !important;
        color: #332A26 !important;
    }
    @media (max-width: 900px) {
        .st-key-home_hero h1.home-hero-title,
        .st-key-home_hero h1.home-hero-title span {
            font-size: clamp(35px, 7.3vw, 70px) !important;
        }
        .st-key-home_hero { padding: 35px 5% 110px !important; }
    }

    /* DETAIL: photo on the left and white information panel on the right. */
    .st-key-detail_layout [data-testid="stHorizontalBlock"] { align-items: stretch; }
    .st-key-detail_photo, .st-key-detail_photo [data-testid="stImage"],
    .st-key-detail_photo [data-testid="stImage"] > div { height: 100%; }
    .st-key-detail_photo img { min-height: 490px; }
    .st-key-detail_layout .detail-aligned-panel { min-height: 490px; }
    @media (max-width: 768px) {
      .st-key-detail_photo img { min-height: 0; height: 300px !important; }
      .st-key-detail_layout .detail-aligned-panel { min-height: 0; }
    }
    /* DETAIL only: panoramic photo and responsive information panel */
    .st-key-detail_photo img {
        width: 100% !important;
        height: 100% !important; min-height: 490px;
        object-fit: cover !important;
        object-position: center !important;
        display: block !important;
    }
    @media (max-width: 640px) {
        .st-key-detail_photo img { height: 260px !important; }
    }

    /* The panel aligns to the top of the adjacent photo. */
    .detail-aligned-panel {
        box-sizing: border-box;
        width: 100%;
        background: #FFFFFF;
        border: 1px solid #E5E5DD;
        padding: clamp(22px, 3vw, 40px);
        margin: 0 0 16px 0;
        border-radius: 2px;
    }
    .st-key-detail_photo { margin-bottom: 0 !important; }
    .detail-info-item {
        border-top: 1px solid #E5E0D9;
        margin-top: 22px;
        padding-top: 18px;
        color: #796F65;
        font-size: 17px;
        line-height: 1.7;
    }
    .detail-info-item b { display: block; color: #45382F; font-size: 16px; margin-bottom: 6px; }
    .festival-official-link {
        display: block;
        text-align: center;
        text-decoration: none !important;
        color: #45382F !important;
        background: #FFFFFF;
        border: 1px solid #D8C7B2;
        padding: 15px 12px;
        margin-top: 25px;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: .4px;
    }
    .festival-official-link:hover { background: #F7F6F2; }
    .other-festivals-heading { margin-top: 50px; }
    .related-festival-card { min-height: 0; }
    .related-festival-card img { height: 195px; }
    .related-festival-card .festival-card-body { min-height: 175px; box-sizing: border-box; }
    @media (max-width: 900px) {
        .related-festival-card { min-height: 0; }
        .related-festival-card .festival-card-body { min-height: 0; }
    }
    </style>
    """).strip(),
    unsafe_allow_html=True,
)



st.markdown("""
<style>
.st-key-home_hero {
    position: relative !important;
}

.st-key-home_hero .st-key-home_language {
    position: absolute !important;
    top: 22px !important;
    right: 4% !important;
    width: auto !important;
    z-index: 50 !important;
    background: transparent !important;
}

.st-key-home_language [data-testid="stRadio"] label,
.st-key-home_language [data-testid="stRadio"] p,
.st-key-home_language [data-testid="stRadio"] span {
    color: white !important;
    font-weight: 600 !important;
}

.st-key-home_language [data-testid="stRadio"] {
    background: transparent !important;
}

.st-key-inner_language {
    margin-top: 22px !important;
    margin-bottom: 10px !important;
}

.st-key-inner_language [data-testid="stRadio"] {
    display: flex;
    justify-content: flex-end;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------



# LANGUAGE_POSITION_FINAL_START
st.markdown("""
<style>
/* HOME과 내부 페이지의 언어 선택 위치 통일 */
.st-key-home_language,
.st-key-inner_language {
    position: fixed !important;
    top: 12px !important;
    right: 24px !important;
    width: auto !important;
    z-index: 999999 !important;
    margin: 0 !important;
    padding: 0 !important;
    background: transparent !important;
}

/* 언어 선택 가로 정렬 */
.st-key-home_language [role="radiogroup"],
.st-key-inner_language [role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    gap: 12px !important;
}

/* 글씨 디자인 */
.st-key-home_language label,
.st-key-inner_language label {
    font-size: 12px !important;
    font-weight: 600 !important;
}

/* 불필요한 여백 제거 */
.st-key-home_language [data-testid="stRadio"],
.st-key-inner_language [data-testid="stRadio"] {
    margin: 0 !important;
    padding: 0 !important;
}
</style>
""", unsafe_allow_html=True)
# LANGUAGE_POSITION_FINAL_END


# CLEAN_HEADER_START
st.markdown("""
<style>
/* Streamlit 기본 상단 툴바 숨기기 */
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stMainMenu"] {
    display: none !important;
}

/* 헤더 자체는 투명하게 유지 */
header[data-testid="stHeader"] {
    background: transparent !important;
    box-shadow: none !important;
}

/* HOME과 내부 페이지 언어 선택 버튼 */
.st-key-home_language,
.st-key-inner_language {
    position: fixed !important;
    top: 12px !important;
    right: 24px !important;
    z-index: 999999 !important;
    width: auto !important;
    margin: 0 !important;
    padding: 0 !important;
    background: transparent !important;
}

/* 언어 선택 가로 정렬 */
.st-key-home_language [role="radiogroup"],
.st-key-inner_language [role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    gap: 12px !important;
}
</style>
""", unsafe_allow_html=True)
# CLEAN_HEADER_END


# TOPLINE_COLOR_START
st.markdown("""
<style>
.topline .brand {
    color: #646477 !important;
}

.topline > div:last-child {
    color: #7A7772 !important;
}
</style>
""", unsafe_allow_html=True)
# TOPLINE_COLOR_END

# 3. 축제 기본 데이터
# ------------------------------------------------------------
# 엑셀 파일에 일부 축제가 없더라도 네 개 축제를 모두 표시합니다.
# 실제 날짜나 입장료를 확인하지 않은 항목은 임의로 만들지 않습니다.

DEFAULT_FESTIVALS = [
    {
        "축제명": "강릉커피축제",
        "지역": "강릉",
        "주소": "강원특별자치도 강릉시",
        "위도": 37.7519,
        "경도": 128.8761,
        "마커": "☕",
        "테마": "COFFEE & COAST",
        "소개": (
            "향긋한 커피와 바다의 풍경을 함께 즐기는 여행. "
            "강릉의 커피 문화와 감각적인 카페를 발견해 보세요."
        ),
        "방문포인트": "로컬 카페, 커피 문화, 바다 풍경",
        "이미지": (
            "https://images.unsplash.com/photo-1442512595331-e89e73853f31"
            "?auto=format&fit=crop&w=1200&q=85"
        ),
    },
    {
        "축제명": "수원화성문화제",
        "지역": "수원",
        "주소": "경기도 수원시",
        "위도": 37.2636,
        "경도": 127.0286,
        "마커": "🏰",
        "테마": "HERITAGE & CULTURE",
        "소개": (
            "유네스코 세계유산 수원화성을 배경으로 "
            "한국의 역사와 전통문화를 만나는 여행."
        ),
        "방문포인트": "수원화성, 전통문화, 역사 탐방",
        "이미지": (
            "https://images.unsplash.com/photo-1519608487953-e999c86e7455"
            "?auto=format&fit=crop&w=1200&q=85"
        ),
    },
    {
        "축제명": "춘천마임축제",
        "지역": "춘천",
        "주소": "강원특별자치도 춘천시",
        "위도": 37.8813,
        "경도": 127.7298,
        "마커": "🎭",
        "테마": "PERFORMANCE & ART",
        "소개": (
            "몸짓과 움직임으로 만나는 공연예술의 세계. "
            "일상에서 벗어나 색다른 문화 경험을 즐겨 보세요."
        ),
        "방문포인트": "거리공연, 마임, 공연예술",
        "이미지": (
            "https://images.unsplash.com/photo-1503095396549-807759245b35"
            "?auto=format&fit=crop&w=1200&q=85"
        ),
    },
    {
        "축제명": "전주비빔밥축제",
        "지역": "전주",
        "주소": "전북특별자치도 전주시",
        "위도": 35.8242,
        "경도": 127.1480,
        "마커": "🍲",
        "테마": "FOOD & LOCAL FLAVOUR",
        "소개": (
            "한국의 대표적인 음식 문화를 탐험하는 미식 여행. "
            "전주의 맛과 지역의 매력을 함께 발견해 보세요."
        ),
        "방문포인트": "비빔밥, 전주 음식문화, 한옥마을",
        "이미지": (
            "https://images.unsplash.com/photo-1490645935967-10de6ba17061"
            "?auto=format&fit=crop&w=1200&q=85"
        ),
    },
]


# ------------------------------------------------------------
# 4. 엑셀 파일 불러오기
# ------------------------------------------------------------

def find_excel_file():
    """현재 폴더 또는 이 파이썬 파일의 폴더에서 엑셀을 찾습니다."""

    current_folder = Path.cwd()
    app_folder = Path(__file__).resolve().parent

    candidates = [
        current_folder / "festival_data_4개.xlsx",
        app_folder / "festival_data_4개.xlsx",
    ]

    for file_path in candidates:
        if file_path.exists():
            return file_path

    return None


def clean_text(value):
    """빈 값이나 NaN을 안전하게 문자열로 바꿉니다."""

    if pd.isna(value):
        return ""

    return str(value).strip()


def find_column(df, possible_names):
    """엑셀의 열 이름이 조금 달라도 해당 열을 찾습니다."""

    columns = {str(col).strip().lower(): col for col in df.columns}

    for name in possible_names:
        key = name.lower()
        if key in columns:
            return columns[key]

    return None


def load_excel_data():
    """엑셀 파일이 있으면 읽고, 없으면 기본 데이터를 사용합니다."""

    bilingual_path = Path(__file__).resolve().parent / "festival_data_bilingual.xlsx"
    file_path = bilingual_path if bilingual_path.exists() else find_excel_file()

    if file_path is None:
        return pd.DataFrame(), False, None

    try:
        excel_df = pd.read_excel(file_path)
        return excel_df, True, None

    except Exception as error:
        return pd.DataFrame(), False, str(error)


# ------------------------------------------------------------
# 5. 엑셀 데이터를 앱에서 사용할 형식으로 정리
# ------------------------------------------------------------

def prepare_festivals(excel_df):
    """
    기본 축제 네 개를 먼저 준비합니다.
    엑셀에 같은 축제가 있으면 엑셀의 실제 정보를 우선 활용합니다.
    """

    festivals = [item.copy() for item in DEFAULT_FESTIVALS]

    if excel_df.empty:
        return festivals

    name_col = find_column(
        excel_df,
        ["축제명", "행사명", "축제명칭", "title", "name", "event_name"],
    )

    if name_col is None:
        return festivals

    region_col = find_column(
        excel_df,
        ["지역", "시도", "지역명", "시군구", "addr1", "주소", "address"],
    )

    address_col = find_column(
        excel_df,
        ["주소", "주소2", "addr1", "addr2", "address", "venue", "장소"],
    )

    lat_col = find_column(
        excel_df,
        ["위도", "latitude", "lat", "mapy"],
    )

    lon_col = find_column(
        excel_df,
        ["경도", "longitude", "lon", "lng", "mapx"],
    )

    date_col = find_column(
        excel_df,
        [
            "축제기간",
            "행사기간",
            "개최기간",
            "기간",
            "event_period",
            "period",
            "date",
        ],
    )

    image_col = find_column(
        excel_df,
        ["이미지", "이미지url", "이미지주소", "firstimage", "image", "image_url"],
    )

    desc_col = find_column(
        excel_df,
        ["소개", "축제소개", "설명", "개요", "overview", "description", "content"],
    )

    venue_col = find_column(
        excel_df,
        ["장소", "개최장소", "venue"],
    )

    # 엑셀의 각 행을 확인합니다.
    for _, row in excel_df.iterrows():
        excel_name = clean_text(row[name_col])

        if not excel_name:
            continue

        # 네 개의 기본 축제 중 이름이 일치하는 항목을 찾습니다.
        target = None

        for festival in festivals:
            if festival["축제명"] in excel_name or excel_name in festival["축제명"]:
                target = festival
                break

        # 네 개 이외의 축제는 현재 시연용 기본 목록에 추가하지 않습니다.
        if target is None:
            continue

        # 빈 셀은 기본값을 유지하고, 값이 있으면 엑셀 정보를 활용합니다.
        if region_col is not None:
            region = clean_text(row[region_col])

            if region:
                target["지역"] = region

        if address_col is not None:
            address = clean_text(row[address_col])

            if address:
                target["주소"] = address

        if lat_col is not None and lon_col is not None:
            try:
                lat = float(row[lat_col])
                lon = float(row[lon_col])

                # 대한민국의 대략적인 위경도 범위를 벗어난 값은 적용하지 않습니다.
                if 33.0 <= lat <= 39.5 and 124.0 <= lon <= 132.0:
                    target["위도"] = lat
                    target["경도"] = lon
            except (TypeError, ValueError):
                pass

        if date_col is not None:
            period = clean_text(row[date_col])

            if period:
                target["기간"] = period

        if "eventstartdate" in row.index and "eventenddate" in row.index:
            start_date = pd.to_datetime(
                str(row["eventstartdate"]).split(".")[0],
                format="%Y%m%d",
                errors="coerce"
            )
            end_date = pd.to_datetime(
                str(row["eventenddate"]).split(".")[0],
                format="%Y%m%d",
                errors="coerce"
            )

            if pd.notna(start_date) and pd.notna(end_date):
                target["기간"] = (
                    f"{start_date:%Y.%m.%d} – "
                    f"{end_date:%Y.%m.%d}"
                )

        if image_col is not None:
            image_url = clean_text(row[image_col])

            if image_url.startswith(("https://", "http://")):
                target["이미지"] = image_url

        if desc_col is not None:
            description = clean_text(row[desc_col])

            if description:
                target["소개"] = description

        if venue_col is not None:
            venue = clean_text(row[venue_col])

            if venue:
                target["장소"] = venue

    return festivals


# ------------------------------------------------------------
# 6. 데이터 불러오기 및 세션 상태 준비
# ------------------------------------------------------------

excel_df, excel_loaded, excel_error = load_excel_data()
festivals = prepare_festivals(excel_df)


# ============================================================
# LANGUAGE SUPPORT / EN - KO
# ============================================================

if "language" not in st.session_state:
    st.session_state.language = "EN"


def tr(en, ko):
    """현재 선택된 언어에 맞는 문장을 반환합니다."""
    return ko if st.session_state.language == "KO" else en


# 기본 데이터에서 영문 번역이 필요한 항목
EN_EXTRA = {
    "강릉커피축제": {
        "지역": "GANGNEUNG",
        "테마": "COFFEE",
        "방문포인트": "Coffee culture, local cafés and coastal scenery",
    },
    "수원화성문화제": {
        "지역": "SUWON",
        "테마": "HERITAGE",
        "방문포인트": "Royal heritage, fortress architecture and cultural performances",
    },
    "춘천마임축제": {
        "지역": "CHUNCHEON",
        "테마": "PERFORMANCE",
        "방문포인트": "Mime, street performances and creative arts",
    },
    "전주비빔밥축제": {
        "지역": "JEONJU",
        "테마": "FOOD",
        "방문포인트": "Bibimbap, traditional flavors and food culture",
    },
}


def display_field(festival, field):
    """내부 한국어 ID는 유지하고 표시 문구만 번역합니다."""
    if st.session_state.language == "KO":
        if field == "소개":
            return festival.get("description_ko") or festival.get("소개", "")
        if field == "장소":
            return festival.get("venue_ko") or festival.get("장소", "")
        return festival.get(field, "")

    english_columns = {
        "축제명": "name_en",
        "소개": "description_en",
        "장소": "venue_en",
        "주소": "venue_en",
    }

    if field in english_columns:
        value = festival.get(english_columns[field])
        if value:
            return value

    name = festival.get("축제명", "")
    translated = EN_EXTRA.get(name, {}).get(field)

    return translated if translated else festival.get(field, "")


# 기존 앱에서 사용하는 축제 데이터에 한영 열 추가
for festival in festivals:
    matching_rows = excel_df[
        excel_df["title"].astype(str).str.strip() == festival["축제명"]
    ] if "title" in excel_df.columns else []

    if len(matching_rows) > 0:
        row = matching_rows.iloc[0]

        for field in [
            "name_en", "description_ko", "description_en",
            "venue_ko", "venue_en", "official_url"
        ]:
            if field in row.index:
                value = clean_text(row[field])
                if value:
                    festival[field] = value


# HTML 디자인은 유지하고 화면에 보이는 고정 문구만 번역
HTML_TRANSLATIONS = {
    "KOREA · LOCAL CULTURE JOURNAL": "한국 · 지역문화 여행",
    "TRAVEL / CULTURE / PEOPLE": "여행 / 문화 / 사람",
    "DISCOVER</span>": "발견하는</span>",
    "KOREA</span>": "한국</span>",
    "THROUGH FESTIVALS</span>": "축제 여행</span>",
    "Beyond Seoul, a thousand stories await.": "서울 너머, 새로운 한국의 이야기를 만나보세요.",
    "01 / EXPLORE THE MAP": "01 / 축제 지도",
    "Four festivals. One Korea.": "네 개의 축제, 하나의 한국.",
    "Select a marker to discover a different side of Korea.": "마커를 선택해 한국의 다양한 매력을 발견해 보세요.",
    "Each symbol represents the local experience waiting for you.": "각 마커에는 특별한 지역 경험이 담겨 있습니다.",
    "CHOOSE YOUR EXPERIENCE": "나만의 여행 선택",
    "Where to next?": "다음 여행지는 어디인가요?",
    "Choose a festival to see its story and travel highlights.": "축제를 선택하고 지역의 이야기를 만나보세요.",
    "02 / FESTIVAL DETAILS": "02 / 축제 상세 정보",
    "Discover the festival.": "축제를 더 자세히 알아보세요.",
    "Explore local culture, one destination at a time.": "지역마다 다른 한국의 문화를 만나보세요.",
    "03 / KEEP EXPLORING": "03 / 다른 축제 둘러보기",
    "Other festivals to explore.": "함께 둘러볼 다른 축제",
    "Discover another local story in Korea.": "한국의 또 다른 지역 이야기를 발견해 보세요.",
    "Local highlights": "주요 즐길 거리",
    "Destination": "여행 지역",
    "Festival period": "축제 기간",
    "FESTIVAL PERIOD": "축제 기간",
    "Venue": "개최 장소",
    "VENUE": "개최 장소",
    "DISCOVER MORE ABOUT THIS FESTIVAL": "축제 공식 홈페이지",
    "<b>DISCOVER</b>": "<b>주요 즐길 거리</b>",
    "Location:": "주소:",
}


def translate_html(html):
    if st.session_state.language != "KO":
        return html

    for english, korean in HTML_TRANSLATIONS.items():
        html = html.replace(english, korean)

    return html


def localized_markdown(body, *args, **kwargs):
    """기존 HTML/CSS를 유지하며 표시 문구만 변경합니다."""
    if isinstance(body, str) and kwargs.get("unsafe_allow_html"):
        body = translate_html(body)

    return st.markdown(body, *args, **kwargs)



if "selected_festival" not in st.session_state:
    st.session_state.selected_festival = "강릉커피축제"

if "pending_festival" in st.session_state:
    st.session_state.selected_festival = st.session_state.pop("pending_festival")

festival_names = [item["축제명"] for item in festivals]

if st.session_state.selected_festival not in festival_names:
    st.session_state.selected_festival = festival_names[0]

if "page" not in st.session_state:
    st.session_state.page = "home"
if "map_zoom_to_festival" not in st.session_state:
    st.session_state.map_zoom_to_festival = False


def on_festival_change():
    st.session_state.selected_festival = st.session_state.festival_selector_widget
    st.session_state.map_zoom_to_festival = True


def go_to(page):
    st.session_state.page = page



def select_festival(festival_name):
    """카드나 선택 메뉴에서 고른 축제를 저장합니다."""

    st.session_state.pending_festival = festival_name
    st.session_state.map_zoom_to_festival = True


selected_festival = next(
    (
        item
        for item in festivals
        if item["축제명"] == st.session_state.selected_festival
    ),
    festivals[0],
)



# ------------------------------------------------------------
# 7. HOME: 한옥 골목 전체 화면 + 동일한 크기의 영문 제목
# ------------------------------------------------------------

if st.session_state.page == "home":
    hero_photo = Path(__file__).resolve().parent / "hanhok.jpg"
    if hero_photo.exists():
        encoded_photo = base64.b64encode(hero_photo.read_bytes()).decode("ascii")
        hero_background = (
            "linear-gradient(90deg, rgba(20,14,12,.35) 0%, "
            "rgba(20,14,12,.17) 56%, rgba(20,14,12,.24) 100%), "
            f"url('data:image/jpeg;base64,{encoded_photo}')"
        )
    else:
        hero_background = "linear-gradient(115deg, #544038, #9E806B)"
        st.warning("홈 배경 사진을 찾지 못했습니다. 사진 JPG 파일을 festival_app.py와 같은 폴더에 넣어 주세요.")

    # HOME에서만 기존 상단 여백을 줄이고 배경을 넓게 표시합니다.
    localized_markdown(
        "<style>"
        ".block-container {max-width: 100% !important; padding: 0 !important; margin: 0 !important;}"
        "header[data-testid='stHeader'] {background: transparent !important; height: 0 !important;}"
        f".st-key-home_hero {{background-image: {hero_background};}}"
        "</style>",
        unsafe_allow_html=True,
    )
    with st.container(key="home_hero"):
        with st.container(key="home_language"):
            st.radio(
                "Language / 언어",
                ["EN", "KO"],
                horizontal=True,
                key="language",
                label_visibility="collapsed",
            )
        localized_markdown(
            '<div class="home-hero-kicker">KOREA · LOCAL CULTURE JOURNAL</div>'
            '<h1 class="home-hero-title">'
            '<span>DISCOVER</span>'
            '<span>KOREA</span>'
            '<span>THROUGH FESTIVALS</span>'
            '</h1>'
            '<div class="home-hero-subtitle">Beyond Seoul, a thousand stories await.</div>',
            unsafe_allow_html=True,
        )
        with st.container(key="home_cta"):
            if st.button(tr("EXPLORE KOREA ↗", "한국 축제 탐색하기 ↗"), key="home_explore"):
                go_to("map")
                st.rerun()

else:
    with st.container(key="inner_language"):
        st.radio(
            "Language / 언어",
            ["EN", "KO"],
            horizontal=True,
            key="language",
            label_visibility="collapsed",
        )

    # 지도와 상세 화면은 기존 상단 브랜드 디자인을 유지합니다.
    localized_markdown(
        compact_html(textwrap.dedent("""
        <div class="topline">
            <div class="brand">KOREA · LOCAL CULTURE JOURNAL</div>
            <div>TRAVEL / CULTURE / PEOPLE</div>
        </div>
        """).strip()),
        unsafe_allow_html=True,
    )



STATION_IDS = {
    "서울": "NAT010000",
    "용산": "NAT010032",
    "청량리": "NAT130126",
    "수원": "NAT010415",
    "강릉": "NAT601936",
    "남춘천": "NAT140840",
    "춘천": "NAT140873",
    "전주": "NAT040257",
}

STATION_EN = {
    "서울": "Seoul",
    "용산": "Yongsan",
    "청량리": "Cheongnyangni",
    "수원": "Suwon",
    "강릉": "Gangneung",
    "남춘천": "Namchuncheon",
    "춘천": "Chuncheon",
    "전주": "Jeonju",
}

FESTIVAL_STATIONS = {
    "강릉커피축제": ("서울", "강릉"),
    "수원화성문화제": ("서울", "수원"),
    "춘천마임축제": ("청량리", "춘천"),
    "전주비빔밥축제": ("서울", "전주"),
}


TRAIN_ICON_HTML = (
    '<svg xmlns="http://www.w3.org/2000/svg" '
    'width="27" height="27" viewBox="0 0 24 24" '
    'fill="none" stroke="#756554" stroke-width="1.7" '
    'stroke-linecap="round" stroke-linejoin="round" '
    'style="vertical-align:-5px;margin-left:9px;" '
    'aria-hidden="true">'
    '<rect x="5" y="3" width="14" height="16" rx="4"/>'
    '<path d="M5 11h14M8 19l-2 3M16 19l2 3"/>'
    '<path d="M8 15h.01M16 15h.01"/>'
    '</svg>'
)



@st.cache_data(ttl=900, show_spinner=False)
def get_train_schedules(dep_id, arr_id, yyyymmdd):
    load_dotenv(r"C:\Users\user\.env", override=True)
    key = os.getenv("TAGO_API_KEY")
    if not key:
        raise ValueError("TAGO_API_KEY is missing from C:\\Users\\user\\.env")
    url = "https://apis.data.go.kr/1613000/TrainInfo/GetStrtpntAlocFndTrainInfo"
    resp = requests.get(url, params={
        "serviceKey": key, "depPlaceId": dep_id, "arrPlaceId": arr_id,
        "depPlandTime": yyyymmdd, "_type": "json",
        "numOfRows": 100, "pageNo": 1,
    }, timeout=20)
    resp.raise_for_status()
    payload = resp.json()["response"]
    header = payload.get("header", {})
    if str(header.get("resultCode")) != "00":
        raise ValueError(header.get("resultMsg", "Train API error"))
    items = (payload.get("body", {}).get("items") or {}).get("item", [])
    if isinstance(items, dict):
        items = [items]
    rows = []
    for train in items or []:
        try:
            dep_time = datetime.strptime(str(train["depplandtime"]), "%Y%m%d%H%M%S")
            arr_time = datetime.strptime(str(train["arrplandtime"]), "%Y%m%d%H%M%S")
        except (KeyError, TypeError, ValueError):
            continue
        duration = int((arr_time - dep_time).total_seconds() // 60)
        charge = train.get("adultcharge")
        try:
            charge_text = f"{int(charge):,} KRW" if charge is not None else "—"
        except (ValueError, TypeError):
            charge_text = "—"
        rows.append({
            "train": str(train.get("traingradename", "Train")),
            "number": str(train.get("trainno", "")),
            "departure": dep_time.strftime("%H:%M"),
            "arrival": arr_time.strftime("%H:%M"),
            "duration": duration,
            "charge": charge_text,
        })
    return sorted(rows, key=lambda x: x["departure"])


if st.session_state.page == "map":
    nav_home, nav_blank = st.columns([1, 6])
    with nav_home:
        if st.button(tr("← HOME", "← 홈"), key="map_home"):
            go_to("home")
            st.rerun()
    # ------------------------------------------------------------
    # 10. 축제 지도
    # ------------------------------------------------------------

    localized_markdown(
        compact_html(textwrap.dedent("""
        <div class="section-kicker">01 / EXPLORE THE MAP</div>
        <div class="section-title">Four festivals. One Korea.</div>
        <div class="section-desc">
            Select a marker to discover a different side of Korea.
            Each symbol represents the local experience waiting for you.
        </div>
        """).strip()),
        unsafe_allow_html=True,
    )

    # 지도 데이터 준비
    map_df = pd.DataFrame(festivals)

    map_df["위도"] = pd.to_numeric(map_df["위도"], errors="coerce")
    map_df["경도"] = pd.to_numeric(map_df["경도"], errors="coerce")
    map_df = map_df.dropna(subset=["위도", "경도"]).copy()

    map_df["선택됨"] = (
        map_df["축제명"] == st.session_state.selected_festival
    )

    # 위경도에 따른 점 크기 및 색상
    map_df["반지름"] = map_df["선택됨"].apply(
        lambda selected: 10000 if selected else 7000
    )

    map_df["색상"] = map_df["선택됨"].apply(
        lambda selected: [138, 102, 80, 235]
        if selected
        else [200, 181, 160, 210]
    )

    # 선택한 축제로 지도 중심 이동. 첫 방문은 한국 전체.
    map_selected = next(
        (item for item in festivals if item["축제명"] == st.session_state.selected_festival),
        festivals[0],
    )
    if st.session_state.map_zoom_to_festival:
        map_latitude = float(map_selected["위도"])
        map_longitude = float(map_selected["경도"])
        map_zoom = 10
    else:
        map_latitude, map_longitude, map_zoom = 36.25, 127.75, 5.4
    view_state = pdk.ViewState(
        latitude=map_latitude,
        longitude=map_longitude,
        zoom=map_zoom,
        pitch=0,
        bearing=0,
    )


    # 지도 레이어를 만들기 전에 툴팁 데이터를 준비
    festival_by_name = {f["축제명"]: f for f in festivals}

    map_df["tooltip_info"] = map_df["축제명"].apply(
        lambda name: (
            display_field(festival_by_name[name], "축제명")
            + " | "
            + display_field(festival_by_name[name], "지역")
            + " · "
            + display_field(festival_by_name[name], "테마")
        )
    )


    map_df["tooltip_name"] = map_df["축제명"].apply(
        lambda name: display_field(festival_by_name[name], "축제명")
    )
    map_df["tooltip_region"] = map_df["축제명"].apply(
        lambda name: display_field(festival_by_name[name], "지역")
    )
    map_df["tooltip_theme"] = map_df["축제명"].apply(
        lambda name: display_field(festival_by_name[name], "테마")
    )

    # 마커 원형 레이어
    circle_layer = pdk.Layer(
        "ScatterplotLayer",
        data=map_df,
        get_position="[경도, 위도]",
        get_radius="반지름",
        get_fill_color="색상",
        get_line_color=[250, 239, 220, 255],
        line_width_min_pixels=2,
        stroked=True,
        filled=True,
        pickable=True,
        auto_highlight=True,
    )

    # 각 지역의 이모지 마커

    # 각 지역의 이모지 마커

    # 각 축제에 맞는 아이콘 이름 지정
    map_df["아이콘키"] = map_df["마커"].map({
        "☕": "coffee",
        "🏰": "castle",
        "🎭": "theater",
        "🍲": "food",
    })

    # 이모티콘 이미지 주소
    icon_data = {
        "coffee": {
            "url": "https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/2615.png",
            "width": 72,
            "height": 72,
            "anchorY": 36,
        },
        "castle": {
            "url": "https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f3f0.png",
            "width": 72,
            "height": 72,
            "anchorY": 36,
        },
        "theater": {
            "url": "https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f3ad.png",
            "width": 72,
            "height": 72,
            "anchorY": 36,
        },
        "food": {
            "url": "https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f372.png",
            "width": 72,
            "height": 72,
            "anchorY": 36,
        },
    }
    map_df["아이콘정보"] = map_df["아이콘키"].map(icon_data)
    # 이모티콘 이미지 표시 레이어
    emoji_layer = pdk.Layer(
        "IconLayer",
        data=map_df,
        get_icon="아이콘정보",
        get_position="[경도, 위도]",
        get_size=40,
        size_scale=1,
        pickable=True,
        billboard=True,
    )


    # 지도 표시 언어
    map_df["표시지역"] = [
        display_field(f, "지역") for f in festivals
    ]
    map_df["표시축제명"] = [
        display_field(f, "축제명") for f in festivals
    ]
    map_df["표시테마"] = [
        display_field(f, "테마") for f in festivals
    ]

    map_df["display_name"] = map_df["표시축제명"]
    map_df["display_region"] = map_df["표시지역"]
    map_df["display_theme"] = map_df["표시테마"]

    # 축제명 레이블
    label_layer = pdk.Layer(
        "TextLayer",
        data=map_df,
        get_position="[경도, 위도]",
        get_text="표시지역",
        get_size=16,
        get_color=[38, 48, 38, 255],
        get_text_anchor="'middle'",
        get_alignment_baseline="'top'",
        get_pixel_offset=[0, 28],
        font_family="Arial",
        billboard=True,
        pickable=True,
    )

    # 툴팁 전용 데이터: 한국어/영어 표시
    festival_lookup = {
        f["축제명"]: f for f in festivals
    }

    map_df["display_name"] = map_df["축제명"].map(
        lambda name: display_field(festival_lookup[name], "축제명")
    )
    map_df["display_region"] = map_df["축제명"].map(
        lambda name: display_field(festival_lookup[name], "지역")
    )
    map_df["display_theme"] = map_df["축제명"].map(
        lambda name: display_field(festival_lookup[name], "테마")
    )

    deck = pdk.Deck(
        layers=[circle_layer, emoji_layer, label_layer],
        initial_view_state=view_state,
        map_style="https://basemaps.cartocdn.com/gl/positron-gl-style/style.json",
        tooltip={
            "html": (
                "<div style='font-size:14px; font-weight:700; "
                "margin-bottom:5px;'>{tooltip_name}</div>"
                "<div style='font-size:11px; opacity:0.8;'>"
                "{tooltip_region} · {tooltip_theme}</div>"
            ),
            "style": {
                "backgroundColor": "#FFFFFF",
                "color": "#29392E",
                "fontSize": "12px",
                "padding": "12px",
            },
        },
    )

    # Train planner: keep the controls in the left map column.


    map_col, select_col = st.columns([1.55, 0.75], gap="large")

    with map_col:
        st.pydeck_chart(
            deck,
            use_container_width=True,
            height=490,
            key=(f"festival_map_{st.session_state.selected_festival}"
                 if st.session_state.map_zoom_to_festival else "festival_map_all"),
        )

        localized_markdown(
            compact_html(textwrap.dedent("""
            <div class="map-caption">
                MAP GUIDE &nbsp; ☕ Gangneung &nbsp; · &nbsp;
                🏰 Suwon &nbsp; · &nbsp; 🎭 Chuncheon &nbsp; · &nbsp;
                🍲 Jeonju

            </div>
            """).strip()),
            unsafe_allow_html=True,
        )

        # Train search lives immediately below the map guide, not below the right cards.
        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='section-kicker'>{tr('PLAN YOUR JOURNEY', '기차 여행 계획하기')}</div>"
            f"<div class='section-title' style='font-size:32px;'>"
            f"{tr('Find Your Train', '열차 시간표 찾기')}"f"{TRAIN_ICON_HTML}</div>",
            unsafe_allow_html=True,
        )
        current_festival = st.session_state.selected_festival
        if st.session_state.get('train_synced_festival') != current_festival:
            default_dep, default_arr = FESTIVAL_STATIONS.get(current_festival, ('서울', '강릉'))
            st.session_state['train_departure'] = default_dep
            st.session_state['train_arrival'] = default_arr
            st.session_state['train_synced_festival'] = current_festival
            st.session_state.pop('train_search_result', None)

        def station_label(station):
            return STATION_EN[station] if st.session_state.language == 'EN' else station

        with st.form('train_search_form'):
            train_dep_col, train_arr_col = st.columns(2)
            with train_dep_col:
                departure_station = st.selectbox(
                    tr('Departure station', '출발역'), list(STATION_IDS),
                    format_func=station_label, key='train_departure',
                )
            with train_arr_col:
                arrival_station = st.selectbox(
                    tr('Arrival station', '도착역'), list(STATION_IDS),
                    format_func=station_label, key='train_arrival',
                )
            train_date = st.date_input(
                tr('Travel date', '여행 날짜'), value=date.today() + timedelta(days=1),
                min_value=date.today(), key='train_travel_date',
            )
            search_pressed = st.form_submit_button(
                tr('SEARCH TRAINS ↗', '열차 검색하기 ↗'), use_container_width=True,
            )

        if search_pressed:
            if departure_station == arrival_station:
                st.warning(tr('Choose two different stations.', '서로 다른 출발역과 도착역을 선택해 줘.'))
                st.session_state.pop('train_search_result', None)
            else:
                with st.spinner(tr('Searching train schedules…', '열차 시간표 조회 중…')):
                    try:
                        schedules = get_train_schedules(
                            STATION_IDS[departure_station], STATION_IDS[arrival_station],
                            train_date.strftime('%Y%m%d'),
                        )
                        st.session_state['train_search_result'] = {
                            'route': (departure_station, arrival_station),
                            'date': train_date.isoformat(), 'trains': schedules,
                        }
                    except (requests.RequestException, ValueError, KeyError, TypeError) as exc:
                        st.session_state.pop('train_search_result', None)
                        st.error(f"{tr('Train search failed', '열차 조회 실패')}: {exc}")

        result = st.session_state.get('train_search_result')
        if result and result['route'] == (departure_station, arrival_station) and result['date'] == train_date.isoformat():
            trains = result['trains']
            st.markdown(f"**{tr('Train schedules', '열차 시간표')}** · {len(trains)} {tr('results', '개 결과')}")
            if not trains:
                st.info(tr(
                    'No direct trains were returned for this route and date. Try another station or check the booking site for transfers.',
                    '해당 날짜와 구간의 직통 열차가 조회되지 않았어. 다른 역을 선택하거나 예매 사이트에서 환승 경로를 확인해 줘.',
                ))
            else:
                for train in trains[:10]:
                    hours, minutes = divmod(train['duration'], 60)
                    duration_text = (
                        f"{hours}h {minutes}m" if st.session_state.language == 'EN'
                        else f"{hours}시간 {minutes}분" if hours else f"{minutes}분"
                    )
                    with st.container(border=True):
                        st.markdown(f"**{escape(train['train'])}** · {escape(train['number'])}")
                        st.markdown(f"### {train['departure']} → {train['arrival']}")
                        st.caption(f"{tr('Duration', '소요시간')}: {duration_text}  ·  {tr('API fare', 'API 제공 요금')}: {train['charge']}")
                if len(trains) > 10:
                    st.caption(tr(
                        f'Showing the first 10 of {len(trains)} trains.',
                        f'전체 {len(trains)}개 중 출발이 빠른 10개 열차를 표시해.',
                    ))
            st.caption(tr(
                'Schedules and fares are API information, not live seat availability. Confirm on the official booking site.',
                '시간표와 요금은 API 제공 정보이며 실시간 잔여 좌석은 아니야. 공식 예매 사이트에서 확인해 줘.',
            ))
            st.link_button(
                tr('BOOK ON KORAIL ↗', '코레일에서 예매하기 ↗'),
                'https://www.letskorail.com/', use_container_width=True,
            )

    with select_col:
        localized_markdown(
            compact_html(textwrap.dedent("""
            <div class="section-kicker">CHOOSE YOUR EXPERIENCE</div>
            <div class="section-title" style="font-size:32px;">
                Where to next?
            </div>
            <div class="section-desc">
                Choose a festival to see its story and travel highlights.
            </div>
            """).strip()),
            unsafe_allow_html=True,
        )

        st.selectbox(
            "축제 선택",
            options=festival_names,
            format_func=lambda name: display_field(
                next(
                    (item for item in festivals if item["축제명"] == name),
                    {"축제명": name}
                ),
                "축제명"
            ),
            key="festival_selector_widget",
            label_visibility="collapsed",
            on_change=on_festival_change,
        )

        # 선택 메뉴를 바꾼 직후 선택 정보를 다시 계산합니다.
        selected_festival = next(
            (
                item
                for item in festivals
                if item["축제명"] == st.session_state.selected_festival
            ),
            festivals[0],
        )

        st.image(
            get_poster_image(
                selected_festival["축제명"],
                selected_festival["이미지"]
            ),
            use_container_width=True,
        )

        localized_markdown(
            compact_html(textwrap.dedent(f"""
            <div class="card-region">
                {display_field(selected_festival, '지역').upper()} ·
                {display_field(selected_festival, '테마')}
            </div>
            <div class="detail-name">
                {selected_festival['마커']}
                {display_field(selected_festival, '축제명')}
            </div>
            <div class="detail-text">
                {display_field(selected_festival, '소개')}
            </div>
            """).strip()),
            unsafe_allow_html=True,
        )

        localized_markdown("<div style='height:0px'></div>", unsafe_allow_html=True)

        localized_markdown(
            compact_html(textwrap.dedent(f"""
            <div class="detail-text">
                <b>DISCOVER</b><br>
                {display_field(selected_festival, '방문포인트')}
            </div>
            """).strip()),
            unsafe_allow_html=True,
        )

        if selected_festival.get("기간"):
            localized_markdown(
                compact_html(textwrap.dedent(f"""
                <div class="detail-text" style="margin-top:6px;">
                    <b>FESTIVAL PERIOD</b><br>
                    {selected_festival['기간']}
                </div>
                """).strip()),
                unsafe_allow_html=True,
            )

        if selected_festival.get("장소"):
            localized_markdown(
                compact_html(textwrap.dedent(f"""
                <div class="detail-text" style="margin-top:6px;">
                    <b>VENUE</b><br>
                    {display_field(selected_festival, '장소')}
                </div>
                """).strip()),
                unsafe_allow_html=True,
            )

        if selected_festival.get("주소"):
            st.caption(f"{tr('Location', '주소')}: {display_field(selected_festival, '주소')}")

        localized_markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
        if st.button(tr("EXPLORE THIS FESTIVAL ↗", "이 축제 알아보기 ↗"), key="map_explore", use_container_width=True):
            go_to("detail")
            st.rerun()



st.markdown('\n<style>\n.related-festival-card {\n    min-height: 0 !important;\n    overflow: hidden;\n}\n.related-festival-card img {\n    display: block !important;\n    width: 100% !important;\n    height: 350px !important;\n    object-fit: contain !important;\n    object-position: center !important;\n    background: #F7F6F2 !important;\n}\n.related-festival-card .festival-card-body {\n    min-height: 175px;\n    box-sizing: border-box;\n}\n</style>\n', unsafe_allow_html=True)


# ============================================================
# 11. FESTIVAL DETAILS (the map is its own page)
# ============================================================




STATION_IDS = {
    "서울": "NAT010000", "용산": "NAT010032", "청량리": "NAT130126",
    "수원": "NAT010415", "강릉": "NAT601936",
    "남춘천": "NAT140840", "춘천": "NAT140873",
    "전주": "NAT040257",
}
STATION_EN = {
    "서울": "Seoul", "용산": "Yongsan", "청량리": "Cheongnyangni", "수원": "Suwon",
    "강릉": "Gangneung", "남춘천": "Namchuncheon",
    "춘천": "Chuncheon", "전주": "Jeonju",
}
FESTIVAL_STATIONS = {
    "강릉커피축제": ("서울", "강릉"),
    "수원화성문화제": ("서울", "수원"),
    "춘천마임축제": ("청량리", "춘천"),
    "전주비빔밥축제": ("서울", "전주"),
}

@st.dialog("🚆 Plan your train journey")
def train_search_dialog():
    festival_name = st.session_state.selected_festival

    default_dep, default_arr = FESTIVAL_STATIONS.get(
        festival_name, ("서울", "강릉")
    )

    def dialog_station_label(station):
        return (
            STATION_EN[station]
            if st.session_state.language == "EN"
            else station
        )

    station_options = list(STATION_IDS)

    with st.form("detail_train_search_form"):
        dep_col, arr_col = st.columns(2)

        with dep_col:
            dep = st.selectbox(
                tr("Departure", "출발역"),
                station_options,
                index=station_options.index(default_dep),
                format_func=dialog_station_label,
            )

        with arr_col:
            arr = st.selectbox(
                tr("Arrival", "도착역"),
                station_options,
                index=station_options.index(default_arr),
                format_func=dialog_station_label,
            )

        travel_date = st.date_input(
            tr("Travel date", "여행 날짜"),
            value=date.today() + timedelta(days=1),
            min_value=date.today(),
        )

        submitted = st.form_submit_button(
            tr("SEARCH TRAINS ↗", "열차 검색하기 ↗"),
            use_container_width=True,
        )

    if submitted:
        if dep == arr:
            st.warning(
                tr(
                    "Please choose different stations.",
                    "출발역과 도착역을 다르게 선택해 줘.",
                )
            )
            st.session_state.pop("detail_train_result", None)
        else:
            try:
                with st.spinner(
                    tr("Searching trains...", "열차 조회 중...")
                ):
                    trains = get_train_schedules(
                        STATION_IDS[dep],
                        STATION_IDS[arr],
                        travel_date.strftime("%Y%m%d"),
                    )

                st.session_state["detail_train_result"] = {
                    "festival": festival_name,
                    "route": (dep, arr),
                    "date": travel_date.isoformat(),
                    "trains": trains,
                }

            except (
                requests.RequestException,
                ValueError,
                KeyError,
                TypeError,
            ) as exc:
                st.session_state.pop("detail_train_result", None)
                st.error(
                    f"{tr('Train search failed', '열차 조회 실패')}: {exc}"
                )

    result = st.session_state.get("detail_train_result")

    if (
        result
        and result["festival"] == festival_name
        and result["route"] == (dep, arr)
        and result["date"] == travel_date.isoformat()
    ):
        trains = result["trains"]

        st.markdown(
            f"**{tr('Train schedules', '열차 시간표')}** "
            f"· {len(trains)} {tr('results', '개 결과')}"
        )

        if not trains:
            st.info(
                tr(
                    "No direct trains found. Try another route or date.",
                    "직통 열차가 조회되지 않았어. 다른 구간이나 날짜를 선택해 줘.",
                )
            )

        for train in trains[:10]:
            hours, minutes = divmod(train["duration"], 60)

            duration_text = (
                f"{hours}h {minutes}m"
                if st.session_state.language == "EN"
                else f"{hours}시간 {minutes}분"
            )

            with st.container(border=True):
                st.markdown(
                    f"**{escape(train['train'])}** "
                    f"· {escape(train['number'])}"
                )
                st.markdown(
                    f"### {train['departure']} → {train['arrival']}"
                )
                st.caption(
                    f"{tr('Duration', '소요시간')}: {duration_text}"
                    f"  ·  {tr('API fare', 'API 제공 요금')}: "
                    f"{train['charge']}"
                )

        if len(trains) > 10:
            st.caption(
                tr(
                    f"Showing the first 10 of {len(trains)} trains.",
                    f"전체 {len(trains)}개 중 10개 열차를 표시해.",
                )
            )

        st.link_button(
            tr("BOOK ON KORAIL ↗", "코레일에서 예매하기 ↗"),
            "https://www.letskorail.com/",
            use_container_width=True,
        )



# ============================================================

# FESTIVAL DIRECTORY - 상세 페이지 전용 왼쪽 메뉴

if st.session_state.page == "detail":

    st.markdown("""
    <style>
    [data-testid="stSidebar"] {
        min-width: 225px !important;
        max-width: 225px !important;
        width: 225px !important;
        background: #F7F6F2;
    }

    [data-testid="stSidebar"] [data-testid="stButton"] button {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #45382F !important;
        justify-content: flex-start !important;
        text-align: left !important;
        padding: 9px 8px !important;
        font-size: 13px !important;
    }

    [data-testid="stSidebar"] [data-testid="stButton"] button:hover {
        background: #EBE4DA !important;
    }
    </style>
    """, unsafe_allow_html=True)

    sidebar_items = [
        ("강릉커피축제", "coffee", "Coffee · Gangneung", "커피 · 강릉"),
        ("수원화성문화제", "landmark", "Hwaseong · Suwon", "화성 · 수원"),
        ("춘천마임축제", "drama", "Mime · Chuncheon", "마임 · 춘천"),
        ("전주비빔밥축제", "utensils", "Bibimbap · Jeonju", "비빔밥 · 전주"),
    ]

    with st.sidebar:
        st.markdown(
            "<div style='font-size:11px;letter-spacing:1.5px;"
            "color:#8B7361;margin:12px 0 18px;'>"
            "EXPLORE FESTIVALS</div>",
            unsafe_allow_html=True,
        )

        for festival_id, icon_name, name_en, name_ko in sidebar_items:

            active = st.session_state.selected_festival == festival_id

            icon_map = {
                "coffee": ":material/coffee:",
                "landmark": ":material/account_balance:",
                "drama": ":material/theater_comedy:",
                "utensils": ":material/restaurant:",
            }

            if st.button(
                tr(name_en, name_ko),
                icon=icon_map[icon_name],
                key=f"sidebar_festival_{festival_id}",
                use_container_width=True,
                type="primary" if active else "secondary",
            ):
                st.session_state.selected_festival = festival_id
                st.session_state.pop("detail_train_result", None)
                st.rerun()

if st.session_state.page == "detail":


    if st.button(tr("← BACK TO MAP", "← 지도로 돌아가기"), key="detail_back_top"):
        st.session_state.map_zoom_to_festival = False
        go_to("map")
        st.rerun()

    selected_festival = next(
        (item for item in festivals if item["축제명"] == st.session_state.selected_festival),
        festivals[0],
    )

    localized_markdown(
        '<div class="section-kicker">02 / FESTIVAL DETAILS</div>'
        '<div class="section-title">Discover the festival.</div>'
        '<div class="section-desc">Explore local culture, one destination at a time.</div>',
        unsafe_allow_html=True,
    )

    festival_urls = {
        "강릉커피축제": "https://www.coffeefestival.net/",
        "수원화성문화제": "https://shfestival.com/",
        "춘천마임축제": "https://www.mimefestival.com/",
        "전주비빔밥축제": "https://tigertrip.net/",
    }
    festival_url = festival_urls.get(selected_festival["축제명"])
    official_button = (
        f'<a class="festival-official-link" href="{escape(festival_url, quote=True)}" '
        'target="_blank" rel="noopener noreferrer">'
        'DISCOVER MORE ABOUT THIS FESTIVAL ↗</a>'
        if festival_url else ""
    )

    period_html = (
        f'<div class="detail-info-item"><b>Festival period</b><div>{escape(str(selected_festival["기간"]))}</div></div>'
        if selected_festival.get("기간") else ""
    )
    venue_html = (
        f'<div class="detail-info-item"><b>Venue</b><div>{escape(str(display_field(selected_festival, '장소')))}</div></div>'
        if selected_festival.get("장소") else ""
    )
    # The highlights and destination are vertical, as in the earlier detail layout.
    detail_html = f"""
    <div class="detail-aligned-panel">
        <div class="section-kicker">{escape(str(display_field(selected_festival, '지역')).upper())} · {escape(str(display_field(selected_festival, '테마')))}</div>
        <div class="detail-name">{escape(str(selected_festival['마커']))} {escape(str(display_field(selected_festival, '축제명')))}</div>
        <div class="detail-text">{escape(str(display_field(selected_festival, '소개')))}</div>
        <div class="detail-info-item"><b>Local highlights</b><div>{escape(str(display_field(selected_festival, '방문포인트')))}</div></div>
        <div class="detail-info-item"><b>Destination</b><div>{escape(str(display_field(selected_festival, '주소')))}</div></div>
        {period_html}
        {venue_html}
        {official_button}
    </div>
    """
    # Photo and information start on the same horizontal line.
    with st.container(key="detail_layout"):
        photo_col, info_col = st.columns([1.18, 1], gap="medium")
        with photo_col:
            with st.container(key="detail_photo"):
                st.image(selected_festival["이미지"], use_container_width=True)
        with info_col:
            localized_markdown(compact_html(detail_html), unsafe_allow_html=True)




    # Beyond the Festival - TourAPI attraction data
    from urllib.parse import quote
    from html import escape

    nearby_df = pd.read_excel(
        "festival_nearby_attractions_complete.xlsx",
        dtype={"contentid": str},
    )
    nearby_df = nearby_df[
        nearby_df["축제명"] == selected_festival["축제명"]
    ].head(3)

    place_names_en = {
        "안목해변": "Anmok Beach",
        "경포호": "Gyeongpo Lake",
        "오죽헌": "Ojukheon",
        "화성행궁": "Hwaseong Haenggung",
        "화홍문": "Hwahongmun Gate",
        "수원시립미술관": "Suwon Museum of Art",
        "소양강스카이워크": "Soyanggang Skywalk",
        "삼악산 호수케이블카": "Samaksan Lake Cable Car",
        "레고랜드": "LEGOLAND Korea Resort",
        "전주한옥마을": "Jeonju Hanok Village",
        "경기전": "Gyeonggijeon Shrine",
        "남부시장": "Jeonju Nambu Market",
    }

    # 여행자에게 도움이 되도록 설명 보강
    descriptions = {
        "안목해변": (
            "Discover Gangneung's famous seaside coffee street. "
            "Enjoy ocean views while exploring cafés along the beach.",
            "강릉의 대표적인 커피거리와 바다를 함께 즐길 수 있는 곳입니다. "
            "해변을 따라 자리한 카페에서 커피를 마시며 바다 풍경을 감상해 보세요."
        ),
        "경포호": (
            "Take a relaxing walk beside one of Gangneung's best-known lakes. "
            "The peaceful scenery makes it a lovely stop during your trip.",
            "강릉을 대표하는 호수 중 하나로, 여유롭게 산책하기 좋은 곳입니다. "
            "잔잔한 호수와 주변 자연 풍경을 감상하며 쉬어 가보세요."
        ),
        "오죽헌": (
            "Visit the historic birthplace associated with Shin Saimdang "
            "and the scholar Yi I. Explore traditional architecture and Korean history.",
            "신사임당과 율곡 이이의 역사로 잘 알려진 문화유산입니다. "
            "전통 건축을 감상하며 조선시대의 역사와 문화를 만나보세요."
        ),
        "화성행궁": (
            "Explore the royal temporary palace inside Suwon Hwaseong. "
            "Its historic courtyards and buildings reveal stories of the Joseon dynasty.",
            "수원화성 안에 자리한 조선시대의 행궁입니다. "
            "아름다운 전각과 마당을 둘러보며 왕실의 역사와 문화를 경험해 보세요."
        ),
        "화홍문": (
            "Admire the elegant northern watergate of Suwon Hwaseong. "
            "Its stone arches and waterside scenery make it a memorable photo stop.",
            "수원화성의 아름다운 북쪽 수문입니다. "
            "석조 아치와 물길이 어우러진 풍경을 감상하고 사진을 남겨보세요."
        ),
        "수원시립미술관": (
            "Discover contemporary art close to Hwaseong Haenggung. "
            "Add a creative stop to your historical journey through Suwon.",
            "화성행궁 가까이에 위치한 현대미술 공간입니다. "
            "수원의 역사 명소를 둘러본 뒤 다양한 미술 전시를 감상해 보세요."
        ),
        "소양강스카이워크": (
            "Step onto a glass walkway overlooking the Soyang River. "
            "Enjoy wide river views and a unique sightseeing experience.",
            "소양강 위로 뻗은 유리 바닥 전망길을 걸어볼 수 있습니다. "
            "탁 트인 강의 풍경과 특별한 전망 체험을 즐겨보세요."
        ),
        "삼악산 호수케이블카": (
            "Ride a cable car above Chuncheon's beautiful landscape. "
            "Enjoy sweeping views of the lake, mountains, and surrounding city.",
            "춘천의 호수와 산을 한눈에 조망할 수 있는 케이블카입니다. "
            "공중에서 펼쳐지는 자연 풍경과 도시 전망을 감상해 보세요."
        ),
        "레고랜드": (
            "Explore a colorful LEGO-themed amusement park in Chuncheon. "
            "Discover imaginative attractions and family-friendly experiences.",
            "춘천에 위치한 레고 테마의 놀이공원입니다. "
            "다채로운 놀이시설과 레고로 꾸며진 공간에서 가족 여행을 즐겨보세요."
        ),
        "전주한옥마을": (
            "Wander through streets lined with traditional Korean houses. "
            "Experience Jeonju's architecture, local food, and cultural atmosphere.",
            "전통 한옥이 모여 있는 전주의 대표 관광지입니다. "
            "골목을 거닐며 한옥 건축과 전주의 음식, 문화적 분위기를 느껴보세요."
        ),
        "경기전": (
            "Discover a historic shrine dedicated to King Taejo, "
            "the founder of the Joseon dynasty. Explore its traditional buildings and grounds.",
            "조선을 건국한 태조의 어진을 모신 역사적인 장소입니다. "
            "전통 건축과 고즈넉한 경내를 둘러보며 조선 왕조의 역사를 만나보세요."
        ),
        "남부시장": (
            "Explore one of Jeonju's traditional markets. "
            "Discover local food, lively stalls, and the city's everyday culture.",
            "전주의 먹거리와 지역 문화를 경험할 수 있는 전통시장입니다. "
            "다양한 음식과 상점들을 둘러보며 활기찬 시장 분위기를 느껴보세요."
        ),
    }

    st.markdown("""
    <style>
    .nearby-section-heading {
        margin-top: 58px;
        margin-bottom: 26px;
    }
    .nearby-card {
        background: #FCFAF6;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 18px rgba(69,56,47,.06);
        display: flex;
        flex-direction: column;
    }
    .nearby-photo {
        width: 100%;
        height: 215px;
        object-fit: cover;
        display: block;
    }
    .nearby-card-content {
        padding: 19px;
        display: flex;
        flex-direction: column;
        flex: 1;
    }
    .nearby-card-name {
        font-size: 19px;
        font-weight: 650;
        color: #45382F;
        line-height: 1.4;
        margin-bottom: 12px;
    }
    .nearby-card-description {
        font-size: 15px;
        line-height: 1.7;
        color: #655C54;
        min-height: 0;
        margin-bottom: 18px;
    }
    .nearby-card-address {
        font-size: 12px;
        color: #8C8176;
        line-height: 1.6;
        margin-top: 12px;
        min-height: 40px;
    }
    .nearby-card-link {
        display: block;
        text-align: center;
        background: #EAE7DC;
        color: #45382F !important;
        border-radius: 8px;
        padding: 13px 10px;
        margin-top: 12px;
        text-decoration: none !important;
        font-size: 13px;
        font-weight: 650;
    }
    .nearby-card-link:hover {
        background: #DDD7C8;
    }
    </style>
    """, unsafe_allow_html=True)

    title = escape(tr("Beyond the Festival", "축제 너머의 여행"))
    subtitle = escape(tr(
        "Discover more around your destination.",
        "축제 주변의 매력적인 관광지도 만나보세요."
    ))
    kicker = escape(tr("EXPLORE NEARBY", "주변 관광지 탐색"))

    # 기존 Discover the festival.과 동일한 제목 스타일
    st.markdown(
        '<div class="nearby-section-heading">'
        f'<div class="section-kicker">{kicker}</div>'
        f'<div class="section-title">{title}</div>'
        f'<div class="section-desc">{subtitle}</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    nearby_cols = st.columns(3, gap="medium")

    for index, (_, place) in enumerate(nearby_df.iterrows()):
        name_ko = str(place["표시이름"])
        name_en = place_names_en.get(name_ko, name_ko)
        desc_en, desc_ko = descriptions.get(name_ko, ("", ""))

        address = place.get("주소", "")
        address = "" if pd.isna(address) else str(address)

        image_url = place.get("이미지URL", "")
        image_url = (
            str(image_url)
            if pd.notna(image_url) and str(image_url).startswith("https://")
            else ""
        )

        # 지도 검색에는 한국어 장소명을 사용하되 화면에는 노출하지 않음
        map_url = (
            "https://www.google.com/maps/search/?api=1&query="
            + quote(name_ko + " " + address)
        )

        photo_html = (
            f'<img class="nearby-photo" '
            f'src="{escape(image_url, quote=True)}" '
            f'alt="{escape(tr(name_en, name_ko), quote=True)}">'
            if image_url else
            '<div class="nearby-photo" style="background:#EAE7DC;"></div>'
        )

        # 영어 모드에서는 한국어 주소를 숨김
        address_html = (
            f'<div class="nearby-card-address">📍 {escape(address)}</div>'
            if tr(False, True) and address else
            ''
        )

        card_html = f"""
        <div class="nearby-card">
            {photo_html}
            <div class="nearby-card-content">
                <div class="nearby-card-name">
                    {escape(tr(name_en, name_ko))}
                </div>
                <div class="nearby-card-description">
                    {escape(tr(desc_en, desc_ko))}
                </div>
                {address_html}
                <a class="nearby-card-link"
                   href="{escape(map_url, quote=True)}"
                   target="_blank" rel="noopener noreferrer">
                    {escape(tr("VIEW LOCATION ↗", "위치 보기 ↗"))}
                </a>
            </div>
        </div>
        """

        with nearby_cols[index]:
            st.html(card_html)

    st.caption(tr(
        "Attraction information: Korea Tourism Organization TourAPI.",
        "관광지 정보 출처: 한국관광공사 TourAPI."
    ))


    # Getting There - after nearby attractions
    st.markdown(
        "<div style='height:20px'></div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        f"<div class='section-kicker'>"
        f"{tr('PLAN YOUR JOURNEY', '여행 계획하기')}</div>"
        f"<div class='section-title'>"
        f"{tr('Getting There', '축제까지 가는 길')}"f"{TRAIN_ICON_HTML}</div>"
        f"<div class='section-desc'>"
        f"{tr('Found somewhere you want to visit? Check the trains for your journey.', '가고 싶은 곳을 찾았다면 열차 시간표를 확인해 보세요.')}"
        f"</div>",
        unsafe_allow_html=True,
    )

    if st.button(
        tr("FIND TRAINS ↗", "열차 시간표 확인하기 ↗"),
        key="detail_getting_there_button",
        use_container_width=True,
    ):
        st.session_state.pop("detail_train_result", None)
        train_search_dialog()


    # ------------------------------------------------------------
    # Festival atmosphere - subtle line illustration
    # ------------------------------------------------------------

    festival_art = {
        "강릉커피축제": """
            <path d="M90 80 H195 V122 Q195 155 143 155 Q90 155 90 122 Z"/>
            <path d="M195 91 H213 Q238 91 238 112 Q238 135 195 135"/>
            <path d="M115 65 Q104 52 117 38"/>
            <path d="M148 65 Q137 52 150 38"/>
            <path d="M180 65 Q169 52 182 38"/>
            <path d="M60 180 Q95 165 130 180 T200 180 T270 180"/>
            <path d="M60 199 Q95 184 130 199 T200 199 T270 199"/>
        """,
        "수원화성문화제": """
            <path d="M45 173 H290"/>
            <path d="M65 173 V118 H270 V173"/>
            <path d="M80 118 L110 87 H225 L255 118 Z"/>
            <path d="M106 87 L126 67 H207 L227 87"/>
            <path d="M146 173 V133 H190 V173"/>
            <path d="M60 150 H95 M240 150 H275"/>
            <path d="M40 190 Q100 182 165 190 T295 190"/>
        """,
        "춘천마임축제": """
            <circle cx="170" cy="60" r="24"/>
            <path d="M170 85 Q155 120 172 151"/>
            <path d="M165 110 Q125 92 99 112"/>
            <path d="M169 112 Q208 88 235 102"/>
            <path d="M172 151 L126 198"/>
            <path d="M172 151 L222 194"/>
            <path d="M75 65 Q100 45 125 65"/>
            <path d="M230 55 Q255 35 280 55"/>
            <path d="M82 220 Q165 208 255 220"/>
        """,
        "전주비빔밥축제": """
            <path d="M50 118 L165 48 L280 118"/>
            <path d="M77 118 H253 V183 H77 Z"/>
            <path d="M104 183 V125 H226 V183"/>
            <path d="M165 125 V183"/>
            <path d="M90 211 Q165 246 240 211"/>
            <path d="M90 211 H240"/>
            <path d="M130 204 Q140 188 150 203"/>
            <path d="M163 204 Q173 188 183 203"/>
            <path d="M194 204 Q204 188 214 203"/>
        """,
    }

    art_paths = festival_art.get(
        selected_festival["축제명"], ""
    )









    # FESTIVAL BACKGROUND START
    import base64

    festival_keys = {
        "강릉커피축제": "gangneung",
        "수원화성문화제": "suwon",
        "춘천마임축제": "chuncheon",
        "전주비빔밥축제": "jeonju",
    }

    festival_key = festival_keys.get(
        selected_festival["축제명"]
    )

    if festival_key:
        image_folder = Path("festival_backgrounds_split")

        sidebar_path = image_folder / f"{festival_key}_sidebar.png"
        main_path = image_folder / f"{festival_key}_main.png"

        sidebar_encoded = base64.b64encode(
            sidebar_path.read_bytes()
        ).decode("ascii")

        main_encoded = base64.b64encode(
            main_path.read_bytes()
        ).decode("ascii")

        st.html(
            f"""
            <style>
            /* 사이드바: 축제를 대표하는 부분 */
            [data-testid="stSidebar"] {{
                background-color: #F7F6F2 !important;
                background-image: url("data:image/png;base64,{sidebar_encoded}") !important;
                background-repeat: no-repeat !important;
                background-size: 420px auto !important;
                background-position: center bottom !important;
            }}

            [data-testid="stSidebar"] > div {{
                background: transparent !important;
            }}

            /* 본문: 대표 그림을 제외한 나머지 풍경 */
            [data-testid="stMainBlockContainer"] {{
                position: relative;
            }}

            [data-testid="stMainBlockContainer"]::after {{
                content: "";
                position: absolute;
                bottom: 0;
                left: 50%;
                transform: translateX(-50%);
                width: 100vw;
                height: 310px;
                background-image: url("data:image/png;base64,{main_encoded}");
                background-position: center bottom;
                background-size: 100% auto;
                background-repeat: no-repeat;
                pointer-events: none;
                opacity: 0.65;
                z-index: 1;
                -webkit-mask-image: linear-gradient(
                    to bottom,
                    transparent 0%,
                    black 35%,
                    black 100%
                );
                mask-image: linear-gradient(
                    to bottom,
                    transparent 0%,
                    black 35%,
                    black 100%
                );
            }}
            </style>
            """
        )
    # FESTIVAL BACKGROUND END



    if st.button(tr("← BACK TO MAP", "← 지도로 돌아가기"), key="detail_back"):
        st.session_state.map_zoom_to_festival = False
        go_to("map")
        st.rerun()

    # ------------------------------------------------------------
    # 13. 데이터 상태 확인
    # ------------------------------------------------------------

    with st.expander("프로젝트 데이터 상태 확인"):
        if excel_loaded:
            st.success("festival_data_4개.xlsx 파일을 불러왔습니다.")
            st.write(f"엑셀 원본 행 수: {len(excel_df)}")
            st.write(f"엑셀 열 이름: {', '.join(map(str, excel_df.columns))}")
            st.caption(
                "기본 축제 네 개를 기준으로 이름이 일치하는 행의 정보를 "
                "적용합니다. 좌표와 기간은 엑셀의 열 이름 및 값 형식에 "
                "따라 적용되지 않을 수도 있습니다."
            )
        else:
            st.warning(
                "엑셀 파일을 찾지 못해 기본 시연용 데이터를 사용 중입니다."
            )

            if excel_error:
                st.caption(f"엑셀 오류: {excel_error}")

            st.caption(
                "festival_app.py와 festival_data_4개.xlsx를 "
                "같은 폴더에 두면 엑셀 데이터를 읽을 수 있습니다."
            )

        st.write(
            pd.DataFrame(
                [
                    {
                        "축제": item["축제명"],
                        "지역": item["지역"],
                        "위도": item["위도"],
                        "경도": item["경도"],
                    }
                    for item in festivals
                ]
            )
        )


    # ------------------------------------------------------------
    # 14. 하단
    # ------------------------------------------------------------

    localized_markdown(
        compact_html(textwrap.dedent("""
        <div class="footer-line">
            DISCOVER KOREA THROUGH FESTIVALS
            <br><br>
            A LOCAL CULTURE & TRAVEL DISCOVERY PROJECT
        </div>
        """).strip()),
        unsafe_allow_html=True,
    )
