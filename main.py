"""
과일 최저가 시즌 알리미
------------------------------------------------------------
한국농수산식품유통공사(aT) "연월별 도,소매가격정보" 공공데이터 API(B552845/perYearMonth/price)를
이용해 과일류(부류코드 400번대) 가격을 조회하고, 어느 달에 가장 저렴한지 보여주는 Streamlit 앱.

API 요청 파라미터 / 응답 필드명, 부류·품목 코드는 함께 제공된 `api_des.xlsx`, `price_code.xlsx`
문서를 그대로 참조해 작성했습니다. (이 두 파일은 배포 시 main.py와 같은 폴더에 두어야
과일 품목 코드 조회가 정상 동작합니다.)

- 서비스 URL: https://apis.data.go.kr/B552845/perYearMonth/price
- 필수 파라미터: serviceKey, returnType, pageNo, numOfRows,
  cond[exmn_ym::GTE](조사연월 시작), cond[exmn_ym::LTE](조사연월 끝)
- 선택 파라미터: cond[ctgry_cd::EQ](부류코드), cond[item_cd::EQ](품목코드) 등
- 응답 item 필드: exmn_ym, se_cd/se_nm(구분: 01 소매, 02 중도매 …),
  ctgry_cd/ctgry_nm, item_cd/item_nm, vrty_cd/vrty_nm, grd_cd/grd_nm,
  unit/unit_sz, pmm_avgprc(월별평균가), pmm_hgprc(월별최고가), pmm_lwprc(월별최저가) 등

API 호출이 실패하거나(키 미등록, 네트워크 오류 등) 응답이 비어 있으면 화면이 비지 않도록
계절성을 반영한 데모 데이터로 자동 대체되며, 이 경우 화면 상단에 안내 배너가 표시됩니다.

Price Status Card의 2줄 분석 문구는 Upstage Solar(모델: solar-open2, OpenAI 호환 API)로
생성합니다. 키는 secrets의 SOLAR_API_KEY에서 불러오며, 응답 속도를 위해 reasoning_effort는
"none"으로 꺼둡니다. Solar 호출이 실패하면(키 미등록 등) 규칙 기반 문구로 자동 대체됩니다.
"""

import html
import json
import math
import os
import random
import re
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

try:
    from openai import OpenAI  # Solar(solar-open2) 호출용. 없어도 앱은 규칙 기반 문구로 정상 동작.
except ImportError:
    OpenAI = None

# ----------------------------------------------------------------------------
# 기본 설정
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="과일 최저가 시즌 알리미",
    page_icon="🍑",
    layout="wide",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICE_CODE_PATH = os.path.join(BASE_DIR, "price_code.xlsx")

# api_des.xlsx 명세 기준
API_ENDPOINT = "https://apis.data.go.kr/B552845/perYearMonth/price"
FRUIT_CATEGORY_CODE = "400"  # 부류코드 400번대 = 과일류

# Upstage Solar (OpenAI 호환 API) — Price Status Card의 2줄 분석 문구 생성에 사용
SOLAR_BASE_URL = "https://api.upstage.ai/v1"
SOLAR_MODEL = "solar-open2"  # 모델명 그대로 사용

# se_cd(구분코드): 01 소매, 02 중도매, 03/07 친환경농산물 — 01→소매, 02→도매로 표시.
# 자몽·블루베리처럼 유통량이 적어 01/02 데이터가 비어 있는 품목도 있을 수 있어, 03/07(친환경)도
# 함께 인식해 두면 "도매/소매 구분"은 없어도 평균가 그래프(막대)로는 표시할 수 있다.
DIVISION_LABELS = {"01": "소매", "02": "도매", "03": "친환경", "07": "친환경"}

# 화면 상단 큰 버튼에 노출할 16개 과일 (요청 순서 그대로)
MAIN_FRUITS = [
    "사과", "배", "복숭아", "포도", "감귤", "단감", "바나나", "참다래",
    "파인애플", "오렌지", "자몽", "레몬", "체리", "망고", "블루베리", "아보카도",
    "수박", "토마토", "딸기", "참외",
]

# 대표 이미지 대용 이모지 + 데모(대체) 데이터 생성을 위한 계절성 파라미터
# cheap_month: 1~12월 중 평년 기준으로 가장 저렴해지는 달(일반적으로 알려진 제철 기준 추정치)
# 실 API가 살아있으면 이 값은 데모 폴백에만 쓰이고, 실제 화면은 API 데이터로 계산됩니다.
FRUIT_INFO = {
    "사과":     {"emoji": "🍎", "base": 3200, "amp": 900,  "cheap_month": 11},
    "배":       {"emoji": "🍐", "base": 3500, "amp": 1000, "cheap_month": 10},
    "복숭아":   {"emoji": "🍑", "base": 4200, "amp": 1600, "cheap_month": 8},
    "포도":     {"emoji": "🍇", "base": 4500, "amp": 1400, "cheap_month": 9},
    "감귤":     {"emoji": "🍊", "base": 2600, "amp": 900,  "cheap_month": 1},
    "단감":     {"emoji": "🍊", "base": 3400, "amp": 900,  "cheap_month": 11},  # 유니코드에 감(persimmon) 전용 이모지가 없어 둥근 주황색 과일 이모지로 대체 (보름달처럼 보이던 원형 이모지 수정)
    "바나나":   {"emoji": "🍌", "base": 2200, "amp": 300,  "cheap_month": 6},
    "참다래":   {"emoji": "🥝", "base": 3800, "amp": 700,  "cheap_month": 12},
    "파인애플": {"emoji": "🍍", "base": 3600, "amp": 500,  "cheap_month": 7},
    "오렌지":   {"emoji": "🍊", "base": 3300, "amp": 700,  "cheap_month": 2},
    "자몽":     {"emoji": "🍊", "base": 2900, "amp": 600,  "cheap_month": 1},  # 유니코드에 자몽 전용 이모지가 없어 감귤류 이모지로 대체 (멜론 이모지 오류 수정)
    "레몬":     {"emoji": "🍋", "base": 3100, "amp": 500,  "cheap_month": 4},
    "체리":     {"emoji": "🍒", "base": 8500, "amp": 3000, "cheap_month": 6},
    "망고":     {"emoji": "🥭", "base": 5200, "amp": 1500, "cheap_month": 7},
    "블루베리": {"emoji": "🫐", "base": 6000, "amp": 1800, "cheap_month": 6},
    "아보카도": {"emoji": "🥑", "base": 3300, "amp": 500,  "cheap_month": 9},
    "수박":     {"emoji": "🍉", "base": 22000, "amp": 8000, "cheap_month": 8},
    "토마토":   {"emoji": "🍅", "base": 3500, "amp": 1200, "cheap_month": 7},
    "딸기":     {"emoji": "🍓", "base": 9000, "amp": 3500, "cheap_month": 1},
    "참외":     {"emoji": "🍋", "base": 4200, "amp": 1600, "cheap_month": 6},  # 유니코드에 참외 전용 이모지가 없어, 초록빛 멜론(🍈)보다 참외의 노란 타원 형태에 더 가까운 레몬 이모지로 대체
}

TAGS = [
    ("#지금_가장_싼_과일", "__CHEAPEST_NOW__"),
    ("#7월_제철_복숭아", "복숭아"),
    ("#수박_최저가", "수박"),
]

FALLBACK_ITEM_CODES = {
    # price_code.xlsx 로딩 실패 시를 대비한 최소 폴백 (ctgry_cd, item_cd)
    "사과": ("400", "411"), "배": ("400", "412"), "복숭아": ("400", "413"),
    "포도": ("400", "414"), "감귤": ("400", "415"), "단감": ("400", "416"),
    "바나나": ("400", "418"), "참다래": ("400", "419"), "파인애플": ("400", "420"),
    "오렌지": ("400", "421"), "자몽": ("400", "423"), "레몬": ("400", "424"),
    "체리": ("400", "425"), "망고": ("400", "428"), "블루베리": ("400", "429"),
    "아보카도": ("400", "430"), "수박": ("200", "221"),
    "토마토": ("200", "225"), "딸기": ("200", "226"), "참외": ("200", "222"),
}


@st.cache_data(ttl=None, show_spinner=False)
def load_item_codes() -> dict:
    """price_code.xlsx의 '품목코드' 시트에서 과일 이름 -> (부류코드, 품목코드) 매핑을 만든다."""
    try:
        items = pd.read_excel(PRICE_CODE_PATH, sheet_name="품목코드")
        items.columns = ["ctgry_cd", "item_cd", "item_nm"]
        lookup = {}
        for _, row in items.iterrows():
            name = str(row["item_nm"]).strip()
            if name not in lookup:  # 첫 매칭만 사용
                lookup[name] = (str(int(row["ctgry_cd"])), str(int(row["item_cd"])))
        return lookup
    except Exception:
        return dict(FALLBACK_ITEM_CODES)


ITEM_CODE_LOOKUP = load_item_codes()
for _name in FRUIT_INFO:
    if _name not in ITEM_CODE_LOOKUP and _name in FALLBACK_ITEM_CODES:
        ITEM_CODE_LOOKUP[_name] = FALLBACK_ITEM_CODES[_name]


@st.cache_data(ttl=None, show_spinner=False)
def load_variety_names() -> dict:
    """price_code.xlsx의 '품종코드' 시트에서 (부류코드, 품목코드) -> 품종명 목록을 만든다."""
    try:
        v = pd.read_excel(PRICE_CODE_PATH, sheet_name="품종코드")
        v.columns = ["ctgry_cd", "item_cd", "vrty_cd", "vrty_nm"]
        result = {}
        for (ctgry_cd, item_cd), grp in v.groupby(["ctgry_cd", "item_cd"]):
            seen, uniq = set(), []
            for n in grp["vrty_nm"]:
                name = str(n).strip()
                if name and name not in seen:
                    seen.add(name)
                    uniq.append(name)
            result[(str(int(ctgry_cd)), str(int(item_cd)))] = uniq
        return result
    except Exception:
        return {}


VARIETY_LOOKUP = load_variety_names()

# 참고용 대표 산지 정보 (일반적으로 널리 알려진 주산지 기준 요약, 공공데이터 API 응답에는
# 포함되지 않는 항목이라 별도로 정리했습니다.)
PRODUCTION_REGIONS = {
    "사과": "경북 청송·안동, 충북 충주 등",
    "배": "충남 천안, 전남 나주 등",
    "복숭아": "충북 음성·충주, 경북 청도 등",
    "포도": "경북 상주·김천, 경기 안성 등 (샤인머스캣은 경남 거창 등)",
    "감귤": "제주",
    "단감": "경남 창원(진영), 전남 나주 등",
    "바나나": "대부분 수입(필리핀 등), 국내는 제주·전남 일부",
    "참다래": "전남 해남·경남 등 (그린키위는 뉴질랜드 수입)",
    "파인애플": "대부분 수입(필리핀 등)",
    "오렌지": "수입(미국 캘리포니아, 호주 등)",
    "자몽": "수입(미국, 남아공 등)",
    "레몬": "수입(미국, 칠레 등)",
    "체리": "수입(미국 워싱턴 등)",
    "망고": "제주·전남 일부 국내산, 대부분 수입(태국·베트남 등)",
    "블루베리": "전북 고창, 전남 등",
    "아보카도": "수입(멕시코·페루 등)",
    "수박": "충북 음성, 전남 고창, 경북 고령 등",
    "토마토": "경남 진주, 부산 등",
    "딸기": "충남 논산, 경남 진주·밀양 등",
    "참외": "경북 성주 등",
}

# 스마트 구매팁 - "좋은 과일 고르는 법" 폴백 문구 (Solar 호출이 실패했을 때만 사용).
# 일반적으로 널리 알려진 신선도 판별 기준을 정리한 것으로, 과장 없이 사실 기반으로 작성했습니다.
FRUIT_SELECTION_TIPS = {
    "사과": "꼭지 주변이 마르지 않고 껍질에 윤기가 돌며, 같은 크기라면 묵직한 것이 수분과 당도가 높습니다.",
    "배": "표면이 매끈하고 은은한 향이 나며, 꼭지 반대편을 눌렀을 때 살짝 탄력이 있는 것이 좋습니다.",
    "복숭아": "붉은빛이 고르게 퍼지고 향이 진하며, 살짝 눌렀을 때 탄력 있게 들어가는 것이 잘 익은 것입니다.",
    "포도": "알이 촘촘히 붙어 있고 표면에 하얀 과분이 남아 있으며, 줄기가 마르지 않은 것이 신선합니다.",
    "감귤": "껍질이 얇고 탄력 있으며, 같은 크기 대비 무거운 것이 즙이 많습니다.",
    "단감": "표면에 상처가 없고 꼭지가 마르지 않았으며, 묵직하고 단단한 것이 좋습니다.",
    "바나나": "껍질에 갈색 반점(슈가스팟)이 있으면 당도가 오른 상태이고, 초록빛이 많이 남아 있으면 며칠 후숙이 필요합니다.",
    "참다래": "손으로 살짝 눌렀을 때 약간 탄력 있게 들어가는 것이 잘 익은 상태이며, 딱딱하면 상온에서 며칠 후숙하세요.",
    "파인애플": "꼭지(잎) 색이 진한 초록이고 향이 은은하게 나며, 밑동을 눌렀을 때 탄력이 있는 것이 좋습니다.",
    "오렌지": "껍질이 매끈하고 탄력 있으며, 같은 크기 대비 무거운 것이 즙이 많습니다.",
    "자몽": "껍질이 얇고 매끈하며 묵직한 것이 즙이 많고, 약간 납작한 모양이 단맛이 강한 편입니다.",
    "레몬": "껍질에 윤기가 있고 단단하며, 무거운 것이 즙이 많습니다.",
    "체리": "꼭지가 초록빛으로 싱싱하고 알에 윤기가 있으며, 주름이나 무른 부분이 없는 것을 고르세요.",
    "망고": "손으로 살짝 눌렀을 때 탄력 있게 들어가고 향이 진하게 나는 것이 잘 익은 상태입니다.",
    "블루베리": "표면에 하얀 과분이 남아 있고 알이 통통하며, 뭉개지거나 물기가 도는 것은 피하세요.",
    "아보카도": "꼭지 부분을 눌렀을 때 살짝 들어가면서 탄력이 있으면 먹기 좋은 상태이고, 단단하면 며칠 후숙하세요.",
    "수박": "두드렸을 때 통통한 소리가 나고, 배꼽(꽃이 떨어진 자리)이 작고 균일한 것이 잘 익은 것입니다.",
    "토마토": "꼭지가 싱싱한 초록색이고 껍질에 윤기가 돌며, 묵직한 것이 과육이 알찹니다.",
    "딸기": "꼭지 바로 아래까지 붉은빛이 고르게 퍼져 있고 윤기가 나는 것이 당도가 높습니다.",
    "참외": "표면 골이 선명하고 향이 진하며, 꼭지 부분이 마르지 않은 것이 신선합니다.",
}
GENERIC_SELECTION_TIP = (
    "표면에 상처나 무른 부분이 없고, 같은 크기 대비 묵직한 것을 고르면 대체로 신선하고 당도가 높습니다."
)
CHANNEL_TIP_DEFAULT_TEMPLATE = (
    "제철 성수기에는 대형마트·재래시장의 특가 물량이 가격 대비 품질이 좋고, 상품성 좋은 것을 "
    "원한다면 {regions} 등 산지 직거래(로컬푸드매장·온라인 산지직송)를 이용하면 신선도와 "
    "가격 모두 유리한 경우가 많습니다."
)


def _format_rich_text(text: str) -> str:
    """마크다운 **볼드**와 줄바꿈만 지원하는 최소 HTML 이스케이프 변환."""
    escaped = html.escape(text)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped)
    return escaped.replace("\n", "<br>")

# ----------------------------------------------------------------------------
# 스타일
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        /* 참고 이미지(핀테크 앱)의 톤앤매너: 화이트 카드 + 은은한 핑크 배경 + 살몬 코럴 포인트 */
        --accent: #f2938a;
        --accent-dark: #e0665a;
        --ink: #241f2b;
        --muted: #9b93a3;
        --chip-mint: #c8f2e3;
        --chip-mint-ink: #1f9a72;
        --chip-lavender: #ded6fb;
        --chip-lavender-ink: #7a68d8;
        --chip-orange: #ffe0bd;
        --chip-orange-ink: #e08a2e;
    }
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(180deg, #fdf1f0 0%, #fbe8e6 100%);
    }
    [data-testid="stHeader"] {
        background: rgba(0,0,0,0);
    }
    .main-copy {
        font-size: 2.2rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        text-align: center;
        margin-bottom: 0.3rem;
        color: var(--ink);
    }
    .sub-copy {
        text-align: center;
        color: var(--muted);
        margin-bottom: 1.4rem;
        font-size: 1rem;
    }
    /* 기본(선택되지 않은) 버튼: 화이트 필, 다크 네이비 텍스트 — 참고 이미지의 화이트 카드 톤 */
    .stButton > button {
        border-radius: 999px;
        border: 1px solid rgba(36,31,43,0.06);
        background-color: #ffffff;
        color: var(--ink);
        font-weight: 600;
        padding: 0.6rem 0.4rem;
        box-shadow: 0 4px 14px rgba(36,31,43,0.06);
        transition: all 0.15s ease-in-out;
    }
    .stButton > button:hover {
        border-color: var(--accent);
        color: var(--accent-dark);
    }
    /* 선택된(primary) 버튼: 참고 이미지의 "Get Started" 버튼처럼 단색 살몬 필 + 화이트 텍스트 */
    .stButton > button[kind="primary"] {
        border-radius: 999px;
        border: none;
        background: var(--accent);
        color: #ffffff;
        font-weight: 700;
        box-shadow: 0 8px 18px rgba(242,147,138,0.4);
    }
    .stButton > button[kind="primary"]:hover {
        background: var(--accent-dark);
        color: #ffffff;
    }
    /* 해시태그는 버튼이 아니라 클릭 가능한 작은 글자로 보이도록 버튼 모양을 전부 제거한다.
       (st.container(key="tag_row")로 실제 DOM 부모를 만들어야 .st-key-tag_row 스코프가 먹는다 —
       예전에는 st.markdown으로 연 <div>가 형제 요소일 뿐이라 아래 규칙이 전혀 적용되지 않았다.) */
    .st-key-tag_row div[data-testid="stHorizontalBlock"] {
        gap: 0.3rem !important;
        flex-wrap: wrap;
    }
    .st-key-tag_row div[data-testid="stColumn"] {
        width: auto !important;
        flex: 0 0 auto !important;
        min-width: 0 !important;
    }
    .st-key-tag_row .stButton {
        width: auto !important;
    }
    .st-key-tag_row .stButton > button,
    .st-key-tag_row .stButton > button:focus,
    .st-key-tag_row .stButton > button:active,
    .st-key-tag_row .stButton > button:focus:not(:active) {
        all: unset;
        display: inline-block;
        cursor: pointer;
        font-size: 0.78rem;
        font-weight: 600;
        line-height: 1.4;
        color: var(--accent-dark);
        padding: 0.1rem 0.3rem;
        white-space: nowrap;
    }
    .st-key-tag_row .stButton > button:hover {
        color: var(--ink);
        text-decoration: underline;
    }
    .status-card {
        border-radius: 28px;
        min-height: 260px;
        padding: 2.2rem 2.4rem;
        display: flex;
        align-items: center;
        gap: 1.8rem;
        background: #ffffff;
        border: none;
        box-shadow: 0 12px 28px rgba(36, 31, 43, 0.08);
        margin-bottom: 1rem;
    }
    .status-card .emoji {
        font-size: 5.5rem;
        line-height: 1;
        flex-shrink: 0;
        filter: drop-shadow(0 6px 10px rgba(36,31,43,0.12));
    }
    .status-card .title {
        font-size: 1.6rem;
        font-weight: 800;
        color: var(--ink);
        margin-bottom: 0.5rem;
    }
    .status-card .analysis-line {
        margin-top: 0.6rem;
        line-height: 1.6;
        color: var(--ink);
    }
    .status-card .extra-info {
        margin-top: 0.7rem;
        font-size: 0.88rem;
        color: var(--muted);
    }
    /* 좁은 화면: 이모지가 고정 크기로 자리를 많이 차지해 글자가 좁은 칸에 억지로
       줄바꿈되는 것을 완화한다. 이모지/여백을 줄여 글자가 쓸 수 있는 폭을 넓힌다. */
    @media (max-width: 600px) {
        .status-card {
            padding: 1.4rem 1.3rem;
            gap: 1rem;
            min-height: 0;
        }
        .status-card .emoji {
            font-size: 3.2rem;
        }
        .status-card .title {
            font-size: 1.25rem;
            margin-bottom: 0.35rem;
        }
        .status-card .analysis-line {
            font-size: 0.92rem;
            line-height: 1.5;
        }
        .status-card .extra-info {
            font-size: 0.82rem;
        }
    }
    .light {
        display: inline-block;
        width: 16px;
        height: 16px;
        border-radius: 50%;
        margin-right: 6px;
        vertical-align: middle;
    }
    .light-on { box-shadow: 0 0 10px currentColor; }
    h4, h5, .stMarkdown h4, .stMarkdown h5 {
        color: var(--ink) !important;
        font-weight: 800 !important;
    }
    /* 과일 버튼 그리드: 한 줄에 5개씩 촘촘하게, 버튼 사이 간격은 좁게 */
    .st-key-fruit_grid div[data-testid="stHorizontalBlock"] {
        gap: 0.5rem !important;
    }
    .st-key-fruit_grid .stButton > button {
        padding: 0.6rem 0.3rem;
        font-size: 0.92rem;
        white-space: nowrap;
    }
    /* 좁은 화면(모바일)에서는 한 줄에 3개씩 접혀서 세로 스크롤을 줄인다 */
    @media (max-width: 700px) {
        .st-key-fruit_grid div[data-testid="stHorizontalBlock"] {
            flex-wrap: wrap;
            gap: 0.5rem !important;
        }
        .st-key-fruit_grid div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            flex: 1 1 30% !important;
            width: 30% !important;
            min-width: 30% !important;
        }
        .st-key-fruit_grid .stButton > button {
            padding: 0.5rem 0.2rem;
            font-size: 0.85rem;
        }
    }
    /* 가격표 커스텀 HTML 테이블 (가운데 정렬) */
    .price-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.92rem;
    }
    .price-table th, .price-table td {
        text-align: center;
        padding: 0.5rem 0.6rem;
        border-bottom: 1px solid rgba(0,0,0,0.06);
    }
    .price-table th {
        color: var(--muted);
        font-weight: 700;
        border-bottom: 2px solid rgba(0,0,0,0.1);
    }
    .price-table tr.event-row td {
        background-color: rgba(242,147,138,0.1);
        font-weight: 700;
    }
    /* 스마트 구매팁 카드: 참고 이미지의 컬러풀한 카테고리 타일처럼, 카드마다 다른
       파스텔 색 아이콘 뱃지를 얹는다 (민트/라벤더/오렌지). */
    .tip-card {
        background: #ffffff;
        border: none;
        border-radius: 22px;
        padding: 1.2rem 1.3rem;
        height: 100%;
        min-height: 172px;
        box-shadow: 0 10px 22px rgba(36,31,43,0.07);
        margin-bottom: 0.6rem;
    }
    .tip-card-emoji {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 2.4rem;
        height: 2.4rem;
        border-radius: 50%;
        font-size: 1.25rem;
        margin-bottom: 0.55rem;
    }
    .tip-card-emoji.mint { background: var(--chip-mint); }
    .tip-card-emoji.lavender { background: var(--chip-lavender); }
    .tip-card-emoji.orange { background: var(--chip-orange); }
    .tip-card-title {
        font-weight: 800;
        font-size: 1.02rem;
        color: var(--ink);
        margin-bottom: 0.45rem;
    }
    .tip-card-text {
        font-size: 0.9rem;
        line-height: 1.6;
        color: var(--muted);
    }
    @media (max-width: 700px) {
        .st-key-tip_cards div[data-testid="stHorizontalBlock"] {
            flex-wrap: wrap;
        }
        .st-key-tip_cards div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            flex: 1 1 100% !important;
            width: 100% !important;
            min-width: 100% !important;
        }
    }
    /* 가격 추이 그래프(Plotly)는 폰트 크기가 Python 쪽에서 픽셀 단위로 고정되어 있어,
       좁은 화면에서는 라벨끼리 겹쳐 보인다. Plotly는 SVG로 렌더링되고 글자 크기가
       인라인 style로 박혀 있는데, 외부 스타일시트의 !important는 CSS 명세상 일반
       인라인 style보다 우선하므로 아래 규칙으로 미디어쿼리에 따라 강제로 축소할 수 있다. */
    @media (max-width: 600px) {
        .js-plotly-plot .xtick text,
        .js-plotly-plot .ytick text {
            font-size: 8px !important;
        }
        .js-plotly-plot .legendtext {
            font-size: 9px !important;
        }
        .js-plotly-plot .scatterlayer text,
        .js-plotly-plot .barlayer text {
            font-size: 9px !important;
        }
        .js-plotly-plot .annotation-text {
            font-size: 10px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# 데이터 조회
# ----------------------------------------------------------------------------

def _extract_items(payload: dict):
    """data.go.kr 표준 응답 포맷(response.body.items.item)에서 리스트를 뽑아낸다."""
    try:
        body = payload["response"]["body"]
        items = body.get("items")
        if items is None:
            return []
        if isinstance(items, dict):
            item = items.get("item", [])
            if isinstance(item, dict):
                return [item]
            return item or []
        if isinstance(items, list):
            return items
    except (KeyError, TypeError):
        pass
    return []


def _call_api(api_key: str, ctgry_cd: str, item_cd: str, start_ym: str, end_ym: str) -> list:
    rows = []
    page_no = 1
    num_of_rows = 1000
    while True:
        params = {
            "serviceKey": api_key,
            "returnType": "json",
            "pageNo": page_no,
            "numOfRows": num_of_rows,
            "cond[exmn_ym::GTE]": start_ym,
            "cond[exmn_ym::LTE]": end_ym,
            "cond[ctgry_cd::EQ]": ctgry_cd,
            "cond[item_cd::EQ]": item_cd,
        }
        resp = requests.get(API_ENDPOINT, params=params, timeout=10)
        resp.raise_for_status()
        items = _extract_items(resp.json())
        if not items:
            break
        rows.extend(items)
        if len(items) < num_of_rows or page_no > 10:
            break
        page_no += 1
    return rows


def _parse_rows(raw_rows: list) -> list:
    parsed = []
    for row in raw_rows:
        se_cd = str(row.get("se_cd", "")).strip()
        division = DIVISION_LABELS.get(se_cd)
        if division is None:
            continue  # 친환경 등 그 외 구분은 이번 화면에서는 제외
        exmn_ym = str(row.get("exmn_ym", "")).strip()
        price_raw = row.get("pmm_avgprc")
        item_nm = row.get("item_nm")
        if len(exmn_ym) != 6 or price_raw in (None, "", "-"):
            continue
        try:
            price = float(str(price_raw).replace(",", "").strip())
        except ValueError:
            continue
        if price <= 0:
            continue
        parsed.append(
            {
                "year": int(exmn_ym[:4]),
                "month": int(exmn_ym[4:6]),
                "item_name": str(item_nm).strip() if item_nm else "",
                "division": division,
                "price": price,
            }
        )
    return parsed


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_fruit_series(fruit_name: str, start_ym: str, end_ym: str) -> tuple[pd.DataFrame, bool]:
    """선택된 과일 1종의 연월별 도/소매 가격을 조회한다.

    반환값: (DataFrame, is_live) — is_live가 False이면 API 호출/파싱 실패로
    계절성 기반 데모 데이터를 사용했다는 뜻.
    """
    ctgry_cd, item_cd = ITEM_CODE_LOOKUP.get(fruit_name, FALLBACK_ITEM_CODES.get(fruit_name, ("400", "")))
    try:
        api_key = st.secrets["FRUITS_API_KEY"]
        if not item_cd:
            raise ValueError("item_cd not found")
        raw_rows = _call_api(api_key, ctgry_cd, item_cd, start_ym, end_ym)
        parsed = _parse_rows(raw_rows)
        if not parsed:
            return _build_demo_dataframe(fruit_name, start_ym, end_ym), False
        return pd.DataFrame(parsed), True
    except Exception:
        return _build_demo_dataframe(fruit_name, start_ym, end_ym), False


def _month_range(start_ym: str, end_ym: str):
    sy, sm = int(start_ym[:4]), int(start_ym[4:])
    ey, em = int(end_ym[:4]), int(end_ym[4:])
    out = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        out.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def _build_demo_dataframe(fruit_name: str, start_ym: str, end_ym: str) -> pd.DataFrame:
    """실 API 연동이 어려운 경우를 대비한, 계절성을 반영한 대체(데모) 데이터."""
    months = _month_range(start_ym, end_ym)
    info = FRUIT_INFO.get(fruit_name, {"base": 4000, "amp": 1000, "cheap_month": 8})
    rnd = random.Random(fruit_name)  # 과일별로 고정된 패턴 유지
    records = []
    for (y, m) in months:
        phase = (m - info["cheap_month"]) % 12
        seasonal = -math.cos(2 * math.pi * phase / 12)  # cheap_month에서 최솟값
        noise = rnd.uniform(-0.04, 0.04)
        retail = info["base"] + info["amp"] * seasonal + info["base"] * noise
        wholesale = retail * rnd.uniform(0.58, 0.68)
        records.append({"year": y, "month": m, "item_name": fruit_name, "division": "소매", "price": round(retail, -1)})
        records.append({"year": y, "month": m, "item_name": fruit_name, "division": "도매", "price": round(wholesale, -1)})
    return pd.DataFrame(records)


# ----------------------------------------------------------------------------
# 날짜 범위 계산 (오늘로부터 최근 12개월)
# ----------------------------------------------------------------------------
today = date.today()
end_year, end_month = today.year, today.month
start_year, start_month = end_year - 1, end_month
if start_month == 0:
    start_month = 12
    start_year -= 1
START_YM = f"{start_year}{start_month:02d}"
END_YM = f"{end_year}{end_month:02d}"

# ----------------------------------------------------------------------------
# 세션 상태
# ----------------------------------------------------------------------------
if "selected_fruit" not in st.session_state:
    st.session_state.selected_fruit = "복숭아"


def select_fruit(name: str):
    st.session_state.selected_fruit = name


@st.cache_data(ttl=3600, show_spinner=False)
def generate_price_commentary(
    fruit: str,
    cheapest_year: int,
    cheapest_month: int,
    latest_year: int,
    latest_month: int,
    cur_price: float,
    position_pct: float,
    light_label: str,
) -> str | None:
    """Solar(solar-open2)로 가격 상태에 대한 2줄 이내 자연어 분석을 생성한다.

    SOLAR_API_KEY가 없거나 호출이 실패하면 None을 반환하며, 호출부에서 규칙 기반
    문구로 대체한다.
    """
    try:
        api_key = st.secrets["SOLAR_API_KEY"]
    except Exception:
        return None

    if OpenAI is None:  # openai 패키지가 설치되지 않은 환경 (requirements.txt 미반영 등)
        return None

    try:
        client = OpenAI(api_key=api_key, base_url=SOLAR_BASE_URL)
        prompt = (
            f"너는 과일 가격 데이터를 설명해주는 쇼핑 도우미야. 아래 데이터를 바탕으로 "
            f"소비자에게 도움이 되는 분석을 정확히 두 줄 이내 한국어 문장으로 작성해줘. "
            f"불릿 기호나 따옴표 없이 문장만 출력해.\n\n"
            f"- 과일: {fruit}\n"
            f"- 최근 1년 중 가장 저렴한 달: {cheapest_year}년 {cheapest_month}월\n"
            f"- 가장 최근 집계월: {latest_year}년 {latest_month}월, 평균가 약 {cur_price:,.0f}원\n"
            f"- 현재 가격은 연중 가격대의 {position_pct:.0f}% 지점이며 상태는 '{light_label}'\n\n"
            f"첫 줄에는 가장 저렴한 달 정보를, 둘째 줄에는 지금 사도 될지에 대한 조언을 담아줘."
        )
        resp = client.chat.completions.create(
            model=SOLAR_MODEL,
            reasoning_effort="none",
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.choices[0].message.content
        return text.strip() if text else None
    except Exception:
        return None


@st.cache_data(ttl=86400, show_spinner=False)
def generate_polished_copy(
    main_copy: str, sub_copy: str, tag1: str, tag2: str, tag3: str, source_note: str
) -> dict | None:
    """Solar(solar-open2)로 페이지 제목/소제목/해시태그/출처 문구를 신뢰감 있게 다듬는다.

    실패하면 None을 반환하고, 호출부에서 원래 문구를 그대로 사용한다. 하루 단위로만
    다시 생성하도록 캐싱해 불필요한 API 호출을 피한다.
    """
    try:
        api_key = st.secrets["SOLAR_API_KEY"]
    except Exception:
        return None
    if OpenAI is None:
        return None

    try:
        client = OpenAI(api_key=api_key, base_url=SOLAR_BASE_URL)
        payload = {
            "main_copy": main_copy,
            "sub_copy": sub_copy,
            "tag1": tag1,
            "tag2": tag2,
            "tag3": tag3,
            "source_note": source_note,
        }
        prompt = (
            "너는 과일 가격 정보 웹앱의 카피라이터야. 아래 JSON에 담긴 UI 문구들을 "
            "의미와 정보량은 그대로 유지하면서, 더 신뢰감 있고 자연스러운 한국어 문장으로 "
            "다듬어줘. main_copy와 sub_copy는 한 줄 문장이어야 하고, tag1~tag3는 "
            "'#'로 시작하고 단어 사이는 '_'로 잇는 해시태그 형식을 유지해야 해. "
            "source_note는 데이터 출처를 설명하는 짧은 문구야. "
            "다른 설명 없이 동일한 키를 가진 JSON 객체 하나만 출력해.\n\n"
            f"{json.dumps(payload, ensure_ascii=False)}"
        )
        resp = client.chat.completions.create(
            model=SOLAR_MODEL,
            reasoning_effort="none",
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.choices[0].message.content or ""
        match = re.search(r"\{.*\}", text, re.S)
        if not match:
            return None
        data = json.loads(match.group(0))
        required_keys = {"main_copy", "sub_copy", "tag1", "tag2", "tag3", "source_note"}
        if not required_keys.issubset(data.keys()):
            return None
        if not all(isinstance(data[k], str) and data[k].strip() for k in required_keys):
            return None
        return data
    except Exception:
        return None


@st.cache_data(ttl=86400, show_spinner=False)
def generate_smart_buying_tips(fruit: str, varieties: tuple, regions: str) -> dict | None:
    """Solar(solar-open2)로 "스마트 구매팁" 3항목(품종별 구매팁/고르는 법/구매 채널)을 생성한다.

    과장되거나 근거 없는 표현 없이 사실 기반으로만 작성하도록 프롬프트에 명시했다.
    실패하면 None을 반환하며, 호출부에서 FRUIT_SELECTION_TIPS 등 규칙 기반 문구로 대체한다.
    하루 1회만 다시 생성하도록 캐싱해 과일당 반복 호출을 피한다.
    """
    try:
        api_key = st.secrets["SOLAR_API_KEY"]
    except Exception:
        return None
    if OpenAI is None:
        return None

    try:
        client = OpenAI(api_key=api_key, base_url=SOLAR_BASE_URL)
        variety_text = ", ".join(varieties) if varieties else "품종 정보 없음"
        prompt = (
            "너는 과일 구매를 도와주는 신뢰할 수 있는 쇼핑 가이드야. 아래 과일에 대해 "
            "실제로 도움이 되는 정확한 정보만 담아 세 항목을 각각 1~3줄 이내 한국어 문장으로 "
            "작성해줘. 과장되거나 근거 없는 표현, 불릿 기호는 쓰지 마.\n\n"
            f"- 과일: {fruit}\n"
            f"- 주요 품종: {variety_text}\n"
            f"- 주산지: {regions}\n\n"
            "1) variety_tip: 품종별로 맛·식감·구매 시기가 어떻게 다른지\n"
            "2) selection_tip: 좋은 품질의 이 과일을 고르는 실용적인 방법(색, 무게, 향, 촉감 등 "
            "구체적 기준)\n"
            "3) channel_tip: 대형마트/재래시장/산지직송(온라인) 중 이 과일을 사기에 유리한 채널과 "
            "그 이유\n\n"
            '다른 설명 없이 정확히 {"variety_tip": "...", "selection_tip": "...", '
            '"channel_tip": "..."} 형태의 JSON 객체 하나만 출력해.'
        )
        resp = client.chat.completions.create(
            model=SOLAR_MODEL,
            reasoning_effort="none",
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.choices[0].message.content or ""
        match = re.search(r"\{.*\}", text, re.S)
        if not match:
            return None
        data = json.loads(match.group(0))
        required_keys = {"variety_tip", "selection_tip", "channel_tip"}
        if not required_keys.issubset(data.keys()):
            return None
        if not all(isinstance(data[k], str) and data[k].strip() for k in required_keys):
            return None
        return data
    except Exception:
        return None


def _chronological_monthly(series: pd.DataFrame, start_ym: str, end_ym: str) -> pd.DataFrame:
    """year+month로 집계한 뒤, 실제 달력 순서(연도 경계를 넘어가도 올바르게)로 정렬한다.

    조회 구간에 포함된 모든 달(현재 달 포함)에 대해 자리(row)를 만들어 두므로, 아직
    집계되지 않은 최신 달이 있어도 x축에서 "지금"의 위치(가장 오른쪽)가 밀리지 않는다.
    데이터가 없는 달은 price가 NaN으로 남는다.
    """
    monthly = series.groupby(["year", "month"])["price"].mean().reset_index()
    full_periods = _month_range(start_ym, end_ym)
    full_df = pd.DataFrame(full_periods, columns=["year", "month"])
    merged = full_df.merge(monthly, on=["year", "month"], how="left")
    return merged


def _latest_with_data(monthly: pd.DataFrame) -> pd.Series | None:
    """price가 채워진 것 중 달력 순서상 가장 마지막(가장 최근) 행을 반환한다."""
    valid = monthly.dropna(subset=["price"])
    if valid.empty:
        return None
    return valid.iloc[-1]


@st.cache_data(ttl=3600, show_spinner=False)
def cheapest_fruit_this_month(start_ym: str, end_ym: str) -> str:
    """가장 최근 집계월 기준, 연중 최저가 대비 가장 저렴해 보이는 과일을 하나 골라준다."""
    best, best_ratio = MAIN_FRUITS[0], math.inf
    for f in MAIN_FRUITS:
        series, _ = fetch_fruit_series(f, start_ym, end_ym)
        if series.empty:
            continue
        monthly = _chronological_monthly(series, start_ym, end_ym)
        latest = _latest_with_data(monthly)
        if latest is None:
            continue
        cur_price = latest["price"]  # 데이터가 있는 것 중 가장 최근 집계월
        year_min = monthly["price"].min()
        if year_min:
            ratio = cur_price / year_min
            if ratio < best_ratio:
                best, best_ratio = f, ratio
    return best


# ----------------------------------------------------------------------------
# 상단: 메인 카피 + 태그 (Solar가 문구를 다듬어주고, 실패 시 기본 문구 사용)
# ----------------------------------------------------------------------------
DEFAULT_MAIN_COPY = "좋아하는 과일, 언제 사야 가장 싸고 맛있을까?"
DEFAULT_SUB_COPY = "아래에서 과일을 골라보세요. 연간 가격 추이로 최적의 구매 시기를 알려드려요."
DEFAULT_TAG_LABELS = ["#지금_가장_싼_과일", "#7월_제철_복숭아", "#수박_최저가"]
DEFAULT_SOURCE_NOTE = "출처: 한국농수산식품유통공사 연월별 도,소매가격정보"

_polished = generate_polished_copy(
    DEFAULT_MAIN_COPY, DEFAULT_SUB_COPY, *DEFAULT_TAG_LABELS, DEFAULT_SOURCE_NOTE
)
if _polished:
    MAIN_COPY = _polished["main_copy"]
    SUB_COPY = _polished["sub_copy"]
    TAG_LABELS = [_polished["tag1"], _polished["tag2"], _polished["tag3"]]
    SOURCE_NOTE = _polished["source_note"]
else:
    MAIN_COPY, SUB_COPY, TAG_LABELS, SOURCE_NOTE = (
        DEFAULT_MAIN_COPY, DEFAULT_SUB_COPY, DEFAULT_TAG_LABELS, DEFAULT_SOURCE_NOTE
    )

st.markdown(f'<div class="main-copy">{html.escape(MAIN_COPY)}</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sub-copy">{html.escape(SUB_COPY)}</div>', unsafe_allow_html=True)

with st.container(key="tag_row"):
    tag_cols = st.columns(len(TAGS) + 6)
    for i, (_, target) in enumerate(TAGS):
        with tag_cols[i]:
            if st.button(TAG_LABELS[i], key=f"tag_{i}"):
                if target == "__CHEAPEST_NOW__":
                    select_fruit(cheapest_fruit_this_month(START_YM, END_YM))
                else:
                    select_fruit(target)

st.write("")

# 과일 버튼 그리드. 데스크톱은 5열이지만, 모바일 폭에서는 CSS로 한 줄에 3개씩 접히도록 해
# 버튼 20개를 세로로 쭉 스크롤하지 않아도 되게 만든다 (아래 st-key-fruit_grid 미디어쿼리 참고).
with st.container(key="fruit_grid"):
    cols_per_row = 5
    for row_start in range(0, len(MAIN_FRUITS), cols_per_row):
        row_fruits = MAIN_FRUITS[row_start: row_start + cols_per_row]
        cols = st.columns(cols_per_row)
        for col, fruit in zip(cols, row_fruits):
            with col:
                emoji = FRUIT_INFO.get(fruit, {}).get("emoji", "🍏")
                is_selected = fruit == st.session_state.selected_fruit
                label = f"✓ {emoji} {fruit}" if is_selected else f"{emoji} {fruit}"
                if st.button(
                    label,
                    key=f"fruit_{fruit}",
                    use_container_width=True,
                    type="primary" if is_selected else "secondary",
                ):
                    select_fruit(fruit)

st.divider()

# ----------------------------------------------------------------------------
# Price Status Card
# ----------------------------------------------------------------------------
selected = st.session_state.selected_fruit
fruit_df, is_live = fetch_fruit_series(selected, START_YM, END_YM)

if not is_live:
    st.info(
        f"⚠️ '{selected}'의 공공데이터 API 응답을 확인하지 못해 계절성을 반영한 데모 데이터로 표시하고 있어요. "
        "`FRUITS_API_KEY`가 secrets에 등록되어 있는지, 인증키 활용 승인이 완료됐는지 확인해 주세요.",
        icon="ℹ️",
    )

overall_monthly = None if fruit_df.empty else _chronological_monthly(fruit_df, START_YM, END_YM)
latest = None if overall_monthly is None else _latest_with_data(overall_monthly)

if fruit_df.empty or latest is None:
    st.error(f"'{selected}'에 대한 데이터가 없어요. 다른 과일을 선택해보세요.")
else:
    # 연도 경계를 넘나드는 12개월 구간이므로 월(1~12) 숫자만으로 정렬/그룹핑하면 안 되고,
    # 반드시 연도까지 함께 고려해 달력 순서대로 다뤄야 한다. 조회 구간의 모든 달(현재 달 포함)에
    # 자리를 만들어 두므로, x축 상 "지금"의 위치는 데이터 발표 여부와 무관하게 항상 맨 오른쪽이 된다.
    cheapest_pos = int(overall_monthly["price"].idxmin())
    cheapest_row = overall_monthly.loc[cheapest_pos]
    cheapest_month = int(cheapest_row["month"])
    cheapest_year = int(cheapest_row["year"])
    year_min = overall_monthly["price"].min()
    year_max = overall_monthly["price"].max()

    latest_month = int(latest["month"])
    latest_year = int(latest["year"])
    cur_price = latest["price"]

    # 최신 데이터가 "지금"으로부터 몇 달이나 떨어져 있는지 확인한다. 감(단감)처럼 특정 시기에만
    # 유통되는 품목은 출하기가 끝나면 몇 달씩 가격 데이터 자체가 없는데, 그 마지막 데이터(주로
    # 출하기=저렴한 시기)만 보고 "지금이 사기 좋은 때"라고 판단하면 실제로는 유통이 끊긴 비수기를
    # 저가 시즌으로 잘못 안내하게 된다. 이를 막기 위해 최신 데이터와 "지금" 사이 격차가 크면
    # 별도의 "출하 시기 아님" 상태로 전환한다.
    STALE_GAP_THRESHOLD = 2  # 이만큼(개월) 이상 최신 데이터가 없으면 비수기로 간주
    latest_pos = int(latest.name)
    current_pos_ref = len(overall_monthly) - 1
    gap_months = current_pos_ref - latest_pos
    is_stale = gap_months >= STALE_GAP_THRESHOLD

    span = max(year_max - year_min, 1e-6)
    position = (cur_price - year_min) / span  # 0(가장 쌈) ~ 1(가장 비쌈)

    if is_stale:
        light_color, light_label = "#9ca3af", "지금은 출하 시기가 아니에요"
        status_suffix = f"(최신 데이터: {latest_year}년 {latest_month}월 기준)"
    elif position <= 0.33:
        light_color, light_label = "#22c55e", "지금이 딱 사기 좋은 때예요"
        status_suffix = f"(최근 데이터 기준 연중 가격대의 {position*100:.0f}% 지점)"
    elif position <= 0.66:
        light_color, light_label = "#eab308", "보통 가격대예요"
        status_suffix = f"(최근 데이터 기준 연중 가격대의 {position*100:.0f}% 지점)"
    else:
        light_color, light_label = "#ef4444", "지금은 비싼 편이에요"
        status_suffix = f"(최근 데이터 기준 연중 가격대의 {position*100:.0f}% 지점)"

    emoji = FRUIT_INFO.get(selected, {}).get("emoji", "🍏")

    if is_stale:
        # 비수기에는 신뢰도를 위해 AI 문구 대신 사실 관계가 확실한 고정 안내문을 사용한다.
        analysis_text = (
            f"**{latest_year}년 {latest_month}월** 이후로는 유통 데이터가 없어 지금 시세를 확인하기 어려워요.\n"
            f"최근 1년 기준 **{cheapest_year}년 {cheapest_month}월**에 가장 저렴했으니 다음 출하기에 참고해보세요."
        )
    else:
        ai_commentary = generate_price_commentary(
            fruit=selected,
            cheapest_year=cheapest_year,
            cheapest_month=cheapest_month,
            latest_year=latest_year,
            latest_month=latest_month,
            cur_price=cur_price,
            position_pct=position * 100,
            light_label=light_label,
        )
        analysis_text = ai_commentary or (
            f"최근 1년 데이터 기준 **{cheapest_year}년 {cheapest_month}월**에 가격이 가장 낮아지는 경향이 있어요.\n"
            f"가장 최근 집계된 **{latest_year}년 {latest_month}월** 평균가는 약 **{cur_price:,.0f}원**이에요."
        )

    sel_ctgry_cd, sel_item_cd = ITEM_CODE_LOOKUP.get(selected, ("", ""))
    varieties = VARIETY_LOOKUP.get((sel_ctgry_cd, sel_item_cd), [])
    variety_text = ", ".join(varieties[:6]) if varieties else "정보 없음"
    region_text = PRODUCTION_REGIONS.get(selected, "정보 없음")

    # 여러 st.* 위젯을 나눠 호출하면 카드 배경(div)이 실제 내용을 감싸지 못하고 잘려 보이므로,
    # 카드 전체를 하나의 HTML 블록으로 만들어 배경이 내용을 온전히 감싸도록 한다.
    card_html = f"""
    <div class="status-card">
        <div class="emoji">{emoji}</div>
        <div>
            <div class="title">{html.escape(selected)}</div>
            <div>
                <span class="light light-on" style="background:{light_color}; color:{light_color};"></span>
                <b>{html.escape(light_label)}</b> {status_suffix}
            </div>
            <div class="analysis-line">{_format_rich_text(analysis_text)}</div>
            <div class="extra-info">🌱 주요 품종: {html.escape(variety_text)} &nbsp;·&nbsp; 🗺️ 대표 산지: {html.escape(region_text)}</div>
        </div>
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)

    st.write("")

    # ------------------------------------------------------------------
    # 연간 가격 추이 (Plotly)
    # ------------------------------------------------------------------
    st.markdown("#### 📈 최근 12개월 가격 추이")

    # 연도가 바뀌는 지점에서만 "YY년"을 함께 표기해 raw 월 숫자만으로는 구분되지 않는
    # 연도 경계(예: 2025년 12월 -> 2026년 1월)를 명확히 한다.
    chart_df = overall_monthly.copy()
    month_labels = []
    prev_year = None
    for _, r in chart_df.iterrows():
        y, m = int(r["year"]), int(r["month"])
        month_labels.append(f"{y}년 {m}월" if y != prev_year else f"{m}월")
        prev_year = y
    chart_df["month_label"] = month_labels
    bar_colors = ["#22c55e" if i == cheapest_pos else "#f2938a" for i in range(len(chart_df))]

    fig = go.Figure()
    division_series = {}  # 이벤트(최저가/지금) 주석의 y좌표 앵커로 재사용

    # 참고 이미지처럼 심플하게: 글로우/그라디언트 채우기/값 라벨을 모두 걷어내고,
    # 얇은 선 + 속이 빈(hollow) 원형 마커만 남긴다. 정확한 가격은 hover 툴팁으로 확인.
    divisions = fruit_df["division"].unique().tolist()
    if "소매" in divisions or "도매" in divisions:
        line_specs = [("소매", "#f2938a"), ("도매", "#57534e")]
        for div_name, color in line_specs:
            if div_name not in divisions:
                continue
            d = _chronological_monthly(
                fruit_df[fruit_df["division"] == div_name], START_YM, END_YM
            )
            d = d.set_index(["year", "month"]).reindex(
                list(zip(chart_df["year"], chart_df["month"]))
            ).reset_index()
            division_series[div_name] = d["price"]

            fig.add_trace(
                go.Scatter(
                    x=chart_df["month_label"],
                    y=d["price"],
                    mode="lines+markers",
                    name=div_name,
                    line=dict(color=color, width=2.2),
                    marker=dict(size=8, color="#ffffff", line=dict(color=color, width=2)),
                    hovertemplate="%{x}<br>" + div_name + ": %{y:,.0f}원<extra></extra>",
                    connectgaps=False,
                )
            )
    else:
        division_series["평균가"] = chart_df["price"]
        fig.add_trace(
            go.Bar(
                x=chart_df["month_label"],
                y=chart_df["price"],
                marker_color=bar_colors,
                marker_line_width=0,
                name="평균가격",
                hovertemplate="%{x}<br>평균가: %{y:,.0f}원<extra></extra>",
            )
        )

    anchor_series = division_series.get("소매", next(iter(division_series.values())))

    # 최저가 달은 참고 이미지의 "적정시기" 밴드처럼, 은은한 배경 면 + 좌우 점선 테두리로
    # 표현한다. 현재 달은 축 위에 점 하나를 콕 찍고 위로 점선을 그어 "지금 여기"를 가리킨다.
    pos_in_chart = cheapest_pos
    current_pos = len(chart_df) - 1  # 조회 구간의 마지막 자리 = 항상 "지금"

    CHEAPEST_COLOR = "#22c55e"
    CURRENT_COLOR = "#8b5cf6"

    fig.add_vrect(
        x0=pos_in_chart - 0.5, x1=pos_in_chart + 0.5,
        fillcolor=CHEAPEST_COLOR, opacity=0.10, line_width=0,
    )
    for edge in (pos_in_chart - 0.5, pos_in_chart + 0.5):
        fig.add_shape(
            type="line", x0=edge, x1=edge, y0=0, y1=1, xref="x", yref="paper",
            line=dict(color=CHEAPEST_COLOR, width=1.3, dash="dot"),
        )
    fig.add_annotation(
        x=chart_df["month_label"].iloc[pos_in_chart], y=1.06, xref="x", yref="paper",
        text="최저가 달", showarrow=False,
        font=dict(size=12, color=CHEAPEST_COLOR, family="Arial Black, Arial"),
    )

    # 현재 달: 점선 세로선 + 축 위의 점 하나로 "지금 시점"을 가리킨다 (배지 박스 없이 담백하게).
    fig.add_shape(
        type="line", x0=current_pos, x1=current_pos, y0=0, y1=1, xref="x", yref="paper",
        line=dict(color=CURRENT_COLOR, width=1.3, dash="dot"),
    )
    fig.add_shape(
        type="circle", xref="x", yref="paper",
        x0=current_pos - 0.09, x1=current_pos + 0.09, y0=-0.035, y1=0.035,
        fillcolor=CURRENT_COLOR, line=dict(color=CURRENT_COLOR, width=0),
    )
    fig.add_annotation(
        x=chart_df["month_label"].iloc[current_pos], y=1.06, xref="x", yref="paper",
        text="지금", showarrow=False,
        font=dict(size=12, color=CURRENT_COLOR, family="Arial Black, Arial"),
    )

    fig.update_layout(
        height=420,
        margin=dict(l=0, r=0, t=44, b=60),
        plot_bgcolor="rgba(255,255,255,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Pretendard, -apple-system, sans-serif", size=12),
        legend=dict(
            orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5,
            font=dict(size=12), bgcolor="rgba(0,0,0,0)",
        ),
        yaxis=dict(
            # 미니멀한 디자인: 축 눈금/숫자와 그리드를 감추고, 정확한 가격은
            # hover 툴팁(hovermode="x unified")으로만 확인하도록 한다.
            title=None,
            showticklabels=False,
            showgrid=False,
            zeroline=False,
            automargin=True,
        ),
        xaxis=dict(
            title=None,
            showgrid=False,
            showline=True,
            linecolor="#e5e7eb",
            tickangle=-45,
            tickfont=dict(size=11, color="#9ca3af"),
            automargin=True,
        ),
        hovermode="x unified",
    )

    with st.container(border=True):
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"responsive": True, "displayModeBar": False, "scrollZoom": False},
        )

        # --------------------------------------------------------------
        # 월별 가격표 (그래프 검증용, 접어두는 심플한 표)
        # --------------------------------------------------------------
        with st.expander("🧾 월별 가격표 보기"):
            table_df = pd.DataFrame({"연월": chart_df["month_label"]})
            if "소매" in divisions or "도매" in divisions:
                for div_name in ["소매", "도매"]:
                    if div_name in divisions:
                        d = _chronological_monthly(
                            fruit_df[fruit_df["division"] == div_name], START_YM, END_YM
                        )
                        d = d.set_index(["year", "month"]).reindex(
                            list(zip(chart_df["year"], chart_df["month"]))
                        ).reset_index()
                        table_df[div_name] = d["price"]
            else:
                table_df["평균가"] = chart_df["price"]

            mark = [""] * len(table_df)
            if current_pos == pos_in_chart:
                mark[pos_in_chart] = "최저가 · 지금"
            else:
                mark[pos_in_chart] = "최저가"
                mark[current_pos] = "지금"
            table_df["구분"] = mark

            price_cols = [c for c in table_df.columns if c not in ("연월", "구분")]
            for c in price_cols:
                table_df[c] = table_df[c].apply(lambda v: f"{v:,.0f}원" if pd.notna(v) else "-")
            table_df = table_df[["연월", *price_cols, "구분"]]

            # st.dataframe은 셀 정렬을 세밀하게 제어하기 어려워, 가운데 정렬이 되는
            # 단순한 HTML 테이블로 직접 그린다.
            header_html = "".join(f"<th>{html.escape(c)}</th>" for c in table_df.columns)
            rows_html = []
            for _, r in table_df.iterrows():
                row_class = ' class="event-row"' if r["구분"] else ""
                cells = "".join(f"<td>{html.escape(str(v))}</td>" for v in r)
                rows_html.append(f"<tr{row_class}>{cells}</tr>")
            table_html = (
                f'<table class="price-table"><thead><tr>{header_html}</tr></thead>'
                f'<tbody>{"".join(rows_html)}</tbody></table>'
            )
            st.markdown(table_html, unsafe_allow_html=True)

    ctgry_cd, item_cd = ITEM_CODE_LOOKUP.get(selected, ("", ""))

    # ----------------------------------------------------------------------------
    # 스마트 구매팁 (Solar로 생성, 실패 시 규칙 기반 문구로 대체)
    # ----------------------------------------------------------------------------
    st.write("")
    st.markdown("#### 🛒 스마트 구매팁")

    varieties_for_tip = VARIETY_LOOKUP.get((ctgry_cd, item_cd), [])
    regions_for_tip = PRODUCTION_REGIONS.get(selected, "국내산·수입 다양")

    tips = generate_smart_buying_tips(selected, tuple(varieties_for_tip), regions_for_tip)
    if tips:
        variety_tip = tips["variety_tip"]
        selection_tip = tips["selection_tip"]
        channel_tip = tips["channel_tip"]
    else:
        variety_names = ", ".join(varieties_for_tip[:4]) if varieties_for_tip else None
        variety_tip = (
            f"주요 품종으로는 {variety_names}이(가) 있으며, 품종에 따라 당도와 수확 시기가 "
            "조금씩 다릅니다." if variety_names else "품종별 상세 정보는 아직 준비 중입니다."
        )
        selection_tip = FRUIT_SELECTION_TIPS.get(selected, GENERIC_SELECTION_TIP)
        channel_tip = CHANNEL_TIP_DEFAULT_TEMPLATE.format(regions=regions_for_tip)

    tip_cards = [
        ("🌱", "mint", "품종별 구매팁", variety_tip),
        ("👀", "lavender", "좋은 과일 고르는 법", selection_tip),
        ("🏪", "orange", "구매 채널 안내", channel_tip),
    ]
    with st.container(key="tip_cards"):
        tip_cols = st.columns(3)
        for col, (emoji, chip, title, text) in zip(tip_cols, tip_cards):
            with col:
                st.markdown(
                    f'<div class="tip-card">'
                    f'<div class="tip-card-emoji {chip}">{emoji}</div>'
                    f'<div class="tip-card-title">{html.escape(title)}</div>'
                    f'<div class="tip-card-text">{_format_rich_text(text)}</div>'
                    f"</div>",
                    unsafe_allow_html=True,
                )

    st.write("")
    st.caption(
        f"데이터 기간: {START_YM[:4]}.{START_YM[4:]} ~ {END_YM[:4]}.{END_YM[4:]}  ·  "
        f"부류코드 {ctgry_cd} / 품목코드 {item_cd}  ·  "
        f"{SOURCE_NOTE}"
        + ("" if is_live else " (데모 데이터)")
    )
