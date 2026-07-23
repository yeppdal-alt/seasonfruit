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
    "참외":     {"emoji": "🍈", "base": 4200, "amp": 1600, "cheap_month": 6},
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
        --accent: #ff8a5b;
        --accent-dark: #e8703f;
        --ink: #171412;
        --muted: #8c8078;
    }
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #fdf1d8 0%, #fbdcc9 45%, #f8c7cf 100%);
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
    /* 기본(선택되지 않은) 버튼: 화이트 필, 블랙 텍스트 — 온보딩 카드의 리스트 옵션과 동일한 톤 */
    .stButton > button {
        border-radius: 999px;
        border: 1px solid rgba(23,20,18,0.08);
        background-color: #ffffff;
        color: var(--ink);
        font-weight: 600;
        padding: 0.6rem 0.4rem;
        box-shadow: 0 2px 8px rgba(23,20,18,0.05);
        transition: all 0.15s ease-in-out;
    }
    .stButton > button:hover {
        border-color: var(--accent);
        color: var(--accent-dark);
    }
    /* 선택된(primary) 버튼: 코럴 그라디언트 필 + 화이트 텍스트 */
    .stButton > button[kind="primary"] {
        border-radius: 999px;
        border: none;
        background: linear-gradient(135deg, var(--accent) 0%, var(--accent-dark) 100%);
        color: #ffffff;
        font-weight: 700;
        box-shadow: 0 6px 16px rgba(232,112,63,0.35);
    }
    .stButton > button[kind="primary"]:hover {
        filter: brightness(1.05);
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
        border-radius: 24px;
        min-height: 260px;
        padding: 2.2rem 2.4rem;
        display: flex;
        align-items: center;
        gap: 1.8rem;
        background: rgba(255, 255, 255, 0.72);
        border: 1px solid rgba(255,255,255,0.9);
        box-shadow: 0 10px 26px rgba(232, 112, 63, 0.16);
        margin-bottom: 1rem;
    }
    .status-card .emoji {
        font-size: 5.5rem;
        line-height: 1;
        flex-shrink: 0;
        filter: drop-shadow(0 6px 10px rgba(232,112,63,0.25));
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
        background-color: rgba(255,138,91,0.08);
        font-weight: 700;
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
    bar_colors = ["#22c55e" if i == cheapest_pos else "#ff8a5b" for i in range(len(chart_df))]

    def _fmt_label(v) -> str:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return ""
        if v >= 10000:
            return f"{v / 10000:.1f}만"
        return f"{v / 1000:.1f}천"

    fig = go.Figure()
    division_series = {}  # 이벤트(최저가/지금) 주석의 y좌표 앵커로 재사용

    divisions = fruit_df["division"].unique().tolist()
    if "소매" in divisions or "도매" in divisions:
        # 소매(위쪽 값표시) / 도매(아래쪽 값표시)를 부드러운 곡선 + 글로우 + 값 라벨로 표현
        line_specs = [
            ("소매", "#ff8a5b", "#c2410c", "top center", "rgba(255,138,91,0.18)"),
            ("도매", "#57534e", "#3f3c38", "bottom center", "rgba(87,83,78,0.08)"),
        ]
        for div_name, color, label_color, text_pos, fill_color in line_specs:
            if div_name not in divisions:
                continue
            d = _chronological_monthly(
                fruit_df[fruit_df["division"] == div_name], START_YM, END_YM
            )
            d = d.set_index(["year", "month"]).reindex(
                list(zip(chart_df["year"], chart_df["month"]))
            ).reset_index()
            division_series[div_name] = d["price"]

            # 은은한 글로우 효과: 같은 선을 굵고 흐리게 한 번 더 깔아준다.
            fig.add_trace(
                go.Scatter(
                    x=chart_df["month_label"],
                    y=d["price"],
                    mode="lines",
                    line=dict(color=color, width=11, shape="spline", smoothing=0.35),
                    opacity=0.13,
                    showlegend=False,
                    hoverinfo="skip",
                )
            )
            fig.add_trace(
                go.Scatter(
                    x=chart_df["month_label"],
                    y=d["price"],
                    mode="lines+markers+text",
                    name=div_name,
                    line=dict(color=color, width=3.5, shape="spline", smoothing=0.35),
                    marker=dict(size=9, color=color, line=dict(color="white", width=2)),
                    text=[_fmt_label(v) for v in d["price"]],
                    textposition=text_pos,
                    textfont=dict(size=12, color=label_color, family="Arial Black, Arial"),
                    fill="tozeroy",
                    fillcolor=fill_color,
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
                text=[_fmt_label(v) for v in chart_df["price"]],
                textposition="outside",
                textfont=dict(size=12, color="#5c3a21", family="Arial Black, Arial"),
                hovertemplate="%{x}<br>평균가: %{y:,.0f}원<extra></extra>",
            )
        )

    # 참고용 평균가 기준선
    anchor_series = division_series.get("소매", next(iter(division_series.values())))
    avg_price = anchor_series.mean()
    if pd.notna(avg_price):
        fig.add_hline(
            y=avg_price,
            line_dash="dot",
            line_color="#c4c4c4",
            line_width=1.3,
            annotation_text=f"평균 {avg_price:,.0f}원",
            annotation_position="right",
            annotation_font=dict(size=11, color="#9ca3af"),
        )

    # 최저가 달은 은은한 배경 하이라이트(면), 현재 달은 또렷한 점선 박스(테두리)로 서로 다른
    # 방식을 사용한다. 표현 방식 자체가 다르기 때문에 두 시점이 같은 달이어도 서로 지워지지
    # 않고 "면 + 테두리"가 함께 보인다.
    pos_in_chart = cheapest_pos
    current_pos = len(chart_df) - 1  # 조회 구간의 마지막 자리 = 항상 "지금"

    CHEAPEST_COLOR = "#22c55e"
    CURRENT_COLOR = "#8b5cf6"

    # 최저가 달: 배경 면 하이라이트
    fig.add_vrect(
        x0=pos_in_chart - 0.5, x1=pos_in_chart + 0.5,
        fillcolor=CHEAPEST_COLOR, opacity=0.13, line_width=0,
    )
    # 현재 달: 점선 테두리 박스 (채우기 없음)
    fig.add_shape(
        type="rect",
        x0=current_pos - 0.5, x1=current_pos + 0.5,
        y0=0, y1=1, xref="x", yref="paper",
        line=dict(color=CURRENT_COLOR, width=2.5, dash="dot"),
        fillcolor="rgba(0,0,0,0)",
    )

    # 두 이벤트가 같은 달일 때도 라벨이 겹치지 않도록 배지를 위아래로 나눠 쌓는다.
    same_month = current_pos == pos_in_chart
    fig.add_annotation(
        x=chart_df["month_label"].iloc[pos_in_chart], y=1.0, xref="x", yref="paper",
        text="🏆 최저가 달", showarrow=False, yshift=(44 if same_month else 18),
        font=dict(size=12, color="white", family="Arial Black, Arial"),
        bgcolor=CHEAPEST_COLOR, borderpad=5,
    )
    fig.add_annotation(
        x=chart_df["month_label"].iloc[current_pos], y=1.0, xref="x", yref="paper",
        text="📍 지금", showarrow=False, yshift=18,
        font=dict(size=12, color="white", family="Arial Black, Arial"),
        bgcolor=CURRENT_COLOR, borderpad=5,
    )

    # 범례를 차트 "아래"로 내려서, 상단은 최저가/지금 뱃지 전용 공간으로 비워둔다.
    # (좁은 화면에서 범례와 뱃지가 같은 줄에서 겹쳐 보이는 것이 모바일 "깨짐"의 주요 원인이었다.)
    fig.update_layout(
        height=460,
        margin=dict(l=10, r=10, t=66, b=70),
        plot_bgcolor="rgba(255,255,255,0.35)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Pretendard, -apple-system, sans-serif", size=12),
        legend=dict(
            orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5,
            font=dict(size=12), bgcolor="rgba(0,0,0,0)",
        ),
        yaxis=dict(
            title=dict(text="가격 (원)", font=dict(size=11, color="#9ca3af")),
            showgrid=True,
            gridcolor="#f2f2f2",
            gridwidth=1,
            zeroline=False,
            tickformat=",",
            automargin=True,
        ),
        xaxis=dict(
            title=None,
            showgrid=True,
            gridcolor="#f8f8f8",
            showline=True,
            linecolor="#e5e7eb",
            tickangle=-45,
            tickfont=dict(size=11),
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
    st.caption(
        f"데이터 기간: {START_YM[:4]}.{START_YM[4:]} ~ {END_YM[:4]}.{END_YM[4:]}  ·  "
        f"부류코드 {ctgry_cd} / 품목코드 {item_cd}  ·  "
        f"{SOURCE_NOTE}"
        + ("" if is_live else " (데모 데이터)")
    )
