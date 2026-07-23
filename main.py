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
"""

import math
import os
import random
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

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

# se_cd(구분코드): 01 소매, 02 중도매, 03/07 친환경농산물 — 이 앱에서는 01→소매, 02→도매로 표시
DIVISION_LABELS = {"01": "소매", "02": "도매"}

# 화면 상단 큰 버튼에 노출할 16개 과일 (요청 순서 그대로)
MAIN_FRUITS = [
    "사과", "배", "복숭아", "포도", "감귤", "단감", "바나나", "참다래",
    "파인애플", "오렌지", "자몽", "레몬", "체리", "망고", "블루베리", "아보카도",
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
    "단감":     {"emoji": "🟠", "base": 3400, "amp": 900,  "cheap_month": 11},
    "바나나":   {"emoji": "🍌", "base": 2200, "amp": 300,  "cheap_month": 6},
    "참다래":   {"emoji": "🥝", "base": 3800, "amp": 700,  "cheap_month": 12},
    "파인애플": {"emoji": "🍍", "base": 3600, "amp": 500,  "cheap_month": 7},
    "오렌지":   {"emoji": "🍊", "base": 3300, "amp": 700,  "cheap_month": 2},
    "자몽":     {"emoji": "🍈", "base": 2900, "amp": 600,  "cheap_month": 1},
    "레몬":     {"emoji": "🍋", "base": 3100, "amp": 500,  "cheap_month": 4},
    "체리":     {"emoji": "🍒", "base": 8500, "amp": 3000, "cheap_month": 6},
    "망고":     {"emoji": "🥭", "base": 5200, "amp": 1500, "cheap_month": 7},
    "블루베리": {"emoji": "🫐", "base": 6000, "amp": 1800, "cheap_month": 6},
    "아보카도": {"emoji": "🥑", "base": 3300, "amp": 500,  "cheap_month": 9},
    "수박":     {"emoji": "🍉", "base": 22000, "amp": 8000, "cheap_month": 8},
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

# ----------------------------------------------------------------------------
# 스타일
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main-copy {
        font-size: 2.1rem;
        font-weight: 800;
        text-align: center;
        margin-bottom: 0.3rem;
        color: #1f2933;
    }
    .sub-copy {
        text-align: center;
        color: #6b7280;
        margin-bottom: 1.4rem;
        font-size: 1rem;
    }
    div[data-testid="stTextInput"] input {
        border-radius: 999px !important;
        padding: 0.9rem 1.3rem !important;
        font-size: 1.1rem !important;
        border: 2px solid #ffb37a !important;
    }
    .stButton > button {
        border-radius: 999px;
        border: 1px solid #f0d9c8;
        background-color: #fff7f0;
        color: #5c3a21;
        font-weight: 600;
        padding: 0.55rem 0.4rem;
        transition: all 0.15s ease-in-out;
    }
    .stButton > button:hover {
        background-color: #ffb37a;
        color: white;
        border-color: #ffb37a;
    }
    .tag-row .stButton > button {
        border-radius: 999px;
        font-size: 0.78rem;
        padding: 0.25rem 0.7rem;
        background-color: #f3f4f6;
        color: #6b7280;
        border: 1px solid #e5e7eb;
        font-weight: 500;
    }
    .status-card {
        border-radius: 20px;
        padding: 1.6rem 1.8rem;
        background: linear-gradient(135deg, #fff9f2 0%, #fff2e2 100%);
        border: 1px solid #ffe3c2;
        box-shadow: 0 6px 18px rgba(0,0,0,0.05);
    }
    .light {
        display: inline-block;
        width: 16px;
        height: 16px;
        border-radius: 50%;
        margin-right: 6px;
    }
    .light-on { box-shadow: 0 0 10px currentColor; }
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
def cheapest_fruit_this_month(start_ym: str, end_ym: str) -> str:
    """이번 달 기준, 연중 최저가 대비 가장 저렴해 보이는 과일을 하나 골라준다."""
    best, best_ratio = MAIN_FRUITS[0], math.inf
    for f in MAIN_FRUITS:
        series, _ = fetch_fruit_series(f, start_ym, end_ym)
        if series.empty:
            continue
        monthly = series.groupby("month")["price"].mean()
        if monthly.empty:
            continue
        latest_month = monthly.index.max()
        cur_price = monthly.loc[latest_month]
        year_min = monthly.min()
        if year_min:
            ratio = cur_price / year_min
            if ratio < best_ratio:
                best, best_ratio = f, ratio
    return best


# ----------------------------------------------------------------------------
# 상단: 메인 카피 + 검색창 + 태그
# ----------------------------------------------------------------------------
st.markdown('<div class="main-copy">좋아하는 과일, 언제 사야 가장 싸고 맛있을까?</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-copy">과일 이름을 검색하거나 아래에서 골라보세요. 연간 가격 추이로 최적의 구매 시기를 알려드려요.</div>', unsafe_allow_html=True)

search_query = st.text_input(
    "fruit_search",
    placeholder="과일 이름을 검색해보세요 (예: 복숭아, 사과, 샤인머스캣)",
    label_visibility="collapsed",
)

st.markdown('<div class="tag-row">', unsafe_allow_html=True)
tag_cols = st.columns(len(TAGS) + 6)
for i, (label, target) in enumerate(TAGS):
    with tag_cols[i]:
        if st.button(label, key=f"tag_{i}"):
            if target == "__CHEAPEST_NOW__":
                select_fruit(cheapest_fruit_this_month(START_YM, END_YM))
            else:
                select_fruit(target)
st.markdown("</div>", unsafe_allow_html=True)

if search_query.strip():
    matches = [f for f in FRUIT_INFO.keys() if search_query.strip() in f]
    if matches:
        select_fruit(matches[0])
    else:
        st.warning(f"'{search_query}'와(과) 일치하는 과일 데이터를 찾지 못했어요. 목록에 있는 과일명으로 다시 시도해보세요.")

st.write("")

# 과일 버튼 그리드 (4 x 4)
cols_per_row = 4
for row_start in range(0, len(MAIN_FRUITS), cols_per_row):
    row_fruits = MAIN_FRUITS[row_start: row_start + cols_per_row]
    cols = st.columns(cols_per_row)
    for col, fruit in zip(cols, row_fruits):
        with col:
            emoji = FRUIT_INFO.get(fruit, {}).get("emoji", "🍏")
            if st.button(f"{emoji} {fruit}", key=f"fruit_{fruit}", use_container_width=True):
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

if fruit_df.empty:
    st.error(f"'{selected}'에 대한 데이터가 없어요. 다른 과일을 선택해보세요.")
else:
    overall_monthly = fruit_df.groupby("month")["price"].mean().reset_index().sort_values("month")

    cheapest_row = overall_monthly.loc[overall_monthly["price"].idxmin()]
    cheapest_month = int(cheapest_row["month"])
    year_min = overall_monthly["price"].min()
    year_max = overall_monthly["price"].max()

    latest_month = int(overall_monthly["month"].iloc[-1])
    cur_price = overall_monthly[overall_monthly["month"] == latest_month]["price"].mean()

    span = max(year_max - year_min, 1e-6)
    position = (cur_price - year_min) / span  # 0(가장 쌈) ~ 1(가장 비쌈)

    if position <= 0.33:
        light_color, light_label = "#22c55e", "지금이 딱 사기 좋은 때예요"
    elif position <= 0.66:
        light_color, light_label = "#eab308", "보통 가격대예요"
    else:
        light_color, light_label = "#ef4444", "지금은 비싼 편이에요"

    emoji = FRUIT_INFO.get(selected, {}).get("emoji", "🍏")

    with st.container():
        st.markdown('<div class="status-card">', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 3])
        with c1:
            st.markdown(f"<div style='font-size:5rem; text-align:center;'>{emoji}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"### {selected}")
            st.markdown(
                f"<span class='light light-on' style='background:{light_color}; color:{light_color};'></span>"
                f"<b>{light_label}</b> (최근 데이터 기준 연중 가격대의 {position*100:.0f}% 지점)",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"최근 1년 데이터 기준 **{cheapest_month}월**에 가격이 가장 낮아지는 경향이 있어요.  \n"
                f"가장 최근 집계된 **{latest_month}월** 평균가는 약 **{cur_price:,.0f}원**이에요."
            )
        st.markdown("</div>", unsafe_allow_html=True)

    st.write("")

    # ------------------------------------------------------------------
    # 연간 가격 추이 (Plotly)
    # ------------------------------------------------------------------
    st.markdown("#### 📈 최근 12개월 가격 추이")

    chart_df = overall_monthly.copy()
    chart_df["month_label"] = chart_df["month"].apply(lambda m: f"{m}월")
    bar_colors = ["#ffb37a" if m != cheapest_month else "#22c55e" for m in chart_df["month"]]

    fig = go.Figure()

    divisions = fruit_df["division"].unique().tolist()
    if "소매" in divisions or "도매" in divisions:
        for div_name, color in [("소매", "#f97316"), ("도매", "#60a5fa")]:
            if div_name in divisions:
                d = (
                    fruit_df[fruit_df["division"] == div_name]
                    .groupby("month")["price"].mean()
                    .reindex(chart_df["month"])
                    .reset_index()
                )
                fig.add_trace(
                    go.Scatter(
                        x=chart_df["month_label"],
                        y=d["price"],
                        mode="lines+markers",
                        name=div_name,
                        line=dict(color=color, width=3),
                        marker=dict(size=7),
                    )
                )
    else:
        fig.add_trace(
            go.Bar(
                x=chart_df["month_label"],
                y=chart_df["price"],
                marker_color=bar_colors,
                name="평균가격",
            )
        )

    # 최저가 달 배경 하이라이트
    pos_in_chart = list(chart_df["month"]).index(cheapest_month)
    fig.add_vrect(
        x0=pos_in_chart - 0.5,
        x1=pos_in_chart + 0.5,
        fillcolor="#22c55e",
        opacity=0.12,
        line_width=0,
        annotation_text="최저가 달",
        annotation_position="top",
    )

    fig.update_layout(
        height=420,
        margin=dict(l=20, r=20, t=40, b=20),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis_title="원 (kg 등 품목별 기준 단위)",
        xaxis_title=None,
    )

    st.plotly_chart(fig, use_container_width=True)

    ctgry_cd, item_cd = ITEM_CODE_LOOKUP.get(selected, ("", ""))
    st.caption(
        f"데이터 기간: {START_YM[:4]}.{START_YM[4:]} ~ {END_YM[:4]}.{END_YM[4:]}  ·  "
        f"부류코드 {ctgry_cd} / 품목코드 {item_cd}  ·  "
        f"출처: 한국농수산식품유통공사 연월별 도,소매가격정보"
        + ("" if is_live else " (데모 데이터)")
    )
