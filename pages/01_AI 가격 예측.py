"""
🍑 과일 가격 예측기 (pages/02_aiprice.py)
-------------------------------------------------------------------------
한국농수산식품유통공사(aT)의 공공데이터로 과일 하나를 골라 "연도별 평균가격"을
모으고, scikit-learn의 아주 단순한 선형회귀로 "연도 -> 평균가격" 관계를 배워보는
페이지입니다.

⚠️ 이 페이지는 배움/체험용 데모예요. 선형회귀는 "가격이 해마다 똑같은 정도로만
오르거나 내린다"고 가정하는 아주 단순한 직선 모델이라, 실제 미래 가격을 정확히
맞히지는 못해요. 특히 실제로 학습에 쓰인 연도 범위를 한참 벗어난 예측(아주 먼 미래나
아주 먼 과거)은 숫자가 비현실적으로 나올 수 있으니 참고용으로만 봐주세요.

초보자를 위해 각 단계마다 "왜 이렇게 하는지"를 한국어 주석으로 최대한 자세히
남겨뒀습니다.
"""

import os
from datetime import date

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score

# ----------------------------------------------------------------------------
# 기본 설정
# ----------------------------------------------------------------------------
st.set_page_config(page_title="과일 가격 예측기", page_icon="🔮", layout="wide")

# 이 파일은 pages/ 폴더 안에 있으므로, 한 단계 위(프로젝트 루트)로 올라가야
# price_code.xlsx를 찾을 수 있어요.
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRICE_CODE_PATH = os.path.join(BASE_DIR, "price_code.xlsx")

# main.py와 동일한 aT(한국농수산식품유통공사) 공공데이터 엔드포인트를 그대로 사용합니다.
API_ENDPOINT = "https://apis.data.go.kr/B552845/perYearMonth/price"

# se_cd(구분코드): 01 소매, 02 중도매, 03/07 친환경농산물
DIVISION_LABELS = {"01": "소매", "02": "도매", "03": "친환경", "07": "친환경"}

# main.py의 20개 과일과 동일한 목록 + 대표 이모지 (이 페이지에서 고를 수 있는 과일들)
FRUIT_EMOJI = {
    "사과": "🍎", "배": "🍐", "복숭아": "🍑", "포도": "🍇", "감귤": "🍊",
    "단감": "🍊", "바나나": "🍌", "참다래": "🥝", "파인애플": "🍍", "오렌지": "🍊",
    "자몽": "🍊", "레몬": "🍋", "체리": "🍒", "망고": "🥭", "블루베리": "🫐",
    "아보카도": "🥑", "수박": "🍉", "토마토": "🍅", "딸기": "🍓", "참외": "🍋",
}
FRUIT_LIST = list(FRUIT_EMOJI.keys())

# price_code.xlsx 로딩이 실패했을 때를 대비한 최소 폴백 (ctgry_cd, item_cd)
FALLBACK_ITEM_CODES = {
    "사과": ("400", "411"), "배": ("400", "412"), "복숭아": ("400", "413"),
    "포도": ("400", "414"), "감귤": ("400", "415"), "단감": ("400", "416"),
    "바나나": ("400", "418"), "참다래": ("400", "419"), "파인애플": ("400", "420"),
    "오렌지": ("400", "421"), "자몽": ("400", "423"), "레몬": ("400", "424"),
    "체리": ("400", "425"), "망고": ("400", "428"), "블루베리": ("400", "429"),
    "아보카도": ("400", "430"), "수박": ("200", "221"),
    "토마토": ("200", "225"), "딸기": ("200", "226"), "참외": ("200", "222"),
}

# 연도별 평균을 낼 때, "이 정도는 관측이 돼야 그 해를 대표할 수 있다"고 보는 최소 개월 수.
# 이보다 적게 관측된 해(예: 데이터가 막 시작된 첫 해, 유통이 뜸했던 해)는 평균을 왜곡시키기
# 때문에 학습에서 제외합니다.
MIN_MONTHS_PER_YEAR = 6


# ----------------------------------------------------------------------------
# 스타일 (따뜻한 톤: main.py와 비슷한 화이트 카드 + 코럴 포인트)
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .aiprice-title { font-size: 2rem; font-weight: 800; color: #241f2b; margin-bottom: 0.2rem; }
    .aiprice-sub { color: #9b93a3; margin-bottom: 1.2rem; }
    .metric-card {
        background: #ffffff; border-radius: 20px; padding: 1.2rem 1.4rem;
        box-shadow: 0 4px 14px rgba(36,31,43,0.06); text-align: center;
    }
    .metric-card .label { color: #9b93a3; font-size: 0.9rem; margin-bottom: 0.3rem; }
    .metric-card .value { font-size: 2.4rem; font-weight: 800; color: #e0665a; }
    .predict-card {
        background: linear-gradient(135deg, #fdf1f0 0%, #fbe8e6 100%);
        border-radius: 24px; padding: 1.6rem; text-align: center;
        border: 1px solid rgba(224,102,90,0.15);
    }
    .predict-card .label { color: #9b93a3; font-size: 1rem; margin-bottom: 0.4rem; }
    .predict-card .value { font-size: 3rem; font-weight: 800; color: #241f2b; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="aiprice-title">🔮 과일 가격 예측기</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="aiprice-sub">과일을 하나 고르면, 지난 연도별 평균가격에 선형회귀(직선 하나)를 '
    '학습시켜서 원하는 연도의 가격을 예측해봐요.</div>',
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# 품목코드 조회 (price_code.xlsx에서 과일 이름 -> (부류코드, 품목코드))
# ----------------------------------------------------------------------------
@st.cache_data(ttl=None, show_spinner=False)
def load_item_codes() -> dict:
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
for _name in FRUIT_LIST:
    if _name not in ITEM_CODE_LOOKUP:
        ITEM_CODE_LOOKUP[_name] = FALLBACK_ITEM_CODES.get(_name, ("400", ""))


# ----------------------------------------------------------------------------
# 공공데이터 API 호출 + 파싱 (main.py와 같은 방식)
# ----------------------------------------------------------------------------
def _call_api(api_key: str, ctgry_cd: str, item_cd: str, start_ym: str, end_ym: str):
    """페이지(pageNo)를 넘겨가며 해당 기간의 모든 데이터를 모은다."""
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
        resp = requests.get(API_ENDPOINT, params=params, timeout=15)
        resp.raise_for_status()
        payload = resp.json()
        try:
            items = payload["response"]["body"]["items"]
            if isinstance(items, dict):
                items = items.get("item", [])
            if isinstance(items, dict):  # 결과가 1건이면 dict 하나로만 오는 경우 방어
                items = [items]
        except (KeyError, TypeError):
            items = []
        if not items:
            break
        rows.extend(items)
        if len(items) < num_of_rows or page_no > 20:  # 너무 많은 페이지를 돌지 않도록 안전장치
            break
        page_no += 1
    return rows


def _parse_unit_size(raw) -> float | None:
    """unit_sz(단위크기) 문자열에서 숫자만 뽑아낸다. (예: "10kg" -> 10.0)"""
    if raw is None:
        return None
    s = str(raw).strip().replace(",", "")
    digits = "".join(ch for ch in s if ch.isdigit() or ch == ".")
    try:
        value = float(digits)
        return value if value > 0 else None
    except ValueError:
        return None


def _unit_size_to_kg(unit: str, size: float) -> float | None:
    """kg 단위로 환산 가능하면 환산하고, 개/상자처럼 무게 단위가 아니면 None을 반환한다."""
    if size is None or size <= 0:
        return None
    u = str(unit or "").strip().lower().replace(" ", "")
    if "kg" in u or "키로" in u or "킬로" in u:
        return size
    if "g" in u or "그램" in u:
        return size / 1000.0
    return None


def _parse_rows(raw_rows: list) -> list:
    """API 원본 응답을 (연도, 월, 구분, 가격) 형태의 표로 바꾼다."""
    parsed = []
    for row in raw_rows:
        se_cd = str(row.get("se_cd", "")).strip()
        division = DIVISION_LABELS.get(se_cd)
        if division is None:
            continue
        exmn_ym = str(row.get("exmn_ym", "")).strip()
        price_raw = row.get("pmm_avgprc")
        if len(exmn_ym) != 6 or price_raw in (None, "", "-"):
            continue
        try:
            price = float(str(price_raw).replace(",", "").strip())
        except ValueError:
            continue
        if price <= 0:
            continue

        # 가능하면 1kg당 가격으로 통일해서, 단위가 다른 달끼리 비교해도 왜곡이 없게 한다.
        unit_size = _parse_unit_size(row.get("unit_sz"))
        weight_kg = _unit_size_to_kg(row.get("unit"), unit_size) if unit_size else None
        if weight_kg:
            unit_price = price / weight_kg
        elif unit_size:
            unit_price = price / unit_size
        else:
            unit_price = price

        parsed.append(
            {
                "year": int(exmn_ym[:4]),
                "month": int(exmn_ym[4:6]),
                "division": division,
                "price": unit_price,
            }
        )
    return parsed


def _build_demo_history(fruit_name: str, start_year: int, end_year: int) -> pd.DataFrame:
    """실 API 연동이 안 될 때를 대비한 장기 가짜 데이터.
    해가 지날수록 물가상승을 완만히 반영해 오르는 추세 + 약간의 계절성 + 노이즈를 섞어서
    "그럴듯한" 연도별 흐름을 만든다. 어디까지나 데모용이라는 점을 화면에도 안내한다."""
    import math
    import random

    rnd = random.Random(fruit_name)
    base = {"사과": 3200, "배": 3500, "복숭아": 4200, "포도": 4500, "감귤": 2600}.get(fruit_name, 3500)
    records = []
    for year in range(start_year, end_year + 1):
        # 연 2%씩 완만하게 오른다고 가정한 인플레이션 추세
        trend = base * (1.02 ** (year - start_year))
        for month in range(1, 13):
            if year == end_year and month > date.today().month:
                break  # 올해는 아직 지나지 않은 달의 데이터는 만들지 않는다
            seasonal = -math.cos(2 * math.pi * month / 12) * base * 0.1
            noise = rnd.uniform(-0.05, 0.05) * base
            price = max(trend + seasonal + noise, 100)
            records.append({"year": year, "month": month, "division": "소매", "price": price})
    return pd.DataFrame(records)


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_price_history(fruit_name: str, start_ym: str, end_ym: str):
    """가능한 한 오래 전 데이터까지 조회한다. (반환값: DataFrame, is_live)"""
    ctgry_cd, item_cd = ITEM_CODE_LOOKUP.get(fruit_name, FALLBACK_ITEM_CODES.get(fruit_name, ("400", "")))
    start_year, end_year = int(start_ym[:4]), int(end_ym[:4])
    try:
        # 🔑 인증키는 절대 코드에 직접 적지 않고, secrets(비밀 금고)에서만 불러옵니다.
        api_key = st.secrets["FRUITS_API_KEY"]
        if not item_cd:
            raise ValueError("item_cd not found")
        raw_rows = _call_api(api_key, ctgry_cd, item_cd, start_ym, end_ym)
        parsed = _parse_rows(raw_rows)
        if not parsed:
            return _build_demo_history(fruit_name, start_year, end_year), False
        return pd.DataFrame(parsed), True
    except Exception:
        return _build_demo_history(fruit_name, start_year, end_year), False


# ----------------------------------------------------------------------------
# 1) 과일 선택
# ----------------------------------------------------------------------------
selected_fruit = st.selectbox(
    "과일을 선택하세요",
    FRUIT_LIST,
    format_func=lambda f: f"{FRUIT_EMOJI.get(f, '🍏')} {f}",
)

today = date.today()
START_YM = "199001"  # 최대한 오래 전부터 요청해본다 (실제로 존재하는 만큼만 돌아온다)
END_YM = f"{today.year}{today.month:02d}"

raw_df, is_live = fetch_price_history(selected_fruit, START_YM, END_YM)

if not is_live:
    st.info(
        "ℹ️ 공공데이터 API 응답을 확인하지 못해, 계절성과 완만한 물가상승 추세를 반영한 "
        "**데모 데이터**로 대신 보여드리고 있어요. `FRUITS_API_KEY`가 secrets에 잘 등록됐는지, "
        "인증키 활용 승인이 완료됐는지 확인해 주세요.",
        icon="ℹ️",
    )

if raw_df.empty:
    st.error(f"'{selected_fruit}'에 대한 데이터를 하나도 찾지 못했어요. 다른 과일을 선택해보세요.")
    st.stop()

# ----------------------------------------------------------------------------
# 2) "소매" 위주로 하나의 구분만 골라 쓴다 (도매·소매를 섞으면 가격 수준 자체가 달라서
#    연도별 추세가 왜곡될 수 있기 때문). 소매가 없으면 있는 구분 중 하나를 그대로 쓴다.
# ----------------------------------------------------------------------------
if "소매" in raw_df["division"].unique():
    price_df = raw_df[raw_df["division"] == "소매"].copy()
    price_basis_label = "소매가"
else:
    only_division = raw_df["division"].unique()[0]
    price_df = raw_df[raw_df["division"] == only_division].copy()
    price_basis_label = only_division

# ----------------------------------------------------------------------------
# 3) 날짜(연월)에서 "연도"만 뽑아 연도별 평균가격으로 묶는다.
#    (year 컬럼은 위에서 이미 exmn_ym의 앞 4자리로 뽑아뒀다.)
# ----------------------------------------------------------------------------
yearly = (
    price_df.groupby("year")
    .agg(평균가격=("price", "mean"), 관측월수=("month", "nunique"))
    .reset_index()
    .rename(columns={"year": "연도"})
    .sort_values("연도")
    .reset_index(drop=True)
)

# ----------------------------------------------------------------------------
# 4) 학습에 쓸 연도를 거른다.
#    - 올해(진행중인 해): 아직 12개월이 다 지나지 않아서 평균이 왜곡되므로 제외
#    - 관측월수가 너무 적은 해(첫 해, 유통이 뜸했던 해 등): 그 해를 대표하지 못하므로 제외
# ----------------------------------------------------------------------------
CURRENT_YEAR = today.year
is_incomplete_year = yearly["연도"] == CURRENT_YEAR
is_sparse_year = yearly["관측월수"] < MIN_MONTHS_PER_YEAR

yearly["제외사유"] = ""
yearly.loc[is_incomplete_year, "제외사유"] = "진행 중인 해"
yearly.loc[~is_incomplete_year & is_sparse_year, "제외사유"] = "관측치 부족"

train_df = yearly[yearly["제외사유"] == ""].copy()
excluded_df = yearly[yearly["제외사유"] != ""].copy()

if len(train_df) < 2:
    st.error(
        "학습에 쓸 수 있는 연도가 2개 미만이에요. (관측치 부족 또는 데이터 자체가 너무 적어요.) "
        "다른 과일을 선택해보세요."
    )
    st.stop()

# ----------------------------------------------------------------------------
# 5) scikit-learn 선형회귀로 "연도 -> 평균가격" 관계를 학습한다.
#    X(입력)는 연도, y(정답)는 그 해의 평균가격이다.
# ----------------------------------------------------------------------------
X_train = train_df["연도"].to_numpy().reshape(-1, 1)  # sklearn은 2차원 입력을 원한다
y_train = train_df["평균가격"].to_numpy()

model = LinearRegression()
model.fit(X_train, y_train)

y_pred_on_train = model.predict(X_train)
r2 = r2_score(y_train, y_pred_on_train)

train_min_year = int(train_df["연도"].min())
train_max_year = int(train_df["연도"].max())

st.divider()

# ----------------------------------------------------------------------------
# 6) R²(결정계수) 카드
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="metric-card">
        <div class="label">R² (결정계수) · {selected_fruit} · {price_basis_label} 기준</div>
        <div class="value">{r2:.3f}</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption(
    "💡 R²은 0~1 사이 값으로(음수가 나올 수도 있어요), **직선 하나만으로 연도별 가격 변화를 "
    "얼마나 잘 설명하는지**를 나타내요. 1에 가까울수록 잘 맞고, 0에 가까울수록(또는 음수라면) "
    "'연도'만으로는 가격을 설명하기 어렵다는 뜻이에요."
)

st.write("")

# ----------------------------------------------------------------------------
# 7) 슬라이더로 연도를 고르면 예측 가격을 보여준다.
# ----------------------------------------------------------------------------
st.subheader("🔮 연도를 선택해서 예측 가격을 확인해보세요")
target_year = st.slider(
    "연도 선택 (1900년 ~ 2100년)",
    min_value=1900,
    max_value=2100,
    value=CURRENT_YEAR,
    step=1,
)

predicted_price = float(model.predict(np.array([[target_year]]))[0])

st.markdown(
    f"""
    <div class="predict-card">
        <div class="label">{target_year}년 {selected_fruit} 예측 평균가격</div>
        <div class="value">{predicted_price:,.0f}원</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# 학습 범위(예: 실제로 관측된 연도들) 밖으로 나가면 "참고용, 조심" 안내를 띄운다.
if target_year < train_min_year or target_year > train_max_year:
    st.warning(
        f"⚠️ **참고용이니 조심해서 봐주세요!** 이 예측은 실제로 학습에 쓰인 연도 범위"
        f"({train_min_year}년 ~ {train_max_year}년) 밖이에요. 선형회귀는 그 범위를 벗어날수록 "
        "빗나갈 가능성이 커지고, 특히 아주 먼 과거나 먼 미래는 숫자 자체가 비현실적일 수 있어요.",
        icon="⚠️",
    )
    if predicted_price <= 0:
        st.caption(
            "😅 예측값이 0원 이하로 나왔어요. 먼 과거로 갈수록 직선이 계속 내려가기만 하는 "
            "선형회귀의 한계가 잘 드러나는 부분이에요."
        )

st.write("")

# ----------------------------------------------------------------------------
# 8) 실제 데이터 점 + 학습한 직선을 Plotly 그래프로 함께 보여준다.
# ----------------------------------------------------------------------------
# 그래프에 그릴 직선의 x축 범위: 학습 범위와 슬라이더로 고른 연도를 모두 포함하도록 넉넉히 잡는다.
line_x_min = min(train_min_year, target_year) - 3
line_x_max = max(train_max_year, target_year) + 3
line_x_min = max(line_x_min, 1900)
line_x_max = min(line_x_max, 2100)
line_years = np.linspace(line_x_min, line_x_max, 100)
line_prices = model.predict(line_years.reshape(-1, 1))

fig = go.Figure()

# 학습에 사용한 연도 범위를 은은한 초록 배경으로 표시 (main.py의 하이라이트 스타일 참고)
fig.add_vrect(
    x0=train_min_year - 0.5, x1=train_max_year + 0.5,
    fillcolor="#22c55e", opacity=0.07, line_width=0,
)
fig.add_annotation(
    x=train_min_year, y=1.06, xref="x", yref="paper", xanchor="left",
    text="학습에 사용한 연도 범위", showarrow=False,
    font=dict(size=11, color="#16a34a"),
)

# 회귀 직선
fig.add_trace(
    go.Scatter(
        x=line_years, y=line_prices, mode="lines", name="학습한 선형회귀 직선",
        line=dict(color="#8b5cf6", width=2.5),
    )
)

# 학습에 실제로 쓰인 연도별 평균가격 점
fig.add_trace(
    go.Scatter(
        x=train_df["연도"], y=train_df["평균가격"], mode="markers", name="실제 연도별 평균가격",
        marker=dict(size=10, color="#f2938a", line=dict(color="#e0665a", width=1)),
        hovertemplate="%{x}년<br>평균가격: %{y:,.0f}원<extra></extra>",
    )
)

# 제외된 연도(진행 중인 해 / 관측치 부족)도 참고용으로 회색 X 표시
if not excluded_df.empty:
    fig.add_trace(
        go.Scatter(
            x=excluded_df["연도"], y=excluded_df["평균가격"], mode="markers",
            name="제외된 연도 (진행중/관측치 부족)",
            marker=dict(size=9, color="#9ca3af", symbol="x"),
            hovertext=excluded_df["제외사유"],
            hovertemplate="%{x}년 (%{hovertext})<br>평균가격: %{y:,.0f}원<extra></extra>",
        )
    )

# 슬라이더로 고른 연도의 예측값을 별 모양으로 강조
fig.add_trace(
    go.Scatter(
        x=[target_year], y=[predicted_price], mode="markers", name=f"{target_year}년 예측값",
        marker=dict(size=16, color="#ef4444", symbol="star", line=dict(color="#ffffff", width=1)),
        hovertemplate=f"{target_year}년 예측<br>%{{y:,.0f}}원<extra></extra>",
    )
)

fig.update_layout(
    height=480,
    margin=dict(l=0, r=0, t=44, b=0),
    plot_bgcolor="rgba(255,255,255,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    xaxis_title="연도",
    yaxis_title=f"평균가격(원, {price_basis_label} 기준)",
    legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
    hovermode="closest",
)

st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

with st.expander("🧾 연도별 원본 데이터 보기"):
    display_df = yearly.copy()
    display_df["평균가격"] = display_df["평균가격"].round(0).astype(int)
    display_df["포함 여부"] = display_df["제외사유"].apply(lambda x: "학습에 사용" if x == "" else f"제외 ({x})")
    st.dataframe(
        display_df[["연도", "평균가격", "관측월수", "포함 여부"]],
        use_container_width=True,
        hide_index=True,
    )
