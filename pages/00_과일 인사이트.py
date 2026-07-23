"""
과일 가격 AI 챗봇 (Streamlit 멀티페이지의 서브페이지)
------------------------------------------------------------
왼쪽에는 미리 정해둔 질문 버튼, 오른쪽에는 채팅창을 두고 Upstage Solar(모델: solar-open2,
OpenAI 호환 API)로 답변을 생성한다.

- API 키는 코드에 직접 쓰지 않고 secrets의 SOLAR_API_KEY에서 불러온다.
- 모델명은 "solar-open2"를 그대로 사용한다.
- 응답 속도를 위해 temperature 대신 reasoning_effort="none"으로 추론을 끈다.
- 답변은 스트리밍(st.write_stream)으로 실시간으로 흘러나오게 표시한다.
- 대화 기록은 st.session_state에 저장해 이전 대화를 기억하며 이어간다.
- 시스템 프롬프트: "너는 따뜻하고 친절한 데이터 분석 선생님이야. 반드시 순수 한국어로만 답해."
- API 호출이 실패하면(키 미등록, 네트워크 오류, openai 미설치 등) 에러 화면 대신 한국어
  안내 메시지를 채팅 말풍선으로 보여준다.
"""

import streamlit as st

try:
    from openai import OpenAI  # Solar(solar-open2) 호출용. 없어도 앱은 안내 메시지로 정상 동작.
except ImportError:
    OpenAI = None

st.set_page_config(page_title="과일 가격 AI 챗봇", page_icon="💬", layout="wide")

SOLAR_BASE_URL = "https://api.upstage.ai/v1"
SOLAR_MODEL = "solar-open2"  # 모델명 그대로 사용
SYSTEM_PROMPT = "너는 따뜻하고 친절한 데이터 분석 선생님이야. 반드시 순수 한국어로만 답해."

FALLBACK_ANSWER = (
    "죄송해요, 지금은 답변을 가져오지 못했어요. 🙏 SOLAR_API_KEY가 제대로 등록되어 있는지, "
    "네트워크 연결이 원활한지 확인한 뒤 다시 시도해 주세요."
)

PRESET_QUESTIONS = [
    "💡 왜 겨울 수박은 비쌀까?",
    "🍎 사과는 언제 가장 저렴할까?",
    "🍊 감귤 가격은 왜 매년 비슷할까?",
    "🌧️ 비가 많이 오면 과일값은 얼마나 오를까?",
    "🥭 수입과일은 환율 영향을 얼마나 받을까?",
    "📦 도매가격과 소매가격은 왜 다를까?",
    "🌡️ 폭염이 과일 가격에 미치는 영향",
    "🚜 올해 생산량이 늘어나면 가격은 어떻게 될까?",
    "💰 가장 가성비 좋은 제철 과일은?",
    "🥝 냉해가 오면 어떤 과일 가격이 오를까?",
]

# ----------------------------------------------------------------------------
# 스타일 (메인 페이지와 통일된 살몬/화이트 카드 톤앤매너)
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    :root {
        --accent: #f2938a;
        --accent-dark: #e0665a;
        --ink: #241f2b;
        --muted: #9b93a3;
    }
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(180deg, #fdf1f0 0%, #fbe8e6 100%);
    }
    [data-testid="stHeader"] { background: rgba(0,0,0,0); }
    .chat-title {
        font-size: 1.8rem;
        font-weight: 800;
        color: var(--ink);
        margin-bottom: 0.2rem;
    }
    .chat-sub {
        color: var(--muted);
        margin-bottom: 1.2rem;
    }
    .preset-card {
        background: #ffffff;
        border-radius: 20px;
        padding: 1rem 1.1rem;
        box-shadow: 0 8px 20px rgba(36,31,43,0.07);
        margin-bottom: 1rem;
    }
    .st-key-preset_list .stButton > button {
        border-radius: 14px;
        border: 1px solid rgba(36,31,43,0.06);
        background-color: #fff7f6;
        color: var(--ink);
        font-weight: 600;
        text-align: left;
        padding: 0.6rem 0.8rem;
        box-shadow: none;
        white-space: normal;
        line-height: 1.4;
    }
    .st-key-preset_list .stButton > button:hover {
        background-color: var(--accent);
        color: #ffffff;
        border-color: var(--accent);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# 세션 상태
# ----------------------------------------------------------------------------
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []  # [{"role": "user"/"assistant", "content": str}, ...]
if "pending_question" not in st.session_state:
    st.session_state.pending_question = None


def _ask_preset(question: str) -> None:
    st.session_state.pending_question = question


def _stream_solar_reply(messages: list):
    """Solar(solar-open2)에 스트리밍으로 질문하고, 조각(delta) 텍스트를 순서대로 내보낸다.

    키 미등록/openai 미설치/네트워크 오류 등 어떤 이유로든 실패하면, 예외를 던지는 대신
    한국어 안내 메시지 한 조각만 내보내고 끝낸다 (호출부에서 에러 화면 없이 자연스럽게
    말풍선에 안내 문구가 표시된다).
    """
    try:
        api_key = st.secrets["SOLAR_API_KEY"]
    except Exception:
        yield FALLBACK_ANSWER
        return
    if OpenAI is None:
        yield FALLBACK_ANSWER
        return

    try:
        client = OpenAI(api_key=api_key, base_url=SOLAR_BASE_URL)
        stream = client.chat.completions.create(
            model=SOLAR_MODEL,
            reasoning_effort="none",
            messages=messages,
            stream=True,
        )
        got_any = False
        for chunk in stream:
            delta = chunk.choices[0].delta.content if chunk.choices else None
            if delta:
                got_any = True
                yield delta
        if not got_any:
            yield FALLBACK_ANSWER
    except Exception:
        yield FALLBACK_ANSWER


st.markdown('<div class="chat-title">💬 과일 가격 AI 챗봇</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="chat-sub">궁금한 걸 왼쪽에서 골라 클릭하거나, 오른쪽 채팅창에 직접 물어보세요.</div>',
    unsafe_allow_html=True,
)

left, right = st.columns([1, 2], gap="medium")

with left:
    st.markdown('<div class="preset-card">', unsafe_allow_html=True)
    st.markdown("**자주 묻는 질문**")
    with st.container(key="preset_list"):
        for i, q in enumerate(PRESET_QUESTIONS):
            st.button(
                q, key=f"preset_{i}", use_container_width=True,
                on_click=_ask_preset, args=(q,),
            )
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    chat_box = st.container(height=520, border=True)
    with chat_box:
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    typed_question = st.chat_input("과일 가격에 대해 무엇이든 물어보세요")
    question = typed_question or st.session_state.pending_question
    st.session_state.pending_question = None

    if question:
        st.session_state.chat_history.append({"role": "user", "content": question})
        with chat_box:
            with st.chat_message("user"):
                st.markdown(question)
            with st.chat_message("assistant"):
                api_messages = [{"role": "system", "content": SYSTEM_PROMPT}]
                api_messages.extend(
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.chat_history
                )
                answer = st.write_stream(_stream_solar_reply(api_messages))
        st.session_state.chat_history.append({"role": "assistant", "content": answer})
