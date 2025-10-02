import io
import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.parser import parse
from streamlit_calendar import calendar
# ... (생략: 나머지 코드는 기존 그대로) ...

@st.cache_data(show_spinner=False)
def load_df(file_bytes: bytes, filename: str):
    """
    업로드 파일 바이트와 파일명으로 DataFrame 로드.
    - 헤더는 3행(1-indexed) -> header=2
    - .xlsx: openpyxl
    - .xls : xlrd(1.2.0) 필요
    """
    ext = (filename or "").lower().strip().split(".")[-1]

    bio = io.BytesIO(file_bytes)

    if ext == "xlsx":
        df = pd.read_excel(bio, header=2, dtype=str, engine="openpyxl")
    elif ext == "xls":
        try:
            df = pd.read_excel(bio, header=2, dtype=str, engine="xlrd")
        except ImportError:
            # 안전한 안내 메시지
            raise RuntimeError(
                "'.xls' 파일을 읽으려면 xlrd==1.2.0이 필요합니다. "
                "requirements.txt에 'xlrd==1.2.0'을 추가한 뒤 다시 배포하세요."
            )
    else:
        raise RuntimeError("지원하지 않는 확장자입니다. .xls 또는 .xlsx 파일을 업로드하세요.")

    return df

# ---------------- App Flow ----------------
st.set_page_config(page_title="수시 일정 캘린더", layout="wide")
st.title("수시 지원/발표 일정 캘린더")
st.caption("엑셀 업로드 → 10·11·12월 달력에 자동 표시. 이벤트 클릭 시 반/이름/V열 값 표시.")

uploaded = st.file_uploader("엑셀 파일(.xlsx/.xls)을 업로드하세요 (헤더는 3행)", type=["xlsx", "xls"])

if uploaded is None:
    st.info("예시: 헤더가 3행에 있고, A/B/D/N/O/P/Q/V 열이 존재하는 엑셀을 올려주세요.")
    st.stop()

# ✅ 파일을 바이트로 읽어 캐시/리로드 이슈 방지
file_bytes = uploaded.read()
df = load_df(file_bytes, uploaded.name)

# (이 아래로는 기존 build_events / 캘린더 표시 로직 그대로 사용)
# events_all, yr = build_events(df)
# ...
