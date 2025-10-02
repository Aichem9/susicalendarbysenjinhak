import io
import importlib
import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.parser import parse
from streamlit_calendar import calendar

st.set_page_config(page_title="수시 일정 캘린더", layout="wide")

st.title("수시 지원/발표 일정 캘린더")

# ▲ 범례(legend)
st.markdown(
    """
**범례**
- <span style='color:blue'>■</span> 면접
- <span style='color:purple'>■</span> 논술
- <span style='color:yellow'>■</span> 1차 발표
- <span style='color:green'>■</span> 2차 발표(최종)
""",
    unsafe_allow_html=True,
)

st.caption(
    "엑셀 업로드 → 10·11·12월 달력에 자동 표시. 이벤트 클릭 시 반/이름/V열 값 표시.\n"
    "senjinhak 수시지원 결과 파일을 다운로드하여 사용하세요."
)

uploaded = st.file_uploader("엑셀 파일(.xlsx/.xls)을 업로드하세요 (헤더는 3행)", type=["xlsx", "xls"])

def safe_date(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    try:
        return parse(str(x)).date()
    except Exception:
        return None

def two_kor(s, n):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    return s[:n]

def extract_class_from_A(a_value):
    if pd.isna(a_value):
        return ""
    s = "".join(str(a_value).strip().split())
    if len(s) >= 3:
        return s[1:3]
    return s

def fc_options(initial_date):
    return {
        "initialView": "dayGridMonth",
        "locale": "ko",
        "firstDay": 0,
        "height": 720,
        "initialDate": initial_date,
        "headerToolbar": {
            "left": "",
            "center": "title",
            "right": ""
        },
        "dayMaxEventRows": True,
        "fixedWeekCount": False,
        "eventDisplay": "block",
        "eventOrder": "title,start"
    }

@st.cache_data(show_spinner=False)
def load_df(file_bytes: bytes, filename: str):
    ext = (filename or "").lower().strip().split(".")[-1]
    bio = io.BytesIO(file_bytes)

    if ext == "xlsx":
        return pd.read_excel(bio, header=2, dtype=str, engine="openpyxl")

    if ext == "xls":
        try:
            importlib.import_module("xlrd")
            return pd.read_excel(bio, header=2, dtype=str, engine="xlrd")
        except Exception:
            st.error(
                "현재 환경에서는 `.xls` 파일을 읽을 수 없습니다.\n"
                "➡️ 해결 방법:\n"
                "1) 파일을 `.xlsx`로 저장해서 다시 업로드하시거나\n"
                "2) runtime.txt를 Python 3.11로 설정하고, requirements.txt에 xlrd==1.2.0 추가 후 재배포하세요."
            )
            st.stop()

    st.error("지원하지 않는 확장자입니다. .xls 또는 .xlsx 파일을 업로드하세요.")
    st.stop()

def build_events(df, target_year=None):
    events = []
    COL_A, COL_B, COL_D = 0, 1, 3
    COL_N, COL_O, COL_P, COL_Q, COL_V = 13, 14, 15, 16, 21

    if target_year is None:
        years = []
        for col in [COL_O, COL_P, COL_Q]:
            for x in df.iloc[:, col].dropna().tolist():
                d = safe_date(x)
                if d:
                    years.append(d.year)
        if years:
            target_year = max(years)
        else:
            target_year = datetime.now().year

    for _, row in df.iterrows():
        try:
            a = row.iloc[COL_A]
            ban = extract_class_from_A(a)
            name = str(row.iloc[COL_B]).strip() if not pd.isna(row.iloc[COL_B]) else ""
            univ = str(row.iloc[COL_D]).strip() if not pd.isna(row.iloc[COL_D]) else ""
            typ = str(row.iloc[COL_N]).strip() if not pd.isna(row.iloc[COL_N]) else ""
            vval = str(row.iloc[COL_V]).strip() if not pd.isna(row.iloc[COL_V]) else ""

            # 전형일(O열) - 면접/논술 색상
            o_date = safe_date(row.iloc[COL_O])
            if o_date and o_date.year == target_year:
                title = f"{ban}/{name}/{two_kor(univ, 2)}/{typ}"
                color = None
                if "면접" in typ:
                    color = "blue"
                elif "논술" in typ:
                    color = "purple"
                events.append({
                    "title": title,
                    "start": o_date.isoformat(),
                    "allDay": True,
                    "color": color,
                    "extendedProps": {
                        "detail": f"{ban} / {name} / {vval}",
                        "cat": "전형일",
                    }
                })

            # 1단계 발표(P열) - 노란색
            p_date = safe_date(row.iloc[COL_P])
            if p_date and p_date.year == target_year:
                title = f"{ban}/{name}/{two_kor(univ, 3)}/{typ}"
                events.append({
                    "title": title,
                    "start": p_date.isoformat(),
                    "allDay": True,
                    "color": "yellow",
                    "extendedProps": {
                        "detail": f"{ban} / {name} / {vval}",
                        "cat": "1단계 발표",
                    }
                })

            # 최종 발표(Q열 = 2차 발표) - 초록색
            q_date = safe_date(row.iloc[COL_Q])
            if q_date and q_date.year == target_year:
                title = f"{ban}/{name}/{two_kor(univ, 3)}/{typ}"
                events.append({
                    "title": title,
                    "start": q_date.isoformat(),
                    "allDay": True,
                    "color": "green",
                    "extendedProps": {
                        "detail": f"{ban} / {name} / {vval}",
                        "cat": "최종 발표",
                    }
                })

        except Exception:
            continue

    return events, target_year

def filter_month_events(events, year, month):
    filtered = []
    for ev in events:
        try:
            d = datetime.fromisoformat(ev["start"])
            if d.year == year and d.month == month:
                filtered.append(ev)
        except Exception:
            continue
    return filtered

if uploaded is None:
    st.info("예시: 헤더가 3행에 있고, A/B/D/N/O/P/Q/V 열이 존재하는 엑셀을 올려주세요. "
            "senjinhak 수시지원 결과 파일을 다운로드하여 사용하세요.")
    st.stop()

file_bytes = uploaded.read()
df = load_df(file_bytes, uploaded.name)

events_all, yr = build_events(df)

clicked_detail_placeholder = st.empty()

# 10월, 11월, 12월을 스크롤로 이어서 출력
for month in [10, 11, 12]:
    st.subheader(f"{yr}년 {month}월")
    month_events = filter_month_events(events_all, yr, month)
    opts = fc_options(f"{yr}-{month:02d}-01")
    state = calendar(events=month_events, options=opts, key=f"cal-{yr}-{month}")
    if state and "eventClick" in state and state["eventClick"]:
        info = state["eventClick"]["event"]
        ext = info.get("extendedProps", {})
        detail = ext.get("detail", "")
        cat = ext.get("cat", "")
        clicked_detail_placeholder.success(f"[{cat}] {detail}")
    st.markdown("---")

st.caption("※ 동일 날짜에 여러 학생 일정이 겹치면 한 화면에 함께 표시됩니다. "
           "senjinhak 수시지원 결과 파일을 다운로드하여 사용하세요.")
