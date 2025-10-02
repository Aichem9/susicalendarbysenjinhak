import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.parser import parse
from streamlit_calendar import calendar

st.set_page_config(page_title="수시 일정 캘린더", layout="wide")

st.title("수시 지원/발표 일정 캘린더")
st.caption("엑셀 업로드 → 10·11·12월 달력에 자동 표시. 이벤트 클릭 시 반/이름/V열 값 표시.")

uploaded = st.file_uploader("엑셀 파일(.xlsx/.xls)을 업로드하세요 (헤더는 3행)", type=["xlsx", "xls"])

# 달력 공통 옵션(FullCalendar)
def fc_options(initial_date):
    return {
        "initialView": "dayGridMonth",
        "locale": "ko",
        "firstDay": 0,
        "height": 720,
        "initialDate": initial_date,      # "YYYY-MM-01"
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
def load_df(file):
    # header=2 => 3행이 헤더(1-indexed)
    df = pd.read_excel(file, header=2, dtype=str)
    return df

def safe_date(x):
    """엑셀의 날짜/문자/공백을 모두 안전하게 날짜로 파싱. 실패 시 None."""
    if pd.isna(x) or str(x).strip() == "":
        return None
    try:
        # 엑셀 직렬 날짜가 문자열로 들어오면 float 캐스팅 시도
        # (단, 이미 "2025-10-03" 같은 문자열이면 parse가 처리)
        return parse(str(x)).date()
    except Exception:
        # 엑셀에서 숫자(직렬)로 올 가능성도 재시도
        try:
            if isinstance(x, (int, float)):
                # pandas가 이미 날짜로 읽었다면 위에서 끝났을 것.
                # 여긴 진짜 직렬일 가능성 낮음. 보수적으로 실패 처리.
                return None
            return None
        except Exception:
            return None

def two_kor(s, n):
    """문자열 s에서 앞 n글자(한글도 글자 단위로) 안전 추출."""
    if pd.isna(s):
        return ""
    s = str(s).strip()
    return s[:n]

def extract_class_from_A(a_value):
    """A열 숫자/문자에서 두 번째, 세 번째 자리(1-indexed)를 '반'으로 사용."""
    if pd.isna(a_value):
        return ""
    s = "".join(str(a_value).strip().split())  # 공백 제거
    if len(s) >= 3:
        return s[1:3]
    return s  # 길이가 짧으면 있는 만큼

def build_events(df, target_year=None):
    """
    df에서 O(전형일), P(1단계 발표), Q(최종 발표)로 이벤트 생성.
    규칙:
      - O: "반/이름/대학2글자/전형종류"
      - P: "반/이름/대학3글자/전형종류"
    클릭 시: "반 / 이름 / V열 숫자"를 하단에 보여줄 수 있도록 extendedProps에 저장.
    """
    events = []

    # 컬럼 인덱스(0-indexed 고정): A=0, B=1, D=3, N=13, O=14, P=15, Q=16, V=21
    COL_A, COL_B, COL_D = 0, 1, 3
    COL_N, COL_O, COL_P, COL_Q, COL_V = 13, 14, 15, 16, 21

    # 타겟 연도 추정: 날짜 컬럼에서 가장 많이 등장하는 연도 또는 현재 연도
    if target_year is None:
        years = []
        for col in [COL_O, COL_P, COL_Q]:
            for x in df.iloc[:, col].dropna().tolist():
                d = safe_date(x)
                if d:
                    years.append(d.year)
        if years:
            # 최빈/최신 중 하나를 택할 수 있으나, 우선 최신 연도로
            target_year = max(years)
        else:
            target_year = datetime.now().year

    for idx, row in df.iterrows():
        try:
            a = row.iloc[COL_A]
            ban = extract_class_from_A(a)  # 반(두 번째/세 번째 자리)
            name = str(row.iloc[COL_B]).strip() if not pd.isna(row.iloc[COL_B]) else ""
            univ = str(row.iloc[COL_D]).strip() if not pd.isna(row.iloc[COL_D]) else ""
            typ = str(row.iloc[COL_N]).strip() if not pd.isna(row.iloc[COL_N]) else ""
            vval = str(row.iloc[COL_V]).strip() if not pd.isna(row.iloc[COL_V]) else ""

            # 전형일(O): 대학 2글자
            o_date = safe_date(row.iloc[COL_O])
            if o_date and o_date.year == target_year:
                title = f"{ban}/{name}/{two_kor(univ, 2)}/{typ}"
                events.append({
                    "title": title,
                    "start": o_date.isoformat(),
                    "allDay": True,
                    "extendedProps": {
                        "detail": f"{ban} / {name} / {vval}",
                        "cat": "전형일",
                    }
                })

            # 1단계 발표(P): 대학 3글자
            p_date = safe_date(row.iloc[COL_P])
            if p_date and p_date.year == target_year:
                title = f"{ban}/{name}/{two_kor(univ, 3)}/{typ}"
                events.append({
                    "title": title,
                    "start": p_date.isoformat(),
                    "allDay": True,
                    "extendedProps": {
                        "detail": f"{ban} / {name} / {vval}",
                        "cat": "1단계 발표",
                    }
                })

            # 최종 발표(Q): (요청 규칙엔 표시 문구가 특정되지 않았지만,
            # 달력에서 같이 보고 싶을 수 있어 동일 포맷으로 추가합니다.
            # 필요 없으면 아래 블록을 주석 처리하세요.)
            q_date = safe_date(row.iloc[COL_Q])
            if q_date and q_date.year == target_year:
                title = f"{ban}/{name}/{two_kor(univ, 3)}/{typ}"
                events.append({
                    "title": title,
                    "start": q_date.isoformat(),
                    "allDay": True,
                    "extendedProps": {
                        "detail": f"{ban} / {name} / {vval}",
                        "cat": "최종 발표",
                    }
                })

        except Exception:
            continue

    return events, target_year

def filter_month_events(events, year, month):
    mm = f"{year:04d}-{month:02d}"
    return [ev for ev in events if str(ev["start"]).startswith(mm)]

# ---------------- App Flow ----------------
if uploaded is None:
    st.info("예시: 헤더가 3행에 있고, A/B/D/N/O/P/Q/V 열이 존재하는 엑셀을 올려주세요.")
    st.stop()

df = load_df(uploaded)
events_all, yr = build_events(df)

# 10, 11, 12월 탭
tabs = st.tabs([f"{yr}년 10월", f"{yr}년 11월", f"{yr}년 12월"])

clicked_detail_placeholder = st.empty()

for tab, month in zip(tabs, [10, 11, 12]):
    with tab:
        month_events = filter_month_events(events_all, yr, month)

        # 초기 날짜는 해당 월의 1일
        opts = fc_options(f"{yr}-{month:02d}-01")

        # streamlit-calendar는 클릭 결과를 state로 반환
        state = calendar(events=month_events, options=opts, key=f"cal-{yr}-{month}")

        # 클릭된 이벤트 처리
        if state and "eventClick" in state and state["eventClick"]:
            info = state["eventClick"]["event"]
            ext = info.get("extendedProps", {})
            detail = ext.get("detail", "")
            cat = ext.get("cat", "")
            clicked_detail_placeholder.success(f"[{cat}] {detail}")

# 겹치는 일정(동일 날짜 다수)도 FullCalendar가 같은 날에 여러 이벤트로 자연스럽게 표시합니다.
st.caption("※ 동일 날짜에 여러 학생 일정이 겹치면 한 화면에 함께 표시됩니다.")
