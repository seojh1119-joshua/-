import streamlit as st

# 자동 번역 충돌 방지 태그 삽입
st.markdown(
    '<meta name="google" content="notranslate">', unsafe_allow_html=True
)

import calendar
from datetime import date, timedelta
from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="당직표 관리 시스템", layout="wide")
st.title("숙직 당직표 관리 앱 (Excel 연동)")

if "staff_df" not in st.session_state:
    st.session_state.staff_df = pd.DataFrame(
        columns=["id", "name", "role", "max_per_month"]
    )

if "schedule_df" not in st.session_state:
    st.session_state.schedule_df = pd.DataFrame(
        columns=["date", "day", "duty_type", "staff_id", "status", "memo"]
    )

if "holidays_df" not in st.session_state:
    st.session_state.holidays_df = pd.DataFrame(columns=["date", "name"])


def to_excel_bytes(staff_df, schedule_df, holidays_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        staff_df.to_excel(writer, index=False, sheet_name="Staff")
        schedule_df.to_excel(writer, index=False, sheet_name="Schedule")
        holidays_df.to_excel(writer, index=False, sheet_name="Holidays")
    return output.getvalue()


def load_excel(file):
    try:
        xls = pd.ExcelFile(file)
        staff = (
            pd.read_excel(xls, "Staff")
            if "Staff" in xls.sheet_names
            else pd.DataFrame()
        )
        sched = (
            pd.read_excel(xls, "Schedule")
            if "Schedule" in xls.sheet_names
            else pd.DataFrame()
        )
        holi = (
            pd.read_excel(xls, "Holidays")
            if "Holidays" in xls.sheet_names
            else pd.DataFrame()
        )
        return staff, sched, holi
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")
        return None, None, None


def generate_schedule(year, month, staff_df, holidays_df):
    days = calendar.monthrange(year, month)[1]
    records = []
    staff_ids = staff_df["id"].tolist()
    idx = 0

    for d in range(1, days + 1):
        dt = date(year, month, d)
        day_name = dt.strftime("%a")
        is_holiday = (
            dt in holidays_df["date"].values or day_name in ["Sat", "Sun"]
        )

        for dtype in (
            ["휴일주간", "휴일야간"] if is_holiday else ["주간", "야간"]
        ):
            if not staff_ids:
                break
            staff_id = staff_ids[idx % len(staff_ids)]
            idx += 1
            records.append(
                {
                    "date": dt,
                    "day": day_name,
                    "duty_type": dtype,
                    "staff_id": staff_id,
                    "status": "자동생성",
                    "memo": "",
                }
            )

    return pd.DataFrame(records)


with st.sidebar:
    st.header("파일 관리")
    up = st.file_uploader("엑셀 업로드 (.xlsx)", type=["xlsx"])
    if up:
        staff, sched, holi = load_excel(up)
        if staff is not None:
            st.session_state.staff_df = staff
            st.session_state.schedule_df = sched
            st.session_state.holidays_df = holi
            st.success("불러오기 완료!")

    st.download_button(
        "현재 데이터 엑셀로 저장",
        data=to_excel_bytes(
            st.session_state.staff_df,
            st.session_state.schedule_df,
            st.session_state.holidays_df,
        ),
        file_name=f"duty_schedule_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

tab1, tab2, tab3, tab4 = st.tabs(
    ["근무자 관리", "근무표 편집", "자동 생성", "통계"]
)

with tab1:
    st.subheader("근무자 명단")
    edited_staff = st.data_editor(
        st.session_state.staff_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "id": st.column_config.TextColumn("ID", required=True),
            "name": st.column_config.TextColumn("이름", required=True),
            "role": st.column_config.SelectboxColumn(
                "역할", options=["팀장", "팀원", "신입"]
            ),
            "max_per_month": st.column_config.NumberColumn(
                "월 최대근무", min_value=0, max_value=31, default=8
            ),
        },
        key="staff_editor",
    )
    st.session_state.staff_df = edited_staff

with tab2:
    st.subheader("근무표 수동 편집")
    if st.session_state.schedule_df.empty:
        st.info("먼저 '자동 생성' 탭에서 초안을 만들거나 엑셀을 업로드하세요.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            year = st.number_input("연도", value=date.today().year, step=1)
        with col2:
            month = st.number_input(
                "월",
                value=date.today().month,
                min_value=1,
                max_value=12,
                step=1,
            )

        mask = (
            pd.to_datetime(st.session_state.schedule_df["date"]).dt.year
            == year
        ) & (
            pd.to_datetime(st.session_state.schedule_df["date"]).dt.month
            == month
        )
        view_df = st.session_state.schedule_df[mask].copy()

        name_map = dict(
            zip(
                st.session_state.staff_df["id"],
                st.session_state.staff_df["name"],
            )
        )
        view_df["staff_name"] = (
            view_df["staff_id"].map(name_map).fillna(view_df["staff_id"])
        )

        edited = st.data_editor(
            view_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "date": st.column_config.DateColumn("날짜", disabled=True),
                "day": st.column_config.TextColumn("요일", disabled=True),
                "duty_type": st.column_config.SelectboxColumn(
                    "근무유형",
                    options=[
                        "주간",
                        "야간",
                        "휴일주간",
                        "휴일야간",
                        "비번",
                        "휴무",
                    ],
                ),
                "staff_id": st.column_config.SelectboxColumn(
                    "근무자ID",
                    options=st.session_state.staff_df["id"].tolist(),
                ),
                "staff_name": st.column_config.TextColumn(
                    "이름", disabled=True
                ),
                "status": st.column_config.SelectboxColumn(
                    "상태", options=["확정", "대기", "변경요청", "취소"]
                ),
                "memo": st.column_config.TextColumn("비고"),
            },
            hide_index=True,
            key="schedule_editor",
        )

        if st.button("수정한 내용 반영"):
            edited = edited.drop(columns=["staff_name"])
            base = st.session_state.schedule_df[~mask]
            st.session_state.schedule_df = pd.concat(
                [base, edited], ignore_index=True
            )
            st.success("반영 완료! (엑셀 저장은 사이드바 버튼 사용)")

with tab3:
    st.subheader("자동 근무표 생성 (라운드로빈 데모)")
    col1, col2 = st.columns(2)
    with col1:
        gen_year = st.number_input(
            "생성 연도", value=date.today().year, step=1, key="gen_y"
        )
    with col2:
        gen_month = st.number_input(
            "생성 월",
            value=date.today().month,
            min_value=1,
            max_value=12,
            step=1,
            key="gen_m",
        )

    if st.button("이번 달 근무표 생성"):
        if st.session_state.staff_df.empty:
            st.warning("근무자 명단을 먼저 입력하세요. (탭 1)")
        else:
            new_sched = generate_schedule(
                gen_year,
                gen_month,
                st.session_state.staff_df,
                st.session_state.holidays_df,
            )
            mask = (
                pd.to_datetime(st.session_state.schedule_df["date"]).dt.year
                == gen_year
            ) & (
                pd.to_datetime(st.session_state.schedule_df["date"]).dt.month
                == gen_month
            )
            base = st.session_state.schedule_df[~mask]
            st.session_state.schedule_df = pd.concat(
                [base, new_sched], ignore_index=True
            )
            st.success(
                f"{gen_year}년 {gen_month}월 근무표 생성 완료! '근무표 편집' 탭에서 확인/수정하세요."
            )

with tab4:
    st.subheader("월별 근무 통계")
    if st.session_state.schedule_df.empty:
        st.info("데이터가 없습니다.")
    else:
        df = st.session_state.schedule_df.copy()
        df["date"] = pd.to_datetime(df["date"])
        df["year_month"] = df["date"].dt.to_period("M").astype(str)
        df = df.merge(
            st.session_state.staff_df[["id", "name"]],
            left_on="staff_id",
            right_on="id",
            how="left",
        )

        ym = st.selectbox(
            "기준 연월", sorted(df["year_month"].unique(), reverse=True)
        )
        sub = df[df["year_month"] == ym]

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 개인별 근무 횟수")
            cnt = sub.groupby("name").size().reset_index(name="count")
            st.bar_chart(cnt.set_index("name"))

        with col2:
            st.markdown("### 근무 유형별 분포")
            dtype_cnt = sub["duty_type"].value_counts().reset_index()
            dtype_cnt.columns = ["duty_type", "count"]
            fig = px.pie(
                dtype_cnt,
                names="duty_type",
                values="count",
                title="근무 유형 비율",
            )
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 공정성 지표 (표준편차)")
        std = cnt["count"].std()
        st.metric("근무 횟수 표준편차", f"{std:.2f} (낮을수록 공정)")
        st.caption(
            "※ 2026 숙직표 관리 앱 Streamlit 프로토타입 (확장 시 FastAPI + React 권장)"
        )
