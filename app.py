import calendar
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# ─────────────────────────────────────────────
# [1] 기본 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="도시가스 공급량 패턴 분석 및 계획",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 단위 환산 상수
MJ_PER_NM3 = 42.563
MJ_TO_GJ = 1.0 / 1000.0

def mj_to_gj(x):
    return x * MJ_TO_GJ if pd.notna(x) else 0

# ─────────────────────────────────────────────
# [2] 데이터 로딩 (하이브리드 방식: 업로드 or 로컬)
# ─────────────────────────────────────────────
def load_data_file(file_uploader, local_path, file_type='excel'):
    """업로더에 파일이 있으면 그걸 쓰고, 없으면 로컬 경로를 시도함"""
    if file_uploader is not None:
        return pd.read_excel(file_uploader) if file_type == 'excel' else pd.read_csv(file_uploader)
    
    # 로컬 파일 시도
    import os
    if os.path.exists(local_path):
        return pd.read_excel(local_path) if file_type == 'excel' else pd.read_csv(local_path)
    
    return None

@st.cache_data(show_spinner=False)
def get_data(daily_file, plan_file, cal_file):
    # 1. 일일 실적 데이터
    df_daily = load_data_file(daily_file, "공급량(일일실적).xlsx")
    
    # 2. 월별 계획 데이터
    # 시트가 여러개일 수 있으므로 엑셀 객체로 로드 후 시트 파싱
    df_plan_raw = None
    if plan_file is not None:
        df_plan_raw = pd.read_excel(plan_file, sheet_name=None)
    else:
        import os
        if os.path.exists("공급량(계획_실적).xlsx"):
            df_plan_raw = pd.read_excel("공급량(계획_실적).xlsx", sheet_name=None)
    
    # 3. 휴일 캘린더
    df_cal = load_data_file(cal_file, "effective_days_calendar.xlsx")

    return df_daily, df_plan_raw, df_cal

# ─────────────────────────────────────────────
# [3] 데이터 전처리 및 분석 로직
# ─────────────────────────────────────────────
def preprocess_daily(df):
    if df is None: return None
    df = df.copy()
    # 컬럼 매핑 (혹시 모를 오타 방지)
    cols = [c for c in df.columns if "일자" in c or "DATE" in c.upper()]
    if not cols: return None
    date_col = cols[0]
    
    df = df.rename(columns={date_col: "일자"})
    df["일자"] = pd.to_datetime(df["일자"])
    df["연도"] = df["일자"].dt.year
    df["월"] = df["일자"].dt.month
    df["일"] = df["일자"].dt.day
    
    # 공급량 컬럼 찾기
    mj_col = [c for c in df.columns if "MJ" in c and "공급" in c][0]
    df = df.rename(columns={mj_col: "공급량(MJ)"})
    return df.dropna(subset=["공급량(MJ)"])

def preprocess_plan(df_dict):
    if df_dict is None: return None
    # '월별계획_실적' 시트 찾기
    sheet_name = [k for k in df_dict.keys() if "월별" in k][0]
    df = df_dict[sheet_name].copy()
    
    # 계획 컬럼 찾기 (숫자형이고 '계획' 포함된 첫번째 컬럼)
    plan_candidates = [c for c in df.columns if "계획" in c and pd.api.types.is_numeric_dtype(df[c])]
    plan_col = plan_candidates[0] if plan_candidates else df.columns[3] # fallback
    
    df = df.rename(columns={plan_col: "계획량(MJ)"})
    return df[["연", "월", "계획량(MJ)"]]

def preprocess_calendar(df):
    if df is None: return None
    # 날짜 컬럼 표준화
    date_col = [c for c in df.columns if "일자" in c or "날짜" in c][0]
    df = df.rename(columns={date_col: "일자"})
    df["일자"] = pd.to_datetime(df["일자"].astype(str), format="%Y%m%d", errors="coerce")
    
    for c in ["공휴일여부", "명절여부"]:
        if c not in df.columns: df[c] = False
        df[c] = df[c].fillna(False).astype(bool)
        
    return df[["일자", "공휴일여부", "명절여부"]]

# ─────────────────────────────────────────────
# [4] 핵심 로직: 패턴 분석 및 계획 수립
# ─────────────────────────────────────────────
def calculate_daily_plan(df_daily, df_plan, df_cal, target_year, target_month, window):
    # 1. 데이터 준비
    daily = preprocess_daily(df_daily)
    plan = preprocess_plan(df_plan)
    cal = preprocess_calendar(df_cal)
    
    if daily is None or plan is None:
        return None, "필수 데이터(일일실적 또는 월별계획)가 누락되었습니다."

    # 2. 과거 데이터 필터링 (최근 N년)
    start_year = target_year - window
    # 타겟 월과 같은 월만 추출
    past_data = daily[(daily["연도"] >= start_year) & 
                      (daily["연도"] < target_year) & 
                      (daily["월"] == target_month)].copy()
    
    if past_data.empty:
        return None, f"최근 {window}년 간 {target_month}월 실적 데이터가 없습니다."
    
    used_years = sorted(past_data["연도"].unique())

    # 3. 요일/휴일 속성 부여 (과거 데이터)
    past_data["weekday"] = past_data["일자"].dt.weekday # 0:월 ~ 6:일
    if cal is not None:
        past_data = past_data.merge(cal, on="일자", how="left").fillna(False)
    else:
        past_data["공휴일여부"] = False
        past_data["명절여부"] = False
    
    past_data["is_weekend"] = (past_data["weekday"] >= 5) | past_data["공휴일여부"] | past_data["명절여부"]
    
    # 4. n번째 요일 로직 (비율 산출용)
    past_data["nth_dow"] = past_data.sort_values("일").groupby(["연도", "weekday"]).cumcount() + 1
    
    # 월별 총량 대비 일별 비율 계산
    past_data["month_total"] = past_data.groupby("연도")["공급량(MJ)"].transform("sum")
    past_data["ratio"] = past_data["공급량(MJ)"] / past_data["month_total"]
    
    # 요일별/n번째별 평균 비율 산출 (평일/주말 구분)
    # 그룹: 평일(월금 / 화수목), 주말
    past_data["day_group"] = np.where(past_data["is_weekend"], "주말", 
                                      np.where(past_data["weekday"].isin([0,4]), "평일1(월금)", "평일2(화수목)"))
    
    # (요일, n번째) 키로 평균 비율 사전 생성
    ratio_map = past_data.groupby(["day_group", "weekday", "nth_dow"])["ratio"].mean().to_dict()
    # fallback용 (n번째 데이터 없을 때 요일 평균)
    ratio_fallback = past_data.groupby(["weekday"])["ratio"].mean().to_dict()

    # 5. 타겟 월 생성 및 적용
    last_day = calendar.monthrange(target_year, target_month)[1]
    dates = pd.date_range(f"{target_year}-{target_month:02d}-01", periods=last_day)
    target = pd.DataFrame({"일자": dates})
    target["일"] = target["일자"].dt.day
    target["weekday"] = target["일자"].dt.weekday
    
    # 타겟 월 휴일 적용
    if cal is not None:
        target = target.merge(cal, on="일자", how="left").fillna(False)
    else:
        target["공휴일여부"] = False
        target["명절여부"] = False
    
    target["is_weekend"] = (target["weekday"] >= 5) | target["공휴일여부"] | target["명절여부"]
    target["nth_dow"] = target.sort_values("일").groupby("weekday").cumcount() + 1
    target["day_group"] = np.where(target["is_weekend"], "주말", 
                                   np.where(target["weekday"].isin([0,4]), "평일1(월금)", "평일2(화수목)"))

    # 비율 매핑
    def get_ratio(row):
        key = (row["day_group"], row["weekday"], row["nth_dow"])
        return ratio_map.get(key, ratio_fallback.get(row["weekday"], 1/last_day))
    
    target["raw_ratio"] = target.apply(get_ratio, axis=1)
    
    # 비율 정규화 (합계 1)
    target["final_ratio"] = target["raw_ratio"] / target["raw_ratio"].sum()
    
    # 계획량 적용
    plan_row = plan[(plan["연"] == target_year) & (plan["월"] == target_month)]
    if plan_row.empty:
        return None, f"{target_year}년 {target_month}월 계획 데이터가 없습니다."
    
    total_plan_mj = plan_row["계획량(MJ)"].values[0]
    target["예상공급량(MJ)"] = (target["final_ratio"] * total_plan_mj).round(0)
    
    # ─────────────────────────────────────────────────────────────
    # [NEW] 아웃라이어 구간 설정 (주차 + 주말여부 그룹핑)
    # ─────────────────────────────────────────────────────────────
    target["WeekNum"] = target["일자"].dt.isocalendar().week
    
    # 그룹: [주차] + [주말여부]
    # 이렇게 하면 같은 주차라도 '평일'과 '주말'의 평균이 따로 계산됨 -> 계단식 상한선 구현
    target["Group_Mean"] = target.groupby(["WeekNum", "is_weekend"])["예상공급량(MJ)"].transform("mean")
    
    target["Upper_Bound"] = target["Group_Mean"] * 1.10
    target["Lower_Bound"] = target["Group_Mean"] * 0.90
    
    target["Is_Outlier"] = (target["예상공급량(MJ)"] > target["Upper_Bound"]) | \
                           (target["예상공급량(MJ)"] < target["Lower_Bound"])
                           
    return target, used_years

# ─────────────────────────────────────────────
# [5] 메인 UI
# ─────────────────────────────────────────────
def main():
    st.title("📊 도시가스 일별 계획 자동 수립")
    st.caption(f"Han형님의 마케팅 기획을 위한 맞춤형 대시보드입니다. (Outlier Check Ver.)")

    # 사이드바: 파일 업로드 및 설정
    with st.sidebar:
        st.header("1. 데이터 파일 설정")
        st.info("파일을 업로드하거나, 이미 서버에 있는 파일을 사용합니다.")
        
        up_daily = st.file_uploader("공급량(일일실적).xlsx", type=["xlsx", "csv"])
        up_plan = st.file_uploader("공급량(계획_실적).xlsx", type=["xlsx", "csv"])
        up_cal = st.file_uploader("effective_days_calendar.xlsx (선택)", type=["xlsx"])
        
        st.divider()
        st.header("2. 분석 조건 설정")
        col1, col2 = st.columns(2)
        with col1:
            t_year = st.number_input("목표 연도", 2024, 2030, 2026)
        with col2:
            t_month = st.selectbox("목표 월", range(1, 13))
            
        window = st.slider("과거 패턴 참조 기간(년)", 1, 10, 3)

    # 데이터 로드
    df_daily_raw, df_plan_raw, df_cal_raw = get_data(up_daily, up_plan, up_cal)

    if df_daily_raw is None or df_plan_raw is None:
        st.warning("👈 왼쪽 사이드바에서 데이터 파일을 업로드해주세요.")
        return

    # 분석 실행
    with st.spinner("패턴 분석 중..."):
        result_df, info_msg = calculate_daily_plan(
            df_daily_raw, df_plan_raw, df_cal_raw, t_year, t_month, window
        )

    if result_df is None:
        st.error(info_msg)
        return

    # 결과 표출
    st.subheader(f"✅ {t_year}년 {t_month}월 일별 공급계획 결과")
    if isinstance(info_msg, list):
        st.success(f"참조한 과거 데이터: {min(info_msg)}년 ~ {max(info_msg)}년 ({len(info_msg)}개년 평균)")

    # MJ -> GJ 변환
    display_df = result_df.copy()
    display_df["예상공급량(GJ)"] = display_df["예상공급량(MJ)"].apply(mj_to_gj)
    display_df["상한선(GJ)"] = display_df["Upper_Bound"].apply(mj_to_gj)
    display_df["하한선(GJ)"] = display_df["Lower_Bound"].apply(mj_to_gj)
    display_df["그룹평균(GJ)"] = display_df["Group_Mean"].apply(mj_to_gj)
    
    # 1. 그래프 그리기
    fig = go.Figure()

    # (1) 막대 그래프 (평일/주말 색상 구분)
    # 주말: 초록색 계열, 평일: 파란색 계열
    colors = np.where(display_df["is_weekend"], "#2ca02c", "#1f77b4")
    
    fig.add_trace(go.Bar(
        x=display_df["일"], y=display_df["예상공급량(GJ)"],
        marker_color=colors,
        name="일별 계획(GJ)",
        opacity=0.8
    ))

    # (2) 상한/하한선 (주중/주말 분리되어 계단식으로 표현됨)
    fig.add_trace(go.Scatter(
        x=display_df["일"], y=display_df["상한선(GJ)"],
        mode='lines', line=dict(width=0), showlegend=False, hoverinfo='skip'
    ))
    fig.add_trace(go.Scatter(
        x=display_df["일"], y=display_df["하한선(GJ)"],
        mode='lines', line=dict(width=0), 
        fill='tonexty', fillcolor='rgba(128, 128, 128, 0.2)',
        name='권장 범위(±10%)'
    ))

    # (3) 아웃라이어 표시 (빨간 X)
    outliers = display_df[display_df["Is_Outlier"]]
    if not outliers.empty:
        fig.add_trace(go.Scatter(
            x=outliers["일"], y=outliers["예상공급량(GJ)"],
            mode='markers',
            marker=dict(color='red', size=12, symbol='x'),
            name='범위 초과(Outlier)'
        ))

    fig.update_layout(
        title=f"{t_year}년 {t_month}월 일별 공급패턴 및 이상치 점검",
        xaxis_title="일 (Day)",
        yaxis_title="공급량 (GJ)",
        legend=dict(orientation="h", y=1.1),
        height=500,
        margin=dict(l=20, r=20, t=80, b=40)
    )
    st.plotly_chart(fig, use_container_width=True)

    # 2. 데이터 테이블 (아웃라이어 강조)
    st.markdown("#### 📋 상세 데이터 (Outlier 강조)")
    
    # 보여줄 컬럼 선택
    cols_show = ["일자", "day_group", "WeekNum", "예상공급량(GJ)", "상한선(GJ)", "하한선(GJ)", "Is_Outlier"]
    table_df = display_df[cols_show].copy()
    
    # 스타일링 함수
    def style_outlier(row):
        color = '#ffcccc' if row["Is_Outlier"] else ''
        return [f'background-color: {color}' for _ in row]

    # 숫자 포맷팅
    table_df["일자"] = table_df["일자"].dt.strftime("%Y-%m-%d")
    for c in ["예상공급량(GJ)", "상한선(GJ)", "하한선(GJ)"]:
        table_df[c] = table_df[c].apply(lambda x: f"{x:,.0f}")
    
    table_df["Is_Outlier"] = table_df["Is_Outlier"].map({True: "⚠️ 초과", False: "-"})
    
    st.dataframe(table_df.style.apply(style_outlier, axis=1), use_container_width=True, height=400)

    # 3. 엑셀 다운로드
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        display_df.to_excel(writer, index=False, sheet_name="일별계획_분석")
    
    st.download_button(
        label="📥 결과 엑셀 다운로드",
        data=output.getvalue(),
        file_name=f"{t_year}_{t_month}_일별계획_분석.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
