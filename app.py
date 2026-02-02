import calendar
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# ─────────────────────────────────────────────
# [설정] Han형님 맞춤형 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="도시가스 공급량: 일별 계획 및 보정 (Han's Version)",
    layout="wide",
)

# ─────────────────────────────────────────────
# 단위/환산 상수
# ─────────────────────────────────────────────
MJ_PER_NM3 = 42.563          # MJ / Nm3
MJ_TO_GJ = 1.0 / 1000.0      # MJ → GJ

def mj_to_gj(x):
    try:
        return x * MJ_TO_GJ
    except Exception:
        return np.nan

def mj_to_m3(x):
    try:
        return x / MJ_PER_NM3
    except Exception:
        return np.nan

# ─────────────────────────────────────────────
# 데이터 불러오기 (캐싱 적용)
# ─────────────────────────────────────────────
@st.cache_data
def load_daily_data():
    # 파일명은 사용하시는 환경에 맞춰 수정해주세요
    excel_path = Path(__file__).parent / "공급량(일일실적).xlsx"
    # 파일이 없을 경우를 대비한 예외처리
    if not excel_path.exists():
        st.error(f"'{excel_path.name}' 파일을 찾을 수 없습니다. 같은 폴더에 넣어주세요.")
        return pd.DataFrame(), pd.DataFrame()

    df_raw = pd.read_excel(excel_path)

    # 필요한 컬럼만 추출
    cols_check = ["일자", "공급량(MJ)", "공급량(M3)", "평균기온(℃)"]
    for c in cols_check:
        if c not in df_raw.columns:
            df_raw[c] = np.nan

    df_raw = df_raw[cols_check].copy()
    df_raw["일자"] = pd.to_datetime(df_raw["일자"])

    df_raw["연도"] = df_raw["일자"].dt.year
    df_raw["월"] = df_raw["일자"].dt.month
    df_raw["일"] = df_raw["일자"].dt.day

    df_temp_all = df_raw.dropna(subset=["평균기온(℃)"]).copy()
    df_model = df_raw.dropna(subset=["공급량(MJ)"]).copy()
    return df_model, df_temp_all

@st.cache_data
def load_monthly_plan() -> pd.DataFrame:
    excel_path = Path(__file__).parent / "공급량(계획_실적).xlsx"
    if not excel_path.exists():
        st.error(f"'{excel_path.name}' 파일을 찾을 수 없습니다.")
        return pd.DataFrame()
        
    df = pd.read_excel(excel_path, sheet_name="월별계획_실적")
    df["연"] = df["연"].astype(int)
    df["월"] = df["월"].astype(int)
    return df

@st.cache_data
def load_effective_calendar() -> pd.DataFrame | None:
    excel_path = Path(__file__).parent / "effective_days_calendar.xlsx"
    if not excel_path.exists():
        return None

    df = pd.read_excel(excel_path)
    if "날짜" not in df.columns:
        return None

    df["일자"] = pd.to_datetime(df["날짜"].astype(str), format="%Y%m%d", errors="coerce")
    
    for col in ["공휴일여부", "명절여부"]:
        if col not in df.columns:
            df[col] = False
            
    df["공휴일여부"] = df["공휴일여부"].fillna(False).astype(bool)
    df["명절여부"] = df["명절여부"].fillna(False).astype(bool)

    return df[["일자", "공휴일여부", "명절여부"]].copy()

# ─────────────────────────────────────────────
# 유틸 함수들
# ─────────────────────────────────────────────
def _find_plan_col(df_plan: pd.DataFrame) -> str:
    candidates = ["계획(사업계획제출_MJ)", "계획(사업계획제출)", "계획_MJ", "계획"]
    for c in candidates:
        if c in df_plan.columns:
            return c
    nums = [c for c in df_plan.columns if pd.api.types.is_numeric_dtype(df_plan[c])]
    return nums[0] if nums else "계획(사업계획제출_MJ)"

def format_table_generic(df, percent_cols=None, temp_cols=None):
    df = df.copy()
    percent_cols = percent_cols or []
    temp_cols = temp_cols or []

    def _fmt_no_comma(x):
        if pd.isna(x): return ""
        try: return f"{int(x)}"
        except: return str(x)

    for col in df.columns:
        if df[col].dtype == bool:
            df[col] = df[col].map(lambda x: "O" if x else "")
            continue
        if col in percent_cols:
            df[col] = df[col].map(lambda x: f"{x:.4f}" if pd.notna(x) else "")
        elif col in temp_cols:
            df[col] = df[col].map(lambda x: f"{x:.2f}" if pd.notna(x) else "")
        elif pd.api.types.is_numeric_dtype(df[col]):
            if col in ["연", "연도", "월", "일", "WeekNum"]:
                df[col] = df[col].map(_fmt_no_comma)
            else:
                df[col] = df[col].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
    return df

def show_table_no_index(df: pd.DataFrame, height: int = 260):
    st.dataframe(df, use_container_width=True, hide_index=True, height=height)

# ─────────────────────────────────────────────
# [핵심] 일별 계획 수립 및 아웃라이어 구간 계산
# ─────────────────────────────────────────────
def make_daily_plan_table(
    df_daily: pd.DataFrame,
    df_plan: pd.DataFrame,
    target_year: int = 2026,
    target_month: int = 1,
    recent_window: int = 3,
):
    cal_df = load_effective_calendar()
    plan_col = _find_plan_col(df_plan)

    # 1. 과거 데이터 조회
    all_years = sorted(df_daily["연도"].unique())
    start_year = target_year - recent_window
    candidate_years = [y for y in range(start_year, target_year) if y in all_years]
    
    if not candidate_years:
        return None, [], pd.DataFrame()

    df_pool = df_daily[(df_daily["연도"].isin(candidate_years)) & (df_daily["월"] == target_month)].copy()
    df_pool = df_pool.dropna(subset=["공급량(MJ)"])
    used_years = sorted(df_pool["연도"].unique().tolist())
    
    if not used_years:
        return None, [], pd.DataFrame()

    df_recent = df_daily[(df_daily["연도"].isin(used_years)) & (df_daily["월"] == target_month)].copy()
    df_recent = df_recent.dropna(subset=["공급량(MJ)"])
    df_recent = df_recent.sort_values(["연도", "일"]).copy()
    df_recent["weekday_idx"] = df_recent["일자"].dt.weekday  # 0=월, 6=일

    # 휴일 정보 병합
    if cal_df is not None:
        df_recent = df_recent.merge(cal_df, on="일자", how="left")
        for c in ["공휴일여부", "명절여부"]:
            if c not in df_recent.columns: df_recent[c] = False
            df_recent[c] = df_recent[c].fillna(False).astype(bool)
    else:
        df_recent["공휴일여부"] = False
        df_recent["명절여부"] = False

    df_recent["is_holiday"] = df_recent["공휴일여부"] | df_recent["명절여부"]
    df_recent["is_weekend"] = (df_recent["weekday_idx"] >= 5) | df_recent["is_holiday"]
    df_recent["is_weekday1"] = (~df_recent["is_weekend"]) & (df_recent["weekday_idx"].isin([0, 4])) # 월,금
    df_recent["is_weekday2"] = (~df_recent["is_weekend"]) & (df_recent["weekday_idx"].isin([1, 2, 3])) # 화수목

    # 월별 총량 대비 비율 계산
    df_recent["month_total"] = df_recent.groupby("연도")["공급량(MJ)"].transform("sum")
    df_recent["ratio"] = df_recent["공급량(MJ)"] / df_recent["month_total"]

    # n번째 요일 계산
    df_recent["nth_dow"] = df_recent.sort_values(["연도", "일"]).groupby(["연도", "weekday_idx"]).cumcount() + 1

    # 그룹별 비율 평균 (Lookup Dictionary 생성)
    # 1) 주말/공휴일
    mask_wend = df_recent["is_weekend"]
    ratio_wend_grp = df_recent[mask_wend].groupby(["weekday_idx", "nth_dow"])["ratio"].mean().to_dict()
    ratio_wend_dow = df_recent[mask_wend].groupby("weekday_idx")["ratio"].mean().to_dict()
    
    # 2) 평일1 (월/금)
    mask_w1 = df_recent["is_weekday1"]
    ratio_w1_grp = df_recent[mask_w1].groupby(["weekday_idx", "nth_dow"])["ratio"].mean().to_dict()
    ratio_w1_dow = df_recent[mask_w1].groupby("weekday_idx")["ratio"].mean().to_dict()
    
    # 3) 평일2 (화수목)
    mask_w2 = df_recent["is_weekday2"]
    ratio_w2_grp = df_recent[mask_w2].groupby(["weekday_idx", "nth_dow"])["ratio"].mean().to_dict()
    ratio_w2_dow = df_recent[mask_w2].groupby("weekday_idx")["ratio"].mean().to_dict()

    # 2. 타겟 월 날짜 생성
    last_day = calendar.monthrange(target_year, target_month)[1]
    date_range = pd.date_range(f"{target_year}-{target_month:02d}-01", periods=last_day, freq="D")
    
    df_target = pd.DataFrame({"일자": date_range})
    df_target["연"] = target_year
    df_target["월"] = target_month
    df_target["일"] = df_target["일자"].dt.day
    df_target["weekday_idx"] = df_target["일자"].dt.weekday
    
    # 타겟 월 휴일 정보
    if cal_df is not None:
        df_target = df_target.merge(cal_df, on="일자", how="left")
        for c in ["공휴일여부", "명절여부"]:
            if c not in df_target.columns: df_target[c] = False
            df_target[c] = df_target[c].fillna(False).astype(bool)
    else:
        df_target["공휴일여부"] = False
        df_target["명절여부"] = False
        
    df_target["is_holiday"] = df_target["공휴일여부"] | df_target["명절여부"]
    df_target["is_weekend"] = (df_target["weekday_idx"] >= 5) | df_target["is_holiday"]
    df_target["is_weekday1"] = (~df_target["is_weekend"]) & (df_target["weekday_idx"].isin([0, 4]))
    df_target["is_weekday2"] = (~df_target["is_weekend"]) & (df_target["weekday_idx"].isin([1, 2, 3]))
    
    weekday_names = ["월", "화", "수", "목", "금", "토", "일"]
    df_target["요일"] = df_target["weekday_idx"].map(lambda i: weekday_names[i])
    df_target["nth_dow"] = df_target.sort_values("일").groupby("weekday_idx").cumcount() + 1
    
    def _get_label(r):
        if r["is_weekend"]: return "주말/공휴일"
        if r["is_weekday1"]: return "평일1(월·금)"
        return "평일2(화·수·목)"
    df_target["구분"] = df_target.apply(_get_label, axis=1)

    # 비율 매핑
    def _pick(r):
        dow, nth = int(r["weekday_idx"]), int(r["nth_dow"])
        key = (dow, nth)
        if r["is_weekend"]:
            return ratio_wend_grp.get(key, ratio_wend_dow.get(dow, np.nan))
        if r["is_weekday1"]:
            return ratio_w1_grp.get(key, ratio_w1_dow.get(dow, np.nan))
        return ratio_w2_grp.get(key, ratio_w2_dow.get(dow, np.nan))
        
    df_target["raw"] = df_target.apply(_pick, axis=1).astype(float)
    
    # 결측치 보정 (전체 평균)
    overall_mean = df_target["raw"].mean()
    df_target["raw"] = df_target["raw"].fillna(overall_mean if pd.notna(overall_mean) else 1.0)
    
    # 정규화 (합계=1)
    raw_sum = df_target["raw"].sum()
    df_target["일별비율"] = df_target["raw"] / raw_sum if raw_sum > 0 else 1.0/last_day

    # 계획 총량 적용
    row_plan = df_plan[(df_plan["연"] == target_year) & (df_plan["월"] == target_month)]
    plan_total = float(row_plan[plan_col].iloc[0]) if not row_plan.empty else np.nan
    df_target["예상공급량(MJ)"] = (df_target["일별비율"] * plan_total).round(0)

    # ─────────────────────────────────────────────────────────────
    # [NEW] Han형님 요청: 주차별 + (주중/주말) 분리 이동평균 및 아웃라이어 감지
    # ─────────────────────────────────────────────────────────────
    # 1. 주차(WeekNum) 생성 (ISO 기준)
    df_target["WeekNum"] = df_target["일자"].dt.isocalendar().week
    
    # 2. 그룹핑: [주차] + [주말여부]
    #    (주말/공휴일은 is_weekend=True 그룹, 나머지는 False 그룹)
    #    이렇게 하면 한 주 내에서 평일 평균, 주말 평균이 따로 계산됩니다.
    group_cols = ["WeekNum", "is_weekend"]
    
    df_target["Group_Mean"] = df_target.groupby(group_cols)["예상공급량(MJ)"].transform("mean")
    
    # 3. 상한/하한 (±10%)
    df_target["Bound_Upper"] = df_target["Group_Mean"] * 1.10
    df_target["Bound_Lower"] = df_target["Group_Mean"] * 0.90
    
    # 4. 아웃라이어 여부 (범위 밖이면 True)
    df_target["is_outlier"] = (df_target["예상공급량(MJ)"] > df_target["Bound_Upper"]) | \
                              (df_target["예상공급량(MJ)"] < df_target["Bound_Lower"])
                              
    return df_target, used_years, plan_total

# ─────────────────────────────────────────────
# 메인 분석 탭
# ─────────────────────────────────────────────
def tab_analysis(df_daily: pd.DataFrame):
    st.header("📅 도시가스 일별 계획 (Outlier Check Ver.)")
    st.caption("Han형님, 요청하신 **주말/주중 분리 상한선**과 **아웃라이어 표시** 기능을 적용했습니다.")

    df_plan = load_monthly_plan()
    
    # 사이드바 컨트롤
    with st.sidebar:
        st.subheader("🛠️ 설정 패널")
        target_year = st.number_input("계획 연도", 2020, 2030, 2026)
        target_month = st.selectbox("계획 월", list(range(1, 13)), index=0)
        recent_window = st.slider("최근 N년 패턴 참조", 1, 10, 3)

    # 계산 실행
    with st.spinner("패턴 분석 및 계획 수립 중..."):
        df_res, used_years, plan_total_mj = make_daily_plan_table(
            df_daily, df_plan, target_year, target_month, recent_window
        )

    if df_res is None:
        st.warning("데이터가 부족하여 계획을 수립할 수 없습니다.")
        return

    # 요약 정보
    st.markdown(f"### 📌 {target_year}년 {target_month}월 분석 결과")
    st.info(f"참조한 과거 연도: **{used_years}** (총 {len(used_years)}개년)")
    
    # 데이터 변환 (MJ -> GJ)
    df_disp = df_res.copy()
    df_disp["예상공급량(GJ)"] = df_disp["예상공급량(MJ)"].apply(mj_to_gj)
    df_disp["상한선(GJ)"] = df_disp["Bound_Upper"].apply(mj_to_gj)
    df_disp["하한선(GJ)"] = df_disp["Bound_Lower"].apply(mj_to_gj)
    df_disp["그룹평균(GJ)"] = df_disp["Group_Mean"].apply(mj_to_gj)

    # ─────────────────────────────────────────────
    # [시각화] Plotly 그래프
    # ─────────────────────────────────────────────
    fig = go.Figure()

    # 1. Bar Chart: 일별 계획
    # 평일/주말 색상 구분
    colors = np.where(df_disp["is_weekend"], "#00CC96", "#636EFA") # 주말: 초록, 평일: 파랑
    
    fig.add_trace(go.Bar(
        x=df_disp["일"], 
        y=df_disp["예상공급량(GJ)"],
        marker_color=colors,
        name="일별 계획(GJ)",
        opacity=0.8
    ))

    # 2. Band Chart: 상한/하한 영역
    # 끊어지는 선을 연결되게 보이려면 x축이 연속적이어야 하는데, 
    # 여기서는 '주중'과 '주말'의 레벨 차이가 급격하므로 Step 형태가 자연스러움.
    
    # 상한선 (투명 선)
    fig.add_trace(go.Scatter(
        x=df_disp["일"], y=df_disp["상한선(GJ)"],
        mode='lines', line=dict(width=0), showlegend=False, hoverinfo='skip'
    ))
    # 하한선 (채우기)
    fig.add_trace(go.Scatter(
        x=df_disp["일"], y=df_disp["하한선(GJ)"],
        mode='lines', line=dict(width=0), 
        fill='tonexty', fillcolor='rgba(100, 100, 100, 0.15)',
        name='허용범위(±10%)', hoverinfo='skip'
    ))
    
    # 그룹 평균선 (점선)
    fig.add_trace(go.Scatter(
        x=df_disp["일"], y=df_disp["그룹평균(GJ)"],
        mode='lines', line=dict(color='gray', dash='dot', width=1),
        name='주간 그룹평균'
    ))

    # 3. Outlier 마커 (빨간 점)
    outliers = df_disp[df_disp["is_outlier"]]
    if not outliers.empty:
        fig.add_trace(go.Scatter(
            x=outliers["일"], y=outliers["예상공급량(GJ)"],
            mode='markers',
            marker=dict(color='red', size=10, symbol='x'),
            name='Outlier (범위 초과)'
        ))

    fig.update_layout(
        title=f"{target_year}년 {target_month}월 일별 계획 및 Outlier 감지",
        xaxis_title="일 (Day)",
        yaxis_title="공급량 (GJ)",
        legend=dict(orientation="h", y=1.1),
        margin=dict(l=20, r=20, t=80, b=40),
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # ─────────────────────────────────────────────
    # 상세 테이블
    # ─────────────────────────────────────────────
    st.markdown("#### 📋 상세 데이터 (Outlier 강조)")
    
    # 테이블용 데이터 정리
    cols_table = [
        "일자", "요일", "WeekNum", "구분", 
        "예상공급량(GJ)", "상한선(GJ)", "하한선(GJ)", "is_outlier"
    ]
    df_table = df_disp[cols_table].copy()
    
    # 아웃라이어 행 강조 스타일링
    def highlight_outlier(row):
        if row["is_outlier"]:
            return ['background-color: #FFEBEB'] * len(row)
        return [''] * len(row)

    # 포맷팅
    df_table["일자"] = df_table["일자"].dt.strftime("%Y-%m-%d")
    for c in ["예상공급량(GJ)", "상한선(GJ)", "하한선(GJ)"]:
        df_table[c] = df_table[c].apply(lambda x: f"{x:,.0f}")
    
    df_table["is_outlier"] = df_table["is_outlier"].map({True: "🚨초과", False: ""})

    st.dataframe(
        df_table.style.apply(highlight_outlier, axis=1),
        use_container_width=True,
        height=400
    )

# ─────────────────────────────────────────────
# App Entry Point
# ─────────────────────────────────────────────
def main():
    df_model, _ = load_daily_data()
    if df_model.empty:
        st.error("데이터 로딩 실패")
        return
        
    tab_analysis(df_model)

if __name__ == "__main__":
    main()
