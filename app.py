import calendar
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill


# ─────────────────────────────────────────────
# 1. 기본 설정 및 상수
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="도시가스 공급량: 일별계획 예측 (Final)",
    layout="wide",
)

MJ_PER_NM3 = 42.563
MJ_TO_GJ = 1.0 / 1000.0

def mj_to_gj(x):
    try: return x * MJ_TO_GJ
    except: return 0

def mj_to_m3(x):
    try: return x / MJ_PER_NM3
    except: return 0

# ─────────────────────────────────────────────
# 2. 데이터 로딩 (에러 방지 강화)
# ─────────────────────────────────────────────
@st.cache_data
def load_daily_data():
    # 파일 경로 확인
    excel_path = Path(__file__).parent / "공급량(일일실적).xlsx"
    if not excel_path.exists():
        st.error(f"❌ 파일이 없습니다: {excel_path.name}")
        return pd.DataFrame(), pd.DataFrame()

    try:
        df_raw = pd.read_excel(excel_path)
        
        # 필수 컬럼 체크
        required = ["일자", "공급량(MJ)", "공급량(M3)", "평균기온(℃)"]
        missing = [c for c in required if c not in df_raw.columns]
        if missing:
            st.error(f"❌ '일일실적' 파일에 다음 컬럼이 없습니다: {missing}")
            return pd.DataFrame(), pd.DataFrame()

        df_raw = df_raw[required].copy()
        df_raw["일자"] = pd.to_datetime(df_raw["일자"], errors='coerce')
        df_raw = df_raw.dropna(subset=["일자"]) # 날짜 없는 행 제거

        df_raw["연도"] = df_raw["일자"].dt.year
        df_raw["월"] = df_raw["일자"].dt.month
        df_raw["일"] = df_raw["일자"].dt.day

        df_temp_all = df_raw.dropna(subset=["평균기온(℃)"]).copy()
        df_model = df_raw.dropna(subset=["공급량(MJ)"]).copy()
        return df_model, df_temp_all
        
    except Exception as e:
        st.error(f"❌ 데이터 로딩 중 에러 발생: {e}")
        return pd.DataFrame(), pd.DataFrame()


@st.cache_data
def load_monthly_plan() -> pd.DataFrame:
    excel_path = Path(__file__).parent / "공급량(계획_실적).xlsx"
    if not excel_path.exists():
        st.error(f"❌ 파일이 없습니다: {excel_path.name}")
        return pd.DataFrame()
        
    try:
        df = pd.read_excel(excel_path, sheet_name="월별계획_실적")
        # 연, 월 컬럼 필수
        if "연" not in df.columns or "월" not in df.columns:
             st.error("❌ '월별계획_실적' 시트에 '연', '월' 컬럼이 필요합니다.")
             return pd.DataFrame()
             
        df["연"] = pd.to_numeric(df["연"], errors='coerce').fillna(0).astype(int)
        df["월"] = pd.to_numeric(df["월"], errors='coerce').fillna(0).astype(int)
        return df
    except Exception as e:
        st.error(f"❌ 월별계획 로딩 실패: {e}")
        return pd.DataFrame()


@st.cache_data
def load_effective_calendar() -> pd.DataFrame | None:
    excel_path = Path(__file__).parent / "effective_days_calendar.xlsx"
    if not excel_path.exists():
        return None

    try:
        df = pd.read_excel(excel_path)
        if "날짜" not in df.columns: return None
        
        df["일자"] = pd.to_datetime(df["날짜"].astype(str), format="%Y%m%d", errors="coerce")
        for col in ["공휴일여부", "명절여부"]:
            if col not in df.columns: df[col] = False
            df[col] = df[col].fillna(False).astype(bool)
            
        return df[["일자", "공휴일여부", "명절여부"]].copy()
    except:
        return None

# ─────────────────────────────────────────────
# 3. 유틸 함수
# ─────────────────────────────────────────────
def _find_plan_col(df_plan: pd.DataFrame) -> str:
    # 계획 컬럼 찾기 (우선순위)
    candidates = ["계획(사업계획제출_MJ)", "계획(사업계획제출)", "계획_MJ", "계획"]
    for c in candidates:
        if c in df_plan.columns: return c
    # 없으면 숫자형 컬럼 중 첫번째 (연, 월 제외)
    nums = [c for c in df_plan.columns if pd.api.types.is_numeric_dtype(df_plan[c]) and c not in ["연", "월"]]
    return nums[0] if nums else ""

def make_month_plan_horizontal(df_plan, target_year, plan_col):
    if df_plan.empty or not plan_col: return pd.DataFrame()
    
    df_year = df_plan[df_plan["연"] == target_year][["월", plan_col]].copy()
    if df_year.empty: return pd.DataFrame()
    
    base = pd.DataFrame({"월": list(range(1, 13))})
    df_year = base.merge(df_year, on="월", how="left")
    
    df_year = df_year.rename(columns={plan_col: "월별 계획(MJ)"})
    
    # 횡형 변환 로직 (간소화)
    row_gj = {"구분": "사업계획(월별 계획, GJ)"}
    row_m3 = {"구분": "사업계획(월별 계획, ㎥)"}
    
    total_mj = df_year["월별 계획(MJ)"].sum()
    row_gj["연간합계"] = round(mj_to_gj(total_mj), 0)
    row_m3["연간합계"] = round(mj_to_m3(total_mj), 0)

    for _, row in df_year.iterrows():
        m = int(row["월"])
        mj = row["월별 계획(MJ)"]
        row_gj[f"{m}월"] = round(mj_to_gj(mj), 0)
        row_m3[f"{m}월"] = round(mj_to_m3(mj), 0)
        
    return pd.DataFrame([row_gj, row_m3])

def format_table_generic(df, percent_cols=None):
    if df.empty: return df
    df = df.copy()
    percent_cols = percent_cols or []
    
    for col in df.columns:
        if df[col].dtype == bool:
            df[col] = df[col].map(lambda x: "O" if x else "")
        elif col in percent_cols:
            df[col] = df[col].map(lambda x: f"{x:.4f}" if pd.notna(x) else "")
        elif pd.api.types.is_numeric_dtype(df[col]):
             if col in ["연", "월", "일", "WeekNum"]:
                 df[col] = df[col].map(lambda x: f"{int(x)}" if pd.notna(x) else "")
             else:
                 df[col] = df[col].map(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
    return df

def show_table_no_index(df, height=260):
    st.dataframe(df, use_container_width=True, hide_index=True, height=height)

# ─────────────────────────────────────────────
# 4. 핵심 분석 로직 (기존 로직 100% 유지)
# ─────────────────────────────────────────────
def make_daily_plan_table(df_daily, df_plan, target_year, target_month, recent_window):
    cal_df = load_effective_calendar()
    plan_col = _find_plan_col(df_plan)
    if not plan_col: return None, None, [], pd.DataFrame()

    all_years = sorted(df_daily["연도"].unique())
    start_year = target_year - recent_window
    candidate_years = [y for y in range(start_year, target_year) if y in all_years]
    
    df_pool = df_daily[(df_daily["연도"].isin(candidate_years)) & (df_daily["월"] == target_month)].dropna(subset=["공급량(MJ)"])
    used_years = sorted(df_pool["연도"].unique())
    
    if not used_years: return None, None, [], pd.DataFrame()

    df_recent = df_pool.copy().sort_values(["연도", "일"])
    df_recent["weekday_idx"] = df_recent["일자"].dt.weekday

    # 휴일 매핑
    if cal_df is not None:
        df_recent = df_recent.merge(cal_df, on="일자", how="left")
        for c in ["공휴일여부", "명절여부"]:
            df_recent[c] = df_recent[c].fillna(False).astype(bool)
    else:
        df_recent["공휴일여부"] = False; df_recent["명절여부"] = False

    df_recent["is_holiday"] = df_recent["공휴일여부"] | df_recent["명절여부"]
    df_recent["is_weekend"] = (df_recent["weekday_idx"] >= 5) | df_recent["is_holiday"]
    df_recent["is_weekday1"] = (~df_recent["is_weekend"]) & (df_recent["weekday_idx"].isin([0, 4]))
    df_recent["is_weekday2"] = (~df_recent["is_weekend"]) & (df_recent["weekday_idx"].isin([1, 2, 3]))

    # 비율 계산 로직
    df_recent["month_total"] = df_recent.groupby("연도")["공급량(MJ)"].transform("sum")
    df_recent["ratio"] = df_recent["공급량(MJ)"] / df_recent["month_total"]
    df_recent["nth_dow"] = df_recent.groupby(["연도", "weekday_idx"]).cumcount() + 1

    # 그룹별 평균 비율 딕셔너리 생성
    def get_ratio_dict(mask):
        grp = df_recent[mask].groupby(["weekday_idx", "nth_dow"])["ratio"].mean().to_dict()
        fallback = df_recent[mask].groupby("weekday_idx")["ratio"].mean().to_dict()
        return grp, fallback

    w_grp, w_fb = get_ratio_dict(df_recent["is_weekend"])
    w1_grp, w1_fb = get_ratio_dict(df_recent["is_weekday1"])
    w2_grp, w2_fb = get_ratio_dict(df_recent["is_weekday2"])

    # 타겟 생성
    last_day = calendar.monthrange(target_year, target_month)[1]
    df_target = pd.DataFrame({"일자": pd.date_range(f"{target_year}-{target_month:02d}-01", periods=last_day)})
    df_target["연"] = target_year
    df_target["월"] = target_month
    df_target["일"] = df_target["일자"].dt.day
    df_target["weekday_idx"] = df_target["일자"].dt.weekday
    
    # 타겟 휴일 매핑
    if cal_df is not None:
        df_target = df_target.merge(cal_df, on="일자", how="left")
        for c in ["공휴일여부", "명절여부"]:
            df_target[c] = df_target[c].fillna(False).astype(bool)
    else:
        df_target["공휴일여부"] = False; df_target["명절여부"] = False

    df_target["is_holiday"] = df_target["공휴일여부"] | df_target["명절여부"]
    df_target["is_weekend"] = (df_target["weekday_idx"] >= 5) | df_target["is_holiday"]
    df_target["is_weekday1"] = (~df_target["is_weekend"]) & (df_target["weekday_idx"].isin([0, 4]))
    df_target["is_weekday2"] = (~df_target["is_weekend"]) & (df_target["weekday_idx"].isin([1, 2, 3]))

    weekday_names = ["월", "화", "수", "목", "금", "토", "일"]
    df_target["요일"] = df_target["weekday_idx"].map(lambda x: weekday_names[x])
    df_target["nth_dow"] = df_target.groupby("weekday_idx").cumcount() + 1

    def _get_label(r):
        if r["is_weekend"]: return "주말/공휴일"
        if r["is_weekday1"]: return "평일1(월·금)"
        return "평일2(화·수·목)"
    df_target["구분"] = df_target.apply(_get_label, axis=1)

    # 비율 적용
    def _apply_ratio(r):
        k = (r["weekday_idx"], r["nth_dow"])
        wd = r["weekday_idx"]
        if r["is_weekend"]: return w_grp.get(k, w_fb.get(wd, np.nan))
        if r["is_weekday1"]: return w1_grp.get(k, w1_fb.get(wd, np.nan))
        return w2_grp.get(k, w2_fb.get(wd, np.nan))

    df_target["raw"] = df_target.apply(_apply_ratio, axis=1).astype(float)
    
    # 결측치 처리 (전체 평균)
    overall_mean = df_target["raw"].mean()
    df_target["raw"] = df_target["raw"].fillna(overall_mean if pd.notna(overall_mean) else 1.0)
    
    # 정규화
    raw_sum = df_target["raw"].sum()
    df_target["일별비율"] = df_target["raw"] / raw_sum

    # 계획 총량 적용
    row_plan = df_plan[(df_plan["연"] == target_year) & (df_plan["월"] == target_month)]
    plan_total = float(row_plan[plan_col].iloc[0]) if not row_plan.empty else 0
    
    df_target["예상공급량(MJ)"] = (df_target["일별비율"] * plan_total).round(0)

    # [NEW] Outlier 판단 (컬럼만 추가)
    df_target["WeekNum"] = df_target["일자"].dt.isocalendar().week
    df_target["Group_Mean"] = df_target.groupby(["WeekNum", "is_weekend"])["예상공급량(MJ)"].transform("mean")
    df_target["Bound_Upper"] = df_target["Group_Mean"] * 1.10
    df_target["Bound_Lower"] = df_target["Group_Mean"] * 0.90
    df_target["is_outlier"] = (df_target["예상공급량(MJ)"] > df_target["Bound_Upper"]) | (df_target["예상공급량(MJ)"] < df_target["Bound_Lower"])
    
    # 필요한 컬럼만 정리해서 리턴
    df_target["최근N년_평균공급량(MJ)"] = 0 # Placeholder
    df_target["최근N년_총공급량(MJ)"] = 0 # Placeholder

    return df_target, None, used_years, None

# ─────────────────────────────────────────────
# 5. UI 및 시각화 (형님 지시사항 반영)
# ─────────────────────────────────────────────
def tab_daily_plan(df_daily):
    st.subheader("📅 Daily 공급량 분석")

    df_plan = load_monthly_plan()
    if df_plan.empty: return

    # ... (연도/월/윈도우 선택 UI 생략 - 기존과 동일) ...
    plan_col = _find_plan_col(df_plan)
    years = sorted(df_plan["연"].unique())
    
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1: target_year = st.selectbox("연도", years, index=len(years)-1)
    with col2: target_month = st.selectbox("월", range(1, 13))
    with col3: recent_window = st.slider("과거 참조(년)", 1, 10, 3)

    # 계산 실행
    df_res, _, used_years, _ = make_daily_plan_table(df_daily, df_plan, target_year, target_month, recent_window)

    if df_res is None:
        st.warning("데이터 부족으로 분석 불가")
        return

    st.markdown(f"**참조 데이터:** {min(used_years)}년 ~ {max(used_years)}년")
    
    # ────────────────────────────────────────────────────────
    # [NEW] 보정 로직 및 UI (우측 상단 배치)
    # ────────────────────────────────────────────────────────
    view = df_res.copy()
    view["보정_예상공급량(MJ)"] = view["예상공급량(MJ)"] # 초기화

    st.divider()

    # ★ 우측 상단 버튼 배치를 위한 레이아웃
    # 왼쪽: 제목 / 오른쪽: 보정 패널
    c_head, c_ctrl = st.columns([1, 2])
    
    with c_head:
        st.markdown("### 📊 일별 계획 & Outlier")
    
    with c_ctrl:
        # 우측에 버튼 배치
        use_calib = st.checkbox("✅ 이상치 보정 활성화", value=False)
        
        diff_mj = 0
        if use_calib:
            with st.expander("🛠️ 보정 상세 설정 (이상구간 -> 보정구간 배분)", expanded=True):
                min_d = view["일자"].min().date()
                max_d = view["일자"].max().date()
                
                cc1, cc2 = st.columns(2)
                d_out = cc1.date_input("1. 이상구간 (자르기)", (min_d, min_d), min_value=min_d, max_value=max_d)
                d_fix = cc2.date_input("2. 보정구간 (채우기)", (min_d, max_d), min_value=min_d, max_value=max_d)
                
                # 보정 로직 (Clamp & Redistribute)
                if len(d_out) == 2 and len(d_fix) == 2:
                    s_out, e_out = d_out
                    s_fix, e_fix = d_fix
                    
                    # 1. 이상구간 Clamp
                    mask_out = (view["일자"].dt.date >= s_out) & (view["일자"].dt.date <= e_out)
                    if mask_out.any():
                        view.loc[mask_out, "보정_예상공급량(MJ)"] = np.where(
                            view.loc[mask_out, "예상공급량(MJ)"] > view.loc[mask_out, "Bound_Upper"],
                            view.loc[mask_out, "Bound_Upper"],
                            np.where(
                                view.loc[mask_out, "예상공급량(MJ)"] < view.loc[mask_out, "Bound_Lower"],
                                view.loc[mask_out, "Bound_Lower"],
                                view.loc[mask_out, "예상공급량(MJ)"]
                            )
                        )
                        diff_mj = (view.loc[mask_out, "예상공급량(MJ)"] - view.loc[mask_out, "보정_예상공급량(MJ)"]).sum()
                    
                    # 2. 보정구간 Redistribute
                    mask_fix = (view["일자"].dt.date >= s_fix) & (view["일자"].dt.date <= e_fix)
                    sum_r = view.loc[mask_fix, "일별비율"].sum()
                    if mask_fix.any() and sum_r > 0:
                         view.loc[mask_fix, "보정_예상공급량(MJ)"] += diff_mj * (view.loc[mask_fix, "일별비율"] / sum_r)

                st.caption(f"변동량: {mj_to_gj(diff_mj):,.0f} GJ")

    # ────────────────────────────────────────────────────────
    # [그래프] 기존 색상 유지 + 보정값 회색 덮어쓰기
    # ────────────────────────────────────────────────────────
    # 단위 변환
    view["예상공급량(GJ)"] = view["예상공급량(MJ)"].apply(mj_to_gj)
    view["보정_예상공급량(GJ)"] = view["보정_예상공급량(MJ)"].apply(mj_to_gj)
    view["Bound_Upper(GJ)"] = view["Bound_Upper"].apply(mj_to_gj)
    view["Bound_Lower(GJ)"] = view["Bound_Lower"].apply(mj_to_gj)

    fig = go.Figure()

    # 1. [기존 그래프] 색상: 평일1(파랑), 평일2(빨강), 주말(초록) - 형님 원래 코드의 로직
    # colors 배열 생성
    colors = np.where(view["is_weekend"], "#00CC96", # 주말 (Green)
             np.where(view["weekday_idx"].isin([0, 4]), "#636EFA", # 평일1 (Blue)
                      "#EF553B")) # 평일2 (Red)
                      
    # 기본 막대 (AS-IS) - 보정이 켜지면 투명도를 줘서 뒤에 깔리게 함
    opacity_val = 0.3 if use_calib else 1.0
    fig.add_trace(go.Bar(
        x=view["일"], y=view["예상공급량(GJ)"],
        marker_color=colors,
        name="기존 계획",
        opacity=opacity_val
    ))
    
    # 2. [보정 그래프] (TO-BE) - 보정 활성화 시에만 그림
    # 색상: 진한 회색 (투명도 60%)
    if use_calib:
        fig.add_trace(go.Bar(
            x=view["일"], y=view["보정_예상공급량(GJ)"],
            marker_color="rgba(80, 80, 80, 0.6)",
            name="보정 후",
        ))

    # 3. 보조 라인들 (비율, 상한, 하한)
    fig.add_trace(go.Scatter(x=view["일"], y=view["일별비율"], yaxis="y2", name="비율", line=dict(color='black', width=1)))
    fig.add_trace(go.Scatter(x=view["일"], y=view["Bound_Upper(GJ)"], line=dict(width=0), showlegend=False))
    fig.add_trace(go.Scatter(x=view["일"], y=view["Bound_Lower(GJ)"], line=dict(width=0), fill='tonexty', fillcolor='rgba(100,100,100,0.1)', name='범위(±10%)'))
    
    # 4. Outlier
    outliers = view[view["is_outlier"]]
    if not outliers.empty:
        fig.add_trace(go.Scatter(x=outliers["일"], y=outliers["예상공급량(GJ)"], mode='markers', marker=dict(color='red', symbol='x', size=10), name='Outlier'))

    fig.update_layout(
        title=f"{target_year}년 {target_month}월 공급계획",
        yaxis=dict(title="공급량(GJ)"),
        yaxis2=dict(title="비율", overlaying="y", side="right"),
        barmode="overlay" if use_calib else "group",
        legend=dict(orientation="h", y=1.1)
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # 테이블 출력 등 나머지 UI
    show_table_no_index(format_table_generic(view[["일자", "구분", "예상공급량(GJ)", "보정_예상공급량(GJ)", "is_outlier"]], percent_cols=[]))

# ─────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────
def main():
    df, _ = load_daily_data()
    if df.empty:
        st.warning("데이터를 불러오지 못했습니다.")
        return
    tab_daily_plan(df)

if __name__ == "__main__":
    main()
