import calendar
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill


# ─────────────────────────────────────────────
# 1. 단위/환산 상수
# ─────────────────────────────────────────────
MJ_PER_NM3 = 42.563
MJ_TO_GJ = 1.0 / 1000.0

def mj_to_gj(x):
    try: return x * MJ_TO_GJ
    except Exception: return np.nan

def mj_to_m3(x):
    try: return x / MJ_PER_NM3
    except Exception: return np.nan


# ─────────────────────────────────────────────
# 2. 기본 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="도시가스 공급량: 일별계획 예측 (Final Ver)",
    layout="wide",
)


# ─────────────────────────────────────────────
# 3. 데이터 불러오기
# ─────────────────────────────────────────────
@st.cache_data
def load_daily_data():
    excel_path = Path(__file__).parent / "공급량(일일실적).xlsx"
    if not excel_path.exists():
        return pd.DataFrame(), pd.DataFrame()

    try:
        df_raw = pd.read_excel(excel_path)
        required = ["일자", "공급량(MJ)", "공급량(M3)", "평균기온(℃)"]
        for c in required:
            if c not in df_raw.columns: df_raw[c] = np.nan

        df_raw = df_raw[required].copy()
        df_raw["일자"] = pd.to_datetime(df_raw["일자"])
        df_raw["연도"] = df_raw["일자"].dt.year
        df_raw["월"] = df_raw["일자"].dt.month
        df_raw["일"] = df_raw["일자"].dt.day

        df_temp_all = df_raw.dropna(subset=["평균기온(℃)"]).copy()
        df_model = df_raw.dropna(subset=["공급량(MJ)"]).copy()
        return df_model, df_temp_all
    except:
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data
def load_monthly_plan() -> pd.DataFrame:
    excel_path = Path(__file__).parent / "공급량(계획_실적).xlsx"
    if not excel_path.exists(): return pd.DataFrame()
    try:
        df = pd.read_excel(excel_path, sheet_name="월별계획_실적")
        if "연" in df.columns: df["연"] = pd.to_numeric(df["연"], errors='coerce').fillna(0).astype(int)
        if "월" in df.columns: df["월"] = pd.to_numeric(df["월"], errors='coerce').fillna(0).astype(int)
        return df
    except: return pd.DataFrame()

@st.cache_data
def load_effective_calendar() -> pd.DataFrame | None:
    excel_path = Path(__file__).parent / "effective_days_calendar.xlsx"
    if not excel_path.exists(): return None
    try:
        df = pd.read_excel(excel_path)
        if "날짜" not in df.columns: return None
        df["일자"] = pd.to_datetime(df["날짜"].astype(str), format="%Y%m%d", errors="coerce")
        for col in ["공휴일여부", "명절여부"]:
            if col not in df.columns: df[col] = False
            df[col] = df[col].fillna(False).astype(bool)
        return df[["일자", "공휴일여부", "명절여부"]].copy()
    except: return None


# ─────────────────────────────────────────────
# 4. 유틸 함수들
# ─────────────────────────────────────────────
def _find_plan_col(df_plan: pd.DataFrame) -> str:
    candidates = ["계획(사업계획제출_MJ)", "계획(사업계획제출)", "계획_MJ", "계획"]
    for c in candidates:
        if c in df_plan.columns: return c
    nums = [c for c in df_plan.columns if pd.api.types.is_numeric_dtype(df_plan[c]) and c not in ["연", "월"]]
    return nums[0] if nums else "계획(사업계획제출_MJ)"

def make_month_plan_horizontal(df_plan: pd.DataFrame, target_year: int, plan_col: str) -> pd.DataFrame:
    if df_plan.empty or not plan_col: return pd.DataFrame()
    df_year = df_plan[df_plan["연"] == target_year][["월", plan_col]].copy()
    if df_year.empty: return pd.DataFrame()
    base = pd.DataFrame({"월": list(range(1, 13))})
    df_year = base.merge(df_year, on="월", how="left")
    df_year = df_year.rename(columns={plan_col: "월별 계획(MJ)"})
    
    total_mj = df_year["월별 계획(MJ)"].sum()
    df_year["월별 계획(GJ)"] = (df_year["월별 계획(MJ)"].apply(mj_to_gj)).round(0)
    df_year["월별 계획(㎥)"] = (df_year["월별 계획(MJ)"].apply(mj_to_m3)).round(0)
    
    total_gj = mj_to_gj(total_mj)
    total_m3 = mj_to_m3(total_mj)
    
    row_gj = {"구분": "사업계획(월별 계획, GJ)"}
    row_m3 = {"구분": "사업계획(월별 계획, ㎥)"}
    
    for _, row in df_year.iterrows():
        m = int(row["월"])
        mj = row["월별 계획(MJ)"]
        row_gj[f"{m}월"] = round(mj_to_gj(mj), 0)
        row_m3[f"{m}월"] = round(mj_to_m3(mj), 0)
    
    row_gj["연간합계"] = round(total_gj, 0)
    row_m3["연간합계"] = round(total_m3, 0)
    
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

def show_table_no_index(df: pd.DataFrame, height: int = 260):
    try: st.dataframe(df, use_container_width=True, hide_index=True, height=height)
    except: st.table(df)

def _format_excel_sheet(ws, freeze="A2", center=True):
    if freeze: ws.freeze_panes = freeze
    if center:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for c in row: c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _add_cumulative_status_sheet(wb, annual_year: int):
    sheet_name = "누적계획현황"
    if sheet_name in wb.sheetnames: return
    ws = wb.create_sheet(sheet_name)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    ws["A1"] = "기준일"; ws["B1"] = pd.Timestamp(f"{annual_year}-01-01")
    ws["B1"].number_format = "yyyy-mm-dd"
    
    headers = ["구분", "목표(GJ)", "누적(GJ)", "목표(m³)", "누적(m³)", "진행률(GJ)"]
    for j, h in enumerate(headers, 1):
        ws.cell(row=3, column=j+1, value=h).border = border
    ws.freeze_panes = "A4"

def _make_display_table_gj_m3(df_mj: pd.DataFrame) -> pd.DataFrame:
    df = df_mj.copy()
    for base_col in ["최근N년_평균공급량(MJ)", "최근N년_총공급량(MJ)", "예상공급량(MJ)", "보정_예상공급량(MJ)"]:
        if base_col not in df.columns: continue
        gj_col = base_col.replace("(MJ)", "(GJ)")
        m3_col = base_col.replace("(MJ)", "(㎥)")
        df[gj_col] = df[base_col].apply(mj_to_gj).round(0)
        df[m3_col] = df[base_col].apply(mj_to_m3).round(0)
    keep = ["연", "월", "일", "요일", "구분", "공휴일여부", "As-Is(기존)", "To-Be(보정)", "Diff(증감)", "일별비율", "is_outlier"]
    
    # Rename for display
    df_disp = df.rename(columns={
        "예상공급량(GJ)": "As-Is(기존)",
        "보정_예상공급량(GJ)": "To-Be(보정)"
    })
    if "To-Be(보정)" in df_disp.columns and "As-Is(기존)" in df_disp.columns:
        df_disp["Diff(증감)"] = df_disp["To-Be(보정)"] - df_disp["As-Is(기존)"]
        
    return df_disp[[c for c in keep if c in df_disp.columns]].copy()


# ─────────────────────────────────────────────
# 5. 핵심 분석 로직 (Daily) - 추세(Trend) 인자 추가
# ─────────────────────────────────────────────
def make_daily_plan_table(df_daily, df_plan, target_year, target_month, recent_window, apply_trend=False):
    cal_df = load_effective_calendar()
    plan_col = _find_plan_col(df_plan)
    
    all_years = sorted(df_daily["연도"].unique())
    start_year = target_year - recent_window
    candidate_years = [y for y in range(start_year, target_year) if y in all_years]
    
    df_pool = df_daily[(df_daily["연도"].isin(candidate_years)) & (df_daily["월"] == target_month)].copy()
    df_pool = df_pool.dropna(subset=["공급량(MJ)"])
    used_years = sorted(df_pool["연도"].unique().tolist())
    if not used_years: return None, None, [], pd.DataFrame()

    df_recent = df_daily[(df_daily["연도"].isin(used_years)) & (df_daily["월"] == target_month)].copy()
    df_recent = df_recent.dropna(subset=["공급량(MJ)"])
    if df_recent.empty: return None, None, used_years, pd.DataFrame()

    df_recent = df_recent.sort_values(["연도", "일"]).copy()
    df_recent["weekday_idx"] = df_recent["일자"].dt.weekday

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

    df_recent["month_total"] = df_recent.groupby("연도")["공급량(MJ)"].transform("sum")
    df_recent["ratio"] = df_recent["공급량(MJ)"] / df_recent["month_total"]
    df_recent["nth_dow"] = df_recent.sort_values(["연도", "일"]).groupby(["연도", "weekday_idx"]).cumcount() + 1

    def get_ratio_dict(mask):
        grp = df_recent[mask].groupby(["weekday_idx", "nth_dow"])["ratio"].mean().to_dict()
        fallback = df_recent[mask].groupby("weekday_idx")["ratio"].mean().to_dict()
        return grp, fallback

    w_grp, w_fb = get_ratio_dict(df_recent["is_weekend"])
    w1_grp, w1_fb = get_ratio_dict(df_recent["is_weekday1"])
    w2_grp, w2_fb = get_ratio_dict(df_recent["is_weekday2"])

    last_day = calendar.monthrange(target_year, target_month)[1]
    df_target = pd.DataFrame({"일자": pd.date_range(f"{target_year}-{target_month:02d}-01", periods=last_day)})
    df_target["연"] = target_year; df_target["월"] = target_month; df_target["일"] = df_target["일자"].dt.day
    df_target["weekday_idx"] = df_target["일자"].dt.weekday
    
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

    def _apply_ratio(r):
        k = (r["weekday_idx"], r["nth_dow"]); wd = r["weekday_idx"]
        if r["is_weekend"]: return w_grp.get(k, w_fb.get(wd, np.nan))
        if r["is_weekday1"]: return w1_grp.get(k, w1_fb.get(wd, np.nan))
        return w2_grp.get(k, w2_fb.get(wd, np.nan))

    df_target["raw"] = df_target.apply(_apply_ratio, axis=1).astype(float)
    overall_mean = df_target["raw"].mean()
    df_target["raw"] = df_target["raw"].fillna(overall_mean if pd.notna(overall_mean) else 1.0)
    
    # ─────────────────────────────────────────────────────────────
    # [NEW] 추세 적용 (Trend) 로직
    # 월초(1.05) ~ 월말(0.95) Linear Decay
    # ─────────────────────────────────────────────────────────────
    if apply_trend:
        # 1일부터 last_day까지 1.05 ~ 0.95로 감소하는 가중치 생성
        days = len(df_target)
        if days > 1:
            trend_factors = np.linspace(1.05, 0.95, days)
            df_target["raw"] = df_target["raw"] * trend_factors

    raw_sum = df_target["raw"].sum()
    df_target["일별비율"] = df_target["raw"] / raw_sum

    row_plan = df_plan[(df_plan["연"] == target_year) & (df_plan["월"] == target_month)]
    plan_total = float(row_plan[plan_col].iloc[0]) if not row_plan.empty else 0
    df_target["예상공급량(MJ)"] = (df_target["일별비율"] * plan_total).round(0)

    # Outlier 계산
    df_target["WeekNum"] = df_target["일자"].dt.isocalendar().week
    df_target["Group_Mean"] = df_target.groupby(["WeekNum", "is_weekend"])["예상공급량(MJ)"].transform("mean")
    df_target["Bound_Upper"] = df_target["Group_Mean"] * 1.10
    df_target["Bound_Lower"] = df_target["Group_Mean"] * 0.90
    df_target["is_outlier"] = (df_target["예상공급량(MJ)"] > df_target["Bound_Upper"]) | (df_target["예상공급량(MJ)"] < df_target["Bound_Lower"])
    
    df_target["최근N년_평균공급량(MJ)"] = 0
    df_target["최근N년_총공급량(MJ)"] = 0

    return df_target, None, used_years, None

def _build_year_daily_plan(df_daily, df_plan, target_year, recent_window):
    all_rows = []
    month_summary_rows = []
    plan_col = _find_plan_col(df_plan)
    
    for m in range(1, 13):
        res, _, _, _ = make_daily_plan_table(df_daily, df_plan, target_year, m, recent_window)
        row_plan = df_plan[(df_plan["연"] == target_year) & (df_plan["월"] == m)]
        plan_total_mj = float(row_plan[plan_col].iloc[0]) if not row_plan.empty else np.nan
        
        if res is not None:
             all_rows.append(res)
             month_summary_rows.append({
                "월": m,
                "월간 계획(GJ)": round(mj_to_gj(plan_total_mj), 0),
                "월간 계획(㎥)": round(mj_to_m3(plan_total_mj), 0)
             })

    if not all_rows: return pd.DataFrame(), pd.DataFrame()
    df_year = pd.concat(all_rows, ignore_index=True)
    return df_year, pd.DataFrame(month_summary_rows)

# ─────────────────────────────────────────────
# 6. UI 및 시각화 (형님 지시 100% 반영)
# ─────────────────────────────────────────────
def tab_daily_plan(df_daily: pd.DataFrame):
    st.subheader("📅 Daily 공급량 분석 — 최근 N년 패턴 기반 일별 계획")

    df_plan = load_monthly_plan()
    plan_col = _find_plan_col(df_plan)
    years_plan = sorted(df_plan["연"].unique())
    
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1: target_year = st.selectbox("계획 연도 선택", years_plan, index=len(years)-1)
    with col2: target_month = st.selectbox("월", range(1, 13))
    with col3: recent_window = st.slider("과거 참조(년)", 1, 10, 3)

    # [NEW] 추세 적용 버튼
    apply_trend = st.checkbox("📉 월말 감소 추세 적용 (기온 상승 고려: 월초 > 월말)", value=False)

    df_result, df_mat, used_years, df_debug = make_daily_plan_table(
        df_daily, df_plan, target_year, target_month, recent_window, apply_trend=apply_trend
    )

    if df_result is None: st.warning("데이터 부족"); return
    
    st.markdown(f"- 실제 학습 연도: {min(used_years)} ~ {max(used_years)}")
    plan_total_gj = mj_to_gj(df_result["예상공급량(MJ)"].sum())
    st.markdown(f"**{target_year}년 {target_month}월 합계:** `{plan_total_gj:,.0f} GJ`")

    # ─────────────── [보정 로직] ───────────────
    view = df_result.copy()
    view["보정_예상공급량(MJ)"] = view["예상공급량(MJ)"]
    
    st.divider()
    
    # ★★★ 그래프 자리 (Placeholder) ★★★
    chart_placeholder = st.empty()
    
    # ─────────────── [UI 하단] 버튼 및 설정 ───────────────
    _, col_btn = st.columns([5, 1]) 
    with col_btn:
        use_calib = st.checkbox("✅ 이상치 보정 활성화", value=False)
        
    diff_mj = 0
    mask_out = pd.Series([False]*len(view))
    mask_fix = pd.Series([False]*len(view)) # 보정(받는) 구간

    if use_calib:
        with st.expander("🛠️ 보정 구간 및 재배분 설정", expanded=True):
            min_d = view["일자"].min().date(); max_d = view["일자"].max().date()
            
            c1, c2 = st.columns(2)
            d_out = c1.date_input("1. 이상구간 (Outlier)", (min_d, min_d), min_value=min_d, max_value=max_d)
            d_fix = c2.date_input("2. 보정 구간 (Redistribution)", (min_d, max_d), min_value=min_d, max_value=max_d)
            
            # [NEW] 스마트 비율 추천 (Smart Default)
            suggested_rate = 0.0
            if isinstance(d_out, tuple) and len(d_out) == 2:
                s, e = d_out
                # 선택된 구간 중 가장 많이 튄(Bound Upper 초과율이 높은) 비율 찾기
                target_rows = view[(view["일자"].dt.date >= s) & (view["일자"].dt.date <= e)]
                if not target_rows.empty:
                    # (Value - Bound) / Value * 100 -> 감량해야할 비율 (음수)
                    # 여기서는 간단히 상한선 대비 초과분만큼 줄이는 걸 추천
                    max_dev = 0
                    for idx, row in target_rows.iterrows():
                        if row["예상공급량(MJ)"] > row["Bound_Upper"]:
                            dev = (row["Bound_Upper"] - row["예상공급량(MJ)"]) / row["예상공급량(MJ)"] * 100
                            if abs(dev) > abs(max_dev): max_dev = dev
                    suggested_rate = round(max_dev, 1)

            cal_rate = st.number_input("조정 비율 (%) - 자동추천됨", min_value=-50.0, max_value=50.0, value=suggested_rate, step=1.0)
            
            # [NEW] 평탄화 버튼
            do_smooth = st.checkbox("🌊 보정 구간 평탄화 (Smoothing) 적용")

            if isinstance(d_out, tuple) and len(d_out) == 2 and isinstance(d_fix, tuple) and len(d_fix) == 2:
                s_out, e_out = d_out
                s_fix, e_fix = d_fix
                
                mask_out = (view["일자"].dt.date >= s_out) & (view["일자"].dt.date <= e_out)
                mask_fix = (view["일자"].dt.date >= s_fix) & (view["일자"].dt.date <= e_fix)

                if mask_out.any():
                    # 1. 이상구간 조정
                    view.loc[mask_out, "보정_예상공급량(MJ)"] = view.loc[mask_out, "예상공급량(MJ)"] * (1 + cal_rate / 100.0)
                    
                    # 2. 차이 계산
                    diff_mj = (view.loc[mask_out, "예상공급량(MJ)"] - view.loc[mask_out, "보정_예상공급량(MJ)"]).sum()
                    
                    # 3. 보정 구간 배분
                    sum_r = view.loc[mask_fix, "일별비율"].sum()
                    if mask_fix.any() and sum_r > 0:
                        # (A) 기본 배분 (비율대로 더하기)
                        view.loc[mask_fix, "보정_예상공급량(MJ)"] += diff_mj * (view.loc[mask_fix, "일별비율"] / sum_r)
                        
                        # (B) 평탄화 옵션 켜지면 -> 해당 구간을 '이상적 패턴(Bound Mean)' 모양으로 맞춤
                        if do_smooth:
                            # 타겟 구간의 새로운 총량
                            target_new_total = view.loc[mask_fix, "보정_예상공급량(MJ)"].sum()
                            # 타겟 구간의 이상적 패턴 (Group Mean 사용)
                            ideal_pattern = view.loc[mask_fix, "Group_Mean"]
                            pattern_sum = ideal_pattern.sum()
                            if pattern_sum > 0:
                                view.loc[mask_fix, "보정_예상공급량(MJ)"] = ideal_pattern * (target_new_total / pattern_sum)

            st.caption(f"변동량: {mj_to_gj(diff_mj):,.0f} GJ")

    # ─────────────── [그래프 그리기] ───────────────
    view["예상공급량(GJ)"] = view["예상공급량(MJ)"].apply(mj_to_gj)
    view["보정_예상공급량(GJ)"] = view["보정_예상공급량(MJ)"].apply(mj_to_gj)
    view["Bound_Upper(GJ)"] = view["Bound_Upper"].apply(mj_to_gj)
    view["Bound_Lower(GJ)"] = view["Bound_Lower"].apply(mj_to_gj)

    fig = go.Figure()

    # 1. [기존 그래프]
    w1 = view[view["구분"] == "평일1(월·금)"].copy()
    w2 = view[view["구분"] == "평일2(화·수·목)"].copy()
    we = view[view["구분"] == "주말/공휴일"].copy()

    fig.add_trace(go.Bar(x=w1["일"], y=w1["예상공급량(GJ)"], name="평일1(월·금)", marker_color="#0059b3", width=0.8))
    fig.add_trace(go.Bar(x=w2["일"], y=w2["예상공급량(GJ)"], name="평일2(화·수·목)", marker_color="#87CEFA", width=0.8))
    fig.add_trace(go.Bar(x=we["일"], y=we["예상공급량(GJ)"], name="주말/공휴일", marker_color="#E60000", width=0.8))

    # 2. [보정 그래프] ★값이 변경된 모든 날짜(뺀 곳 + 더한 곳) 회색 표시★
    if use_calib:
        # 오차(1MJ) 이상 차이나는 모든 곳을 회색으로 덮음
        mask_changed = (abs(view["예상공급량(MJ)"] - view["보정_예상공급량(MJ)"]) > 1)
        if mask_changed.any():
            target_view = view[mask_changed]
            fig.add_trace(go.Bar(
                x=target_view["일"], 
                y=target_view["보정_예상공급량(GJ)"],
                marker_color="rgba(80, 80, 80, 0.7)", # 진한 회색
                name="보정됨(To-Be)",
                width=0.8
            ))

    fig.add_trace(go.Scatter(x=view["일"], y=view["일별비율"], yaxis="y2", name="비율", line=dict(color='#FF8A80', width=2)))
    fig.add_trace(go.Scatter(x=view["일"], y=view["Bound_Upper(GJ)"], mode='lines', line=dict(width=0), showlegend=False))
    fig.add_trace(go.Scatter(x=view["일"], y=view["Bound_Lower(GJ)"], mode='lines', line=dict(width=0), fill='tonexty', fillcolor='rgba(100,100,100,0.45)', name='범위(±10%)'))
    
    outliers = view[view["is_outlier"]]
    if not outliers.empty:
        fig.add_trace(go.Scatter(x=outliers["일"], y=outliers["예상공급량(GJ)"], mode='markers', marker=dict(color='red', symbol='x', size=10), name='Outlier'))

    fig.update_layout(
        title=f"{target_year}년 {target_month}월 공급계획",
        xaxis_title="일",
        yaxis=dict(title="공급량(GJ)"),
        yaxis2=dict(title="비율", overlaying="y", side="right"),
        barmode="overlay", 
        legend=dict(orientation="h", y=1.1)
    )
    
    chart_placeholder.plotly_chart(fig, use_container_width=True)
    
    # UI 순서: 1.분배기준 -> 2.월별 -> 3.일별
    st.divider()
    st.markdown("### 🧩 1. 일별 공급량 분배 기준")
    st.markdown(
        """
- **주말/공휴일/명절**: **'요일(토/일) + 그 달의 n번째' 기준 평균** (공휴일/명절도 주말 패턴으로 묶음)
- **평일**: '평일1(월·금)' / '평일2(화·수·목)'로 구분  
  기본은 **'요일 + 그 달의 n번째(1째 월요일, 2째 월요일...)' 기준 평균**
- 일부 케이스 데이터가 부족하면 **'요일 평균'으로 보정**
- 마지막에 **일별비율 합계가 1이 되도록 정규화(raw / SUM(raw))**
        """.strip()
    )

    st.markdown("#### 📌 2. 월별 계획량(1~12월) & 연간 총량")
    df_plan_h = make_month_plan_horizontal(df_plan, target_year, plan_col)
    show_table_no_index(format_table_generic(df_plan_h), height=160)

    st.markdown("#### 📋 3. 일별 비율, 예상 공급량 테이블")
    
    view_show = _make_display_table_gj_m3(view)
    if "is_outlier" in view_show.columns:
        view_show["is_outlier"] = view_show["is_outlier"].map({True: "🚨", False: ""})

    view_show = format_table_generic(view_show, percent_cols=["일별비율"])
    show_table_no_index(view_show, height=520)

    # ─────────────── [To-Be 다운로드] ───────────────
    st.markdown("#### 💾 6. 데이터 다운로드")
    
    col_down1, col_down2 = st.columns(2)
    
    with col_down1:
        if use_calib:
            st.info("💡 보정된(To-Be) 데이터를 다운로드할 수 있습니다.")
            buffer_tobe = BytesIO()
            # 헤더 정리
            download_df = view[["일자", "요일", "구분", "예상공급량(GJ)", "보정_예상공급량(GJ)", "is_outlier"]].copy()
            download_df.columns = ["일자", "요일", "구분", "As-Is(기존)", "To-Be(보정)", "Outlier"]
            download_df["Diff(증감)"] = download_df["To-Be(보정)"] - download_df["As-Is(기존)"]
            
            with pd.ExcelWriter(buffer_tobe, engine="openpyxl") as writer:
                download_df.to_excel(writer, index=False, sheet_name="To-Be_일별계획")
            st.download_button(
                label="📥 To-Be(보정후) 일별계획 다운로드", 
                data=buffer_tobe.getvalue(), 
                file_name=f"{target_year}_{target_month:02d}_ToBe_일별공급계획.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            buffer = BytesIO()
            excel_df = _make_display_table_gj_m3(view)
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                excel_df.to_excel(writer, index=False, sheet_name="일별계획")
            st.download_button(
                label="📥 일별계획 다운로드 (As-Is)", 
                data=buffer.getvalue(), 
                file_name=f"{target_year}_{target_month:02d}_일별공급계획.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with col_down2:
        st.markdown("**🗂️ 연간 전체 계획 다운로드**")
        annual_year = st.selectbox("연간 계획 연도", years_plan, index=years_plan.index(target_year) if target_year in years_plan else 0)
        buffer_year = BytesIO()
        df_year_daily, df_month_summary = _build_year_daily_plan(df_daily, df_plan, int(annual_year), recent_window)
        with pd.ExcelWriter(buffer_year, engine="openpyxl") as writer:
            df_year_daily.to_excel(writer, index=False, sheet_name="연간")
            df_month_summary.to_excel(writer, index=False, sheet_name="월 요약")
            _add_cumulative_status_sheet(writer.book, int(annual_year))
        st.download_button(
            label="📥 연간 계획 다운로드", 
            data=buffer_year.getvalue(), 
            file_name=f"{annual_year}_연간_일별공급계획.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            key="download_annual_excel"
        )


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def main():
    df, _ = load_daily_data()
    mode = st.sidebar.radio("좌측 탭 선택", ("📅 Daily 공급량 분석",), index=0)
    if mode == "📅 Daily 공급량 분석":
        st.title("도시가스 공급량 — 일별계획 예측")
        tab_daily_plan(df_daily=df)

if __name__ == "__main__":
    main()
