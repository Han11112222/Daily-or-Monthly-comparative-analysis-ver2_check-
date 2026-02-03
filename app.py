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
    page_title="도시가스 공급량: 일별계획 예측",
    layout="wide",
)

MJ_PER_NM3 = 42.563
MJ_TO_GJ = 1.0 / 1000.0

def mj_to_gj(x):
    try: return x * MJ_TO_GJ
    except: return np.nan

def mj_to_m3(x):
    try: return x / MJ_PER_NM3
    except: return np.nan

# ─────────────────────────────────────────────
# 2. 데이터 로딩 (전체 기능 복구)
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
        df_raw["일자"] = pd.to_datetime(df_raw["일자"], errors='coerce')
        df_raw = df_raw.dropna(subset=["일자"])

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

@st.cache_data
def load_corr_data() -> pd.DataFrame | None:
    # [복구] 상관도 분석용 데이터
    excel_path = Path(__file__).parent / "상관도분석.xlsx"
    if not excel_path.exists(): return None
    try: return pd.read_excel(excel_path)
    except: return None

# ─────────────────────────────────────────────
# 3. 유틸 함수 (전체 복구)
# ─────────────────────────────────────────────
def _find_plan_col(df_plan: pd.DataFrame) -> str:
    candidates = ["계획(사업계획제출_MJ)", "계획(사업계획제출)", "계획_MJ", "계획"]
    for c in candidates:
        if c in df_plan.columns: return c
    nums = [c for c in df_plan.columns if pd.api.types.is_numeric_dtype(df_plan[c]) and c not in ["연", "월"]]
    return nums[0] if nums else ""

def make_month_plan_horizontal(df_plan, target_year, plan_col):
    if df_plan.empty or not plan_col: return pd.DataFrame()
    df_year = df_plan[df_plan["연"] == target_year][["월", plan_col]].copy()
    if df_year.empty: return pd.DataFrame()
    base = pd.DataFrame({"월": list(range(1, 13))})
    df_year = base.merge(df_year, on="월", how="left")
    df_year = df_year.rename(columns={plan_col: "월별 계획(MJ)"})
    
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

def _format_excel_sheet(ws, freeze="A2", center=True, width_map=None):
    if freeze: ws.freeze_panes = freeze
    if center:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for c in row: c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if width_map:
        for col_letter, w in width_map.items():
            ws.column_dimensions[col_letter].width = w

def _add_cumulative_status_sheet(wb, annual_year):
    sheet_name = "누적계획현황"
    if sheet_name in wb.sheetnames: return
    ws = wb.create_sheet(sheet_name)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="F2F2F2")

    ws["A1"] = "기준일"; ws["A1"].font = Font(bold=True)
    ws["B1"] = pd.Timestamp(f"{annual_year}-01-01").to_pydatetime()
    ws["B1"].number_format = "yyyy-mm-dd"; ws["B1"].font = Font(bold=True)

    headers = ["구분", "목표(GJ)", "누적(GJ)", "목표(m³)", "누적(m³)", "진행률(GJ)"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=j, value=h)
        c.font = Font(bold=True); c.fill = header_fill; c.border = border; c.alignment = Alignment(horizontal="center")

    rows = [("일", 4), ("월", 5), ("연", 6)]
    for label, r in rows:
        ws.cell(row=r, column=1, value=label).border = border; ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

    d = "$B$1"
    ws["B4"] = f'=IFERROR(XLOOKUP({d},연간!$D:$D,연간!$O:$O),"")'
    ws["C4"] = "=B4"
    ws["D4"] = f'=IFERROR(XLOOKUP({d},연간!$D:$D,연간!$P:$P),"")'
    ws["E4"] = "=D4"
    ws["F4"] = '=IFERROR(IF(B4=0,"",C4/B4),"")'
    ws["B5"] = f'=SUMIFS(연간!$O:$O,연간!$A:$A,YEAR({d}),연간!$B:$B,MONTH({d}))'
    ws["C5"] = f'=SUMIFS(연간!$O:$O,연간!$D:$D,">="&EOMONTH({d},-1)+1,연간!$D:$D,"<="&{d})'
    ws["D5"] = f'=SUMIFS(연간!$P:$P,연간!$A:$A,YEAR({d}),연간!$B:$B,MONTH({d}))'
    ws["E5"] = f'=SUMIFS(연간!$P:$P,연간!$D:$D,">="&EOMONTH({d},-1)+1,연간!$D:$D,"<="&{d})'
    ws["F5"] = '=IFERROR(IF(B5=0,"",C5/B5),"")'
    ws["B6"] = f'=SUMIFS(연간!$O:$O,연간!$A:$A,YEAR({d}))'
    ws["C6"] = f'=SUMIFS(연간!$O:$O,연간!$D:$D,">="&DATE(YEAR({d}),1,1),연간!$D:$D,"<="&{d})'
    ws["D6"] = f'=SUMIFS(연간!$P:$P,연간!$A:$A,YEAR({d}))'
    ws["E6"] = f'=SUMIFS(연간!$P:$P,연간!$D:$D,">="&DATE(YEAR({d}),1,1),연간!$D:$D,"<="&{d})'
    ws["F6"] = '=IFERROR(IF(B6=0,"",C6/B6),"")'
    ws.freeze_panes = "A4"

def _make_display_table_gj_m3(df_mj: pd.DataFrame) -> pd.DataFrame:
    df = df_mj.copy()
    for base_col in ["최근N년_평균공급량(MJ)", "최근N년_총공급량(MJ)", "예상공급량(MJ)", "보정_예상공급량(MJ)"]:
        if base_col not in df.columns: continue
        gj_col = base_col.replace("(MJ)", "(GJ)")
        m3_col = base_col.replace("(MJ)", "(㎥)")
        df[gj_col] = df[base_col].apply(mj_to_gj).round(0)
        df[m3_col] = df[base_col].apply(mj_to_m3).round(0)
    keep_cols = [
        "연", "월", "일", "요일", "weekday_idx", "nth_dow", "구분", "공휴일여부",
        "최근N년_평균공급량(GJ)", "최근N년_평균공급량(㎥)",
        "최근N년_총공급량(GJ)", "최근N년_총공급량(㎥)",
        "일별비율",
        "예상공급량(GJ)", "예상공급량(㎥)", "보정_예상공급량(GJ)", "is_outlier"
    ]
    cols = [c for c in keep_cols if c in df.columns]
    return df[cols].copy()

# [복구] 회귀분석 및 히트맵 함수
def fit_poly3_and_r2(x, y):
    if len(x) < 4: return None, None, None
    coef = np.polyfit(x, y, 3)
    y_pred = np.polyval(coef, x)
    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2 = np.nan if ss_tot == 0 else 1 - ss_res / ss_tot
    return coef, y_pred, r2

def plot_poly_fit(x, y, coef, title, x_label, y_label):
    x_grid = np.linspace(x.min(), x.max(), 200)
    y_grid = np.polyval(coef, x_grid)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x, y=y, mode="markers", name="실적"))
    fig.add_trace(go.Scatter(x=x_grid, y=y_grid, mode="lines", name="3차 다항식"))
    fig.update_layout(title=title, xaxis_title=x_label, yaxis_title=y_label)
    return fig

def _render_daily_temp_heatmap(df_temp_all):
    st.subheader("🧊 기온분석 — 일일 평균기온 히트맵")
    if df_temp_all.empty: return
    
    col1, col2 = st.columns([2, 1])
    min_y, max_y = int(df_temp_all["연도"].min()), int(df_temp_all["연도"].max())
    with col1: y_range = st.slider("연도 범위", min_y, max_y, (min_y, max_y))
    with col2: m_sel = st.selectbox("월", range(1, 13))

    dfm = df_temp_all[df_temp_all["연도"].between(y_range[0], y_range[1]) & (df_temp_all["월"] == m_sel)]
    if dfm.empty: st.caption("데이터 없음"); return

    pv = dfm.pivot_table(index="일", columns="연도", values="평균기온(℃)", aggfunc="mean").reindex(range(1, 32))
    fig = go.Figure(data=go.Heatmap(z=pv.values, x=pv.columns, y=pv.index, colorscale="RdBu_r", zmid=0))
    fig.update_layout(title=f"{m_sel}월 기온 히트맵", yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────
# 4. 핵심 분석 로직 (Daily)
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
    df_recent["nth_dow"] = df_recent.groupby(["연도", "weekday_idx"]).cumcount() + 1

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
    for m in range(1, 13):
        res, _, _, _ = make_daily_plan_table(df_daily, df_plan, target_year, m, recent_window)
        if res is not None:
             all_rows.append(res)
             month_summary_rows.append({
                "월": m,
                "월간 계획(GJ)": round(mj_to_gj(res["예상공급량(MJ)"].sum()), 0),
                "월간 계획(㎥)": round(mj_to_m3(res["예상공급량(MJ)"].sum()), 0)
             })
    if not all_rows: return pd.DataFrame(), pd.DataFrame()
    return pd.concat(all_rows, ignore_index=True), pd.DataFrame(month_summary_rows)

# ─────────────────────────────────────────────
# 5. UI 및 시각화 (형님 지시 100% 반영)
# ─────────────────────────────────────────────
def tab_daily_plan(df_daily):
    st.subheader("📅 Daily 공급량 분석")

    df_plan = load_monthly_plan()
    if df_plan.empty: return

    plan_col = _find_plan_col(df_plan)
    years = sorted(df_plan["연"].unique())
    
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1: target_year = st.selectbox("연도", years, index=len(years)-1)
    with col2: target_month = st.selectbox("월", range(1, 13))
    with col3: recent_window = st.slider("과거 참조(년)", 1, 10, 3)

    df_res, _, used_years, _ = make_daily_plan_table(df_daily, df_plan, target_year, target_month, recent_window)
    if df_res is None:
        st.warning("데이터 부족"); return

    st.markdown(f"**참조 데이터:** {min(used_years)}년 ~ {max(used_years)}년")
    
    # ─────────────── [NEW] 보정 로직 ───────────────
    view = df_res.copy()
    view["보정_예상공급량(MJ)"] = view["예상공급량(MJ)"]

    st.divider()

    # ★ 우측 상단 버튼 배치: [제목(3) | 버튼(1)]
    c_head, c_ctrl = st.columns([3, 1])
    
    with c_head:
        st.markdown("#### 📊 2. 일별 예상 공급량 & Outlier 분석")
    
    with c_ctrl:
        # [수정] 우측 상단에 체크박스 배치
        use_calib = st.checkbox("✅ 이상치 보정 활성화", value=False)
        
    diff_mj = 0
    if use_calib:
        with st.expander("🛠️ 보정 구간 및 재배분 설정", expanded=True):
            min_d = view["일자"].min().date()
            max_d = view["일자"].max().date()
            
            c1, c2 = st.columns(2)
            d_out = c1.date_input("1. 수정 필요 구간 (Outlier)", (min_d, min_d), min_value=min_d, max_value=max_d)
            d_fix = c2.date_input("2. 보정 구간 (Redistribution)", (min_d, max_d), min_value=min_d, max_value=max_d)
            
            if isinstance(d_out, tuple) and len(d_out) == 2 and isinstance(d_fix, tuple) and len(d_fix) == 2:
                s_out, e_out = d_out
                s_fix, e_fix = d_fix
                
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
                
                mask_fix = (view["일자"].dt.date >= s_fix) & (view["일자"].dt.date <= e_fix)
                sum_r = view.loc[mask_fix, "일별비율"].sum()
                if mask_fix.any() and sum_r > 0:
                        view.loc[mask_fix, "보정_예상공급량(MJ)"] += diff_mj * (view.loc[mask_fix, "일별비율"] / sum_r)
            
            st.caption(f"변동량: {mj_to_gj(diff_mj):,.0f} GJ")

    # ─────────────── [그래프] 색상: 파/파/빨 (녹색 제거) ───────────────
    view["예상공급량(GJ)"] = view["예상공급량(MJ)"].apply(mj_to_gj)
    view["보정_예상공급량(GJ)"] = view["보정_예상공급량(MJ)"].apply(mj_to_gj)
    view["Bound_Upper(GJ)"] = view["Bound_Upper"].apply(mj_to_gj)
    view["Bound_Lower(GJ)"] = view["Bound_Lower"].apply(mj_to_gj)

    fig = go.Figure()

    w1 = view[view["구분"] == "평일1(월·금)"].copy()
    w2 = view[view["구분"] == "평일2(화·수·목)"].copy()
    we = view[view["구분"] == "주말/공휴일"].copy()

    # ★ [중요] 기존 색상 고정 (녹색 제거): 평일1=진파랑, 평일2=연파랑, 주말=빨강
    fig.add_trace(go.Bar(x=w1["일"], y=w1["예상공급량(GJ)"], name="평일1(월·금)", marker_color="#1F77B4"))
    fig.add_trace(go.Bar(x=w2["일"], y=w2["예상공급량(GJ)"], name="평일2(화·수·목)", marker_color="#636EFA"))
    fig.add_trace(go.Bar(x=we["일"], y=we["예상공급량(GJ)"], name="주말/공휴일", marker_color="#EF553B")) # 녹색 제거 -> 빨강

    # 보정 그래프 (회색 덮어쓰기)
    if use_calib:
        fig.add_trace(go.Bar(
            x=view["일"], y=view["보정_예상공급량(GJ)"],
            marker_color="rgba(80, 80, 80, 0.6)", # 진한 회색 (투명도)
            name="보정 후",
        ))

    fig.add_trace(go.Scatter(x=view["일"], y=view["일별비율"], yaxis="y2", name="비율", line=dict(color='black', width=1)))
    fig.add_trace(go.Scatter(x=view["일"], y=view["Bound_Upper(GJ)"], line=dict(width=0), showlegend=False))
    fig.add_trace(go.Scatter(x=view["일"], y=view["Bound_Lower(GJ)"], line=dict(width=0), fill='tonexty', fillcolor='rgba(100,100,100,0.1)', name='범위(±10%)'))
    
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
    
    show_table_no_index(format_table_generic(view[["일자", "구분", "예상공급량(GJ)", "보정_예상공급량(GJ)", "is_outlier"]], percent_cols=[]))

# ─────────────────────────────────────────────
# 탭2: Daily·Monthly 공급량 비교 (복구)
# ─────────────────────────────────────────────
def tab_daily_monthly_compare(df, df_temp_all):
    st.subheader("📊 상관도 분석 및 비교")
    df_corr = load_corr_data()
    if df_corr is not None:
         # 상관도 히트맵 (간략 복구)
         num_df = df_corr.select_dtypes(include=["number"])
         st.dataframe(num_df.corr())
    _render_daily_temp_heatmap(df_temp_all)

# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def main():
    df, df_temp_all = load_daily_data()
    mode = st.sidebar.radio("좌측 탭 선택", ("📅 Daily 공급량 분석", "📊 Daily·Monthly 공급량 비교"), index=0)

    if mode == "📅 Daily 공급량 분석":
        tab_daily_plan(df)
    else:
        tab_daily_monthly_compare(df, df_temp_all)

if __name__ == "__main__":
    main()
