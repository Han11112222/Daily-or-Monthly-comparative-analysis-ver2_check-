import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
import requests
from pathlib import Path
from urllib.parse import quote
from sklearn.linear_model import LinearRegression
from typing import Dict, List, Optional, Tuple

# ─────────────────────────────────────────────────────────
# 🟢 기본 설정 & 폰트
# ─────────────────────────────────────────────────────────
st.set_page_config(page_title="도시가스 계획/실적 분석", layout="wide")

def set_korean_font():
    ttf = Path(__file__).parent / "NanumGothic-Regular.ttf"
    if ttf.exists():
        try:
            import matplotlib as mpl
            mpl.font_manager.fontManager.addfont(str(ttf))
            mpl.rcParams["font.family"] = "NanumGothic"
            mpl.rcParams["axes.unicode_minus"] = False
        except: pass

set_korean_font()

# 🟢 설정 정보
GITHUB_USER = "HanYeop"
REPO_NAME = "GasProject"
DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"

# 🟢 용도 매핑
USE_COL_TO_GROUP = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "일반용": "영업용",
    "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용",
    "수송용(CNG)": "수송용", "수송용(BIO)": "수송용",
    "열병합용": "열병합", "열병합용1": "열병합", "열병합용2": "열병합",
    "연료전지용": "연료전지", "열전용설비용": "열전용설비용"
}

# ─────────────────────────────────────────────────────────
# 1. 데이터 로드 및 전처리
# ─────────────────────────────────────────────────────────
def _clean_base(df):
    out = df.copy()
    if "Unnamed: 0" in out.columns: out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def make_long(plan_df, actual_df):
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)
    records = []
    
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue
            group = USE_COL_TO_GROUP.get(col)
            if not group: continue
            
            base = df[["연", "월"]].copy()
            base["그룹"] = group
            base["용도"] = col
            base["계획/실적"] = label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            records.append(base)
            
    if not records: return pd.DataFrame()
    long_df = pd.concat(records, ignore_index=True)
    return long_df.dropna(subset=["연", "월"])

def load_data_simple(uploaded_file=None):
    try:
        if uploaded_file:
            return pd.ExcelFile(uploaded_file, engine='openpyxl')
        elif Path(DEFAULT_SALES_XLSX).exists():
            return pd.ExcelFile(DEFAULT_SALES_XLSX, engine='openpyxl')
        return None
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")
        return None

# ─────────────────────────────────────────────────────────
# 2. [기능 1] 실적 분석
# ─────────────────────────────────────────────────────────
def render_analysis_dashboard(long_df, unit_label):
    st.subheader(f"📊 실적 분석 ({unit_label})")
    
    df_act = long_df[long_df['계획/실적'] == '실적'].copy()
    df_act = df_act[df_act['연'] <= 2025] 
    
    all_years = sorted(df_act['연'].unique())
    if not all_years:
        st.error("분석할 실적 데이터가 없습니다.")
        return

    default_years = all_years[-3:] if len(all_years) >= 3 else all_years
    
    st.markdown("##### 📅 분석할 연도를 선택하세요 (다중 선택)")
    selected_years = st.multiselect(
        "연도 선택",
        options=all_years,
        default=default_years,
        label_visibility="collapsed"
    )
    
    if not selected_years:
        st.warning("연도를 1개 이상 선택해주세요.")
        return

    df_filtered = df_act[df_act['연'].isin(selected_years)]
    st.markdown("---")

    # [그래프 1] 월별 실적 추이
    st.markdown(f"#### 📈 월별 실적 추이 ({', '.join(map(str, selected_years))})")
    df_mon_compare = df_filtered.groupby(['연', '월'])['값'].sum().reset_index()
    
    fig1 = px.line(
        df_mon_compare, x='월', y='값', color='연', markers=True,
        title="월별 실적 추이 비교"
    )
    fig1.update_layout(xaxis=dict(tickmode='linear', dtick=1), yaxis_title=unit_label)
    st.plotly_chart(fig1, use_container_width=True)
    
    st.markdown("##### 📋 월별 상세 수치")
    pivot_mon = df_mon_compare.pivot(index='월', columns='연', values='값').fillna(0)
    st.dataframe(pivot_mon.style.format("{:,.0f}"), use_container_width=True)
    
    st.markdown("---")

    # [그래프 2] 연도별 용도 누적
    st.markdown(f"#### 🧱 연도별 용도 구성비 ({', '.join(map(str, selected_years))})")
    df_yr_usage = df_filtered.groupby(['연', '그룹'])['값'].sum().reset_index()
    
    fig2 = px.bar(
        df_yr_usage, x='연', y='값', color='그룹',
        title="연도별 판매량 및 용도 구성", text_auto='.2s'
    )
    fig2.update_layout(xaxis_type='category', yaxis_title=unit_label)
    st.plotly_chart(fig2, use_container_width=True)
    
    st.markdown("##### 📋 용도별 상세 수치")
    pivot_usage = df_yr_usage.pivot(index='연', columns='그룹', values='값').fillna(0)
    pivot_usage['합계'] = pivot_usage.sum(axis=1)
    st.dataframe(pivot_usage.style.format("{:,.0f}"), use_container_width=True)

# ─────────────────────────────────────────────────────────
# 3. [기능 2] 2035 예측 (5가지 추세 분석 모델 적용)
# ─────────────────────────────────────────────────────────
def holt_linear_trend(y, n_preds):
    """지수 평활법 (Holt's Linear Trend) 간단 구현"""
    if len(y) < 2: return np.full(n_preds, y[0])
    
    alpha = 0.8  # 최근 데이터 가중치 (0~1)
    beta = 0.2   # 추세 가중치 (0~1)
    
    level = y[0]
    trend = y[1] - y[0]
    
    # 학습
    for val in y[1:]:
        prev_level = level
        level = alpha * val + (1 - alpha) * (prev_level + trend)
        trend = beta * (level - prev_level) + (1 - beta) * trend
        
    # 예측
    preds = []
    for i in range(1, n_preds + 1):
        preds.append(level + i * trend)
    return np.array(preds)

def render_prediction_2035(long_df, unit_label):
    st.subheader(f"🔮 2035 장기 예측 ({unit_label})")
    
    # 🔴 예측 방법 선택 (5가지)
    st.markdown("##### 📊 추세 분석 모델 선택")
    pred_method = st.radio(
        "분석 방법",
        [
            "1. 선형 추세 (Linear)", 
            "2. 2차 곡선 (Quadratic)", 
            "3. 로그 추세 (Logarithmic)",
            "4. 지수 평활 (Holt's Trend)",
            "5. 연평균 성장률 (CAGR)"
        ],
        index=0,
        horizontal=True
    )

    # 모델 설명
    if "선형" in pred_method:
        st.info("💡 **[선형 추세]** 가장 기본적인 분석법으로, 일정한 기울기로 증가하거나 감소하는 직선 패턴을 예측합니다.")
    elif "2차" in pred_method:
        st.info("💡 **[2차 곡선]** 변화의 속도가 가속화되거나 둔화되는 곡선 패턴을 예측합니다.")
    elif "로그" in pred_method:
        st.info("💡 **[로그 추세]** 초반에는 빠르게 변하다가 점차 완만해지는(성숙기 진입) 패턴을 예측합니다.")
    elif "지수 평활" in pred_method:
        st.info("💡 **[지수 평활(Holt)]** 가장 최근의 실적 데이터와 추세에 더 높은 가중치를 두어 민감하게 예측합니다.")
    else:
        st.info("💡 **[CAGR]** 과거 기간의 연평균 성장률(%)을 그대로 적용하여 미래에도 같은 비율로 성장한다고 가정합니다.")

    # 데이터 준비
    df_act = long_df[(long_df['계획/실적'] == '실적') & (long_df['연'] <= 2025)].copy()
    df_train = df_act.groupby(['연', '그룹'])['값'].sum().reset_index()
    
    groups = df_train['그룹'].unique()
    future_years = np.arange(2026, 2036).reshape(-1, 1)
    results = []
    
    progress = st.progress(0)
    
    for i, grp in enumerate(groups):
        sub = df_train[df_train['그룹'] == grp]
        if len(sub) < 2: continue
        
        # 최근 5년 데이터 사용 (트렌드 반영 보정)
        sub_recent = sub.tail(5)
        if len(sub_recent) < 2: sub_recent = sub
            
        X = sub_recent['연'].values
        y = sub_recent['값'].values
        
        pred = []
        
        # 🟢 1. 선형 회귀 (Linear)
        if "선형" in pred_method:
            model = LinearRegression()
            model.fit(X.reshape(-1,1), y)
            pred = model.predict(future_years)
            
        # 🟢 2. 2차 곡선 (Quadratic)
        elif "2차" in pred_method:
            try:
                coeffs = np.polyfit(X, y, 2)
                p = np.poly1d(coeffs)
                pred = p(future_years.flatten())
            except: # 에러시 선형 대체
                model = LinearRegression()
                model.fit(X.reshape(-1,1), y)
                pred = model.predict(future_years)

        # 🟢 3. 로그 추세 (Logarithmic)
        elif "로그" in pred_method:
            try:
                # x축을 log로 변환하여 선형회귀 (Y = a + b * ln(X))
                # 연도 숫자가 크므로(2025) log값 차이가 미미할 수 있어, index(1,2,3...)로 변환해서 적용
                X_idx = np.arange(1, len(X) + 1).reshape(-1, 1)
                X_future_idx = np.arange(len(X) + 1, len(X) + 11).reshape(-1, 1)
                
                model = LinearRegression()
                model.fit(np.log(X_idx), y)
                pred = model.predict(np.log(X_future_idx))
            except:
                model = LinearRegression()
                model.fit(X.reshape(-1,1), y)
                pred = model.predict(future_years)

        # 🟢 4. 지수 평활 (Holt's Trend)
        elif "지수 평활" in pred_method:
            pred = holt_linear_trend(y, n_preds=10)

        # 🟢 5. CAGR
        else:
            try:
                start_val = y[0] if y[0] > 0 else 1
                end_val = y[-1]
                n = len(y) - 1
                if n > 0 and start_val > 0 and end_val > 0:
                    cagr = (end_val / start_val) ** (1/n) - 1
                else:
                    cagr = 0
                
                current_val = end_val
                temp_pred = []
                for _ in range(10):
                    current_val = current_val * (1 + cagr)
                    temp_pred.append(current_val)
                pred = np.array(temp_pred)
            except:
                pred = np.full(10, y[-1])

        # 음수 방지
        pred = [max(0, p) for p in pred]

        # 결과 저장
        for yr, v in zip(sub['연'], sub['값']):
            results.append({'연': yr, '그룹': grp, '판매량': v, 'Type': '실적'})
        for yr, v in zip(future_years.flatten(), pred):
            results.append({'연': yr, '그룹': grp, '판매량': v, 'Type': '예측'})
            
        progress.progress((i+1)/len(groups))
    progress.empty()
    
    df_res = pd.DataFrame(results)
    
    # 그래프 1: 전체 추세선
    st.markdown("#### 📈 전체 장기 전망 (추세선)")
    fig_line = px.line(
        df_res, x='연', y='판매량', color='그룹', 
        line_dash='Type', markers=True,
        title=f"용도별 장기 추세 ({unit_label}) - {pred_method.split('.')[1]}"
    )
    fig_line.add_vrect(x0=2025.5, x1=2035.5, fillcolor="green", opacity=0.1, annotation_text="예측 구간")
    st.plotly_chart(fig_line, use_container_width=True)
    
    st.markdown("---")
    
    # 그래프 2: 스택바
    st.markdown("#### 🧱 2035년 미래 예측 상세")
    df_forecast_only = df_res[df_res['Type'] == '예측']
    
    fig_stack = px.bar(
        df_forecast_only, x='연', y='판매량', color='그룹',
        title=f"향후 10년 공급량 예측 구성비", text_auto='.2s'
    )
    fig_stack.update_layout(xaxis_type='category', yaxis_title=unit_label)
    st.plotly_chart(fig_stack, use_container_width=True)
    
    # 표 & 다운로드
    st.markdown("##### 📋 상세 데이터")
    piv = df_forecast_only.pivot_table(index='연', columns='그룹', values='판매량')
    piv['합계'] = piv.sum(axis=1)
    
    st.dataframe(piv.style.format("{:,.0f}"), use_container_width=True)
    st.download_button(
        label="💾 예측 데이터 다운로드",
        data=piv.to_csv().encode('utf-8-sig'),
        file_name=f"forecast_2035.csv",
        mime="text/csv"
    )

# ─────────────────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("🔥 도시가스 판매량 분석 & 예측")
    
    with st.sidebar:
        st.header("설정")
        uploaded = None
        if not Path(DEFAULT_SALES_XLSX).exists():
            st.warning(f"⚠️ '{DEFAULT_SALES_XLSX}' 파일이 없습니다.")
            uploaded = st.file_uploader("엑셀 파일 업로드", type="xlsx")
        else:
            st.success(f"✅ '{DEFAULT_SALES_XLSX}' 파일 연결됨")
            if st.checkbox("다른 파일 업로드하기"):
                uploaded = st.file_uploader("엑셀 파일 업로드", type="xlsx")

        st.markdown("---")
        mode = st.radio("분석 모드", ["1. 실적 분석", "2. 2035 예측"])
        unit = st.radio("단위", ["부피 (천m?)", "열량 (GJ)"])

    xls = load_data_simple(uploaded)
    if xls is None: return

    try:
        if unit.startswith("부피"):
            df_p = xls.parse("계획_부피")
            df_a = xls.parse("실적_부피")
            unit_label = "천m?"
        else:
            df_p = xls.parse("계획_열량")
            df_a = xls.parse("실적_열량")
            unit_label = "GJ"
            
        long_df = make_long(df_p, df_a)
        
    except Exception as e:
        st.error(f"시트 로드 실패: {e}")
        return

    if mode.startswith("1"):
        render_analysis_dashboard(long_df, unit_label)
    else:
        render_prediction_2035(long_df, unit_label)

if __name__ == "__main__":
    main()
