"""해외법인 손익 리포트 - Streamlit 웹앱 버전.

Report_P1 / Report_P2 (엑셀 버전)와 동일한 계산 정의를 report_logic.py에서
가져와 사용한다. 이 앱은 엑셀의 표를 그대로 웹에서 보여주는 것에 더해,
Page 2의 차이내역(물량/롤마진/제경비/영업외)을 막대그래프로 시각화한다.
"""

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from report_logic import CORPS, compute_report, load_raw, period_options

st.set_page_config(page_title="해외법인 손익 리포트", layout="wide")

NAVY = "#1F3864"
GREEN = "#2E7D32"
RED = "#C62828"
GREY = "#8C8C8C"
COMPONENT_COLORS = {
    "물량차이": "#4C78A8",
    "롤마진차이": "#72B7B2",
    "제경비차이": "#E45756",
    "영업외차이": "#F2A65A",
}

ROW_ORDER = CORPS + ["합계"]

# 권역별 그룹핑 - 표 배경색과 그래프 배경 밴드에 동일하게 사용
REGION_MEMBERS = {
    "중국": ["중국1", "중국2", "중국3", "중국4"],
    "인도": ["인도1", "인도2", "인도3"],
    "유럽": ["유럽1", "유럽2", "유럽3", "유럽4"],
    "미주": ["미주1", "미주2", "미주3", "미주4"],
}
REGION_COLORS = {
    "중국": "#DCE9F9",  # light blue
    "인도": "#FDEBD3",  # light amber
    "유럽": "#E2F0E4",  # light green
    "미주": "#F0E4F7",  # light lavender
}
CORP_REGION = {corp: region for region, members in REGION_MEMBERS.items() for corp in members}


@st.cache_data
def _load_period_options():
    df = load_raw()
    return period_options(df)


GROUP_DIVIDER = "3px solid #1F3864"


def style_table(df: pd.DataFrame, decimals=1, group_dividers=()) -> "pd.io.formats.style.Styler":
    """group_dividers: 왼쪽에 굵은 구분선을 그릴 컬럼명 목록 (계획/실적/차이 등 열 그룹 구분용)."""
    numeric_cols = [c for c in df.columns if c != "법인"]

    def row_background(row):
        corp = row["법인"]
        if corp == "합계":
            style = "font-weight: bold; background-color: #D9E1F2"
        else:
            region_color = REGION_COLORS.get(CORP_REGION.get(corp))
            style = f"background-color: {region_color}" if region_color else ""
        return [style for _ in row]

    def color_negative(v):
        if isinstance(v, (int, float)) and v < 0:
            return f"color: {RED}"
        return ""

    fmt = {c: f"{{:,.{decimals}f}}" for c in numeric_cols}
    styler = (
        df.style.format(fmt)
        .apply(row_background, axis=1)
        .map(color_negative, subset=numeric_cols)
    )
    if "검증(차이내역합-경상이익차이)" in df.columns:
        check_col = "검증(차이내역합-경상이익차이)"
        styler = styler.map(
            lambda v: "background-color: #FFEB9C" if abs(v) > 1e-6 else "",
            subset=[check_col],
        )
    for col in group_dividers:
        styler = styler.set_properties(subset=[col], **{"border-left": GROUP_DIVIDER})
        styler = styler.set_table_styles(
            [{"selector": f"th.col_heading.col{df.columns.get_loc(col)}", "props": f"border-left: {GROUP_DIVIDER}"}],
            overwrite=False,
        )
    return styler


def top_diff_table(page1: pd.DataFrame, n=4) -> pd.DataFrame:
    """경상이익차이 절대값 기준 상위 n개 법인 + 기타 합계 + 합산, 세로 (n+2)행."""
    corp_rows = page1[page1["법인"] != "합계"][["법인", "차이-경상이익"]].copy()
    corp_rows["절대값"] = corp_rows["차이-경상이익"].abs()
    top = corp_rows.sort_values("절대값", ascending=False, kind="stable").head(n)
    others_sum = corp_rows.loc[~corp_rows["법인"].isin(top["법인"]), "차이-경상이익"].sum()
    total_sum = corp_rows["차이-경상이익"].sum()

    rows = [{"구분": r["법인"], "경상이익차이": r["차이-경상이익"], "내용": ""} for _, r in top.iterrows()]
    rows.append({"구분": "기타", "경상이익차이": others_sum, "내용": ""})
    rows.append({"구분": "합산", "경상이익차이": total_sum, "내용": ""})
    return pd.DataFrame(rows)


def style_top_table(df: pd.DataFrame, decimals=1) -> "pd.io.formats.style.Styler":
    def row_style(row):
        if row["구분"] == "합산":
            return ["font-weight: bold; background-color: #D9E1F2" for _ in row]
        if row["구분"] == "기타":
            return ["font-style: italic; color: #666666" for _ in row]
        return ["" for _ in row]

    def color_negative(v):
        if isinstance(v, (int, float)) and v < 0:
            return f"color: {RED}"
        return ""

    # 내용 칸을 가장 넓게 - 나중에 코멘트를 텍스트로 채워 넣을 자리
    col_widths = {"구분": "15%", "경상이익차이": "15%", "내용": "70%"}
    table_styles = [{"selector": "", "props": "table-layout: fixed; width: 100%;"}]
    for col, width in col_widths.items():
        idx = df.columns.get_loc(col)
        table_styles.append({"selector": f"th.col_heading.col{idx}", "props": f"width: {width};"})
        table_styles.append({"selector": f"td.col{idx}", "props": f"width: {width};"})

    return (
        df.style.format({"경상이익차이": f"{{:,.{decimals}f}}"})
        .apply(row_style, axis=1)
        .map(color_negative, subset=["경상이익차이"])
        .set_table_styles(table_styles, overwrite=False)
    )


def add_region_bands(fig: go.Figure) -> None:
    """법인별 그래프 x축 뒤에 권역별(중국/인도/유럽/미주) 배경 밴드를 표 배경색과 동일하게 표시."""
    for region, members in REGION_MEMBERS.items():
        first_idx = ROW_ORDER.index(members[0])
        last_idx = ROW_ORDER.index(members[-1])
        fig.add_vrect(
            x0=first_idx - 0.5,
            x1=last_idx + 0.5,
            fillcolor=REGION_COLORS[region],
            opacity=0.6,
            line_width=0,
            layer="below",
            annotation_text=region,
            annotation_position="top",
            annotation_font=dict(size=12, color="#555555"),
        )


def page1_chart(page1: pd.DataFrame) -> go.Figure:
    d = page1.set_index("법인").loc[ROW_ORDER].reset_index()
    colors = [NAVY if r == "합계" else (GREEN if v >= 0 else RED) for r, v in zip(d["법인"], d["차이-경상이익"])]
    fig = go.Figure(
        go.Bar(x=d["법인"], y=d["차이-경상이익"], marker_color=colors, text=d["차이-경상이익"].round(1), textposition="outside")
    )
    fig.update_layout(
        title="법인별 경상이익차이 (실적-계획)",
        yaxis_title="백만$",
        xaxis_title="",
        height=420,
        margin=dict(t=60, b=20),
        showlegend=False,
    )
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=ROW_ORDER)
    add_region_bands(fig)
    fig.add_hline(y=0, line_color="#999999", line_width=1)
    return fig


def page2_chart(page2: pd.DataFrame) -> go.Figure:
    d = page2.set_index("법인").loc[ROW_ORDER].reset_index()
    fig = go.Figure()
    for comp, color in COMPONENT_COLORS.items():
        fig.add_trace(go.Bar(name=comp, x=d["법인"], y=d[comp], marker_color=color))
    fig.add_trace(
        go.Scatter(
            name="경상이익차이(검증)",
            x=d["법인"],
            y=d["경상이익차이"],
            mode="markers+text",
            marker=dict(symbol="diamond", size=11, color="black"),
            text=d["경상이익차이"].round(1),
            textposition="bottom center",
            textfont=dict(size=11, color="black"),
        )
    )
    fig.update_layout(
        barmode="relative",
        title="법인별 계획비 주요 차이내역 (물량/롤마진/제경비/영업외 → 경상이익차이)",
        yaxis_title="백만$",
        xaxis_title="",
        height=460,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=80, b=20),
    )
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=ROW_ORDER)
    add_region_bands(fig)
    fig.add_hline(y=0, line_color="#999999", line_width=1)
    return fig


st.title("해외법인 손익 리포트")

options = _load_period_options()
default_idx = len(options) - 1  # 가장 최근 조합을 기본값으로
selected = st.selectbox("기준월 선택", options, index=default_idx)

page1, page2 = compute_report(selected)

tab1, tab2 = st.tabs(["Page 1 · 계획/실적 요약", "Page 2 · 경상이익 차이내역"])

with tab1:
    with st.expander("표 보기", expanded=True):
        st.table(
            style_table(page1, group_dividers=["계획-판매량", "실적-판매량", "차이-경상이익"]),
            hide_index=True,
        )
    with st.expander("그래프 보기", expanded=True):
        st.plotly_chart(page1_chart(page1), width="stretch")
    with st.expander("주요 차이법인 보기", expanded=True):
        st.table(style_top_table(top_diff_table(page1)), hide_index=True)

with tab2:
    with st.expander("표 보기", expanded=True):
        st.table(style_table(page2, group_dividers=["경상이익차이", "물량차이"]), hide_index=True)
    with st.expander("그래프 보기", expanded=True):
        st.plotly_chart(page2_chart(page2), width="stretch")
    with st.expander("주요 차이법인 보기", expanded=True):
        st.table(style_top_table(top_diff_table(page1)), hide_index=True)
