"""대시보드 차트 빌더.

규칙(dataviz)
- 형태를 먼저 고르고, 색은 마지막에 역할로 배정한다.
- 카테고리 색은 고정 슬롯 순서로만 쓰고 순위에 따라 다시 칠하지 않는다.
- y축은 하나만 둔다(이중 축 금지).
- 시리즈가 2개 이상이면 범례를 항상 두고, 모든 차트에는 '표로 보기'를 함께 제공한다.
"""

from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from viz_theme import Palette, style


def _bar_edges(fig, pal: Palette, width: int = 1):
    """막대 사이 간격 - 테두리가 아니라 배경색 틈으로 분리한다."""
    fig.update_traces(marker_line_color=pal.c["surface"], marker_line_width=width,
                      selector=dict(type="bar"))
    try:
        fig.update_traces(marker_cornerradius=4, selector=dict(type="bar"))
    except Exception:
        pass


def hbar(data: pd.DataFrame, cat: str, val: str, pal: Palette, *,
         height: int | None = None, color: str | None = None, suffix: str = "건",
         text_fmt: str = "{:,.0f}", hover_fmt: str = "%{x:,.0f}") -> go.Figure:
    """가로 막대 - 이름이 긴 명목형 범주의 크기 비교용. 한 시리즈 = 한 색."""
    d = data.sort_values(val, ascending=True)
    fig = go.Figure(go.Bar(
        x=d[val], y=d[cat].astype(str), orientation="h",
        marker_color=color or pal.slot(0),
        text=[text_fmt.format(v) for v in d[val]],
        textposition="outside", cliponaxis=False,
        textfont=dict(color=pal.c["ink2"], size=15, family="system-ui"),
        hovertemplate=f"<b>%{{y}}</b><br>{hover_fmt}{suffix}<extra></extra>",
    ))
    _bar_edges(fig, pal)
    h = height or max(220, 30 * len(d) + 60)
    style(fig, pal, height=h, showlegend=False, bottom=32)
    fig.update_xaxes(showgrid=True, showticklabels=False, range=[0, d[val].max() * 1.16])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=pal.c["ink2"], size=16))
    return fig


def vbar_line(x, bars, line, pal: Palette, *, bar_name: str, line_name: str,
              height: int = 300, xtitle: str = "", suffix: str = "건") -> go.Figure:
    """막대 + 이동평균선. 같은 단위·같은 축이므로 이중 축이 아니다."""
    fig = go.Figure()
    fig.add_bar(x=x, y=bars, name=bar_name, marker_color=pal.slot(0),
                hovertemplate=f"%{{x}}<br>{bar_name} %{{y:,.0f}}{suffix}<extra></extra>")
    fig.add_scatter(x=x, y=line, name=line_name, mode="lines",
                    line=dict(color=pal.slot(1), width=2, shape="spline", smoothing=0.5),
                    hovertemplate=f"%{{x}}<br>{line_name} %{{y:,.1f}}{suffix}<extra></extra>")
    _bar_edges(fig, pal)
    style(fig, pal, height=height, xtitle=xtitle, margin_l=36, bottom=60)
    fig.update_layout(hovermode="x unified")
    # '2024-01' 같은 라벨을 plotly가 날짜로 해석해 영문 눈금으로 바꾸지 않게 한다
    fig.update_xaxes(showgrid=False, type="category")
    return fig


def vbar(x, y, pal: Palette, *, height: int = 300, xtitle: str = "", suffix: str = "건",
        text_fmt: str = "{:,.0f}", value_labels: list[str] | None = None) -> go.Figure:
    """단일 시리즈 세로 막대 - 값 라벨을 막대 위에 직접 찍는다(범례가 필요 없는 경우).

    `value_labels`를 주면 막대 위 텍스트를 그 값으로 대신한다(예: '관찰불가' 표시).
    """
    labels = value_labels if value_labels is not None else [text_fmt.format(v) for v in y]
    fig = go.Figure(go.Bar(
        x=list(x), y=list(y), marker_color=pal.slot(0),
        text=labels, textposition="outside", cliponaxis=False,
        textfont=dict(color=pal.c["ink2"], size=15, family="system-ui"),
        hovertemplate=f"%{{x}}<br>%{{y:,.1f}}{suffix}<extra></extra>",
    ))
    _bar_edges(fig, pal)
    style(fig, pal, height=height, showlegend=False, xtitle=xtitle, margin_l=36, bottom=56)
    fig.update_xaxes(showgrid=False, type="category")
    ymax = max([v for v in y if v is not None] or [0])
    fig.update_yaxes(range=[0, ymax * 1.25 if ymax else 1])
    return fig


def paired_bar(categories, before, after, pal: Palette, *, height: int = 300,
              xtitle: str = "", suffix: str = "건", before_name: str = "전",
              after_name: str = "후", value_fmt: str = "{:,.0f}") -> go.Figure:
    """카테고리별 전/후 막대 - '전'/'후'가 곧 시리즈이므로 슬롯을 고정 배정한다
    (전=슬롯0, 후=슬롯1 - 두 차트에 걸쳐 같은 뜻이면 항상 같은 색)."""
    fig = go.Figure()
    fig.add_bar(x=list(categories), y=list(before), name=before_name, marker_color=pal.slot(0),
                text=[value_fmt.format(v) for v in before], textposition="outside",
                cliponaxis=False, textfont=dict(color=pal.c["ink2"], size=15),
                hovertemplate=f"%{{x}} · {before_name}<br>%{{y:,.2f}}{suffix}<extra></extra>")
    fig.add_bar(x=list(categories), y=list(after), name=after_name, marker_color=pal.slot(1),
                text=[value_fmt.format(v) for v in after], textposition="outside",
                cliponaxis=False, textfont=dict(color=pal.c["ink2"], size=15),
                hovertemplate=f"%{{x}} · {after_name}<br>%{{y:,.2f}}{suffix}<extra></extra>")
    _bar_edges(fig, pal)
    style(fig, pal, height=height, xtitle=xtitle, margin_l=36, bottom=56)
    fig.update_layout(barmode="group")
    fig.update_xaxes(showgrid=False, type="category")
    ymax = max(list(before) + list(after) or [0])
    fig.update_yaxes(range=[0, ymax * 1.3 if ymax else 1])
    return fig


def multiline(df: pd.DataFrame, x: str, y: str, series: str, pal: Palette, *,
              height: int = 320, xtitle: str = "", suffix: str = "건",
              order: list | None = None) -> go.Figure:
    """시리즈별 라인. 색은 시리즈 이름에 고정 배정하여 필터에도 유지된다."""
    names = order or sorted(df[series].dropna().unique().tolist())
    cmap = {n: pal.slot(i) for i, n in enumerate(names)}
    fig = go.Figure()
    for n in names:
        d = df[df[series] == n]
        fig.add_scatter(x=d[x], y=d[y], name=str(n), mode="lines+markers",
                        line=dict(color=cmap[n], width=2),
                        marker=dict(size=8, line=dict(width=2, color=pal.c["surface"])),
                        hovertemplate=f"%{{x}} · {n}<br>%{{y:,.0f}}{suffix}<extra></extra>")
    style(fig, pal, height=height, xtitle=xtitle, margin_l=36, bottom=56)
    fig.update_layout(hovermode="x unified")
    return fig


def stacked_bar(pivot: pd.DataFrame, pal: Palette, *, height: int = 340,
                xtitle: str = "", suffix: str = "건", other: str = "기타") -> go.Figure:
    """누적 막대 - 구성비 변화용. 세그먼트 사이는 배경색 2px 틈으로 분리.

    꼬리를 묶은 '기타'는 중립 회색으로 둔다 - 색 슬롯은 실제 범주에만 쓴다.
    """
    fig = go.Figure()
    slot = 0
    for col in pivot.columns:
        if str(col) == other:
            color = pal.c["axis"]
        else:
            color = pal.slot(slot)
            slot += 1
        fig.add_bar(x=pivot.index.astype(str), y=pivot[col], name=str(col),
                    marker_color=color,
                    marker_line=dict(color=pal.c["surface"], width=2),
                    hovertemplate=f"%{{x}} · {col}<br>%{{y:,.0f}}{suffix}<extra></extra>")
    style(fig, pal, height=height, xtitle=xtitle, margin_l=36, bottom=60)
    fig.update_layout(barmode="stack", hovermode="x unified", legend_traceorder="normal")
    fig.update_xaxes(showgrid=False, type="category")
    return fig


def heatmap(mat: pd.DataFrame, pal: Palette, *, height: int | None = None,
            xtitle: str = "", ytitle: str = "", suffix: str = "건") -> go.Figure:
    """연속형 크기 - 파랑 단일 색상 램프 + 스케일 범례. 값은 표로도 제공한다."""
    fig = go.Figure(go.Heatmap(
        z=mat.values, x=mat.columns.astype(str), y=mat.index.astype(str),
        colorscale=[[i / (len(pal.seq) - 1), c] for i, c in enumerate(pal.seq)],
        xgap=2, ygap=2,
        colorbar=dict(title=dict(text=suffix, side="top", font=dict(size=15, color=pal.c["muted"])),
                      thickness=10, outlinewidth=0, len=0.85,
                      tickfont=dict(size=14, color=pal.c["muted"])),
        hovertemplate=f"%{{y}} · %{{x}}<br>%{{z:,.0f}}{suffix}<extra></extra>",
    ))
    h = height or max(260, 28 * len(mat.index) + 120)
    style(fig, pal, height=h, showlegend=False, xtitle=xtitle, ytitle=ytitle,
          margin_l=8, bottom=90)
    fig.update_xaxes(showgrid=False, tickangle=-35, tickfont=dict(size=15, color=pal.c["ink2"]))
    fig.update_yaxes(showgrid=False, tickfont=dict(size=15, color=pal.c["ink2"]))
    return fig


def block(title: str, fig: go.Figure, table: pd.DataFrame, key: str, *,
          caption: str | None = None):
    """차트 + 캡션 + 표 보기(값 접근성) + CSV 내려받기를 한 덩어리로 렌더."""
    st.markdown(f"**{title}**")
    if caption:
        st.caption(caption)
    st.plotly_chart(fig, key=f"c_{key}", theme=None,
                    config={"displayModeBar": False, "responsive": True})
    with st.expander("표로 보기 · 내려받기"):
        st.dataframe(table, hide_index=True, height="auto")
        st.download_button("CSV 내려받기", table.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"{key}.csv", mime="text/csv", key=f"d_{key}")
