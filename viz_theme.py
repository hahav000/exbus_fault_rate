"""차트 팔레트 · Plotly 레이아웃 템플릿.

`dataviz` 레퍼런스 팔레트를 그대로 사용한다(검증된 기본값).
- 카테고리형: 고정 순서로만 배정하며 순위에 따라 재배색하지 않는다.
- 연속형(히트맵): 파랑 단일 색상 명도 램프 + 스케일 범례.
- 라이트/다크 각각 별도 스텝을 사용한다(자동 반전 아님).
"""

from __future__ import annotations

import streamlit as st

FONT = "Pretendard, system-ui, -apple-system, 'Malgun Gothic', 'Segoe UI', sans-serif"

# 카테고리형 8슬롯 (고정 순서)
CATEGORICAL = {
    "light": ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
              "#e87ba4", "#008300", "#4a3aa7", "#e34948"],
    "dark": ["#3987e5", "#d95926", "#199e70", "#c98500",
             "#d55181", "#008300", "#9085e9", "#e66767"],
}

# 연속형: 파랑 단일 색상 램프 (밝음 → 어두움)
SEQUENTIAL = {
    "light": ["#f2f7fe", "#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#104281"],
    "dark": ["#11223a", "#104281", "#184f95", "#256abf", "#3987e5", "#6da7ec", "#9ec5f4"],
}

# 상태색 (고정 - 시리즈 색으로 재사용 금지)
STATUS = {"good": "#0ca30c", "warning": "#fab219", "serious": "#ec835a", "critical": "#d03b3b"}

CHROME = {
    "light": {
        "surface": "#fcfcfb",
        "page": "#f9f9f7",
        "ink": "#0b0b0b",
        "ink2": "#52514e",
        "muted": "#898781",
        "grid": "#e1e0d9",
        "axis": "#c3c2b7",
        "up_good": "#006300",
        "border": "rgba(11,11,11,0.10)",
    },
    "dark": {
        "surface": "#1a1a19",
        "page": "#0d0d0d",
        "ink": "#ffffff",
        "ink2": "#c3c2b7",
        "muted": "#898781",
        "grid": "#2c2c2a",
        "axis": "#383835",
        "up_good": "#0ca30c",
        "border": "rgba(255,255,255,0.10)",
    },
}


def theme_mode() -> str:
    """뷰어가 보고 있는 테마(light/dark)를 반환."""
    try:
        t = getattr(st.context, "theme", None)
        if t is not None and getattr(t, "type", None) in ("light", "dark"):
            return t.type
    except Exception:
        pass
    base = st.get_option("theme.base")
    return "dark" if base == "dark" else "light"


class Palette:
    """현재 테마에 맞는 색 슬롯 모음."""

    def __init__(self, mode: str):
        self.mode = mode
        self.series = CATEGORICAL[mode]
        self.seq = SEQUENTIAL[mode]
        self.c = CHROME[mode]

    def slot(self, i: int) -> str:
        """카테고리 슬롯 색. 8개를 넘는 시리즈는 만들지 않는다(그 전에 '기타'로 묶을 것)."""
        if i >= len(self.series):
            raise ValueError("카테고리 색은 8슬롯까지만 사용한다 - 그 이상은 '기타'로 묶거나 분할할 것")
        return self.series[i]


def style(fig, pal: Palette, *, height: int = 320, showlegend: bool | None = None,
          xtitle: str = "", ytitle: str = "", margin_l: int = 8, bottom: int = 44):
    """모든 차트에 공통 크롬을 적용한다 - 얇은 헤어라인 격자, 낮은 채도의 축."""
    c = pal.c
    fig.update_layout(
        template="plotly_white" if pal.mode == "light" else "plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, size=16, color=c["ink2"]),
        height=height,
        margin=dict(l=margin_l, r=12, t=8, b=bottom),
        hoverlabel=dict(font=dict(family=FONT, size=16), bgcolor=c["surface"],
                        bordercolor=c["axis"], font_color=c["ink"]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    font=dict(color=c["ink2"], size=16), bgcolor="rgba(0,0,0,0)",
                    title_text=""),
        bargap=0.28,
    )
    if showlegend is not None:
        fig.update_layout(showlegend=showlegend)
    axis = dict(
        showgrid=True, gridcolor=c["grid"], gridwidth=1, griddash="solid",
        zeroline=False, linecolor=c["axis"], linewidth=1,
        tickfont=dict(color=c["muted"], size=15), title_font=dict(color=c["ink2"], size=16),
        automargin=True,
    )
    fig.update_xaxes(**axis, title_text=xtitle)
    fig.update_yaxes(**axis, title_text=ytitle)
    return fig
