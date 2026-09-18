"""PPT 보고서용 슬라이드 데이터를 만든다.

대시보드 화면과 같은 계산 로직(charts.py / data_loader.py)을 그대로 쓰되, 두 가지를
고정한다 - PPT는 화면과 달리 재현 가능해야 하고, 어떤 뷰어의 화면에서 열어도 똑같이
읽혀야 하기 때문:

- **항상 라이트 팔레트.** 뷰어의 다크모드와 무관하게, 프린트/투사 환경에서 흔한
  흰 배경 슬라이드에 맞춘다.
- **항상 전체 기간(raw).** 대시보드의 사이드바 필터(기간·지역·고속사 등)는 반영하지
  않는다 - "보고서 생성" 버튼을 누른 시점의 우연한 필터 상태에 따라 내용이 달라지면
  안 되기 때문이다.

이 모듈은 순수 계산만 하고(파일 I/O나 st.* 호출 없음), 결과를 `Slide` 리스트로 반환한다.
실제 PPTX 파일로 굳히는 것은 `ppt_export.py`가 담당한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd
import plotly.graph_objects as go

import charts as ch
import data_loader as dl
from viz_theme import Palette

PAL = Palette("light")  # 보고서는 항상 라이트 고정


@dataclass
class Slide:
    title: str
    figs: list[go.Figure]           # 1장이면 [fig], 2장이면 나란히 배치
    bullets: list[str]
    fig_widths: list[float] = field(default_factory=list)  # 상대 비율, 비우면 균등분할


def _pct(n: int, total: int) -> str:
    return f"{n / total * 100:.0f}%" if total else "0%"


def build_yoy_slide(raw: pd.DataFrame) -> Slide | None:
    """연도별 월간 추이 비교 (추세 탭과 동일한 로직 - 2개월 미만인 연도는 제외)."""
    mon = raw.groupby(["연도", "월"], as_index=False).size().rename(columns={"size": "건수"})
    months = mon.groupby("연도")["월"].nunique()
    drawn = sorted(int(y) for y in months[months >= 2].index)
    if not drawn:
        return None

    fig = ch.multiline(mon[mon["연도"].isin(drawn)], "월", "건수", "연도", PAL,
                       height=760, xtitle="월", order=drawn)
    fig.update_xaxes(tickmode="array", tickvals=list(range(1, 13)),
                     ticktext=[f"{m}월" for m in range(1, 13)], showgrid=False)

    totals = mon[mon["연도"].isin(drawn)].groupby("연도")["건수"].sum()
    latest, prev = drawn[-1], (drawn[-2] if len(drawn) >= 2 else None)
    bullets = [f"집계 연도: {', '.join(str(y) + '년' for y in drawn)}",
              f"{latest}년 누계 {int(totals[latest]):,}건" +
              (f" (부분 기간, {int(months[latest])}개월치)" if months[latest] < 12 else "")]
    if prev is not None:
        # 같은 개월 수만큼만 비교해야 공정하다(부분 기간 연도 대비 왜곡 방지)
        n_month = int(months[latest])
        common_months = mon[(mon["연도"].isin([prev, latest])) & (mon["월"] <= n_month)]
        t = common_months.groupby("연도")["건수"].sum()
        if prev in t.index and latest in t.index and t[prev] > 0:
            diff = t[latest] - t[prev]
            bullets.append(f"{prev}년 동기(1~{n_month}월) 대비 {diff:+,}건 "
                           f"({diff / t[prev] * 100:+.0f}%)")
    return Slide("연도별 월간 추이 비교", [fig], bullets)


def build_monthly_count_slide(raw: pd.DataFrame) -> Slide:
    """월별 장애 건수 (전체 기간) - 추세 탭과 동일."""
    m = (raw.groupby("연월라벨", as_index=False).size()
           .rename(columns={"size": "건수"}).sort_values("연월라벨"))
    m["3개월 이동평균"] = m["건수"].rolling(3).mean().round(1)
    fig = ch.vbar_line(m["연월라벨"], m["건수"], m["3개월 이동평균"], PAL, height=760,
                       bar_name="월간 건수", line_name="3개월 이동평균")
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=15))

    last3 = m.tail(3)["건수"].sum()
    prev3 = m.iloc[-6:-3]["건수"].sum() if len(m) >= 6 else None
    bullets = [f"집계 기간: {m['연월라벨'].iloc[0]} ~ {m['연월라벨'].iloc[-1]} ({len(m)}개월)",
              f"최근 3개월 합계 {int(last3):,}건"]
    if prev3:
        diff = last3 - prev3
        bullets.append(f"직전 3개월 대비 {diff:+,.0f}건 ({diff / prev3 * 100:+.0f}%)")
    return Slide("월별 장애 건수 (전체 기간)", [fig], bullets)


def build_focus_monthly_mix_slide(raw: pd.DataFrame, focus_faults: list[str]) -> Slide | None:
    """관심 장애유형 월별 구성 (전체 기간) - 관심 장애유형 탭과 동일."""
    foc = raw[raw["장애 대분류"].isin(focus_faults)]
    if foc.empty:
        return None
    months = sorted(raw["연월라벨"].unique())
    mon_piv = (foc.groupby(["연월라벨", "장애 대분류"]).size().unstack(fill_value=0)
                  .reindex(index=months, columns=focus_faults, fill_value=0))
    fig = ch.stacked_bar(mon_piv, PAL, height=760)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=15))

    total = int(mon_piv.values.sum())
    share = _pct(total, len(raw))
    top = mon_piv.sum().sort_values(ascending=False)
    bullets = [f"이 {len(focus_faults)}종 합계 {total:,}건 (전체의 {share})",
              f"최다 유형: {top.index[0]} {int(top.iloc[0]):,}건"]
    return Slide("관심 장애유형 - 월별 구성 (전체 기간)", [fig], bullets)


def build_field_trend_figure(raw: pd.DataFrame, focus_faults: list[str], pal: Palette, *,
                             height: int = 360, tick_size: int = 14,
                             start_year: int = 2024) -> tuple[go.Figure, pd.DataFrame]:
    """필드 장애 월별 추이(예: LTE·BMS·승차연결지연) + 개선활동 두 건(수리포인트 시작·
    AFC-승차단말 펌웨어 배포) 시점 표시. 대시보드(개선효과 탭)와 PPT 보고서가 정확히
    같은 차트를 그리도록 이 함수 하나로 모았다."""
    f_trend = raw[raw["연도"] >= start_year]
    months = sorted(f_trend["연월라벨"].unique())
    mon = (f_trend[f_trend["장애 대분류"].isin(focus_faults)]
           .groupby(["연월라벨", "장애 대분류"]).size().unstack(fill_value=0)
           .reindex(index=months, columns=focus_faults, fill_value=0))
    long_df = mon.reset_index().melt(id_vars="연월라벨", var_name="장애 대분류", value_name="건수")
    fig = ch.multiline(long_df, "연월라벨", "건수", "장애 대분류", pal, height=height,
                       order=focus_faults)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=tick_size), type="category",
                     categoryorder="array", categoryarray=months)

    for date, y_anchor, label in [
        (dl.REPAIR_POINT_START, 1.0,
         f"수리포인트 시작({dl.REPAIR_POINT_START.month}/{dl.REPAIR_POINT_START.day})"),
        (dl.FIRMWARE_DEPLOY_DATE, 0.9,
         f"AFC-승차단말 펌웨어({dl.FIRMWARE_DEPLOY_DATE.month}/{dl.FIRMWARE_DEPLOY_DATE.day})"),
    ]:
        x_label = f"{date.year}-{date.month:02d}"
        if x_label in months:
            fig.add_vline(x=x_label, line_dash="dot", line_width=1.5, line_color=pal.c["muted"])
            fig.add_annotation(x=x_label, y=y_anchor, yref="paper", showarrow=False,
                               text=label, font=dict(size=tick_size, color=pal.c["muted"]),
                               xanchor="left", yanchor="top")
    return fig, mon.reset_index()


def build_field_trend_slide(raw: pd.DataFrame, focus_faults: list[str]) -> Slide | None:
    """필드 장애 월별 추이(2024년~) - 개선효과 탭의 같은 이름 차트와 동일하다."""
    if raw[raw["연도"] >= 2024].empty:
        return None
    fig, table = build_field_trend_figure(raw, focus_faults, PAL, height=760, tick_size=15)

    last = table.iloc[-1]
    last_total = int(last[focus_faults].sum())
    bullets = [
        f"집계 기간: {table['연월라벨'].iloc[0]} ~ {table['연월라벨'].iloc[-1]}",
        f"{table['연월라벨'].iloc[-1]} 합계 {last_total:,}건",
        f"점선: 수리포인트 시작 {dl.REPAIR_POINT_START:%Y-%m-%d} · "
        f"AFC-승차단말 펌웨어 배포 {dl.FIRMWARE_DEPLOY_DATE:%Y-%m-%d}",
    ]

    # AFC-승차단말 펌웨어는 전 차량 일괄 배포라 차량별 대조군이 없어, 배포 전/후를
    # 월 환산 발생률로 비교한다 - 개선효과 탭의 펌웨어 효과 KPI와 같은 계산.
    fw_stats = dl.before_after_monthly_rate(raw, dl.FIRMWARE_TARGET_FAULTS,
                                            dl.FIRMWARE_DEPLOY_DATE)
    if fw_stats["days_since"] >= 0:
        b, a = fw_stats["before_rate"], fw_stats["after_rate"]
        diff_pct = f"{(a - b) / b * 100:+.0f}%" if b else "N/A"
        bullets.append(
            f"AFC 펌웨어(승차연결지연·통신불량 기준): 배포전 월평균 {b:.1f}건/월 → "
            f"배포후 월평균 {a:.1f}건/월 ({diff_pct}, {fw_stats['days_since']}일 관찰)"
        )

    bullets.append("계절 요인 등 다른 변수도 함께 작용할 수 있어, 이 그래프만으로 인과를 "
                   "단정하지 않습니다.")
    return Slide("필드 장애 추이 (관심 장애유형)", [fig], bullets)


def build_intervention_slide(raw: pd.DataFrame, iv: pd.DataFrame,
                             fleet_size_est: int | None = None) -> Slide | None:
    """외장 LTE모뎀 전환 · B/D(AFC/BMS) 보드교체 현황 - 개선효과 탭과 동일한 2개 차트."""
    n_new = int(iv["외장모뎀_신규전환"].sum())
    n_bd = int(iv["BD보드_교체"].sum())
    n_bd_only = int(iv["BD보드_단독"].sum())
    if n_new == 0 and n_bd == 0:
        return None

    m_new = (iv[iv["외장모뎀_신규전환"]]
             .groupby(iv.loc[iv["외장모뎀_신규전환"], "일자"].dt.to_period("M").astype(str))
             .size())
    if len(m_new):
        m_new = m_new.reindex(
            pd.period_range(m_new.index.min(), m_new.index.max(), freq="M").astype(str),
            fill_value=0)
    cum = m_new.cumsum()
    fig1 = ch.vbar_line(m_new.index, m_new.values, cum.values, PAL, height=680,
                        bar_name="월별 신규전환", line_name="누적 전환대수")
    fig1.update_xaxes(tickangle=-45, tickfont=dict(size=14))

    bd = iv[iv["BD보드_교체"]].copy()
    bd["구분"] = bd["BD보드_단독"].map({True: "B/D 단독", False: "외장모뎀 동반"})
    piv = (bd.groupby([bd["일자"].dt.to_period("M").astype(str), "구분"]).size()
             .unstack(fill_value=0).reindex(columns=["B/D 단독", "외장모뎀 동반"], fill_value=0))
    fig2 = ch.stacked_bar(piv, PAL, height=680, other="__none__")
    fig2.update_xaxes(tickangle=-45, tickfont=dict(size=14))

    n_new_veh = iv.loc[iv["외장모뎀_신규전환"], "차량번호"].nunique()
    n_bd_veh = iv.loc[iv["BD보드_교체"], "차량번호"].nunique()
    bullets = [f"외장 LTE모뎀 신규전환 {n_new:,}건 ({n_new_veh:,}대)",
              f"B/D(AFC/BMS) 보드교체 {n_bd:,}건 ({n_bd_veh:,}대) - "
              f"단독 {n_bd_only:,}건 · 외장모뎀 동반 {n_bd - n_bd_only:,}건"]
    if fleet_size_est:
        bullets.append(f"추정 운영 차량 약 {fleet_size_est:,}대 대비 - 전환 약 "
                       f"{_pct(n_new_veh, fleet_size_est)}, 보드교체 약 "
                       f"{_pct(n_bd_veh, fleet_size_est)} 진행")
    return Slide("외장 LTE모뎀 전환 · B/D(AFC/BMS) 보드교체 현황",
                [fig1, fig2], bullets, fig_widths=[1, 1])


def build_before_after_slide(raw: pd.DataFrame, iv: pd.DataFrame,
                             focus_faults: list[str]) -> Slide | None:
    """조치 전/후 90일 관심 장애 발생률 비교 - 개선효과 탭과 동일."""
    studies = {
        "외장모뎀 신규전환": dl.vehicle_event_study(raw, iv["외장모뎀_신규전환"],
                                                  focus_faults=focus_faults),
        "B/D교체(단독)": dl.vehicle_event_study(raw, iv["BD보드_단독"],
                                              focus_faults=focus_faults),
        "B/D교체(전체)": dl.vehicle_event_study(raw, iv["BD보드_교체"],
                                              focus_faults=focus_faults),
    }
    cats, before_rate, after_rate, ns = [], [], [], []
    for name, r in studies.items():
        if r.empty:
            continue
        cats.append(name)
        before_rate.append(r["전"].mean())
        after_rate.append(r["후"].mean())
        ns.append(len(r))
    if not cats:
        return None

    fig = ch.paired_bar(cats, before_rate, after_rate, PAL, height=760,
                        suffix="건/대", before_name="전 90일", after_name="후 90일",
                        value_fmt="{:.2f}")
    fig.update_xaxes(tickfont=dict(size=17))

    bullets = [f"막대는 차량 1대당 평균 발생건수(관심 장애 {len(focus_faults)}종 기준)",
              "관찰 차량수(n): " + " · ".join(f"{c}={n}대" for c, n in zip(cats, ns))]
    for c, b, a in zip(cats, before_rate, after_rate):
        arrow = "개선" if a < b else ("악화" if a > b else "동일")
        bullets.append(f"{c}: {b:.2f} → {a:.2f} ({arrow})")
    bullets.append("정식 대조군 비교가 아닌 참고용 신호입니다.")
    return Slide("조치 전/후 관심 장애 발생률 비교", [fig], bullets)


def build_report_slides(raw: pd.DataFrame, focus_faults: list[str],
                        fleet_size_est: int | None = None) -> list[Slide]:
    """요청받은 5종 슬라이드(월별 장애 건수 · 관심유형 월별구성 · 외장모뎀/B/D 현황 ·
    전후비교 · 필드 장애 추이)를 순서대로 만든다. 데이터가 없어 만들 수 없는 슬라이드는
    조용히 건너뛴다(빈 슬라이드를 넣지 않는다).

    연도별 월간 추이 비교는 뺐다(사용자 요청) - 대신 개선활동 두 건(수리포인트·펌웨어)
    시점이 함께 표시되는 필드 장애 추이를 넣는다.
    """
    iv = dl.extract_intervention_events(raw)
    builders = [
        build_monthly_count_slide(raw),
        build_focus_monthly_mix_slide(raw, focus_faults),
        build_intervention_slide(raw, iv, fleet_size_est),
        build_before_after_slide(raw, iv, focus_faults[:3]),
        build_field_trend_slide(raw, focus_faults),
    ]
    return [s for s in builders if s is not None]
