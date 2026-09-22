"""고속버스 단말기 장애실적 대시보드.

`고속 시외 YYYY년.xlsm` 파일들의 `고속Rawdata` 시트를 읽어 분석한다.
파일이 갱신되면(매주) 자동으로 다시 읽는다.

실행:  streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

import charts as ch
import data_loader as dl
import ppt_export
import report_builder
from viz_theme import Palette, theme_mode

BASE_DIR = Path(__file__).resolve().parent

# 사용자가 지목한 관심 장애유형 - 전부 통신/연결 계열이다.
# 이 순서가 곧 색 슬롯 순서(고정) - 어떤 필터를 걸어도 같은 유형은 같은 색을 유지한다.
FOCUS_FAULTS = ["LTE모뎀망연결지연", "BMS접속지연", "승차연결지연", "승차단말기통신불량"]

# 사용자 제공 추정치 - 고속사별로 매달 대수가 조금씩 바뀌어 정확한 값이 아니라 대략치.
# 커버리지(전환률 등)를 어림잡을 때만 참고용으로 쓰고, 정밀한 분모로는 쓰지 않는다.
FLEET_SIZE_EST = 2000

# AFC-승차단말간 통신보완코드가 적용된 펌웨어를 전 차량에 일괄 배포한 날짜(dl 모듈이
# 진짜 출처 - PPT 보고서의 필드 장애 추이 슬라이드도 같은 값을 본다).
# 개별 차량마다 시행일이 다른 외장모뎀 전환/B/D교체와 달리 - 이건 한 날짜에 전체가 바뀌는
# 사건이라 차량별 대조군이 없다. 그래서 차량별 전/후 비교가 아니라 전체(모집단) 수준의
# 시계열 전/후 비교로 본다.
FIRMWARE_DATE = dl.FIRMWARE_DEPLOY_DATE
FIRMWARE_TARGET_FAULTS = dl.FIRMWARE_TARGET_FAULTS

st.set_page_config(page_title="고속버스 장애실적 대시보드", page_icon="🚌",
                   layout="wide", initial_sidebar_state="expanded")

PAL = Palette(theme_mode())

st.markdown(f"""
<style>
  .block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1500px; }}
  h1, h2, h3 {{ letter-spacing: -0.01em; }}
  [data-testid="stMetric"] {{
      background: {PAL.c['surface']}; border: 1px solid {PAL.c['border']};
      border-radius: 10px; padding: 14px 16px;
  }}
  [data-testid="stMetricLabel"] p {{ color: {PAL.c['muted']}; font-size: 0.80rem; }}
  [data-testid="stMetricValue"] {{ font-size: 1.6rem; }}
  /* 값이 말줄임으로 잘리지 않게 한다 - 숫자가 잘리면 타일이 쓸모없다 */
  [data-testid="stMetricValue"] div {{ overflow: visible; text-overflow: clip; white-space: normal; }}
  [data-testid="stMetricDelta"] div {{ overflow: visible; text-overflow: clip; white-space: normal; }}
  hr {{ border-color: {PAL.c['border']}; }}
</style>
""", unsafe_allow_html=True)


# ── 데이터 적재 (항상 이 폴더에 둔 파일을 읽는다) ───────────────────────
files = dl.discover_files(BASE_DIR)
if not files:
    st.error(f"`{BASE_DIR}` 에서 `.xlsm` 파일을 찾지 못했습니다. "
             "`고속 시외 2025년.xlsm` 같은 파일을 이 폴더에 두고 다시 실행하세요.")
    st.stop()
try:
    raw = dl.load_data(tuple(str(p) for p in files), dl.fingerprint(files))
except Exception as exc:  # 엑셀이 열려 있거나 시트 구성이 바뀐 경우
    st.error(f"엑셀을 읽지 못했습니다: {exc}")
    st.info("파일이 엑셀에서 열려 있으면 닫은 뒤 새로고침하세요. "
            f"시트 이름은 `{dl.SHEET}`, 헤더는 3행이어야 합니다.")
    st.stop()
fault_labels = []
for p in files:
    ts = pd.Timestamp(p.stat().st_mtime, unit="s", tz="UTC").tz_convert("Asia/Seoul")
    fault_labels.append(f"{p.name} · 갱신 {ts:%Y-%m-%d %H:%M} (폴더)")

# E-PASS 고속수리(수리센터 재불량 로그) - 있으면 읽고, 없으면 '개선효과' 탭에서만 안내한다.
repair_log = pd.DataFrame()
repair_error = None
repair_files = dl.discover_repair_files(BASE_DIR)
if repair_files:
    try:
        repair_log = dl.load_repair_data(tuple(str(p) for p in repair_files),
                                          dl.fingerprint(repair_files))
    except Exception as exc:
        repair_error = str(exc)

# 고속사별 차량대수 (`고속사별차량대수.txt`) - 있으면 고속사별 장애를 "건수"가 아니라
# "차량 1대당 건수"로 보정해 보여준다. 없으면 건수 기준으로 되돌아간다.
FLEET_SIZES = dl.load_company_fleet_sizes(BASE_DIR)


# ── 사이드바: 데이터 현황 + PPT 보고서 ──────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 데이터")
    for label in fault_labels:
        st.caption(label)
    st.caption(f"총 {len(raw):,}건 · {raw['일자'].min():%Y-%m-%d} ~ {raw['일자'].max():%Y-%m-%d}")
    if st.button("🔄 새로고침", width="stretch"):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.markdown("### 📊 PPT 보고서")
    st.caption("연도별 월간 추이 · 월별 장애 건수 · 관심유형 월별구성 · 외장모뎀/B/D "
               "현황 · 전후비교, 5장을 자동으로 만듭니다(항상 전체 기간 기준, "
               "사이드바 필터와 무관). 생성 후 바로 아래 나타나는 다운로드 버튼을 "
               "누르면 각자 PC의 **다운로드** 폴더로 저장됩니다(이 앱을 외부 서버에 "
               "올려도 동일합니다 - 브라우저가 각자의 컴퓨터로 받습니다).")
    if st.button("📊 PPT 보고서 생성", width="stretch"):
        with st.spinner("차트를 그리고 PPT로 묶는 중… (10~20초)"):
            slides = report_builder.build_report_slides(raw, FOCUS_FAULTS, FLEET_SIZE_EST)
            if not slides:
                st.error("보고서를 만들 데이터가 부족합니다.")
            else:
                st.session_state["pptx_bytes"] = ppt_export.build_pptx(
                    slides,
                    title="고속버스 단말기 장애실적 · 개선효과 보고서",
                    subtitle=f"{raw['일자'].min():%Y-%m-%d} ~ {raw['일자'].max():%Y-%m-%d}",
                    footer=f"생성일 {pd.Timestamp.now():%Y-%m-%d} · "
                          f"자동 생성(고속버스 장애실적 대시보드)",
                )
                st.session_state["pptx_name"] = (
                    f"장애실적_개선효과_보고서_{pd.Timestamp.now():%Y%m%d_%H%M%S}.pptx")
    if "pptx_bytes" in st.session_state:
        st.download_button(
            "⬇️ PPT 다운로드", st.session_state["pptx_bytes"],
            file_name=st.session_state["pptx_name"],
            mime=("application/vnd.openxmlformats-officedocument"
                 ".presentationml.presentation"),
            width="stretch",
        )


# ── 필터 (모든 차트가 같은 구간을 본다) ────────────────────────────────
with st.sidebar:
    st.divider()
    st.markdown("### 🔎 필터")

    dmin, dmax = raw["일자"].min().date(), raw["일자"].max().date()
    preset = st.radio("기간", ["전체", "최근 12주", "최근 26주", "올해", "직접 선택"],
                      index=0, horizontal=False)
    if preset == "전체":
        d_from, d_to = dmin, dmax
    elif preset == "최근 12주":
        d_from, d_to = (raw["일자"].max() - pd.Timedelta(weeks=12)).date(), dmax
    elif preset == "최근 26주":
        d_from, d_to = (raw["일자"].max() - pd.Timedelta(weeks=26)).date(), dmax
    elif preset == "올해":
        d_from, d_to = pd.Timestamp(dmax.year, 1, 1).date(), dmax
    else:
        picked = st.date_input("직접 선택", value=(dmin, dmax),
                               min_value=dmin, max_value=dmax)
        d_from, d_to = picked if isinstance(picked, tuple) and len(picked) == 2 else (dmin, dmax)

    def multi(label: str, col: str):
        opts = sorted(raw[col].dropna().unique().tolist())
        return st.multiselect(label, opts, default=[], placeholder="전체")

    f_region = multi("처리지역", "처리지역")
    f_base = multi("처리거점", "처리거점")
    f_comp = multi("고속사", "고속사")
    f_dev = multi("단말기 구분", "단말기 구분")
    f_fault = multi("장애 대분류", "장애 대분류")
    f_action = multi("처리 대분류", "처리 대분류")
    f_staff = multi("처리담당자", "처리담당자")
    f_type = multi("접수구분", "접수구분")

mask = (raw["일자"].dt.date >= d_from) & (raw["일자"].dt.date <= d_to)
for col, sel in [("처리지역", f_region), ("처리거점", f_base), ("고속사", f_comp),
                 ("단말기 구분", f_dev), ("장애 대분류", f_fault),
                 ("처리 대분류", f_action), ("처리담당자", f_staff), ("접수구분", f_type)]:
    if sel:
        mask &= raw[col].isin(sel)
df = raw[mask].copy()

st.title("🚌 고속버스 단말기 장애실적 대시보드")
st.caption(f"{d_from:%Y-%m-%d} ~ {d_to:%Y-%m-%d} · 선택 {len(df):,}건 / 전체 {len(raw):,}건"
           + ("  ·  필터 적용중" if len(df) != len(raw) else ""))

if df.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다. 사이드바에서 필터를 완화하세요.")
    st.stop()

wk = dl.weekly_counts(df)

# 관심 장애유형 탭 전용 범위 - 사이드바의 다른 필터(기간·지역·고속사·담당자 등)는
# 그대로 따르되, '장애 대분류' 필터만은 적용하지 않는다. 이 탭 자체가 특정 장애유형을
# 골라 보는 화면이라, 장애대분류 필터에 다른 값이 걸려 있으면 화면이 비어버리기 때문.
_scope_mask = (raw["일자"].dt.date >= d_from) & (raw["일자"].dt.date <= d_to)
for _col, _sel in [("처리지역", f_region), ("처리거점", f_base), ("고속사", f_comp),
                    ("단말기 구분", f_dev), ("처리 대분류", f_action),
                    ("처리담당자", f_staff), ("접수구분", f_type)]:
    if _sel:
        _scope_mask &= raw[_col].isin(_sel)
scope = raw[_scope_mask].copy()
foc = scope[scope["장애 대분류"].isin(FOCUS_FAULTS)].copy()
# repair_log/repair_error는 위 '데이터 적재' 단계에서 이미 채워져 있다
# (사이드바 필터와 무관하게 항상 전체 기간 - 단말기 S/N에는 지역·고속사 정보가 없다).


# ── KPI ─────────────────────────────────────────────────────────────
def kpis():
    """3열 2행 - 6열로 늘어놓으면 좁은 화면에서 숫자가 말줄임으로 잘린다."""
    c = st.columns(3)
    c[0].metric("총 장애 건수", f"{len(df):,}건")

    last4 = int(wk["건수"].tail(4).sum())
    prev4 = int(wk["건수"].tail(8).head(4).sum()) if len(wk) >= 8 else None
    delta = None if not prev4 else f"{last4 - prev4:+,}건 ({(last4 - prev4) / prev4 * 100:+.0f}%)"
    c[1].metric("최근 4주", f"{last4:,}건", delta=delta, delta_color="inverse",
                help="직전 4주 대비. 장애는 줄어드는 것이 좋으므로 감소를 초록으로 표시합니다.")

    c[2].metric("주 평균", f"{wk['건수'].mean():.1f}건",
                help=f"집계 주차 {len(wk)}주 기준")

    c = st.columns(3)
    veh = df["차량번호"].nunique()
    c[0].metric("대상 차량", f"{veh:,}대",
                help="기간 내 장애가 1건 이상 발생한 고유 차량 수")

    rep = int(df["단말기교체"].sum())
    c[1].metric("단말기 교체", f"{rep:,}건", delta=f"전체의 {rep / len(df) * 100:.0f}%",
                delta_color="off", delta_arrow="off")

    topf = df["장애 대분류"].value_counts()
    c[2].metric(f"최다 장애유형 · {topf.index[0]}", f"{topf.iloc[0]:,}건",
                delta=f"전체의 {topf.iloc[0] / len(df) * 100:.0f}%",
                delta_color="off", delta_arrow="off")


kpis()
st.divider()

tabs = st.tabs(["📈 추세", "📡 관심 장애유형", "🛠️ 개선효과", "🧩 장애·처리 유형",
                "🗺️ 지역·고속사", "🚍 차량·단말기", "🔧 자재·담당자", "📄 데이터"])


# ── 1. 추세 ─────────────────────────────────────────────────────────
with tabs[0]:
    left, right = st.columns([3, 2])

    with left:
        recent = wk.tail(26).copy()
        recent["4주 이동평균"] = wk["건수"].rolling(4).mean().tail(26).round(1)
        fig = ch.vbar_line(recent["주차라벨"], recent["건수"], recent["4주 이동평균"], PAL,
                           bar_name="주간 건수", line_name="4주 이동평균", height=330)
        fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
        ch.block("주차별 장애 건수 (최근 26주)", fig,
                 recent[["주차라벨", "시작일", "종료일", "건수", "4주 이동평균"]], "wk26",
                 caption="주차는 원본 파일의 주차 구분(목~수)을 그대로 따릅니다.")

    with right:
        mon = (df.groupby(["연도", "월"], as_index=False)
                 .size().rename(columns={"size": "건수"}))
        # 한 달치도 안 되는 연도는 점 하나로만 찍혀 오히려 읽기를 방해한다
        months = mon.groupby("연도")["월"].nunique()
        drawn = sorted(int(y) for y in months[months >= 2].index)
        dropped = sorted(int(y) for y in months[months < 2].index)
        piv = mon.pivot(index="월", columns="연도", values="건수").reset_index()
        if not drawn:
            st.markdown("**연도별 월간 추이 비교**")
            st.info("선택 구간에 두 달 이상인 연도가 없어 비교선을 그릴 수 없습니다. "
                    "기간을 넓혀 보세요.")
            st.dataframe(piv, hide_index=True, height="auto")
        else:
            fig = ch.multiline(mon[mon["연도"].isin(drawn)], "월", "건수", "연도", PAL,
                               height=330, xtitle="월", order=drawn)
            fig.update_xaxes(tickmode="array", tickvals=list(range(1, 13)),
                             ticktext=[f"{m}월" for m in range(1, 13)], showgrid=False)
            note = "같은 달을 연도끼리 겹쳐 봅니다. 색은 연도에 고정됩니다."
            partial = [y for y in drawn if int(months[y]) < 12]
            if partial:
                note += f"  ·  {', '.join(f'{y}년' for y in partial)}은 부분 기간입니다."
            if dropped:
                note += (f"  ·  {', '.join(f'{y}년' for y in dropped)}은 1개월 미만이라 "
                         "선에서 제외했습니다(표에는 있음).")
            ch.block("연도별 월간 추이 비교", fig, piv, "yoy", caption=note)

    st.divider()
    left, right = st.columns([3, 2])

    with left:
        m = (df.groupby("연월라벨", as_index=False).size()
               .rename(columns={"size": "건수"}).sort_values("연월라벨"))
        m["3개월 이동평균"] = m["건수"].rolling(3).mean().round(1)
        fig = ch.vbar_line(m["연월라벨"], m["건수"], m["3개월 이동평균"], PAL,
                           bar_name="월간 건수", line_name="3개월 이동평균", height=320)
        fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
        ch.block("월별 장애 건수 (전체 기간)", fig, m, "monthly")

    with right:
        names = ["월", "화", "수", "목", "금", "토", "일"]
        dow = (df.groupby("요일", as_index=False).size()
                 .rename(columns={"size": "건수"}))
        dow["요일명"] = dow["요일"].map(lambda i: names[i])
        fig = ch.hbar(dow[["요일명", "건수"]], "요일명", "건수", PAL, height=320)
        fig.update_yaxes(categoryorder="array", categoryarray=names[::-1])
        ch.block("요일별 접수 분포", fig, dow[["요일명", "건수"]], "dow",
                 caption="현장 출동이 몰리는 요일을 확인합니다.")

    st.divider()
    top5 = df["장애 대분류"].value_counts().head(5).index.tolist()
    d5 = df.copy()
    d5["유형"] = d5["장애 대분류"].where(d5["장애 대분류"].isin(top5), "기타")
    piv = (d5.pivot_table(index="연월라벨", columns="유형", values="일자",
                          aggfunc="size", fill_value=0)
             .reindex(columns=top5 + ["기타"], fill_value=0).sort_index())
    fig = ch.stacked_bar(piv, PAL, height=360)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
    ch.block("월별 장애유형 구성 (상위 5종 + 기타)", fig, piv.reset_index(), "mix",
             caption="상위 5종 외에는 '기타'로 묶습니다 — 색은 6슬롯까지만 씁니다.")


# ── 2. 관심 장애유형 (LTE모뎀망연결지연 · BMS접속지연 · 승차연결지연 · 승차단말기통신불량)
with tabs[1]:
    st.markdown("**통신·연결 계열 4종 장애**")
    st.caption(
        "이 탭은 사이드바의 '장애 대분류' 필터와 무관하게 이 4종만 봅니다 "
        "(기간·지역·고속사 등 나머지 필터는 그대로 적용됩니다). "
        f"선택 범위 {len(scope):,}건 중 {len(foc):,}건({len(foc) / max(len(scope), 1) * 100:.0f}%)."
    )

    if foc.empty:
        st.info("선택한 범위에는 이 4종 장애가 없습니다. 기간이나 다른 필터를 넓혀 보세요.")
    else:
        counts = foc["장애 대분류"].value_counts()
        foc_wk = dl.weekly_counts(foc)
        last4 = int(foc_wk["건수"].tail(4).sum())
        prev4 = int(foc_wk["건수"].tail(8).head(4).sum()) if len(foc_wk) >= 8 else None
        delta = None if not prev4 else f"{last4 - prev4:+,}건 ({(last4 - prev4) / prev4 * 100:+.0f}%)"

        c = st.columns(5)
        c[0].metric("4종 합계", f"{len(foc):,}건", delta=delta, delta_color="inverse",
                    help="최근 4주 vs 직전 4주. 장애는 줄어드는 게 좋으므로 감소를 초록으로 표시합니다.")
        for i, name in enumerate(FOCUS_FAULTS, start=1):
            n = int(counts.get(name, 0))
            c[i].metric(name, f"{n:,}건",
                        delta=f"{n / max(len(scope), 1) * 100:.0f}%",
                        delta_color="off", delta_arrow="off",
                        help="선택 범위 전체(장애대분류 필터 제외) 대비 비중")

        st.divider()
        left, right = st.columns([3, 2])

        with left:
            weeks_all = (scope[["주차키", "주차라벨"]].drop_duplicates()
                                .sort_values("주차키").tail(26))
            wk_piv = (foc.groupby(["주차키", "장애 대분류"]).size().unstack(fill_value=0)
                         .reindex(index=weeks_all["주차키"], columns=FOCUS_FAULTS, fill_value=0))
            wk_piv.index = pd.Index(weeks_all["주차라벨"].values, name="주차라벨")
            long = wk_piv.reset_index().melt(id_vars="주차라벨", var_name="장애 대분류",
                                              value_name="건수")
            fig = ch.multiline(long, "주차라벨", "건수", "장애 대분류", PAL, height=340,
                               order=FOCUS_FAULTS)
            fig.update_xaxes(tickangle=-45, tickfont=dict(size=14), type="category",
                             categoryorder="array", categoryarray=list(wk_piv.index))
            ch.block("주차별 추이 (최근 26주)", fig, wk_piv.reset_index(), "foc_wk",
                     caption="0건인 주는 값이 실제로 0이며, 데이터 누락이 아닙니다.")

        with right:
            comp = foc["고속사"].value_counts().rename_axis("고속사").reset_index(name="건수")
            if FLEET_SIZES:
                rated = dl.with_fleet_rate(comp, "고속사", FLEET_SIZES)
                top_rate = (rated.dropna(subset=["차량대수"])
                                 .sort_values("대당건수", ascending=False).head(10))
                ch.block(
                    "고속사별 4종 발생률 (차량 1대당) Top 10",
                    ch.hbar(top_rate, "고속사", "대당건수", PAL, height=340, suffix="건/대",
                           text_fmt="{:.2f}", hover_fmt="%{x:,.2f}"),
                    rated.sort_values("대당건수", ascending=False, na_position="last"),
                    "foc_comp",
                    caption="차량 1대당 발생 건수 기준입니다 - 차량이 많은 고속사가 단순 건수만으로 "
                           "불리하게 보이는 걸 막습니다.",
                )
            else:
                ch.block("고속사별 4종 합계 Top 10",
                         ch.hbar(comp.head(10), "고속사", "건수", PAL, height=340),
                         comp, "foc_comp",
                         caption="고속사별로 얼마나 자주 겪는지 봅니다 - `고속사별차량대수.txt`를 "
                                "이 폴더에 두면 차량 1대당 발생률로 볼 수 있습니다.")

        st.divider()
        left, right = st.columns([3, 2])

        with left:
            months_all = sorted(scope["연월라벨"].unique())
            mon_piv = (foc.groupby(["연월라벨", "장애 대분류"]).size().unstack(fill_value=0)
                          .reindex(index=months_all, columns=FOCUS_FAULTS, fill_value=0))
            fig = ch.stacked_bar(mon_piv, PAL, height=340, other="__none__")
            fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
            ch.block("월별 구성 (전체 기간)", fig, mon_piv.reset_index(), "foc_mon",
                     caption="4종의 월별 건수 구성입니다. 색은 유형에 고정됩니다.")

        with right:
            comps10 = foc["고속사"].value_counts().head(10).index.tolist()
            mat = (foc[foc["고속사"].isin(comps10)]
                   .pivot_table(index="고속사", columns="장애 대분류", values="일자",
                                aggfunc="size", fill_value=0)
                   .reindex(index=comps10, columns=FOCUS_FAULTS, fill_value=0))
            ch.block("고속사 × 유형 히트맵", ch.heatmap(mat, PAL, height=340),
                     mat.reset_index(), "foc_hm",
                     caption="어느 고속사가 어떤 유형에 유독 취약한지 봅니다.")

        st.divider()
        left, right = st.columns(2)
        with left:
            t = dl.top_n(foc, "처리지역", 10)
            ch.block("처리지역별 4종 합계", ch.hbar(t, "처리지역", "건수", PAL, height=300),
                     t, "foc_region")
        with right:
            t = dl.top_n(foc, "처리 대분류", 8)
            ch.block("처리 방식", ch.hbar(t, "처리 대분류", "건수", PAL, height=300),
                     t, "foc_action", caption="이 4종 장애가 실제로 어떻게 조치되는지 봅니다.")

        st.divider()
        st.markdown("**원본 데이터 (4종만)**")
        cols = ["일자", "주차라벨", "처리지역", "처리거점", "고속사", "차량번호",
                "장애 대분류", "처리 대분류", "장애 접수 내용",
                "세부 조치 내용 (현장 확인 증상 / 조치내용)", "처리담당자"]
        view = foc[cols].sort_values("일자", ascending=False)
        st.dataframe(view, hide_index=True, height=360,
                     column_config={"일자": st.column_config.DateColumn("일자", format="YYYY-MM-DD")})
        st.download_button("CSV 내려받기", view.to_csv(index=False).encode("utf-8-sig"),
                           file_name="관심장애유형_4종.csv", mime="text/csv", key="d_foc")


# ── 2. 개선효과 (외장 LTE모뎀 전환 · 수리포인트) ────────────────────────
with tabs[2]:
    st.markdown("**외장 LTE모뎀 전환 · 수리포인트가 필드 장애에 미친 효과**")
    st.caption(
        "이 탭은 사이드바 필터와 무관하게 전체 기간을 봅니다. 2025년 하반기부터 일부 단말의 "
        "내장 LTE모뎀을 외장으로 전환 중이고, 2026년 2월부터는 수리센터(E-PASS)에서 "
        "LTE·BMS·승차연결지연 관련 보드 수리포인트를 잡아 수리를 진행하고 있습니다."
    )

    # ---- 외장 LTE모뎀 전환 · B/D 보드교체 현황 (세부 조치 내용 텍스트에서 추출) ----
    st.markdown("#### 📶 외장 LTE모뎀 전환 · B/D(AFC/BMS) 보드교체 현황")
    st.caption(
        "`세부 조치 내용` 칸의 문구로 구분합니다 — `외장모뎀적용`은 이번 조치로 외장 LTE모뎀을 "
        "새로 장착했다는 뜻이고, `외장모뎀적용차량`은 이 차량이 이미 전환된 상태라는 맥락 표기일 "
        "뿐 신규 전환이 아닙니다. `B/D교체적용`은 운전자단말기 내 AFC보드·BMS보드 중 하나 이상을 "
        "교체했다는 뜻입니다(오타 변형도 함께 잡습니다). `외장모뎀적용`이 표준 문구로 자리잡기 "
        "전(2025-05-29 이전)에는 `외장형모뎀교체` · `외장모뎀교체` · `외장모뎀 장착`처럼 다르게 "
        "적힌 경우도 신규 전환으로 함께 잡습니다 — 이 파일에는 제조사(텔라딘·CNSLINK 등) 정보는 "
        "없어 구분하지 못하며, 고장 출동과 함께 이뤄진 건만 잡히므로 전체 설치 대수보다는 "
        "적게 집계될 수 있습니다."
    )
    iv = dl.extract_intervention_events(raw)
    n_new = int(iv["외장모뎀_신규전환"].sum())
    n_already = int(iv["외장모뎀_기적용"].sum())
    n_bd = int(iv["BD보드_교체"].sum())
    n_bd_only = int(iv["BD보드_단독"].sum())

    if n_new == 0 and n_bd == 0:
        st.info("장애실적 세부 조치 내용에서 외장모뎀·B/D보드 관련 기록을 찾지 못했습니다.")
    else:
        n_new_veh = iv.loc[iv["외장모뎀_신규전환"], "차량번호"].nunique()
        n_bd_veh = iv.loc[iv["BD보드_교체"], "차량번호"].nunique()

        c = st.columns(4)
        c[0].metric("외장모뎀 신규전환", f"{n_new:,}건",
                    help=f"전환 차량 {n_new_veh:,}대 · "
                         f"{iv.loc[iv['외장모뎀_신규전환'],'일자'].min():%Y-%m-%d} ~ "
                         f"{iv.loc[iv['외장모뎀_신규전환'],'일자'].max():%Y-%m-%d}")
        c[1].metric("이미 전환된 차량(참고)", f"{n_already:,}건",
                    help="신규 전환이 아니라, 다른 조치 중 '이미 외장모뎀 적용 상태'라고 "
                         "달린 맥락 표기입니다.")
        c[2].metric("B/D 보드교체(AFC/BMS)", f"{n_bd:,}건",
                    delta=f"외장모뎀 동반 {n_bd - n_bd_only:,}건 · 단독 {n_bd_only:,}건",
                    delta_color="off", delta_arrow="off")
        c[3].metric("B/D 보드교체 차량", f"{n_bd_veh:,}대")
        st.caption(
            f"사용자 제공 추정치 기준 운영 차량 약 {FLEET_SIZE_EST:,}대(대략치) 대비 — "
            f"외장모뎀 전환 차량은 약 {n_new_veh / FLEET_SIZE_EST * 100:.0f}%, "
            f"B/D 보드교체 차량은 약 {n_bd_veh / FLEET_SIZE_EST * 100:.0f}%까지 진행됐습니다."
        )

        left, right = st.columns(2)
        with left:
            m_new = (iv[iv["외장모뎀_신규전환"]].groupby(iv.loc[iv["외장모뎀_신규전환"], "일자"]
                     .dt.to_period("M").astype(str)).size())
            m_new = m_new.reindex(pd.period_range(m_new.index.min(), m_new.index.max(), freq="M")
                                  .astype(str), fill_value=0) if len(m_new) else m_new
            cum = m_new.cumsum()
            fig = ch.vbar_line(m_new.index, m_new.values, cum.values, PAL, height=300,
                               bar_name="월별 신규전환", line_name="누적 전환대수")
            fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
            tbl = pd.DataFrame({"연월": m_new.index, "신규전환": m_new.values, "누적": cum.values})
            ch.block("월별 외장모뎀 신규전환", fig, tbl, "ext_new_mon",
                     caption="2025년 하반기부터 시작해 2026년에도 계속 이어지고 있습니다.")
        with right:
            bd = iv[iv["BD보드_교체"]].copy()
            bd["구분"] = bd["BD보드_단독"].map({True: "B/D 단독", False: "외장모뎀 동반"})
            piv = (bd.groupby([bd["일자"].dt.to_period("M").astype(str), "구분"]).size()
                     .unstack(fill_value=0).reindex(columns=["B/D 단독", "외장모뎀 동반"], fill_value=0))
            fig = ch.stacked_bar(piv, PAL, height=300, other="__none__")
            fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
            ch.block("월별 B/D 보드교체", fig, piv.reset_index(), "bd_mon",
                     caption="외장모뎀 교체와 같은 방문에서 함께 이뤄진 경우가 많습니다.")

        with st.expander("전환/보드교체 이력 (원본)"):
            cols = ["일자", "연월라벨", "고속사", "차량번호", "장애 대분류", "처리 대분류",
                    dl.ACTION_TEXT_COL, "교체전", "교체후"]
            hist = iv[iv["외장모뎀_신규전환"] | iv["BD보드_교체"]][cols].sort_values("일자", ascending=False)
            st.dataframe(hist, hide_index=True, height=320,
                         column_config={"일자": st.column_config.DateColumn(format="YYYY-MM-DD")})
            st.download_button("CSV 내려받기", hist.to_csv(index=False).encode("utf-8-sig"),
                               file_name="전환_보드교체_이력.csv", mime="text/csv", key="d_iv")

        st.divider()

        # ---- 전후 비교: 조치가 실제로 재발을 줄였는가 ----
        st.markdown("#### ⚖️ 전후 비교 — 조치 이후 관심 장애(LTE·BMS·승차연결지연)가 줄었는가")
        st.caption(
            "**어떻게 계산하나**: 조치를 받은 차량마다 시행일 기준 전 90일 · 후 90일 동안 "
            "그 차량에서 관심 장애가 몇 건 났는지 셉니다(한 차량에 조치가 여러 번이면 "
            "가장 이른 시행일 기준). 그 다음 조치받은 차량 전체에 대해 '전 건수'와 "
            "'후 건수'를 각각 평균 낸 값이 아래 막대(건/대)입니다 — "
            "**평균 발생률 = (전체 차량의 발생 건수 합) ÷ (관찰 차량 수)**. "
            "설치 직후 14일은 같은 방문의 재작업일 수 있어 '후' 집계에서 뺐고, "
            "조치 후 90일이 아직 안 지난(최근에 조치한) 차량은 관찰기간 부족으로 뺐습니다."
        )
        FOCUS3 = FOCUS_FAULTS[:3]
        studies = {
            "외장모뎀 신규전환": dl.vehicle_event_study(raw, iv["외장모뎀_신규전환"], focus_faults=FOCUS3),
            "B/D교체(단독)": dl.vehicle_event_study(raw, iv["BD보드_단독"], focus_faults=FOCUS3),
            "B/D교체(전체)": dl.vehicle_event_study(raw, iv["BD보드_교체"], focus_faults=FOCUS3),
        }
        cats, before_rate, after_rate, ns = [], [], [], []
        before_sums, after_sums = [], []
        for name, r in studies.items():
            if r.empty:
                continue
            cats.append(name)
            before_sums.append(int(r["전"].sum()))
            after_sums.append(int(r["후"].sum()))
            before_rate.append(r["전"].mean())
            after_rate.append(r["후"].mean())
            ns.append(len(r))
        if not cats:
            st.info("아직 90일 관찰 기간을 채운 조치 건이 없습니다.")
        else:
            fig = ch.paired_bar(cats, before_rate, after_rate, PAL, height=320,
                                suffix="건/대", before_name="전 90일", after_name="후 90일",
                                value_fmt="{:.2f}")
            fig.update_xaxes(tickfont=dict(size=15))
            cap = ("막대는 차량 1대당 평균 발생건수. 관찰 차량수(n): "
                  + " · ".join(f"{c}={n}대" for c, n in zip(cats, ns)))
            tbl = pd.DataFrame({"구분": cats, "관찰차량수": ns,
                               "전_합계건수": before_sums, "전_평균(건/대)": before_rate,
                               "후_합계건수": after_sums, "후_평균(건/대)": after_rate})
            ch.block("조치 전/후 관심 장애 발생률", fig, tbl, "event_study", caption=cap)
            # 실제 숫자로 "0.10건/대가 어디서 나왔나"를 바로 확인시킨다
            ex_i = 0
            st.caption(
                f"🔢 예시로 계산 과정을 풀면 — **{cats[ex_i]}**: 조치받은 차량 "
                f"{ns[ex_i]}대에서 전 90일 동안 합쳐서 {before_sums[ex_i]}건 발생 → "
                f"{before_sums[ex_i]}건 ÷ {ns[ex_i]}대 = {before_rate[ex_i]:.2f}건/대. "
                f"후 90일은 합쳐서 {after_sums[ex_i]}건 → {after_sums[ex_i]}건 ÷ "
                f"{ns[ex_i]}대 = {after_rate[ex_i]:.2f}건/대. (나머지 구분도 같은 방식 - "
                "위 '표로 보기'에서 전_합계건수·후_합계건수로 확인할 수 있습니다.)"
            )

            # 이 해석 문구는 고정 텍스트가 아니라 매번 raw로 다시 계산한다 - 데이터가
            # 바뀌면(새 파일 업로드 등) 결과 방향(개선/악화)이 뒤집힐 수 있으므로, 그때도
            # 틀린 말을 하지 않도록 실제 before_rate/after_rate에서 문장을 만든다.
            rates = dict(zip(cats, zip(before_rate, after_rate, ns)))

            def _verdict(before: float, after: float) -> str:
                if before == 0:
                    return "개선" if after == 0 else "악화"
                change = (after - before) / before
                if change <= -0.10:
                    return "개선"
                if change >= 0.10:
                    return "악화"
                return "변화 미미"

            msgs = []
            if "외장모뎀 신규전환" in rates:
                b, a, n = rates["외장모뎀 신규전환"]
                verdict = _verdict(b, a)
                new_mask = iv["외장모뎀_신규전환"]
                lte_pct = (iv.loc[new_mask, "장애 대분류"].eq("LTE모뎀망연결지연").mean() * 100
                          if new_mask.sum() else None)
                if verdict != "개선":
                    reason = (f"전환이 대개 LTE 고장을 계기로 이뤄져(신규전환의 {lte_pct:.0f}%가 "
                             f"LTE모뎀망연결지연으로 촉발) 표본 자체가 문제 차량에 쏠려 있고, "
                             if lte_pct else "")
                    msgs.append(
                        f"**외장모뎀 신규전환은 이 검증에서 개선 신호가 뚜렷하지 않습니다** "
                        f"({b:.2f}건/대 → {a:.2f}건/대, n={n}대) — {reason}"
                        f"표본이 작아 몇 대의 우연한 재발에도 비율이 크게 흔들립니다."
                    )
                else:
                    msgs.append(f"**외장모뎀 신규전환도 개선 신호**를 보입니다 "
                               f"({b:.2f}건/대 → {a:.2f}건/대, n={n}대).")

            bd_parts = []
            for key, label in [("B/D교체(단독)", "단독"), ("B/D교체(전체)", "전체")]:
                if key in rates:
                    b, a, n = rates[key]
                    bd_parts.append(f"{label} {b:.2f}→{a:.2f}건/대(n={n}대, {_verdict(b, a)})")
            if bd_parts:
                bd_verdicts = [_verdict(*rates[k][:2]) for k in
                              ["B/D교체(단독)", "B/D교체(전체)"] if k in rates]
                lead = ("**B/D 보드교체는 반대로 뚜렷한 개선**을 보입니다"
                       if all(v == "개선" for v in bd_verdicts) and
                       (not msgs or "개선" not in msgs[0][:20])
                       else "**B/D 보드교체**는")
                msgs.append(f"{lead} — {', '.join(bd_parts)}. 특히 외장모뎀과 무관하게 "
                           "보드만 교체한 건(단독)의 결과가 이렇다면 우연만은 아닌 것으로 "
                           "보입니다.")

            msgs.append("다만 차량별 전/후 비교일 뿐 정식 대조군 비교가 아니므로, "
                       "둘 다 참고용 신호로 보시는 게 안전합니다.")
            st.warning(" ".join(msgs))

            with st.expander("차량별 전/후 상세"):
                for name, r in studies.items():
                    if r.empty:
                        continue
                    st.markdown(f"**{name}** ({len(r):,}대)")
                    st.dataframe(r.sort_values("이벤트일", ascending=False), hide_index=True,
                                height=220,
                                column_config={"이벤트일": st.column_config.DateColumn(
                                    format="YYYY-MM-DD")})

    st.divider()

    # ---- E-PASS 수리포인트 ----
    st.markdown("#### 🔧 수리포인트 (E-PASS 고속수리)")
    if repair_error:
        st.error(f"수리 로그를 읽지 못했습니다: {repair_error}")
    elif repair_log.empty:
        st.info(
            "`E-PASS고속수리_2026.xlsx` 파일을 찾지 못했습니다. "
            f"`{BASE_DIR}` 폴더에 그 이름의 파일을 두면 이 섹션이 채워집니다."
        )
    else:
        rep = repair_log
        res = dl.build_recurrence(rep, raw)
        last_field_date = raw["일자"].max()

        n_total = len(rep)
        n_driver = int((rep["기기"] == "운전자").sum())
        n_board = int((rep["기기"] == "승차").sum())
        n_lte = int((rep["카테고리"] == "LTE/모뎀").sum())
        n_indoor_repeat = int(rep["재불량여부"].sum())
        obs = res[(last_field_date - res["수리일"]).dt.days >= 60]
        n_obs = len(obs)
        n_recur60 = int((obs["경과일"] <= 60).sum()) if n_obs else 0

        c = st.columns(4)
        c[0].metric("수리 건수", f"{n_total:,}건",
                    help=f"{rep['수리일'].min():%Y-%m-%d} ~ {rep['수리일'].max():%Y-%m-%d}"
                         f" · 운전자단말기 {n_driver:,}건 · 승차단말기 {n_board:,}건")
        c[1].metric("LTE/모뎀 관련", f"{n_lte:,}건",
                    delta=f"전체의 {n_lte / n_total * 100:.0f}%",
                    delta_color="off", delta_arrow="off")
        c[2].metric("수리센터 내 재불량률", f"{n_indoor_repeat / n_total * 100:.0f}%",
                    delta=f"{n_indoor_repeat:,}건 · 1차 수리 후 재검사에서 재발견",
                    delta_color="off", delta_arrow="off",
                    help="확인/재확인 검사 중 같은 단말이 다시 불량으로 잡혀 재작업한 비율")
        c[3].metric("60일 내 현장 재발률",
                    f"{n_recur60 / n_obs * 100:.0f}%" if n_obs else "N/A",
                    delta=f"관찰 가능 {n_obs:,}건 중 {n_recur60:,}건",
                    delta_color="off", delta_arrow="off",
                    help="수리일로부터 60일 이상 지난 건만 대상으로, 그 안에 같은 S/N이 "
                         "현장에서 다시 `교체전`(고장 이탈)으로 잡혔는지 확인한 비율입니다.")

        left, right = st.columns(2)
        with left:
            cats = ["LTE/모뎀", "BMS", "승차연결", "화면/터치", "바코드", "전원/부팅", "기타"]
            mon_piv = (rep.groupby(["연월라벨", "카테고리"]).size().unstack(fill_value=0)
                          .reindex(columns=cats, fill_value=0))
            fig = ch.stacked_bar(mon_piv, PAL, height=320)
            fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
            ch.block("월별 수리 건수 (증상 계열별)", fig, mon_piv.reset_index(), "rep_mon",
                     caption="LTE/모뎀 계열이 절반 이상을 차지합니다 — 색은 관심 장애유형 "
                             "탭과 같은 순서(LTE·BMS·승차연결)를 씁니다.")
        with right:
            t = dl.top_n(rep, "상세수리내역", 10)
            ch.block("상세 수리 내역 Top 10", ch.hbar(t, "상세수리내역", "건수", PAL, height=320),
                     t, "rep_detail", caption="실제로 보드에서 어떤 조치를 하는지 봅니다.")

        st.divider()
        left, right = st.columns([3, 2])
        with left:
            rmon = res.copy()
            rmon["연월"] = rmon["수리일"].dt.to_period("M").astype(str)
            rmon["관찰가능"] = (last_field_date - rmon["수리일"]).dt.days >= 60
            g = (rmon.groupby("연월")
                     .agg(수리건수=("SN", "size"), 관찰가능건수=("관찰가능", "sum"))
                     .reset_index())
            recur60 = (rmon[rmon["관찰가능"] & (rmon["경과일"] <= 60)]
                       .groupby("연월").size().rename("60일내재발"))
            g = g.merge(recur60, on="연월", how="left")
            g["60일내재발"] = g["60일내재발"].fillna(0).astype(int)
            drawable = g[g["관찰가능건수"] > 0].copy()
            drawable["재발률"] = drawable["60일내재발"] / drawable["관찰가능건수"] * 100
            censored = g.loc[g["관찰가능건수"] == 0, "연월"].tolist()
            if drawable.empty:
                st.info("아직 60일 관찰 기간을 채운 수리 건이 없습니다.")
            else:
                # 표본이 작은 달(관찰가능 10건 미만)은 비율이 크게 흔들리므로 n을 함께 보여
                # 과잉해석을 막는다 - 2026-02가 대표적(관찰가능 6건뿐).
                labels = [f"{v:.0f}%(n={n})" for v, n in
                         zip(drawable["재발률"], drawable["관찰가능건수"])]
                fig = ch.vbar(drawable["연월"], drawable["재발률"], PAL, height=300,
                             suffix="%", value_labels=labels)
                fig.update_xaxes(tickangle=-45, tickfont=dict(size=14))
                cap = ("수리일로부터 60일이 지난 건만 계산합니다(관찰기간 확보 전 건은 제외). "
                      "막대 위 n은 그 달의 관찰가능 건수 - 표본이 작은 달(2월 등)은 참고만 하세요.")
                if censored:
                    cap += f" 제외: {', '.join(censored)}(아직 60일 미경과)."
                ch.block("월별 60일 내 현장 재발률", fig, g, "rep_recur", caption=cap)
        with right:
            recur = res[res["재발여부"]]
            recur_faults = dl.top_n(recur, "재발장애", 8)
            if recur_faults.empty:
                st.info("현장에서 재발이 확인된 건이 없습니다.")
            else:
                n_same = int(recur["재발장애"].eq("LTE모뎀망연결지연").sum())
                cap = (f"수리 후 다시 고장 났을 때, 어떤 장애로 잡히는지 봅니다. 재발 {len(recur)}건 중 "
                      f"{n_same}건({n_same / len(recur) * 100:.0f}%)만 같은 LTE 계열이고, "
                      "나머지는 다른 증상(주로 운전자화면불량)이라 - 60일 재발률을 "
                      "'LTE 수리가 실패한 비율'로 곧이곧대로 읽기보다는 상한선으로 보는 편이 안전합니다.")
                ch.block("재발 시 장애 유형", ch.hbar(recur_faults, "재발장애", "건수", PAL, height=300),
                         recur_faults, "rep_recur_type", caption=cap)

        st.divider()
        st.markdown("**수리 방식별 60일 내 현장 재발률**")
        st.caption(
            "여기서 말하는 `Main B/D`는 E-PASS 수리 로그의 `주 수리 내역` 분류입니다 - "
            "운전자단말기의 **메인보드**(LTE모뎀 회로가 실장된 보드) 전체를 가리키며, "
            "위 '외장 LTE모뎀 전환 · B/D(AFC/BMS) 보드교체 현황'에서 말한 `B/D교체적용`"
            "(AFC보드·BMS보드)과는 **다른 보드**입니다 - 이름은 같은 'B/D'(보드)지만 "
            "가리키는 물리 보드가 다릅니다."
        )
        rtype = res.copy()
        rtype["구분"] = rtype["주 수리 내역"].where(
            rtype["주 수리 내역"].isin(["Main B/D 교체", "Main B/D 수리", "MODEM"]),
            "기타(NDF·I/O·LCD 등)")
        rtype["관찰가능"] = (last_field_date - rtype["수리일"]).dt.days >= 60
        rows = []
        for name in ["Main B/D 교체", "Main B/D 수리", "MODEM", "기타(NDF·I/O·LCD 등)"]:
            sub = rtype[rtype["구분"] == name]
            obs = sub[sub["관찰가능"]]
            n = len(obs)
            r60 = int((obs["경과일"] <= 60).sum()) if n else 0
            rows.append({"구분": name, "전체건수": len(sub), "관찰가능건수": n, "60일내재발": r60,
                        "재발률(%)": round(r60 / n * 100, 1) if n else None})
        rt = pd.DataFrame(rows)
        drawable_rt = rt[rt["관찰가능건수"] > 0]
        if drawable_rt.empty:
            st.info("아직 60일 관찰 기간을 채운 건이 없습니다.")
        else:
            labels = [f"{v:.0f}%(n={n})" for v, n in
                     zip(drawable_rt["재발률(%)"], drawable_rt["관찰가능건수"])]
            fig = ch.vbar(drawable_rt["구분"], drawable_rt["재발률(%)"], PAL, height=300, suffix="%",
                         value_labels=labels)
            cap = "막대 위 n은 관찰가능 건수(수리 후 60일이 지난 건만)."
            skipped = rt.loc[rt["관찰가능건수"] == 0, "구분"].tolist()
            if skipped:
                cap += f" {', '.join(skipped)}은 전부 최근(예: 8/31 일괄 시행)이라 아직 관찰 불가."
            ch.block("수리 방식별 60일 내 현장 재발률", fig, rt, "rep_type_recur", caption=cap)

        with st.expander("현장 재발 상세 내역"):
            cols = ["수리일", "SN", "카테고리", "상세수리내역", "재발일", "경과일", "재발장애"]
            rv = res.loc[res["재발여부"], cols].sort_values("수리일", ascending=False)
            st.dataframe(rv, hide_index=True, height=320,
                         column_config={
                             "수리일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                             "재발일": st.column_config.DateColumn(format="YYYY-MM-DD"),
                             "경과일": st.column_config.NumberColumn(format="%d일"),
                         })
            st.download_button("CSV 내려받기", rv.to_csv(index=False).encode("utf-8-sig"),
                               file_name="수리후_현장재발.csv", mime="text/csv", key="d_recur")

    st.divider()

    # ---- 관심 장애유형 4종 월별 추이 (2024년~, 개선활동 시점 표시) ----
    # PPT 보고서의 같은 이름 슬라이드와 정확히 같은 차트를 쓴다(report_builder 참고).
    st.markdown("#### 📉 필드 장애 추이 (2024년~, LTE·BMS·승차연결지연·승차단말기통신불량)")
    if raw[raw["연도"] >= 2024].empty:
        st.info("2024년 이후 데이터가 없습니다.")
    else:
        fig, table = report_builder.build_field_trend_figure(raw, FOCUS_FAULTS, PAL, height=360)
        ch.block("월별 추이 (2024년~)", fig, table, "field_trend_monthly",
                 caption="점선 이후가 각 개선활동의 적용 기간입니다. 월 단위라 해당 월 전체가 "
                         "표시되며, 실제 시작일은 각각 2/5, 8/6입니다. 계절 요인 등 다른 "
                         "변수도 함께 작용할 수 있어, 이 그래프만으로 인과를 단정하지는 않습니다.")

    st.divider()

    # ---- AFC-승차단말 통신보완 펌웨어 (2026-08-06 배포) 효과 ----
    st.markdown("#### 🎫 AFC-승차단말 통신보완 펌웨어 효과 (2026-08-06 배포)")
    st.caption(
        "이 펌웨어는 특정 차량만이 아니라 전 차량에 같은 날 배포됐습니다 - 그래서 외장모뎀"
        "전환·B/D교체처럼 차량별 전/후를 비교할 대조군이 없습니다. 대신 배포일 전후로 "
        "`승차연결지연`·`승차단말기통신불량` 전체(모집단) 발생 추이가 꺾였는지를 봅니다."
    )
    fw_targets = raw[raw["장애 대분류"].isin(FIRMWARE_TARGET_FAULTS)]
    last_date = raw["일자"].max()
    stats = dl.before_after_monthly_rate(raw, FIRMWARE_TARGET_FAULTS, FIRMWARE_DATE)
    days_since = stats["days_since"]

    if days_since < 0:
        st.info("아직 배포일 전입니다.")
    else:
        before_rate, after_rate = stats["before_rate"], stats["after_rate"]

        c = st.columns(4)
        c[0].metric("배포 후 경과일", f"{days_since:,}일",
                    help=f"데이터 최신일 {last_date:%Y-%m-%d} 기준")
        c[1].metric("배포 후 발생건수", f"{stats['after_count']:,}건",
                    delta=f"월 환산 {after_rate:.1f}건/월", delta_color="off", delta_arrow="off")
        c[2].metric("배포 전 월평균(최근 6개월)", f"{before_rate:.1f}건/월")
        diff_pct = ((after_rate - before_rate) / before_rate * 100) if before_rate else None
        c[3].metric("월 환산 증감", f"{after_rate - before_rate:+.1f}건/월",
                    delta=f"{diff_pct:+.0f}%" if diff_pct is not None else None,
                    delta_color="inverse")

        months_all = sorted(raw.loc[raw["연도"] >= 2024, "연월라벨"].unique())
        mon_fw = (fw_targets.groupby(["연월라벨", "장애 대분류"]).size().unstack(fill_value=0)
                            .reindex(index=months_all, columns=FIRMWARE_TARGET_FAULTS,
                                    fill_value=0))
        long_fw = (mon_fw.reset_index()
                        .melt(id_vars="연월라벨", var_name="장애 대분류", value_name="건수"))
        # 색은 관심 장애유형 탭과 같은 슬롯(승차연결지연=3번째, 승차단말기통신불량=4번째)을
        # 그대로 유지한다 - multiline()의 order 위치가 곧 색 슬롯이라, 앞의 LTE·BMS 두
        # 자리를 빈 값으로 채워 넣어 자리를 맞춘다(범례에는 뜨지 않도록 아래에서 감춘다).
        fig_fw = ch.multiline(long_fw, "연월라벨", "건수", "장애 대분류", PAL, height=320,
                              order=FOCUS_FAULTS)
        fig_fw.for_each_trace(lambda t: t.update(visible=False) if t.name not in
                              FIRMWARE_TARGET_FAULTS else ())
        fig_fw.update_layout(showlegend=True)
        fig_fw.update_xaxes(tickangle=-45, tickfont=dict(size=14), type="category",
                            categoryorder="array", categoryarray=months_all)
        if "2026-08" in months_all:
            fig_fw.add_vline(x="2026-08", line_dash="dot", line_width=1.5,
                             line_color=PAL.c["muted"])
            fig_fw.add_annotation(x="2026-08", y=1.0, yref="paper", showarrow=False,
                                  text="펌웨어 배포(8/6)", font=dict(size=14, color=PAL.c["muted"]),
                                  xanchor="left", yanchor="top")
        ch.block("월별 추이 - 승차연결지연 · 승차단말기통신불량", fig_fw,
                 mon_fw.reset_index(), "firmware_trend",
                 caption=f"배포 후 관찰 기간이 {days_since}일로 아직 짧아(1~2개월치) 참고용 "
                         "신호입니다 - 몇 주 더 지켜봐야 추세로 확정할 수 있습니다.")


# ── 3. 장애·처리 유형 ────────────────────────────────────────────────
with tabs[3]:
    left, right = st.columns(2)
    with left:
        t = dl.top_n(df, "장애 대분류", 12)
        ch.block("장애 대분류 Top 12", ch.hbar(t, "장애 대분류", "건수", PAL),
                 t, "fault", caption=f"전체 {df['장애 대분류'].nunique()}종 중 상위 12종")
    with right:
        t = dl.top_n(df, "처리 대분류", 12)
        ch.block("처리 대분류 Top 12", ch.hbar(t, "처리 대분류", "건수", PAL),
                 t, "action", caption=f"전체 {df['처리 대분류'].nunique()}종 중 상위 12종")

    st.divider()
    left, right = st.columns([2, 3])
    with left:
        t = dl.top_n(df, "단말기 구분", 8)
        ch.block("단말기 구분별 건수", ch.hbar(t, "단말기 구분", "건수", PAL, height=280),
                 t, "device")
    with right:
        combo = (df.groupby(["장애 대분류", "처리 대분류"], as_index=False)
                   .size().rename(columns={"size": "건수"})
                   .sort_values("건수", ascending=False).head(15))
        combo["비중"] = (combo["건수"] / len(df) * 100).round(1)
        st.markdown("**장애 → 처리 조합 Top 15**")
        st.caption("어떤 증상이 어떤 조치로 끝나는지 — 조치 표준화 여지를 봅니다.")
        st.dataframe(combo, hide_index=True, height=360,
                     column_config={"비중": st.column_config.NumberColumn("비중(%)", format="%.1f%%")})

    st.divider()
    fy = df["장애 대분류"].value_counts().head(8).index.tolist()
    mat = (df[df["장애 대분류"].isin(fy)]
           .pivot_table(index="장애 대분류", columns="연월라벨", values="일자",
                        aggfunc="size", fill_value=0)
           .reindex(fy).sort_index(axis=1))
    fig = ch.heatmap(mat, PAL, xtitle="", ytitle="")
    ch.block("장애유형 × 월 히트맵 (상위 8종)", fig, mat.reset_index(), "fhm",
             caption="특정 유형이 몰린 시기를 찾습니다. 색이 진할수록 건수가 많습니다.")


# ── 3. 지역·고속사 ───────────────────────────────────────────────────
with tabs[4]:
    left, right = st.columns(2)
    with left:
        t = dl.top_n(df, "처리지역", 12)
        ch.block("처리지역별 건수", ch.hbar(t, "처리지역", "건수", PAL), t, "region")
    with right:
        t = dl.top_n(df, "처리거점", 12)
        ch.block("처리거점별 건수", ch.hbar(t, "처리거점", "건수", PAL), t, "base")

    st.divider()
    all_counts = df["고속사"].value_counts().rename_axis("고속사").reset_index(name="건수")
    all_counts["비중"] = (all_counts["건수"] / len(df) * 100).round(1)
    if FLEET_SIZES:
        rated = dl.with_fleet_rate(all_counts, "고속사", FLEET_SIZES)
        unrated = rated.loc[rated["차량대수"].isna(), "고속사"].tolist()
        top_rate = (rated.dropna(subset=["차량대수"])
                         .sort_values("대당건수", ascending=False).head(15))
        cap = ("차량 1대당 발생 건수 = 전체 기간 누적 장애 건수 ÷ 고속사별 차량대수(제공 파일 "
              "기준) - 단순 건수가 아니라 차량 규모를 반영한 비율입니다.")
        if unrated:
            cap += f" 차량대수 미확인: {', '.join(unrated)}(비율 계산 제외, 표에는 건수만 있음)."
        ch.block(
            "고속사별 장애 발생률 (차량 1대당) Top 15",
            ch.hbar(top_rate, "고속사", "대당건수", PAL, suffix="건/대",
                   text_fmt="{:.2f}", hover_fmt="%{x:,.2f}"),
            rated.sort_values("대당건수", ascending=False, na_position="last"),
            "company", caption=cap,
        )
    else:
        ch.block("고속사별 장애 건수 Top 15",
                 ch.hbar(all_counts.head(15), "고속사", "건수", PAL),
                 all_counts, "company",
                 caption="`고속사별차량대수.txt`를 이 폴더에 두면 차량 1대당 발생률로 볼 수 있습니다.")

    st.divider()
    comps = df["고속사"].value_counts().head(10).index.tolist()
    faults = df["장애 대분류"].value_counts().head(8).index.tolist()
    mat = (df[df["고속사"].isin(comps) & df["장애 대분류"].isin(faults)]
           .pivot_table(index="고속사", columns="장애 대분류", values="일자",
                        aggfunc="size", fill_value=0)
           .reindex(index=comps, columns=faults, fill_value=0))
    ch.block("고속사 × 장애유형 히트맵", ch.heatmap(mat, PAL), mat.reset_index(), "chm",
             caption="사업자별로 유독 잦은 증상이 있는지 봅니다 (상위 10사 × 상위 8종).")


# ── 4. 차량·단말기 ───────────────────────────────────────────────────
with tabs[5]:
    veh = (df.groupby("차량번호", as_index=False)
             .agg(건수=("일자", "size"), 최초=("일자", "min"), 최근=("일자", "max"),
                  고속사=("고속사", "first"))
             .sort_values("건수", ascending=False))

    c = st.columns(4)
    c[0].metric("장애 발생 차량", f"{len(veh):,}대")
    c[1].metric("차량당 평균", f"{veh['건수'].mean():.2f}건")
    multi_v = int((veh["건수"] >= 3).sum())
    c[2].metric("3회 이상 반복 차량", f"{multi_v:,}대",
                delta=f"전체 차량의 {multi_v / len(veh) * 100:.0f}%",
                delta_color="off", delta_arrow="off")
    c[3].metric("반복 차량 발생 건수",
                f"{int(veh.loc[veh['건수'] >= 3, '건수'].sum()):,}건",
                delta=f"전체의 {veh.loc[veh['건수'] >= 3, '건수'].sum() / len(df) * 100:.0f}%",
                delta_color="off", delta_arrow="off")
    st.caption(
        f"사용자 제공 추정치 기준 운영 차량은 약 {FLEET_SIZE_EST:,}대(고속사별 매달 대수 변동) — "
        f"선택 기간에 장애가 접수된 {len(veh):,}대는 그 약 {len(veh) / FLEET_SIZE_EST * 100:.0f}%에 "
        "해당합니다. 정확한 전체 대수가 아니라 대략치라 참고용으로만 보세요."
    )

    st.divider()
    left, right = st.columns([3, 2])
    with left:
        top_v = veh.head(20).copy()
        top_v["라벨"] = top_v["차량번호"].astype(str) + " · " + top_v["고속사"].astype(str)
        ch.block("반복 장애 차량 Top 20",
                 ch.hbar(top_v[["라벨", "건수"]], "라벨", "건수", PAL),
                 top_v[["차량번호", "고속사", "건수", "최초", "최근"]], "veh",
                 caption="같은 차량이 반복 출동을 만드는지 — 교체보다 근본원인 점검 대상입니다.")
    with right:
        bins = pd.cut(veh["건수"], [0, 1, 2, 3, 5, 10, 10**6],
                      labels=["1회", "2회", "3회", "4~5회", "6~10회", "11회+"])
        dist = bins.value_counts().reindex(["1회", "2회", "3회", "4~5회", "6~10회", "11회+"])
        dist = dist.rename_axis("발생횟수").reset_index(name="차량수")
        fig = ch.hbar(dist, "발생횟수", "차량수", PAL, height=340, suffix="대")
        fig.update_yaxes(categoryorder="array",
                         categoryarray=["11회+", "6~10회", "4~5회", "3회", "2회", "1회"])
        ch.block("차량별 장애 발생 횟수 분포", fig, dist, "vdist")

    st.divider()
    st.markdown("**교체 단말기 재고장 추적**")
    st.caption("`교체전` 칸에 같은 S/N이 여러 번 등장하면, 회수·재투입된 단말기가 "
               "다시 고장 났을 가능성이 있습니다.")
    sn = df["교체전"].dropna()
    sn = sn[sn.str.len() >= 6]
    if sn.empty:
        st.info("이 구간에는 단말기 S/N 기록이 없습니다.")
    else:
        vc = sn.value_counts()
        rep_sn = vc[vc >= 2]
        c = st.columns(3)
        c[0].metric("S/N 기록 건수", f"{len(sn):,}건")
        c[1].metric("고유 S/N", f"{vc.size:,}개")
        c[2].metric("2회 이상 회수 S/N", f"{rep_sn.size:,}개",
                    delta=f"해당 {int(rep_sn.sum()):,}건",
                    delta_color="off", delta_arrow="off")
        if rep_sn.size:
            detail = (df[df["교체전"].isin(rep_sn.head(20).index)]
                      [["교체전", "일자", "차량번호", "고속사", "장애 대분류", "처리 대분류", "교체후"]]
                      .sort_values(["교체전", "일자"]))
            st.dataframe(detail, hide_index=True, height=320)
            st.download_button("CSV 내려받기", detail.to_csv(index=False).encode("utf-8-sig"),
                               file_name="재고장_SN.csv", mime="text/csv", key="d_sn")


# ── 5. 자재·담당자 ───────────────────────────────────────────────────
with tabs[6]:
    parts = (df[dl.PART_COLS].sum().rename_axis("자재").reset_index(name="수량"))
    parts = parts[parts["수량"] > 0].sort_values("수량", ascending=False)

    left, right = st.columns(2)
    with left:
        st.markdown("**자재 사용량**")
        if parts.empty:
            st.info("이 구간에는 자재 사용 수량이 입력되어 있지 않습니다. "
                    "(원본의 GPS안테나~베이스플레이트 칸이 비어 있음)")
        else:
            st.caption(f"총 {int(parts['수량'].sum()):,}개 · {len(parts)}종")
            st.plotly_chart(ch.hbar(parts, "자재", "수량", PAL, suffix="개"),
                            key="c_parts", theme=None,
                            config={"displayModeBar": False, "responsive": True})
            st.download_button("CSV 내려받기", parts.to_csv(index=False).encode("utf-8-sig"),
                               file_name="자재사용량.csv", mime="text/csv", key="d_parts")

    with right:
        t = dl.top_n(df, "처리담당자", 15)
        ch.block("담당자별 처리 건수", ch.hbar(t, "처리담당자", "건수", PAL), t, "staff",
                 caption="현장 처리 인원별 부하를 봅니다.")

    st.divider()
    staff = df["처리담당자"].value_counts().head(10).index.tolist()
    bases = df["처리거점"].value_counts().head(8).index.tolist()
    mat = (df[df["처리담당자"].isin(staff) & df["처리거점"].isin(bases)]
           .pivot_table(index="처리담당자", columns="처리거점", values="일자",
                        aggfunc="size", fill_value=0)
           .reindex(index=staff, columns=bases, fill_value=0))
    ch.block("담당자 × 거점 배치 현황", ch.heatmap(mat, PAL), mat.reset_index(), "shm",
             caption="담당자가 어느 거점을 맡고 있는지 — 인력 재배치 검토용.")


# ── 6. 원본 데이터 ───────────────────────────────────────────────────
with tabs[7]:
    st.markdown("**필터가 적용된 원본 데이터**")
    cols = ["일자", "주차라벨", "접수구분", "처리지역", "처리거점", "고속사", "차량번호",
            "단말기 구분", "장애 대분류", "처리 대분류", "장애 접수 내용",
            "세부 조치 내용 (현장 확인 증상 / 조치내용)", "처리담당자", "교체전", "교체후",
            "원본파일"]
    view = df[cols].sort_values("일자", ascending=False)
    st.caption(f"{len(view):,}행")
    st.dataframe(view, hide_index=True, height=560,
                 column_config={"일자": st.column_config.DateColumn("일자", format="YYYY-MM-DD")})
    st.download_button("전체 CSV 내려받기", view.to_csv(index=False).encode("utf-8-sig"),
                       file_name="고속Rawdata_필터결과.csv", mime="text/csv", key="d_raw")

    st.divider()
    with st.expander("표기 정규화 내역 — 원본의 오타·공백을 이렇게 통일했습니다"):
        st.markdown("집계가 갈라지지 않도록 아래 값을 통합했습니다. 원본 파일은 수정하지 않습니다.")
        rows = []
        for label, fix in [("처리지역", dl.REGION_FIX), ("처리거점", dl.BASE_FIX),
                           ("고속사", dl.COMPANY_FIX), ("처리담당자", dl.STAFF_FIX),
                           ("처리 대분류", dl.ACTION_FIX)]:
            for k, v in fix.items():
                rows.append({"컬럼": label, "원본 표기": k, "통일 표기": v})
        st.dataframe(pd.DataFrame(rows), hide_index=True, height=340)
        st.caption("이 밖에 모든 텍스트 칸의 앞뒤 공백과 중복 공백을 제거했고, "
                   "월·연도는 원본 '월' 칸 대신 `일자`에서 다시 계산합니다.")
