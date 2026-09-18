"""고속 장애실적 Rawdata 로더.

`고속 시외 YYYY년.xlsm` 파일들의 `고속Rawdata` 시트를 읽어 하나로 합치고,
수기 입력 과정에서 생긴 표기 흔들림(공백/오타)을 정규화한다.

시트 규격
    3행 = 헤더, 4행부터 데이터 (pandas 기준 header=2)
    B열(No.) ~ AF열(베이스플레이트)

파일은 매주 갱신되므로 캐시 키에 파일의 수정시각·크기를 넣어
새 파일이 저장되면 자동으로 다시 읽도록 한다.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import streamlit as st

SHEET = "고속Rawdata"
HEADER_ROW = 2  # 0-based -> 엑셀 3행

# 개선활동 시행일 - 대시보드(app.py)와 PPT 보고서(report_builder.py)가 같은 값을 보도록
# 여기 한 곳에 둔다.
REPAIR_POINT_START = pd.Timestamp("2026-02-05")     # 수리센터 수리포인트 적용 시작
FIRMWARE_DEPLOY_DATE = pd.Timestamp("2026-08-06")   # AFC-승차단말 통신보완 펌웨어 배포일
FIRMWARE_TARGET_FAULTS = ["승차연결지연", "승차단말기통신불량"]  # 이 펌웨어가 겨냥하는 장애

# 자재(부품) 사용량 컬럼
PART_COLS = [
    "GPS안테나", "AUDIO케이블", "VIDEO케이블", "행선 케이블", "Y케이블", "AFC케이블",
    "BMS케이블", "젠터(AUDIO)", "쿠션", "LTE,Wifi안테나", "심지봉",
    "육각볼트(와셔포함)", "베이스플레이트",
]

DIM_COLS = ["접수구분", "처리지역", "처리거점", "고속사", "단말기 구분",
            "장애 대분류", "처리 대분류", "처리담당자"]

TEXT_COLS = ["장애 접수 내용", "세부 조치 내용 (현장 확인 증상 / 조치내용)"]

# ── 표기 정규화 사전 ────────────────────────────────────────────────
# 처리지역: 'OO터미널' 표기를 정본으로 삼는다(원본에서 우세한 형태).
REGION_FIX = {
    "광주터민널": "광주터미널", "세트럴": "센트럴", "센터럴": "센트럴",
    "센트를": "센트럴", "센틀럴": "센트럴",
    "광주": "광주터미널", "부산": "부산터미널", "인천": "인천터미널",
    "성남": "성남터미널", "수원": "수원터미널", "동대구": "동대구터미널",
}

# 처리거점: '경부선'은 지역명이 거점 칸에 잘못 들어간 경우 -> 서울
BASE_FIX = {"인전": "인천", "경부선": "서울"}

# 고속사: 지역명이 잘못 들어간 값은 회사로 집계되지 않도록 분리
COMPANY_FIX = {"금혹고속": "금호고속", "경부선": "(미확인)", "동서울": "(미확인)"}

# 처리담당자: 1건짜리 오타를 실제 담당자로 되돌린다
# (김선곤=광주 1건 -> 광주 전담 김성곤. 김선웅은 별도 인물일 수 있어 건드리지 않음)
STAFF_FIX = {"이경핀": "이경필", "박벙태": "박병태", "김선곤": "김성곤"}

# 처리 대분류: 띄어쓰기/표기 변형 통합
ACTION_FIX = {
    "GPS 안테나 교체": "GPS안테나교체", "GPS안테나 교체": "GPS안테나교체",
    "LCD 교체": "LCD교체", "LTE 모뎀교체": "LTE모뎀교체",
    "바코드리더기 교체": "바코드리더기교체", "바코드리더기 기구물교체": "바코드리더기기구물교체",
    "승차단말기 교체": "승차단말기교체", "운전단말기교체": "운전자단말기교체",
    "외장형모뎀교체": "외장모뎀교체", "AFC I/O 케이블 교체": "AFC I/O케이블교체",
}

_WS = re.compile(r"\s+")
_WEEK = re.compile(r"^(\d{1,2})월\s*(\d{1,2})주$")


def discover_files(base_dir: str | Path = ".") -> list[Path]:
    """`고속 시외 ____년.xlsm` 파일을 찾는다. 엑셀 임시 잠금파일(~$)은 제외."""
    base = Path(base_dir)
    files = [p for p in base.glob("*.xlsm") if not p.name.startswith("~$")]
    return sorted(files, key=lambda p: p.name)


def fingerprint(paths: list[Path]) -> tuple:
    """파일이 갱신되면 캐시가 자동으로 무효화되도록 하는 키."""
    out = []
    for p in paths:
        try:
            s = p.stat()
            out.append((p.name, int(s.st_mtime), s.st_size))
        except OSError:
            out.append((p.name, 0, 0))
    return tuple(out)


def _clean_text(s: pd.Series) -> pd.Series:
    """앞뒤 공백 제거 + 연속 공백을 1칸으로."""
    return (s.astype("string")
             .str.replace(_WS, " ", regex=True)
             .str.strip())


def _business_week(dates: pd.Series, labels: pd.Series) -> tuple[pd.Series, pd.Series]:
    """원본 '주차' 라벨(예: 3월2주)에 연도를 붙여 정렬 가능한 주차 키를 만든다.

    이 업무의 주차는 목요일에 시작하므로 12월 말 날짜가 '1월1주'로 잡힌다.
    날짜의 연도만 쓰면 연말/연초가 뒤섞이므로 월-라벨 불일치를 보정한다.
    """
    m = labels.str.extract(_WEEK)
    wk_month = pd.to_numeric(m[0], errors="coerce")
    wk_no = pd.to_numeric(m[1], errors="coerce")
    d_month = dates.dt.month
    year = dates.dt.year.astype("float")
    year = year.where(~((d_month == 12) & (wk_month == 1)), year + 1)
    year = year.where(~((d_month == 1) & (wk_month == 12)), year - 1)

    key = (year.astype("Int64").astype("string")
           + "-" + wk_month.astype("Int64").astype("string").str.zfill(2)
           + "-" + wk_no.astype("Int64").astype("string").str.zfill(2))
    disp = (year.astype("Int64").astype("string").str[2:] + "년 "
            + wk_month.astype("Int64").astype("string") + "월"
            + wk_no.astype("Int64").astype("string") + "주")
    # 라벨이 깨진 행은 날짜 기준 주(월요일 시작)로 대체
    bad = key.isna()
    if bad.any():
        wk_start = dates.dt.to_period("W-WED").dt.start_time
        key = key.mask(bad, wk_start.dt.strftime("%Y-%m-%d"))
        disp = disp.mask(bad, wk_start.dt.strftime("%y년 %m/%d주"))
    return key, disp


def _read_one(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET, header=HEADER_ROW, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df["원본파일"] = path.name
    return df


def _normalize_field(df: pd.DataFrame) -> pd.DataFrame:
    """concat한 원시 프레임에 정규화·파생 컬럼을 적용한다.

    경로로 읽은 것이든 업로드로 읽은 것이든 이 한 함수를 거치면 동일한 결과가 되도록,
    표기 정규화·파생 컬럼 계산 로직을 여기 한 곳에만 둔다.
    """
    df = df.copy()
    df["일자"] = pd.to_datetime(df["일자"], errors="coerce")
    df = df[df["일자"].notna()].copy()

    for c in DIM_COLS + TEXT_COLS + ["주차", "차량번호", "교체전", "교체후"]:
        if c in df.columns:
            df[c] = _clean_text(df[c])

    df["처리지역"] = df["처리지역"].replace(REGION_FIX)
    df["처리거점"] = df["처리거점"].replace(BASE_FIX)
    df["고속사"] = df["고속사"].replace(COMPANY_FIX)
    df["처리담당자"] = df["처리담당자"].fillna("(미기재)").replace(STAFF_FIX)
    df["처리 대분류"] = df["처리 대분류"].replace(ACTION_FIX)

    df["연도"] = df["일자"].dt.year
    df["월"] = df["일자"].dt.month                      # 원본 '월' 문자열 대신 날짜에서 유도
    df["연월"] = df["일자"].dt.to_period("M").dt.to_timestamp()
    df["연월라벨"] = df["일자"].dt.strftime("%Y-%m")
    df["요일"] = df["일자"].dt.dayofweek                 # 0=월
    df["주차키"], df["주차라벨"] = _business_week(df["일자"], df["주차"])

    for c in PART_COLS:
        df[c] = pd.to_numeric(df.get(c), errors="coerce").fillna(0.0)
    df["자재사용합계"] = df[PART_COLS].sum(axis=1)

    df["단말기교체"] = df["처리 대분류"].str.contains("단말기교체", na=False)

    df = df.sort_values("일자", ignore_index=True)
    return df


@st.cache_data(show_spinner="엑셀 파일을 읽는 중…", persist="disk")
def load_data(paths: tuple[str, ...], _fp: tuple) -> pd.DataFrame:
    """폴더에 있는 여러 연도 파일을 읽어 정규화된 하나의 데이터프레임으로 반환.

    `_fp`는 캐시 무효화용 지문이며 함수 안에서는 쓰지 않는다.
    """
    frames = [_read_one(Path(p)) for p in paths]
    return _normalize_field(pd.concat(frames, ignore_index=True))


@st.cache_data(show_spinner="업로드한 장애접수 엑셀을 읽는 중…")
def load_data_from_uploads(files) -> pd.DataFrame:
    """사이드바에서 업로드한 장애접수 엑셀(최대 2개)을 읽어 정규화한다.

    업로드 파일은 세션 메모리에만 있어 디스크 캐시(persist)는 쓰지 않는다 -
    브라우저 업로드 위젯의 파일 객체(UploadedFile)는 내용으로 해시되므로,
    같은 파일을 다시 올리지 않는 한 재계산되지 않는다.
    """
    frames = []
    for f in files:
        f.seek(0)
        df = pd.read_excel(f, sheet_name=SHEET, header=HEADER_ROW, engine="openpyxl")
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        df["원본파일"] = getattr(f, "name", "업로드파일")
        frames.append(df)
    return _normalize_field(pd.concat(frames, ignore_index=True))


def weekly_counts(df: pd.DataFrame) -> pd.DataFrame:
    """주차별 건수 (주차키로 정렬, 표시는 주차라벨)."""
    g = (df.groupby(["주차키", "주차라벨"], as_index=False)
           .agg(건수=("일자", "size"), 시작일=("일자", "min"), 종료일=("일자", "max"))
           .sort_values("주차키", ignore_index=True))
    return g


def top_n(df: pd.DataFrame, col: str, n: int = 10, other: bool = False) -> pd.DataFrame:
    """빈도 상위 n개. other=True면 나머지를 '기타'로 묶는다."""
    vc = df[col].fillna("(미기재)").value_counts()
    head = vc.head(n)
    out = head.rename_axis(col).reset_index(name="건수")
    if other and len(vc) > n:
        rest = int(vc.iloc[n:].sum())
        out = pd.concat([out, pd.DataFrame({col: ["기타"], "건수": [rest]})], ignore_index=True)
    return out


# ═══════════════════════════════════════════════════════════════════
# E-PASS 고속수리 (수리센터 재불량 로그) - 개선활동 효과 분석용
#
# `E-PASS고속수리_YYYY.xlsx`의 'B610 고속 운전자 재불량' 시트는 개별 단말기
# (S/N 단위)의 수리 이력을 담고 있다. 이 S/N은 고속Rawdata의 `교체전`/`교체후`
# 칸과 같은 체계라서, 두 데이터를 S/N으로 이어 붙이면
#   "수리센터에서 고친 단말기가 현장에 재투입된 뒤 다시 고장 났는가"
# 를 직접 추적할 수 있다.
# ═══════════════════════════════════════════════════════════════════

REPAIR_SHEETS = ["B610 고속 운전자 재불량", "B610 고속 승차 재불량"]

# 증상 텍스트(접수+실장애 내역)를 장애실적 쪽 관심 장애유형과 맞춰 분류한다.
# 순서가 우선순위 - 위에서부터 먼저 매치되는 것으로 정한다.
_REPAIR_CATEGORY_RULES = [
    ("LTE/모뎀", r"LTE|모뎀|접속\s*지연|연결\s*지연|끊김"),
    ("BMS", r"BMS"),
    ("승차연결", r"승차.*(연결|접속)"),
    ("화면/터치", r"화면|LCD|터치|멈춤"),
    ("바코드", r"바코드"),
    ("전원/부팅", r"전원|부팅"),
]


def _classify_repair(text: pd.Series) -> pd.Series:
    out = pd.Series("기타", index=text.index)
    remaining = pd.Series(True, index=text.index)
    for name, pat in _REPAIR_CATEGORY_RULES:
        hit = remaining & text.str.contains(pat, regex=True, na=False)
        out[hit] = name
        remaining &= ~hit
    return out


def discover_repair_files(base_dir: str | Path = ".") -> list[Path]:
    """`E-PASS고속수리_*.xlsx` 파일을 찾는다. 엑셀 임시 잠금파일(~$)은 제외."""
    base = Path(base_dir)
    files = [p for p in base.glob("E-PASS고속수리*.xlsx") if not p.name.startswith("~$")]
    return sorted(files, key=lambda p: p.name)


def _read_repair_sheet(path: Path, sheet: str) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=0, engine="openpyxl")
    except (ValueError, KeyError):
        return None  # 시트가 없는 버전의 파일일 수 있다
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df[df["S/N"].notna()].copy()
    df["기기"] = "운전자" if "운전자" in sheet else "승차"
    df["원본파일"] = path.name
    return df


def _normalize_repair(df: pd.DataFrame) -> pd.DataFrame:
    """concat한 원시 수리 로그 프레임에 정규화·파생 컬럼을 적용한다(경로/업로드 공통)."""
    df = df.copy()
    df["수리일"] = pd.to_datetime(df["수 리 일"], errors="coerce")
    df = df[df["수리일"].notna()].copy()

    # S/N은 필드 데이터의 교체전/교체후와 같은 9자리 체계 - 문자열로 맞춘다
    df["SN"] = pd.to_numeric(df["S/N"], errors="coerce")
    df = df[df["SN"].notna()].copy()
    df["SN"] = df["SN"].astype("Int64").astype(str)

    for c in ["월", "주차", "접수 장애 내역", "실 장애 내역", "주 수리 내역",
              "상세수리내역", "수리의견", "수리/NDF"]:
        if c in df.columns:
            df[c] = _clean_text(df[c])

    txt = df["접수 장애 내역"].fillna("") + " " + df["실 장애 내역"].fillna("")
    df["카테고리"] = _classify_repair(txt)

    df["재불량여부"] = df["재불량"].fillna("").astype(str).str.strip().eq("재불량")
    df["동일재불량여부"] = df["동일재불량"].fillna("").astype(str).str.strip().eq("재불량")

    df["연월"] = df["수리일"].dt.to_period("M").dt.to_timestamp()
    df["연월라벨"] = df["수리일"].dt.strftime("%Y-%m")
    df["주차키"], df["주차라벨"] = _business_week(df["수리일"], df["주차"])

    df = df.sort_values("수리일", ignore_index=True)
    return df


@st.cache_data(show_spinner="수리 로그를 읽는 중…", persist="disk")
def load_repair_data(paths: tuple[str, ...], _fp: tuple) -> pd.DataFrame:
    """폴더의 수리센터 재불량 로그를 읽어 정규화한다. `_fp`는 캐시 무효화용 지문."""
    frames = []
    for p in paths:
        for sheet in REPAIR_SHEETS:
            d = _read_repair_sheet(Path(p), sheet)
            if d is not None and not d.empty:
                frames.append(d)
    if not frames:
        return pd.DataFrame()
    return _normalize_repair(pd.concat(frames, ignore_index=True))


def _read_repair_sheet_upload(f, sheet: str) -> pd.DataFrame | None:
    try:
        f.seek(0)
        df = pd.read_excel(f, sheet_name=sheet, header=0, engine="openpyxl")
    except (ValueError, KeyError):
        return None  # 시트가 없는 버전의 파일일 수 있다
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df[df["S/N"].notna()].copy()
    df["기기"] = "운전자" if "운전자" in sheet else "승차"
    df["원본파일"] = getattr(f, "name", "업로드파일")
    return df


@st.cache_data(show_spinner="업로드한 수리현황 엑셀을 읽는 중…")
def load_repair_from_uploads(files) -> pd.DataFrame:
    """사이드바에서 업로드한 수리현황(E-PASS) 엑셀을 읽어 정규화한다."""
    frames = []
    for f in files:
        for sheet in REPAIR_SHEETS:
            d = _read_repair_sheet_upload(f, sheet)
            if d is not None and not d.empty:
                frames.append(d)
    if not frames:
        return pd.DataFrame()
    return _normalize_repair(pd.concat(frames, ignore_index=True))


def build_recurrence(rep: pd.DataFrame, field: pd.DataFrame) -> pd.DataFrame:
    """각 수리 건에 대해, 수리일 이후 그 S/N이 현장에서 다시 `교체전`(고장 이탈)으로
    잡히는 가장 이른 시점을 찾아 붙인다. `merge_asof`로 SN별 다음 사건을 찾는 방식이라
    수리 건수가 늘어도 빠르다.

    반환 컬럼 추가: 재발여부, 재발일, 경과일(수리~재발, 일), 재발장애(장애 대분류)
    """
    if rep.empty:
        return rep.assign(재발여부=pd.Series(dtype=bool), 재발일=pd.NaT,
                          경과일=pd.Series(dtype=float), 재발장애=None)

    # merge_asof는 조인 키의 dtype이 정확히 같아야 한다 - 두 쪽 모두 확장 문자열
    # dtype(string[python]/string[pyarrow])일 수 있어, 순수 object/str로 맞춘다.
    # merge_asof는 `by` 그룹별이 아니라 `on` 컬럼 전체가 오름차순 정렬되어 있어야 한다.
    left = rep[["SN", "수리일"]].reset_index().sort_values("수리일")
    left["SN"] = left["SN"].astype(object).map(str)
    right = (field.dropna(subset=["교체전"])[["교체전", "일자", "장애 대분류"]]
             .rename(columns={"교체전": "SN"})
             .sort_values("일자"))
    right["SN"] = right["SN"].astype(object).map(str)

    merged = pd.merge_asof(left, right, left_on="수리일", right_on="일자", by="SN",
                           direction="forward", allow_exact_matches=False)
    merged = merged.set_index("index").reindex(rep.index)

    out = rep.copy()
    out["재발여부"] = merged["일자"].notna()
    out["재발일"] = merged["일자"]
    out["경과일"] = (merged["일자"] - out["수리일"]).dt.days
    out["재발장애"] = merged["장애 대분류"]
    return out


# ═══════════════════════════════════════════════════════════════════
# 현장 개선활동 이벤트 (세부 조치 내용 텍스트에서 추출)
#
# `세부 조치 내용 (현장 확인 증상 / 조치내용)` 칸의 자유 텍스트에 아래 패턴이
# 박혀 있다 - 사용자 확인:
#   '외장모뎀적용'      = 이번 조치로 외장 LTE모뎀을 새로 장착 (신규 전환)
#   '외장모뎀적용차량'  = 이 차량은 이미 외장모뎀이 적용된 상태라는 맥락 표기
#                        (이번 조치와 무관 - 신규 전환이 아니다)
#   'B/D교체적용'       = 운전자단말기 내 AFC보드·BMS보드 중 하나 이상을 교체
# 오타 변형(적용→적요, 차량→차랑)도 실제 데이터에 있어 함께 잡는다.
# ═══════════════════════════════════════════════════════════════════

ACTION_TEXT_COL = TEXT_COLS[1]

_EXT_NEW_RE = r"외장모뎀적[용요](?!차[량랑])"
_EXT_ALREADY_RE = r"외장모뎀적[용요]차[량랑]"
_EXT_REMOVED_RE = r"외장모뎀철거"
_BD_REPLACED_RE = r"B/D교체적용"


def extract_intervention_events(df: pd.DataFrame,
                                text_col: str = ACTION_TEXT_COL) -> pd.DataFrame:
    """세부 조치 내용에서 개선활동 이벤트 플래그를 뽑아 원본에 컬럼으로 붙인다."""
    txt = df[text_col].fillna("")
    out = df.copy()
    out["외장모뎀_신규전환"] = txt.str.contains(_EXT_NEW_RE, regex=True, na=False)
    out["외장모뎀_기적용"] = txt.str.contains(_EXT_ALREADY_RE, regex=True, na=False)
    out["외장모뎀_철거"] = txt.str.contains(_EXT_REMOVED_RE, na=False)
    out["BD보드_교체"] = txt.str.contains(_BD_REPLACED_RE, na=False)
    out["BD보드_단독"] = out["BD보드_교체"] & ~out["외장모뎀_신규전환"] & ~out["외장모뎀_기적용"]
    return out


def vehicle_event_study(df: pd.DataFrame, event_mask: pd.Series, *, focus_faults: list[str],
                        window_days: int = 90, buffer_days: int = 14) -> pd.DataFrame:
    """차량별로 개선조치 최초 시행일 기준 전/후 `window_days`일 동안 관심 장애가 몇 건
    발생했는지 비교한다(차량번호 기준 - 이 이벤트들은 S/N보다 차량번호가 항상 채워져 있다).

    - 이벤트 이후 `buffer_days`일은 집계에서 뺀다 - 설치 직후 며칠간의 재작업은 같은
      방문의 연장선일 뿐, '수리 후 새로 재발'한 것으로 보기 어렵다.
    - 이벤트 이후 `window_days`를 다 채우지 못한(최근에 조치한) 차량은 제외한다
      (우측 절단 방지 - E-PASS 수리 재발률과 같은 원리).
    - 한 차량에 이벤트가 여러 번 있으면 가장 이른 시점만 기준으로 삼는다.
    """
    last_date = df["일자"].max()
    events = (df.loc[event_mask, ["일자", "차량번호"]]
                .rename(columns={"일자": "이벤트일"})
                .sort_values("이벤트일")
                .drop_duplicates("차량번호", keep="first"))
    events = events[events["이벤트일"] <= last_date - pd.Timedelta(days=window_days)]

    focus = df.loc[df["장애 대분류"].isin(focus_faults), ["차량번호", "일자"]]
    rows = []
    for veh, t0 in zip(events["차량번호"], events["이벤트일"]):
        d = focus.loc[focus["차량번호"] == veh, "일자"]
        before = int(((d >= t0 - pd.Timedelta(days=window_days)) & (d < t0)).sum())
        after = int(((d > t0 + pd.Timedelta(days=buffer_days))
                     & (d <= t0 + pd.Timedelta(days=window_days))).sum())
        rows.append((veh, t0, before, after))
    return pd.DataFrame(rows, columns=["차량번호", "이벤트일", "전", "후"])


def before_after_monthly_rate(df: pd.DataFrame, target_faults: list[str],
                              event_date: pd.Timestamp, *,
                              pre_window_days: int = 180) -> dict:
    """전 차량에 같은 날 일괄 적용된 조치(펌웨어 배포 등)의 전/후를 월 환산 발생률로
    비교한다 - 차량별 대조군이 없는 조치는 `vehicle_event_study`처럼 차량별 전/후를
    비교할 수 없어, 전체(모집단) 수준의 시계열로 본다.

    - 배포 전: `event_date` 기준 최근 `pre_window_days`일(기본 180일=6개월)의 건수를
      30일 기준으로 환산.
    - 배포 후: `event_date`부터 데이터 최신일까지 실제 경과일 기준으로 환산 - 최신일이
      배포일 이후여야 계산되며(우측 절단 방지), 관찰 기간이 짧을수록 표본이 작아
      비율이 크게 흔들릴 수 있다.
    """
    sub = df[df["장애 대분류"].isin(target_faults)]
    last_date = df["일자"].max()
    days_since = (last_date - event_date).days

    before_win = sub[(sub["일자"] >= event_date - pd.Timedelta(days=pre_window_days))
                     & (sub["일자"] < event_date)]
    after_win = sub[sub["일자"] >= event_date] if days_since >= 0 else sub.iloc[0:0]

    before_rate = len(before_win) / pre_window_days * 30
    after_rate = (len(after_win) / days_since * 30) if days_since > 0 else 0.0

    return {
        "days_since": days_since,
        "pre_window_days": pre_window_days,
        "before_count": len(before_win),
        "after_count": len(after_win),
        "before_rate": before_rate,
        "after_rate": after_rate,
    }
