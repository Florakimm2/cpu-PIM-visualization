"""
PIM 출원 동향 시각화 자동화 스크립트 - 디자인 개선 반영본

기능
1. 엑셀 데이터 자동 로드
2. 국가코드 / 출원일 / 출원인 기준 집계
3. 여러 그래프를 반복문으로 대량 생성
4. style.py 기반 디자인 테마 통일
5. 3D 그래프 생성
6. PNG 이미지와 PDF 리포트 저장

실행 예시
python pim_trend_visualizer.py --input "전데검완_등록.xlsx" --output "pim_visual_outputs"
python pim_trend_visualizer.py --input "전데검완_등록.xlsx" --alias "출원인_삼성SK 출원인 중복.xlsx" --output "pim_visual_outputs"
"""

from __future__ import annotations

import argparse
import re
import warnings
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_pdf import PdfPages
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401  # 3D projection 등록용

from style import apply_report_style, polish_axes
import config

warnings.filterwarnings("ignore", category=UserWarning)


# ============================================================
# 1. 공통 유틸
# ============================================================

def safe_filename(text: str, max_len: int = 80) -> str:
    """파일명으로 쓰기 어려운 문자를 정리한다."""
    text = str(text).strip()
    text = re.sub(r"[\\/:*?\"<>|]", "_", text)
    text = re.sub(r"\s+", "_", text)
    return text[:max_len] if text else "unknown"


def set_korean_font() -> Optional[str]:
    """한국어가 깨지지 않도록 사용 가능한 한글 폰트를 자동 선택한다."""
    candidates = [
        "AppleGothic",       # macOS
        "Malgun Gothic",     # Windows
        "NanumGothic",       # Linux/Colab
        "Noto Sans CJK KR",
        "Noto Sans KR",
        "Arial Unicode MS",
    ]
    available = {font.name for font in fm.fontManager.ttflist}
    selected = next((font for font in candidates if font in available), None)

    if selected:
        plt.rcParams["font.family"] = selected
    plt.rcParams["axes.unicode_minus"] = False
    return selected


def apply_theme() -> None:
    """
    전체 그래프의 디자인을 통일한다.

    기존 config 기반 테마 대신 style.py의 apply_report_style()를 우선 적용한다.
    다만 style.py에서 AppleGothic을 고정해둔 경우를 대비해, 사용 가능한 한글 폰트를
    다시 찾아서 rcParams에 덮어쓴다.
    """
    font_name = set_korean_font()

    try:
        apply_report_style()
    except Exception:
        # style.py가 수정 중이거나 일부 환경에서 실패해도 기본 그래프는 생성되게 한다.
        sns.set_theme(
            style=getattr(config, "SEABORN_STYLE", "whitegrid"),
            context=getattr(config, "SEABORN_CONTEXT", "talk"),
        )

    if font_name:
        plt.rcParams["font.family"] = font_name

    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = getattr(config, "FIG_DPI", 180)
    plt.rcParams["savefig.dpi"] = getattr(config, "FIG_DPI", 180)
    plt.rcParams["axes.prop_cycle"] = plt.cycler(color=sns.color_palette("deep"))


def save_figure(
    fig: plt.Figure,
    output_dir: Path,
    filename: str,
    pdf: Optional[PdfPages] = None,
) -> None:
    """PNG 저장 + PDF에 페이지 추가."""
    output_dir.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()

    if config.SAVE_PNG:
        fig.savefig(output_dir / f"{filename}.png", bbox_inches="tight")

    if config.SAVE_PDF and pdf is not None:
        pdf.savefig(fig, bbox_inches="tight")

    plt.close(fig)


def annotate_bar_values(ax: plt.Axes, fmt: str = "{:.0f}") -> None:
    """막대 그래프에 수치 라벨을 붙인다."""
    for container in ax.containers:
        labels = []
        for value in container.datavalues:
            if pd.isna(value):
                labels.append("")
            else:
                labels.append(fmt.format(value))
        ax.bar_label(container, labels=labels, padding=3, fontsize=9)


def shorten_label(text: object, max_len: int = 32) -> str:
    """축 라벨이 너무 길 때 적당히 줄인다."""
    value = str(text)
    return value if len(value) <= max_len else value[:max_len - 1] + "…"


def add_horizontal_bar_labels(
    ax: plt.Axes,
    values: pd.Series,
    total: Optional[float] = None,
    suffix: str = "건",
) -> None:
    """가로 막대 오른쪽에 값 또는 비율 라벨을 붙인다."""
    max_value = float(values.max()) if len(values) else 0
    offset = max_value * 0.015 if max_value else 0.1

    for patch in ax.patches:
        width = patch.get_width()
        if total and total > 0:
            label = f"{int(width)}{suffix} ({width / total * 100:.1f}%)"
        else:
            label = f"{int(width)}{suffix}"

        ax.text(
            width + offset,
            patch.get_y() + patch.get_height() / 2,
            label,
            va="center",
            ha="left",
            fontsize=9,
            color="#333333",
        )

    ax.set_xlim(0, max_value * 1.22 if max_value else 1)


def style_heatmap_axis(ax: plt.Axes, title: str, subtitle: str, xlabel: str, ylabel: str) -> None:
    """Heatmap 전용 제목/부제 스타일."""
    ax.set_title(title, loc="left", pad=22, fontsize=18, fontweight="bold")
    ax.text(
        0,
        1.02,
        subtitle,
        transform=ax.transAxes,
        ha="left",
        va="bottom",
        fontsize=10,
        color="#666666",
    )
    ax.set_xlabel(xlabel, labelpad=10)
    ax.set_ylabel(ylabel, labelpad=10)


def style_3d_axis(ax: plt.Axes) -> None:
    """Matplotlib 3D 그래프를 조금 더 차분하게 보이도록 정리한다."""
    for axis in [ax.xaxis, ax.yaxis, ax.zaxis]:
        axis.pane.set_facecolor((0.98, 0.98, 0.98, 1.0))
        axis.pane.set_edgecolor((0.86, 0.86, 0.86, 1.0))

    ax.grid(True, alpha=0.25)


# ============================================================
# 2. 데이터 로드 및 전처리
# ============================================================

def read_excel_data(input_path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if not input_path.exists():
        raise FileNotFoundError(f"입력 엑셀 파일을 찾을 수 없습니다: {input_path}")

    if sheet_name:
        return pd.read_excel(input_path, sheet_name=sheet_name)

    # sheet_name=None을 넣으면 전체 시트 dict가 되므로, 첫 번째 시트명을 따로 찾는다.
    xls = pd.ExcelFile(input_path)
    first_sheet = xls.sheet_names[0]
    print(f"[INFO] 사용 시트: {first_sheet}")
    return pd.read_excel(input_path, sheet_name=first_sheet)


def split_applicants(value: object) -> list[str]:
    """출원인 다중 표기를 분리한다. 쉼표는 회사명에 자주 들어가므로 분리 기준에서 제외한다."""
    if pd.isna(value):
        return ["미상"]
    text = str(value).strip()
    if not text:
        return ["미상"]

    # WIPS 데이터에서 다중 값은 보통 | 또는 ; 형태로 들어오는 경우가 많음.
    parts = re.split(r"\s*\|\s*|\s*;\s*", text)
    parts = [p.strip() for p in parts if p.strip()]
    return parts if parts else ["미상"]




def make_applicant_key(value: object) -> str:
    """
    출원인명 비교용 키를 만든다.

    목적:
    - 대소문자 차이 제거
    - 마침표/쉼표/공백 차이 완화
    - Co., Ltd. / Inc. 같은 법인 표기 차이 완화

    주의:
    - 이 함수만으로 모든 회사를 자동 병합하지 않는다.
    - 최종 병합 기준은 alias 엑셀에서 가져온 수동 매핑이다.
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()
    text = text.replace("㈜", "주식회사")
    text = re.sub(r"\s+", " ", text)
    text = text.upper()

    # 비교를 어렵게 하는 기호 제거
    text = re.sub(r"[\.,，·・ㆍ\(\)\[\]\{\}]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    # 자주 등장하는 법인 표기 완화
    removable_words = [
        "CO", "LTD", "INC", "INCORPORATED", "CORPORATION", "CORP",
        "COMPANY", "LIMITED", "LLC", "PLC",
    ]
    tokens = [t for t in text.split() if t not in removable_words]
    return " ".join(tokens).strip()


def read_applicant_alias_map(alias_path: Optional[Path]) -> dict[str, str]:
    """
    출원인명 alias 엑셀을 읽어서 {비교용키: 대표출원인명} 매핑을 만든다.

    지원 형식 1: 권장 long 형식
        대표출원인 | 출원인표기
        Samsung Electronics | 삼성전자주식회사
        Samsung Electronics | Samsung Electronics Co., Ltd.

    지원 형식 2: 현재 사용 중인 wide-pair 형식
        A열: 삼성전자 alias 목록, B열: 삼성전자 원자료 나열
        C열: SK hynix alias 목록, D열: SK hynix 원자료 나열
        E열: Intel alias 목록, F열: Intel 원자료 나열

    wide-pair 형식에서는 각 2열 묶음의 왼쪽 열만 alias 기준으로 사용한다.
    오른쪽 열은 원자료 검토용 목록으로 보고 매핑에는 사용하지 않는다.
    """
    if alias_path is None:
        return {}

    alias_path = Path(alias_path)
    if not alias_path.exists():
        raise FileNotFoundError(f"출원인 alias 엑셀 파일을 찾을 수 없습니다: {alias_path}")

    alias_df = pd.read_excel(alias_path, sheet_name=0, header=None)
    alias_df = alias_df.dropna(how="all").dropna(axis=1, how="all")

    alias_map: dict[str, str] = {}

    def add_alias(alias_value: object, canonical_value: object) -> None:
        if pd.isna(alias_value) or pd.isna(canonical_value):
            return
        alias = str(alias_value).strip()
        canonical = str(canonical_value).strip()
        if not alias or not canonical:
            return

        # 합계/총계/숫자만 있는 행은 매핑에서 제외
        if alias in {"합계", "총계", "TOTAL", "Total", "total"}:
            return
        if re.fullmatch(r"\d+(\.\d+)?", alias):
            return

        # 다중 출원인 결합 문자열은 원자료에서 split_applicants()로 먼저 분리되므로
        # alias 매핑에서는 그대로 등록하지 않는다. 잘못하면 공동출원인을 특정 회사로 흡수할 수 있다.
        if "|" in alias or ";" in alias:
            return

        key = make_applicant_key(alias)
        if key:
            alias_map[key] = canonical

        canonical_key = make_applicant_key(canonical)
        if canonical_key:
            alias_map[canonical_key] = canonical

    # long 형식 감지: 첫 행에 대표/표준/alias/출원인 같은 헤더가 있는 경우
    first_row_values = [str(v).strip() for v in alias_df.iloc[0].tolist() if pd.notna(v)]
    header_text = " ".join(first_row_values)
    looks_like_long = (
        len(alias_df.columns) >= 2
        and any(word in header_text for word in ["대표", "표준", "canonical", "Canonical"])
        and any(word in header_text for word in ["alias", "Alias", "표기", "출원인"])
    )

    if looks_like_long:
        long_df = pd.read_excel(alias_path, sheet_name=0)
        columns = list(long_df.columns)
        canonical_col = next(
            (c for c in columns if any(k in str(c) for k in ["대표", "표준", "canonical", "Canonical"])),
            columns[0],
        )
        alias_col = next(
            (c for c in columns if any(k in str(c) for k in ["alias", "Alias", "표기", "출원인"])),
            columns[1],
        )
        for _, row in long_df.iterrows():
            add_alias(row[alias_col], row[canonical_col])
    else:
        # wide-pair 형식: 0,2,4...번째 열을 alias 열로 사용하고, 오른쪽 열은 검토용 원자료로 무시
        col_count = len(alias_df.columns)
        for col_idx in range(0, col_count, 2):
            series = alias_df.iloc[:, col_idx].dropna()
            series = series[series.astype(str).str.strip() != ""]
            if series.empty:
                continue

            canonical = str(series.iloc[0]).strip()
            for alias in series:
                add_alias(alias, canonical)

    print(f"[INFO] 출원인 alias 매핑 수: {len(alias_map):,}")
    return alias_map


def normalize_applicant_name(value: object, alias_map: dict[str, str]) -> str:
    """출원인명을 alias 매핑 기준 대표출원인명으로 바꾼다."""
    if pd.isna(value):
        return "미상"

    original = str(value).strip()
    if not original:
        return "미상"

    key = make_applicant_key(original)
    return alias_map.get(key, original)

def preprocess(df: pd.DataFrame, alias_map: Optional[dict[str, str]] = None) -> pd.DataFrame:
    required_cols = [config.COUNTRY_COL, config.DATE_COL, config.APPLICANT_COL]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(
            "필수 컬럼이 없습니다: " + ", ".join(missing) +
            f"\n현재 컬럼: {list(df.columns)}"
        )

    data = df.copy()
    data[config.COUNTRY_COL] = data[config.COUNTRY_COL].fillna("미상").astype(str).str.strip()
    data[config.APPLICANT_COL] = data[config.APPLICANT_COL].fillna("미상").astype(str).str.strip()
    data[config.DATE_COL] = pd.to_datetime(data[config.DATE_COL], errors="coerce")

    # 출원일 없는 행은 연도 트렌드 분석에서 의미가 약하므로 제외
    data = data.dropna(subset=[config.DATE_COL]).copy()
    data["출원연도"] = data[config.DATE_COL].dt.year.astype(int)

    # 출원인 다중 표기 처리
    if config.EXPLODE_APPLICANTS:
        data["출원인_정규화"] = data[config.APPLICANT_COL].apply(split_applicants)
        data = data.explode("출원인_정규화")
    else:
        data["출원인_정규화"] = data[config.APPLICANT_COL]

    data["출원인_원문"] = data["출원인_정규화"].fillna("미상").astype(str).str.strip()
    data = data[data["출원인_원문"] != ""].copy()

    alias_map = alias_map or {}
    data["출원인_정규화"] = data["출원인_원문"].apply(lambda x: normalize_applicant_name(x, alias_map))

    return data


# ============================================================
# 3. 집계 테이블 생성
# ============================================================

def make_summary_tables(data: pd.DataFrame) -> dict[str, pd.DataFrame]:
    country_counts = (
        data.groupby(config.COUNTRY_COL)
        .size()
        .reset_index(name="출원건수")
        .sort_values("출원건수", ascending=False)
    )

    yearly_counts = (
        data.groupby("출원연도")
        .size()
        .reset_index(name="출원건수")
        .sort_values("출원연도")
    )

    applicant_counts = (
        data.groupby("출원인_정규화")
        .size()
        .reset_index(name="출원건수")
        .sort_values("출원건수", ascending=False)
    )

    country_year = (
        data.groupby(["출원연도", config.COUNTRY_COL])
        .size()
        .reset_index(name="출원건수")
        .sort_values(["출원연도", config.COUNTRY_COL])
    )

    applicant_year = (
        data.groupby(["출원연도", "출원인_정규화"])
        .size()
        .reset_index(name="출원건수")
        .sort_values(["출원연도", "출원인_정규화"])
    )

    applicant_country = (
        data.groupby(["출원인_정규화", config.COUNTRY_COL])
        .size()
        .reset_index(name="출원건수")
        .sort_values("출원건수", ascending=False)
    )

    return {
        "country_counts": country_counts,
        "yearly_counts": yearly_counts,
        "applicant_counts": applicant_counts,
        "country_year_counts": country_year,
        "applicant_year_counts": applicant_year,
        "applicant_country_counts": applicant_country,
    }


def save_summary_tables(tables: dict[str, pd.DataFrame], output_dir: Path) -> None:
    if not config.SAVE_CSV_TABLES:
        return
    table_dir = output_dir / "tables"
    table_dir.mkdir(parents=True, exist_ok=True)
    for name, table in tables.items():
        table.to_csv(table_dir / f"{name}.csv", index=False, encoding="utf-8-sig")




def save_applicant_normalization_audit(data: pd.DataFrame, output_dir: Path) -> None:
    """출원인명 정규화가 어떻게 적용됐는지 검토용 CSV를 저장한다."""
    if not config.SAVE_CSV_TABLES:
        return
    if "출원인_원문" not in data.columns:
        return

    table_dir = output_dir / "tables"
    table_dir.mkdir(parents=True, exist_ok=True)

    audit = (
        data.groupby(["출원인_원문", "출원인_정규화"])
        .size()
        .reset_index(name="출원건수")
        .sort_values(["출원인_정규화", "출원건수"], ascending=[True, False])
    )
    audit.to_csv(table_dir / "applicant_normalization_audit.csv", index=False, encoding="utf-8-sig")

# ============================================================
# 4. 그래프 함수
# ============================================================

def plot_total_yearly_trend(yearly_counts: pd.DataFrame) -> plt.Figure:
    """연도별 출원 추이: 막대 + 3년 이동평균 + 최고점 주석."""
    df = yearly_counts.copy().sort_values("출원연도")
    df["이동평균"] = df["출원건수"].rolling(window=3, min_periods=1).mean()

    fig, ax = plt.subplots(figsize=(12, 6))

    bars = ax.bar(
        df["출원연도"],
        df["출원건수"],
        width=0.72,
        alpha=0.86,
        label="연도별 출원 건수",
    )

    ax.plot(
        df["출원연도"],
        df["이동평균"],
        linewidth=2.8,
        marker="o",
        label="3년 이동평균",
    )

    if not df.empty:
        max_row = df.loc[df["출원건수"].idxmax()]
        ymax = max(float(df["출원건수"].max()), 1.0)
        ax.annotate(
            f"최고점: {int(max_row['출원연도'])}년\n{int(max_row['출원건수'])}건",
            xy=(max_row["출원연도"], max_row["출원건수"]),
            xytext=(max_row["출원연도"], ymax * 1.13),
            ha="center",
            fontsize=10,
            arrowprops=dict(arrowstyle="->", lw=1.2),
        )

        # 2020년 이후 구간이 실제 데이터에 있을 때만 최근 구간을 은은하게 표시
        if df["출원연도"].max() >= 2020:
            ax.axvspan(2020, df["출원연도"].max(), alpha=0.08)

        ax.set_ylim(0, ymax * 1.28)
        ax.set_xticks(sorted(df["출원연도"].unique()))

    polish_axes(
        ax,
        title="PIM/PNM 출원은 최근 구간에서 집중적으로 증가",
        subtitle="2010년 이후 공개·등록 특허 기준, 연도별 출원 건수 및 3년 이동평균",
        xlabel="출원연도",
        ylabel="출원 건수",
        source="Data: WIPS 검색 결과 기반 자체 집계",
    )

    ax.legend(loc="upper left")

    for bar in bars:
        height = bar.get_height()
        if height > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height,
                f"{int(height)}",
                ha="center",
                va="bottom",
                fontsize=8,
                color="#444444",
            )

    return fig


def plot_country_bar(country_counts: pd.DataFrame) -> plt.Figure:
    """국가별 출원 건수: 가로 막대 + 점유율 라벨."""
    top = country_counts.head(config.TOP_N_COUNTRIES).copy()
    top = top.sort_values("출원건수", ascending=True)
    total = float(country_counts["출원건수"].sum())

    fig, ax = plt.subplots(figsize=config.FIGSIZE_MEDIUM)
    sns.barplot(
        data=top,
        y=config.COUNTRY_COL,
        x="출원건수",
        ax=ax,
        orient="h",
        color=sns.color_palette("deep")[0],
    )

    add_horizontal_bar_labels(ax, top["출원건수"], total=total)

    # title=f"국가코드별 출원 비중 TOP {config.TOP_N_COUNTRIES}",
    polish_axes(
        ax,
        title=f"국가코드별 출원 비중",
        subtitle="막대 길이는 출원 건수, 괄호 안 수치는 전체 대비 비율",
        xlabel="출원 건수",
        ylabel="국가코드",
        source="Data: WIPS 검색 결과 기반 자체 집계",
    )

    return fig


def plot_applicant_bar(applicant_counts: pd.DataFrame) -> plt.Figure:
    """주요 출원인 TOP 그래프: 긴 회사명에 맞춘 가로 막대."""
    top = applicant_counts.head(config.TOP_N_APPLICANTS).copy()
    top = top.sort_values("출원건수", ascending=True)
    top["출원인_표시"] = top["출원인_정규화"].apply(lambda x: shorten_label(x, 38))

    fig, ax = plt.subplots(figsize=(13, 9))
    sns.barplot(
        data=top,
        y="출원인_표시",
        x="출원건수",
        ax=ax,
        orient="h",
        color=sns.color_palette("deep")[1],
    )

    add_horizontal_bar_labels(ax, top["출원건수"], total=None)

    polish_axes(
        ax,
        title=f"상위 {config.TOP_N_APPLICANTS}개 출원인이 PIM/PNM 특허 출원을 주도[출원인 명칭 정리 부족]",
        # subtitle="출원인 명칭 기준 단순 집계. 계열사·표기 차이는 별도 정규화 필요",
        xlabel="출원 건수",
        ylabel="출원인",
        source="Data: WIPS 검색 결과 기반 자체 집계",
    )

    return fig


def plot_country_year_stacked_bar(country_year: pd.DataFrame, country_counts: pd.DataFrame) -> plt.Figure:
    top_countries = country_counts.head(config.TOP_N_COUNTRIES)[config.COUNTRY_COL].tolist()
    subset = country_year[country_year[config.COUNTRY_COL].isin(top_countries)]
    pivot = subset.pivot_table(
        index="출원연도",
        columns=config.COUNTRY_COL,
        values="출원건수",
        aggfunc="sum",
        fill_value=0,
    ).sort_index()

    fig, ax = plt.subplots(figsize=config.FIGSIZE_WIDE)
    pivot.plot(kind="bar", stacked=True, ax=ax, width=0.82)

    polish_axes(
        ax,
        title=f"연도별 출원 증가를 주도한 국가는 어디인가[임시]",
        subtitle=f"상위 {config.TOP_N_COUNTRIES}개 국가코드 기준 누적 막대그래프",
        xlabel="출원연도",
        ylabel="출원 건수",
        source="Data: WIPS 검색 결과 기반 자체 집계",
    )
    ax.legend(title="국가코드", bbox_to_anchor=(1.02, 1), loc="upper left")

    return fig


def plot_country_year_heatmap(country_year: pd.DataFrame, country_counts: pd.DataFrame) -> plt.Figure:
    top_countries = country_counts.head(config.TOP_N_COUNTRIES)[config.COUNTRY_COL].tolist()
    subset = country_year[country_year[config.COUNTRY_COL].isin(top_countries)]
    pivot = subset.pivot_table(
        index=config.COUNTRY_COL,
        columns="출원연도",
        values="출원건수",
        aggfunc="sum",
        fill_value=0,
    )
    pivot = pivot.loc[[c for c in top_countries if c in pivot.index]]

    fig, ax = plt.subplots(figsize=config.FIGSIZE_WIDE)
    sns.heatmap(
        pivot,
        annot=True,
        fmt=".0f",
        linewidths=0.5,
        cmap="YlGnBu",
        ax=ax,
        cbar_kws={"label": "출원 건수"},
    )
    style_heatmap_axis(
        ax,
        title="국가별 출원 집중 시점 비교",
        subtitle="색이 진할수록 해당 국가·연도 조합의 출원 건수가 많음",
        xlabel="출원연도",
        ylabel="국가코드",
    )

    return fig


def plot_applicant_country_heatmap(
    applicant_country: pd.DataFrame,
    applicant_counts: pd.DataFrame,
    country_counts: pd.DataFrame,
) -> plt.Figure:
    top_applicants = applicant_counts.head(config.TOP_N_APPLICANTS_3D)["출원인_정규화"].tolist()
    top_countries = country_counts.head(config.TOP_N_COUNTRIES_3D)[config.COUNTRY_COL].tolist()
    subset = applicant_country[
        applicant_country["출원인_정규화"].isin(top_applicants)
        & applicant_country[config.COUNTRY_COL].isin(top_countries)
    ]
    pivot = subset.pivot_table(
        index="출원인_정규화",
        columns=config.COUNTRY_COL,
        values="출원건수",
        aggfunc="sum",
        fill_value=0,
    )
    pivot = pivot.reindex(index=top_applicants, columns=top_countries, fill_value=0)
    pivot.index = [shorten_label(name, 36) for name in pivot.index]

    fig, ax = plt.subplots(figsize=(14, 9))
    sns.heatmap(
        pivot,
        annot=True,
        fmt=".0f",
        linewidths=0.5,
        cmap="Blues",
        ax=ax,
        cbar_kws={"label": "출원 건수"},
    )
    style_heatmap_axis(
        ax,
        title="주요 출원인의 국가별 포트폴리오 분포",
        subtitle="상위 출원인과 상위 국가코드 기준 출원 건수 매트릭스",
        xlabel="국가코드",
        ylabel="출원인",
    )

    return fig


def plot_applicant_country_bubble(
    applicant_country: pd.DataFrame,
    applicant_counts: pd.DataFrame,
    country_counts: pd.DataFrame,
) -> plt.Figure:
    """출원인 x 국가코드 x 출원건수를 버블 매트릭스로 표현한다."""
    top_applicants = applicant_counts.head(config.TOP_N_APPLICANTS_3D)["출원인_정규화"].tolist()
    top_countries = country_counts.head(config.TOP_N_COUNTRIES_3D)[config.COUNTRY_COL].tolist()

    plot_df = applicant_country[
        applicant_country["출원인_정규화"].isin(top_applicants)
        & applicant_country[config.COUNTRY_COL].isin(top_countries)
    ].copy()

    applicants = list(reversed(top_applicants))
    countries = top_countries

    applicant_map = {name: i for i, name in enumerate(applicants)}
    country_map = {name: i for i, name in enumerate(countries)}

    plot_df["x"] = plot_df[config.COUNTRY_COL].map(country_map)
    plot_df["y"] = plot_df["출원인_정규화"].map(applicant_map)

    fig, ax = plt.subplots(figsize=(12, 8))

    max_count = plot_df["출원건수"].max() if not plot_df.empty else 1
    sizes = 120 + (plot_df["출원건수"] / max_count) * 1700

    ax.scatter(
        plot_df["x"],
        plot_df["y"],
        s=sizes,
        alpha=0.62,
        edgecolors="white",
        linewidth=1.4,
    )

    for _, row in plot_df.iterrows():
        if row["출원건수"] >= 1:
            ax.text(
                row["x"],
                row["y"],
                str(int(row["출원건수"])),
                ha="center",
                va="center",
                fontsize=8,
                color="#222222",
            )

    ax.set_xticks(range(len(countries)))
    ax.set_xticklabels(countries)
    ax.set_yticks(range(len(applicants)))
    ax.set_yticklabels([shorten_label(name, 34) for name in applicants])

    polish_axes(
        ax,
        title="출원인별 국가 포트폴리오를 한눈에 비교",
        subtitle="버블 크기는 해당 출원인의 국가별 출원 건수를 의미",
        xlabel="국가코드",
        ylabel="출원인",
        source="Data: WIPS 검색 결과 기반 자체 집계",
    )
    ax.grid(True, alpha=0.25)

    return fig


def plot_3d_country_year(country_year: pd.DataFrame, country_counts: pd.DataFrame) -> plt.Figure:
    top_countries = country_counts.head(config.TOP_N_COUNTRIES_3D)[config.COUNTRY_COL].tolist()
    subset = country_year[country_year[config.COUNTRY_COL].isin(top_countries)].copy()
    years = sorted(subset["출원연도"].unique())

    pivot = subset.pivot_table(
        index=config.COUNTRY_COL,
        columns="출원연도",
        values="출원건수",
        aggfunc="sum",
        fill_value=0,
    ).reindex(index=top_countries, columns=years, fill_value=0)

    fig = plt.figure(figsize=config.FIGSIZE_3D)
    ax = fig.add_subplot(111, projection="3d")

    xpos, ypos = np.meshgrid(np.arange(len(years)), np.arange(len(top_countries)))
    xpos = xpos.flatten()
    ypos = ypos.flatten()
    zpos = np.zeros_like(xpos)
    dz = pivot.to_numpy().flatten()
    dx = np.full_like(xpos, 0.55, dtype=float)
    dy = np.full_like(ypos, 0.55, dtype=float)

    ax.bar3d(xpos, ypos, zpos, dx, dy, dz, shade=True, alpha=0.82)
    ax.set_title("3D 출원 동향: 출원연도 x 국가코드 x 출원건수", pad=20, fontweight="bold")
    ax.set_xlabel("출원연도", labelpad=12)
    ax.set_ylabel("국가코드", labelpad=12)
    ax.set_zlabel("출원건수", labelpad=10)
    ax.set_xticks(np.arange(len(years)) + 0.27)
    ax.set_xticklabels(years, rotation=45, ha="right")
    ax.set_yticks(np.arange(len(top_countries)) + 0.27)
    ax.set_yticklabels(top_countries)
    ax.view_init(elev=25, azim=-55)
    style_3d_axis(ax)

    return fig


def plot_3d_applicant_country(
    applicant_country: pd.DataFrame,
    applicant_counts: pd.DataFrame,
    country_counts: pd.DataFrame,
) -> plt.Figure:
    top_applicants = applicant_counts.head(config.TOP_N_APPLICANTS_3D)["출원인_정규화"].tolist()
    top_countries = country_counts.head(config.TOP_N_COUNTRIES_3D)[config.COUNTRY_COL].tolist()

    subset = applicant_country[
        applicant_country["출원인_정규화"].isin(top_applicants)
        & applicant_country[config.COUNTRY_COL].isin(top_countries)
    ]

    pivot = subset.pivot_table(
        index="출원인_정규화",
        columns=config.COUNTRY_COL,
        values="출원건수",
        aggfunc="sum",
        fill_value=0,
    ).reindex(index=top_applicants, columns=top_countries, fill_value=0)

    fig = plt.figure(figsize=(15, 10))
    ax = fig.add_subplot(111, projection="3d")

    xpos, ypos = np.meshgrid(np.arange(len(top_countries)), np.arange(len(top_applicants)))
    xpos = xpos.flatten()
    ypos = ypos.flatten()
    zpos = np.zeros_like(xpos)
    dz = pivot.to_numpy().flatten()
    dx = np.full_like(xpos, 0.55, dtype=float)
    dy = np.full_like(ypos, 0.55, dtype=float)

    ax.bar3d(xpos, ypos, zpos, dx, dy, dz, shade=True, alpha=0.82)
    ax.set_title("3D 출원 분포: 국가코드 x 출원인 x 출원건수", pad=20, fontweight="bold")
    ax.set_xlabel("국가코드", labelpad=12)
    ax.set_ylabel("출원인", labelpad=18)
    ax.set_zlabel("출원건수", labelpad=10)
    ax.set_xticks(np.arange(len(top_countries)) + 0.27)
    ax.set_xticklabels(top_countries)
    ax.set_yticks(np.arange(len(top_applicants)) + 0.27)
    ax.set_yticklabels([shorten_label(name, 24) for name in top_applicants], fontsize=8)
    ax.view_init(elev=28, azim=-48)
    style_3d_axis(ax)

    return fig


# ============================================================
# 5. 반복 생성 그래프
# ============================================================

def generate_country_detail_charts(
    data: pd.DataFrame,
    country_counts: pd.DataFrame,
    output_img_dir: Path,
    pdf: Optional[PdfPages],
) -> None:
    """국가별 연도 추이 그래프를 반복 생성한다."""
    top_countries = country_counts.head(config.TOP_N_COUNTRIES)[config.COUNTRY_COL].tolist()
    all_years = sorted(data["출원연도"].unique())

    for country in top_countries:
        subset = data[data[config.COUNTRY_COL] == country]
        yearly = subset.groupby("출원연도").size().reset_index(name="출원건수").sort_values("출원연도")

        fig, ax = plt.subplots(figsize=config.FIGSIZE_MEDIUM)
        sns.lineplot(data=yearly, x="출원연도", y="출원건수", marker="o", linewidth=2.5, ax=ax)
        ax.fill_between(yearly["출원연도"], yearly["출원건수"], alpha=0.12)
        ax.set_xticks(all_years)

        for _, row in yearly.iterrows():
            ax.text(
                row["출원연도"],
                row["출원건수"],
                str(int(row["출원건수"])),
                ha="center",
                va="bottom",
                fontsize=8,
                color="#444444",
            )

        polish_axes(
            ax,
            title=f"국가코드 {country}: 연도별 PIM/PNM 출원 동향",
            subtitle="국가별 상세 추이 그래프",
            xlabel="출원연도",
            ylabel="출원 건수",
            source="Data: WIPS 검색 결과 기반 자체 집계",
        )

        save_figure(fig, output_img_dir / "country_detail", f"country_{safe_filename(country)}_yearly_trend", pdf)


def generate_applicant_detail_charts(
    data: pd.DataFrame,
    applicant_counts: pd.DataFrame,
    output_img_dir: Path,
    pdf: Optional[PdfPages],
) -> None:
    """상위 출원인별 연도 추이 그래프를 반복 생성한다."""
    top_applicants = applicant_counts.head(config.TOP_N_APPLICANTS)["출원인_정규화"].tolist()
    all_years = sorted(data["출원연도"].unique())

    for applicant in top_applicants:
        subset = data[data["출원인_정규화"] == applicant]
        yearly = subset.groupby("출원연도").size().reset_index(name="출원건수").sort_values("출원연도")

        fig, ax = plt.subplots(figsize=config.FIGSIZE_MEDIUM)
        sns.lineplot(data=yearly, x="출원연도", y="출원건수", marker="o", linewidth=2.5, ax=ax)
        ax.fill_between(yearly["출원연도"], yearly["출원건수"], alpha=0.12)
        ax.set_xticks(all_years)

        for _, row in yearly.iterrows():
            ax.text(
                row["출원연도"],
                row["출원건수"],
                str(int(row["출원건수"])),
                ha="center",
                va="bottom",
                fontsize=8,
                color="#444444",
            )

        polish_axes(
            ax,
            title=f"출원인별 연도 추이: {shorten_label(applicant, 45)}",
            subtitle="상위 출원인별 상세 추이 그래프",
            xlabel="출원연도",
            ylabel="출원 건수",
            source="Data: WIPS 검색 결과 기반 자체 집계",
        )

        save_figure(fig, output_img_dir / "applicant_detail", f"applicant_{safe_filename(applicant)}_yearly_trend", pdf)


# ============================================================
# 6. 전체 실행
# ============================================================

def build_all_visuals(input_path: Path, output_dir: Path, sheet_name: Optional[str], alias_path: Optional[Path] = None) -> None:
    apply_theme()

    output_dir.mkdir(parents=True, exist_ok=True)
    image_dir = output_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)

    raw = read_excel_data(input_path, sheet_name)
    alias_map = read_applicant_alias_map(alias_path)
    data = preprocess(raw, alias_map=alias_map)
    tables = make_summary_tables(data)
    save_summary_tables(tables, output_dir)
    save_applicant_normalization_audit(data, output_dir)

    print(f"[INFO] 원본 행 수: {len(raw):,}")
    print(f"[INFO] 분석 행 수: {len(data):,}")
    print(f"[INFO] 출원연도 범위: {data['출원연도'].min()} ~ {data['출원연도'].max()}")
    print(f"[INFO] 국가 수: {data[config.COUNTRY_COL].nunique():,}")
    print(f"[INFO] 정규화 후 출원인 수: {data['출원인_정규화'].nunique():,}")
    if "출원인_원문" in data.columns:
        print(f"[INFO] 정규화 전 출원인 표기 수: {data['출원인_원문'].nunique():,}")

    pdf_path = output_dir / config.PDF_FILENAME
    pdf_context = PdfPages(pdf_path) if config.SAVE_PDF else None

    try:
        # 기본 요약 그래프
        figures = [
            (plot_total_yearly_trend(tables["yearly_counts"]), "01_total_yearly_trend"),
            (plot_country_bar(tables["country_counts"]), "02_country_count_top"),
            (plot_applicant_bar(tables["applicant_counts"]), "03_applicant_count_top"),
            (
                plot_country_year_stacked_bar(
                    tables["country_year_counts"],
                    tables["country_counts"],
                ),
                "04_country_year_stacked_bar",
            ),
            (
                plot_country_year_heatmap(
                    tables["country_year_counts"],
                    tables["country_counts"],
                ),
                "05_country_year_heatmap",
            ),
            (
                plot_applicant_country_heatmap(
                    tables["applicant_country_counts"],
                    tables["applicant_counts"],
                    tables["country_counts"],
                ),
                "06_applicant_country_heatmap",
            ),
            (
                plot_3d_country_year(
                    tables["country_year_counts"],
                    tables["country_counts"],
                ),
                "07_3d_country_year",
            ),
            (
                plot_3d_applicant_country(
                    tables["applicant_country_counts"],
                    tables["applicant_counts"],
                    tables["country_counts"],
                ),
                "08_3d_applicant_country",
            ),
            (
                plot_applicant_country_bubble(
                    tables["applicant_country_counts"],
                    tables["applicant_counts"],
                    tables["country_counts"],
                ),
                "09_applicant_country_bubble",
            ),
        ]

        for fig, filename in figures:
            save_figure(fig, image_dir, filename, pdf_context)

        # 반복 생성 그래프
        generate_country_detail_charts(data, tables["country_counts"], image_dir, pdf_context)
        generate_applicant_detail_charts(data, tables["applicant_counts"], image_dir, pdf_context)

    finally:
        if pdf_context is not None:
            pdf_context.close()

    print(f"[DONE] 이미지 저장 폴더: {image_dir.resolve()}")
    if config.SAVE_PDF:
        print(f"[DONE] PDF 저장 파일: {pdf_path.resolve()}")
    if config.SAVE_CSV_TABLES:
        print(f"[DONE] 집계 CSV 저장 폴더: {(output_dir / 'tables').resolve()}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PIM 출원 동향 시각화 자동화")
    parser.add_argument("--input", type=str, default=str(config.INPUT_EXCEL), help="입력 엑셀 파일 경로")
    parser.add_argument("--output", type=str, default=str(config.OUTPUT_DIR), help="출력 폴더 경로")
    parser.add_argument("--sheet", type=str, default=config.SHEET_NAME, help="사용할 시트명. 생략 시 첫 번째 시트 사용")
    parser.add_argument(
        "--alias",
        type=str,
        default=None,
        help="출원인명 통합 기준 엑셀 파일 경로. 예: 출원인_삼성SK 출원인 중복.xlsx",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    build_all_visuals(
        input_path=Path(args.input),
        output_dir=Path(args.output),
        sheet_name=args.sheet,
        alias_path=Path(args.alias) if args.alias else None,
    )
