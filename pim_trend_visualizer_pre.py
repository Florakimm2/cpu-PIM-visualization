"""
PIM 출원 동향 시각화 자동화 스크립트

기능
1. 엑셀 데이터 자동 로드
2. 국가코드 / 출원일 / 출원인 기준 집계
3. 여러 그래프를 반복문으로 대량 생성
4. 디자인 테마 통일
5. 3D 그래프 생성
6. PNG 이미지와 PDF 리포트 저장

실행 예시
python pim_trend_visualizer.py --input "전데검완_등록.xlsx" --output "pim_visual_outputs"
"""

from __future__ import annotations

import argparse
import re
import warnings
from pathlib import Path
from typing import Iterable, Optional

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
    """전체 그래프의 디자인을 통일한다."""
    font_name = set_korean_font()
    if font_name:
        sns.set_theme(
            style=config.SEABORN_STYLE,
            context=config.SEABORN_CONTEXT,
            font=font_name,
        )
    else:
        sns.set_theme(
            style=config.SEABORN_STYLE,
            context=config.SEABORN_CONTEXT,
        )
    plt.rcParams["axes.prop_cycle"] = plt.cycler(color=sns.color_palette("deep"))

    plt.rcParams.update({
        "figure.dpi": config.FIG_DPI,
        "savefig.dpi": config.FIG_DPI,
        "axes.titleweight": "bold",
        "axes.labelweight": "bold",
        "axes.titlesize": 17,
        "axes.labelsize": 13,
        "xtick.labelsize": 10,
        "ytick.labelsize": 10,
        "legend.fontsize": 10,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "grid.alpha": 0.35,
    })


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


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
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

    data["출원인_정규화"] = data["출원인_정규화"].fillna("미상").astype(str).str.strip()
    data = data[data["출원인_정규화"] != ""].copy()

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


# ============================================================
# 4. 그래프 함수
# ============================================================
def plot_yearly_trend_pretty(yearly_counts, output_path):
    df = yearly_counts.copy()
    df = df.sort_values("출원연도")
    df["이동평균"] = df["출원건수"].rolling(window=3, min_periods=1).mean()

    fig, ax = plt.subplots(figsize=(12, 6))

    bars = ax.bar(
        df["출원연도"],
        df["출원건수"],
        width=0.72,
        alpha=0.85,
        label="연도별 출원 건수"
    )

    ax.plot(
        df["출원연도"],
        df["이동평균"],
        linewidth=2.8,
        marker="o",
        label="3년 이동평균"
    )

    max_row = df.loc[df["출원건수"].idxmax()]
    ax.annotate(
        f"최고점: {int(max_row['출원연도'])}년\n{int(max_row['출원건수'])}건",
        xy=(max_row["출원연도"], max_row["출원건수"]),
        xytext=(max_row["출원연도"], max_row["출원건수"] * 1.15),
        ha="center",
        fontsize=10,
        arrowprops=dict(arrowstyle="->", lw=1.2)
    )

    if df["출원연도"].max() >= 2020:
        ax.axvspan(2020, df["출원연도"].max(), alpha=0.08)

    polish_axes(
        ax,
        title="PIM/PNM 출원은 최근 구간에서 집중적으로 증가",
        subtitle="2010년 이후 공개·등록 특허 기준, 연도별 출원 건수 및 3년 이동평균",
        xlabel="출원연도",
        ylabel="출원 건수",
        source="Data: WIPS 검색 결과 기반 자체 집계"
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
                color="#444444"
            )

    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
'''
def plot_total_yearly_trend(yearly_counts: pd.DataFrame) -> plt.Figure:
    fig, ax = plt.subplots(figsize=config.FIGSIZE_WIDE)
    sns.lineplot(
        data=yearly_counts,
        x="출원연도",
        y="출원건수",
        marker="o",
        linewidth=2.5,
        ax=ax,
    )
    ax.set_title("PIM 전체 출원 동향: 연도별 출원건수")
    ax.set_xlabel("출원연도")
    ax.set_ylabel("출원건수")
    ax.set_xticks(sorted(yearly_counts["출원연도"].unique()))
    for _, row in yearly_counts.iterrows():
        ax.text(row["출원연도"], row["출원건수"], str(int(row["출원건수"])), ha="center", va="bottom", fontsize=9)
    return fig
'''

def plot_country_bar(country_counts: pd.DataFrame) -> plt.Figure:
    top = country_counts.head(config.TOP_N_COUNTRIES).copy()
    fig, ax = plt.subplots(figsize=config.FIGSIZE_MEDIUM)
    sns.barplot(data=top, y=config.COUNTRY_COL, x="출원건수", ax=ax, color=sns.color_palette("deep")[0])
    ax.set_title(f"국가코드별 출원건수 TOP {config.TOP_N_COUNTRIES}")
    ax.set_xlabel("출원건수")
    ax.set_ylabel("국가코드")
    annotate_bar_values(ax)
    return fig


def plot_applicant_bar(applicant_counts: pd.DataFrame) -> plt.Figure:
    top = applicant_counts.head(config.TOP_N_APPLICANTS).copy()
    fig, ax = plt.subplots(figsize=(14, 9))
    sns.barplot(data=top, y="출원인_정규화", x="출원건수", ax=ax, color=sns.color_palette("deep")[1])
    ax.set_title(f"주요 출원인별 출원건수 TOP {config.TOP_N_APPLICANTS}")
    ax.set_xlabel("출원건수")
    ax.set_ylabel("출원인")
    annotate_bar_values(ax)
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
    ax.set_title(f"연도별·국가코드별 출원건수: TOP {config.TOP_N_COUNTRIES} 국가")
    ax.set_xlabel("출원연도")
    ax.set_ylabel("출원건수")
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
    sns.heatmap(pivot, annot=True, fmt=".0f", linewidths=0.5, cmap="YlGnBu", ax=ax)
    ax.set_title("국가코드 x 출원연도 출원건수 Heatmap")
    ax.set_xlabel("출원연도")
    ax.set_ylabel("국가코드")
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

    fig, ax = plt.subplots(figsize=(14, 9))
    sns.heatmap(pivot, annot=True, fmt=".0f", linewidths=0.5, cmap="Blues", ax=ax)
    ax.set_title("주요 출원인 x 국가코드 출원건수 Heatmap")
    ax.set_xlabel("국가코드")
    ax.set_ylabel("출원인")
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

    ax.bar3d(xpos, ypos, zpos, dx, dy, dz, shade=True, alpha=0.85)
    ax.set_title("3D 출원 동향: 출원연도 x 국가코드 x 출원건수", pad=20)
    ax.set_xlabel("출원연도", labelpad=12)
    ax.set_ylabel("국가코드", labelpad=12)
    ax.set_zlabel("출원건수", labelpad=10)
    ax.set_xticks(np.arange(len(years)) + 0.27)
    ax.set_xticklabels(years, rotation=45, ha="right")
    ax.set_yticks(np.arange(len(top_countries)) + 0.27)
    ax.set_yticklabels(top_countries)
    ax.view_init(elev=25, azim=-55)
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

    ax.bar3d(xpos, ypos, zpos, dx, dy, dz, shade=True, alpha=0.85)
    ax.set_title("3D 출원 분포: 국가코드 x 출원인 x 출원건수", pad=20)
    ax.set_xlabel("국가코드", labelpad=12)
    ax.set_ylabel("출원인", labelpad=18)
    ax.set_zlabel("출원건수", labelpad=10)
    ax.set_xticks(np.arange(len(top_countries)) + 0.27)
    ax.set_xticklabels(top_countries)
    ax.set_yticks(np.arange(len(top_applicants)) + 0.27)
    ax.set_yticklabels([name[:24] for name in top_applicants], fontsize=8)
    ax.view_init(elev=28, azim=-48)
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

    for country in top_countries:
        subset = data[data[config.COUNTRY_COL] == country]
        yearly = subset.groupby("출원연도").size().reset_index(name="출원건수").sort_values("출원연도")

        fig, ax = plt.subplots(figsize=config.FIGSIZE_MEDIUM)
        sns.lineplot(data=yearly, x="출원연도", y="출원건수", marker="o", linewidth=2.5, ax=ax)
        ax.set_title(f"국가코드 {country}: 연도별 PIM 출원 동향")
        ax.set_xlabel("출원연도")
        ax.set_ylabel("출원건수")
        ax.set_xticks(sorted(data["출원연도"].unique()))
        for _, row in yearly.iterrows():
            ax.text(row["출원연도"], row["출원건수"], str(int(row["출원건수"])), ha="center", va="bottom", fontsize=9)

        save_figure(fig, output_img_dir / "country_detail", f"country_{safe_filename(country)}_yearly_trend", pdf)


def generate_applicant_detail_charts(
    data: pd.DataFrame,
    applicant_counts: pd.DataFrame,
    output_img_dir: Path,
    pdf: Optional[PdfPages],
) -> None:
    """상위 출원인별 연도 추이 그래프를 반복 생성한다."""
    top_applicants = applicant_counts.head(config.TOP_N_APPLICANTS)["출원인_정규화"].tolist()

    for applicant in top_applicants:
        subset = data[data["출원인_정규화"] == applicant]
        yearly = subset.groupby("출원연도").size().reset_index(name="출원건수").sort_values("출원연도")

        fig, ax = plt.subplots(figsize=config.FIGSIZE_MEDIUM)
        sns.lineplot(data=yearly, x="출원연도", y="출원건수", marker="o", linewidth=2.5, ax=ax)
        ax.set_title(f"출원인별 연도 추이: {applicant[:45]}")
        ax.set_xlabel("출원연도")
        ax.set_ylabel("출원건수")
        ax.set_xticks(sorted(data["출원연도"].unique()))
        for _, row in yearly.iterrows():
            ax.text(row["출원연도"], row["출원건수"], str(int(row["출원건수"])), ha="center", va="bottom", fontsize=9)

        save_figure(fig, output_img_dir / "applicant_detail", f"applicant_{safe_filename(applicant)}_yearly_trend", pdf)


# ============================================================
# 6. 전체 실행
# ============================================================

def build_all_visuals(input_path: Path, output_dir: Path, sheet_name: Optional[str]) -> None:
    apply_theme()

    output_dir.mkdir(parents=True, exist_ok=True)
    image_dir = output_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)

    raw = read_excel_data(input_path, sheet_name)
    data = preprocess(raw)
    tables = make_summary_tables(data)
    save_summary_tables(tables, output_dir)

    print(f"[INFO] 원본 행 수: {len(raw):,}")
    print(f"[INFO] 분석 행 수: {len(data):,}")
    print(f"[INFO] 출원연도 범위: {data['출원연도'].min()} ~ {data['출원연도'].max()}")
    print(f"[INFO] 국가 수: {data[config.COUNTRY_COL].nunique():,}")
    print(f"[INFO] 출원인 수: {data['출원인_정규화'].nunique():,}")

    pdf_path = output_dir / config.PDF_FILENAME
    pdf_context = PdfPages(pdf_path) if config.SAVE_PDF else None

    try:
        # 기본 요약 그래프
        figures = [
            (plot_total_yearly_trend(tables["yearly_counts"]), "01_total_yearly_trend"),
            (plot_country_bar(tables["country_counts"]), "02_country_count_top"),
            (plot_applicant_bar(tables["applicant_counts"]), "03_applicant_count_top"),
            (plot_country_year_stacked_bar(tables["country_year_counts"], tables["country_counts"]), "04_country_year_stacked_bar"),
            (plot_country_year_heatmap(tables["country_year_counts"], tables["country_counts"]), "05_country_year_heatmap"),
            (plot_applicant_country_heatmap(tables["applicant_country_counts"], tables["applicant_counts"], tables["country_counts"]), "06_applicant_country_heatmap"),
            (plot_3d_country_year(tables["country_year_counts"], tables["country_counts"]), "07_3d_country_year"),
            (plot_3d_applicant_country(tables["applicant_country_counts"], tables["applicant_counts"], tables["country_counts"]), "08_3d_applicant_country"),
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
    return parser.parse_args()


if __name__ == "__main__":
    apply_report_style()
    args = parse_args()
    build_all_visuals(
        input_path=Path(args.input),
        output_dir=Path(args.output),
        sheet_name=args.sheet,
    )
