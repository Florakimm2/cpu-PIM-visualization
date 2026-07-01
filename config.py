from pathlib import Path

# =========================
# 기본 입력/출력 설정
# =========================
# 실행 시 --input 옵션으로 바꿀 수 있음.
INPUT_EXCEL = Path("전데검완_등록.xlsx")
OUTPUT_DIR = Path("pim_visual_outputs")
SHEET_NAME = None  # None이면 첫 번째 시트 사용. 예: "다운로드"

# =========================
# 분석 기준 컬럼명
# =========================
COUNTRY_COL = "국가코드"
DATE_COL = "출원일"
APPLICANT_COL = "출원인"

# =========================
# 그래프 범위 설정
# =========================
TOP_N_COUNTRIES = 8
TOP_N_APPLICANTS = 15
TOP_N_APPLICANTS_3D = 10
TOP_N_COUNTRIES_3D = 8

# 출원인이 "A | B"처럼 여러 명이면 분리해서 집계할지 여부
# 회사명에 쉼표가 들어가는 경우가 많으므로 쉼표 기준 분리는 하지 않음.
EXPLODE_APPLICANTS = True

# =========================
# 저장 옵션
# =========================
SAVE_PNG = True
SAVE_PDF = True
SAVE_CSV_TABLES = True
PDF_FILENAME = "pim_application_trend_report.pdf"

# =========================
# 디자인 테마 설정
# =========================
SEABORN_STYLE = "whitegrid"
SEABORN_CONTEXT = "talk"
FIG_DPI = 180
FIGSIZE_WIDE = (14, 8)
FIGSIZE_MEDIUM = (11, 7)
FIGSIZE_3D = (13, 9)
