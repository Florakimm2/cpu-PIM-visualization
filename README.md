# PIM 출원 동향 시각화 자동화 코드

## 목적
엑셀 특허 데이터에서 `국가코드`, `출원일`, `출원인`을 기준으로 PIM 출원 동향 그래프를 자동 생성한다.

## 생성되는 결과물
- `images/`: PNG 그래프 이미지
- `images/country_detail/`: 국가별 연도 추이 반복 생성 이미지
- `images/applicant_detail/`: 출원인별 연도 추이 반복 생성 이미지
- `tables/`: 집계 CSV 파일
- `pim_application_trend_report.pdf`: 전체 그래프를 묶은 PDF 리포트

## 주요 그래프
1. 전체 연도별 출원 추이
2. 국가코드별 출원건수 TOP 그래프
3. 주요 출원인별 출원건수 TOP 그래프
4. 연도별·국가별 누적 막대그래프
5. 국가코드 x 출원연도 Heatmap
6. 출원인 x 국가코드 Heatmap
7. 3D 그래프: 출원연도 x 국가코드 x 출원건수
8. 3D 그래프: 국가코드 x 출원인 x 출원건수
9. 국가별 상세 그래프 반복 생성
10. 출원인별 상세 그래프 반복 생성

## 실행 방법
```bash
pip install -r requirements.txt
python pim_trend_visualizer.py --input "전데검완_등록.xlsx" --output "pim_visual_outputs"
```

시트명을 직접 지정하려면:
```bash
python pim_trend_visualizer.py --input "전데검완_등록.xlsx" --sheet "다운로드" --output "pim_visual_outputs"
```

## 설정 변경
`config.py`에서 다음 값을 바꾸면 된다.

```python
TOP_N_COUNTRIES = 8
TOP_N_APPLICANTS = 15
TOP_N_APPLICANTS_3D = 10
TOP_N_COUNTRIES_3D = 8
EXPLODE_APPLICANTS = True
```

## 주의
- 회사명에는 쉼표가 자주 들어가므로 출원인 분리 기준으로 쉼표는 사용하지 않는다.
- 출원인이 `A | B` 또는 `A; B`처럼 들어간 경우에만 분리한다.
