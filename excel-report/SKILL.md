---
name: excel-report
description: "CSV 파일을 서식이 적용된 Excel 파일로 변환. 헤더 색상, 천 단위 콤마, 합계 행, 차트 포함."
user-invocable: true
---

# Excel Report

CSV → 서식 Excel 변환 스킬. `scripts/excel_report.py` 실행.

## 요구사항

- Python 3, `openpyxl` (`pip install openpyxl`)

## 사용법

사용자가 CSV 파일 경로를 주면:

```bash
python /Users/goyeongjun/skills/excel-report/scripts/excel_report.py <csv_path>
```

출력: 입력 파일과 같은 위치에 `.xlsx` 저장.

## 적용 서식

- **헤더**: 파란 배경(`4472C4`), 흰 굵은 글씨
- **숫자 열**: 천 단위 콤마(`#,##0`)
- **합계 행**: 마지막 행에 숫자 열 SUM, 굵게
- **차트**: 첫 번째 숫자 열로 바 차트 삽입 (시트 우측)

## 워크플로

1. 파일 경로 확인 → csv 맞는지 체크
2. `python … excel_report.py <path>` 실행
3. 생성된 `.xlsx` 경로 사용자에게 알림
