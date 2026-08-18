#!/usr/bin/env python3
"""
classify_chc1.py가 만든 전체(10컬럼, 판정근거 포함) 분류 결과에서
메일 발송용 6컬럼(제품명/업체명/허가일자/주성분/포장단위/CHC1코드)만
추려낸 경량 리포트를 만든다.

서식:
- 헤더(1행): 굵게 + 회색 배경 + 가운데 정렬
- 폰트: 나눔고딕 (헤더 11pt / 본문 10pt) -- xlsx는 폰트 이름만 저장하므로
  받는 사람 PC에 나눔고딕이 없으면 엑셀이 자동으로 다른 폰트로 대체한다.
- 주성분 컬럼: 줄바꿈 처리, 너비는 고정값 유지 (원래도 길어서 줄바꿈 필요)
- 나머지 컬럼: 실제 내용 길이에 맞춰 한 줄로 보이도록 자동 폭 조정

사용법:
    python3 export_email_report.py --source CHC1_분류_202607.xlsx --out CHC1_리포트_202607.xlsx
"""
import argparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

EMAIL_COLUMNS = ["제품명", "업체명", "허가일자", "주성분", "포장단위", "CHC1코드"]

FONT_NAME = "나눔고딕"
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True)
BODY_FONT = Font(name=FONT_NAME, size=10)
HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
PLAIN_ALIGN = Alignment(vertical="top")

INGREDIENT_COLUMN = "주성분"
INGREDIENT_WIDTH = 45  # 줄바꿈으로 보여줄 것이므로 폭은 고정 유지

# 한글은 라틴 문자보다 시각적으로 넓으므로, 한글 비중이 큰 컬럼은 자동폭 계산 시
# 문자 수에 가중치를 곱해 폭을 넉넉히 잡는다.
KOREAN_HEAVY_COLUMNS = {"제품명", "업체명", "포장단위"}
KOREAN_WIDTH_FACTOR = 1.15
MIN_WIDTH = 10
MAX_WIDTH = 60
WIDTH_PADDING = 3


def _auto_width(values, korean_heavy: bool) -> float:
    max_len = max((len(str(v)) for v in values if v not in (None, "")), default=MIN_WIDTH)
    factor = KOREAN_WIDTH_FACTOR if korean_heavy else 1.0
    return min(max(max_len * factor + WIDTH_PADDING, MIN_WIDTH), MAX_WIDTH)


def build_email_report(source_path: str, out_path: str, sheet_title: str = "CHC1_리포트"):
    src_wb = load_workbook(source_path, data_only=True)
    src_ws = src_wb.active

    header_row = [c.value for c in next(src_ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {name: header_row.index(name) for name in EMAIL_COLUMNS}

    data_rows = []
    for row in src_ws.iter_rows(min_row=2, values_only=True):
        data_rows.append([row[col_idx[name]] for name in EMAIL_COLUMNS])

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.append(EMAIL_COLUMNS)

    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN

    for data_row in data_rows:
        ws.append(data_row)

    ingr_col_num = EMAIL_COLUMNS.index(INGREDIENT_COLUMN) + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = WRAP_ALIGN if cell.column == ingr_col_num else PLAIN_ALIGN

    for i, name in enumerate(EMAIL_COLUMNS, start=1):
        col_letter = ws.cell(row=1, column=i).column_letter
        if name == INGREDIENT_COLUMN:
            ws.column_dimensions[col_letter].width = INGREDIENT_WIDTH
        else:
            values = [r[i - 1] for r in data_rows] + [name]
            ws.column_dimensions[col_letter].width = _auto_width(values, name in KOREAN_HEAVY_COLUMNS)

    wb.save(out_path)
    return len(data_rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="classify_chc1.py 결과 엑셀 경로")
    parser.add_argument("--out", required=True, help="메일 발송용 경량 엑셀 출력 경로")
    args = parser.parse_args()

    n = build_email_report(args.source, args.out)
    print(f"{n}건 -> {args.out} 저장 완료 (컬럼: {', '.join(EMAIL_COLUMNS)})")


if __name__ == "__main__":
    main()
