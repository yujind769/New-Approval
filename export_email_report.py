#!/usr/bin/env python3
"""
classify_chc1.py가 만든 전체(10컬럼, 판정근거 포함) 분류 결과에서
메일 발송용 6컬럼(제품명/업체명/허가일자/주성분/포장단위/CHC1코드)만
추려낸 경량 리포트를 만든다.

사용법:
    python3 export_email_report.py --source CHC1_분류_202607.xlsx --out CHC1_리포트_202607.xlsx
"""
import argparse

from openpyxl import Workbook, load_workbook

EMAIL_COLUMNS = ["제품명", "업체명", "허가일자", "주성분", "포장단위", "CHC1코드"]


def build_email_report(source_path: str, out_path: str, sheet_title: str = "CHC1_리포트"):
    src_wb = load_workbook(source_path, data_only=True)
    src_ws = src_wb.active

    header_row = [c.value for c in next(src_ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {name: header_row.index(name) for name in EMAIL_COLUMNS}

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.append(EMAIL_COLUMNS)

    row_count = 0
    for row in src_ws.iter_rows(min_row=2, values_only=True):
        ws.append([row[col_idx[name]] for name in EMAIL_COLUMNS])
        row_count += 1

    widths = [38, 26, 12, 45, 30, 28]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    wb.save(out_path)
    return row_count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="classify_chc1.py 결과 엑셀 경로")
    parser.add_argument("--out", required=True, help="메일 발송용 경량 엑셀 출력 경로")
    args = parser.parse_args()

    n = build_email_report(args.source, args.out)
    print(f"{n}건 -> {args.out} 저장 완료 (컬럼: {', '.join(EMAIL_COLUMNS)})")


if __name__ == "__main__":
    main()
