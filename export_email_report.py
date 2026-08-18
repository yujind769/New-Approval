#!/usr/bin/env python3
"""
classify_chc1.py가 만든 전체(10컬럼, 판정근거 포함) 분류 결과에서
메일 발송용 6컬럼만 추려낸 경량 리포트를 만든다.

컬럼 순서: 제품명 / 업체명 / 허가일자 / CHC1코드 / 주성분 / 포장단위

서식:
- 시트 제목: "{YY}년 {M}월 OTC" (year_month 지정 시)
- 헤더(1행): 굵게 + 회색 배경 + 가운데 정렬
- 폰트: 나눔고딕 (헤더 11pt / 본문 10pt) -- xlsx는 폰트 이름만 저장하므로
  받는 사람 PC에 나눔고딕이 없으면 엑셀이 자동으로 다른 폰트로 대체한다.
- 제품명/업체명/허가일자: 헤더+본문 모두 가운데 정렬
- 주성분: 줄바꿈 처리, 폭을 넉넉하게(기존 대비 1.5배)
- 포장단위: 줄바꿈 처리, 폭을 절반 수준으로 축소
- CHC1코드: 내용 길이에 맞춘 자동 폭(한 줄 표기)

사용법:
    python3 export_email_report.py --source CHC1_분류_202607.xlsx --out CHC1_리포트_202607.xlsx --year-month 202607
"""
import argparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

EMAIL_COLUMNS = ["제품명", "업체명", "허가일자", "CHC1코드", "주성분", "포장단위"]

FONT_NAME = "나눔고딕"
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True)
BODY_FONT = Font(name=FONT_NAME, size=10)
HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center")

# 컬럼별 본문 정렬: (가운데정렬 여부, 줄바꿈 여부)
COLUMN_ALIGN_SPEC = {
    "제품명": (True, True),
    "업체명": (True, False),
    "허가일자": (True, False),
    "CHC1코드": (True, False),
    "주성분": (False, True),
    "포장단위": (False, True),
}
FIXED_WIDTHS = {"주성분": 68, "포장단위": 30}  # 주성분: 기존 45의 1.5배 / 포장단위: 기존 60의 절반

KOREAN_HEAVY_COLUMNS = {"제품명", "업체명", "포장단위"}
KOREAN_WIDTH_FACTOR = 1.15
MIN_WIDTH = 10
MAX_WIDTH = 60
WIDTH_PADDING = 3


def _auto_width(values, korean_heavy: bool) -> float:
    max_len = max((len(str(v)) for v in values if v not in (None, "")), default=MIN_WIDTH)
    factor = KOREAN_WIDTH_FACTOR if korean_heavy else 1.0
    return min(max(max_len * factor + WIDTH_PADDING, MIN_WIDTH), MAX_WIDTH)


def _sheet_title(year_month: str | None) -> str:
    if not year_month:
        return "CHC1_리포트"
    yy = year_month[2:4]
    month = str(int(year_month[4:6]))
    return f"{yy}년 {month}월 OTC"


def build_email_report(source_path: str, out_path: str, year_month: str | None = None):
    src_wb = load_workbook(source_path, data_only=True)
    src_ws = src_wb.active

    header_row = [c.value for c in next(src_ws.iter_rows(min_row=1, max_row=1))]
    col_idx = {name: header_row.index(name) for name in EMAIL_COLUMNS}

    data_rows = []
    for row in src_ws.iter_rows(min_row=2, values_only=True):
        data_rows.append([row[col_idx[name]] for name in EMAIL_COLUMNS])

    wb = Workbook()
    ws = wb.active
    ws.title = _sheet_title(year_month)[:31]
    ws.append(EMAIL_COLUMNS)

    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN

    for data_row in data_rows:
        ws.append(data_row)

    col_align = {}
    for name in EMAIL_COLUMNS:
        centered, wrapped = COLUMN_ALIGN_SPEC[name]
        col_align[name] = Alignment(
            horizontal="center" if centered else None,
            vertical="center" if centered else "top",
            wrap_text=wrapped,
        )

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            name = EMAIL_COLUMNS[cell.column - 1]
            cell.font = BODY_FONT
            cell.alignment = col_align[name]

    for i, name in enumerate(EMAIL_COLUMNS, start=1):
        col_letter = ws.cell(row=1, column=i).column_letter
        if name in FIXED_WIDTHS:
            ws.column_dimensions[col_letter].width = FIXED_WIDTHS[name]
        else:
            values = [r[i - 1] for r in data_rows] + [name]
            ws.column_dimensions[col_letter].width = _auto_width(values, name in KOREAN_HEAVY_COLUMNS)

    wb.save(out_path)
    return len(data_rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="classify_chc1.py 결과 엑셀 경로")
    parser.add_argument("--out", required=True, help="메일 발송용 경량 엑셀 출력 경로")
    parser.add_argument("--year-month", default=None, help="시트 제목용 허가월(YYYYMM)")
    args = parser.parse_args()

    n = build_email_report(args.source, args.out, args.year_month)
    print(f"{n}건 -> {args.out} 저장 완료 (컬럼: {', '.join(EMAIL_COLUMNS)})")


if __name__ == "__main__":
    main()
