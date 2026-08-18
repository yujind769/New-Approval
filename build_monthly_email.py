#!/usr/bin/env python3
"""
월별 CHC1 분류 결과로 메일 제목/본문을 렌더링하고, 발송용 6컬럼 엑셀을 만든다.
실제 발송(Gmail)은 이 스크립트가 하지 않는다 -- Gmail MCP 도구는 Claude
세션 안에서만 호출 가능하므로, 이 스크립트는 제목/본문 텍스트와 첨부파일
경로까지만 준비한다.

사용법:
    python3 build_monthly_email.py --source CHC1_분류_202607.xlsx --year-month 202607
"""
import argparse
from collections import Counter

from openpyxl import load_workbook

import export_email_report as eer

SUBJECT_TEMPLATE = "[공유] {year}년 {month}월 OTC 품목허가현황 공유의 건"

BODY_TEMPLATE = """안녕하세요, CH개발기획팀 AI봇입니다.

{year}년 {month}월 OTC 품목허가현황 공유드립니다.
{month}월 OTC 신규 허가 건수는 총 {total}건입니다.


[CHC1 분류 별 허가 건수]
{breakdown}
"""


def render_email(source_path: str, year_month: str):
    wb = load_workbook(source_path, data_only=True)
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    chc1_idx = header.index("CHC1코드")

    codes = [row[chc1_idx] for row in ws.iter_rows(min_row=2, values_only=True) if row[chc1_idx]]
    total = len(codes)
    counts = Counter(codes)

    breakdown_lines = [f"{code} - {n}건" for code, n in sorted(counts.items())]
    breakdown = "\n".join(breakdown_lines)

    year, month = year_month[:4], str(int(year_month[4:6]))

    subject = SUBJECT_TEMPLATE.format(year=year, month=month)
    body = BODY_TEMPLATE.format(year=year, month=month, total=total, breakdown=breakdown)
    return subject, body, total


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="classify_chc1.py 결과 엑셀 경로")
    parser.add_argument("--year-month", required=True, help="허가월(YYYYMM)")
    parser.add_argument("--attachment-out", default=None, help="발송용 6컬럼 엑셀 출력 경로")
    args = parser.parse_args()

    subject, body, total = render_email(args.source, args.year_month)

    attachment_out = args.attachment_out or f"OTC 신규 품목허가현황_{args.year_month}.xlsx"
    n = eer.build_email_report(args.source, attachment_out, args.year_month)

    print("=== SUBJECT ===")
    print(subject)
    print("=== BODY ===")
    print(body)
    print("=== ATTACHMENT ===")
    print(f"{attachment_out} ({n}건)")


if __name__ == "__main__":
    main()
