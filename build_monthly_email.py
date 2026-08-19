#!/usr/bin/env python3
"""
월별 CHC1 분류 결과로 메일 제목/본문(텍스트+HTML)을 렌더링하고, 발송용
6컬럼 엑셀을 만든다. 실제 발송(Gmail)은 이 스크립트가 하지 않는다 --
Gmail MCP 도구는 Claude 세션 안에서만 호출 가능하므로, 이 스크립트는
제목/본문/첨부파일 경로까지만 준비한다.

HTML 본문은 <p> 태그(기본 여백 때문에 줄간격이 벌어져 보임) 대신,
텍스트 버전과 완전히 동일한 줄 단위(<br>)로 그대로 옮겨서 두 버전의
줄바꿈/여백이 항상 일치하도록 한다. [CHC1 분류 별 허가 건수] 줄만 굵게,
마지막 안내문구 줄만 작은 회색 글씨로 스타일을 얹는다.

사용법:
    python3 build_monthly_email.py --source CHC1_분류_202607.xlsx --year-month 202607
"""
import argparse
import html
from collections import Counter

from openpyxl import load_workbook

import export_email_report as eer

SUBJECT_TEMPLATE = "[공유] {year}년 {month}월 OTC 품목허가현황 공유의 건"

BODY_FONT_SIZE_PT = 11
FOOTER_FONT_SIZE_PT = 9
FOOTER_COLOR = "#888888"
HEADING_LINE = "[CHC1 분류 별 허가 건수]"
DIVIDER = "-" * 40  # 텍스트 버전의 구분선; HTML에서는 이 줄을 <hr>로 치환한다
FOOTER_LINES = [
    "⚠️ 본 메일은 발신 전용으로 자동 발송되었습니다(회신 불가).",
    "✅ 시스템 문의: CH개발기획팀 장유진 <yujin00@daewoong.co.kr>",
]


def _breakdown(source_path: str):
    wb = load_workbook(source_path, data_only=True)
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    chc1_idx = header.index("CHC1코드")

    codes = [row[chc1_idx] for row in ws.iter_rows(min_row=2, values_only=True) if row[chc1_idx]]
    total = len(codes)
    counts = Counter(codes)
    lines = [f"{code} - {n}건" for code, n in sorted(counts.items())]
    return total, lines


def _body_lines(year: str, month: str, total: int, breakdown_lines: list[str]) -> list[str]:
    """본문을 줄 단위 리스트로 만든다 (텍스트/HTML 버전이 항상 같은 줄 구조를 쓰도록)."""
    return [
        "안녕하세요,",
        "CH개발기획팀 AI봇 - 허가현황 알리미 🤖 입니다.",
        "",
        f"{year}년 {month}월 OTC 품목허가현황 공유드립니다.",
        f"{month}월 OTC 신규 허가 건수는 총 {total}건입니다.",
        "",
        "",
        HEADING_LINE,
        *breakdown_lines,
        "",
        DIVIDER,
        *FOOTER_LINES,
    ]


def render_email(source_path: str, year_month: str):
    total, breakdown_lines = _breakdown(source_path)
    year, month = year_month[:4], str(int(year_month[4:6]))
    lines = _body_lines(year, month, total, breakdown_lines)

    subject = SUBJECT_TEMPLATE.format(year=year, month=month)
    body = "\n".join(lines) + "\n"

    html_parts = []
    for i, line in enumerate(lines):
        if line == DIVIDER:
            html_parts.append('<hr style="border:none;border-top:1px solid #dddddd;margin:12px 0;">')
            continue
        escaped = html.escape(line) if line else "&nbsp;"
        if line == HEADING_LINE:
            escaped = f"<b>{escaped}</b>"
        elif line in FOOTER_LINES:
            escaped = f'<span style="font-size:{FOOTER_FONT_SIZE_PT}pt;color:{FOOTER_COLOR};">{escaped}</span>'
        is_last = i == len(lines) - 1
        html_parts.append(escaped if is_last else escaped + "<br>")

    html_body = (
        f'<div style="font-family:\'나눔고딕\',sans-serif;font-size:{BODY_FONT_SIZE_PT}pt;color:#000000;">\n'
        + "\n".join(html_parts)
        + "\n</div>"
    )
    return subject, body, html_body, total


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, help="classify_chc1.py 결과 엑셀 경로")
    parser.add_argument("--year-month", required=True, help="허가월(YYYYMM)")
    parser.add_argument("--attachment-out", default=None, help="발송용 6컬럼 엑셀 출력 경로")
    args = parser.parse_args()

    subject, body, html_body, total = render_email(args.source, args.year_month)

    attachment_out = args.attachment_out or f"OTC 신규 품목허가현황_{args.year_month}.xlsx"
    n = eer.build_email_report(args.source, attachment_out, args.year_month)

    print("=== SUBJECT ===")
    print(subject)
    print("=== BODY (plain text fallback) ===")
    print(body)
    print("=== HTML BODY ===")
    print(html_body)
    print("=== ATTACHMENT ===")
    print(f"{attachment_out} ({n}건)")


if __name__ == "__main__":
    main()
