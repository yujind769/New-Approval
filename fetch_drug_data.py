#!/usr/bin/env python3
"""
공공데이터포털 - 의약품제품허가정보 서비스(DrugPrdtPrmsnInfoService07) 조회 스크립트

전체 페이지를 순회하며 데이터를 수집한 뒤,
- SPCLTY_PBLC == "일반의약품"
- ITEM_PERMIT_DATE 가 "202607" 로 시작
조건에 맞는 항목만 골라 엑셀 파일(ITEM_NAME/ENTP_NAME/ITEM_PERMIT_DATE/ITEM_INGR_NAME)로 저장한다.

사용법:
    python3 fetch_drug_data.py [--service-key KEY] [--out OUTPUT.xlsx]

기본 서비스키는 아래 DEFAULT_SERVICE_KEY(디코딩된 형태)를 사용한다.
"""
import argparse
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

from openpyxl import Workbook

BASE_URL = "https://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnDtlInq06"
DEFAULT_SERVICE_KEY = "0b9kSMyZHTA6Vzot3jmEbmeaUS4YpWUvBeMqSRuekpFY7lKRmQpRCirnEwN5nACMpC2b9X5ChC9Cz3xIu8P8dQ=="
NUM_OF_ROWS = 500
TARGET_CLASS = "일반의약품"
TARGET_PERMIT_PREFIX = "202607"
MAX_RETRIES = 4
RETRY_DELAY_SEC = 2


def fetch_page(service_key: str, page_no: int, num_of_rows: int) -> ET.Element:
    """지정 페이지를 XML로 요청하고 루트 Element를 반환한다."""
    params = {
        "serviceKey": service_key,
        "pageNo": str(page_no),
        "numOfRows": str(num_of_rows),
        "type": "xml",
    }
    url = f"{BASE_URL}?{urllib.parse.urlencode(params)}"

    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            with urllib.request.urlopen(url, timeout=60) as resp:
                data = resp.read()
            return ET.fromstring(data)
        except Exception as e:  # 네트워크 오류/타임아웃 등
            last_err = e
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY_SEC * attempt)
    raise RuntimeError(f"페이지 {page_no} 요청 실패 (재시도 {MAX_RETRIES}회 초과): {last_err}")


def parse_items(root: ET.Element):
    items = []
    for item_el in root.iter("item"):
        record = {child.tag: (child.text or "").strip() for child in item_el}
        items.append(record)
    return items


def get_total_count(root: ET.Element) -> int:
    total_el = root.find(".//totalCount")
    if total_el is not None and total_el.text and total_el.text.isdigit():
        return int(total_el.text)
    return 0


def get_result_code(root: ET.Element) -> str:
    code_el = root.find(".//resultCode")
    return code_el.text.strip() if code_el is not None and code_el.text else ""


def fetch_all_items(service_key: str, num_of_rows: int = NUM_OF_ROWS):
    all_items = []
    page_no = 1

    first_root = fetch_page(service_key, page_no, num_of_rows)
    result_code = get_result_code(first_root)
    if result_code not in ("", "00"):
        msg_el = first_root.find(".//resultMsg")
        msg = msg_el.text if msg_el is not None else "알 수 없는 오류"
        raise RuntimeError(f"API 오류 (resultCode={result_code}): {msg}")

    total_count = get_total_count(first_root)
    all_items.extend(parse_items(first_root))
    print(f"totalCount={total_count}, page {page_no} 수집: {len(all_items)}건")

    total_pages = (total_count + num_of_rows - 1) // num_of_rows if total_count else 1
    for page_no in range(2, total_pages + 1):
        root = fetch_page(service_key, page_no, num_of_rows)
        page_items = parse_items(root)
        all_items.extend(page_items)
        print(f"page {page_no}/{total_pages} 수집: 누적 {len(all_items)}건")

    return all_items


def filter_items(items):
    filtered = []
    for it in items:
        if it.get("SPCLTY_PBLC") == TARGET_CLASS and it.get("ITEM_PERMIT_DATE", "").startswith(TARGET_PERMIT_PREFIX):
            filtered.append(it)
    return filtered


def write_excel(items, out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "일반의약품_202607"

    headers = ["제품명", "업체명", "허가일자", "주성분"]
    ws.append(headers)

    for it in items:
        ws.append([
            it.get("ITEM_NAME", ""),
            it.get("ENTP_NAME", ""),
            it.get("ITEM_PERMIT_DATE", ""),
            it.get("ITEM_INGR_NAME", ""),
        ])

    widths = [40, 30, 14, 50]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--service-key", default=DEFAULT_SERVICE_KEY, help="공공데이터포털 서비스키 (디코딩된 형태)")
    parser.add_argument("--out", default="일반의약품_202607_허가목록.xlsx", help="출력 엑셀 파일 경로")
    parser.add_argument("--num-of-rows", type=int, default=NUM_OF_ROWS, help="페이지당 요청 건수")
    args = parser.parse_args()

    print("데이터 수집을 시작합니다...")
    items = fetch_all_items(args.service_key, args.num_of_rows)
    print(f"총 {len(items)}건 수집 완료")

    filtered = filter_items(items)
    print(f"조건(일반의약품 / 허가일자 202607*)에 맞는 항목: {len(filtered)}건")

    write_excel(filtered, args.out)
    print(f"엑셀 파일 저장 완료: {args.out}")


if __name__ == "__main__":
    sys.exit(main())
