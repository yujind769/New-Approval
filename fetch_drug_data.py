#!/usr/bin/env python3
"""
공공데이터포털 - 의약품제품허가정보 서비스(DrugPrdtPrmsnInfoService07) 조회 스크립트

전체 페이지를 순회하며 데이터를 수집한 뒤,
- SPCLTY_PBLC == "일반의약품"
- ITEM_PERMIT_DATE 가 "202607" 로 시작
조건에 맞는 항목만 골라 엑셀 파일(ITEM_NAME/ENTP_NAME/ITEM_PERMIT_DATE/ITEM_INGR_NAME)로 저장한다.

페이지별로 수집한 원본 데이터를 --cache 경로에 JSON Lines로 즉시 저장하므로,
네트워크 오류 등으로 중간에 중단되어도 다시 실행하면 이미 받은 페이지는
건너뛰고 이어서 받는다(재시작 시 처음부터 다시 받지 않음).

사용법:
    python3 fetch_drug_data.py [--service-key KEY] [--out OUTPUT.xlsx] [--cache CACHE.jsonl]

기본 서비스키는 아래 DEFAULT_SERVICE_KEY(디코딩된 형태)를 사용한다.
"""
import argparse
import json
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
MAX_RETRIES = 8
REQUEST_TIMEOUT_SEC = 90
RETRY_DELAY_SEC = 3
RETRY_DELAY_CAP_SEC = 30


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
            with urllib.request.urlopen(url, timeout=REQUEST_TIMEOUT_SEC) as resp:
                data = resp.read()
            return ET.fromstring(data)
        except Exception as e:  # 네트워크 오류/타임아웃 등
            last_err = e
            if attempt < MAX_RETRIES:
                delay = min(RETRY_DELAY_SEC * attempt, RETRY_DELAY_CAP_SEC)
                print(f"  페이지 {page_no} 요청 실패({attempt}/{MAX_RETRIES}): {e} -> {delay}초 후 재시도")
                time.sleep(delay)
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


def load_cache(cache_path: str):
    """이미 저장된 페이지 캐시를 읽어 {page_no: items} 형태로 반환한다."""
    cached_pages = {}
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                rec = json.loads(line)
                cached_pages[rec["page"]] = rec["items"]
    except FileNotFoundError:
        pass
    return cached_pages


def append_cache(cache_path: str, page_no: int, items):
    with open(cache_path, "a", encoding="utf-8") as f:
        f.write(json.dumps({"page": page_no, "items": items}, ensure_ascii=False) + "\n")
        f.flush()


def fetch_all_items(service_key: str, num_of_rows: int, cache_path: str):
    cached_pages = load_cache(cache_path)
    if cached_pages:
        print(f"캐시에서 {len(cached_pages)}개 페이지 이어받기 감지")

    if 1 in cached_pages:
        first_root = None
        total_count = None
    else:
        first_root = fetch_page(service_key, 1, num_of_rows)
        result_code = get_result_code(first_root)
        if result_code not in ("", "00"):
            msg_el = first_root.find(".//resultMsg")
            msg = msg_el.text if msg_el is not None else "알 수 없는 오류"
            raise RuntimeError(f"API 오류 (resultCode={result_code}): {msg}")
        total_count = get_total_count(first_root)
        page_items = parse_items(first_root)
        append_cache(cache_path, 1, page_items)
        cached_pages[1] = page_items
        print(f"totalCount={total_count}, page 1 수집: {len(page_items)}건")

    if total_count is None:
        # 캐시에 이미 1페이지가 있는 경우, totalCount 파악을 위해 재조회하지 않고
        # 캐시된 페이지 수 기준으로 마지막 페이지까지 순차 진행 후 빈 응답에서 멈춘다.
        # 안전하게 다시 1페이지를 조회해 totalCount만 얻는다.
        probe_root = fetch_page(service_key, 1, num_of_rows)
        total_count = get_total_count(probe_root)

    total_pages = (total_count + num_of_rows - 1) // num_of_rows if total_count else 1

    for page_no in range(2, total_pages + 1):
        if page_no in cached_pages:
            continue
        root = fetch_page(service_key, page_no, num_of_rows)
        page_items = parse_items(root)
        append_cache(cache_path, page_no, page_items)
        cached_pages[page_no] = page_items
        collected = sum(len(v) for k, v in cached_pages.items() if k <= page_no)
        print(f"page {page_no}/{total_pages} 수집: 누적 {collected}건")

    all_items = []
    for page_no in range(1, total_pages + 1):
        all_items.extend(cached_pages.get(page_no, []))
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
    parser.add_argument("--cache", default=".drug_data_cache.jsonl", help="페이지별 원본 데이터 캐시 파일(중단 시 이어받기용)")
    args = parser.parse_args()

    print("데이터 수집을 시작합니다...")
    items = fetch_all_items(args.service_key, args.num_of_rows, args.cache)
    print(f"총 {len(items)}건 수집 완료")

    filtered = filter_items(items)
    print(f"조건(일반의약품 / 허가일자 202607*)에 맞는 항목: {len(filtered)}건")

    write_excel(filtered, args.out)
    print(f"엑셀 파일 저장 완료: {args.out}")


if __name__ == "__main__":
    sys.exit(main())
