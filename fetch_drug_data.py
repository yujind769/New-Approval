#!/usr/bin/env python3
"""
공공데이터포털 - 의약품제품허가정보 서비스(DrugPrdtPrmsnInfoService07) 조회 스크립트

item_permit_date 파라미터(YYYYMM 6자리)로 서버측에서 해당 허가월만 걸러서
조회한다 (검증됨: item_permit_date=202607 -> totalCount=122, 요청 1회, 약 15초).
전체 42,968건을 페이지네이션으로 다 받아온 뒤 클라이언트에서 필터링하던
이전 방식보다 훨씬 빠르고, 매월 반복 실행하기에도 적합하다.

받아온 월별 데이터 중 ETC_OTC_CODE(전문/일반 구분)가 지정한 값과 일치하는
항목만 골라 엑셀 파일(제품명/업체명/허가일자/주성분/포장단위)로 저장한다.

사용법:
    python3 fetch_drug_data.py [--year-month YYYYMM] [--etc-otc-code 일반의약품]
                                [--service-key KEY] [--out OUTPUT.xlsx]

--year-month 를 생략하면 지난달(YYYYMM)을 자동으로 사용한다.
기본 서비스키는 아래 DEFAULT_SERVICE_KEY(디코딩된 형태)를 사용한다.
"""
import argparse
import datetime
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

from openpyxl import Workbook

BASE_URL = "https://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07/getDrugPrdtPrmsnDtlInq06"
DEFAULT_SERVICE_KEY = "0b9kSMyZHTA6Vzot3jmEbmeaUS4YpWUvBeMqSRuekpFY7lKRmQpRCirnEwN5nACMpC2b9X5ChC9Cz3xIu8P8dQ=="
NUM_OF_ROWS = 500
DEFAULT_ETC_OTC_CODE = "일반의약품"
MAX_RETRIES = 5
REQUEST_TIMEOUT_SEC = 90
RETRY_DELAY_SEC = 3
RETRY_DELAY_CAP_SEC = 20
FALLBACK_NUM_OF_ROWS = 100  # 월별 건수가 많아 500 요청이 실패할 때의 보조 크기


def fetch_page_raw(service_key: str, page_no: int, num_of_rows: int, year_month: str) -> ET.Element:
    """지정 pageNo/numOfRows로 item_permit_date=year_month 조회를 재시도와 함께 수행한다."""
    params = {
        "serviceKey": service_key,
        "pageNo": str(page_no),
        "numOfRows": str(num_of_rows),
        "type": "xml",
        "item_permit_date": year_month,
    }
    url = f"{BASE_URL}?{urllib.parse.urlencode(params)}"

    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            with urllib.request.urlopen(url, timeout=REQUEST_TIMEOUT_SEC) as resp:
                data = resp.read()
            return ET.fromstring(data)
        except Exception as e:  # 네트워크 오류/타임아웃/504 등
            last_err = e
            if attempt < MAX_RETRIES:
                delay = min(RETRY_DELAY_SEC * attempt, RETRY_DELAY_CAP_SEC)
                print(f"  pageNo={page_no} numOfRows={num_of_rows} 요청 실패({attempt}/{MAX_RETRIES}): {e} -> {delay}초 후 재시도")
                time.sleep(delay)
    raise RuntimeError(f"pageNo={page_no} numOfRows={num_of_rows} 요청 실패 (재시도 {MAX_RETRIES}회 초과): {last_err}")


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


def fetch_month_items(service_key: str, year_month: str, num_of_rows: int = NUM_OF_ROWS):
    try:
        first_root = fetch_page_raw(service_key, 1, num_of_rows, year_month)
    except Exception as e:
        if num_of_rows == FALLBACK_NUM_OF_ROWS:
            raise
        print(f"  numOfRows={num_of_rows} 실패, numOfRows={FALLBACK_NUM_OF_ROWS}로 축소 재시도: {e}")
        return fetch_month_items(service_key, year_month, FALLBACK_NUM_OF_ROWS)

    result_code = get_result_code(first_root)
    if result_code not in ("", "00"):
        msg_el = first_root.find(".//resultMsg")
        msg = msg_el.text if msg_el is not None else "알 수 없는 오류"
        raise RuntimeError(f"API 오류 (resultCode={result_code}): {msg}")

    total_count = get_total_count(first_root)
    items = parse_items(first_root)
    print(f"item_permit_date={year_month} totalCount={total_count}, page 1 수집: {len(items)}건")

    total_pages = (total_count + num_of_rows - 1) // num_of_rows if total_count else 1
    for page_no in range(2, total_pages + 1):
        root = fetch_page_raw(service_key, page_no, num_of_rows, year_month)
        page_items = parse_items(root)
        items.extend(page_items)
        print(f"page {page_no}/{total_pages} 수집: 누적 {len(items)}건")

    return items


def filter_items(items, etc_otc_code: str):
    return [it for it in items if it.get("ETC_OTC_CODE") == etc_otc_code]


INGR_CODE_PREFIX_RE = re.compile(r"^\[[^\]]*\]")


def clean_ingredient_names(raw: str) -> str:
    """MAIN_ITEM_INGR의 '[M223062]브롬헥신염산염|[M223211]...' 형식에서
    성분 코드를 제거하고 ', '로 구분된 성분명만 남긴다."""
    if not raw:
        return ""
    names = [INGR_CODE_PREFIX_RE.sub("", part).strip() for part in raw.split("|")]
    return ", ".join(name for name in names if name)


def write_excel(items, out_path: str, sheet_title: str):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]  # 엑셀 시트명 31자 제한

    headers = ["제품명", "업체명", "허가일자", "주성분", "포장단위"]
    ws.append(headers)

    for it in items:
        ws.append([
            it.get("ITEM_NAME", ""),
            it.get("ENTP_NAME", ""),
            it.get("ITEM_PERMIT_DATE", ""),
            clean_ingredient_names(it.get("MAIN_ITEM_INGR", "")),
            it.get("PACK_UNIT", ""),
        ])

    widths = [40, 30, 14, 50, 40]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

    wb.save(out_path)


def default_year_month() -> str:
    """지난달을 YYYYMM 형식으로 반환한다."""
    today = datetime.date.today()
    last_day_of_prev_month = today.replace(day=1) - datetime.timedelta(days=1)
    return last_day_of_prev_month.strftime("%Y%m")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--service-key", default=DEFAULT_SERVICE_KEY, help="공공데이터포털 서비스키 (디코딩된 형태)")
    parser.add_argument("--year-month", default=None, help="조회할 허가월(YYYYMM). 생략 시 지난달")
    parser.add_argument("--etc-otc-code", default=DEFAULT_ETC_OTC_CODE, help="필터링할 전문/일반 구분 값 (기본: 일반의약품)")
    parser.add_argument("--out", default=None, help="출력 엑셀 파일 경로 (생략 시 자동 생성)")
    parser.add_argument("--num-of-rows", type=int, default=NUM_OF_ROWS, help="페이지당 요청 건수")
    args = parser.parse_args()

    year_month = args.year_month or default_year_month()
    out_path = args.out or f"{args.etc_otc_code}_{year_month}_허가목록.xlsx"

    print(f"{year_month} 허가 데이터 수집을 시작합니다...")
    items = fetch_month_items(args.service_key, year_month, args.num_of_rows)
    print(f"총 {len(items)}건 수집 완료")

    filtered = filter_items(items, args.etc_otc_code)
    print(f"조건({args.etc_otc_code} / 허가월 {year_month})에 맞는 항목: {len(filtered)}건")

    write_excel(filtered, out_path, f"{args.etc_otc_code}_{year_month}")
    print(f"엑셀 파일 저장 완료: {out_path}")


if __name__ == "__main__":
    sys.exit(main())
