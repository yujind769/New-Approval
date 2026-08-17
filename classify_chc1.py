#!/usr/bin/env python3
"""
월별 일반의약품 허가 데이터에 CHC1 코드를 매겨 엑셀로 저장한다.

fetch_drug_data.py로 해당 월 데이터를 받아온 뒤, chc1_classifier.py의
규칙(제품명/성분 완전일치 -> 한방 처방명 매칭 -> 성분 유사도 매칭)을
순서대로 적용한다. 규칙으로 해결되지 않는 항목은 CHC1을 비워두고
"LLM판단필요"로 표시한다 -- 이 항목들은 Claude(또는 사람)가 직접 검토해서
채워야 하는, 자동화할 수 없는 부분이다.

사용법:
    python3 classify_chc1.py [--year-month YYYYMM] [--reference data/chc1_reference.xlsx]
                              [--out OUTPUT.xlsx]
"""
import argparse

from openpyxl import Workbook

import fetch_drug_data as fd
import chc1_classifier as clf

CHC1_KOR = {
    "01_COUGH COLD&OTH RESP PROD": "01_기침, 감기 및 기타 호흡기용제",
    "02_PAIN RELIEF": "02_진통제",
    "03_DIGEST & OTH INTEST PROD": "03_소화기 및 기타 장 관련 제품",
    "04_VITAM.MINER.&NUTRIT.SUPPL": "04_비타민, 무기물 및 영양 보충제",
    "05_TONICS & OTHER STIMULANTS": "05_강장제 및 기타자극제",
    "06_SKIN TREATMENT": "06_피부 치료제",
    "07_EYE CARE": "07_눈 관리용제",
    "08_EAR CARE": "08_귀 관리용제",
    "09_MOUTH TREATMENT PRODUCTS": "09_구강 치료제",
    "10_CIRCULATORY PRODUCTS": "10_혈액 순환용제",
    "11_ANTINAUSEAN.": "11_항구토제",
    "12_URINARY&REPRODUCTIVE CARE": "12_비뇨기관 및 생식기관용제",
    "13_CALMING AND SLEEPING": "13_수면 유도 및 기분안정제",
    "14_WEIGHT MANAGEMENT PROD.": "14_체중 감소제",
    "17_HABIT TREATMENT": "17_습관 개선제",
    "18_MISCELLANEOUS": "18_기타",
    "49_PLASTERS": "49_첩부제",
    "82_BEAUTY PRODUCTS FOR WOMEN": "82_여성용 미용제제",
    "83_UNISEX BEAUTY PRODUCTS": "83_남녀공용 미용제제",
}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--service-key", default=fd.DEFAULT_SERVICE_KEY)
    parser.add_argument("--year-month", default=None, help="조회할 허가월(YYYYMM). 생략 시 지난달")
    parser.add_argument("--etc-otc-code", default=fd.DEFAULT_ETC_OTC_CODE)
    parser.add_argument("--reference", default=clf.DEFAULT_REFERENCE_PATH)
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    year_month = args.year_month or fd.default_year_month()
    out_path = args.out or f"CHC1_분류_{year_month}.xlsx"

    print(f"{year_month} 허가 데이터 수집 중...")
    items = fd.fetch_month_items(args.service_key, year_month)
    filtered = fd.filter_items(items, args.etc_otc_code)
    print(f"조건에 맞는 항목: {len(filtered)}건")

    print(f"참조 데이터 로딩 중: {args.reference}")
    ref = clf.load_reference(args.reference)

    wb = Workbook()
    ws = wb.active
    ws.title = f"CHC1_{year_month}"[:31]
    headers = ["제품명", "업체명", "허가일자", "주성분", "포장단위",
               "CHC1코드", "CHC1_한글명", "판정방법", "신뢰도", "근거메모"]
    ws.append(headers)

    need_review = 0
    for it in filtered:
        chc1, method, conf, note = clf.classify_item(it, ref)
        if chc1 is None:
            need_review += 1
        ws.append([
            it.get("ITEM_NAME", ""),
            it.get("ENTP_NAME", ""),
            it.get("ITEM_PERMIT_DATE", ""),
            fd.clean_ingredient_names(it.get("MAIN_ITEM_INGR", "")),
            it.get("PACK_UNIT", ""),
            chc1 or "",
            CHC1_KOR.get(chc1, "") if chc1 else "",
            method,
            conf,
            note,
        ])

    widths = [38, 26, 12, 45, 30, 26, 26, 22, 8, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    wb.save(out_path)
    print(f"엑셀 저장 완료: {out_path}")
    print(f"자동 분류 완료: {len(filtered) - need_review}/{len(filtered)}건, "
          f"검수 필요(LLM판단필요): {need_review}건")


if __name__ == "__main__":
    main()
