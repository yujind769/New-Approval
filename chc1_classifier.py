#!/usr/bin/env python3
"""
CHC1(Consumer HealthCare 1단계) 코드 분류 모듈

data/chc1_reference.xlsx (Sheet1=CHC1 코드 체계, Sheet2=과거 라벨링 사례)를
참조 데이터로 사용해서, fetch_drug_data.py가 API에서 받아온 제품 항목에
CHC1 코드를 매기려고 시도한다.

분류 단계 (순서대로 시도, 먼저 성공하는 단계를 채택):
  1. 제품명 완전일치           -> 신뢰도 높음
  2. 정규화 성분셋 완전일치     -> 신뢰도 높음
  3. 외용 진통제 파스/플라스타 네이밍 휴리스틱 -> 신뢰도 중간
  4. 한방 복합제 2단계 로직     -> 신뢰도 높음(고전처방명 매칭) / 낮음(미인식 기본값 18)
  5. 성분셋 유사도(Jaccard) 매칭 -> 신뢰도 중간
  6. 위 전부 실패              -> CHC1 없음, "LLM판단필요"(신뢰도 낮음) 플래그만 남김
     (이 단계는 결정론적 규칙으로 대체할 수 없는, 사람 또는 LLM의 실제 판단이
     필요한 항목이다. 매달 이 스크립트를 돌린 뒤, "LLM판단필요"로 남은 행만
     남고 채워 넣는 반자동 워크플로를 전제로 한다.)

한방 복합제 판별은 성분명(영문)이 뿌리/줄기/열매 등 생약재 특유의 단어로만
구성돼 있는지를 휴리스틱으로 판단한다. 저명 처방명 목록(KNOWN_HERBAL_FORMULAS)은
발견되는 대로 계속 추가한다.

'파스'/'플라스타' 네이밍은 흔히 CHC1의 49_PLASTERS(첩부제)로 오인하기 쉽지만,
참조데이터상 이 네이밍의 외용 진통소염 제품(캄파/멘톨/살리실산 계열)은 예외 없이
02_PAIN RELIEF로 분류돼 있다 -- 49_PLASTERS는 티눈/각질 제거, 상처케어(액상밴드)
등 진통과 무관한 첩부제 전용 카테고리다. dl-/l- 이성질체 접두사 때문에 성분셋이
참조데이터와 문자열 그대로 일치하지 않아 아래 4/5단계에서 놓치는 경우가 많아
이름 기반 휴리스틱을 성분셋 매칭보다 먼저 시도한다.
"""
import re
from collections import Counter, defaultdict

from openpyxl import load_workbook

DEFAULT_REFERENCE_PATH = "data/chc1_reference.xlsx"

# --- 한방 복합제 판별/처방명 매칭 ---

HERBAL_PART_KEYWORDS = {
    "ROOT", "RHIZOME", "BARK", "SEED", "FRUIT", "TUBER", "HERB", "LEAF",
    "FLOWER", "PEEL", "RESIN", "RADIX", "RHIZOMA", "COLLA", "GALLSTONE",
    "MUSK", "BORNEOL", "POLLEN", "CORNU", "SCLEROTIUM", "EXTRACT",
    "OYSTER", "SHELL", "GINSENG", "STEM", "PERICARP", "KERNEL",
}

KNOWN_HERBAL_FORMULAS = [
    "천왕보심단", "반하사심탕", "평위산", "갈근탕", "구풍해독탕",
    "은교산", "소청룡탕", "쌍화탕", "우황청심원",
    # 새로 발견되는 저명 처방명은 여기 계속 추가
]


def is_herbal_complex(main_ingr_eng: str) -> bool:
    """영문 성분명이 3종 이상이고 대부분 생약재 표현이면 한방 '복합제'로 간주한다.
    단일/소수 성분의 표준화된 생약 추출물(예: 쿠쿠르비트종자유엑스 단일제)은
    '처방명'이 없는 서양 의약품에 가까운 판단 대상이라 제외한다 -- herbal-name
    lookup이 아니라 일반 성분 매칭/LLM 판단 경로로 보내야 한다."""
    if not main_ingr_eng:
        return False
    parts = [p.strip().upper() for p in main_ingr_eng.split("/") if p.strip()]
    if len(parts) < 3:
        return False
    herbal_like = sum(
        1 for p in parts if any(kw in p for kw in HERBAL_PART_KEYWORDS)
    )
    return herbal_like / len(parts) >= 0.6


def classify_herbal(product_name: str, sheet2_rows):
    """한방 복합제 2단계 로직.
    1) 저명 처방명이 제품명에 포함 -> 참조데이터에서 그 처방명의 실제 CHC1 분포를 다수결로 채택
    2) 못 찾으면 -> 18_MISCELLANEOUS, 신뢰도 낮음(검수 필요)
    """
    for formula in KNOWN_HERBAL_FORMULAS:
        if formula in product_name:
            matches = [r[1] for r in sheet2_rows if r[2] and formula in r[2]]
            if matches:
                counter = Counter(matches)
                top_chc, cnt = counter.most_common(1)[0]
                ratio = cnt / len(matches)
                return (
                    top_chc,
                    "고전처방명 매칭",
                    "높음" if ratio >= 0.9 else "중간",
                    f"'{formula}' 학습자료 {len(matches)}건 중 {ratio:.0%} 일치",
                )
    return (
        "18_MISCELLANEOUS",
        "한방처방 미인식(기본값)",
        "낮음",
        "학습자료에서 처방명을 찾지 못함 - 검수 필요",
    )


# --- 외용 진통제 파스/플라스타 판별 ---

PATCH_NAME_KEYWORDS = ("파스", "플라스타", "카타플라스마")
TOPICAL_ANALGESIC_KEYWORDS = {
    "CAMPHOR", "MENTHOL", "MENTHA", "SALICYL", "CAPSICUM", "CAPSAICIN",
    "NONIVAMIDE", "FELBINAC", "KETOPROFEN", "DICLOFENAC", "IBUPROFEN",
    "PIROXICAM", "INDOMETHACIN", "FLURBIPROFEN", "LOXOPROFEN",
    "HYDROXYTOLUIC", "THYMOL", "NICOTINIC",
}


def is_pain_relief_patch(product_name: str, main_ingr_eng: str) -> bool:
    """제품명에 파스/플라스타류 표현이 있고 외용 진통소염 성분이 주성분이면
    49_PLASTERS가 아니라 02_PAIN RELIEF로 봐야 하는 경우를 잡아낸다."""
    if not any(kw in product_name for kw in PATCH_NAME_KEYWORDS):
        return False
    eng_upper = (main_ingr_eng or "").upper()
    return any(kw in eng_upper for kw in TOPICAL_ANALGESIC_KEYWORDS)


# --- 성분명 정규화 (염/제형/농도 표기 차이 흡수) ---

NOISE_TOKENS = {
    "HYDROCHLORIDE", "HCL", "SODIUM", "POTASSIUM", "CALCIUM", "MAGNESIUM",
    "SULFATE", "SULPHATE", "MALEATE", "ACETATE", "CITRATE", "TARTRATE",
    "PHOSPHATE", "BROMIDE", "HYDROBROMIDE", "MESYLATE", "BESYLATE",
    "FUMARATE", "SUCCINATE", "GLUCONATE", "LACTATE", "NITRATE",
    "DIHYDRATE", "MONOHYDRATE", "HYDRATE", "ANHYDROUS", "POWDER",
    "DRIED", "EXTRACT", "CONCENTRATED", "CONC", "PURIFIED", "CRUDE",
    "COMPOUND", "SOLUTION", "GEL", "MIXTURE",
}


def normalize_ingredient(name: str) -> str:
    name = name.upper().strip()
    name = re.sub(r"\d+(\.\d+)?\s*%", " ", name)
    name = re.sub(r"[0-9]+(/[0-9]+)?", " ", name)
    name = re.sub(r"[().,\-]", " ", name)
    tokens = [t for t in name.split() if t and t not in NOISE_TOKENS]
    return " ".join(sorted(tokens))


def norm_ingr_set(raw: str, sep: str) -> frozenset:
    if not raw:
        return frozenset()
    parts = [p.strip() for p in raw.replace(sep, "/").split("/") if p.strip()]
    return frozenset(normalize_ingredient(p) for p in parts if normalize_ingredient(p))


def jaccard(a: frozenset, b: frozenset) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


# --- 참조 데이터 로딩 ---

class Reference:
    def __init__(self, sheet2_rows, name_to_chc, ingr_set_to_chc):
        self.sheet2_rows = sheet2_rows
        self.name_to_chc = name_to_chc
        self.ingr_set_to_chc = ingr_set_to_chc


def load_reference(path: str = DEFAULT_REFERENCE_PATH) -> Reference:
    wb = load_workbook(path, data_only=True)
    ws2 = wb["Sheet2"]
    rows = list(ws2.iter_rows(values_only=True))[2:]  # 헤더 2행 스킵

    name_to_chc = defaultdict(Counter)
    ingr_set_to_chc = defaultdict(Counter)
    for r in rows:
        if r[1] is None:
            continue
        chc1, name, _pack, molecule = r[1], r[2], r[3], r[4]
        if name:
            name_to_chc[name.strip()][chc1] += 1
        if molecule:
            s = norm_ingr_set(molecule, "+")
            if s:
                ingr_set_to_chc[s][chc1] += 1

    return Reference(rows, name_to_chc, ingr_set_to_chc)


# --- 메인 분류 함수 ---

FUZZY_THRESHOLD = 0.5


def classify_item(item: dict, ref: Reference):
    """item: fetch_drug_data.py가 파싱한 원본 dict (ITEM_NAME, MAIN_ITEM_INGR, MAIN_INGR_ENG 등).
    반환: (chc1코드 또는 None, 판정방법, 신뢰도, 근거메모)
    """
    name = (item.get("ITEM_NAME") or "").strip()
    eng = item.get("MAIN_INGR_ENG", "")

    # 1. 제품명 완전일치
    if name in ref.name_to_chc:
        counter = ref.name_to_chc[name]
        top_chc, cnt = counter.most_common(1)[0]
        ratio = cnt / sum(counter.values())
        return top_chc, "제품명 완전일치", "높음" if ratio >= 0.9 else "중간", f"학습자료 {sum(counter.values())}건 중 {ratio:.0%} 일치"

    # 2. 정규화 성분셋 완전일치
    key = norm_ingr_set(eng, "/")
    if key in ref.ingr_set_to_chc:
        counter = ref.ingr_set_to_chc[key]
        top_chc, cnt = counter.most_common(1)[0]
        ratio = cnt / sum(counter.values())
        return top_chc, "정규화 성분일치", "높음" if ratio >= 0.9 else "중간", f"학습자료 {sum(counter.values())}건 중 {ratio:.0%} 일치"

    # 3. 외용 진통제 파스/플라스타 네이밍 휴리스틱
    if is_pain_relief_patch(name, eng):
        return (
            "02_PAIN RELIEF",
            "파스/플라스타 네이밍 휴리스틱",
            "중간",
            "제품명에 파스/플라스타 계열 표현 + 캄파/멘톨/살리실산 등 외용진통성분 -- "
            "49_PLASTERS(첩부제)는 티눈/상처케어 전용이며, 참조데이터상 파스류는 전부 02_PAIN RELIEF",
        )

    # 4. 한방 복합제 2단계 로직
    if is_herbal_complex(eng):
        return classify_herbal(name, ref.sheet2_rows)

    # 5. 성분셋 유사도(Jaccard) 매칭
    best_score = 0.0
    best_chcs = []
    for s, counter in ref.ingr_set_to_chc.items():
        sc = jaccard(key, s)
        if sc > best_score:
            best_score = sc
            best_chcs = [counter]
        elif sc == best_score and sc > 0:
            best_chcs.append(counter)

    if best_score >= FUZZY_THRESHOLD:
        votes = Counter()
        for c in best_chcs:
            votes.update(c)
        top_chc, _ = votes.most_common(1)[0]
        return top_chc, f"유사매칭(J={best_score:.2f})", "중간", f"후보 {len(best_chcs)}개 다수결"

    # 5. 실패 -> LLM 판단 필요
    return None, "LLM판단필요", "낮음", f"자동 매칭 실패(최고 J={best_score:.2f}) - 검수 필요"
