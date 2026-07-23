# -*- coding: utf-8 -*-
"""
wines.xlsx('와인목록' 시트) → data/wines.json 변환
사용법:  python tools/convert_wines.py   (wine-map 폴더에서 실행)
- 필수 열 누락, 잘못된 지역/타입 값은 줄 번호와 함께 오류로 알려주고 변환을 중단합니다.
"""
import json, os, sys
import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE, "wines.xlsx")
OUT = os.path.join(BASE, "data", "wines.json")

REGION_KO_TO_ID = {
    "샤블리": "chablis", "코트 드 뉘": "cote-de-nuits", "코트 드 본": "cote-de-beaune",
    "코트 샬로네즈": "cote-chalonnaise", "마코네": "maconnais",
    "몽타뉴 드 랭스": "montagne-de-reims", "발레 드 라 마른": "vallee-de-la-marne",
    "코트 데 블랑": "cote-des-blancs", "코트 드 세잔": "cote-de-sezanne",
    "코트 데 바르": "cote-des-bar",
}
TYPE_KO_TO_ID = {"레드": "red", "화이트": "white", "스파클링": "sparkling", "로제": "rose"}

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["와인목록"]

wines, errors = [], []
for row in ws.iter_rows(min_row=2, values_only=True):
    if not any(row):
        continue
    r = [(str(v).strip() if v is not None else "") for v in (list(row) + [""] * 11)[:11]]
    name_ko, name_orig, prod_ko, prod_orig, region_ko, app, grape, type_ko, vintage, price, desc = r
    line = len(wines) + len(errors) + 2  # 대략적 줄 번호 표시용
    if not name_ko:
        continue
    missing = [label for label, v in
               [("와인명(한글)", name_ko), ("생산자(한글)", prod_ko),
                ("지역", region_ko), ("아펠라시옹", app), ("품종", grape), ("타입", type_ko)] if not v]
    if missing:
        errors.append(f"[{name_ko or '이름없음'}] 필수 항목 누락: {', '.join(missing)}")
        continue
    if region_ko not in REGION_KO_TO_ID:
        errors.append(f"[{name_ko}] 알 수 없는 지역: '{region_ko}' (드롭다운 목록에서 선택)")
        continue
    if type_ko not in TYPE_KO_TO_ID:
        errors.append(f"[{name_ko}] 알 수 없는 타입: '{type_ko}' (레드/화이트/스파클링/로제)")
        continue
    wines.append({
        "nameKo": name_ko, "nameOrig": name_orig,
        "producerKo": prod_ko, "producerOrig": prod_orig,
        "region": REGION_KO_TO_ID[region_ko], "appellation": app,
        "grape": grape, "type": TYPE_KO_TO_ID[type_ko],
        "vintage": vintage, "price": price, "desc": desc,
    })

if errors:
    print("!! 변환 중단 — 아래 오류를 고친 뒤 다시 실행하세요:")
    for e in errors:
        print("  -", e)
    sys.exit(1)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(wines, f, ensure_ascii=False, indent=1)
print(f"OK: {len(wines)}종 변환 완료 → {OUT}")
