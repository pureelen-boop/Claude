"""
전산원본 + 출하내역서 스캔 zip -> 업체명으로 파일명이 바뀐 zip 생성

아티팩트(웹 도구)의 JPEG 매칭 기능과 완전히 같은 규칙을 쓴다:
면세지급액/과세공급가액/부가세/과세지급액/총지급예정액계/순매출금액/매출수수료
중 3개 이상이 한 업체와만 동시에 일치해야 "확정"으로 이름을 바꾸고,
그렇지 않으면 "확인필요_" 접두사를 붙여 원본 그대로 zip에 남긴다.

파일명은 사용자가 메일에 첨부할 때 그 파일명이 메일 제목으로 그대로
쓰이기 때문에, "{번호}. {월}월 {업체명} 매출내역 입니다.jpg" 형식으로
맞춘다. 같은 업체 스캔이 여러 장(출하내역서가 여러 페이지로 나뉜 경우,
또는 중복 스캔)이면 번호를 "9-1", "9-2"처럼 세분화한다.

사용법:
    python rename_jpegs.py 전산원본.xls 출하내역서스캔.zip 결과.zip --date "7월 31일"
"""
import argparse
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

import pandas as pd

SOURCE_SHEET = "수수료매출내역(거래처별)"
MATCH_FIELDS = ["면세지급액", "과세공급가액", "부가세", "과세지급액", "총지급예정액계", "순매출금액", "매출수수료"]
NAME_FIELD = "업체명(텍스트 언급)"
FOOTER_RE = re.compile(r"공급자\s*\(?\s*출\s*하\s*자\s*\)?\s*[:：]?\s*(.+?)\s*\(?\s*인\s*\)?\s*$")
NUM_RE = re.compile(r"\d[\d,]{2,}")
# 업체 상호에 흔히 붙는 법인 형태 표현. 예를 들어 전산원본엔 "로움에스농업회사법인
# 유한회사(파파스컷)"로 등록돼 있어도, 출하내역서 위에는 "로움에스농업회사법인"까지만
# (풋터가 잘리거나 상품명에 그렇게) 찍혀 있는 경우가 있다. 이런 접두/접미 법인
# 표현과 괄호 안 내용을 떼어내고 남는 "핵심 상호"로 비교해야 실제로 매칭된다.
CORP_SUFFIXES = ["주식회사", "유한회사", "합자회사", "합명회사", "영농조합법인", "농업회사법인", "협동조합"]
PAREN_RE = re.compile(r"\([^)]*\)")
# 출하내역서 상단 "매출기간 : 2026년07월01일 ~ 2026년07월31일" 문구에서
# 종료일을 뽑는다. --date를 안 주면 이 종료일을 작성일자로 자동 사용한다.
DATE_RANGE_RE = re.compile(
    r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일\s*~\s*(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일"
)


def to_number(v):
    if v is None:
        return 0
    s = str(v).replace(",", "").strip()
    if s == "" or s.lower() == "nan":
        return 0
    try:
        return float(s)
    except ValueError:
        return 0


def load_vendors(xls_path):
    raw = pd.read_excel(xls_path, sheet_name=SOURCE_SHEET, header=None)
    data_rows = raw.iloc[2:].reset_index(drop=True)
    vendors = []
    for _, row in data_rows.iterrows():
        name = row[1]
        biz_no = row[2]
        if pd.isna(name) or str(name).strip() == "":
            continue
        if pd.isna(biz_no) or str(biz_no).strip() == "":
            continue  # 하이포스 합계행 제외
        taxed_total = to_number(row[21])
        supply = round(taxed_total / 1.1) if taxed_total else 0
        vat = taxed_total - supply
        vendors.append({
            "seq": len(vendors) + 1,  # 전산원본/결과물3와 같은 순서의 번호
            "name": str(name).strip(),
            "면세지급액": to_number(row[22]),
            "과세공급가액": supply,
            "부가세": vat,
            "과세지급액": taxed_total,
            "총지급예정액계": to_number(row[23]),
            "순매출금액": to_number(row[8]),
            "매출수수료": to_number(row[15]),
        })
    return vendors


def parse_month(date_str):
    if not date_str:
        return None
    m = re.search(r"(\d{1,2})\s*월", date_str)
    return int(m.group(1)) if m else None


def ocr_text(path):
    r = subprocess.run(["tesseract", str(path), "stdout", "-l", "kor"], capture_output=True, text=True)
    return r.stdout


def extract_footer_name(text):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in reversed(lines):
        if "공급자" in line and ("출하" in line or "인" in line):
            m = FOOTER_RE.search(line)
            if m:
                return m.group(1).strip()
    return None


def extract_end_date(text):
    """'매출기간 : YYYY년MM월DD일 ~ YYYY년MM월DD일'에서 종료일(월,일)을
    뽑는다. 관례상 결과물3의 작성일자 = 매출기간 종료일."""
    m = DATE_RANGE_RE.search(text)
    if not m:
        return None
    return int(m.group(5)), int(m.group(6))


def core_name(name):
    """상호에서 괄호 안 내용과 법인 형태 표현을 떼어낸 "핵심 상호".
    "로움에스농업회사법인 유한회사(파파스컷)" -> "로움에스", "주식회사매일봄" -> "매일봄"."""
    n = PAREN_RE.sub("", name)
    for suf in CORP_SUFFIXES:
        n = n.replace(suf, "")
    return n.strip()


def extract_numbers(text):
    nums = set()
    for m in NUM_RE.finditer(text):
        try:
            nums.add(float(m.group(0).replace(",", "")))
        except ValueError:
            pass
    return nums


def match_vendor(ocr_numbers, vendors, text):
    scored = []
    for v in vendors:
        fields = [f for f in MATCH_FIELDS if v[f] and v[f] in ocr_numbers]
        # 업체명(또는 핵심 상호)이 이미지 어디에든 그대로 적혀 있으면(풋터
        # 서명란뿐 아니라 상품명 줄에도 - 예: "매일봄" 브랜드 상품) 이것도
        # 하나의 매칭 항목으로 센다. 금액이 우연히 겹쳐도 이름이 안 겹치면
        # 걸러내는 안전장치 역할도 겸한다 (classify()의 상충 검사 참고).
        core = core_name(v["name"])
        if core and len(core) >= 2 and core in text:
            fields.append(NAME_FIELD)
        if fields:
            scored.append({"vendor": v["name"], "seq": v["seq"], "count": len(fields), "fields": fields})
    scored.sort(key=lambda s: -s["count"])
    return scored


def classify(ocr_name, scored):
    """(업체명, seq, 확정여부, 근거) 반환. 업체를 못 정하면 seq는 None."""
    if scored and scored[0]["count"] >= 3 and (len(scored) == 1 or scored[1]["count"] < scored[0]["count"]):
        s = scored[0]
        # 금액만으로 3개 이상 맞아떨어져도, 정작 이 업체 이름은 이미지 어디에도
        # 안 보이는데 다른 업체 이름은 보인다면 - 숫자가 우연히 겹쳤을 위험
        # 신호다. 실제로 "로움에스"(정육) 스캔이 순전히 숫자만으로 "스퀴즈"로
        # 잘못 확정된 사례가 있어서 추가한 안전장치. 이럴 땐 확정하지 않는다.
        if NAME_FIELD not in s["fields"]:
            conflicting = [x for x in scored if x is not s and NAME_FIELD in x["fields"]]
            if conflicting:
                c = conflicting[0]
                return s["vendor"], s["seq"], False, (
                    f"금액은 '{s['vendor']}'와 일치({s['count']}개)하지만 "
                    f"이미지엔 '{c['vendor']}'라는 이름이 보여 상충 — 확인 필요"
                )
        return s["vendor"], s["seq"], True, "·".join(s["fields"]) + f" 일치 ({s['count']}개)"
    if len(scored) > 1 and scored[1]["count"] == scored[0]["count"]:
        tied = [s for s in scored if s["count"] == scored[0]["count"]]
        # 동점이어도 그중 이름까지 일치하는 후보가 단 하나뿐이면 그걸로 정한다.
        name_tied = [s for s in tied if NAME_FIELD in s["fields"]]
        if len(name_tied) == 1:
            s = name_tied[0]
            return s["vendor"], s["seq"], True, "·".join(s["fields"]) + f" 일치 ({s['count']}개, 동점 후보 중 이름으로 확정)"
        tied_names = [s["vendor"] for s in tied]
        return scored[0]["vendor"], scored[0]["seq"], False, " / ".join(tied_names) + f" 중 어디인지 애매함 (각 {scored[0]['count']}개 항목 일치)"
    if scored:
        s = scored[0]
        return s["vendor"], s["seq"], False, f"'{s['vendor']}'로 추정 ({'·'.join(s['fields'])}만 일치, {s['count']}개뿐)"
    if ocr_name:
        return None, None, False, f"이름은 '{ocr_name}'로 읽혔지만 일치하는 금액이 없음"
    return None, None, False, "글자를 거의 읽지 못함 (스캔 상태 확인 필요)"


def build_filename(seq, page, page_count, month, vendor):
    """번호(-페이지). 월 업체명 매출내역 입니다.jpg — 메일 첨부 시 이 파일명이
    그대로 메일 제목으로 쓰이므로 이 형식을 임의로 바꾸지 않는다."""
    num = f"{seq}-{page}" if page_count > 1 else str(seq)
    month_part = f"{month}월 " if month else ""
    return f"{num}. {month_part}{vendor} 매출내역 입니다.jpg"


def main(xls_path, zip_path, out_zip_path, date_str=None):
    vendors = load_vendors(xls_path)
    print(f"전산원본에서 업체 {len(vendors)}곳 로드")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(zip_path) as zin:
            names = [n for n in zin.namelist() if not n.endswith("/") and re.search(r"\.(jpe?g|png)$", n, re.I)]
            zin.extractall(tmp, members=names)

        # 1차: 파일마다 OCR로 업체를 정한다 (아직 파일명은 정하지 않는다 —
        # 같은 업체가 몇 장인지 다 봐야 9-1/9-2 식 번호를 매길 수 있다).
        recognized = []  # dict: order, src, orig_base, vendor, seq, confirmed, reason
        for i, name in enumerate(names, start=1):
            src = tmp / name
            orig_base = Path(name).name
            print(f"[{i}/{len(names)}] {orig_base} 인식 중...")
            text = ocr_text(src)
            ocr_name = extract_footer_name(text)
            numbers = extract_numbers(text)
            scored = match_vendor(numbers, vendors, text)
            vendor, seq, confirmed, reason = classify(ocr_name, scored)
            recognized.append({
                "order": i, "src": src, "orig_base": orig_base,
                "vendor": vendor, "seq": seq, "confirmed": confirmed, "reason": reason,
                "date_detected": extract_end_date(text),
            })

        # 작성일자 결정: --date를 줬으면 그대로 쓰고, 안 줬으면 스캔들의
        # "매출기간" 종료일을 다수결로 자동 인식한다.
        if date_str:
            month = parse_month(date_str)
            dm = re.search(r"(\d{1,2})\s*일", date_str)
            day = int(dm.group(1)) if dm else None
            print(f"작성일자: {date_str} (직접 지정)")
        else:
            detected = [r["date_detected"] for r in recognized if r["date_detected"]]
            if detected:
                from collections import Counter
                (month, day), n = Counter(detected).most_common(1)[0]
                print(f"작성일자 자동 감지: {month}월 {day}일 "
                      f"(출하내역서 '매출기간' 종료일 기준, {n}/{len(detected)}장에서 인식)")
                if len(set(detected)) > 1:
                    print(f"  주의: 스캔들의 매출기간이 서로 다르게 인식됐습니다 {sorted(set(detected))} "
                          f"— 가장 많이 나온 값을 썼습니다. 결과가 이상하면 --date로 직접 지정하세요.")
            else:
                month, day = None, None
                print("경고: 출하내역서에서 '매출기간'을 인식하지 못해 작성일자를 자동으로 못 정했습니다. "
                      "--date를 직접 지정해주세요 (파일명에 'N월'도 안 붙습니다).")

        # 2차: 확정된 것들은 업체별로 묶어서 9-1/9-2 번호를 매긴다.
        confirmed_by_vendor = {}
        for r in recognized:
            if r["confirmed"] and r["vendor"]:
                confirmed_by_vendor.setdefault(r["vendor"], []).append(r)

        results = {}  # order -> (arcname, src, status, reason)
        for vendor, items in confirmed_by_vendor.items():
            seq = items[0]["seq"]
            for page, r in enumerate(items, start=1):
                arcname = build_filename(seq, page, len(items), month, vendor)
                results[r["order"]] = (arcname, r["src"], "확정", r["reason"])

        for r in recognized:
            if r["confirmed"] and r["vendor"]:
                continue  # 위에서 이미 처리
            arcname = (f"확인필요_{r['vendor']}_{r['orig_base']}" if r["vendor"]
                       else f"확인필요_{r['orig_base']}")
            results[r["order"]] = (arcname, r["src"], "확인 필요", r["reason"])

        report_rows = []
        out_entries = []  # (arcname, source_path)
        for order in sorted(results):
            arcname, src, status, reason = results[order]
            out_entries.append((arcname, src))
            report_rows.append({"원본파일명": recognized[order - 1]["orig_base"], "매칭결과": arcname, "상태": status, "근거": reason})

        with zipfile.ZipFile(out_zip_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for arcname, src in out_entries:
                zout.write(src, arcname)

    report_df = pd.DataFrame(report_rows)
    report_path = str(Path(out_zip_path).with_suffix("")) + "_매칭결과.csv"
    report_df.to_csv(report_path, index=False, encoding="utf-8-sig")

    confirmed_n = sum(1 for r in report_rows if r["상태"] == "확정")
    print(f"완료: 확정 {confirmed_n}장 / 확인 필요 {len(report_rows) - confirmed_n}장")
    print(f"zip: {out_zip_path}")
    print(f"매칭결과표: {report_path}")
    if month and day:
        print(f"작성일자(최종): {month}월 {day}일  # 결과물3 생성 시 이 값을 그대로 넘기세요")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("xls_path")
    parser.add_argument("zip_path")
    parser.add_argument("out_zip_path")
    parser.add_argument("--date", dest="date_str", default=None,
                         help='작성일자, 예: "7월 31일" — 파일명의 N월에 쓰인다')
    args = parser.parse_args()
    main(args.xls_path, args.zip_path, args.out_zip_path, args.date_str)
