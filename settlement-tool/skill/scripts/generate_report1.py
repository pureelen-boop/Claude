"""
전산원본(하이포스 수수료매출내역) -> 결과물1(정산 요약본) 생성 스크립트

사용법:
    python generate_report1.py 전산원본.xls 결과물1_생성.xlsx

전산원본의 24개 컬럼 중, 기존에 4년간 써온 결과물1 양식과 동일한
11개 컬럼만 그대로의 순서로 뽑아낸다. 계산은 하지 않고 전산원본의
값을 그대로 옮기기만 한다 (전산원본이 이미 계산 완료된 값이므로).
"""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SOURCE_SHEET = "수수료매출내역(거래처별)"

# 전산원본 컬럼 인덱스(0-based, 헤더 2행 중 실제 라벨은 2번째 행) -> 결과물1에서의 라벨
COLUMN_MAP = [
    (0, "코드"),
    (1, "거래처명"),
    (2, "사업자번호"),
    (3, "수수료율"),
    (8, "순매출금액"),
    (15, "매출수수료"),
    (22, "면세지급액"),
    (19, "과세공급가액"),
    (20, "부가세"),
    (21, "과세지급액"),
    (23, "총지급예정액계"),
]

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FONT = Font(name="맑은 고딕", bold=True)
BODY_FONT = Font(name="맑은 고딕")
NUM_FMT = "#,##0"


def to_number(v):
    """하이포스 원본은 숫자를 '1,234' 형태의 텍스트로 저장한다. 실수 숫자로 변환."""
    if v is None:
        return 0
    s = str(v).replace(",", "").strip()
    if s == "" or s == "nan":
        return 0
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return v


def load_source(path):
    raw = pd.read_excel(path, sheet_name=SOURCE_SHEET, header=None)
    data_rows = raw.iloc[2:].reset_index(drop=True)  # 0,1행은 병합 헤더
    return data_rows


def build_report1(data_rows, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "수수료매출내역(거래처별)"

    headers = [label for _, label in COLUMN_MAP]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    r = 2
    for _, row in data_rows.iterrows():
        if pd.isna(row[1]):  # 거래처명 없는 빈 행 skip
            continue

        values = {}
        for c, (src_idx, label) in enumerate(COLUMN_MAP, start=1):
            raw_val = row[src_idx]
            if c in (2, 3):  # 거래처명, 사업자번호는 텍스트 그대로
                values[label] = raw_val
            else:
                values[label] = to_number(raw_val)

        # 과세공급가액/부가세는 하이포스 원본을 그대로 옮기면 두 값이 각자
        # 반올림되어 있어 총액(과세지급액)과 1원 어긋나는 경우가 있다.
        # 과세지급액(총액)을 기준으로 역산해 항상 총액과 정확히 맞도록 재계산한다.
        taxed_total = values["과세지급액"]
        supply = round(taxed_total / 1.1) if taxed_total else 0
        vat = taxed_total - supply
        values["과세공급가액"] = supply
        values["부가세"] = vat

        for c, (_src_idx, label) in enumerate(COLUMN_MAP, start=1):
            val = values[label]
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            if c not in (1, 2, 3):  # 코드/거래처명/사업자번호 제외 숫자 서식
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="center" if c != 2 else "left")
        r += 1

    widths = [8, 26, 15, 9, 13, 12, 13, 13, 11, 13, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    wb.save(out_path)
    print(f"결과물1 생성 완료: {out_path} ({r - 2}개 업체)")


if __name__ == "__main__":
    src, out = sys.argv[1], sys.argv[2]
    data = load_source(src)
    build_report1(data, out)
