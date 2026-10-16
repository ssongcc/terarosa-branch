# -*- coding: utf-8 -*-
"""
지점 원두 발주서 → 우체국 택배 접수파일 변환 로직 (앱/스크립트 공용)

규칙
  1) [어센틱에스프레소블렌드 1kg] : 지점별 수량(=kg) → 15kg 기준 박스 분할(올림), 박스별 고르게 분배
  2) 그 외 품명 뒤 중량(g/kg)이 붙은 원두 : 지점별 중량 합산 → 15kg 기준 박스 분할(올림)
  3) 1박스 = 1행, 수취인명 뒤 (박스번호/총박스)  예) (1/3) (2/3) (3/3), 1박스면 (1/1)
  4) 어센틱 박스 품목(F열) = '어센틱에스프레소블렌드1kg 10kg' 처럼 박스별 중량 표기
  5) 강릉 지점(본점·경포·임당·사천·신라모노), 주소파일에 없는 지점은 제외
"""
import io, re, math, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BOX_KG = 15
AUTHENTIC = "어센틱에스프레소블렌드1kg"           # 공백 제거 후 비교
AUTHENTIC_LABEL = "어센틱에스프레소블렌드1kg"     # 송장 품목(F열) 표기
EXCLUDE = {"본점", "경포", "임당", "사천", "신라모노"}
# 발주서 헤더 → 주소파일 수취인명에 포함된 키워드 (헤더와 다를 때만)
ALIAS = {
    "서종": "서종점", "예당": "예술의전당", "양재": "양재역", "포스코": "포스코센터",
    "현대": "국립현대미술관", "중문": "중문", "청계": "청계광장", "구의": "구의역",
    "잠실한강": "잠실한강공원", "역삼": "역삼역", "용산": "용산아이파크",
    "스타필드하남": "스타필드하남", "청주": "청주지웰", "교육서비스": "교육서비스",
    "IFC": "IFC몰", "현대송도": "현대아울렛송도", "현대동대문": "현대아울렛동대문",
    "성수": "성수", "신세계대구": "신세계대구", "청량리": "청량리",
}
WEIGHT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(kg|g)", re.I)
HEADERS = ["수취인명", "우편번호", "상세주소", "연락처 1 일반전화", "연락처 2 휴대폰", "품목", "메모"]


def s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).replace("\xa0", " ").replace("\n", " ").strip()


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def boxes(kg):
    kg = round(kg, 6)
    return math.ceil(kg / BOX_KG - 1e-9) if kg > 0 else 0


def split_even(kg, n):
    """kg를 n박스에 고르게 분배. 정수 kg면 정수로, 남는 1kg은 앞 박스부터"""
    if n <= 0:
        return []
    kg = round(kg, 6)
    if float(kg).is_integer():
        base, rem = divmod(int(kg), n)
        return [base + (1 if i < rem else 0) for i in range(n)]
    part = round(kg / n, 2)
    return [part] * (n - 1) + [round(kg - part * (n - 1), 2)]


ADDR_KEYS = ["name", "zip", "addr", "tel1", "tel2", "item", "memo"]


def load_addresses(src):
    """src: 경로 / file-like(xlsx) / 주소 리스트(Secrets의 [[addresses]])"""
    if isinstance(src, (list, tuple)):
        rows = []
        for a in src:
            row = [s(a.get(k, "")) for k in ADDR_KEYS] if hasattr(a, "get") else [s(x) for x in list(a)[:7]]
            if row[0]:
                rows.append(row + [""] * (7 - len(row)))
        return rows
    ws = openpyxl.load_workbook(src, data_only=True).worksheets[0]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not s(r[0]):
            continue
        rows.append([s(r[i]) if i < len(r) else "" for i in range(7)])
    return rows


def match_address(header, addrs):
    key = ALIAS.get(header, header)
    hits = [a for a in addrs if key in a[0].replace("[[원두]테라로사", "")]
    if len(hits) == 1:
        return hits[0], ""
    return None, ("주소 미등록 → 제외" if not hits else "주소 중복매칭 → 확인필요")


def process(order_src, addr_src):
    """
    반환 dict
      date(date|None), mmdd(str), upload(list[list]), summary(list[list]),
      skipped(list[str]), missing(list[(지점, 어센틱kg, 기타kg)])
    summary 행: [발주서지점, 수취인명, 어센틱kg, 어센틱박스, 기타kg, 기타박스, 총박스, 비고]
    """
    ws = openpyxl.load_workbook(order_src, data_only=True).worksheets[0]

    hrow = next((r for r in range(1, 21)
                 if s(ws.cell(r, 1).value).replace(" ", "") == "품명"), None)
    if hrow is None:
        raise ValueError("발주서에서 '품   명' 헤더행을 찾지 못했습니다. 지점 원두 발주서가 맞는지 확인하세요.")

    branches = [(c, s(ws.cell(hrow, c).value)) for c in range(12, ws.max_column + 1)
                if s(ws.cell(hrow, c).value)]          # L열(본점)부터

    auth_rows, other_rows, skipped = [], [], []
    for r in range(hrow + 1, ws.max_row + 1):
        pname = s(ws.cell(r, 1).value)
        if not pname:
            continue
        if pname.replace(" ", "") == AUTHENTIC:
            auth_rows.append(r)
            continue
        m = WEIGHT_RE.search(pname)
        if m:
            w = float(m.group(1)) / (1000 if m.group(2).lower() == "g" else 1)
            other_rows.append((r, w))
        else:
            skipped.append(pname)
    if not auth_rows:
        raise ValueError("발주서에 '어센틱에스프레소블렌드 1kg' 행이 없습니다.")

    addrs = load_addresses(addr_src)

    d = ws.cell(hrow - 2, 1).value
    date = d.date() if isinstance(d, datetime.datetime) else (d if isinstance(d, datetime.date) else None)
    mmdd = date.strftime("%m%d") if date else datetime.date.today().strftime("%m%d")

    upload, summary, missing = [], [], []
    for c, bname in branches:
        a_kg = round(sum(num(ws.cell(r, c).value) for r in auth_rows), 6)
        o_kg = round(sum(num(ws.cell(r, c).value) * w for r, w in other_rows), 6)
        a_box, o_box = boxes(a_kg), boxes(o_kg)
        if bname in EXCLUDE:
            summary.append([bname, "", a_kg, a_box, o_kg, o_box, 0, "강릉 지점 제외"])
            continue
        addr, note = match_address(bname, addrs)
        if addr is None:
            if a_kg + o_kg == 0:
                note += " (발주없음)"
            else:
                missing.append((bname, a_kg, o_kg))
            summary.append([bname, "", a_kg, a_box, o_kg, o_box, 0, note])
            continue
        total = a_box + o_box
        items = [f"{AUTHENTIC_LABEL} {q:g}kg" for q in split_even(a_kg, a_box)] + [addr[5]] * o_box
        for i, item in enumerate(items, 1):
            upload.append([f"{addr[0]}({i}/{total})"] + addr[1:5] + [item, addr[6]])
        summary.append([bname, addr[0], a_kg, a_box, o_kg, o_box, total, "" if total else "발주없음"])

    return dict(date=date, mmdd=mmdd, upload=upload, summary=summary,
                skipped=skipped, missing=missing)


def build_workbook(res, src_name=""):
    """결과 dict → xlsx bytes"""
    upload, summary, skipped = res["upload"], res["summary"], res["skipped"]
    wb = openpyxl.Workbook()
    f10 = Font(name="맑은 고딕", size=10)

    # ── 업로드용 (우체국 업로드 양식) ──
    ws = wb.active
    ws.title = "업로드용"
    ws.append(HEADERS)
    for row in upload:
        ws.append([v if v != "" else None for v in row])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = f10
            cell.number_format = "@"
            cell.alignment = Alignment(horizontal="left")
    for col, w in zip("ABCDEFG", [38, 8, 70, 16, 16, 32, 40]):
        ws.column_dimensions[col].width = w

    # ── 박스산출내역 (검증용) ──
    TERRA, LINEN, CREAM, OFF, CHAR, STONE = "8B3A2A", "EDE5DC", "FBF8F5", "F5F0EB", "2C2C2C", "D6CEC8"
    bd = Border(*(Side(style="thin", color=STONE),) * 4)
    ws2 = wb.create_sheet("박스산출내역")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:H1")
    ws2["A1"] = f"지점 원두 우체국 박스 산출내역  ({src_name})"
    ws2["A1"].font = Font(name="맑은 고딕", size=14, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill("solid", fgColor=TERRA)
    ws2.row_dimensions[1].height = 26
    ws2.merge_cells("A2:H2")
    ws2["A2"] = "기준: 15kg/박스 올림 · 어센틱에스프레소블렌드 1kg 별도 박스(고르게 분배) · 그 외 중량표기 원두 합산"
    ws2["A2"].font = Font(name="맑은 고딕", size=9, color="6B6056")
    ws2.append([])
    ws2.append(["발주서 지점", "수취인명", "어센틱1kg(kg)", "어센틱 박스", "기타원두(kg)", "기타 박스", "총 박스", "비고"])
    hr = ws2.max_row
    for c in range(1, 9):
        cell = ws2.cell(hr, c)
        cell.font = Font(name="맑은 고딕", size=9, bold=True, color=CHAR)
        cell.fill = PatternFill("solid", fgColor=LINEN)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = bd
    start = hr + 1
    for i, row in enumerate(summary):
        ws2.append(row)
        r = ws2.max_row
        fill = PatternFill("solid", fgColor=CREAM if i % 2 == 0 else OFF)
        if "미등록" in row[7] and row[2] + row[4] > 0:
            fill = PatternFill("solid", fgColor="F0DDD7")
        for c in range(1, 9):
            cell = ws2.cell(r, c)
            bold = c == 7 and bool(row[6])
            cell.font = Font(name="맑은 고딕", size=9, bold=bold, color=TERRA if bold else CHAR)
            cell.fill = fill
            cell.border = bd
            if c in (3, 5):
                cell.number_format = "0.##"
            if 3 <= c <= 7:
                cell.alignment = Alignment(horizontal="right")
    end = ws2.max_row
    ws2.append(["합계 (접수대상)", "",
                f'=SUMIFS(C{start}:C{end},G{start}:G{end},">0")', f'=SUMIFS(D{start}:D{end},G{start}:G{end},">0")',
                f'=SUMIFS(E{start}:E{end},G{start}:G{end},">0")', f'=SUMIFS(F{start}:F{end},G{start}:G{end},">0")',
                f"=SUM(G{start}:G{end})", f"업로드용 행수: {len(upload)}"])
    # 강릉 지점 소계 / 전체 총량(강릉 포함)
    ws2.append(["강릉 지점 소계", "",
                f'=SUMIFS(C{start}:C{end},H{start}:H{end},"강릉 지점 제외")', "",
                f'=SUMIFS(E{start}:E{end},H{start}:H{end},"강릉 지점 제외")', "", "", "우체국 접수 제외"])
    ws2.append(["전체 총량 (강릉 포함)", "",
                f"=SUM(C{start}:C{end})", "", f"=SUM(E{start}:E{end})", "", "", "발주서 전 지점 합계"])
    for r in range(ws2.max_row - 2, ws2.max_row + 1):
        for c in range(1, 9):
            cell = ws2.cell(r, c)
            cell.font = Font(name="맑은 고딕", size=10, bold=True, color=TERRA)
            cell.fill = PatternFill("solid", fgColor=LINEN)
            cell.border = bd
            if c in (3, 5):
                cell.number_format = "0.##"
    if skipped:
        ws2.append([])
        ws2.append(["※ 중량표기 없어 제외된 품목: " + ", ".join(skipped)])
        ws2.cell(ws2.max_row, 1).font = Font(name="맑은 고딕", size=8, color="6B6056")
    for col, w in zip("ABCDEFGH", [14, 36, 13, 11, 13, 10, 9, 30]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = f"A{start}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
