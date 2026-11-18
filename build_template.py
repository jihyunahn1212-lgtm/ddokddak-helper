"""엑셀 2016/2019용 자동 정리 템플릿 생성 (가상 샘플 데이터 포함).
- 원본붙여넣기: 시스템 로우데이터를 A1부터 붙여넣는 시트(열 순서가 달라도 '열 이름'으로 찾음)
- 정리결과: 날짜순 정렬·서식 통일·누적잔액·검증이 자동 계산되는 시트
- 계산: 중간 계산용 시트(건드리지 않기)
수식은 신형 함수(SORT/FILTER/XLOOKUP)를 쓰지 않는다."""
import os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

base = os.path.dirname(os.path.abspath(__file__))
MAXR = 501                      # 처리 가능한 마지막 행 (데이터 500행)
RAW = "원본붙여넣기"
CAL = "계산"
RES = "정리결과"
HDR = f"{RAW}!$A$1:$Z$1"
BODY = f"{RAW}!$A$2:$Z${MAXR}"

wb = Workbook()
raw = wb.active; raw.title = RAW
res = wb.create_sheet(RES)
cal = wb.create_sheet(CAL)
cal.sheet_properties.tabColor = "999999"
bold = Font(bold=True)
fill = PatternFill("solid", fgColor="DDEBF7")

# ---------- 원본붙여넣기: 가상 로우데이터 ----------
raw.append(["전표번호", "일자", "거래처", "입금", "출금", "적요"])
for r in [
    ("V003", "2026-11-03", "가상상사C", "830,000", None, "매출입금"),
    ("V001", date(2026, 11, 2), "가상상사A", 1200000, None, "매출입금"),
    ("V002", date(2026, 11, 3), "가상상사B", None, "450,000", "외주비"),
    ("V005", date(2026, 11, 6), "가상상사E", 560000, None, "매출입금"),
    ("V004", date(2026, 11, 5), "가상상사D", None, 12000, "수수료"),
    ("V005", date(2026, 11, 6), "가상상사E", 560000, None, "매출입금"),
    ("V006", date(2026, 11, 9), "가상상사F", None, 980000, "임차료"),
]:
    raw.append(list(r))
for c in raw[1]:
    c.font = bold; c.fill = fill
raw["H1"] = "← A1부터 시스템 로우데이터를 붙여넣으세요(헤더 포함, 최대 500행). 열 순서는 달라도 됩니다."
raw["H1"].font = Font(color="C00000", bold=True)


def rawcell(r, name):
    """계산 시트의 r행에 해당하는 원본 값 (열 이름으로 찾음)"""
    return f'INDEX({BODY},{r-1},MATCH("{name}",{HDR},0))'


# ---------- 계산 시트 ----------
for col, t in zip("ABCDEFG", ["(참고)원본행", "일자", "전표번호", "입금", "출금", "정렬순번", "중복판정"]):
    cal[f"{col}1"] = t; cal[f"{col}1"].font = bold
for r in range(2, MAXR + 1):
    xd, xi, xo = rawcell(r, "일자"), rawcell(r, "입금"), rawcell(r, "출금")
    cal[f"A{r}"] = r
    cal[f"B{r}"] = f'=IF({xd}="","",IF(ISNUMBER({xd}),{xd},IFERROR(DATEVALUE({xd}),"변환오류")))'
    cal[f"C{r}"] = f'=IF(B{r}="","",{rawcell(r, "전표번호")}&"")'
    cal[f"D{r}"] = f'=IF(B{r}="","",IFERROR(IF(ISNUMBER({xi}),{xi},VALUE(SUBSTITUTE({xi},",",""))),0))'
    cal[f"E{r}"] = f'=IF(B{r}="","",IFERROR(IF(ISNUMBER({xo}),{xo},VALUE(SUBSTITUTE({xo},",",""))),0))'
    cal[f"F{r}"] = f'=IF(ISNUMBER(B{r}),COUNTIF(B$2:B${MAXR},"<"&B{r})+COUNTIF(B$2:B{r},B{r}),"")'
    cal[f"G{r}"] = (f'=IF(C{r}="","",IF(COUNTIFS(C$2:C{r},C{r},B$2:B{r},B{r},D$2:D{r},D{r},E$2:E{r},E{r})>1,'
                    f'"중복 의심",""))')
    cal[f"B{r}"].number_format = "yyyy-mm-dd"

# ---------- 정리결과 ----------
res.append(["일자", "전표번호", "거래처", "적요", "입금", "출금", "전표심사완료", "누적잔액", "확인메모"])
for c in res[1]:
    c.font = bold; c.fill = fill
pos = f'MATCH(ROW()-1,{CAL}!$F$2:$F${MAXR},0)'
for r in range(2, MAXR + 1):
    res[f"A{r}"] = f'=IFERROR(INDEX({CAL}!$B$2:$B${MAXR},{pos}),"")'
    res[f"B{r}"] = f'=IFERROR(INDEX({CAL}!$C$2:$C${MAXR},{pos}),"")'
    res[f"C{r}"] = f'=IFERROR(INDEX({BODY},{pos},MATCH("거래처",{HDR},0))&"","")'
    res[f"D{r}"] = f'=IFERROR(INDEX({BODY},{pos},MATCH("적요",{HDR},0))&"","")'
    res[f"E{r}"] = f'=IFERROR(INDEX({CAL}!$D$2:$D${MAXR},{pos}),"")'
    res[f"F{r}"] = f'=IFERROR(INDEX({CAL}!$E$2:$E${MAXR},{pos}),"")'
    prev = "$M$3" if r == 2 else f"H{r-1}"
    res[f"H{r}"] = f'=IF(A{r}="","",{prev}+E{r}-F{r})'
    res[f"I{r}"] = f'=IFERROR(INDEX({CAL}!$G$2:$G${MAXR},{pos}),"")'
    res[f"A{r}"].number_format = "yyyy-mm-dd"
    for col in "EFH":
        res[f"{col}{r}"].number_format = "#,##0"
res.freeze_panes = "A2"
dv = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
res.add_data_validation(dv); dv.add(f"G2:G{MAXR}")
res.conditional_formatting.add(
    f"A2:I{MAXR}", FormulaRule(formula=['$I2<>""'], fill=PatternFill("solid", bgColor="FFF2CC")))

# 오른쪽: 입력값 + 검증
res["K1"] = "[가상 샘플] 입력값(가정) — 매주 본인 값으로 교체"; res["K1"].font = bold
res["K2"] = "정렬 기준"; res["M2"] = "일자 오름차순, 같은 날짜는 원본 순서"
res["K3"] = "이전 주 마감 잔액"; res["M3"] = 1000000
res["K4"] = "재무장부 잔액"; res["M4"] = 2148000
res["K5"] = "허용 오차"; res["M5"] = 0
res["K6"] = "최종 잔액(정리결과)"; res["K6"].font = bold
res["M6"] = f'=IF(COUNT(A2:A{MAXR})=0,M3,INDEX(H2:H{MAXR},COUNT(A2:A{MAXR})))'
res["M6"].font = bold
for c in ("M3", "M4", "M6"):
    res[c].number_format = "#,##0"
for c in ("M3", "M4", "M5"):
    res[c].fill = PatternFill("solid", fgColor="FFF2CC")

for col, t in zip("KLMNO", ["검증 항목", "기준값", "결과값", "차이", "판정"]):
    res[f"{col}7"] = t; res[f"{col}7"].font = bold; res[f"{col}7"].fill = fill


def rawcol(name):
    return f'INDEX({BODY},0,MATCH("{name}",{HDR},0))'


def amt_sum(name):
    return f'=SUMPRODUCT(--("0"&SUBSTITUTE({rawcol(name)},",","")))'


checks = [
    # (이름, 기준값, 결과값, 엄격판정 여부)
    ("1. 잔액 재계산 일치", f"=M3+SUM(E2:E{MAXR})-SUM(F2:F{MAXR})", "=M6", False),
    ("2. 재무장부 잔액 대조", "=M4", "=M6", False),
    ("3. 행 수 일치(원본 vs 정리)", f'=COUNTA({rawcol("전표번호")})', f"=COUNT(A2:A{MAXR})", True),
    ("4. 입금 합계 일치(원본 vs 정리)", amt_sum("입금"), f"=SUM(E2:E{MAXR})", False),
    ("5. 출금 합계 일치(원본 vs 정리)", amt_sum("출금"), f"=SUM(F2:F{MAXR})", False),
    ("6. 일자 변환 오류 행 수", 0, f'=COUNTIF({CAL}!B2:B{MAXR},"변환오류")', True),
    ("7. 처리 한도(500행) 초과 데이터", 0, f"=COUNTA({RAW}!A{MAXR+1}:Z2000)", True),
]
for i, (name, ref, resf, strict) in enumerate(checks, start=8):
    res[f"K{i}"] = name; res[f"L{i}"] = ref; res[f"M{i}"] = resf
    res[f"N{i}"] = f"=M{i}-L{i}"
    cond = f"N{i}=0" if strict else f"ABS(N{i})<=$M$5"
    res[f"O{i}"] = f'=IF(ISERROR(N{i}),"오류",IF({cond},"통과","불일치"))'
    for col in "LMN":
        res[f"{col}{i}"].number_format = "#,##0"
res["K16"] = "참고: 중복 의심 행 수(노란색 행)"; res["M16"] = f'=COUNTIF(I2:I{MAXR},"중복 의심")'
res.conditional_formatting.add("O8:O14", CellIsRule(operator="equal", formula=['"불일치"'], font=Font(color="C00000", bold=True)))
res.conditional_formatting.add("O8:O14", CellIsRule(operator="equal", formula=['"오류"'], font=Font(color="C00000", bold=True)))
res.conditional_formatting.add("O8:O14", CellIsRule(operator="equal", formula=['"통과"'], font=Font(color="00803C", bold=True)))

for col, w in zip("ABCDEFGHIJKLMNO", (12, 10, 12, 12, 12, 12, 13, 14, 12, 3, 34, 16, 18, 16, 10)):
    res.column_dimensions[col].width = w
for col in "ABCDEFG":
    cal.column_dimensions[col].width = 12

wb.move_sheet(RES, offset=-1)  # 순서: 정리결과, 원본붙여넣기, 계산
out = os.path.join(base, "정리_템플릿_가상샘플.xlsx")
wb.save(out)
print("저장:", out)
