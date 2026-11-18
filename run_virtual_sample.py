"""최종_프롬프트.md 를 '가상 샘플'로 실행하는 스크립트. 모든 값은 가상이며 가정은 파일 안에 표시함."""
import os
from datetime import date, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

base = os.path.dirname(os.path.abspath(__file__))
in_dir = os.path.join(base, "샘플_입력")
os.makedirs(in_dir, exist_ok=True)

# ---- 1) 가상 원본 (열 순서가 표준과 다르고, 일부 금액이 텍스트, 중복 전표 1건 포함) ----
src = Workbook(); ws = src.active
ws.append(["전표번호", "일자", "거래처", "입금", "출금", "적요"])
raw = [
    ("V003", "2026-11-03", "가상상사C", "830,000", None, "매출입금"),
    ("V001", date(2026, 11, 2), "가상상사A", 1200000, None, "매출입금"),
    ("V002", date(2026, 11, 3), "가상상사B", None, "450,000", "외주비"),
    ("V005", date(2026, 11, 6), "가상상사E", 560000, None, "매출입금"),
    ("V004", date(2026, 11, 5), "가상상사D", None, 12000, "수수료"),
    ("V005", date(2026, 11, 6), "가상상사E", 560000, None, "매출입금"),  # 중복(의도적)
    ("V006", date(2026, 11, 9), "가상상사F", None, 980000, "임차료"),
]
for r in raw:
    ws.append(list(r))
src_path = os.path.join(in_dir, "가상_원본.xlsx")
src.save(src_path)

# ---- 2) 가상 입력값 (가정) ----
OPENING = 1_000_000      # 가정: 이전 주 마감 잔액
LEDGER = 2_148_000       # 가정: 재무장부 잔액 (중복 제외 시 맞는 값)
TOL = 0                  # 가정: 허용 오차
STD = ["일자", "전표번호", "거래처", "적요", "입금", "출금"]

# ---- 3) 정리 ----
def num(v):
    if v in (None, ""):
        return 0
    return int(str(v).replace(",", ""))

def dt(v):
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else datetime.strptime(v, "%Y-%m-%d").date()

wb_in = load_workbook(src_path)
sh = wb_in.active
hdr = [c.value for c in sh[1]]
rows = []
for r in sh.iter_rows(min_row=2, values_only=True):
    d = dict(zip(hdr, r))
    rows.append({"일자": dt(d["일자"]), "전표번호": d["전표번호"], "거래처": d["거래처"],
                 "적요": d["적요"], "입금": num(d["입금"]), "출금": num(d["출금"])})
orig_n = len(rows)
orig_in = sum(r["입금"] for r in rows)
orig_out = sum(r["출금"] for r in rows)

seen, dups = set(), []
for r in rows:
    key = (r["전표번호"], r["일자"], r["입금"], r["출금"])
    if key in seen:
        dups.append(r)
    seen.add(key)
rows.sort(key=lambda r: (r["일자"], r["전표번호"]))

# ---- 4) 출력 (시트 1개: 왼쪽 정리본 / 오른쪽 입력값·검증) ----
out = Workbook()
a = out.active; a.title = "정리결과"
bold = Font(bold=True); fill = PatternFill("solid", fgColor="DDEBF7")
warn = PatternFill("solid", fgColor="FFF2CC")

# 왼쪽: 정리본 A~I (I = 확인메모)
a.append(STD + ["전표심사완료", "누적잔액", "확인메모"])
for c in a[1]:
    c.font = bold; c.fill = fill
n = len(rows)
seen2 = set()
for i, r in enumerate(rows, start=2):
    key = (r["전표번호"], r["일자"], r["입금"], r["출금"])
    memo = "중복 의심" if key in seen2 else None
    seen2.add(key)
    a.append([r[k] for k in STD] + [None, None, memo])
    a.cell(i, 1).number_format = "yyyy-mm-dd"
    for col in (5, 6, 8):
        a.cell(i, col).number_format = "#,##0"
    prev = "$M$3" if i == 2 else f"H{i-1}"
    a.cell(i, 8).value = f"={prev}+E{i}-F{i}"
    if memo:
        for col in range(1, 10):
            a.cell(i, col).fill = warn
last = n + 1
a.cell(last + 1, 7).value = "최종 잔액"; a.cell(last + 1, 7).font = bold
a.cell(last + 1, 8).value = f"=H{last}"; a.cell(last + 1, 8).number_format = "#,##0"; a.cell(last + 1, 8).font = bold
dv = DataValidation(type="list", formula1='"O,X"', allow_blank=True)
a.add_data_validation(dv); dv.add(f"G2:G{last}")
a.freeze_panes = "A2"

# 오른쪽: K~O (입력값 + 검증), K=라벨 L=기준값 M=값 N=차이 O=판정
a["K1"] = "[가상 샘플] 입력값(가정)"; a["K1"].font = bold
a["K2"] = "정렬 기준(가정)"; a["M2"] = "일자 오름차순, 같으면 전표번호순"
a["K3"] = "이전 주 마감 잔액"; a["M3"] = OPENING
a["K4"] = "재무장부 잔액"; a["M4"] = LEDGER
a["K5"] = "허용 오차"; a["M5"] = TOL
for c in ("M3", "M4"):
    a[c].number_format = "#,##0"

for col, t in zip("KLMNO", ["검증 항목", "기준값", "결과값", "차이", "판정"]):
    a[f"{col}7"] = t; a[f"{col}7"].font = bold; a[f"{col}7"].fill = fill
rng_in, rng_out = f"E2:E{last}", f"F2:F{last}"
checks = [
    ("1. 잔액 재계산 일치", f"=M3+SUM({rng_in})-SUM({rng_out})", f"=H{last}"),
    ("2. 재무장부 잔액 대조", "=M4", f"=H{last}"),
    ("3. 행 수 일치(원본 vs 정리본)", orig_n, f"=COUNTA(B2:B{last})"),
    ("4. 입금 합계 일치(원본 vs 정리본)", orig_in, f"=SUM({rng_in})"),
    ("5. 출금 합계 일치(원본 vs 정리본)", orig_out, f"=SUM({rng_out})"),
]
for i, (name, ref, res) in enumerate(checks, start=8):
    a[f"K{i}"] = name; a[f"L{i}"] = ref; a[f"M{i}"] = res
    a[f"N{i}"] = f"=M{i}-L{i}"
    a[f"O{i}"] = f'=IF(ABS(N{i})<=$M$5,"통과","불일치")'
    for col in "LMN":
        a[f"{col}{i}"].number_format = "#,##0"

a["K14"] = "확인 필요 행(삭제하지 않고 표시만 함, 왼쪽 표에 노란색)"; a["K14"].font = bold
for col, t in zip("KLMNOP", ["유형", "일자", "전표번호", "거래처", "입금", "출금"]):
    a[f"{col}15"] = t; a[f"{col}15"].font = bold; a[f"{col}15"].fill = fill
for j, d in enumerate(dups, start=16):
    for col, val in zip("KLMNOP", ["중복 의심", d["일자"], d["전표번호"], d["거래처"], d["입금"], d["출금"]]):
        a[f"{col}{j}"] = val
    a[f"L{j}"].number_format = "yyyy-mm-dd"

for col, w in zip("ABCDEFGHIJKLMNOP", (12, 10, 12, 12, 12, 12, 13, 14, 12, 3, 34, 16, 18, 16, 10, 12)):
    a.column_dimensions[col].width = w

out_path = os.path.join(base, "정리결과_가상샘플.xlsx")
out.save(out_path)

# ---- 5) 파이썬 독립 재계산 (엑셀 수식 결과와 비교용) ----
bal = OPENING
for r in rows:
    bal += r["입금"] - r["출금"]
print("저장:", out_path)
print("행 수:", orig_n, "| 입금합:", orig_in, "| 출금합:", orig_out)
print("파이썬 재계산 최종잔액:", bal, "| 장부잔액:", LEDGER, "| 차이:", bal - LEDGER)
print("중복 의심 행:", [(d["전표번호"], d["입금"], d["출금"]) for d in dups])
