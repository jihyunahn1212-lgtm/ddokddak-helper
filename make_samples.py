"""가상 샘플 다운로드 파일 3개 생성 (실제 데이터 아님).
부서마다 열 순서/서식이 다른 상황을 흉내 낸다."""
import os
from datetime import date
from openpyxl import Workbook

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "샘플_다운로드")
os.makedirs(out, exist_ok=True)

# 1) 영업1팀: 표준 열 순서, 날짜/숫자 정상
wb = Workbook(); ws = wb.active
ws.append(["일자", "부서코드", "거래처", "계정과목", "금액", "비고"])
ws.append([date(2026, 11, 2), "S100", "가상상사A", "매출", 1200000, ""])
ws.append([date(2026, 11, 3), "S100", "가상상사B", "매출", 830000, "분할"])
ws.append([date(2026, 11, 5), "S100", "가상상사C", "매입", -450000, ""])
wb.save(os.path.join(out, "영업1팀_20261110.xlsx"))

# 2) 영업2팀: 열 순서가 다름
wb = Workbook(); ws = wb.active
ws.append(["금액", "거래처", "일자", "계정과목", "비고", "부서코드"])
ws.append([560000, "가상상사D", date(2026, 11, 4), "매출", "", "S200"])
ws.append([-120000, "가상상사E", date(2026, 11, 6), "매입", "반품", "S200"])
wb.save(os.path.join(out, "영업2팀_20261110.xlsx"))

# 3) 관리팀: 날짜/금액이 텍스트, 열 순서 다름, 비고 열 없음
wb = Workbook(); ws = wb.active
ws.append(["부서코드", "일자", "계정과목", "거래처", "금액"])
ws.append(["M300", "2026-11-02", "비용", "가상상사F", "98,000"])
ws.append(["M300", "2026-11-07", "비용", "가상상사G", "1,250,000"])
wb.save(os.path.join(out, "관리팀_20261110.xlsx"))

print("생성 완료:", out)
