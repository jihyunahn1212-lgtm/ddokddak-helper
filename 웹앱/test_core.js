// 핵심 로직 테스트: node test_core.js  (가상 샘플 기준)
const fs = require('fs');
const path = require('path');
const DD = require('./core.js');

(async () => {
  const fail = [];
  const expect = (cond, msg) => { if (!cond) fail.push(msg); };
  const root = path.join(__dirname, '..', '샘플_입력');

  // 1) 작은 샘플
  const t1 = await DD.parseXlsx(fs.readFileSync(path.join(root, '가상_원본.xlsx')));
  const res = DD.processTable(t1, DD.autoMap(t1.headers));
  console.log('검증:', res.checks.map((c) => c.name + '=' + c.status).join(' | '));
  expect(res.rows.length === 7, '행 수 7');
  expect(res.checks.length === 4 && res.checks.every((c) => c.status === '통과'), '검증 4개 모두 통과');
  expect(res.dupCount === 1 && res.rows[5].dup, '중복 1건(V005)');
  expect(res.daily.length === 5, '일자별 합계 5일');
  const d3 = res.daily.find((d) => d.date === '2026-11-03');
  expect(d3.count === 2 && d3.inc === 830000 && d3.out === 450000, '11-03 합계');
  expect(!('잔액' in res.rows[0]), '누적잔액 제거');

  // 2) 자금요청 (전일잔액 1,234,567 / 목표 5,000,000)
  const base = { date: '2026-11-03', prevBal: 1234567, target: 5000000, unit: 10000 };
  const f = DD.fundRequest(res.rows, base);
  console.log('자금요청:', f.inc, f.out, f.exact, f.rounded, f.final, f.afterFinal, f.diffTarget);
  expect(f.exact === 3385433 && f.rounded === 3390000 && f.final === 3390000, '계산값·올림값·기본 최종값');
  expect(f.afterFinal === 5004567 && f.diffTarget === 4567, '예상 잔액·목표 대비');
  const f2 = DD.fundRequest(res.rows, { ...base, final: 3000000 });
  expect(f2.final === 3000000 && f2.afterFinal === 4614567 && f2.diffTarget === -385433, '수정값(적게) 반영');
  const f3 = DD.fundRequest(res.rows, { ...base, final: 4500000 });
  expect(f3.diffTarget === 1114567, '수정값(많이) 반영');
  // 본사 역송(마이너스): 전일잔액 9,000,000 / 11-03 입금 830,000 출금 450,000 → 잔액 9,380,000, 4,380,000 역송하면 5,000,000
  const fr = DD.fundRequest(res.rows, { ...base, prevBal: 9000000, final: -4380000 });
  expect(fr.exact === -4380000 && fr.rounded === 0, '초과 시 계산값 음수, 올림값 0');
  expect(fr.final === -4380000 && fr.afterFinal === 5000000 && fr.diffTarget === 0, '마이너스(역송) 반영: 예상 잔액 500만원');
  expect(DD.fundRequest(res.rows, { ...base, prevBal: 9000000, final: -1000000 }).afterFinal === 8380000, '일부만 역송');
  expect(DD.fundRequest(res.rows, { ...base, unit: 1000000 }).rounded === 4000000, '백만 원 올림');
  expect(DD.fundRequest(res.rows, { ...base, prevBal: 9000000 }).rounded === 0, '목표 초과 시 0원');
  expect(DD.fundRequest(res.rows, { ...base, date: '2026-12-31' }).count === 0, '해당 날짜 없음');

  // 2-1) 일자 수정(V001: 2026-11-02 → 2026-11-05, 원본 3행)
  const ed = DD.processTable(t1, DD.autoMap(t1.headers), { dateOverrides: { 3: '2026-11-05' } });
  const v1 = ed.rows.find((r) => r.전표번호 === 'V001');
  expect(v1.일자 === '2026-11-05' && v1.edited && v1.memo.includes('원래 2026-11-02'), '일자 수정 반영·메모');
  expect(!ed.daily.find((d) => d.date === '2026-11-02'), '11-02 합계 사라짐');
  const d5 = ed.daily.find((d) => d.date === '2026-11-05');
  expect(d5.count === 2 && d5.inc === 1200000 && d5.out === 12000, '11-05 합계에 합쳐짐');
  expect(DD.fundRequest(ed.rows, { date: '2026-11-05', prevBal: 0, target: 5000000, unit: 10000 }).inc === 1200000, '자금요청도 수정 날짜 기준');
  expect(ed.rows.map((r) => r.id).sort().join() === res.rows.map((r) => r.id).sort().join(), '체크용 id는 날짜를 고쳐도 동일');
  expect(ed.rows[ed.rows.length - 2].전표번호 !== 'V001' || ed.rows[ed.rows.length - 2].일자 >= '2026-11-05', '날짜순 재정렬');
  const bad = DD.processTable(t1, DD.autoMap(t1.headers), { dateOverrides: { 3: '2026-11-02' } });
  expect(!bad.rows.find((r) => r.전표번호 === 'V001').edited, '원래 날짜로 돌리면 수정 표시 없음');
  expect(ed.checks.every((c) => c.status === '통과'), '일자 수정 후에도 검증 통과');

  // 2-2) 중복 의심 행 제외 (두 번째 V005 = id 'V005#2')
  const ex = DD.processTable(t1, DD.autoMap(t1.headers), { excluded: { 'V005#2': true } });
  expect(ex.rows.length === 7 && ex.excludedCount === 1, '제외한 행도 표에는 남음');
  expect(ex.dupCount === 0 && !ex.rows.some((r) => r.dup), '제외하면 중복 의심 해소');
  expect(ex.totals.inc === 2590000 && ex.totals.count === 6, '제외한 행은 합계에서 빠짐');
  expect(ex.daily.find((d) => d.date === '2026-11-06').count === 1, '일자별 합계에서도 빠짐');
  expect(ex.checks.every((c) => c.status === '통과'), '제외해도 원본 읽기 검증은 통과');
  expect(DD.fundRequest(ex.rows, { date: '2026-11-06', prevBal: 0, target: 5000000, unit: 10000 }).inc === 560000, '자금요청에서도 빠짐');
  const ex1 = DD.processTable(t1, DD.autoMap(t1.headers), { excluded: { V005: true } });
  expect(ex1.dupCount === 0 && ex1.totals.inc === 2590000, '첫 번째를 제외해도 나머지는 중복 아님');

  // 2-3) 누락 항목 직접 추가 (예정 출금 2,000,000원)
  const ad = DD.processTable(t1, DD.autoMap(t1.headers), { extraRows: [{ uid: 1, 일자: '2026-11-10', 거래처: '가상X', 적요: '예정', 입금: 0, 출금: 2000000 }] });
  const nrow = ad.rows.find((r) => r.added);
  expect(ad.rows.length === 8 && ad.addedCount === 1 && nrow.memo.includes('직접 추가'), '추가한 행 표시');
  expect(ad.totals.out === 1442000 + 2000000 && ad.totals.count === 8, '추가한 행이 합계에 포함');
  expect(ad.daily[ad.daily.length - 1].date === '2026-11-10', '일자별 합계에 새 날짜');
  expect(ad.checks.every((c) => c.status === '통과'), '추가해도 원본 읽기 검증 통과');
  expect(DD.fundRequest(ad.rows, { date: '2026-11-10', prevBal: 0, target: 5000000, unit: 10000 }).out === 2000000, '자금요청에 추가분 반영');
  const ad2 = DD.processTable(t1, DD.autoMap(t1.headers), { extraRows: [{ uid: 1, 일자: '2026-11-10', 거래처: '', 적요: '', 입금: 0, 출금: 5 }], dateOverrides: { 1000001: '2026-11-12' } });
  expect(ad2.rows.find((r) => r.added).일자 === '2026-11-12', '추가한 행도 일자 수정 가능');

  // 3) 엑셀 내보내기
  const out = DD.buildXlsx(res, { V001: true }, f2);
  fs.writeFileSync(path.join(__dirname, 'test_출력.xlsx'), out);
  const back = await DD.parseXlsx(out);
  expect(back.headers.join(',').startsWith('일자,전표번호,거래처,적요,입금,출금,전표심사완료,확인메모'), '출력 헤더(누적잔액 없음)');
  expect(back.rows[0][6] === 'O' && back.rows[1][6] === 'X', '심사완료 O/X');
  expect(JSON.stringify(back.rows).includes('최종 요청시재(수정 반영)'), '내보내기에 최종 요청시재 포함');

  // 3-1) 제외·추가가 들어간 내보내기: 본표에는 제외한 행이 없고 오른쪽에 제외 목록이 있음
  const both = DD.processTable(t1, DD.autoMap(t1.headers), { excluded: { 'V005#2': true }, extraRows: [{ uid: 1, 일자: '2026-11-10', 거래처: '가상X', 적요: '예정', 입금: 0, 출금: 2000000 }] });
  const outB = DD.buildXlsx(both, {}, null);
  const backB = await DD.parseXlsx(outB);
  const mainRows = backB.rows.filter((r) => r[0] !== null && r[0] !== undefined && r[0] !== '' && /^\d{4}-/.test(String(r[0])) && r[1] !== null);
  expect(mainRows.length === 7, '내보내기 본표 7행(제외 1 빼고 추가 1 더함)');
  expect(JSON.stringify(backB.rows).includes('제외한 행(합계에서 뺌)') && JSON.stringify(backB.rows).includes('직접 추가'), '제외 목록·직접 추가 표시');

  // 4) 큰 금액 주자금 샘플
  const t2 = await DD.parseXlsx(fs.readFileSync(path.join(root, '가상_주자금_원본.xlsx')));
  const r2 = DD.processTable(t2, DD.autoMap(t2.headers));
  expect(r2.rows.length === 22 && r2.dupCount === 1, '큰 샘플 22행·중복 1');
  expect(r2.checks.every((c) => c.status === '통과'), '큰 샘플 검증 통과');
  const mon = DD.fundRequest(r2.rows, { date: '2026-11-09', prevBal: 30000000, target: 5000000, unit: 10000 });
  expect(mon.exact === 428450000, '월요일 요청시재 428,450,000');
  const fri = DD.fundRequest(r2.rows, { date: '2026-11-13', prevBal: 20000000, target: 5000000, unit: 10000 });
  expect(fri.exact === 128000000, '금요일 요청시재 128,000,000');

  // 5) 매일 바뀌는 응원 한마디
  const q0 = DD.dailyQuote(new Date(2026, 10, 18)), q1 = DD.dailyQuote(new Date(2026, 10, 19));
  expect(q0 !== q1, '다음 날은 다른 문구');
  expect(DD.dailyQuote(new Date(2026, 10, 18, 0, 5)) === DD.dailyQuote(new Date(2026, 10, 18, 23, 55)), '같은 날은 하루 종일 같은 문구');
  expect(new Set(DD.QUOTES).size === DD.QUOTES.length && DD.QUOTES.length >= 31, '문구 중복 없음, 한 달치 이상');
  let allDiff = true;
  for (let i = 0; i < DD.QUOTES.length; i++) { if (DD.dailyQuote(new Date(2026, 0, 1 + i)) === DD.dailyQuote(new Date(2026, 0, 2 + i))) allDiff = false; }
  expect(allDiff, '연속된 날은 항상 다른 문구');
  expect(DD.QUOTES.every((s) => s.length <= 45), '문구는 한 줄 길이');

  console.log(fail.length ? '실패: ' + fail.join(' / ') : '모든 테스트 통과');
  process.exit(fail.length ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
