/* 똑딱 정리 도우미 - 핵심 로직 (외부 라이브러리 없음, 브라우저/Node 공용)
 * 읽기: .xlsx(zip) 직접 해석 / 쓰기: 최소 .xlsx 직접 생성 / 정리·검증 계산 */
(function (root) {
  'use strict';
  const dec = new TextDecoder('utf-8');
  const enc = new TextEncoder();

  /* ---------------- zip 읽기 ---------------- */
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(u8) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  async function inflateRaw(u8) {
    const stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  async function unzip(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let e = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
    }
    if (e < 0) throw new Error('엑셀(.xlsx) 파일이 아닌 것 같아요');
    const count = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const files = {};
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true);
      const xlen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const loff = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      const ln = dv.getUint16(loff + 26, true);
      const lx = dv.getUint16(loff + 28, true);
      const start = loff + 30 + ln + lx;
      files[name] = { method, data: u8.subarray(start, start + csize) };
      p += 46 + nlen + xlen + clen;
    }
    return files;
  }
  async function readText(files, name) {
    const f = files[name];
    if (!f) return null;
    if (f.method === 0) return dec.decode(f.data);
    if (f.method === 8) return dec.decode(await inflateRaw(f.data));
    throw new Error('지원하지 않는 압축 방식이에요');
  }

  /* ---------------- xlsx 읽기 ---------------- */
  function unesc(s) {
    return s
      .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }
  function attrs(str) {
    const o = {};
    const re = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
    let m;
    while ((m = re.exec(str))) o[m[1]] = unesc(m[2]);
    return o;
  }
  function textOf(xml) {
    xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = '';
    const re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let m;
    while ((m = re.exec(xml))) out += unesc(m[1]);
    return out;
  }
  function colIndex(ref) {
    const m = /^([A-Z]+)/.exec(ref);
    let n = 0;
    for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  const BUILTIN_DATE = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
  function isDateFmt(code) {
    const c = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[ymdhs]/i.test(c);
  }
  function serialToISO(n) {
    const d = new Date(Math.round((n - 25569) * 86400000));
    return d.toISOString().slice(0, 10);
  }

  async function parseXlsx(buf) {
    const files = await unzip(buf);
    const wbx = await readText(files, 'xl/workbook.xml');
    if (!wbx) throw new Error('엑셀(.xlsx) 파일이 아닌 것 같아요');
    const sm = /<sheet\b([^>]*?)\/?>/.exec(wbx);
    if (!sm) throw new Error('시트를 찾을 수 없어요');
    const sa = attrs(sm[1]);
    let target = 'worksheets/sheet1.xml';
    const rels = await readText(files, 'xl/_rels/workbook.xml.rels');
    if (rels && sa['r:id']) {
      const re = /<Relationship\b([^>]*?)\/?>/g;
      let m;
      while ((m = re.exec(rels))) {
        const a = attrs(m[1]);
        if (a.Id === sa['r:id']) { target = a.Target; break; }
      }
    }
    const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
    const sheetXml = await readText(files, path);
    if (!sheetXml) throw new Error('시트 내용을 읽을 수 없어요');

    const shared = [];
    const ssx = await readText(files, 'xl/sharedStrings.xml');
    if (ssx) {
      const re = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
      let m;
      while ((m = re.exec(ssx))) shared.push(m[1] ? textOf(m[1]) : '');
    }

    const dateXf = [];
    const stx = await readText(files, 'xl/styles.xml');
    if (stx) {
      const custom = {};
      const re = /<numFmt\b([^>]*?)\/?>/g;
      let m;
      while ((m = re.exec(stx))) {
        const a = attrs(m[1]);
        custom[a.numFmtId] = a.formatCode || '';
      }
      const cx = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stx);
      if (cx) {
        const re2 = /<xf\b([^>]*?)\/?>/g;
        let m2;
        while ((m2 = re2.exec(cx[1]))) {
          const id = parseInt(attrs(m2[1]).numFmtId || '0', 10);
          dateXf.push(BUILTIN_DATE.has(id) || (custom[id] !== undefined && isDateFmt(custom[id])));
        }
      }
    }

    const grid = [];
    const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    let rm;
    let autoRow = 0;
    while ((rm = rowRe.exec(sheetXml))) {
      const ra = attrs(rm[1]);
      const rIdx = ra.r ? parseInt(ra.r, 10) - 1 : autoRow;
      autoRow = rIdx + 1;
      const cells = [];
      const body = rm[2] || '';
      const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm;
      let autoCol = 0;
      while ((cm = cRe.exec(body))) {
        const ca = attrs(cm[1]);
        const ci = ca.r ? colIndex(ca.r) : autoCol;
        autoCol = ci + 1;
        const inner = cm[2] || '';
        const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
        let val = null;
        const t = ca.t || 'n';
        if (t === 's' && vm) val = shared[parseInt(vm[1], 10)];
        else if (t === 'inlineStr') val = textOf(inner);
        else if (t === 'str' && vm) val = unesc(vm[1]);
        else if (t === 'b' && vm) val = vm[1] === '1' ? 'TRUE' : 'FALSE';
        else if (t === 'e') val = null;
        else if (vm && vm[1] !== '') {
          const num = parseFloat(vm[1]);
          const s = ca.s !== undefined ? parseInt(ca.s, 10) : 0;
          val = dateXf[s] ? serialToISO(num) : num;
        }
        cells[ci] = val;
      }
      grid[rIdx] = cells;
    }
    const rows = [];
    for (const r of grid) {
      if (!r) continue;
      const dense = Array.from(r, (v) => (v === undefined ? null : v));
      if (dense.some((v) => v !== null && String(v).trim() !== '')) rows.push(dense);
    }
    if (!rows.length) throw new Error('데이터가 비어 있어요');
    const headers = rows[0].map((h) => (h === null ? '' : String(h).trim()));
    return { sheetName: sa.name || '', headers, rows: rows.slice(1) };
  }

  /* ---------------- 값 변환 ---------------- */
  function normAmount(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/[,\s원]/g, '');
    if (s === '') return 0;
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
  }
  function validISO(y, m, d) {
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d
      ? `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}` : null;
  }
  function normDate(v) {
    if (v === null || v === undefined || String(v).trim() === '') return null;
    if (typeof v === 'number') {
      if (v >= 19000101 && v <= 29991231) return validISO(Math.floor(v / 10000), Math.floor(v / 100) % 100, v % 100);
      if (v > 20000 && v < 80000) return serialToISO(v);
      return null;
    }
    const s = String(v).trim();
    let m = /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/.exec(s);
    if (m) return validISO(+m[1], +m[2], +m[3]);
    m = /^(\d{4})(\d{2})(\d{2})$/.exec(s);
    if (m) return validISO(+m[1], +m[2], +m[3]);
    return null;
  }

  /* ---------------- 정리·검증 ---------------- */
  const REQUIRED = ['일자', '전표번호', '입금', '출금'];
  const OPTIONAL = ['거래처', '적요'];

  function autoMap(headers) {
    const norm = (s) => String(s).replace(/\s+/g, '');
    const map = {};
    for (const f of REQUIRED.concat(OPTIONAL)) {
      const i = headers.findIndex((h) => norm(h) === f);
      map[f] = i >= 0 ? i : -1;
    }
    return map;
  }

  /* opts: { dateOverrides: {원본행번호: 'YYYY-MM-DD'}, excluded: {id: true}, extraRows: [{uid, 일자, 거래처, 적요, 입금, 출금}] }
   * - 제외(excluded)한 행은 표에는 남기되 합계·일자별 합계·자금요청·중복 판정에서 뺀다
   * - extraRows(직접 추가한 행)는 원본 행과 똑같이 합계에 들어간다 */
  function processTable(table, map, opts) {
    opts = opts || {};
    const over = opts.dateOverrides || {};
    const exc = opts.excluded || {};
    const miss = REQUIRED.filter((f) => map[f] === undefined || map[f] < 0);
    if (miss.length) throw new Error('필수 열을 골라 주세요: ' + miss.join(', '));
    const get = (row, f) => (map[f] >= 0 ? row[map[f]] : null);
    const items = [];
    table.rows.forEach((row, idx) => {
      const errors = [];
      const orig = normDate(get(row, '일자'));
      const ov = over[idx + 2];
      const date = ov || orig;
      const edited = !!ov && ov !== orig;
      if (!date) errors.push('일자 변환 오류');
      let inc = normAmount(get(row, '입금'));
      let out = normAmount(get(row, '출금'));
      if (Number.isNaN(inc) || Number.isNaN(out)) errors.push('금액 변환 오류');
      if (Number.isNaN(inc)) inc = 0;
      if (Number.isNaN(out)) out = 0;
      const txt = (f) => { const v = get(row, f); return v === null || v === undefined ? '' : String(v); };
      items.push({ srcRow: idx + 2, 일자: date, 전표번호: txt('전표번호'), 거래처: txt('거래처'), 적요: txt('적요'), 입금: inc, 출금: out, errors, dup: false, edited, origDate: orig, added: false });
    });
    const occ = {};
    items.forEach((r) => { // 체크·제외 저장용 id는 원본 순서 기준(날짜를 고쳐도 바뀌지 않음)
      occ[r.전표번호] = (occ[r.전표번호] || 0) + 1;
      r.id = r.전표번호 + (occ[r.전표번호] > 1 ? '#' + occ[r.전표번호] : '');
    });
    const sourceItems = items.slice();
    (opts.extraRows || []).forEach((x) => {
      const srcRow = 1000000 + x.uid;
      const orig = normDate(x.일자);
      const ov = over[srcRow];
      const date = ov || orig;
      items.push({
        srcRow, id: 'ADD' + x.uid, 일자: date, 전표번호: '', 거래처: x.거래처 || '', 적요: x.적요 || '',
        입금: x.입금 || 0, 출금: x.출금 || 0, errors: date ? [] : ['일자 변환 오류'], dup: false,
        edited: !!ov && ov !== orig, origDate: orig, added: true,
      });
    });
    items.forEach((r) => { r.excluded = !r.added && !!exc[r.id]; });

    const good = items.filter((r) => r.일자).sort((a, b) => (a.일자 < b.일자 ? -1 : a.일자 > b.일자 ? 1 : a.srcRow - b.srcRow));
    const bad = items.filter((r) => !r.일자);
    const sorted = good.concat(bad);
    const seen = new Set();
    sorted.forEach((r) => {
      if (!r.excluded) {
        const key = [r.전표번호, r.일자, r.입금, r.출금].join('|');
        if (seen.has(key)) { r.dup = true; r.errors.push('중복 의심'); }
        seen.add(key);
      }
      const notes = r.errors.slice();
      if (r.edited) notes.push('일자 수정됨(원래 ' + (r.origDate || '오류') + ')');
      if (r.added) notes.push('직접 추가');
      if (r.excluded) notes.push('제외됨');
      r.memo = notes.join(', ');
    });
    const active = sorted.filter((r) => !r.excluded);

    const sumRaw = (f) => table.rows.reduce((s, r) => { const v = normAmount(get(r, f)); return s + (Number.isNaN(v) ? 0 : v); }, 0);
    const sumSrc = (f) => sourceItems.reduce((s, r) => s + r[f], 0);
    const mk = (name, expected, actual) => {
      const diff = actual - expected;
      return { name, expected, actual, diff, status: Math.abs(diff) < 0.005 ? '통과' : '불일치' };
    };
    // 행 수·합계 검증은 "원본 파일을 빠짐없이 읽었는지" 확인한다(제외한 행도 포함해서 비교)
    const checks = [
      mk('행 수 일치(원본 vs 정리)', table.rows.length, sourceItems.length),
      mk('입금 합계 일치', sumRaw('입금'), sumSrc('입금')),
      mk('출금 합계 일치', sumRaw('출금'), sumSrc('출금')),
      mk('변환 오류 행 수', 0, active.filter((r) => r.errors.some((e) => e.includes('변환'))).length),
    ];

    const dm = new Map();
    active.forEach((r) => {
      if (!r.일자) return;
      const d = dm.get(r.일자) || { date: r.일자, count: 0, inc: 0, out: 0, flagged: 0 };
      d.count += 1; d.inc += r.입금; d.out += r.출금; if (r.errors.length) d.flagged += 1;
      dm.set(r.일자, d);
    });
    return {
      rows: sorted, checks,
      dupCount: active.filter((r) => r.dup).length,
      daily: [...dm.values()],
      totals: { inc: active.reduce((s, r) => s + r.입금, 0), out: active.reduce((s, r) => s + r.출금, 0), count: active.length },
      excludedCount: sorted.length - active.length,
      addedCount: sorted.filter((r) => r.added).length,
    };
  }

  /* ---------------- xlsx 쓰기 ---------------- */
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function colName(i) {
    let s = '';
    for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
    return s;
  }
  function isoToSerial(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    return Date.UTC(y, m - 1, d) / 86400000 + 25569;
  }
  // 셀: {t:'s'|'n'|'d', v, s:스타일번호}  스타일: 0기본 1날짜 2숫자 3헤더 4노랑글자 5노랑날짜 6노랑숫자 7굵게 8굵은숫자
  function sheetXml(grid, widths) {
    let rowsXml = '';
    grid.forEach((row, r) => {
      if (!row) return;
      let cs = '';
      row.forEach((c, ci) => {
        if (c === undefined || c === null) return;
        const ref = colName(ci) + (r + 1);
        const s = c.s ? ` s="${c.s}"` : '';
        if (c.t === 'n') cs += `<c r="${ref}"${s}><v>${c.v}</v></c>`;
        else if (c.t === 'd') cs += `<c r="${ref}"${s}><v>${isoToSerial(c.v)}</v></c>`;
        else cs += `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(c.v)}</t></is></c>`;
      });
      rowsXml += `<row r="${r + 1}">${cs}</row>`;
    });
    const cols = widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${cols}</cols><sheetData>${rowsXml}</sheetData></worksheet>`;
  }
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Malgun Gothic"/></font><font><b/><sz val="11"/><name val="Malgun Gothic"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/><xf numFmtId="3" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="3" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  function zipStore(entries) {
    const parts = [];
    const central = [];
    let offset = 0;
    const u16 = (n) => new Uint8Array([n & 255, (n >> 8) & 255]);
    const u32 = (n) => new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
    const cat = (...a) => { const len = a.reduce((s, x) => s + x.length, 0); const o = new Uint8Array(len); let p = 0; for (const x of a) { o.set(x, p); p += x.length; } return o; };
    for (const [name, text] of entries) {
      const nameB = enc.encode(name);
      const data = enc.encode(text);
      const crc = crc32(data);
      const local = cat(u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(0x21), u32(crc), u32(data.length), u32(data.length), u16(nameB.length), u16(0), nameB, data);
      central.push(cat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0x21), u32(crc), u32(data.length), u32(data.length), u16(nameB.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), nameB));
      parts.push(local);
      offset += local.length;
    }
    const cd = cat(...central);
    const end = cat(u32(0x06054b50), u16(0), u16(0), u16(entries.length), u16(entries.length), u32(cd.length), u32(offset), u16(0));
    return cat(...parts, cd, end);
  }

  function buildXlsx(result, checked, fund) {
    const grid = [];
    const set = (r, c, cell) => { (grid[r] = grid[r] || [])[c] = cell; };
    const S = (v, s) => ({ t: 's', v, s });
    ['일자', '전표번호', '거래처', '적요', '입금', '출금', '전표심사완료', '확인메모'].forEach((h, i) => set(0, i, S(h, 3)));
    const shown = result.rows.filter((x) => !x.excluded);
    shown.forEach((r, i) => {
      const y = r.dup ? 4 : 0;
      const row = i + 1;
      if (r.일자) set(row, 0, { t: 'd', v: r.일자, s: r.dup ? 5 : 1 }); else set(row, 0, S('(변환 오류)', y));
      set(row, 1, S(r.전표번호, y));
      set(row, 2, S(r.거래처, y));
      set(row, 3, S(r.적요, y));
      set(row, 4, { t: 'n', v: r.입금, s: r.dup ? 6 : 2 });
      set(row, 5, { t: 'n', v: r.출금, s: r.dup ? 6 : 2 });
      set(row, 6, S(checked[r.id] ? 'O' : 'X', y));
      set(row, 7, S(r.memo, y));
    });

    // 오른쪽(K열~): 검증 → 일자별 합계 → 자금요청 계산
    let r = 0;
    ['검증 항목', '기준값', '결과값', '차이', '판정'].forEach((h, i) => set(r, 10 + i, S(h, 3)));
    result.checks.forEach((c) => {
      r += 1;
      set(r, 10, S(c.name, 0));
      set(r, 11, { t: 'n', v: c.expected, s: 2 });
      set(r, 12, { t: 'n', v: c.actual, s: 2 });
      set(r, 13, { t: 'n', v: c.diff, s: 2 });
      set(r, 14, S(c.status, 7));
    });
    r += 1; set(r, 10, S('중복 의심 행 수(노란색)', 0)); set(r, 12, { t: 'n', v: result.dupCount, s: 2 });

    r += 2;
    ['일자별 합계', '건수', '입금 합계', '출금 합계'].forEach((h, i) => set(r, 10 + i, S(h, 3)));
    result.daily.forEach((d) => {
      r += 1;
      set(r, 10, { t: 'd', v: d.date, s: 1 });
      set(r, 11, { t: 'n', v: d.count, s: 2 });
      set(r, 12, { t: 'n', v: d.inc, s: 2 });
      set(r, 13, { t: 'n', v: d.out, s: 2 });
    });
    if (result.daily.length) {
      r += 1;
      set(r, 10, S('합계', 7));
      set(r, 11, { t: 'n', v: result.daily.reduce((s, d) => s + d.count, 0), s: 8 });
      set(r, 12, { t: 'n', v: result.daily.reduce((s, d) => s + d.inc, 0), s: 8 });
      set(r, 13, { t: 'n', v: result.daily.reduce((s, d) => s + d.out, 0), s: 8 });
    }

    const excl = result.rows.filter((x) => x.excluded);
    if (excl.length) {
      r += 2;
      ['제외한 행(합계에서 뺌)', '일자', '거래처', '입금', '출금'].forEach((h, i) => set(r, 10 + i, S(h, 3)));
      excl.forEach((x) => {
        r += 1;
        set(r, 10, S(x.전표번호, 0));
        if (x.일자) set(r, 11, { t: 'd', v: x.일자, s: 1 }); else set(r, 11, S('(변환 오류)', 0));
        set(r, 12, S(x.거래처, 0));
        set(r, 13, { t: 'n', v: x.입금, s: 2 });
        set(r, 14, { t: 'n', v: x.출금, s: 2 });
      });
    }

    if (fund) {
      r += 2;
      set(r, 10, S('자금요청 계산', 3)); set(r, 11, S('', 3)); set(r, 12, S('', 3));
      const line = (label, cell) => { r += 1; set(r, 10, S(label, 0)); set(r, 12, cell); };
      const N = (v, s) => ({ t: 'n', v, s: s || 2 });
      line('요청일', S(fund.date, 0));
      line('전일잔액(입력)', N(fund.prevBal));
      line('요청일 입금 합계', N(fund.inc));
      line('요청일 출금 합계', N(fund.out));
      line('목표 잔액', N(fund.target));
      line('요청시재(계산값)', N(fund.exact > 0 ? fund.exact : 0, 8));
      line('올림 단위', N(fund.unit));
      line('요청시재(올림값)', N(fund.rounded, 8));
      line('최종 요청시재(수정 반영)', N(fund.final, 8));
      line('최종 요청 시 예상 잔액', N(fund.afterFinal));
      line('목표 잔액 대비 차이', N(fund.diffTarget));
    }

    const entries = [
      ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
      ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="정리결과" sheetId="1" r:id="rId1"/></sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['xl/styles.xml', STYLES],
      ['xl/worksheets/sheet1.xml', sheetXml(grid, [12, 10, 14, 14, 15, 15, 13, 20, 3, 3, 30, 16, 16, 16, 18])],
    ];
    return zipStore(entries);
  }

  /* 자금요청: 요청시재 = 목표잔액 - (전일잔액 + 요청일 입금 - 요청일 출금)
   * final(최종 요청시재)은 사용자가 수정한 값. 없으면 올림값을 사용 */
  function fundRequest(rows, o) {
    const day = rows.filter((r) => r.일자 === o.date && !r.excluded);
    const inc = day.reduce((s, r) => s + r.입금, 0);
    const out = day.reduce((s, r) => s + r.출금, 0);
    const base = o.prevBal + inc - out;
    const exact = o.target - base;
    const rounded = exact > 0 ? Math.ceil(exact / o.unit) * o.unit : 0;
    const final = o.final !== undefined && o.final !== null ? o.final : rounded;
    return {
      date: o.date, prevBal: o.prevBal, target: o.target, unit: o.unit,
      count: day.length, flagged: day.filter((r) => r.errors.length).length,
      inc, out, base, exact, rounded, final,
      afterFinal: base + final, diffTarget: base + final - o.target,
    };
  }

  /* 매일 바뀌는 응원 한마디: 날짜(로컬)로 문구를 정해서 하루 동안은 같고 다음 날 바뀐다 */
  const QUOTES = [
    "숫자는 거짓말을 하지 않아요. 오늘도 차분히, 정확하게!",
    "작은 확인 하나가 큰 실수를 막아 줘요.",
    "천천히 가도 괜찮아요. 맞게 가는 게 제일 빨라요.",
    "오늘 맞춘 한 줄이 내일의 여유가 돼요.",
    "완벽하지 않아도 괜찮아요. 확인하고 고치면 돼요.",
    "꾸준함은 가장 조용하고 가장 강한 실력이에요.",
    "어려운 날에도 장부는 결국 맞춰져요. 당신 덕분에요.",
    "한 번에 하나씩 하면, 어느새 다 끝나 있어요.",
    "당신이 꼼꼼히 본 만큼 회사는 든든해져요.",
    "오늘의 수고가 쌓여서 신뢰가 돼요.",
    "차이가 나는 날은, 원인을 찾는 실력이 자라는 날이에요.",
    "급할수록 한 번 더 확인! 그게 프로의 습관이에요.",
    "잘 쉬는 것도 일의 일부예요. 물 한 잔 어때요?",
    "오늘도 돈의 길을 바르게 안내하는 당신, 멋져요.",
    "실수를 발견했다면, 이미 한 걸음 앞서 있는 거예요.",
    "지금 하는 정리가 미래의 나를 도와줘요.",
    "마감은 지나가고, 쌓은 실력은 남아요.",
    "오늘 하루도 끝까지 맞춘 당신에게 박수를!",
    "어제보다 한 가지 더 편해졌다면 충분히 성장한 거예요.",
    "정확함은 최고의 친절이에요.",
    "막히면 잠깐 멈추고 숨 고르기. 답은 곧 보여요.",
    "당신의 꼼꼼함은 팀의 안전망이에요.",
    "오늘 할 일을 하나 끝냈다면, 이미 잘하고 있는 거예요.",
    "복잡한 일도 순서대로 풀면 단순해져요.",
    "자랑스러워해도 돼요. 아무도 모르게 많은 걸 지켜내고 있어요.",
    "시작이 반이에요. 이미 파일을 열었잖아요!",
    "틀려도 괜찮아요. 확인하는 사람이 있다는 게 중요해요.",
    "오늘의 한 걸음이 내일의 지름길이 돼요.",
    "맑은 머리로 맞추는 숫자가 가장 정확해요.",
    "힘든 날일수록 작은 칭찬을 스스로에게 해 주세요.",
    "당신이 있어서 오늘도 정산이 무사히 끝나요.",
    "성실함은 결국 가장 빛나는 재능이에요.",
    "한 줄 한 줄 쌓다 보면 큰 그림이 완성돼요.",
    "오늘의 나, 어제의 나보다 조금 더 단단해졌어요.",
    "서두르지 않아도 돼요. 당신의 속도가 정답이에요.",
    "맞춰 놓고 나면 마음이 가벼워져요. 조금만 더 힘내요!",
    "당신의 정리 덕분에 누군가는 오늘 편하게 일해요.",
    "마지막까지 확인하는 당신, 정말 든든해요.",
    "오늘도 수고 많았어요. 내일은 더 가벼울 거예요.",
  ];
  function dailyQuote(d) {
    const day = Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000);
    return QUOTES[((day % QUOTES.length) + QUOTES.length) % QUOTES.length];
  }

  const api = { dailyQuote, QUOTES, fundRequest, parseXlsx, autoMap, processTable, buildXlsx, normAmount, normDate, REQUIRED, OPTIONAL };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.DD = api;
})(typeof window !== 'undefined' ? window : globalThis);
