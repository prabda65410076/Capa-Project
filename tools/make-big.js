#!/usr/bin/env node
// A workbook of the real size (after CAPA_Setup) for timing: node tools/make-big.js <out.xlsx> [items] [machines]
//   1-2: <items> rows x (A:CG inputs + CH:AGS 792 values + AGU:AHF), 1-3: <machines> rows x 792 results,
//   BY ITEM CODE, PART METHOD, 1-1.Plan, ST_SUM, CAPA_RULES (very hidden, mode VALUES).
'use strict';
const fs = require('fs');
const zlib = require('zlib');
const { once } = require('events');

const out = process.argv[2] || 'big.xlsx';
const N_ITEMS = +(process.argv[3] || 42000);
const N_MC = +(process.argv[4] || 400);
const N_CHILD = 3000, N_PLAN = 5000;
const FIRST_ROW = 7, COL_PLAN = 86, BLOCK = 66, N_PROC = 65, COL_QTY = 879, CAPA_COL = 73, CAPA_FLAG = 8;
const BULLET = '●';
let seed = 12345;
const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };

const letters = (c) => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; };
const L = []; for (let c = 1; c <= 900; c++) L[c] = letters(c);
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const num = (r, c, v) => '<c r="' + L[c] + r + '"><v>' + v + '</v></c>';
const str = (r, c, v) => '<c r="' + L[c] + r + '" t="inlineStr"><is><t>' + esc(v) + '</t></is></c>';
const row = (r, cells) => '<row r="' + r + '">' + cells.join('') + '</row>';
const SHEET_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>';
const SHEET_TAIL = '</sheetData></worksheet>';

const specs = [];
for (let i = 0; i < 50; i++) specs.push(+(4 + i * 0.37).toFixed(2));
const kinds = ['NORMAL', 'NORMAL', 'NORMAL', 'FMCNC1', 'FMCNC2', 'HAND'];

function* sheetPlan() {
  yield SHEET_HEAD;
  for (let i = 0; i < N_PLAN; i++) {
    const r = 4 + i, cells = [str(r, 3, 'ASSY-' + i)];
    for (let m = 0; m < 12; m++) cells.push(num(r, 7 + m, Math.floor(rnd() * 500)));
    yield row(r, cells);
  }
  yield SHEET_TAIL;
}

function* sheetMethod() {
  yield SHEET_HEAD;
  yield row(4, [str(4, 1, 'FMCNC1 machines'), str(4, 4, 'FMCNC1'), num(4, 5, 7)]);
  yield row(5, [str(5, 1, 'HAND machines'), str(5, 4, 'HAND')]);
  yield row(6, [str(6, 1, 'FMCNC2 machines'), str(6, 4, 'FMCNC2'), num(6, 5, 7.5)]);
  for (let i = 0; i < 200; i++) {
    const r = 8 + i, how = ['FMCNC1', 'FMCNC2', 'HAND'][i % 3];
    yield row(r, [str(r, 1, 'CH-' + (i * 7)), str(r, 2, how === 'HAND' ? 'BD' : 'FM'), str(r, 3, how)].concat(i % 2 ? [num(r, 4, 5 + (i % 9))] : []));
  }
  yield SHEET_TAIL;
}

function* sheet12() {
  yield SHEET_HEAD;
  const r3 = [], r4 = [], r5 = [str(5, 6, 'ASSY P/N'), str(5, 18, 'FM'), str(5, 19, 'BD')];
  for (let k = 1; k <= N_PROC; k++) {
    r4.push(str(4, 20 + k, k >= 8 && k <= 15 ? 'Forming' : 'Group ' + Math.ceil(k / 10)));
    r5.push(str(5, 20 + k, 'Process ' + k));
  }
  for (let m = 0; m < 12; m++) {
    const c = COL_PLAN + m * BLOCK;
    r4.push(str(4, c, 'Plan'));
    for (let k = 1; k <= N_PROC; k++) r3.push(num(3, c + k, rnd() < 0.8 ? 1 : 1.1));
  }
  yield row(3, r3);
  yield row(4, r4);
  yield row(5, r5);
  for (let i = 0; i < N_ITEMS; i++) {
    const r = FIRST_ROW + i;
    const cells = [str(r, 6, 'ASSY-' + Math.floor(rnd() * N_PLAN * 1.05)), str(r, 8, 'CH-' + (i % N_CHILD)), str(r, 9, 'Part ' + i),
      num(r, 11, 1 + (i % 3)), num(r, 12, specs[i % specs.length]), num(r, 16, i % 7 ? 1 : 0),
      str(r, 18, i % 3 ? 'FM' : '0'), str(r, 19, i % 4 ? 'BD' : '0')];
    for (let k = 1; k <= N_PROC; k++) if (rnd() < 0.35) cells.push(num(r, 20 + k, Math.round(rnd() * 6000) / 100));
    for (let c = COL_PLAN; c < COL_PLAN + 12 * BLOCK; c++) cells.push(num(r, c, rnd() * 1000));
    for (let m = 0; m < 12; m++) cells.push(num(r, COL_QTY + m, Math.floor(rnd() * 800)));
    yield row(r, cells);
  }
  yield SHEET_TAIL;
}

function* sheetSum() {
  yield SHEET_HEAD;
  yield row(5, [str(5, 8, 'ChildP/N'), str(5, 12, 'OD'), str(5, 13, 'FM'), str(5, 14, 'BD')]);
  for (let i = 0; i < 20; i++) yield row(FIRST_ROW + i, [num(FIRST_ROW + i, 12, specs[i]), num(FIRST_ROW + i, 87, 1)]);
  yield SHEET_TAIL;
}

function* sheet13() {
  yield SHEET_HEAD;
  const r3 = [], r5 = [];
  for (let m = 0; m < 12; m++) for (let k = 0; k < 66; k++) {
    const c = CAPA_COL + m * BLOCK + k;
    r3.push(num(3, c, 20 + (m % 3)));
    r5.push(str(5, c, 'L1'));
  }
  yield row(3, r3);
  yield row(4, [str(4, 8, 'A/T'), str(4, 73, 'A/T')]);
  yield row(5, r5);
  for (let i = 0; i < N_MC; i++) {
    const r = FIRST_ROW + i, kind = kinds[i % kinds.length];
    const cells = [str(r, 1, 'L1'), str(r, 2, kind), str(r, 3, 'MC-' + i), num(r, 5, 16), num(r, 7, specs[i % specs.length])];
    for (let k = 1; k <= N_PROC; k++) {
      const on = kind === 'HAND' ? k === 16 : kind.startsWith('FMCNC') ? k >= 8 && k <= 15 : rnd() < 0.5;
      if (on) cells.push(str(r, CAPA_FLAG + k - 1, BULLET));
    }
    for (let c = CAPA_COL; c < CAPA_COL + 12 * BLOCK; c++) cells.push(num(r, c, rnd()));
    yield row(r, cells);
  }
  yield SHEET_TAIL;
}

function* sheetItem() {
  yield SHEET_HEAD;
  for (let i = 0; i < N_CHILD; i++) {
    const r = 4 + i, cells = [str(r, 7, 'CH-' + i)];
    for (let m = 0; m < 13; m++) cells.push(num(r, 8 + m, Math.floor(rnd() * 900)));
    yield row(r, cells);
  }
  yield SHEET_TAIL;
}

function rules() {
  const k1 = "'PART METHOD'!R4C4", k2 = "'PART METHOD'!R6C4", kb = "'PART METHOD'!R5C4";
  const list = [];
  const last12 = FIRST_ROW + N_ITEMS - 1, last13 = FIRST_ROW + N_MC - 1;
  for (let m = 0; m < 12; m++) {
    const col = COL_QTY + m, pc = COL_PLAN + m * BLOCK;
    list.push(['1-2.ITEM', '=IF(RC16>0,RC[' + (pc - col) + ']*RC11,0)', FIRST_ROW + ' ' + col + ' ' + last12 + ' ' + col]);
  }
  for (let k = 1; k <= 66; k++) {
    const fl = CAPA_FLAG + k - 1, cnt = 'COUNTIFS(C7,RC7,C' + fl + ',"' + BULLET + '",C1,R5C,C2,';
    let f;
    if (k === 66) f = '=SUM(RC[-65]:RC[-1])';
    else if (k >= 8 && k <= 15) {
      f = '=IF(RC' + fl + '="' + BULLET + '",SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7,ST_SUM!C13,IF(RC2=' + k1 + ',"FMCNC1",IF(RC2=' + k2 +
        ',"FMCNC2","STD")))/IF(RC2=' + k1 + ',' + cnt + k1 + '),IF(RC2=' + k2 + ',' + cnt + k2 + '),' + cnt + '"<>"&' + k1 + ',C2,"<>"&' +
        k2 + ')))/R3C,0)/3600/RC5';
    } else if (k === 16) {
      f = '=IF(RC' + fl + '="' + BULLET + '",SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7,ST_SUM!C14,IF(RC2=' + kb + ',"HAND","STD"))/' + cnt +
        'IF(RC2=' + kb + ',' + kb + ',"<>"&' + kb + '))/R3C,0)/3600/RC5';
    } else {
      f = '=IF(RC' + fl + '="' + BULLET + '",SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7)/COUNTIFS(C7,RC7,C' + fl + ',"' + BULLET + '",C1,R5C)/R3C,0)/3600/RC5';
    }
    const rects = [];
    for (let m = 0; m < 12; m++) { const c = CAPA_COL + m * BLOCK + k - 1; rects.push(FIRST_ROW + ' ' + c + ' ' + last13 + ' ' + c); }
    list.push(['1-3.CAPA', f, rects.join(',')]);
  }
  for (let m = 0; m < 12; m++) list.push(['BY ITEM CODE', '=SUMIF(ST_SUM!C8,RC7,ST_SUM!C' + (COL_QTY + m) + ')', '4 ' + (8 + m) + ' ' + (3 + N_CHILD) + ' ' + (8 + m)]);
  const checks = [['1-2.ITEM', '$A$1:$' + L[890] + '$' + last12], ['1-3.CAPA', '$A$1:$' + L[CAPA_COL + 12 * BLOCK - 1] + '$' + last13],
    ['BY ITEM CODE', '$A$1:$T$' + (3 + N_CHILD)]];
  return { list, checks };
}

function* sheetRules(R) {
  yield SHEET_HEAD;
  yield row(1, [str(1, 1, 'CAPA formula rules - written by macro CAPA_Setup. Do not edit.'), str(1, 3, 'VALUES')]);
  let r = 3;
  const n = Math.max(R.list.length, R.checks.length);
  for (let i = 0; i < n; i++, r++) {
    const cells = [];
    if (i < R.list.length) cells.push(str(r, 1, R.list[i][0]), str(r, 2, R.list[i][1]), str(r, 3, R.list[i][2]));
    if (i < R.checks.length) cells.push(str(r, 6, R.checks[i][0]), str(r, 7, 'CAPA_CHK_' + (i + 1)), str(r, 8, R.checks[i][1]));
    yield row(r, cells);
  }
  yield SHEET_TAIL;
}

// ---- zip ---------------------------------------------------------------------------------
async function main() {
  const fd = fs.openSync(out, 'w');
  let off = 0;
  const cd = [];
  const write = (b) => { fs.writeSync(fd, b); off += b.length; };
  const header = (name, crc, csize, usize) => {
    const nb = Buffer.from(name);
    const h = Buffer.alloc(30);
    h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(20, 4); h.writeUInt16LE(8, 8);
    h.writeUInt32LE(crc >>> 0, 14); h.writeUInt32LE(csize, 18); h.writeUInt32LE(usize, 22); h.writeUInt16LE(nb.length, 26);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(8, 10);
    c.writeUInt32LE(crc >>> 0, 16); c.writeUInt32LE(csize, 20); c.writeUInt32LE(usize, 24); c.writeUInt16LE(nb.length, 28);
    c.writeUInt32LE(off, 42);
    cd.push(Buffer.concat([c, nb]));
    write(Buffer.concat([h, nb]));
  };
  const addBuf = (name, s) => {
    const b = Buffer.from(s), z = zlib.deflateRawSync(b);
    header(name, zlib.crc32(b), z.length, b.length);
    write(z);
  };
  const addGen = async (name, gen) => {
    const tmp = out + '.part';
    const ws = fs.createWriteStream(tmp);
    const def = zlib.createDeflateRaw({ level: 3 });
    def.pipe(ws);
    let crc = 0, usize = 0, buf = [], n = 0;
    const flush = async () => {
      const b = Buffer.from(buf.join('')); buf = []; n = 0;
      crc = zlib.crc32(b, crc); usize += b.length;
      if (!def.write(b)) await once(def, 'drain');
    };
    for (const s of gen) { buf.push(s); n += s.length; if (n > 4e6) await flush(); }
    await flush();
    def.end();
    await once(ws, 'finish');
    const csize = fs.statSync(tmp).size;
    header(name, crc, csize, usize);
    const rfd = fs.openSync(tmp, 'r');
    const chunk = Buffer.alloc(1 << 24);
    let k;
    while ((k = fs.readSync(rfd, chunk, 0, chunk.length, null)) > 0) write(chunk.subarray(0, k));
    fs.closeSync(rfd);
    fs.unlinkSync(tmp);
    console.log(name, (usize / 1048576).toFixed(0) + ' MB XML');
  };
  const names = ['1-1.Plan', 'PART METHOD', '1-2.ITEM', 'ST_SUM', '1-3.CAPA', 'BY ITEM CODE', 'CAPA_RULES'];
  const R = rules();
  addBuf('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    names.map((n, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') +
    '</Types>');
  addBuf('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  addBuf('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
    names.map((n, i) => '<sheet name="' + esc(n) + '" sheetId="' + (i + 1) + '"' + (n === 'CAPA_RULES' ? ' state="veryHidden"' : '') + ' r:id="rId' + (i + 1) + '"/>').join('') +
    '</sheets><definedNames>' + R.checks.map((c, i) => '<definedName name="CAPA_CHK_' + (i + 1) + '" hidden="1">\'' + esc(c[0]) + '\'!' + c[1] + '</definedName>').join('') +
    '</definedNames></workbook>');
  addBuf('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    names.map((n, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
    '</Relationships>');
  const gens = [sheetPlan(), sheetMethod(), sheet12(), sheetSum(), sheet13(), sheetItem(), sheetRules(R)];
  for (let i = 0; i < names.length; i++) await addGen('xl/worksheets/sheet' + (i + 1) + '.xml', gens[i]);
  const cdBuf = Buffer.concat(cd);
  const e = Buffer.alloc(22);
  e.writeUInt32LE(0x06054b50, 0); e.writeUInt16LE(cd.length, 8); e.writeUInt16LE(cd.length, 10);
  e.writeUInt32LE(cdBuf.length, 12); e.writeUInt32LE(off, 16);
  write(cdBuf);
  write(e);
  fs.closeSync(fd);
  console.log(out, (fs.statSync(out).size / 1048576).toFixed(1) + ' MB');
}
main().catch((e) => { console.error(e); process.exit(1); });
