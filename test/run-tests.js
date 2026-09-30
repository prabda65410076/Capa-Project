#!/usr/bin/env node
// Tests of the CAPA web engine. Run: node test/run-tests.js   (after python3 test/make_fixtures.py)
'use strict';
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCore, fileFrom } = require('../tools/core.js');

const FIX = path.join(__dirname, 'fixtures');
const TMP = path.join(__dirname, 'fixtures', 'out');
let failed = 0, passed = 0;

function check(name, ok, detail) {
  if (ok) { passed++; return; }
  failed++;
  console.log('FAIL', name, detail === undefined ? '' : detail);
}

function same(a, b) {
  if (typeof a === 'number' && typeof b === 'number') {
    const m = Math.max(1, Math.abs(a), Math.abs(b));
    return Math.abs(a - b) <= m * 1e-9;
  }
  return JSON.stringify(a) === JSON.stringify(b);
}

// LibreOffice counts TRUE as 1 inside SUM / COUNT ranges, Excel does not: the Excel values.
const EXCEL_NOT_LO = { 'TEST!C14': 2.5, 'TEST!C15': 4 };
// openpyxl cannot write a cell that holds empty text: such a cell is blank in the fixtures.
const EMPTY_TEXT = new Set(['TEST!C46']);

//------------------------------------------------------------------------------
//  Engine semantics on a small workbook
//------------------------------------------------------------------------------
function evalBook(c, sheets, formulas) {
  const wb = new c.__x.Workbook(null, null);
  for (const [name, cells] of Object.entries(sheets)) {
    const sh = wb.addSheet(name);
    for (const [a, v] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      sh.put(+m[2], c.colNumber(m[1]), v);
    }
  }
  c.engReset(wb);
  const list = [];
  for (const [ref, f] of Object.entries(formulas)) {
    const [sn, a] = ref.split('!');
    const m = /^([A-Z]+)(\d+)$/.exec(a);
    const sh = wb.sheet(sn), r = +m[2], col = c.colNumber(m[1]);
    const r1c1 = f.startsWith('=R') || f.includes('RC') ? f : c.a1ToR1C1(f.slice(1), r, col);
    const ru = c.engAddRule(sh.index, r1c1, { kind: 'cell' });
    c.engAddCell(sh.index, r, col, ru);
    list.push([ref, sh, r, col, ru]);
  }
  c.engCompileAll();
  const bad = {};
  for (const [ref, , , , ru] of list) if (c.engRuleBad(ru) !== null) bad[ref] = c.engRuleBad(ru);
  c.engBuildCells();
  c.engEvaluate('');
  c.engWrite(1);
  const out = {};
  for (const [ref, sh, r, col] of list) {
    const v = sh.get(r, col);
    out[ref] = v instanceof c.__x.XErr ? { e: v.text } : v;
  }
  return { out, bad };
}

function testSemantics() {
  const c = loadCore();
  const E = (t) => ({ e: t });
  const data = {
    S: {
      A1: 5, A2: '5', A3: ' 5', A4: '05', A5: '$5', A6: 'abc', A7: 'ABC', A8: true, A9: c.errOf('#N/A'), A10: '',
      A12: 7, A13: 'a*c', B1: 1, B2: 2, B3: 4, B4: 8, B5: 16, B6: 32, B7: 64, B8: 128, B9: 256, B10: 512, B11: 1024,
      B12: 2048, B13: 4096, D1: 'TRUE', D2: 'maybe', E1: 10, E2: 20, E3: 30, F2: 'x',
      H1: 'b', H2: 'a', H3: 3, I1: 'B!', I2: 'A!', I3: 'three',
    },
  };
  const f = {
    // criteria (VBA comments: measured in Excel)
    'S!C1': '=SUMIF(A1:A13,5,B1:B13)',          // 5 "5" " 5" "05" "$5" -> 1+2+4+8+16
    'S!C2': '=SUMIF(A1:A13,C20,B1:B13)',        // blank criteria cell = 0: nothing
    'S!C3': '=SUMIF(A1:A13,"",B1:B13)',         // blank (A11) and empty text (A10)
    'S!C4': '=SUMIF(A1:A13,"=",B1:B13)',        // only blank
    'S!C5': '=SUMIF(A1:A13,"<>",B1:B13)',       // not blank
    'S!C6': '=SUMIF(A1:A13,"<>5",B1:B13)',      // everything but the number 5
    'S!C7': '=SUMIF(A1:A13,">4",B1:B13)',       // numbers only: 5, 7
    'S!C8': '=SUMIF(A1:A13,"abc",B1:B13)',      // case ignored
    'S!C9': '=SUMIF(A1:A13,"a?c",B1:B13)',      // wildcard (a*c matches too)
    'S!C10': '=SUMIF(A1:A13,"a~*c",B1:B13)',    // escaped *
    'S!C11': '=SUMIF(A1:A13,"TRUE",B1:B13)',    // the logical value
    'S!C12': '=SUMIF(A1:A13,"#N/A",B1:B13)',    // an error matches the same error
    'S!C13': '=COUNTIF(A1:A20,"")',             // blank rows up to the range end count
    'S!C14': '=COUNTIFS(A1:A13,">=5",B1:B13,"<100")',
    'S!C15': '=SUMIF(A1:A13,"<>abc",B1:B13)',
    // arithmetic and conversions
    'S!C16': '=-2^2',
    'S!C17': '=(-8)^(1/3)',
    'S!C18': '=1-(1-2^-50)',
    'S!C19': '=1-(1-2^-49)',
    'S!C21': '=A2+1',
    'S!C22': '=A5*2',
    'S!C23': '=A6+1',
    'S!C24': '=A8+1',
    'S!C25': '=IF(D1,1,2)',
    'S!C26': '=IF(D2,1,2)',
    'S!C27': '=IF(A9,1,2)',
    'S!C28': '=E1:E3*2',                        // implicit intersection (row 28: none) -> #VALUE!
    'S!F1': '=E1:E3*2',                         // row 1 -> E1
    'S!F3': '=SUM(E1:E3)+E1:E3',                // row 3 -> E3
    'S!C29': '=0.1+0.2&""',
    'S!C30': '=1/3&""',
    'S!C31': '=A1/0',
    'S!C32': '=MID("abcdef",2,3)&MID("abc",0,1)',
    'S!C33': '=INT(-2.5)+INT(2.5)',
    'S!C34': '=VLOOKUP("a",H1:I3,2,FALSE)',
    'S!C35': '=VLOOKUP(3,H1:I3,2,FALSE)',
    'S!C36': '=VLOOKUP("3",H1:I3,2,FALSE)',     // text does not find a number
    'S!C37': '=VLOOKUP("A*",H1:I3,2,0)',
    'S!C38': '=VLOOKUP("a",H1:I3,3,FALSE)',     // column beyond the table
    'S!C39': '=IFERROR(A9,"err")',
    'S!C40': '=IFERROR(A1,"err")',
    'S!C41': '=A11',                            // blank -> 0
    'S!C42': '=A11=""',
    'S!C43': '="a"<"B"',
    'S!C44': '=1<"0"',                          // numbers < text
    'S!C45': '="x"<TRUE',                       // text < logical
    'S!C46': '=SUM(A1:A13)',                    // an error in the range wins
    'S!C47': '=SUM(A2,A8,1)',                   // text and TRUE in referenced cells do not count
    'S!C48': '=COUNT(A1:A13,A2,"x")',
    'S!C49': '=NA()',
    'S!C50': '=TRUE()+FALSE()',
    'S!C51': '=10%*A1',
    'S!C52': '=IF(A1>1,"big")&IF(A1<1,"small")',
    'S!C53': '=SUMIFS(B1:B13,A1:A13,"<>5",A1:A13,"<>abc")',
  };
  const { out, bad } = evalBook(c, data, f);
  const want = {
    'S!C1': 31, 'S!C2': 0, 'S!C3': 512 + 1024, 'S!C4': 1024, 'S!C5': 8191 - 1024, 'S!C6': 8191 - 1,
    'S!C7': 1 + 2048, 'S!C8': 32 + 64, 'S!C9': 32 + 64 + 4096, 'S!C10': 4096, 'S!C11': 128, 'S!C12': 256, 'S!C13': 9,
    'S!C14': 1, 'S!C15': 8191 - 32 - 64, 'S!C16': 4, 'S!C17': -2, 'S!C18': 0, 'S!C19': Math.pow(2, -49),
    'S!C21': 6, 'S!C22': 10, 'S!C23': E('#VALUE!'), 'S!C24': 2, 'S!C25': 1, 'S!C26': E('#VALUE!'), 'S!C27': E('#N/A'),
    'S!C28': E('#VALUE!'), 'S!F1': 20, 'S!F3': 90, 'S!C29': '0.3', 'S!C30': '0.333333333333333', 'S!C31': E('#DIV/0!'),
    'S!C32': E('#VALUE!'), 'S!C33': -1, 'S!C34': 'A!', 'S!C35': 'three', 'S!C36': E('#N/A'), 'S!C37': 'A!',
    'S!C38': E('#REF!'), 'S!C39': 'err', 'S!C40': 5, 'S!C41': 0, 'S!C42': true, 'S!C43': true, 'S!C44': true,
    'S!C45': true, 'S!C46': E('#N/A'), 'S!C47': 1, 'S!C48': 2, 'S!C49': E('#N/A'), 'S!C50': 1, 'S!C51': 0.5, 'S!C52': 'bigFALSE',
    'S!C53': 8191 - 1 - 32 - 64,
  };
  check('semantics: no compile errors', Object.keys(bad).length === 0, JSON.stringify(bad));
  for (const k of Object.keys(want)) check('semantics ' + k + ' ' + f[k], same(out[k], want[k]), JSON.stringify(out[k]) + ' want ' + JSON.stringify(want[k]));

  // formulas the engine does not do
  const c2 = loadCore();
  const r2 = evalBook(c2, { S: { A1: 1 } }, {
    'S!B1': '=ROUND(A1,0)', 'S!B2': '=Plan_Range+1', 'S!B3': '={1,2}', 'S!B4': "=[1]Sheet1!A1", 'S!B5': '=A1+',
  });
  check('unsupported formulas are reported', Object.keys(r2.bad).length === 5, JSON.stringify(r2.bad));
  // approximate VLOOKUP stops (as in the VBA engine), circular references stop
  for (const [name, fm] of [['vlookup TRUE', { 'S!B1': '=VLOOKUP(1,A1:A2,1,TRUE)' }], ['circular', { 'S!B1': '=B2', 'S!B2': '=B1+1' }]]) {
    let msg = '';
    try { evalBook(loadCore(), { S: { A1: 1 } }, fm); } catch (e) { msg = e.message; }
    check(name + ' stops with a message', msg.length > 0, msg);
  }
  // A1 <-> R1C1
  check('a1ToR1C1', c.a1ToR1C1("SUMIFS('1-2.X'!CI:CI,$L:$L,$G7)+A$1+1:1", 7, 73) ===
    "=SUMIFS('1-2.X'!C[14],C12,RC7)+R1C[-72]+R[-6]", c.a1ToR1C1("SUMIFS('1-2.X'!CI:CI,$L:$L,$G7)+A$1+1:1", 7, 73));
  check('r1c1ToA1', c.r1c1ToA1("=SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7)/R3C", 7, 73) === '=SUMIFS(ST_SUM!CI:CI,ST_SUM!$L:$L,$G7)/BU$3',
    c.r1c1ToA1("=SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7)/R3C", 7, 73));
  // text that is a number
  const tn = [[' 5 ', 5], ['$1,000.5', 1000.5], ['(5)', -5], ['5%', 0.05], ['-$5', -5], ['1e3', 1000], ['1,00', null],
    ['1-2', null], ['abc', null], ['', null], ['.5', 0.5], ['5.', 5]];
  for (const [s, v] of tn) check('textToNum ' + JSON.stringify(s), c.textToNum(s) === v, c.textToNum(s));
  check('cnc2Shots', c.cnc2Shots(3, 1) === 2 && c.cnc2Shots(5, 2) === 3 && c.cnc2Shots(4, 2) === 2 && c.cnc2Shots(2.5, 1) === 1.25,
    [c.cnc2Shots(3, 1), c.cnc2Shots(5, 2), c.cnc2Shots(4, 2), c.cnc2Shots(2.5, 1)]);
}

//------------------------------------------------------------------------------
//  Fixture workbooks
//------------------------------------------------------------------------------
function wv(c, v) { return v instanceof c.__x.XErr ? { e: v.text } : v; }

async function recalcFile(file) {
  const c = loadCore();
  await c.apiCall('load', { file: fileFrom(path.join(FIX, file)) });
  const r = await c.apiCall('recalc', {});
  return { c, r };
}

function unexpectedChanges(r) {
  const bad = [];
  for (const ch of r.changes) {
    for (const x of ch.ex) {
      const k = ch.sheet + '!' + x.a;
      if (k in EXCEL_NOT_LO || EMPTY_TEXT.has(k)) continue;
      bad.push(k + ' ' + x.old + ' -> ' + x.val);
    }
  }
  return bad;
}

function compareExpected(c, getter, expected, label) {
  let n = 0, bad = 0;
  const shown = [];
  for (const [sh, r, col, v0] of expected) {
    const k = sh + '!' + c.addr(r, col);
    let v = v0;
    if (k in EXCEL_NOT_LO) v = EXCEL_NOT_LO[k];
    const got = getter(sh, r, col);
    n++;
    if (EMPTY_TEXT.has(k) && (got === '' || got === null) && (v === '' || v === null)) continue;
    if (!same(got, v)) {
      bad++;
      if (shown.length < 12) shown.push(k + ' got ' + JSON.stringify(got) + ' want ' + JSON.stringify(v));
    }
  }
  check(label + ' (' + n + ' cells)', bad === 0, bad + ' differ: ' + shown.join('; '));
}

async function testFixtures() {
  for (const f of ['A_formulas.xlsx', 'A_values.xlsx']) {
    const { r } = await recalcFile(f);
    const bad = unexpectedChanges(r);
    check(f + ': every result equals the value in the file', bad.length === 0, bad.slice(0, 10).join('; '));
    check(f + ': ST_SUM equals the oracle', r.stSumDiff === 0, r.stSumDiff);
    check(f + ': no unsupported formula', r.unsupportedCount === 0, JSON.stringify(r.unsupported.slice(0, 3)));
    check(f + ': PART METHOD notes', /ไม่พบในชีต 1-2/.test(r.methodNote) && /แถวไม่ถูกใช้/.test(r.methodNote) &&
      /เกือบตรง/.test(r.methodNote), r.methodNote);
  }
  // inputs of B, values of A in the cells: the recalculation must give B
  const expected = JSON.parse(fs.readFileSync(path.join(FIX, 'expected_B.json'), 'utf8'));
  const { c, r } = await recalcFile('AB_values.xlsx');
  const wb = c.__x.session.wb;
  check('AB: something changed', r.changedCells > 1000, r.changedCells);
  compareExpected(c, (sh, row, col) => wv(c, wb.sheet(sh).get(row, col)), expected, 'AB_values recalculated = B');

  // export, read back
  fs.mkdirSync(TMP, { recursive: true });
  const t0 = Date.now();
  const ex = await c.apiCall('export', { opts: { detail: true } });   // with the item values of 1-2
  const buf = Buffer.from(await ex.blob.arrayBuffer());
  const outPath = path.join(TMP, 'AB_exported.xlsx');
  fs.writeFileSync(outPath, buf);
  check('export name', ex.name === 'AB_values_web.xlsx', ex.name);
  console.log('  export', ((Date.now() - t0) / 1000).toFixed(2) + 's', buf.length, 'bytes');
  const c2 = loadCore();
  const wbFull = await c2.loadWorkbook(fileFrom(outPath), {}, null);   // every cell, 1-2 CH:AGS too
  compareExpected(c2, (sh, row, col) => wv(c2, wbFull.sheet(sh).get(row, col)), expected, 'exported file holds B');
  await c2.apiCall('load', { file: fileFrom(outPath) });
  const r2 = await c2.apiCall('recalc', {});
  const bad2 = unexpectedChanges(r2);
  check('exported file: a second recalculation changes nothing', bad2.length === 0, bad2.slice(0, 10).join('; '));
  // other programs can read it
  const py = execFileSync('python3', ['-c', `
import openpyxl, json, sys
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb['1-3.CAPA']
print(json.dumps([wb.sheetnames, ws['BU7'].value, ws.sheet_state, wb['CAPA_RULES'].sheet_state]))
`, outPath]).toString();
  check('openpyxl reads the exported file', /1-3\.CAPA/.test(py), py);
  const loDir = path.join(TMP, 'lo');
  fs.rmSync(loDir, { recursive: true, force: true });
  let loOk = true;
  try {
    execFileSync('soffice', ['--headless', '--norestore', '--convert-to', 'csv', '--outdir', loDir, outPath], { stdio: 'pipe', timeout: 300000 });
  } catch (e) { loOk = false; }
  check('LibreOffice opens the exported file', loOk && fs.existsSync(path.join(loDir, 'AB_exported.csv')));

  // export without the item values of sheet 1-2 (CH:AGS keeps only the Plan columns)
  const ex2 = await c.apiCall('export', { opts: { detail: false } });
  const out2 = path.join(TMP, 'AB_light.xlsx');
  fs.writeFileSync(out2, Buffer.from(await ex2.blob.arrayBuffer()));
  const c3 = loadCore();
  const wb3 = await c3.loadWorkbook(fileFrom(out2), {}, null);
  const s12 = wb3.sheet('1-2.ITEM');
  let planOk = true, procEmpty = true;
  for (let row = 7; row < 7 + 64; row++) {
    for (let m = 0; m < 12; m++) {
      const pc = 86 + m * 66;
      const want = expected.find((e) => e[0] === '1-2.ITEM' && e[1] === row && e[2] === pc);
      if (want && !same(wv(c3, s12.get(row, pc)), want[3])) planOk = false;
      if (s12.get(row, pc + 1) !== null) procEmpty = false;
    }
  }
  check('light export keeps the Plan columns', planOk);
  check('light export leaves out the process columns', procEmpty);
  console.log('  light export', fs.statSync(out2).size, 'bytes');

  // structure check: rows inserted after the conversion
  const c4 = loadCore();
  await c4.apiCall('load', { file: fileFrom(path.join(FIX, 'A_values.xlsx')) });
  const n0 = c4.__x.session.wb.names.find((n) => n.name === 'CAPA_CHK_1');
  n0.text = n0.text.replace(/\$(\d+)$/, (m, d) => '$' + (+d + 1));
  let msg = '';
  try { await c4.apiCall('recalc', {}); } catch (e) { msg = e.message; }
  check('inserted rows are found', /แทรกหรือลบแถว/.test(msg), msg);
}

(async () => {
  const t0 = Date.now();
  testSemantics();
  if (fs.existsSync(path.join(FIX, 'AB_values.xlsx'))) await testFixtures();
  else console.log('(no fixtures: run python3 test/make_fixtures.py)');
  console.log(passed + ' passed, ' + failed + ' failed (' + ((Date.now() - t0) / 1000).toFixed(1) + 's)');
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.log(e.stack); process.exit(1); });
