'use strict';
//==============================================================================
//  CAPA formula engine: a port of modCapaEngine (VBA).
//
//  Every formula is a rule (sheet + R1C1 text) that is compiled to tokens in
//  reverse Polish notation and evaluated the way Excel does. Formulas may use
//  numbers, text, TRUE/FALSE, error values, R1C1 references (other sheets,
//  ranges, whole columns or rows), + - * / ^ & % = <> < > <= >= and the
//  functions IF IFERROR SUM COUNT SUMIF SUMIFS COUNTIF COUNTIFS VLOOKUP INT MID NA
//  (and TRUE() / FALSE(), which LibreOffice writes for TRUE / FALSE).
//==============================================================================

const MAX_ROW = 1048576;
const MAX_COL = 16384;

// ---- tokens ------------------------------------------------------------------
const T_NUM = 1, T_MISS = 2, T_REF = 3;
const T_ADD = 10, T_SUB = 11, T_MUL = 12, T_DIV = 13, T_POW = 14, T_CAT = 15;
const T_EQ = 20, T_NE = 21, T_LT = 22, T_LE = 23, T_GT = 24, T_GE = 25;
const T_NEG = 30, T_PLUS = 31, T_PCT = 32;
const T_JMPF = 40, T_JMP = 41, T_JNE = 42, T_FUNC = 50;
const T_ARK = 60, T_ARR = 61, T_IFCK = 62;

const F_SUM = 1, F_COUNT = 2, F_SUMIF = 3, F_SUMIFS = 4, F_COUNTIF = 5, F_COUNTIFS = 6;
const F_VLOOKUP = 7, F_INT = 8, F_MID = 9, F_NA = 10;

const RF_R1 = 1, RF_C1 = 2, RF_R2 = 4, RF_C2 = 8, RF_COLS = 16, RF_ROWS = 32;

// ---- criteria ------------------------------------------------------------------
const K_NUM = 1, K_TEXT = 2, K_LIKE = 3, K_BOOL = 4, K_ERR = 5, K_EMPTYSTR = 6, K_BLANK = 7;
const OP_EQ = 0, OP_NE = 1, OP_LT = 2, OP_LE = 3, OP_GT = 4, OP_GE = 5;

class CompileError extends Error {}

// ---- state -------------------------------------------------------------------
let engWb = null;
let shs = [];                 // per sheet: {sheet, br1, bc1, br2, bc2, ids, cdone, nRow, nCol}
let nRule = 0, ruSheet = [], ruText = [], ruTk0 = [], ruTk1 = [], ruSrc = [], ruBad = [], ruByKey = new Map();
let nTk = 0, tkCap = 0;
let tkOp = null, tkA = null, tkB = null, tkC = null, tkD = null, tkE = null, tkF = null, tkV = [];
let nCell = 0, ceSheet = null, ceRow = null, ceCol = null, ceRule = null, ceState = null, ceSkip = null;
let ceVal = [], ceOld = [];
let capS = [], capR = [], capC = [], capU = [];
let sp = 0, stCap = 0;
let stK = null, stV = [], stS = null, stR1 = null, stC1 = null, stR2 = null, stC2 = null;
let curRow = 0, curCol = 0, curTok = 0;
let tcN = null, tcS = null, tcC = null, tcR1 = null, tcR2 = null;
let aggCache = new Map(), ixByKey = new Map(), ix = [];
let psS = '', psI = 0, psN = 0, psSheet = 0;
let wantSh = [];
let engTimeNote = '';

//==============================================================================
//  Model: rules and formula cells
//==============================================================================
function engReset(wb) {
  engWb = wb;
  shs = wb.sheets.map((sheet) => ({ sheet, br1: 0, bc1: 0, br2: -1, bc2: -1, ids: null, cdone: null, nRow: 0, nCol: 0 }));
  nRule = 0; ruSheet = [0]; ruText = ['']; ruTk0 = [0]; ruTk1 = [0]; ruSrc = [null]; ruBad = [null];
  ruByKey = new Map();
  nTk = 0; tkCap = 0; tkV = [];
  tkOp = tkA = tkB = tkC = tkD = tkE = tkF = null;
  growTokens();
  nCell = 0;
  capS = []; capR = []; capC = []; capU = [];
}

// Rule of formula f (R1C1, with "=") in sheet s. src: where it comes from (for messages).
function engAddRule(s, f, src) {
  const k = s + '\t' + f;
  const found = ruByKey.get(k);
  if (found !== undefined) return found;
  nRule++;
  ruSheet[nRule] = s; ruText[nRule] = f; ruTk0[nRule] = 0; ruTk1[nRule] = -1; ruSrc[nRule] = src; ruBad[nRule] = null;
  ruByKey.set(k, nRule);
  return nRule;
}

function engAddCell(s, r, c, ru) {
  capS.push(s); capR.push(r); capC.push(c); capU.push(ru);
}

// Compiles every rule. A rule that cannot be compiled gets ruBad (the message); its
// cells are left out by engBuildCells.
function engCompileAll() {
  for (let ru = 1; ru <= nRule; ru++) {
    try {
      compileRule(ru);
    } catch (e) {
      if (!(e instanceof CompileError)) throw e;
      ruBad[ru] = e.message;
      ruTk1[ru] = ruTk0[ru] - 1;
    }
  }
}

function engRuleBad(ru) { return ruBad[ru]; }
function engRuleText(ru) { return ruText[ru]; }
function engRuleSrc(ru) { return ruSrc[ru]; }
function engRuleCount() { return nRule; }
function engCellCount() { return nCell; }
function engTokenCount() { return nTk; }

// Box and cell ids per sheet; cells are numbered by sheet, column, row.
// ids[c - bc1] is an Int32Array over rows br1..br2 (null for a column without rule cells).
function engBuildCells() {
  const n = capS.length;
  for (const S of shs) { S.br1 = 0; S.br2 = -1; S.bc1 = 0; S.bc2 = -1; S.ids = null; S.cdone = null; }
  const cellsOfRule = new Int32Array(nRule + 1);
  for (let k = 0; k < n; k++) {
    if (ruBad[capU[k]] !== null) continue;
    const S = shs[capS[k]], r = capR[k], c = capC[k];
    if (S.br2 < 0) { S.br1 = S.br2 = r; S.bc1 = S.bc2 = c; continue; }
    if (r < S.br1) S.br1 = r;
    if (r > S.br2) S.br2 = r;
    if (c < S.bc1) S.bc1 = c;
    if (c > S.bc2) S.bc2 = c;
  }
  for (const S of shs) {
    if (S.br2 >= 0) {
      S.ids = new Array(S.bc2 - S.bc1 + 1).fill(null);
      S.cdone = new Uint8Array(S.bc2 - S.bc1 + 1);
    }
  }
  for (let k = 0; k < n; k++) {
    if (ruBad[capU[k]] !== null) continue;
    const S = shs[capS[k]], r = capR[k], c = capC[k];
    let col = S.ids[c - S.bc1];
    if (col === null) col = S.ids[c - S.bc1] = new Int32Array(S.br2 - S.br1 + 1);
    if (col[r - S.br1] !== 0) {
      throw new UserError('เซลล์ ' + S.sheet.name + '!' + addr(r, c) + ' มีสูตรซ้ำสองครั้งในกฎสูตร (CAPA_RULES)');
    }
    col[r - S.br1] = -capU[k];
  }
  let m = 0;
  for (let k = 0; k < n; k++) if (ruBad[capU[k]] === null) m++;
  const sz = m + 1;
  ceSheet = new Int32Array(sz); ceRow = new Int32Array(sz); ceCol = new Int32Array(sz); ceRule = new Int32Array(sz);
  ceState = new Uint8Array(sz); ceSkip = new Uint8Array(sz);
  ceVal = new Array(sz).fill(null); ceOld = new Array(sz).fill(null);
  nCell = 0;
  for (let s = 0; s < shs.length; s++) {
    const S = shs[s];
    if (S.br2 < 0) continue;
    for (let c = S.bc1; c <= S.bc2; c++) {
      const col = S.ids[c - S.bc1];
      if (col === null) continue;
      for (let i = 0; i < col.length; i++) {
        if (col[i] < 0) {
          nCell++;
          ceRule[nCell] = -col[i];
          col[i] = nCell;
          ceSheet[nCell] = s; ceRow[nCell] = S.br1 + i; ceCol[nCell] = c;
          cellsOfRule[ceRule[nCell]]++;
        }
      }
    }
  }
  capS = []; capR = []; capC = []; capU = [];
  return cellsOfRule;
}

// Id of the rule cell at sheet s, row r, column c (0 = none).
function engCellId(s, r, c) {
  const S = shs[s];
  if (S === undefined || r < S.br1 || r > S.br2 || c < S.bc1 || c > S.bc2) return 0;
  const col = S.ids[c - S.bc1];
  if (col === null) return 0;
  const id = col[r - S.br1];
  return id > 0 ? id : 0;
}

function engCellInfo(id) {
  return { rule: ceRule[id], val: ceVal[id], old: ceOld[id], state: ceState[id], skip: ceSkip[id] };
}

function cellName(id) {
  return shs[ceSheet[id]].sheet.name + '!' + addr(ceRow[id], ceCol[id]);
}

function growTokens() {
  const n = tkCap === 0 ? 4096 : tkCap * 2;
  const g = (a) => { const b = new Int32Array(n + 1); if (a !== null) b.set(a); return b; };
  tkOp = g(tkOp); tkA = g(tkA); tkB = g(tkB); tkC = g(tkC); tkD = g(tkD); tkE = g(tkE); tkF = g(tkF);
  tkCap = n;
}

function newTok(op) {
  nTk++;
  if (nTk > tkCap) growTokens();
  tkOp[nTk] = op; tkA[nTk] = 0; tkB[nTk] = 0; tkC[nTk] = 0; tkD[nTk] = 0; tkE[nTk] = 0; tkF[nTk] = 0;
  tkV[nTk] = null;
  return nTk;
}

function growStack() {
  const n = stCap === 0 ? 256 : stCap * 2;
  const g = (a, T) => { const b = new T(n + 1); if (a !== null) b.set(a); return b; };
  stK = g(stK, Uint8Array); stS = g(stS, Int32Array); stR1 = g(stR1, Int32Array); stC1 = g(stC1, Int32Array);
  stR2 = g(stR2, Int32Array); stC2 = g(stC2, Int32Array);
  stCap = n;
}

//==============================================================================
//  Evaluation of all cells
//==============================================================================
// Evaluates every rule (only the sheets whose name starts with onlyPrefix, if given).
function engEvaluate(onlyPrefix, onProgress) {
  onlyPrefix = onlyPrefix || '';
  wantSh = shs.map((S) => onlyPrefix.length === 0 || S.sheet.name.slice(0, onlyPrefix.length) === onlyPrefix);
  for (const S of shs) {
    const sh = S.sheet;
    S.nRow = Math.max(sh.maxRow, S.br2);
    S.nCol = Math.max(sh.maxCol, S.bc2);
    if (sh.virt !== null) {
      if (sh.virt.r2 > S.nRow) S.nRow = sh.virt.r2;
      if (sh.virt.c2 > S.nCol) S.nCol = sh.virt.c2;
    }
    if (S.cdone !== null) S.cdone.fill(0);
  }
  ceState.fill(0);
  ceVal.fill(null);
  ceOld.fill(null);
  aggCache = new Map();
  ixByKey = new Map();
  ix = [null];
  tcN = new Int32Array(nTk + 2); tcS = new Int32Array(nTk + 2); tcC = new Int32Array(nTk + 2);
  tcR1 = new Int32Array(nTk + 2); tcR2 = new Int32Array(nTk + 2);
  sp = 0;
  if (stCap === 0) growStack();
  for (let id = 1; id <= nCell; id++) {
    if (wantSh[ceSheet[id]]) {
      ceOld[id] = shs[ceSheet[id]].sheet.get(ceRow[id], ceCol[id]);
      if (ceSkip[id]) { ceVal[id] = ceOld[id]; ceState[id] = 2; }
    }
  }
  let t = Date.now(), lastS = -1, nDone = 0;
  engTimeNote = '';
  for (let id = 1; id <= nCell; id++) {
    if (ceState[id] !== 2 && wantSh[ceSheet[id]]) {
      if (ceSheet[id] !== lastS) {
        if (lastS >= 0) {
          engTimeNote += (engTimeNote ? '; ' : '') + shs[lastS].sheet.name + ' ' + ((Date.now() - t) / 1000).toFixed(2) + 's';
          t = Date.now();
        }
        lastS = ceSheet[id];
        if (onProgress) onProgress(shs[lastS].sheet.name, nDone);
      }
      evalCell(id);
      nDone++;
    }
  }
  if (lastS >= 0) engTimeNote += (engTimeNote ? '; ' : '') + shs[lastS].sheet.name + ' ' + ((Date.now() - t) / 1000).toFixed(2) + 's';
  return nDone;
}

function engTimeNoteText() { return engTimeNote; }

// Cells whose result differs from the value in the file (numbers within 1E-9 relative).
// Returns {count, list: [{sheet, r, c, old, val}]}.
function engCompare(maxList) {
  let n = 0;
  const list = [];
  for (let id = 1; id <= nCell; id++) {
    if (ceState[id] === 2 && !ceSkip[id] && wantSh[ceSheet[id]]) {
      if (!sameValue(ceOld[id], ceVal[id])) {
        n++;
        if (list.length < maxList) {
          list.push({ sheet: shs[ceSheet[id]].sheet.name, r: ceRow[id], c: ceCol[id], old: ceOld[id], val: ceVal[id] });
        }
      }
    }
  }
  return { count: n, list };
}

// Writes the results into the sheets of the model. mode 0: only cells whose result changed
// (numbers by more than 1E-9); 1: every evaluated cell. Returns the number of cells written.
function engWrite(mode) {
  let n = 0;
  for (let id = 1; id <= nCell; id++) {
    if (ceState[id] !== 2 || ceSkip[id] || !wantSh[ceSheet[id]]) continue;
    const o = ceOld[id], v = ceVal[id];
    let chg;
    if (mode > 0) chg = true;
    else if (typeof o === 'number' && typeof v === 'number') chg = o !== v && !sameValue(o, v);
    else chg = !sameValue(o, v);
    if (chg) {
      shs[ceSheet[id]].sheet.set(ceRow[id], ceCol[id], v);
      ceOld[id] = v;
      n++;
    }
  }
  return n;
}

// Marks a rule cell to be left alone (it keeps the value it has).
function engSkipCell(id) { ceSkip[id] = 1; }

//==============================================================================
//  Compiler: formula text (R1C1) -> tokens in reverse Polish notation.
//  IF and IFERROR become jumps, so only the branch Excel evaluates is evaluated.
//==============================================================================
function cerr(msg) { return new CompileError(msg); }

function compileRule(ru) {
  ruTk0[ru] = nTk + 1;
  try {
    const f = ruText[ru];
    if (f.charAt(0) !== '=') throw cerr('ไม่ใช่สูตร');
    psS = f.slice(1); psI = 0; psN = psS.length; psSheet = ruSheet[ru];
    pCompare();
    skipWs();
    if (psI < psN) throw cerr("อ่านไม่ได้ตรง '" + psS.charAt(psI) + "'");
    ruTk1[ru] = nTk;
    peephole(ru);
  } catch (e) {
    nTk = ruTk0[ru] - 1;
    if (e instanceof CompileError) throw e;
    throw cerr(String(e && e.message || e));
  }
}

// Fuses frequent token sequences of a rule (its tokens are the last ones):
//   cell, constant, comparison, IF-jump  -> T_IFCK
//   constant, arithmetic                 -> T_ARK
//   cell, arithmetic                     -> T_ARR
function peephole(ru) {
  const t0 = ruTk0[ru], n = ruTk1[ru] - t0 + 1;
  if (n < 2) return;
  const op = new Int32Array(n + 1), a = new Int32Array(n + 1), b = new Int32Array(n + 1), c = new Int32Array(n + 1);
  const d = new Int32Array(n + 1), e = new Int32Array(n + 1), f = new Int32Array(n + 1), v = new Array(n + 1);
  const tgt = new Uint8Array(n + 2), map = new Int32Array(n + 2);
  for (let k = 1; k <= n; k++) {
    const i = t0 + k - 1;
    op[k] = tkOp[i]; a[k] = tkA[i]; b[k] = tkB[i]; c[k] = tkC[i]; d[k] = tkD[i]; e[k] = tkE[i]; f[k] = tkF[i];
    v[k] = tkV[i];
  }
  for (let k = 1; k <= n; k++) {
    if (op[k] === T_JMP || op[k] === T_JNE) tgt[a[k] - t0 + 1] = 1;
    else if (op[k] === T_JMPF) { tgt[a[k] - t0 + 1] = 1; tgt[b[k] - t0 + 1] = 1; }
  }
  nTk = t0 - 1;
  let k = 1, nt;
  while (k <= n) {
    if (k + 3 <= n && op[k] === T_REF && isCellTok(f[k], b[k], c[k], d[k], e[k]) && op[k + 1] === T_NUM &&
        op[k + 2] >= T_EQ && op[k + 2] <= T_GE && op[k + 3] === T_JMPF && !(tgt[k + 1] || tgt[k + 2] || tgt[k + 3])) {
      nt = newTok(T_IFCK);
      tkA[nt] = a[k]; tkB[nt] = b[k]; tkC[nt] = c[k]; tkF[nt] = (f[k] & 3) + 256 * op[k + 2];
      tkV[nt] = v[k + 1]; tkD[nt] = a[k + 3]; tkE[nt] = b[k + 3];
      map[k] = map[k + 1] = map[k + 2] = map[k + 3] = nt;
      k += 4;
      continue;
    }
    if (k + 1 <= n && op[k + 1] >= T_ADD && op[k + 1] <= T_POW && !tgt[k + 1]) {
      if (op[k] === T_NUM && typeof v[k] === 'number') {
        nt = newTok(T_ARK);
        tkF[nt] = op[k + 1]; tkV[nt] = v[k];
        map[k] = map[k + 1] = nt;
        k += 2;
        continue;
      }
      if (op[k] === T_REF && isCellTok(f[k], b[k], c[k], d[k], e[k])) {
        nt = newTok(T_ARR);
        tkA[nt] = a[k]; tkB[nt] = b[k]; tkC[nt] = c[k]; tkF[nt] = (f[k] & 3) + 256 * op[k + 1];
        map[k] = map[k + 1] = nt;
        k += 2;
        continue;
      }
    }
    nt = newTok(op[k]);
    tkA[nt] = a[k]; tkB[nt] = b[k]; tkC[nt] = c[k]; tkD[nt] = d[k]; tkE[nt] = e[k]; tkF[nt] = f[k]; tkV[nt] = v[k];
    map[k] = nt;
    k++;
  }
  map[n + 1] = nTk + 1;
  for (let i = t0; i <= nTk; i++) {
    const o = tkOp[i];
    if (o === T_JMP || o === T_JNE) tkA[i] = map[tkA[i] - t0 + 1];
    else if (o === T_JMPF) { tkA[i] = map[tkA[i] - t0 + 1]; tkB[i] = map[tkB[i] - t0 + 1]; }
    else if (o === T_IFCK) { tkD[i] = map[tkD[i] - t0 + 1]; tkE[i] = map[tkE[i] - t0 + 1]; }
  }
  ruTk1[ru] = nTk;
}

// One cell (not a range, not whole columns or rows).
function isCellTok(f, r1, c1, r2, c2) {
  if ((f & (RF_COLS | RF_ROWS)) !== 0) return false;
  if (r1 !== r2 || c1 !== c2) return false;
  if (((f & RF_R1) !== 0) !== ((f & RF_R2) !== 0)) return false;
  if (((f & RF_C1) !== 0) !== ((f & RF_C2) !== 0)) return false;
  return true;
}

function peek1() { return psI < psN ? psS.charAt(psI) : ''; }
function peek2() { return psS.substr(psI, 2); }

function skipWs() {
  while (psI < psN) {
    const ch = psS.charAt(psI);
    if (ch === ' ' || ch === '\r' || ch === '\n' || ch === '\t') psI++;
    else break;
  }
}

function expect(ch) {
  skipWs();
  if (peek1() !== ch) throw cerr("ต้องมี '" + ch + "'");
  psI++;
}

function pCompare() {
  pConcat();
  for (;;) {
    skipWs();
    let op;
    const p2 = peek2();
    if (p2 === '<>') { op = T_NE; psI += 2; }
    else if (p2 === '<=') { op = T_LE; psI += 2; }
    else if (p2 === '>=') { op = T_GE; psI += 2; }
    else {
      const p1 = peek1();
      if (p1 === '=') op = T_EQ;
      else if (p1 === '<') op = T_LT;
      else if (p1 === '>') op = T_GT;
      else break;
      psI++;
    }
    pConcat();
    newTok(op);
  }
}

function pConcat() {
  pAdd();
  for (;;) {
    skipWs();
    if (peek1() !== '&') break;
    psI++;
    pAdd();
    newTok(T_CAT);
  }
}

function pAdd() {
  pMul();
  for (;;) {
    skipWs();
    const ch = peek1();
    let op;
    if (ch === '+') op = T_ADD;
    else if (ch === '-') op = T_SUB;
    else break;
    psI++;
    pMul();
    newTok(op);
  }
}

function pMul() {
  pPow();
  for (;;) {
    skipWs();
    const ch = peek1();
    let op;
    if (ch === '*') op = T_MUL;
    else if (ch === '/') op = T_DIV;
    else break;
    psI++;
    pPow();
    newTok(op);
  }
}

function pPow() {
  pUnary();
  for (;;) {
    skipWs();
    if (peek1() !== '^') break;
    psI++;
    pUnary();
    newTok(T_POW);
  }
}

// In Excel a leading minus binds tighter than ^ :  -2^2 = 4
function pUnary() {
  skipWs();
  const ch = peek1();
  if (ch === '-') { psI++; pUnary(); newTok(T_NEG); }
  else if (ch === '+') { psI++; pUnary(); newTok(T_PLUS); }
  else pPercent();
}

function pPercent() {
  pPrimary();
  for (;;) {
    skipWs();
    if (peek1() !== '%') break;
    psI++;
    newTok(T_PCT);
  }
}

function pPrimary() {
  skipWs();
  if (psI >= psN) throw cerr('สูตรจบก่อนกำหนด');
  const ch = psS.charAt(psI);
  let t;
  if (ch === '(') {
    psI++;
    pCompare();
    expect(')');
  } else if (ch === '"') {
    t = newTok(T_NUM); tkV[t] = readString();
  } else if (ch === '#') {
    t = newTok(T_NUM); tkV[t] = readError();
  } else if ((ch >= '0' && ch <= '9') || ch === '.') {
    t = newTok(T_NUM); tkV[t] = readNumber();
  } else if (ch === "'") {
    if (!tryRef(readQuotedSheet())) throw cerr('ต้องมีการอ้างอิงเซลล์หลังชื่อชีต');
  } else if (ch === '{') {
    throw cerr('ไม่รองรับค่าคงที่แบบ array {…}');
  } else if (ch === '[') {
    throw cerr('ไม่รองรับการอ้างอิงไปไฟล์อื่น');
  } else {
    if ((ch === 'R' || ch === 'C') && tryRef(psSheet)) return;
    const st = psI;
    while (psI < psN && isNameChar(psS.charAt(psI))) psI++;
    if (psI === st) throw cerr("อ่านไม่ได้ตรง '" + ch + "'");
    const nm = psS.slice(st, psI);
    if (peek1() === '!') {
      psI++;
      if (!tryRef(sheetOf(nm))) throw cerr('ต้องมีการอ้างอิงเซลล์หลังชื่อชีต');
    } else if (peek1() === '(') {
      psI++;
      pFunc(nm.toUpperCase());
    } else if (nm.toUpperCase() === 'TRUE' || nm.toUpperCase() === 'FALSE') {
      t = newTok(T_NUM); tkV[t] = nm.toUpperCase() === 'TRUE';
    } else {
      throw cerr('ไม่รองรับชื่อ (Name) "' + nm + '"');
    }
  }
}

function isNameChar(ch) {
  if (ch === '') return false;
  const c = ch.charCodeAt(0);
  return (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || ch === '_' || ch === '.' ||
    ch === '\\' || c > 127;
}

// R1C1 reference at psI (cell, range, whole columns or whole rows). Restores psI when there is none.
function tryRef(s) {
  const save = psI;
  const p1 = readRC();
  if (p1 === null) { psI = save; return false; }
  let ch = peek1();
  if (ch === '(' || ch === '!' || isNameChar(ch)) { psI = save; return false; }
  let p2 = p1;
  if (ch === ':') {
    psI++;
    p2 = readRC();
    if (p2 === null) throw cerr("ต้องมีการอ้างอิงหลัง ':'");
    if (p2.hasR !== p1.hasR || p2.hasC !== p1.hasC) throw cerr('ไม่รองรับช่วงเซลล์แบบนี้');
    ch = peek1();
    if (ch === '(' || ch === '!' || isNameChar(ch)) throw cerr("อ่านไม่ได้ตรง '" + ch + "'");
  }
  const t = newTok(T_REF);
  let f = 0;
  tkA[t] = s;
  if (p1.hasR && p1.hasC) {
    tkB[t] = p1.r; tkC[t] = p1.c; tkD[t] = p2.r; tkE[t] = p2.c;
    if (p1.rRel) f |= RF_R1;
    if (p1.cRel) f |= RF_C1;
    if (p2.rRel) f |= RF_R2;
    if (p2.cRel) f |= RF_C2;
  } else if (p1.hasC) {
    tkB[t] = 1; tkD[t] = MAX_ROW; tkC[t] = p1.c; tkE[t] = p2.c;
    f = RF_COLS;
    if (p1.cRel) f |= RF_C1;
    if (p2.cRel) f |= RF_C2;
  } else {
    tkC[t] = 1; tkE[t] = MAX_COL; tkB[t] = p1.r; tkD[t] = p2.r;
    f = RF_ROWS;
    if (p1.rRel) f |= RF_R1;
    if (p2.rRel) f |= RF_R2;
  }
  tkF[t] = f;
  return true;
}

function readRC() {
  let hasR = false, hasC = false, r = 0, c = 0, rRel = false, cRel = false;
  if (peek1() === 'R') {
    hasR = true; psI++;
    const o = readOffset();
    if (o === null) return null;
    r = o.v; rRel = o.rel;
  }
  if (peek1() === 'C') {
    hasC = true; psI++;
    const o = readOffset();
    if (o === null) return null;
    c = o.v; cRel = o.rel;
  }
  if (!hasR && !hasC) return null;
  return { hasR, hasC, r, c, rRel, cRel };
}

// After R or C:  [n] = relative, n = absolute, nothing = relative 0.
function readOffset() {
  if (peek1() === '[') {
    psI++;
    const st = psI;
    if (peek1() === '-' || peek1() === '+') psI++;
    while (psI < psN && psS.charAt(psI) >= '0' && psS.charAt(psI) <= '9') psI++;
    if (peek1() !== ']' || psI === st) return null;
    const v = parseInt(psS.slice(st, psI), 10);
    if (!(v === v)) return null;
    psI++;
    return { v, rel: true };
  }
  const ch = peek1();
  if (ch >= '0' && ch <= '9' && ch !== '') {
    const st = psI;
    while (psI < psN && psS.charAt(psI) >= '0' && psS.charAt(psI) <= '9') psI++;
    return { v: parseInt(psS.slice(st, psI), 10), rel: false };
  }
  return { v: 0, rel: true };
}

function readQuotedSheet() {
  let nm = '';
  psI++;
  for (;;) {
    if (psI >= psN) throw cerr('เครื่องหมาย \' ไม่ครบคู่');
    const ch = psS.charAt(psI);
    if (ch === "'") {
      if (psS.charAt(psI + 1) === "'") { nm += "'"; psI += 2; } else { psI++; break; }
    } else {
      nm += ch; psI++;
    }
  }
  if (peek1() !== '!') throw cerr("ต้องมี '!' หลังชื่อชีต");
  psI++;
  if (nm.indexOf('[') >= 0) throw cerr('ไม่รองรับการอ้างอิงไปไฟล์อื่น');
  return sheetOf(nm);
}

function sheetOf(nm) {
  const sh = engWb.sheet(nm);
  if (sh === null) throw cerr("ไม่พบชีต '" + nm + "'");
  return sh.index;
}

function readString() {
  let s = '';
  psI++;
  for (;;) {
    if (psI >= psN) throw cerr('ข้อความในสูตรไม่มีเครื่องหมาย " ปิด');
    const ch = psS.charAt(psI);
    if (ch === '"') {
      if (psS.charAt(psI + 1) === '"') { s += '"'; psI += 2; } else { psI++; break; }
    } else {
      s += ch; psI++;
    }
  }
  return s;
}

function readNumber() {
  const st = psI;
  while (psI < psN && /[0-9.]/.test(psS.charAt(psI))) psI++;
  if (psI < psN && (psS.charAt(psI) === 'E' || psS.charAt(psI) === 'e') && /[0-9+-]/.test(psS.charAt(psI + 1))) {
    psI += 2;
    while (psI < psN && psS.charAt(psI) >= '0' && psS.charAt(psI) <= '9') psI++;
  }
  return vbVal(psS.slice(st, psI));
}

// VBA Val: the leading number of the text ("1.5.2" -> 1.5).
function vbVal(s) {
  const m = /^[0-9]*\.?[0-9]*(?:[eE][+-]?[0-9]+)?/.exec(s);
  const x = Number(m[0]);
  if (x === x) return x;
  const m2 = /^[0-9]*\.?[0-9]*/.exec(s);
  const y = Number(m2[0]);
  return y === y ? y : 0;
}

function readError() {
  for (const len of [7, 6, 5, 4]) {
    const e = errLiteral(psS.substr(psI, len));
    if (e !== null) { psI += e.text.length; return e; }
  }
  throw cerr('ค่า error ที่ไม่รู้จัก');
}

function pArg() {
  skipWs();
  if (peek1() === ',' || peek1() === ')') newTok(T_MISS);
  else pCompare();
}

function pFunc(nm) {
  let fid = 0, argc = 0, t, j1, j2;
  if (nm.slice(0, 6) === '_XLFN.') nm = nm.slice(6);
  switch (nm) {
    case 'IF':
      pArg();
      expect(',');
      j1 = newTok(T_JMPF);
      pArg();
      j2 = newTok(T_JMP);
      tkA[j1] = nTk + 1;                          // FALSE: the third argument
      skipWs();
      if (peek1() === ',') {
        psI++;
        pArg();
      } else {
        t = newTok(T_NUM); tkV[t] = false;        // IF(test, x) gives FALSE
      }
      expect(')');
      tkA[j2] = nTk + 1;
      tkB[j1] = nTk + 1;                          // error in the test: skip both
      return;
    case 'IFERROR':
      pArg();
      expect(',');
      j1 = newTok(T_JNE);
      pArg();
      expect(')');
      tkA[j1] = nTk + 1;
      return;
    case 'TRUE': case 'FALSE':
      skipWs();
      expect(')');
      t = newTok(T_NUM); tkV[t] = nm === 'TRUE';
      return;
    case 'SUM': fid = F_SUM; break;
    case 'COUNT': fid = F_COUNT; break;
    case 'SUMIF': fid = F_SUMIF; break;
    case 'SUMIFS': fid = F_SUMIFS; break;
    case 'COUNTIF': fid = F_COUNTIF; break;
    case 'COUNTIFS': fid = F_COUNTIFS; break;
    case 'VLOOKUP': fid = F_VLOOKUP; break;
    case 'INT': fid = F_INT; break;
    case 'MID': fid = F_MID; break;
    case 'NA': fid = F_NA; break;
    default: throw cerr('ไม่รองรับฟังก์ชัน ' + nm);
  }
  skipWs();
  if (peek1() === ')') {
    psI++;
  } else {
    for (;;) {
      pArg();
      argc++;
      skipWs();
      if (peek1() === ',') psI++; else break;
    }
    expect(')');
  }
  let ok = true;
  switch (fid) {
    case F_SUM: case F_COUNT: ok = argc >= 1; break;
    case F_SUMIF: ok = argc >= 2 && argc <= 3; break;
    case F_SUMIFS: ok = argc >= 3 && argc % 2 === 1; break;
    case F_COUNTIF: ok = argc === 2; break;
    case F_COUNTIFS: ok = argc >= 2 && argc % 2 === 0; break;
    case F_VLOOKUP: ok = argc >= 3 && argc <= 4; break;
    case F_INT: ok = argc === 1; break;
    case F_MID: ok = argc === 3; break;
    case F_NA: ok = argc === 0; break;
  }
  if (!ok) throw cerr('จำนวนอาร์กิวเมนต์ของ ' + nm + ' ไม่ถูกต้อง');
  t = newTok(T_FUNC); tkA[t] = fid; tkB[t] = argc;
}

//==============================================================================
//  Evaluation
//==============================================================================
function evalCell(id) {
  if (ceState[id] === 1) throw new UserError('สูตรอ้างอิงวนกลับมาที่ตัวเอง (circular reference) ที่ ' + cellName(id));
  ceState[id] = 1;
  const sr = curRow, sc = curCol;
  curRow = ceRow[id]; curCol = ceCol[id];
  ceVal[id] = runRule(ceRule[id], curRow, curCol);
  curRow = sr; curCol = sc;
  ceState[id] = 2;
}

// Value of a cell; a formula cell that is not done yet is evaluated first.
function getV(s, r, c) {
  const S = shs[s];
  if (r >= S.br1 && r <= S.br2 && c >= S.bc1 && c <= S.bc2) {
    const col = S.ids[c - S.bc1];
    if (col !== null) {
      const id = col[r - S.br1];
      if (id > 0) {
        if (ceState[id] !== 2) evalCell(id);
        return ceVal[id];
      }
    }
  }
  return S.sheet.get(r, c);
}

// Evaluates the formula cells inside a range (before a function reads the range).
function ensureRange(s, r1, c1, r2, c2) {
  const S = shs[s];
  if (S.br2 < 0) return;
  if (r1 < S.br1) r1 = S.br1;
  if (r2 > S.br2) r2 = S.br2;
  if (c1 < S.bc1) c1 = S.bc1;
  if (c2 > S.bc2) c2 = S.bc2;
  if (r1 > r2 || c1 > c2) return;
  const whole = r1 === S.br1 && r2 === S.br2;
  for (let c = c1; c <= c2; c++) {
    if (S.cdone[c - S.bc1]) continue;
    const col = S.ids[c - S.bc1];
    if (col !== null) {
      for (let r = r1; r <= r2; r++) {
        const id = col[r - S.br1];
        if (id > 0 && ceState[id] !== 2) evalCell(id);
      }
    }
    if (whole) S.cdone[c - S.bc1] = 1;
  }
}

function runRule(ru, r0, c0) {
  const base = sp;
  let pc = ruTk0[ru];
  const pEnd = ruTk1[ru];
  let a, b, ok, isT, f, x;
  while (pc <= pEnd) {
    const op = tkOp[pc];
    switch (op) {
      case T_REF:
        sp++;
        if (sp > stCap) growStack();
        f = tkF[pc];
        stK[sp] = 1; stS[sp] = tkA[pc];
        stR1[sp] = (f & RF_R1) ? r0 + tkB[pc] : tkB[pc];
        stC1[sp] = (f & RF_C1) ? c0 + tkC[pc] : tkC[pc];
        stR2[sp] = (f & RF_R2) ? r0 + tkD[pc] : tkD[pc];
        stC2[sp] = (f & RF_C2) ? c0 + tkE[pc] : tkE[pc];
        if (f > 0) fixRef(sp);
        break;
      case T_NUM:
        sp++;
        if (sp > stCap) growStack();
        stK[sp] = 0; stV[sp] = tkV[pc];
        break;
      case T_IFCK: {
        a = tokCell(pc, r0, c0);
        const kv = tkV[pc], cmp = tkF[pc] >> 8;
        if (a === null && typeof kv === 'string' && cmp === T_EQ) {
          ok = true; isT = kv.length === 0;           // blank = "" only
        } else if (a instanceof XErr) {
          ok = false;
        } else if (kv instanceof XErr) {
          ok = false; a = kv;
        } else {
          ok = true;
          x = xCompare(a, kv);
          switch (cmp) {
            case T_EQ: isT = x === 0; break;
            case T_NE: isT = x !== 0; break;
            case T_LT: isT = x < 0; break;
            case T_LE: isT = x <= 0; break;
            case T_GT: isT = x > 0; break;
            default: isT = x >= 0; break;
          }
        }
        if (ok) {
          if (!isT) pc = tkD[pc] - 1;
        } else {
          sp++;
          if (sp > stCap) growStack();
          stK[sp] = 0; stV[sp] = a;
          pc = tkE[pc] - 1;
        }
        break;
      }
      case T_ARK:
        a = stK[sp] === 0 ? stV[sp] : scalar(sp);
        stK[sp] = 0; stV[sp] = arithVals(tkF[pc], a, tkV[pc], pc === pEnd);
        break;
      case T_ARR:
        b = tokCell(pc, r0, c0);
        a = stK[sp] === 0 ? stV[sp] : scalar(sp);
        stK[sp] = 0; stV[sp] = arithVals(tkF[pc] >> 8, a, b, pc === pEnd);
        break;
      case T_DIV: case T_MUL: case T_ADD: case T_SUB: case T_POW:  // Excel sets a last + or - that cancels out to 0
        b = scalar(sp); sp--; a = scalar(sp);
        stK[sp] = 0; stV[sp] = arithVals(op, a, b, pc === pEnd);
        break;
      case T_JMPF:
        a = scalar(sp);
        ok = true;
        if (typeof a === 'boolean') isT = a;
        else if (typeof a === 'number') isT = a !== 0;
        else if (a === null) isT = false;
        else if (typeof a === 'string') {
          if (strCompText(a, 'TRUE') === 0) isT = true;
          else if (strCompText(a, 'FALSE') === 0) isT = false;
          else { ok = false; a = ERR_VALUE; }
        } else ok = false;
        if (ok) {
          sp--;
          if (!isT) pc = tkA[pc] - 1;
        } else {
          stK[sp] = 0; stV[sp] = a;
          pc = tkB[pc] - 1;
        }
        break;
      case T_JMP:
        pc = tkA[pc] - 1;
        break;
      case T_JNE:
        a = scalar(sp);
        if (a instanceof XErr) {
          sp--;
        } else {
          stK[sp] = 0; stV[sp] = a;
          pc = tkA[pc] - 1;
        }
        break;
      case T_EQ: case T_NE: case T_LT: case T_LE: case T_GT: case T_GE:
        doCompare(op);
        break;
      case T_FUNC:
        curTok = pc;
        callFunc(tkA[pc], tkB[pc]);
        break;
      case T_MISS:
        sp++;
        if (sp > stCap) growStack();
        stK[sp] = 2; stV[sp] = null;
        break;
      case T_NEG: case T_PLUS: case T_PCT:
        doUnary(op);
        break;
      case T_CAT:
        doConcat();
        break;
    }
    pc++;
  }
  a = scalar(sp);
  if (a === null) a = 0;                 // a formula that points at a blank cell shows 0
  sp = base;
  return a;
}

function fixRef(i) {
  let t;
  if (stR1[i] > stR2[i]) { t = stR1[i]; stR1[i] = stR2[i]; stR2[i] = t; }
  if (stC1[i] > stC2[i]) { t = stC1[i]; stC1[i] = stC2[i]; stC2[i] = t; }
  if (stR1[i] < 1 || stC1[i] < 1 || stR2[i] > MAX_ROW || stC2[i] > MAX_COL) {
    stK[i] = 0; stV[i] = ERR_REF;
  }
}

// Value of stack entry i. A range where one value is expected gives the cell in the
// row (or column) of the formula (implicit intersection), else #VALUE!.
function scalar(i) {
  const k = stK[i];
  if (k === 0) return stV[i];
  if (k === 1) {
    const r1 = stR1[i], c1 = stC1[i], r2 = stR2[i], c2 = stC2[i];
    if (r1 === r2 && c1 === c2) return getV(stS[i], r1, c1);
    if (c1 === c2 && curRow >= r1 && curRow <= r2) return getV(stS[i], curRow, c1);
    if (r1 === r2 && curCol >= c1 && curCol <= c2) return getV(stS[i], r1, curCol);
    return ERR_VALUE;
  }
  return null;
}

// a <op> b with Excel's conversions; isLast: the last step of the formula.
function arithVals(op, a, b, isLast) {
  let x, y;
  if (typeof a === 'number') x = a; else { x = toNum(a); if (typeof x !== 'number') return x; }
  if (typeof b === 'number') y = b; else { y = toNum(b); if (typeof y !== 'number') return y; }
  let r;
  switch (op) {
    case T_ADD: r = x + y; if (isLast) r = cancelSum(r, x, y); break;
    case T_SUB: r = x - y; if (isLast) r = cancelSum(r, x, y); break;
    case T_MUL: r = x * y; break;
    case T_DIV:
      if (y === 0) return ERR_DIV0;
      r = x / y;
      break;
    default:
      return xPow(x, y);
  }
  if (r !== r || r === Infinity || r === -Infinity) return ERR_NUM;
  return r;
}

// Value of the single cell of a fused token (#REF! off the sheet).
function tokCell(pc, r0, c0) {
  const r = (tkF[pc] & RF_R1) ? r0 + tkB[pc] : tkB[pc];
  const c = (tkF[pc] & RF_C1) ? c0 + tkC[pc] : tkC[pc];
  if (r < 1 || c < 1 || r > MAX_ROW || c > MAX_COL) return ERR_REF;
  return getV(tkA[pc], r, c);
}

// Excel returns 0 when + or - cancels out to within 2^-49 of the larger operand
// (measured: 1-(1-2^-50) gives 0, 1-(1-2^-49) does not).
function cancelSum(r, x, y) {
  if (r === 0) return r;
  let m = Math.abs(x);
  if (Math.abs(y) > m) m = Math.abs(y);
  if (Math.abs(r) >= m * 0.000000000000002) return r;
  let p = Math.pow(2, Math.floor(Math.log(m) / Math.log(2)));
  if (p > m) p = p / 2;
  if (p * 2 <= m) p = p * 2;
  if (Math.abs(r) < p * 1.77635683940025E-15) return 0;
  return r;
}

function addCancel(acc, x) {
  const r = acc + x;
  if (r !== 0 && Math.abs(r) < Math.abs(x) * 0.000000000000002) return cancelSum(r, acc, x);
  return r;
}

function xPow(x, y) {
  if (x === 0 && y === 0) return ERR_NUM;
  if (x === 0 && y < 0) return ERR_DIV0;
  if (x < 0 && y !== Math.floor(y)) {
    const k = 1 / y;                             // odd roots of negative numbers work in Excel
    const rk = vbRound(k);
    if (Math.abs(k - rk) < 0.000000001 && Math.abs(rk) < 2147483648 && rk % 2 !== 0) {
      const r = -Math.pow(-x, y);
      return isFinite(r) ? r : ERR_NUM;
    }
    return ERR_NUM;
  }
  const r = Math.pow(x, y);
  return isFinite(r) ? r : ERR_NUM;
}

function doCompare(op) {
  const b = scalar(sp);
  sp--;
  const a = scalar(sp);
  stK[sp] = 0;
  if (a instanceof XErr) { stV[sp] = a; return; }
  if (b instanceof XErr) { stV[sp] = b; return; }
  const c = xCompare(a, b);
  switch (op) {
    case T_EQ: stV[sp] = c === 0; break;
    case T_NE: stV[sp] = c !== 0; break;
    case T_LT: stV[sp] = c < 0; break;
    case T_LE: stV[sp] = c <= 0; break;
    case T_GT: stV[sp] = c > 0; break;
    default: stV[sp] = c >= 0; break;
  }
}

function doUnary(op) {
  const a = scalar(sp);
  stK[sp] = 0;
  if (op === T_PLUS) { stV[sp] = a; return; }
  const x = toNum(a);
  if (typeof x !== 'number') stV[sp] = x;
  else if (op === T_NEG) stV[sp] = -x;
  else stV[sp] = x / 100;
}

function doConcat() {
  const b = scalar(sp);
  sp--;
  const a = scalar(sp);
  stK[sp] = 0;
  const s1 = toText(a);
  if (s1 instanceof XErr) { stV[sp] = s1; return; }
  const s2 = toText(b);
  if (s2 instanceof XErr) { stV[sp] = s2; return; }
  stV[sp] = s1 + s2;
}

function callFunc(fid, argc) {
  const base = sp - argc + 1;
  let res;
  switch (fid) {
    case F_SUM: res = fnSum(base, argc); break;
    case F_COUNT: res = fnCount(base, argc); break;
    case F_SUMIF: res = fnSumIf(base, argc); break;
    case F_SUMIFS: res = fnIfs(true, base, argc); break;
    case F_COUNTIF: case F_COUNTIFS: res = fnIfs(false, base, argc); break;
    case F_VLOOKUP: res = fnVlookup(base, argc); break;
    case F_INT: res = fnInt(base); break;
    case F_MID: res = fnMid(base); break;
    case F_NA: res = ERR_NA; break;
  }
  sp = base;
  if (sp > stCap) growStack();
  stK[sp] = 0; stV[sp] = res;
}

//==============================================================================
//  Functions
//==============================================================================
function fnSum(base, argc) {
  let acc = 0;
  for (let i = base; i < base + argc; i++) {
    const k = stK[i];
    if (k === 1) {                                 // range: numbers only, first error wins
      const s = stS[i];
      ensureRange(s, stR1[i], stC1[i], stR2[i], stC2[i]);
      const S = shs[s];
      const r2 = Math.min(stR2[i], S.nRow), c2 = Math.min(stC2[i], S.nCol);
      for (let r = stR1[i]; r <= r2; r++) {
        for (let c = stC1[i]; c <= c2; c++) {
          const v = getV(s, r, c);
          if (typeof v === 'number') acc = addCancel(acc, v);
          else if (v instanceof XErr) return v;
        }
      }
    } else if (k === 0) {                          // value: TRUE = 1, number text counts
      const v = stV[i];
      if (typeof v === 'number') acc = addCancel(acc, v);
      else if (typeof v === 'boolean') { if (v) acc = addCancel(acc, 1); }
      else if (typeof v === 'string') {
        const x = toNum(v);
        if (typeof x !== 'number') return x;
        acc = addCancel(acc, x);
      } else if (v instanceof XErr) return v;
    }
  }
  return acc;
}

function fnCount(base, argc) {
  let n = 0;
  for (let i = base; i < base + argc; i++) {
    const k = stK[i];
    if (k === 1) {
      const s = stS[i];
      ensureRange(s, stR1[i], stC1[i], stR2[i], stC2[i]);
      const S = shs[s];
      const r2 = Math.min(stR2[i], S.nRow), c2 = Math.min(stC2[i], S.nCol);
      for (let r = stR1[i]; r <= r2; r++) {
        for (let c = stC1[i]; c <= c2; c++) if (typeof getV(s, r, c) === 'number') n++;
      }
    } else if (k === 0) {
      const v = stV[i];
      if (typeof v === 'number' || typeof v === 'boolean') n++;
      else if (typeof v === 'string' && textToNum(v) !== null) n++;
    }
  }
  return n;
}

// SUMIF(range, criteria, [sum_range]): sum_range takes the size of range.
function fnSumIf(base, argc) {
  const sb = argc === 3 ? base + 2 : base;
  if (stK[base] !== 1 || stK[sb] !== 1) return ERR_VALUE;
  const fast = fastSumIf(base, sb);
  if (fast !== undefined) return fast;
  return critCore(true, stS[sb], stR1[sb], stC1[sb], stR2[base] - stR1[base] + 1, stC2[base] - stC1[base] + 1, base, 1);
}

// SUMIF over one column with a plain criterion (a number, or text without an operator,
// wildcard, TRUE/FALSE or error name): the same result as critCore, straight from the
// column index. Returns undefined when it does not apply.
function fastSumIf(base, sb) {
  if (stC1[base] !== stC2[base]) return undefined;
  const tok = curTok;
  const cv = scalar(base + 1);
  let key;
  if (typeof cv === 'number') {
    key = numKey(cv);
  } else if (typeof cv === 'string') {
    if (cv.length === 0) return undefined;
    const ch = cv.charAt(0);
    if (ch === '<' || ch === '>' || ch === '=' || ch === '#') return undefined;
    if (hasWild(cv)) return undefined;
    if (strCompText(cv, 'TRUE') === 0 || strCompText(cv, 'FALSE') === 0) return undefined;
    key = critCellKey(cv);
  } else if (cv === null) {
    key = 'n0';                                    // a blank criteria cell means 0
  } else {
    return undefined;
  }
  const s = stS[base], r1 = stR1[base], c = stC1[base];
  const d = rowsWithData(s, r1, stR2[base] - r1 + 1);
  if (d === 0) return 0;
  const r2 = r1 + d - 1;
  let n;
  if (tcN[tok] > 0 && tcS[tok] === s && tcC[tok] === c && tcR1[tok] === r1 && tcR2[tok] === r2) {
    n = tcN[tok];
  } else {
    n = getIndex(s, c, r1, r2, false);
    tcN[tok] = n; tcS[tok] = s; tcC[tok] = c; tcR1[tok] = r1; tcR2[tok] = r2;
  }
  const g = ix[n].keys.get(key);
  if (g === undefined) return 0;
  ensureRange(stS[sb], stR1[sb], stC1[sb], stR1[sb] + d - 1, stC1[sb]);
  const X = ix[n];
  let acc = 0;
  for (let p = X.gStart[g]; p < X.gStart[g + 1]; p++) {
    const off = X.gRows[p] - r1;
    const v = getV(stS[sb], stR1[sb] + off, stC1[sb]);
    if (typeof v === 'number') acc = addCancel(acc, v);
    else if (v instanceof XErr) return v;
  }
  return acc;
}

// SUMIFS(sum_range, range1, crit1, ...) and COUNTIF(S)(range1, crit1, ...): all ranges the same size.
function fnIfs(isSum, base, argc) {
  const pb = isSum ? base + 1 : base, nP = isSum ? (argc - 1) >> 1 : argc >> 1, first = base;
  if (stK[first] !== 1) return ERR_VALUE;
  const nR = stR2[first] - stR1[first] + 1, nC = stC2[first] - stC1[first] + 1;
  for (let k = 1; k <= nP; k++) {
    const ri = pb + 2 * (k - 1);
    if (stK[ri] !== 1) return ERR_VALUE;
    if (stR2[ri] - stR1[ri] + 1 !== nR || stC2[ri] - stC1[ri] + 1 !== nC) return ERR_VALUE;
  }
  if (isSum) {
    if (nP === 1) {
      const fast = fastSumIf(pb, base);            // SUMIFS with one criterion = SUMIF
      if (fast !== undefined) return fast;
    }
    return critCore(true, stS[base], stR1[base], stC1[base], nR, nC, pb, nP);
  }
  return critCore(false, 0, 0, 0, nR, nC, pb, nP);
}

function valKeyPart(v) {
  if (v === null) return 'E:';
  if (typeof v === 'number') return 'n:' + v;
  if (typeof v === 'string') return 's:' + v;
  if (typeof v === 'boolean') return 'b:' + v;
  return 'x:' + v.text;
}

// Sum or count over the rows where every criterion matches. The criteria pairs
// (range, criterion) are on the stack from pb on. Results are cached.
function critCore(isSum, sumS, sumR1, sumC1, nR, nC, pb, nP) {
  const tok = curTok;
  const cr = new Array(nP + 1), cv = new Array(nP + 1), rs = new Int32Array(nP + 1), rr1 = new Int32Array(nP + 1),
    rc1 = new Int32Array(nP + 1);
  for (let k = 1; k <= nP; k++) {
    const ri = pb + 2 * (k - 1);
    rs[k] = stS[ri]; rr1[k] = stR1[ri]; rc1[k] = stC1[ri];
    cv[k] = scalar(ri + 1);
  }
  for (let k = 1; k <= nP; k++) cr[k] = parseCrit(cv[k]);
  let driver = 0;
  if (nC === 1) {
    for (let k = 1; k <= nP; k++) if (isIndexable(cr[k])) { driver = k; break; }
  }
  // one criterion that an index answers: no result cache needed
  const useCache = nP > 1 || driver === 0;
  let key = '';
  if (useCache) {
    key = (isSum ? 'S' + sumS + ':' + sumR1 + ':' + sumC1 : 'C') + ':' + nR + ':' + nC;
    for (let k = 1; k <= nP; k++) key += '|' + rs[k] + ':' + rr1[k] + ':' + rc1[k] + '=' + valKeyPart(cv[k]);
    const hit = aggCache.get(key);
    if (hit !== undefined) return hit;
  }
  // rows beyond the last used row are blank in every range
  let maxOff = 0, d;
  for (let k = 1; k <= nP; k++) { d = rowsWithData(rs[k], rr1[k], nR); if (d > maxOff) maxOff = d; }
  if (isSum) { d = rowsWithData(sumS, sumR1, nR); if (d > maxOff) maxOff = d; }
  if (maxOff > 0) {
    for (let k = 1; k <= nP; k++) ensureRange(rs[k], rr1[k], rc1[k], rr1[k] + maxOff - 1, rc1[k] + nC - 1);
    if (isSum) ensureRange(sumS, sumR1, sumC1, sumR1 + maxOff - 1, sumC1 + nC - 1);
  }
  if (useCache) {
    const hit = aggCache.get(key);
    if (hit !== undefined) return hit;
  }
  let acc = 0, cnt = 0, res;
  done: {
    if (driver > 0) {
      // only the rows whose value equals the criterion (from an index of that column)
      d = rowsWithData(rs[driver], rr1[driver], nR);
      if (d > 0) {
        const r2 = rr1[driver] + d - 1;
        let n;
        if (tcN[tok] > 0 && tcS[tok] === rs[driver] && tcC[tok] === rc1[driver] && tcR1[tok] === rr1[driver] &&
            tcR2[tok] === r2) {
          n = tcN[tok];
        } else {
          n = getIndex(rs[driver], rc1[driver], rr1[driver], r2, false);
          tcN[tok] = n; tcS[tok] = rs[driver]; tcC[tok] = rc1[driver]; tcR1[tok] = rr1[driver]; tcR2[tok] = r2;
        }
        const X = ix[n];
        const g = X.keys.get(critIxKey(cr[driver]));
        if (g !== undefined) {
          for (let p = X.gStart[g]; p < X.gStart[g + 1]; p++) {
            const off = X.gRows[p] - rr1[driver];
            let ok = true;
            for (let k = 1; k <= nP; k++) {
              if (k !== driver && !critMatch(getV(rs[k], rr1[k] + off, rc1[k]), cr[k])) { ok = false; break; }
            }
            if (ok) {
              if (isSum) {
                const v = getV(sumS, sumR1 + off, sumC1);
                if (typeof v === 'number') acc = addCancel(acc, v);
                else if (v instanceof XErr) { res = v; break done; }
              } else {
                cnt++;
              }
            }
          }
        }
      }
    } else {
      for (let off = 0; off < maxOff; off++) {
        for (let j = 0; j < nC; j++) {
          let ok = true;
          for (let k = 1; k <= nP; k++) {
            if (!critMatch(getV(rs[k], rr1[k] + off, rc1[k] + j), cr[k])) { ok = false; break; }
          }
          if (ok) {
            if (isSum) {
              const v = getV(sumS, sumR1 + off, sumC1 + j);
              if (typeof v === 'number') acc = addCancel(acc, v);
              else if (v instanceof XErr) { res = v; break done; }
            } else {
              cnt++;
            }
          }
        }
      }
      if (!isSum && nR > maxOff) {                 // the blank rows count when every criterion matches blank
        let ok = true;
        for (let k = 1; k <= nP; k++) if (!critMatch(null, cr[k])) { ok = false; break; }
        if (ok) cnt += (nR - maxOff) * nC;
      }
    }
    res = isSum ? acc : cnt;
  }
  if (useCache) aggCache.set(key, res);
  return res;
}

function rowsWithData(s, r1, nR) {
  let d = shs[s].nRow - r1 + 1;
  if (d < 0) d = 0;
  if (d > nR) d = nR;
  return d;
}

// VLOOKUP(value, table, column, FALSE): exact match only.
function fnVlookup(base, argc) {
  let lv = scalar(base);
  if (lv instanceof XErr) return lv;
  const tb = base + 1;
  if (stK[tb] !== 1) {
    if (stK[tb] === 0 && stV[tb] instanceof XErr) return stV[tb];
    return ERR_NA;
  }
  let col = toNum(scalar(base + 2));
  if (typeof col !== 'number') return col;
  col = Math.trunc(col);
  if (col < 1) return ERR_VALUE;
  if (col > stC2[tb] - stC1[tb] + 1) return ERR_REF;
  let exact = false;
  if (argc >= 4) {
    if (stK[base + 3] === 2) {
      exact = true;                                // VLOOKUP(x, t, 2, ) : an empty argument is FALSE
    } else {
      const rl = scalar(base + 3);
      if (typeof rl === 'boolean') exact = !rl;
      else if (typeof rl === 'number') exact = rl === 0;
      else if (rl === null) exact = true;
      else if (rl instanceof XErr) return rl;
      else if (strCompText(String(rl), 'FALSE') === 0) exact = true;
      else if (strCompText(String(rl), 'TRUE') !== 0) return ERR_VALUE;
    }
  }
  if (!exact) {
    throw new UserError('VLOOKUP แบบค่าประมาณ (TRUE) คำนวณไม่ได้ (เหมือน macro เดิม) ที่ ' +
      shs[stS[tb]].sheet.name + ' สูตรในเซลล์ ' + addr(curRow, curCol));
  }
  const s = stS[tb];
  const rEnd = Math.min(stR2[tb], shs[s].nRow);
  let found = 0;
  if (rEnd >= stR1[tb]) {
    if (typeof lv === 'string' && hasWild(lv)) {
      const re = likeRegex(lv.toUpperCase());
      ensureRange(s, stR1[tb], stC1[tb], rEnd, stC1[tb]);
      for (let r = stR1[tb]; r <= rEnd; r++) {
        const v = getV(s, r, stC1[tb]);
        if (typeof v === 'string' && re.test(v.toUpperCase())) { found = r; break; }
      }
    } else {
      const n = getIndex(s, stC1[tb], stR1[tb], rEnd, true);
      if (lv === null) lv = 0;                     // a blank lookup value looks for 0
      const g = ix[n].keys.get(lookupKey(lv));
      if (g !== undefined) found = ix[n].gRows[ix[n].gStart[g]];
    }
  }
  if (found === 0) return ERR_NA;
  let v = getV(s, found, stC1[tb] + col - 1);
  if (v === null) v = 0;
  return v;
}

function fnInt(base) {
  const x = toNum(scalar(base));
  if (typeof x !== 'number') return x;
  return Math.floor(x);
}

function fnMid(base) {
  const t = toText(scalar(base));
  if (t instanceof XErr) return t;
  let st = toNum(scalar(base + 1));
  if (typeof st !== 'number') return st;
  let n = toNum(scalar(base + 2));
  if (typeof n !== 'number') return n;
  st = Math.trunc(st); n = Math.trunc(n);
  if (st < 1 || n < 0) return ERR_VALUE;
  if (st > t.length) return '';
  if (n > 32767) n = 32767;
  return t.substr(st - 1, n);
}

// Index of one column (rows r1..r2): value key -> rows, in row order.
// lookupMode: VLOOKUP keys (number text stays text), else criteria keys.
function getIndex(s, c, r1, r2, lookupMode) {
  const key = s + ':' + c + ':' + r1 + ':' + r2 + ':' + lookupMode;
  let n = ixByKey.get(key);
  if (n !== undefined) return n;
  ensureRange(s, r1, c, r2, c);
  n = ixByKey.get(key);
  if (n !== undefined) return n;
  const keys = new Map();
  const len = Math.max(0, r2 - r1 + 1);
  const grp = new Int32Array(len + 1), cnt = [0];
  let nG = 0;
  for (let r = r1; r <= r2; r++) {
    const v = getV(s, r, c);
    const k = lookupMode ? lookupKey(v) : critCellKey(v);
    if (k.length > 0) {
      let g = keys.get(k);
      if (g === undefined) { nG++; g = nG; keys.set(k, g); cnt[g] = 0; }
      grp[r - r1 + 1] = g;
      cnt[g]++;
    }
  }
  const gStart = new Int32Array(nG + 2);
  gStart[1] = 1;
  for (let g = 1; g <= nG; g++) gStart[g + 1] = gStart[g] + cnt[g];
  const gRows = new Int32Array(gStart[nG + 1] + 1);
  const fill = gStart.slice();
  for (let i = 1; i <= len; i++) {
    const g = grp[i];
    if (g > 0) { gRows[fill[g]] = r1 + i - 1; fill[g]++; }
  }
  ix.push({ keys, gStart, gRows });
  n = ix.length - 1;
  ixByKey.set(key, n);
  return n;
}

//==============================================================================
//  Criteria
//==============================================================================
// Measured in Excel: a number matches numbers and number text ("5", " 5", "05", "$5");
// a blank criteria cell means 0; "" matches blank and empty text, "=" only blank,
// "<>" anything not blank; "<>5" leaves out only the number 5; > < >= <= look at
// numbers only (or text only); text ignores case; * ? ~ are wildcards; "TRUE" means
// the logical value; an error matches the same error.
function parseCrit(v) {
  const c = { op: OP_EQ, kind: K_NUM, num: 0, txt: '', b: false, e: null, re: null };
  if (typeof v === 'number') { c.num = v; return c; }
  if (v === null) return c;
  if (typeof v === 'boolean') { c.kind = K_BOOL; c.b = v; return c; }
  if (v instanceof XErr) { c.kind = K_ERR; c.e = v; return c; }
  let s = String(v), hasOp = false;
  const p2 = s.slice(0, 2), p1 = s.charAt(0);
  if (p2 === '<>') { c.op = OP_NE; s = s.slice(2); hasOp = true; }
  else if (p2 === '<=') { c.op = OP_LE; s = s.slice(2); hasOp = true; }
  else if (p2 === '>=') { c.op = OP_GE; s = s.slice(2); hasOp = true; }
  else if (p1 === '<') { c.op = OP_LT; s = s.slice(1); hasOp = true; }
  else if (p1 === '>') { c.op = OP_GT; s = s.slice(1); hasOp = true; }
  else if (p1 === '=') { s = s.slice(1); hasOp = true; }
  let x, e;
  if (s.length === 0) {
    c.kind = hasOp ? K_BLANK : K_EMPTYSTR;
  } else if ((x = textToNum(s)) !== null) {
    c.kind = K_NUM; c.num = x;
  } else if (strCompText(s, 'TRUE') === 0) {
    c.kind = K_BOOL; c.b = true;
  } else if (strCompText(s, 'FALSE') === 0) {
    c.kind = K_BOOL; c.b = false;
  } else if ((e = errLiteral(s)) !== null) {
    c.kind = K_ERR; c.e = e;
  } else if ((c.op === OP_EQ || c.op === OP_NE) && hasWild(s)) {
    c.kind = K_LIKE; c.re = likeRegex(s.toUpperCase());
  } else {
    c.kind = K_TEXT; c.txt = s.toUpperCase();
  }
  return c;
}

function critMatch(v, c) {
  let m = false;
  switch (c.kind) {
    case K_NUM:
      if (c.op === OP_EQ) {
        if (typeof v === 'number') return numEq(v, c.num);
        if (typeof v === 'string') { const x = textToNum(v); return x !== null && numEq(x, c.num); }
        return false;
      }
      if (c.op === OP_NE) return typeof v === 'number' ? !numEq(v, c.num) : true;
      if (typeof v === 'number') return opHolds(c.op, numCmp(v, c.num));
      return false;
    case K_TEXT:
      if (typeof v === 'string') {
        if (c.op === OP_EQ) return v === c.txt || v.toUpperCase() === c.txt || strCompText(v, c.txt) === 0;
        return opHolds(c.op, strCompText(v, c.txt));
      }
      return c.op === OP_NE;
    case K_LIKE:
      if (typeof v === 'string') m = c.re.test(v.toUpperCase());
      return c.op === OP_NE ? !m : m;
    case K_BOOL:
      if (typeof v === 'boolean') m = v === c.b;
      return c.op === OP_NE ? !m : m && c.op === OP_EQ;
    case K_ERR:
      if (v instanceof XErr) m = v.text === c.e.text;
      return c.op === OP_NE ? !m : m && c.op === OP_EQ;
    case K_EMPTYSTR:
      if (v === null) return true;
      return typeof v === 'string' && v.length === 0;
    case K_BLANK:
      if (c.op === OP_NE) return v !== null;
      if (c.op === OP_EQ) return v === null;
      return false;
  }
  return false;
}

function opHolds(op, cmp) {
  switch (op) {
    case OP_EQ: return cmp === 0;
    case OP_NE: return cmp !== 0;
    case OP_LT: return cmp < 0;
    case OP_LE: return cmp <= 0;
    case OP_GT: return cmp > 0;
    default: return cmp >= 0;
  }
}

function isIndexable(c) {
  return c.op === OP_EQ && (c.kind === K_NUM || c.kind === K_TEXT || c.kind === K_BOOL || c.kind === K_ERR);
}

// Key a criterion matches in a criteria index (see critCellKey).
function critIxKey(c) {
  switch (c.kind) {
    case K_NUM: return numKey(c.num);
    case K_TEXT: return 's' + c.txt;
    case K_BOOL: return 'b' + c.b;
    default: return 'e' + c.e.text;
  }
}

// Cell key for criteria: number text is a number, text ignores case, blank has no key.
function critCellKey(v) {
  if (typeof v === 'number') return numKey(v);
  if (typeof v === 'string') {
    const x = textToNum(v);
    return x !== null ? numKey(x) : 's' + v.toUpperCase();
  }
  if (typeof v === 'boolean') return 'b' + v;
  if (v instanceof XErr) return 'e' + v.text;
  return '';
}

// Key for VLOOKUP: a number only matches a number, text only text (case ignored).
function lookupKey(v) {
  if (typeof v === 'number') return numKey(v);
  if (typeof v === 'string') return 's' + v.toUpperCase();
  if (typeof v === 'boolean') return 'b' + v;
  return '';
}

// Key of a number, 15 significant digits; -0 is 0.
function numKey(x) {
  return x === 0 ? 'n0' : 'n' + k15(x);
}

function hasWild(s) {
  return s.indexOf('*') >= 0 || s.indexOf('?') >= 0 || s.indexOf('~') >= 0;
}

// Excel wildcards (* ? and ~ as escape) -> a regular expression over upper case text.
function likeRegex(u) {
  let out = '^';
  for (let i = 0; i < u.length; i++) {
    const ch = u.charAt(i);
    if (ch === '~' && i < u.length - 1) {
      i++;
      out += u.charAt(i).replace(/[\\^$.*+?()[\]{}|/-]/g, '\\$&');
    } else if (ch === '*') {
      out += '[\\s\\S]*';
    } else if (ch === '?') {
      out += '[\\s\\S]';
    } else {
      out += ch.replace(/[\\^$.*+?()[\]{}|/-]/g, '\\$&');
    }
  }
  return new RegExp(out + '$');
}
