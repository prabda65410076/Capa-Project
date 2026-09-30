'use strict';
//==============================================================================
//  CAPA calculation: a port of modCapaCalc (VBA), the part CAPA_Recalc runs.
//
//    * ST x coefficient x Plan for every item of sheet 1-2 (12 months x 65
//      processes), added up per OD (spec) into sheet ST_SUM. The 1-2 values
//      CH:AGS are not stored: they are computed when something reads them.
//    * Parts of sheet PART METHOD (FMCNC1 / FMCNC2 / HAND) are kept apart.
//    * Q'TY AGU:AHF of sheet 1-2 added up per ChildP/N into ST_SUM.
//    * Every other formula is evaluated by the engine (engine.js): from the
//      very hidden sheet CAPA_RULES (after CAPA_Setup) and from the cells.
//==============================================================================

// ---- sheet 1-2 (column numbers) -------------------------------------------------
const FIRST_ROW = 7;          // first item row in 1-2 and in 1-3
const COEF_ROW = 3;           // coefficient per process and month
const HDR_ROW = 4;            // "Plan" / process group headers
const COL_PN = 6;             // F   ASSY P/N (key into 1-1.Plan)
const COL_OD = 12;            // L   OD = spec used by sheet 1-3
const COL_ST = 21;            // U   first ST column (U..CG)
const N_PROC = 65;
const COL_PLAN = 86;          // CH  Plan JAN; a new month every 66 columns
const BLOCK = 66;             // Plan column + 65 process columns
const N_MONTH = 12;
const COL_QTY = 879;          // AGU..AHF  =IF($P>0, Plan*$K, 0)
const COL_CHILD = 8;          // H   ChildP/N
const COL_MK_FM = 18;         // R   PROCESS mark "FM", or FMCNC1 / FMCNC2
const COL_MK_BD = 19;         // S   PROCESS mark "BD", or HAND
// ---- sheet 1-1.Plan ---------------------------------------------------------------
const PLAN_ROW = 4;
const PLAN_COL_PN = 3;        // C
const PLAN_COL_M1 = 7;        // G..R = JAN..DEC
// ---- sheet 1-3 ----------------------------------------------------------------------
const CAPA_COL = 73;          // BU  first result column (JAN A/T)
const CAPA_FLAG = 8;          // H   first flag column (H..BT)

const SUM_SHEET = 'ST_SUM';
const METHOD_SHEET = 'PART METHOD';
const RULES_SHEET = 'CAPA_RULES';
const MS_FMKIND = [4, 4];     // D4  Kind (1-3 column B) of the FMCNC1 machines
const MS_BDKIND = [5, 4];     // D5  Kind of the HAND machines
const MS_FM2KIND = [6, 4];    // D6  Kind of the FMCNC2 machines
const MS_FMSEC = [4, 5];      // E4  FMCNC1: sec per operation when column D is empty
const MS_FM2SEC = [6, 5];     // E6  FMCNC2: sec per shot when column D is empty
const MS_FIRST = 8;           // first part row: A Child P/N, B shop, C method, D time
const FM_P1 = 8, FM_P2 = 15;  // Forming processes 8-15: Reduce .. Press Joint
const FM_STOP = 11;           // Stopper: FMCNC2 pairs Reduce (8), Expand (9) and Stopper
const BD_P = 16;              // Bending Pipe
const COL_PCS = 11;           // K   pcs. per assy (1-2)
const COL_CH_FM = 13;         // ST_SUM M: FMCNC1 / FMCNC2 / STD
const COL_CH_BD = 14;         // ST_SUM N: HAND / STD
const NM_FM1 = 'FMCNC1', NM_FM2 = 'FMCNC2', NM_HAND = 'HAND', NM_STD = 'STD';
const FIRST_RULE_ROW = 3;
const BULLET = '●';
const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

//------------------------------------------------------------------------------
//  Small helpers (as in the VBA module)
//------------------------------------------------------------------------------
function findSheet(wb, prefix) {
  const sh = wb.findPrefix(prefix);
  if (sh === null) throw new UserError("ไม่พบชีตที่ชื่อขึ้นต้นด้วย '" + prefix + "'");
  return sh;
}

function lastRow(ws, col) { return ws.lastRowIn(col); }

function at(ws, rc) { return ws.get(rc[0], rc[1]); }

function a1Of(rc) { return addr(rc[0], rc[1]); }

// VBA Trim$: spaces only.
function vbTrim(s) { return s.replace(/^ +| +$/g, ''); }

// CStr of a cell value ("" for an empty or error cell).
function cellText(v) {
  if (v === null || v instanceof XErr) return '';
  if (typeof v === 'number') return cstr(v);
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  return v;
}

// First line of a cell value as text.
function firstLine(v) {
  const s = cellText(v);
  const p = s.indexOf('\n');
  return p >= 0 ? s.slice(0, p) : s;
}

function fmtInt(x) {
  return Math.round(x).toLocaleString('en-US');
}

const RE_PLAINNUM = /^[ \t\n\r\f\v]*[-+]?([0-9]+\.?[0-9]*|\.[0-9]+)([eE][-+]?[0-9]+)?[ \t\n\r\f\v]*$/;
function isPlainNumber(s) {
  if (s.length === 0) return false;
  const ch = s.charCodeAt(0) | 32;
  if (ch >= 97 && ch <= 122) return false;
  return RE_PLAINNUM.test(s);
}

// Excel arithmetic on a cell value (ToNumber of the macro): blank = 0, TRUE = 1, numeric text =
// number, other text = #VALUE!, errors pass through; text of only spaces counts as blank.
// Returns a number or an XErr.
function calcToNumber(x) {
  if (x === null) return 0;
  if (typeof x === 'number') return x;
  if (typeof x === 'boolean') return x ? 1 : 0;
  if (x instanceof XErr) return x;
  if (/^ *$/.test(x.replace(/ /g, ' '))) return 0;
  if (isPlainNumber(x)) {
    const v = Number(x);
    return isFinite(v) ? v : ERR_VALUE;
  }
  return ERR_VALUE;
}

// Key for matching like SUMIF / SUMIFS criteria: text is case-insensitive, a number and the
// same number stored as text match. Blank gives "".
function critKey(x) {
  if (x === null) return '';
  if (typeof x === 'string') {
    if (x.length === 0) return '';
    if (isPlainNumber(x)) return 'N' + k15(Number(x));
    return 'S' + x.toUpperCase();
  }
  if (typeof x === 'boolean') return 'B' + x;
  if (x instanceof XErr) return 'E' + x.text;
  return 'N' + k15(x);
}

function getSumSheet(wb) {
  const ws = wb.sheet(SUM_SHEET);
  if (ws === null) throw new UserError('ยังไม่มีชีต ST_SUM: ต้องรัน CAPA_Setup ใน Excel ก่อนหนึ่งครั้ง');
  return ws;
}

function checkLayout(ws12) {
  for (let m = 0; m < N_MONTH; m++) {
    const c = COL_PLAN + m * BLOCK;
    if (cellText(ws12.get(HDR_ROW, c)) !== 'Plan') {
      throw new UserError("รูปแบบชีต 1-2 ไม่ตรงกับที่ macro คาดไว้: เซลล์ " + addr(HDR_ROW, c) +
        " ต้องเป็นคำว่า 'Plan' (อาจมีการแทรกหรือลบคอลัมน์)");
    }
  }
}

// Writes the block get(i, j) (n rows x w columns) at row0 / col0 and clears the older rows
// below it (only when something differs, like WriteBlockIfChanged). Counts the cells that
// differ from the file in stats.
function writeBlockIfChanged(ws, row0, col0, n, w, get, stats) {
  const lastOld = ws.usedLastRow();
  let changed = false;
  outer:
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < w; j++) {
      if (differs(ws.get(row0 + i, col0 + j), get(i, j))) { changed = true; break outer; }
    }
  }
  if (!changed && lastOld >= row0 + n) {
    outer2:
    for (let r = row0 + n; r <= lastOld; r++) {
      for (let c = col0; c < col0 + w; c++) if (ws.get(r, c) !== null) { changed = true; break outer2; }
    }
  }
  if (!changed) return;
  const last = Math.max(lastOld, row0 + n - 1);
  for (let r = row0; r <= last; r++) {
    for (let c = col0; c < col0 + w; c++) {
      const v = r - row0 < n ? get(r - row0, c - col0) : null;
      const old = ws.get(r, c);
      if (differs(old, v)) {
        ws.set(r, c, v);
        if (stats && !sameValue(old, v)) stats.n++;
      }
    }
  }
}

//------------------------------------------------------------------------------
//  Core: Plan per item, ST x coefficient x Plan, totals per OD -> ST_SUM
//------------------------------------------------------------------------------
// Value of one process cell of sheet 1-2 on the slow path (errors, PART METHOD times).
// r: item index (0 = row 7), k: process 1..65, j: column 1..792 of CH:AGS.
function slowCell(T, r, k, j, pv, pe) {
  let useOv = false, ov = 0, sv = 0, e = null, t;
  if (T.ovr[r]) {
    if (T.fmMode[r] > 0 && k >= FM_P1 && k <= FM_P2) {
      // operations x sec per operation (not the coefficient) x Plan
      useOv = true;
      ov = T.fmSec[r];
      if (k <= FM_P1 + 1 || k === FM_STOP) ov = ov * T.pairF[r];
      t = calcToNumber(T.ws12.get(FIRST_ROW + r, COL_ST + k - 1));
      if (t instanceof XErr) e = t; else sv = t;
      if (e === null) e = pe;
    } else if (k === BD_P && T.bdM[r] && T.bdT[r] !== null) {
      // sec per piece x pcs. (column K) x Plan
      useOv = true;
      ov = T.bdT[r];
      t = calcToNumber(T.pcs[r]);
      if (t instanceof XErr) e = t; else sv = t;
      if (e === null) e = pe;
    }
  }
  if (!useOv) {
    t = calcToNumber(T.ws12.get(FIRST_ROW + r, COL_ST + k - 1));
    if (t instanceof XErr) e = t; else sv = t;
    if (e === null) e = T.coefE[j];
    if (e === null) e = pe;
  }
  if (e === null) return useOv ? sv * ov * pv : sv * T.coefV[j] * pv;
  return e;
}

// Value of sheet 1-2 CH:AGS (what CAPA_Recalc writes there).
function totalsCell(T, row, col) {
  const r = row - FIRST_ROW, j = col - COL_PLAN + 1;
  const m = Math.floor((j - 1) / BLOCK) + 1, k = (j - 1) % BLOCK;
  const q = r * N_MONTH + m - 1;
  const pe = T.planErr === null ? null : T.planErr[q];
  if (k === 0) return pe !== null ? pe : T.planNum[q];
  const pv = pe !== null ? 0 : T.planNum[q];
  if (T.rowErr[r] || T.coefErr[m] || pe !== null || T.ovr[r]) return slowCell(T, r, k, j, pv, pe);
  return pv !== 0 ? T.stV[(k - 1) * T.nRow + r] * T.coefV[j] * pv : 0;
}

function buildTotals(wb, writePlan, rep) {
  const wsPlan = findSheet(wb, '1-1.');
  const ws12 = findSheet(wb, '1-2.');
  ws12.virt = null;
  checkLayout(ws12);
  const wsSum = getSumSheet(wb);
  const nCol = N_MONTH * BLOCK;

  // --- 1) plan quantity per P/N, matched like SUMIF -----------------------------
  const planKeys = new Map();
  let planV = null, planE = null, nKey = 0;
  const lastP = lastRow(wsPlan, PLAN_COL_PN);
  if (lastP >= PLAN_ROW) {
    const nP = lastP - PLAN_ROW + 1;
    planV = new Float64Array(nP * N_MONTH);
    planE = new Array(nP * N_MONTH).fill(null);
    for (let r = 0; r < nP; r++) {
      const key = critKey(wsPlan.get(PLAN_ROW + r, PLAN_COL_PN));
      if (key.length === 0) continue;
      let idx = planKeys.get(key);
      if (idx === undefined) { idx = nKey++; planKeys.set(key, idx); }
      for (let m = 0; m < N_MONTH; m++) {
        const q = idx * N_MONTH + m;
        // SUMIF adds numbers only; an error in a matching row makes the result an error
        if (planE[q] === null) {
          const v = wsPlan.get(PLAN_ROW + r, PLAN_COL_M1 + m);
          if (typeof v === 'number') planV[q] += v;
          else if (v instanceof XErr) planE[q] = v;
        }
      }
    }
  }

  // --- 2) item rows of sheet 1-2 ---------------------------------------------------
  const lastR = lastRow(ws12, COL_PN);
  if (lastR < FIRST_ROW) throw new UserError('ไม่พบแถวสินค้าในชีต 1-2 (คอลัมน์ F)');
  const nRow = lastR - FIRST_ROW + 1;
  const child = new Array(nRow), pcs = new Array(nRow), od = new Array(nRow), pn = new Array(nRow);
  for (let r = 0; r < nRow; r++) {
    pn[r] = ws12.get(FIRST_ROW + r, COL_PN);
    od[r] = ws12.get(FIRST_ROW + r, COL_OD);
    child[r] = ws12.get(FIRST_ROW + r, COL_CHILD);
    pcs[r] = ws12.get(FIRST_ROW + r, COL_PCS);
  }

  // parts of sheet PART METHOD (matched on ChildP/N like SUMIF)
  //   fmMode 0 = not FM CNC, 1 = FMCNC1, 2 = FMCNC2; fmSec = sec per operation (column D or E4 / E6)
  const fmMode = new Uint8Array(nRow), bdM = new Uint8Array(nRow), ovr = new Uint8Array(nRow);
  const bdT = new Array(nRow).fill(null), fmSec = new Float64Array(nRow), pairF = new Float64Array(nRow);
  checkMethodKinds(wb);
  const info = { methodBad: 0 };
  const meth = readMethods(wb, info);
  const used = new Set();
  const secSet = [null, null, null];
  if (meth.size > 0) {
    for (let r = 0; r < nRow; r++) {
      const key = critKey(child[r]);
      if (key.length === 0 || !meth.has(key)) continue;
      const it = meth.get(key);
      fmMode[r] = it[0]; bdM[r] = it[2] ? 1 : 0; bdT[r] = it[3];
      if (fmMode[r] > 0) {
        if (it[1] === null) {
          if (secSet[fmMode[r]] === null) secSet[fmMode[r]] = methodSetting(wb, fmMode[r]);
          fmSec[r] = secSet[fmMode[r]];
        } else {
          fmSec[r] = it[1];
        }
      }
      ovr[r] = fmMode[r] > 0 || (bdM[r] && bdT[r] !== null) ? 1 : 0;
      used.add(key);
    }
  }
  if (writePlan) writeMethodMarks(ws12, nRow, fmMode, bdM);

  const coefV = new Float64Array(nCol + 1), coefE = new Array(nCol + 1).fill(null), coefErr = new Uint8Array(N_MONTH + 1);
  for (let j = 1; j <= nCol; j++) {
    const t = calcToNumber(ws12.get(COEF_ROW, COL_PLAN + j - 1));
    if (t instanceof XErr) { coefE[j] = t; coefErr[Math.floor((j - 1) / BLOCK) + 1] = 1; } else coefV[j] = t;
  }

  // plan per item row (this is what the old SUMIF columns showed)
  const planNum = new Float64Array(nRow * N_MONTH);
  let planErr = null;
  for (let r = 0; r < nRow; r++) {
    const key = critKey(pn[r]);
    const idx = key.length > 0 ? planKeys.get(key) : undefined;
    if (idx === undefined) continue;
    for (let m = 0; m < N_MONTH; m++) {
      const e = planE[idx * N_MONTH + m];
      if (e !== null) {
        if (planErr === null) planErr = new Array(nRow * N_MONTH).fill(null);
        planErr[r * N_MONTH + m] = e;
      } else {
        planNum[r * N_MONTH + m] = planV[idx * N_MONTH + m];
      }
    }
  }

  // OD groups: one ST_SUM row per distinct OD (blank OD is a group too), and apart
  // from it the parts of sheet PART METHOD (FMCNC1 / FMCNC2 / HAND) of that OD
  const grpKeys = new Map(), grpOf = new Int32Array(nRow), grpVal = [], grpFm = [], grpBd = [];
  let nGrp = 0;
  for (let r = 0; r < nRow; r++) {
    const key = critKey(od[r]) + '\t' + fmMode[r] + (bdM[r] ? 'M' : 'S');
    let g = grpKeys.get(key);
    if (g === undefined) {
      g = nGrp++;
      grpKeys.set(key, g);
      grpVal[g] = od[r];
      grpFm[g] = fmMode[r] > 0 ? (fmMode[r] === 1 ? NM_FM1 : NM_FM2) : NM_STD;
      grpBd[g] = bdM[r] ? NM_HAND : NM_STD;
    }
    grpOf[r] = g;
  }

  // --- 3) ST x coefficient x Plan: totals per OD --------------------------------------
  //        (month by month, process by process, the rows in order: the same additions in
  //        the same order as the macro, so the totals are the same to the last bit)
  const stV = new Float64Array(N_PROC * nRow), rowErr = new Uint8Array(nRow);
  for (let k = 1; k <= N_PROC; k++) {
    const off = (k - 1) * nRow;
    for (let r = 0; r < nRow; r++) {
      const raw = ws12.get(FIRST_ROW + r, COL_ST + k - 1);
      if (typeof raw === 'number') {
        stV[off + r] = raw;
      } else if (raw !== null) {
        const t = calcToNumber(raw);
        if (t instanceof XErr) rowErr[r] = 1; else stV[off + r] = t;
      }
    }
  }
  for (let r = 0; r < nRow; r++) {                     // FMCNC2: Reduce / Expand / Stopper take
    pairF[r] = 1;                                     // shots / operations of the time
    if (fmMode[r] === 2) {
      const sv = stV[(FM_P1 - 1) * nRow + r] + stV[FM_P1 * nRow + r] + stV[(FM_STOP - 1) * nRow + r];
      if (sv > 0) pairF[r] = cnc2Shots(sv, pcs[r]) / sv;
    }
  }
  const T = {
    ws12, nRow, lastR, stV, rowErr, coefV, coefE, coefErr, planNum, planErr, fmMode, fmSec, pairF, bdM, bdT, ovr, pcs,
    nGrp, grpOf, grpVal, grpFm, grpBd, tot: null, totE: null,
  };
  const tot = new Float64Array(nGrp * nCol), totE = new Array(nGrp * nCol).fill(null);
  const pvR = new Float64Array(nRow), peR = new Array(nRow), slow = new Uint8Array(nRow);
  for (let m = 1; m <= N_MONTH; m++) {
    const c = (m - 1) * BLOCK + 1;                     // Plan column of this month (1 = CH)
    for (let r = 0; r < nRow; r++) {
      const q = r * N_MONTH + m - 1;
      const pe = planErr === null ? null : planErr[q];
      peR[r] = pe;
      pvR[r] = pe !== null ? 0 : planNum[q];
      // slow path: same error order as =$U7*CI$3*CH7, times from sheet PART METHOD
      slow[r] = rowErr[r] || coefErr[m] || pe !== null || ovr[r] ? 1 : 0;
    }
    for (let k = 1; k <= N_PROC; k++) {
      const j = c + k, cv = coefV[j], kOff = (k - 1) * nRow, tj = j - 1;
      for (let r = 0; r < nRow; r++) {
        if (slow[r]) {
          // SUMIFS keeps the first error (the rows are still added in order)
          const x = slowCell(T, r, k, j, pvR[r], peR[r]);
          const q = grpOf[r] * nCol + tj;
          if (typeof x === 'number') tot[q] += x;
          else if (totE[q] === null) totE[q] = x;
        } else {
          const pv = pvR[r];
          if (pv !== 0) tot[grpOf[r] * nCol + tj] += stV[kOff + r] * cv * pv;
        }
      }
    }
  }
  T.tot = tot;
  T.totE = totE;
  if (writePlan) {
    ws12.virt = {
      r1: FIRST_ROW, r2: lastR, c1: COL_PLAN, c2: COL_PLAN + nCol - 1,
      get: (rr, cc) => totalsCell(T, rr, cc),
      keep: (cc) => (cc - COL_PLAN) % BLOCK === 0,       // the Plan columns (read by AGU:AHF)
      nextKeep: (cc) => COL_PLAN + Math.ceil((cc - COL_PLAN) / BLOCK) * BLOCK,
    };
  }

  // --- 4) write ST_SUM ---------------------------------------------------------------
  const stats = rep ? rep.stSum : null;
  writeBlockIfChanged(wsSum, FIRST_ROW, COL_OD, nGrp, 1, (i) => grpVal[i], stats);
  writeBlockIfChanged(wsSum, FIRST_ROW, COL_PLAN, nGrp, nCol, (i, j) => {
    if (j % BLOCK === 0) return null;                    // leave the Plan columns empty
    const q = i * nCol + j;
    return totE[q] !== null ? totE[q] : tot[q];
  }, stats);
  writeBlockIfChanged(wsSum, FIRST_ROW, COL_CH_FM, nGrp, 2, (i, j) => (j === 0 ? grpFm[i] : grpBd[i]), stats);

  // --- 5) sheet PART METHOD: parts not found, work that no machine takes -----------
  if (rep) {
    rep.methodNote = writePlan ? channelNote(wb, T, meth.size - used.size, info.methodBad) : '';
    rep.items = nRow;
    rep.groups = nGrp;
    rep.methodParts = meth.size;
  }
  return T;
}

// Q'TY columns AGU:AHF of sheet 1-2 added up per ChildP/N -> ST_SUM (read by sheet BY ITEM CODE).
function buildQtyTotals(wb, rep) {
  const ws12 = findSheet(wb, '1-2.');
  const wsSum = getSumSheet(wb);
  const lastR = lastRow(ws12, COL_PN);
  if (lastR < FIRST_ROW) return;
  const nRow = lastR - FIRST_ROW + 1;
  // Only ChildP/N with a non-zero or error value are listed; SUMIF of any other ChildP/N is 0 either way.
  const qKeys = new Map(), qTot = [], qVal = [];
  for (let r = 0; r < nRow; r++) {
    const row = FIRST_ROW + r;
    let hasQty = false;
    for (let m = 0; m < N_MONTH; m++) {
      const v = ws12.get(row, COL_QTY + m);
      if ((typeof v === 'number' && v !== 0) || v instanceof XErr) { hasQty = true; break; }
    }
    if (!hasQty) continue;
    const key = critKey(ws12.get(row, COL_CHILD));
    let g = qKeys.get(key);
    if (g === undefined) {
      g = qTot.length;
      qKeys.set(key, g);
      qVal.push(ws12.get(row, COL_CHILD));
      qTot.push(new Array(N_MONTH).fill(0));
    }
    const t = qTot[g];
    for (let m = 0; m < N_MONTH; m++) {                 // SUMIF: numbers add up, first error wins
      if (t[m] instanceof XErr) continue;
      const v = ws12.get(row, COL_QTY + m);
      if (typeof v === 'number') t[m] += v;
      else if (v instanceof XErr) t[m] = v;
    }
  }
  const n = Math.max(1, qTot.length);
  const stats = rep ? rep.stSum : null;
  writeBlockIfChanged(wsSum, FIRST_ROW, COL_CHILD, n, 1, (i) => (i < qVal.length ? qVal[i] : null), stats);
  writeBlockIfChanged(wsSum, FIRST_ROW, COL_QTY, n, N_MONTH, (i, j) => (i < qTot.length ? qTot[i][j] : null), stats);
}

//------------------------------------------------------------------------------
//  Sheet PART METHOD
//------------------------------------------------------------------------------
// ChildP/N key -> [FM mode 0/1/2, FM time, BD HAND, BD time]; a time is null when not given.
// Rows with another Shop / Method are counted in info.methodBad. The names of v10-v12
// (CNC1 / CNC2 / MANUAL, v9: CNC) still count.
function readMethods(wb, info) {
  const d = new Map();
  info.methodBad = 0;
  const ws = wb.sheet(METHOD_SHEET);
  if (ws === null) return d;
  const lastR = lastRow(ws, 1);
  if (lastR < MS_FIRST) return d;
  for (let row = MS_FIRST; row <= lastR; row++) {
    const key = critKey(ws.get(row, 1));
    if (key.length === 0) continue;
    const shop = vbTrim(firstLine(ws.get(row, 2))).toUpperCase();
    const how = firstLine(ws.get(row, 3)).replace(/ /g, '').toUpperCase();
    const it = d.has(key) ? d.get(key).slice() : [0, null, false, null];
    let mode = 0;
    if (shop === 'FM') {
      if (how === NM_FM1 || how === 'CNC1' || how === 'CNC') mode = 1;
      else if (how === NM_FM2 || how === 'CNC2') mode = 2;
    }
    if (mode > 0) {
      it[0] = mode; it[1] = methodTime(ws.get(row, 4), row);
      d.set(key, it);
    } else if (shop === 'BD' && (how === NM_HAND || how === 'MANUAL')) {
      it[2] = true; it[3] = methodTime(ws.get(row, 4), row);
      d.set(key, it);
    } else {
      info.methodBad++;
    }
  }
  return d;
}

function methodTime(v, row) {
  if (v === null) return null;
  if (typeof v === 'string' && vbTrim(v).length === 0) return null;
  const x = calcToNumber(v);
  if (x instanceof XErr || x < 0) {
    throw new UserError('ชีต ' + METHOD_SHEET + ' แถว ' + row + ': เวลาต้องเป็นตัวเลข (วินาที) หรือเว้นว่าง');
  }
  return x;
}

// Sheet 1-2 PROCESS marks: a row marked FM (column R) shows FMCNC1 / FMCNC2 when its ChildP/N is
// FMCNC1 / FMCNC2 in sheet PART METHOD, a row marked BD (column S) shows HAND when it is HAND,
// and they go back to FM / BD when the part leaves the list. Other cells stay as they are.
function writeMethodMarks(ws12, nRow, fmMode, bdM) {
  const wantFm = ['FM', NM_FM1, NM_FM2], wantBd = ['BD', NM_HAND];
  const fmMarks = 'FM|CNC1|CNC2|' + NM_FM1 + '|' + NM_FM2, bdMarks = 'BD|' + NM_HAND;
  if (vbTrim(cellText(ws12.get(HDR_ROW + 1, COL_MK_FM))).toUpperCase() !== 'FM' ||
      vbTrim(cellText(ws12.get(HDR_ROW + 1, COL_MK_BD))).toUpperCase() !== 'BD') return;
  for (let r = 0; r < nRow; r++) {
    const row = FIRST_ROW + r;
    const a = ws12.get(row, COL_MK_FM), b = ws12.get(row, COL_MK_BD);
    const na = markText(a, fmMarks, wantFm[fmMode[r]]), nb = markText(b, bdMarks, wantBd[bdM[r]]);
    if (differs(a, na)) ws12.set(row, COL_MK_FM, na);
    if (differs(b, nb)) ws12.set(row, COL_MK_BD, nb);
  }
}

// want when v is one of the texts in marks ("A|B", any case, spaces around it) and is not
// exactly want, else v.
function markText(v, marks, want) {
  if (typeof v !== 'string') return v;
  const s = vbTrim(v).toUpperCase();
  if (s.length === 0 || ('|' + marks + '|').indexOf('|' + s + '|') < 0) return v;
  return v !== want ? want : v;
}

// Seconds per operation of FMCNC1 (mode 1, cell E4) or FMCNC2 (mode 2, cell E6).
function methodSetting(wb, mode) {
  const rc = mode === 1 ? MS_FMSEC : MS_FM2SEC;
  const ws = wb.sheet(METHOD_SHEET);
  const x = calcToNumber(ws === null ? null : at(ws, rc));
  if (x instanceof XErr || x <= 0) {
    throw new UserError('ชีต ' + METHOD_SHEET + ' ' + a1Of(rc) + ': เวลาของ ' + (mode === 1 ? NM_FM1 : NM_FM2) +
      ' ต้องเป็นตัวเลขวินาทีต่อครั้งที่มากกว่า 0 (หรือใส่เวลาของชิ้นงานนั้นในคอลัมน์ D)');
  }
  return x;
}

// The 1-3 formulas compare column B with D4, D6 (FMCNC1 / FMCNC2) and D5 (HAND) like Excel's
// "=" (whole text, any case): an empty name would make every machine without a Kind one of
// those machines, and a space around the name would match no machine.
function checkMethodKinds(wb) {
  const ws = wb.sheet(METHOD_SHEET);
  if (ws === null) return;
  const adr = [MS_FMKIND, MS_FM2KIND, MS_BDKIND], nm = [NM_FM1, NM_FM2, NM_HAND], kn = [];
  for (let i = 0; i < 3; i++) {
    kn[i] = cellText(at(ws, adr[i]));
    if (vbTrim(kn[i]).length === 0) {
      throw new UserError('ชีต ' + METHOD_SHEET + ' ' + a1Of(adr[i]) + ' ต้องมีชื่อ Kind ของเครื่อง (ชีต 1-3 คอลัมน์ B) เช่น ' + nm[i]);
    }
    if (kn[i] !== vbTrim(kn[i]) || kn[i].indexOf('\n') >= 0 || kn[i].indexOf('\r') >= 0) {
      throw new UserError('ชีต ' + METHOD_SHEET + ' ' + a1Of(adr[i]) + ": ลบช่องว่างหน้า/หลังชื่อ Kind '" + vbTrim(kn[i]) +
        "' (หรือการขึ้นบรรทัดใหม่) ออก เพราะชีต 1-3 เทียบข้อความทั้งหมด");
    }
    for (let j = 0; j < i; j++) {
      if (strCompText(kn[i], kn[j]) === 0) {
        throw new UserError('ชีต ' + METHOD_SHEET + ' ' + a1Of(adr[j]) + ' และ ' + a1Of(adr[i]) + ' ต้องเป็นชื่อ Kind ที่ต่างกัน');
      }
    }
  }
}

// FMCNC2 shots for n operations of Reduce / Expand / Stopper on one row of sheet 1-2 (n counts
// all pcs. of the assy): one shot does two operations of the same piece. The operations are
// spread over the pieces as evenly as possible and a piece with q operations takes (q + 1) \ 2
// shots. pcs. that is not a whole number from 1 counts as 1 piece; n that is not a whole number
// gives n / 2.
function cnc2Shots(n, pcsV) {
  if (Math.abs(n - vbRound(n)) > 0.000000001 || n > 1000000) return n / 2;
  let p = calcToNumber(pcsV);
  if (p instanceof XErr || p < 1 || p > 1000000 || Math.abs(p - vbRound(p)) > 0.000000001) p = 1;
  const nn = vbRound(n), pc = vbRound(p);
  const q = Math.floor(nn / pc), rm = nn - q * pc;
  return (pc - rm) * Math.floor((q + 1) / 2) + rm * Math.floor((q + 2) / 2);
}

// Warnings: listed parts not in sheet 1-2, and work of the PART METHOD parts (or of the other
// parts) that no machine of sheet 1-3 takes.
function channelNote(wb, T, notFound, methodBad) {
  const ws13 = findSheet(wb, '1-3.');
  const wsM = wb.sheet(METHOD_SHEET);
  const ws12 = T.ws12;
  const nCol = N_MONTH * BLOCK;
  const fm1 = wsM ? cellText(at(wsM, MS_FMKIND)) : '';
  const fm2 = wsM ? cellText(at(wsM, MS_FM2KIND)) : '';
  const bdKind = wsM ? cellText(at(wsM, MS_BDKIND)) : '';
  const lastR = ws13.usedLastRow();
  const rows = [];
  for (let row = FIRST_ROW; row <= lastR; row++) rows.push(row);
  const lines = new Set();
  let msg = '', n = 0, lost = 0;
  for (let g = 0; g < T.nGrp; g++) {
    const spec = critKey(T.grpVal[g]);
    for (let k = FM_P1; k <= BD_P; k++) {                 // Forming 8-15 and Bending Pipe
      let need = 0;
      for (let m = 1; m <= N_MONTH; m++) need += T.tot[g * nCol + (m - 1) * BLOCK + k];
      if (need === 0) continue;
      let special, kind;
      if (k === BD_P) {
        special = T.grpBd[g] !== NM_STD; kind = bdKind;
      } else {
        special = T.grpFm[g] !== NM_STD;
        kind = T.grpFm[g] === NM_FM2 ? fm2 : fm1;
      }
      let found = false, anyKind = false;
      for (const row of rows) {
        const flag = ws13.get(row, CAPA_FLAG + k - 1);
        if (flag !== BULLET) continue;
        if (critKey(ws13.get(row, 7)) !== spec) continue;
        const rowKind = cellText(ws13.get(row, 2));
        if (special) {
          if (strCompText(rowKind, kind) === 0) { found = true; break; }
        } else {
          const isSpecial = k === BD_P ? strCompText(rowKind, bdKind) === 0 :
            strCompText(rowKind, fm1) === 0 || strCompText(rowKind, fm2) === 0;
          if (!isSpecial) { found = true; break; }
          anyKind = true;
        }
      }
      if (!found && (special || anyKind)) {
        let what;
        if (special) {
          what = (k === BD_P ? NM_HAND : T.grpFm[g]) + " -> ไม่มีเครื่องที่ Kind = '" + kind + "'";
        } else if (k === BD_P) {
          what = "ชิ้นงานอื่น -> ทุกเครื่องเป็น Kind '" + bdKind + "'";
        } else {
          what = "ชิ้นงานอื่น -> ทุกเครื่องเป็น Kind '" + fm1 + "' หรือ '" + fm2 + "'";
        }
        lost += need;
        const ln = '- spec ' + cellText(T.grpVal[g]) + ', ' + firstLine(ws12.get(HDR_ROW + 1, COL_ST + k - 1)) + ': ' + what;
        if (!lines.has(ln)) {
          lines.add(ln);
          n++;
          if (n <= 8) msg += '\n' + ln;
        }
      }
    }
  }
  if (n > 8) msg += '\n- ... และอีก ' + (n - 8) + ' รายการ';
  if (n > 0) {
    msg = '\n\nชีต ' + METHOD_SHEET + ': งานประมาณ ' + fmtInt(lost / 3600) + ' ชั่วโมงต่อปี ไม่ถูกนับให้เครื่องใดในชีต 1-3 ' +
      '(ตรวจสอบคอลัมน์ B Kind และเครื่องหมาย ● ในคอลัมน์ H..BT):' + msg;
  }
  // a Kind that is almost one of D4 / D6 / D5 counts as another machine for the formulas
  let near = 0, nearRow = 0, nearText = '';
  for (const row of rows) {
    const rowKind = cellText(ws13.get(row, 2));
    if (strCompText(rowKind, fm1) === 0 || strCompText(rowKind, fm2) === 0 || strCompText(rowKind, bdKind) === 0) continue;
    const clean = vbTrim(rowKind.replace(/[\r\n]/g, ' '));
    let isNear = false;
    for (const kk of [fm1, fm2, bdKind]) if (strCompText(clean, kk) === 0) isNear = true;
    if (/^FMCNC/.test(clean.replace(/ /g, '').toUpperCase())) isNear = true;
    if (isNear) {
      near++;
      if (nearRow === 0) { nearRow = row; nearText = rowKind; }
    }
  }
  if (near > 0) {
    msg += '\n\nชีต 1-3 คอลัมน์ B: มี ' + near + " แถวที่ Kind เกือบตรงแต่ไม่ตรงกับชีต " + METHOD_SHEET + " (D4 '" + fm1 +
      "', D6 '" + fm2 + "', D5 '" + bdKind + "') เช่นแถว " + nearRow + " '" + nearText +
      "' จึงนับเป็นเครื่องอื่น ให้พิมพ์ชื่อให้ตรงกับ D4 / D6 / D5 หรือแก้ D4 / D6";
  }
  if (notFound > 0) msg += '\n\nชีต ' + METHOD_SHEET + ': มี ChildP/N ' + notFound + ' รายการที่ไม่พบในชีต 1-2 (คอลัมน์ H)';
  if (methodBad > 0) {
    msg += '\n\nชีต ' + METHOD_SHEET + ': ' + methodBad + ' แถวไม่ถูกใช้ เพราะ Shop / Method ต้องเป็น FM + ' + NM_FM1 +
      ' หรือ ' + NM_FM2 + ' หรือ BD + ' + NM_HAND;
  }
  return msg.replace(/^\n+/, '');
}

//------------------------------------------------------------------------------
//  Formula rules (CAPA_RULES) and structure checks
//------------------------------------------------------------------------------
// "VALUES" (formulas are rules, cells hold values), "FORMULAS" (the formulas are back in the
// cells) or "" (no rules yet).
function rulesMode(wb) {
  const ws = wb.sheet(RULES_SHEET);
  return ws === null ? '' : cellText(ws.get(1, 3));
}

// True when the formula rules of sheet 1-3 still pick the ST_SUM channels by the names of
// v10-v12 (CNC1 / CNC2 / MANUAL).
function rulesUseOldNames(wb) {
  const ws = wb.sheet(RULES_SHEET);
  if (ws === null) return false;
  const n = lastRow(ws, 2);
  for (let r = FIRST_RULE_ROW; r <= n; r++) {
    const s = ws.get(r, 2);
    if (typeof s !== 'string') continue;
    if (s.indexOf('ST_SUM!C' + COL_CH_FM + ',IF(') >= 0 || s.indexOf('ST_SUM!C' + COL_CH_BD + ',IF(') >= 0) {
      if (s.indexOf('"CNC1"') >= 0 || s.indexOf('"CNC2"') >= 0 || s.indexOf('"MANUAL"') >= 0 || s.indexOf('"CNC"') >= 0) {
        return true;
      }
    }
  }
  return false;
}

// "Sheet!$A$1:$B$2" -> {sheet, r1, c1, r2, c2} or null.
function parseRefersTo(text) {
  const t = String(text).replace(/^=/, '');
  const k = t.lastIndexOf('!');
  if (k < 0) return null;
  let sheet = t.slice(0, k);
  if (sheet.charAt(0) === "'") sheet = sheet.slice(1, -1).replace(/''/g, "'");
  const rg = parseAddress(t.slice(k + 1));
  if (rg === null) return null;
  rg.sheet = sheet;
  return rg;
}

function parseAddress(s) {
  const m = /^\$?([A-Z]{1,3})\$?([0-9]+)(?::\$?([A-Z]{1,3})\$?([0-9]+))?$/i.exec(String(s).trim());
  if (!m) return null;
  const c1 = colNumber(m[1].toUpperCase()), r1 = +m[2];
  const c2 = m[3] ? colNumber(m[3].toUpperCase()) : c1, r2 = m[4] ? +m[4] : r1;
  return { r1, c1, r2, c2 };
}

// "" when no row or column was inserted or deleted where the rules work, else a message.
function checkStructure(wb) {
  const ws = wb.sheet(RULES_SHEET);
  if (ws === null) return '';
  const n = lastRow(ws, 6);
  if (n < FIRST_RULE_ROW) return '';
  let msg = '';
  for (let r = FIRST_RULE_ROW; r <= n; r++) {
    const shName = cellText(ws.get(r, 6)), nm = cellText(ws.get(r, 7)), adr = cellText(ws.get(r, 8));
    const dn = nm ? wb.definedName(nm) : null;
    const rg = dn ? parseRefersTo(dn.text) : null;
    const target = rg ? wb.sheet(rg.sheet) : null;
    if (target === null) {
      msg += '\n- ชีต ' + shName + ': ถูกลบ หรือมีการลบแถว/คอลัมน์';
    } else if (target.name !== shName) {
      msg += '\n- ชีต ' + shName + ' ถูกเปลี่ยนชื่อเป็น ' + target.name;
    } else {
      const want = parseAddress(adr);
      if (want === null || want.r1 !== rg.r1 || want.c1 !== rg.c1 || want.r2 !== rg.r2 || want.c2 !== rg.c2) {
        msg += '\n- ชีต ' + shName + ': มีการแทรกหรือลบแถว/คอลัมน์';
      }
    }
  }
  if (msg.length === 0) return '';
  return 'ชีตมีการเปลี่ยนแปลงหลังจากที่นำสูตรออกไปแล้ว:' + msg + '\n\n' +
    'จึงไม่รู้ว่าผลลัพธ์ควรอยู่ที่เซลล์ใด ให้ยกเลิกการแก้ไขนั้น (หรือเปิดไฟล์เดิมที่ยังไม่ได้บันทึก) ' +
    'แล้วใน Excel ให้รัน CAPA_EditFormulas แก้ไขตามต้องการ แล้วรัน CAPA_Setup';
}

// All formulas as engine rules: the formulas in the cells, and in VALUES mode the rules of
// sheet CAPA_RULES (for the cells that hold no formula). A formula the engine cannot do keeps
// the value it has in the file (listed in rep.unsupported).
function captureRules(wb, valuesMode, rep) {
  engReset(wb);
  const ws12 = wb.findPrefix('1-2.');
  const unsupported = rep.unsupported;
  const sharedRule = new Map();
  let nFile = 0;
  for (const sh of wb.sheets) {
    const up = sh.name.toUpperCase();
    if (up === RULES_SHEET || up === SUM_SHEET) continue;
    for (const [key, fr] of sh.formulas) {
      const r = keyRow(key), c = keyCol(key);
      if (sh === ws12 && r >= FIRST_ROW && c >= COL_PLAN && c < COL_PLAN + N_MONTH * BLOCK) continue;
      if (fr.kind === 'array' || fr.kind === 'other') {
        unsupported.push({ sheet: sh.name, r, c, why: fr.kind === 'array' ? 'สูตรแบบ array' : 'สูตรตารางข้อมูล (data table)' });
        continue;
      }
      let ru;
      if (fr.kind === 'shared') {
        const k = sh.index + ':' + fr.si;
        ru = sharedRule.get(k);
        if (ru === undefined) {
          const ms = sh.shared.get(fr.si);
          if (ms === undefined) {
            unsupported.push({ sheet: sh.name, r, c, why: 'สูตรแบบ shared ที่ไม่มีต้นแบบ' });
            continue;
          }
          ru = engAddRule(sh.index, a1ToR1C1(ms.text, ms.r, ms.c), { kind: 'cell' });
          sharedRule.set(k, ru);
        }
      } else {
        ru = engAddRule(sh.index, a1ToR1C1(fr.text, r, c), { kind: 'cell' });
      }
      engAddCell(sh.index, r, c, ru);
      nFile++;
    }
  }
  let nRules = 0;
  if (valuesMode) {
    const ws = wb.sheet(RULES_SHEET);
    const n = lastRow(ws, 1);
    for (let row = FIRST_RULE_ROW; row <= n; row++) {
      const shName = cellText(ws.get(row, 1));
      if (shName.length === 0) continue;
      const target = wb.sheet(shName);
      if (target === null) {
        throw new UserError("ไม่พบชีต '" + shName + "' (ที่กฎสูตรใช้อยู่) ถูกเปลี่ยนชื่อหรือลบไปหรือไม่?");
      }
      const ru = engAddRule(target.index, cellText(ws.get(row, 2)), { kind: 'rules', row });
      const rects = cellText(ws.get(row, 3));
      for (const part of rects.split(',')) {
        const q = part.trim().split(/ +/).map(Number);
        if (q.length !== 4 || q.some((x) => !(x >= 1))) continue;
        for (let c = q[1]; c <= q[3]; c++) {
          for (let r = q[0]; r <= q[2]; r++) {
            if (target.hasFormula(r, c)) continue;        // typed after the conversion: its own formula is used
            engAddCell(target.index, r, c, ru);
            nRules++;
          }
        }
      }
    }
  }
  engCompileAll();
  for (let ru = 1; ru <= engRuleCount(); ru++) {
    const bad = engRuleBad(ru);
    if (bad === null) continue;
    const src = engRuleSrc(ru);
    if (src.kind === 'rules') {
      throw new UserError('สูตรนี้ในกฎสูตร (CAPA_RULES แถว ' + src.row + ') คำนวณไม่ได้:\n' + engRuleText(ru) + '\n(' + bad + ')');
    }
  }
  const cellsOfRule = engBuildCells();
  // the cells of formulas that cannot be compiled keep their values
  for (const sh of wb.sheets) {
    for (const [key, fr] of sh.formulas) {
      if (fr.kind === 'array' || fr.kind === 'other') continue;
      if (fr.kind === 'shared' && !sh.shared.has(fr.si)) continue;
      const r = keyRow(key), c = keyCol(key);
      if (engCellId(sh.index, r, c) !== 0) continue;
      if (sh === ws12 && r >= FIRST_ROW && c >= COL_PLAN && c < COL_PLAN + N_MONTH * BLOCK) continue;
      const up = sh.name.toUpperCase();
      if (up === RULES_SHEET || up === SUM_SHEET) continue;
      const text = fr.kind === 'shared' ? (sh.shared.get(fr.si) || {}).text : fr.text;
      let why = 'คำนวณไม่ได้';
      if (text !== undefined) {
        const f = fr.kind === 'shared' ? a1ToR1C1(sh.shared.get(fr.si).text, sh.shared.get(fr.si).r, sh.shared.get(fr.si).c) :
          a1ToR1C1(text, r, c);
        const ru = ruByKey.get(sh.index + '\t' + f);
        if (ru !== undefined && engRuleBad(ru) !== null) why = engRuleBad(ru);
      }
      unsupported.push({ sheet: sh.name, r, c, why, formula: text === undefined ? '' : '=' + text });
    }
  }
  rep.rules = engRuleCount();
  rep.fileFormulaCells = nFile;
  rep.ruleCells = nRules;
  rep.cells = engCellCount();
  rep.tokens = engTokenCount();
  return cellsOfRule;
}

//------------------------------------------------------------------------------
//  CAPA_Recalc
//------------------------------------------------------------------------------
function runRecalc(wb, progress) {
  const rep = {
    mode: '', methodNote: '', unsupported: [], stSum: { n: 0 }, times: [], changes: [], changedCells: 0,
    items: 0, groups: 0, rules: 0, cells: 0, tokens: 0,
  };
  const tick = (() => { let t = Date.now(); return (name) => { const n = Date.now(); rep.times.push([name, (n - t) / 1000]); t = n; }; })();
  const say = (text) => { if (progress) progress({ stage: 'calc', text }); };
  if (wb.sheet(SUM_SHEET) === null) {
    throw new UserError('ยังไม่มีชีต ST_SUM: ต้องรัน CAPA_Setup ใน Excel ก่อนหนึ่งครั้ง (ใช้ไฟล์ที่แปลงด้วย macro เวอร์ชัน v13 แล้ว)');
  }
  rep.mode = rulesMode(wb);
  const valuesMode = rep.mode === 'VALUES';
  if (valuesMode) {
    const msg = checkStructure(wb);
    if (msg.length > 0) throw new UserError(msg);
    if (rulesUseOldNames(wb)) {
      throw new UserError('สูตรของชีต 1-3 ยังใช้ชื่อวิธีการผลิตของ v10-v12 (CNC1 / CNC2 / MANUAL) แต่ ST_SUM ใช้ ' +
        NM_FM1 + ' / ' + NM_FM2 + ' / ' + NM_HAND + ' แล้ว ให้รัน CAPA_EditFormulas แล้ว CAPA_Setup ใน Excel');
    }
  }
  say('คำนวณ ST × coefficient × Plan ของทุกรายการในชีต 1-2...');
  const T = buildTotals(wb, true, rep);
  tick('ST × coefficient × Plan (1-2 → ST_SUM)');
  say('เตรียมสูตรของทุกชีต...');
  captureRules(wb, valuesMode, rep);
  tick('อ่านและแปลงสูตร');
  say("คำนวณสูตรของชีต 1-2 (Q'TY)...");
  engEvaluate('1-2.');
  engWrite(0);
  buildQtyTotals(wb, rep);
  tick("Q'TY ของชีต 1-2 → ST_SUM");
  say('คำนวณสูตรของทุกชีต...');
  engEvaluate('', (name) => say('คำนวณสูตรของชีต ' + name + '...'));
  engWrite(0);
  tick('สูตรของทุกชีต');
  rep.evalNote = engTimeNoteText();
  // what changed compared with the file
  for (const sh of wb.sheets) {
    let n = 0;
    const ex = [];
    for (const k of sh.dirty) {
      const o = sh.orig.get(k), v = sh.getStored(keyRow(k), keyCol(k));
      if (!sameValue(o, v)) {
        n++;
        if (ex.length < 200) ex.push({ r: keyRow(k), c: keyCol(k), old: o, val: v });
      }
    }
    if (n > 0) rep.changes.push({ sheet: sh.name, index: sh.index, n, ex });
    rep.changedCells += n;
  }
  rep.workload = workload(T);
  return rep;
}

// Hours of work per process and month (all specs): ST x coefficient x Plan / 3600, as
// [hours, error]: error is the first error of a spec (its hours are not in the sum).
function workload(T) {
  const nCol = N_MONTH * BLOCK;
  const ws12 = T.ws12;
  const procs = [];
  for (let k = 1; k <= N_PROC; k++) {
    const name = firstLine(ws12.get(HDR_ROW + 1, COL_ST + k - 1)) || ('Process ' + k);
    const group = firstLine(ws12.get(HDR_ROW, COL_ST + k - 1));
    const months = [];
    for (let m = 1; m <= N_MONTH; m++) {
      const j = (m - 1) * BLOCK + k;
      let sum = 0, err = null;
      for (let g = 0; g < T.nGrp; g++) {
        const e = T.totE[g * nCol + j];
        if (e !== null) { if (err === null) err = e.text; } else sum += T.tot[g * nCol + j];
      }
      months.push([sum / 3600, err]);
    }
    procs.push({ k, name, group, months });
  }
  return { procs, months: MONTHS };
}
