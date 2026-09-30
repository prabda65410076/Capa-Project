'use strict';
//==============================================================================
//  What the page asks for: load a file, recalculate, read cells, export.
//  Runs in a Web Worker (worker glue in the page) or, if that is not possible,
//  on the page itself.
//==============================================================================

const session = { wb: null, rep: null, unsupported: null };

// Values sent to the page: an error becomes {e: '#N/A'}.
function wireVal(v) {
  return v instanceof XErr ? { e: v.text } : v;
}

function fmtVal(v) {
  if (v === null) return '';
  if (typeof v === 'number') {
    if (v === 0) return '0';
    const a = Math.abs(v);
    if (a >= 1e15 || a < 1e-4) return v.toExponential(4).replace('e', 'E');
    if (Number.isInteger(v)) return v.toLocaleString('en-US');
    return v.toLocaleString('en-US', { maximumFractionDigits: a >= 1000 ? 2 : a >= 1 ? 4 : 6 });
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof XErr) return v.text;
  return v;
}

function skipRegion(wb, sh) {
  return sh === wb.findPrefix('1-2.') ? { r1: FIRST_ROW, c1: COL_PLAN, c2: COL_PLAN + N_MONTH * BLOCK - 1 } : null;
}

function sheetList(wb) {
  return wb.sheets.map((sh) => {
    let rows = sh.maxRow, cols = sh.maxCol;
    if (sh.virt !== null) { rows = Math.max(rows, sh.virt.r2); cols = Math.max(cols, sh.virt.c2); }
    return { name: sh.name, state: sh.state, rows, cols, formulas: sh.formulas.size };
  });
}

async function apiLoad(file, progress) {
  session.wb = null;
  session.rep = null;
  session.unsupported = null;
  const t0 = Date.now();
  const wb = await loadWorkbook(file, { skipFor: skipRegion }, progress);
  session.wb = wb;
  return {
    fileName: wb.fileName, size: file.size, readTime: (Date.now() - t0) / 1000,
    mode: rulesMode(wb), sheets: sheetList(wb),
  };
}

async function apiRecalc(progress) {
  const wb = session.wb;
  if (wb === null) throw new UserError('ยังไม่ได้เปิดไฟล์');
  const t0 = Date.now();
  const rep = runRecalc(wb, progress);
  session.rep = rep;
  session.unsupported = new Map();
  for (const u of rep.unsupported) {
    const sh = wb.sheet(u.sheet);
    if (sh) session.unsupported.set(sh.index + ':' + cellKey(u.r, u.c), u);
  }
  return {
    time: (Date.now() - t0) / 1000,
    mode: rep.mode,
    times: rep.times,
    evalNote: rep.evalNote,
    methodNote: rep.methodNote,
    items: rep.items, groups: rep.groups, methodParts: rep.methodParts,
    rules: rep.rules, cells: rep.cells, tokens: rep.tokens,
    stSumDiff: rep.stSum.n,
    changedCells: rep.changedCells,
    changes: rep.changes.map((c) => ({
      sheet: c.sheet, index: c.index, n: c.n,
      ex: c.ex.map((x) => ({ r: x.r, c: x.c, a: addr(x.r, x.c), old: fmtVal(x.old), val: fmtVal(x.val) })),
    })),
    unsupportedCount: rep.unsupported.length,
    unsupported: rep.unsupported.slice(0, 50).map((u) => ({ sheet: u.sheet, a: addr(u.r, u.c), why: u.why, formula: u.formula || '' })),
    workload: rep.workload,
    sheets: sheetList(wb),
  };
}

// Cells r1..r2 x c1..c2 of a sheet for the viewer: text and kind flags per cell (row by row).
// kind: 1 number, 2 error, 4 calculated, 8 changed, 16 not calculated (kept), 32 TRUE/FALSE
function apiBlock(si, r1, c1, r2, c2) {
  const wb = session.wb;
  const sh = wb.sheets[si];
  const nr = Math.max(0, r2 - r1 + 1), nc = Math.max(0, c2 - c1 + 1);
  const texts = new Array(nr * nc), kinds = new Uint8Array(nr * nc);
  const vt = sh.virt, isSum = sh.name.toUpperCase() === SUM_SHEET, uns = session.unsupported;
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const i = (r - r1) * nc + (c - c1);
      const v = sh.get(r, c);
      texts[i] = fmtVal(v);
      let k = 0;
      if (typeof v === 'number') k |= 1;
      else if (v instanceof XErr) k |= 2;
      else if (typeof v === 'boolean') k |= 32;
      if (session.rep !== null) {
        const key = cellKey(r, c);
        if ((vt !== null && r >= vt.r1 && r <= vt.r2 && c >= vt.c1 && c <= vt.c2) || engCellId(si, r, c) > 0 ||
            (isSum && r >= FIRST_ROW && v !== null)) k |= 4;
        if (sh.dirty.has(key) && !sameValue(sh.orig.get(key), sh.getStored(r, c))) k |= 8;
        if (uns !== null && uns.has(si + ':' + key)) k |= 16;
      }
      kinds[i] = k;
    }
  }
  return { texts, kinds };
}

function apiCell(si, r, c) {
  const wb = session.wb;
  const sh = wb.sheets[si];
  const key = cellKey(r, c);
  const v = sh.get(r, c);
  const out = { sheet: sh.name, a: addr(r, c), value: fmtVal(v), raw: wireVal(v), type: v === null ? 'ว่าง' :
    typeof v === 'number' ? 'ตัวเลข' : typeof v === 'string' ? 'ข้อความ' : typeof v === 'boolean' ? 'TRUE/FALSE' : 'error' };
  if (sh.dirty.has(key)) {
    const o = sh.orig.get(key);
    if (!sameValue(o, sh.getStored(r, c))) out.old = fmtVal(o) || '(ว่าง)';
  }
  const fr = sh.formulas.get(key);
  if (fr !== undefined) {
    if (fr.kind === 'shared') {
      const ms = sh.shared.get(fr.si);
      if (ms) out.formula = r1c1ToA1(a1ToR1C1(ms.text, ms.r, ms.c), r, c);
    } else if (fr.text) {
      out.formula = '=' + fr.text;
    }
    out.source = 'สูตรในเซลล์';
  }
  const id = session.rep !== null ? engCellId(si, r, c) : 0;
  if (id > 0) {
    const info = engCellInfo(id);
    const src = engRuleSrc(info.rule);
    if (!out.formula) out.formula = r1c1ToA1(engRuleText(info.rule), r, c);
    out.r1c1 = engRuleText(info.rule);
    out.source = src.kind === 'rules' ? 'กฎสูตรในชีต CAPA_RULES แถว ' + src.row : 'สูตรในเซลล์';
  }
  const vt = sh.virt;
  if (vt !== null && r >= vt.r1 && r <= vt.r2 && c >= vt.c1 && c <= vt.c2) {
    const j = c - COL_PLAN, m = Math.floor(j / BLOCK), k = j % BLOCK;
    out.source = 'คำนวณโดยเว็บ (CAPA_Recalc)';
    out.formula = k === 0 ? 'Plan ' + MONTHS[m] + ' = SUMIF ของชีต 1-1.Plan ตาม ASSY P/N (คอลัมน์ F)' :
      'ST (' + addr(r, COL_ST + k - 1) + ') × coefficient (' + addr(COEF_ROW, c) + ') × Plan (' +
      addr(r, COL_PLAN + m * BLOCK) + ')';
  }
  if (sh.name.toUpperCase() === SUM_SHEET && r >= FIRST_ROW) {
    out.source = 'คำนวณโดยเว็บ (CAPA_Recalc): ผลรวมของชีต 1-2 ตาม OD / ChildP/N';
  }
  if (session.unsupported !== null) {
    const u = session.unsupported.get(si + ':' + key);
    if (u) out.note = 'เว็บคำนวณสูตรนี้ไม่ได้ (' + u.why + ') จึงใช้ค่าเดิมในไฟล์';
  }
  return out;
}

// Finds text in a sheet (value or formula). Returns the first cells found.
function apiFind(si, text, max) {
  const sh = session.wb.sheets[si];
  const u = String(text).toUpperCase();
  const out = [];
  if (u.length === 0) return out;
  const rows = sh.maxRow;
  for (let c = 1; c <= sh.maxCol && out.length < max; c++) {
    const col = sh.cols[c];
    if (col === undefined) continue;
    for (let r = 1; r <= rows && out.length < max; r++) {
      const v = sh.getStored(r, c);
      if (v !== null && fmtVal(v).toUpperCase().indexOf(u) >= 0) out.push({ r, c, a: addr(r, c), value: fmtVal(v) });
    }
  }
  out.sort((a, b) => a.r - b.r || a.c - b.c);
  return out;
}

async function apiExport(opts, progress) {
  if (session.wb === null || session.rep === null) throw new UserError('ยังไม่ได้คำนวณ');
  const blob = await exportWorkbook(session.wb, opts, progress);
  const name = session.wb.fileName.replace(/(\.[^.]+)?$/, (ext) => '_web' + (ext || '.xlsx'));
  return { blob, name };
}

// One entry point for the page: cmd + args -> result (progress(p) for long steps).
async function apiCall(cmd, args, progress) {
  switch (cmd) {
    case 'load': return apiLoad(args.file, progress);
    case 'recalc': return apiRecalc(progress);
    case 'block': return apiBlock(args.si, args.r1, args.c1, args.r2, args.c2);
    case 'cell': return apiCell(args.si, args.r, args.c);
    case 'find': return apiFind(args.si, args.text, args.max || 50);
    case 'export': return apiExport(args.opts || {}, progress);
    default: throw new Error('unknown command ' + cmd);
  }
}

function errorText(e) {
  if (e instanceof UserError) return { user: true, message: e.message };
  if (e && e.name === 'RangeError' && /call stack/i.test(e.message)) {
    return { user: true, message: 'สูตรอ้างอิงต่อกันลึกเกินไป (เกินขีดจำกัดของเบราว์เซอร์)' };
  }
  return { user: false, message: (e && (e.stack || e.message)) || String(e) };
}
