'use strict';
//==============================================================================
//  Reading .xlsx / .xlsm: workbook, shared strings and the sheets.
//  The sheet XML is scanned as a byte stream (SheetScanner); the same scanner
//  is used to write a sheet back (xlsxwrite.js).
//==============================================================================

const utf8Dec = new TextDecoder('utf-8');
const utf8Enc = new TextEncoder();

function xmlUnescape(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#[0-9]+);/g, (m, g) => {
    switch (g) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
    }
    const cp = g[1] === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
    return String.fromCodePoint(cp);
  });
}

// OOXML escapes of characters XML cannot hold: _x000D_ etc.
function xUnescape(s) {
  if (s.indexOf('_x') < 0) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
}

function parseAttrs(s) {
  const a = {};
  const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(s)) !== null) a[m[1]] = xmlUnescape(m[2] !== undefined ? m[2] : m[3]);
  return a;
}

function resolvePath(base, target) {
  if (target.charAt(0) === '/') return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

function relsPath(part) {
  const i = part.lastIndexOf('/');
  return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels';
}

async function readXmlText(zip, name) {
  const e = zipEntry(zip, name);
  if (e === null) return null;
  return utf8Dec.decode(await zipReadAll(zip, e));
}

// Relationships of a part: [{id, type, target}] (target resolved against the part).
async function readRels(zip, part) {
  const xml = await readXmlText(zip, relsPath(part));
  const out = [];
  if (xml === null) return out;
  const re = /<Relationship\b([^>]*?)\/?>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const a = parseAttrs(m[1]);
    const external = a.TargetMode === 'External';
    out.push({ id: a.Id, type: a.Type || '', target: external ? a.Target : resolvePath(part, a.Target || '') });
  }
  return out;
}

function parseSharedStrings(xml) {
  const out = [];
  const reSi = /<si>([\s\S]*?)<\/si>|<si\/>/g;
  const reT = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
  let m;
  while ((m = reSi.exec(xml)) !== null) {
    let inner = m[1] || '';
    if (inner.indexOf('<rPh') >= 0) inner = inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let s = '', t;
    reT.lastIndex = 0;
    while ((t = reT.exec(inner)) !== null) s += t[1];
    out.push(xUnescape(xmlUnescape(s)));
  }
  return out;
}

//------------------------------------------------------------------------------
//  Byte helpers
//------------------------------------------------------------------------------
const B_LT = 60, B_GT = 62, B_SL = 47, B_EQ = 61, B_QU = 34, B_AP = 39;

function bytesOf(s) { return utf8Enc.encode(s); }
const SEQ_SHEETDATA = bytesOf('<sheetData');
const SEQ_ROW_END = bytesOf('</row>');
const SEQ_IS_END = bytesOf('</is>');
const A_R = bytesOf('r'), A_T = bytesOf('t'), A_S = bytesOf('s'), A_SI = bytesOf('si'), A_SPANS = bytesOf('spans');

function indexOfSeq(buf, seq, from) {
  const first = seq[0], m = seq.length, lim = buf.length - m;
  let i = from;
  while (i <= lim) {
    i = buf.indexOf(first, i);
    if (i < 0 || i > lim) return -1;
    let k = 1;
    while (k < m && buf[i + k] === seq[k]) k++;
    if (k === m) return i;
    i++;
  }
  return -1;
}

function isWs(b) { return b === 32 || b === 9 || b === 10 || b === 13; }

// Start of the value of attribute `name` in buf[s..e) (a start tag after its element
// name), or -1. attrEnd is set to the position of the closing quote.
let attrEnd = 0, attrNameStart = 0;
function findAttr(buf, s, e, name) {
  let i = s;
  const nl = name.length;
  while (i < e) {
    const b = buf[i];
    if (isWs(b)) { i++; continue; }
    const ns = i;
    while (i < e) {
      const c = buf[i];
      if (c === B_EQ || isWs(c) || c === B_GT || c === B_SL) break;
      i++;
    }
    const ne = i;
    while (i < e && isWs(buf[i])) i++;
    if (i >= e || buf[i] !== B_EQ) { i++; continue; }
    i++;
    while (i < e && isWs(buf[i])) i++;
    const q = buf[i];
    if (q !== B_QU && q !== B_AP) return -1;
    const vs = i + 1;
    const ve = buf.indexOf(q, vs);
    if (ve < 0 || ve >= e) return -1;
    if (ne - ns === nl) {
      let k = 0;
      while (k < nl && buf[ns + k] === name[k]) k++;
      if (k === nl) { attrEnd = ve; attrNameStart = ns; return vs; }
    }
    i = ve + 1;
  }
  return -1;
}

function asciiStr(buf, s, e) {
  let out = '';
  for (let i = s; i < e; i++) out += String.fromCharCode(buf[i]);
  return out;
}

function utf8Str(buf, s, e) {
  return xmlUnescape(utf8Dec.decode(buf.subarray(s, e)));
}

// "AB12" at buf[s..e) -> row * 16384 + col - 1, or -1.
function parseRefBytes(buf, s, e) {
  let i = s, c = 0, r = 0;
  while (i < e) {
    const b = buf[i];
    if (b >= 65 && b <= 90) c = c * 26 + (b - 64);
    else if (b >= 97 && b <= 122) c = c * 26 + (b - 96);
    else if (b !== 36) break;
    i++;
  }
  while (i < e) {
    const b = buf[i];
    if (b >= 48 && b <= 57) r = r * 10 + (b - 48);
    else if (b !== 36) break;
    i++;
  }
  if (c < 1 || r < 1) return -1;
  return r * KEY_COLS + (c - 1);
}

function parseIntBytes(buf, s, e) {
  let n = 0;
  for (let i = s; i < e; i++) {
    const b = buf[i];
    if (b < 48 || b > 57) break;
    n = n * 10 + (b - 48);
  }
  return n;
}

// Reference of the cell whose tag starts at buf[p] ('<c'), tag attributes end at e:
// row * 16384 + col - 1, or -1. Fast for Excel's own form <c r="AB12" ...>.
function cellRefAt(buf, p, e) {
  if (buf[p + 2] === 32 && buf[p + 3] === 114 && buf[p + 4] === B_EQ && buf[p + 5] === B_QU) {
    let i = p + 6, c = 0, r = 0, b = buf[i];
    while (b >= 65 && b <= 90) { c = c * 26 + (b - 64); b = buf[++i]; }
    while (b >= 48 && b <= 57) { r = r * 10 + (b - 48); b = buf[++i]; }
    if (b === B_QU && c > 0 && r > 0) return r * KEY_COLS + (c - 1);
  }
  const vs = findAttr(buf, p + 2, e, A_R);
  return vs >= 0 ? parseRefBytes(buf, vs, attrEnd) : -1;
}

// Position of the next "</row>" from p, or -1 when it is not in buf yet.
function findRowEnd(buf, p) {
  const n = buf.length;
  let i = p;
  for (;;) {
    i = buf.indexOf(119, i);                      // 'w' is rare inside a row
    if (i < 0 || i + 1 >= n) return -1;
    if (buf[i + 1] === B_GT && buf[i - 1] === 111 && buf[i - 2] === 114 && buf[i - 3] === B_SL && buf[i - 4] === B_LT) return i - 4;
    i++;
  }
}

// Start of the first cell after column c2 in the row part buf[p..re) (re: its "</row>").
function tailStart(buf, p, re, c2) {
  let t = re, q = re;
  for (;;) {
    const i = buf.lastIndexOf(B_LT, q - 1);
    if (i < p) break;
    q = i;
    if (buf[i + 1] !== 99 || !(isWs(buf[i + 2]) || buf[i + 2] === B_GT || buf[i + 2] === B_SL)) continue;
    const gt = buf.indexOf(B_GT, i);
    const k = cellRefAt(buf, i, buf[gt - 1] === B_SL ? gt - 1 : gt);
    if (k < 0 || keyCol(k) <= c2) break;
    t = i;
  }
  return t;
}

//------------------------------------------------------------------------------
//  SheetScanner: splits a sheet XML stream into rows and cells.
//  h.raw(buf, s, e)                       bytes outside rows and cells
//  h.sheetData(buf, s, e, selfClose)      the <sheetData> tag
//  h.rowOpen(buf, s, e, r, selfClose)     -> true: pass the whole row on to raw();
//                                            {c1, c2}: leave out the cells of columns c1..c2
//  h.cell(buf, s, tagEnd, e, r, c, selfClose)
//  h.rowClose(buf, s, e)
//  h.sheetDataEnd(buf, s, e)
//------------------------------------------------------------------------------
class SheetScanner {
  constructor(h) {
    this.h = h;
    this.st = 0;          // 0 before <sheetData>, 1 inside, 2 after
    this.carry = null;
    this.row = 0;
    this.col = 0;
    this.passRow = false;
    this.skipCols = null;   // {c1, c2}: columns of the current row that are left out
  }

  push(chunk) {
    let buf = chunk;
    if (this.carry !== null) {
      buf = new Uint8Array(this.carry.length + chunk.length);
      buf.set(this.carry);
      buf.set(chunk, this.carry.length);
    }
    const p = this.run(buf, false);
    this.carry = p < buf.length ? buf.slice(p) : null;
  }

  finish() {
    if (this.carry === null) return;
    const buf = this.carry;
    this.carry = null;
    const p = this.run(buf, true);
    if (p < buf.length) {
      if (this.st === 1) throw new UserError('ข้อมูลชีตในไฟล์ไม่ครบ (ไฟล์เสีย)');
      this.h.raw(buf, p, buf.length);
    }
  }

  run(buf, final) {
    const h = this.h, n = buf.length;
    let p = 0;
    if (this.st === 0) {
      const i = indexOfSeq(buf, SEQ_SHEETDATA, 0);
      if (i < 0) {
        const keep = final ? n : Math.max(0, n - SEQ_SHEETDATA.length);
        if (keep > 0) h.raw(buf, 0, keep);
        return keep;
      }
      const gt = buf.indexOf(B_GT, i);
      if (gt < 0) {
        if (i > 0) h.raw(buf, 0, i);
        return i;
      }
      if (i > 0) h.raw(buf, 0, i);
      const sc = buf[gt - 1] === B_SL;
      h.sheetData(buf, i, gt + 1, sc);
      p = gt + 1;
      this.st = sc ? 2 : 1;
    }
    while (this.st === 1 && p < n) {
      if (this.passRow) {
        const i = indexOfSeq(buf, SEQ_ROW_END, p);
        if (i < 0) {
          const keep = Math.max(p, n - SEQ_ROW_END.length);
          if (keep > p) h.raw(buf, p, keep);
          return keep;
        }
        if (i > p) h.raw(buf, p, i);
        h.raw(buf, i, i + 6);
        p = i + 6;
        this.passRow = false;
        continue;
      }
      let q = p;
      while (q < n && buf[q] !== B_LT) q++;
      if (q > p) { h.raw(buf, p, q); p = q; }
      if (p + 1 >= n) break;
      const b1 = buf[p + 1];
      if (b1 === 99 && p + 2 < n && (isWs(buf[p + 2]) || buf[p + 2] === B_GT || buf[p + 2] === B_SL)) {
        // <c ...> ... </c>
        const gt = buf.indexOf(B_GT, p + 2);
        if (gt < 0) break;
        const sc = buf[gt - 1] === B_SL;
        let e;
        if (sc) {
          e = gt + 1;
        } else {
          e = -1;
          let i = gt + 1;
          for (;;) {
            i = buf.indexOf(B_LT, i);
            if (i < 0 || i + 3 >= n) break;
            if (buf[i + 1] === B_SL && buf[i + 2] === 99 && buf[i + 3] === B_GT) { e = i + 4; break; }
            i++;
          }
          if (e < 0) break;
        }
        let r = this.row, c;
        const k = cellRefAt(buf, p, sc ? gt - 1 : gt);
        if (k >= 0) { r = keyRow(k); c = keyCol(k); } else c = this.col + 1;
        this.col = c;
        const sk = this.skipCols;
        if (sk !== null && c >= sk.c1 && c <= sk.c2) {
          // the cells of columns c1..c2 are left out: go on with the first cell after them
          const re = findRowEnd(buf, p);
          if (re < 0) break;
          p = tailStart(buf, p, re, sk.c2);
          this.col = sk.c2;
          continue;
        }
        h.cell(buf, p, gt + 1, e, r, c, sc);
        p = e;
      } else if (b1 === 114) {
        // <row ...>
        if (p + 4 >= n) break;
        if (buf[p + 2] !== 111 || buf[p + 3] !== 119) { p = this.other(buf, p); if (p < 0) return -p - 1; continue; }
        const gt = buf.indexOf(B_GT, p + 4);
        if (gt < 0) break;
        const sc = buf[gt - 1] === B_SL;
        const vs = findAttr(buf, p + 4, sc ? gt - 1 : gt, A_R);
        const r = vs >= 0 ? parseIntBytes(buf, vs, attrEnd) : this.row + 1;
        this.row = r;
        this.col = 0;
        const pass = h.rowOpen(buf, p, gt + 1, r, sc);
        p = gt + 1;
        this.skipCols = null;
        if (pass === true) { if (!sc) this.passRow = true; }
        else if (pass) this.skipCols = pass;
      } else if (b1 === B_SL) {
        if (buf[p + 2] === 114) {
          if (p + 6 > n) break;
          h.rowClose(buf, p, p + 6);
          p += 6;
        } else if (buf[p + 2] === 115) {
          if (p + 12 > n) break;
          h.sheetDataEnd(buf, p, p + 12);
          p += 12;
          this.st = 2;
        } else {
          p = this.other(buf, p);
          if (p < 0) return -p - 1;
        }
      } else {
        p = this.other(buf, p);
        if (p < 0) return -p - 1;
      }
    }
    if (this.st === 2 && p < n) {
      h.raw(buf, p, n);
      return n;
    }
    return p;
  }

  // Any other element inside sheetData: passed on as raw bytes. Returns -(p + 1) when incomplete.
  other(buf, p) {
    const gt = buf.indexOf(B_GT, p);
    if (gt < 0) return -(p + 1);
    this.h.raw(buf, p, gt + 1);
    return gt + 1;
  }
}

//------------------------------------------------------------------------------
//  Loading the values and formulas of a sheet into the model
//------------------------------------------------------------------------------
class LoadHandler {
  constructor(sheet, sst, date1904) {
    this.sh = sheet;
    this.sst = sst;
    this.date1904 = date1904;
    this.skip = sheet.skip;
  }
  raw() {}
  sheetData() {}
  rowOpen(buf, s, e, r) {
    if (r > this.sh.rows) this.sh.rows = r;
    return this.skip !== null && r >= this.skip.r1 ? this.skip : false;
  }
  rowClose() {}
  sheetDataEnd() {}

  cell(buf, s, tagEnd, e, r, c, sc) {
    if (sc) return;
    const sk = this.skip;
    if (sk !== null && r >= sk.r1 && c >= sk.c1 && c <= sk.c2) return;
    const sh = this.sh;
    let t = 0;          // 0 number, 1 shared string, 2 str, 3 inlineStr, 4 bool, 5 error, 6 date
    const tv = findAttr(buf, s + 2, tagEnd - 1, A_T);
    if (tv >= 0 && tv < attrEnd) {
      const b0 = buf[tv], b1 = buf[tv + 1];
      if (b0 === 115) t = b1 === 116 ? 2 : 1;
      else if (b0 === 105) t = 3;
      else if (b0 === 98) t = 4;
      else if (b0 === 101) t = 5;
      else if (b0 === 100) t = 6;
    }
    const end = e - 4;
    let i = tagEnd, vS = -1, vE = -1, isS = -1, isE = -1, hasF = false;
    while (i < end) {
      i = buf.indexOf(B_LT, i);
      if (i < 0 || i >= end) break;
      const b1 = buf[i + 1], b2 = buf[i + 2];
      const gt = buf.indexOf(B_GT, i);
      if (b1 === 118 && (b2 === B_GT || isWs(b2) || b2 === B_SL)) {
        if (buf[gt - 1] === B_SL) { vS = vE = gt + 1; i = gt + 1; continue; }
        vS = gt + 1;
        vE = buf.indexOf(B_LT, vS);
        i = vE + 4;
      } else if (b1 === 102 && (b2 === B_GT || isWs(b2) || b2 === B_SL)) {
        hasF = true;
        const fsc = buf[gt - 1] === B_SL;
        const tagE = fsc ? gt - 1 : gt;
        let kind = 'n';
        const ft = findAttr(buf, i + 2, tagE, A_T);
        if (ft >= 0) {
          const k0 = buf[ft];
          kind = k0 === 115 ? 'shared' : k0 === 97 ? 'array' : k0 === 110 ? 'n' : 'other';
        }
        let si = -1;
        if (kind === 'shared') {
          const sv = findAttr(buf, i + 2, tagE, A_SI);
          if (sv >= 0) si = parseIntBytes(buf, sv, attrEnd);
        }
        let text = '';
        if (fsc) {
          i = gt + 1;
        } else {
          const tE = buf.indexOf(B_LT, gt + 1);
          text = utf8Str(buf, gt + 1, tE);
          i = tE + 4;
        }
        const rec = { kind, text, si };
        if (kind === 'shared' && text.length > 0) sh.shared.set(si, { text, r, c });
        sh.formulas.set(cellKey(r, c), rec);
      } else if (b1 === 105 && b2 === 115) {
        isS = gt + 1;
        isE = indexOfSeq(buf, SEQ_IS_END, isS);
        if (isE < 0) isE = end;
        i = isE + 5;
      } else {
        // extLst or anything else: skip its content
        if (buf[gt - 1] === B_SL) { i = gt + 1; continue; }
        const nameEnd = (() => { let k = i + 1; while (k < gt && !isWs(buf[k])) k++; return k; })();
        const close = bytesOf('</' + asciiStr(buf, i + 1, nameEnd) + '>');
        const ce = indexOfSeq(buf, close, gt + 1);
        i = ce < 0 || ce > end ? end : ce + close.length;
      }
    }
    let v = null;
    switch (t) {
      case 0:
        if (vE > vS) {
          const x = Number(asciiStr(buf, vS, vE));
          v = x === x ? x : null;
        }
        break;
      case 1:
        if (vE > vS) {
          const s2 = this.sst[parseIntBytes(buf, vS, vE)];
          v = s2 === undefined ? '' : s2;
        }
        break;
      case 2:
        if (vS >= 0) v = xUnescape(utf8Str(buf, vS, vE));
        break;
      case 3:
        if (isS >= 0) v = parseSharedStrings('<si>' + utf8Dec.decode(buf.subarray(isS, isE)) + '</si>')[0];
        else if (vS >= 0) v = xUnescape(utf8Str(buf, vS, vE));
        break;
      case 4:
        if (vE > vS) v = buf[vS] === 49 || buf[vS] === 116;
        break;
      case 5:
        if (vE > vS) v = errOf(asciiStr(buf, vS, vE));
        break;
      case 6:
        if (vE > vS) {
          const ms = Date.parse(asciiStr(buf, vS, vE));
          if (ms === ms) v = ms / 86400000 + (this.date1904 ? 24107 : 25569);
        }
        break;
    }
    if (v !== null) sh.put(r, c, v);
    else if (hasF) {
      if (r > sh.maxRow) sh.maxRow = r;
      if (c > sh.maxCol) sh.maxCol = c;
    }
  }
}

async function scanEntry(zip, entry, handler, onBytes) {
  const sc = new SheetScanner(handler);
  for await (const chunk of bigChunks(await zipStream(zip, entry), 1 << 20)) {
    sc.push(chunk);
    if (onBytes) await onBytes(chunk.length);
  }
  sc.finish();
}

//------------------------------------------------------------------------------
//  The whole workbook
//------------------------------------------------------------------------------
// opts.skipFor(sheetName) -> {r1, c1, c2} or null: cells not stored.
// progress({stage, text, done, total}) may return a promise.
async function loadWorkbook(file, opts, progress) {
  opts = opts || {};
  const zip = await zipOpen(file);
  const wb = new Workbook(file, zip);
  let wbPath = 'xl/workbook.xml';
  for (const rel of await readRels(zip, '')) {
    if (/\/officeDocument$/.test(rel.type)) wbPath = rel.target;
  }
  const wbXml = /\.xml$/i.test(wbPath) ? await readXmlText(zip, wbPath) : null;
  if (wbXml === null) {
    if (/\.bin$/i.test(wbPath)) throw new UserError('ไฟล์ .xlsb (Excel Binary) อ่านไม่ได้ กรุณาบันทึกเป็น .xlsm ก่อน');
    throw new UserError('ไม่พบข้อมูล workbook ในไฟล์ (ไม่ใช่ไฟล์ Excel?)');
  }
  const rels = await readRels(zip, wbPath);
  const relById = new Map(rels.map((r) => [r.id, r]));
  const pr = /<workbookPr\b([^>]*?)\/?>/.exec(wbXml);
  if (pr) {
    const a = parseAttrs(pr[1]);
    wb.date1904 = a.date1904 === '1' || a.date1904 === 'true';
  }

  let sst = [];
  const sstRel = rels.find((r) => /\/sharedStrings$/.test(r.type));
  if (sstRel) {
    if (progress) await progress({ stage: 'read', text: 'อ่านข้อความในไฟล์ (shared strings)...', done: 0, total: 1 });
    const xml = await readXmlText(zip, sstRel.target);
    if (xml !== null) sst = parseSharedStrings(xml);
  }

  const reSheet = /<sheet\b([^>]*?)\/?>/g;
  let m;
  const localNames = [];
  while ((m = reSheet.exec(wbXml)) !== null) {
    const a = parseAttrs(m[1]);
    let rid = null;
    for (const k of Object.keys(a)) if (k !== 'sheetId' && /(^|:)id$/.test(k)) rid = a[k];
    const rel = relById.get(rid);
    localNames.push(a.name);
    if (!rel || !/\/worksheet$/.test(rel.type)) continue;
    const entry = zipEntry(zip, rel.target);
    if (entry === null) continue;
    const sh = wb.addSheet(a.name);
    sh.state = a.state || 'visible';
    sh.path = entry.name;
    sh.entry = entry;
  }
  const reName = /<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g;
  while ((m = reName.exec(wbXml)) !== null) {
    const a = parseAttrs(m[1]);
    const local = a.localSheetId !== undefined ? parseInt(a.localSheetId, 10) : -1;
    wb.names.push({ name: a.name, text: xmlUnescape(m[2]), local, hidden: a.hidden === '1' });
  }

  let total = 0, done = 0;
  for (const sh of wb.sheets) total += sh.entry.usize;
  for (const sh of wb.sheets) {
    sh.skip = opts.skipFor ? opts.skipFor(wb, sh) : null;
    const text = 'อ่านชีต ' + sh.name + '...';
    if (progress) await progress({ stage: 'read', text, done, total });
    let last = 0;
    await scanEntry(zip, sh.entry, new LoadHandler(sh, sst, wb.date1904), async (n) => {
      done += n;
      last += n;
      if (progress && last > 4000000) { last = 0; await progress({ stage: 'read', text, done, total }); }
    });
  }
  return wb;
}
