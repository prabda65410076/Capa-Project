'use strict';
//==============================================================================
//  Writing the workbook back. Only the sheets with changed cells are rewritten,
//  cell by cell as a stream; every other part of the file (formats, charts, the
//  VBA project, ...) is copied byte for byte. A changed cell keeps its style and
//  its formula (FORMULAS mode: only the value Excel shows is updated).
//==============================================================================

class ByteWriter {
  constructor() {
    this.size = 1 << 20;
    this.buf = new Uint8Array(this.size);
    this.n = 0;
    this.out = [];
  }
  room(k) {
    if (this.n + k > this.buf.length) {
      this.flush();
      if (k > this.buf.length) this.buf = new Uint8Array(k);
    }
  }
  flush() {
    if (this.n > 0) {
      this.out.push(this.buf.subarray(0, this.n));
      this.buf = new Uint8Array(this.size);
      this.n = 0;
    }
  }
  bytes(src, s, e) {
    const k = e - s;
    if (k <= 0) return;
    if (k > this.size) { this.flush(); this.out.push(src.slice(s, e)); return; }
    this.room(k);
    this.buf.set(src.subarray(s, e), this.n);
    this.n += k;
  }
  ascii(str) {
    const k = str.length;
    this.room(k);
    const b = this.buf;
    let n = this.n;
    for (let i = 0; i < k; i++) b[n++] = str.charCodeAt(i);
    this.n = n;
  }
  text(str) {
    let ascii = true;
    for (let i = 0; i < str.length; i++) if (str.charCodeAt(i) > 127) { ascii = false; break; }
    if (ascii) { this.ascii(str); return; }
    const enc = utf8Enc.encode(str);
    this.room(enc.length);
    this.buf.set(enc, this.n);
    this.n += enc.length;
  }
  take() {
    this.flush();
    const o = this.out;
    this.out = [];
    return o;
  }
}

// Text for a <t> / <v> element.
function xmlText(s) {
  let t = s;
  if (t.indexOf('_x') >= 0) t = t.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_');
  t = t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return t.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g,
    (ch) => '_x' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_');
}

function numXml(x) {
  if (x === 0) return '0';
  const s = String(x);
  return s.indexOf('e') >= 0 ? s.replace('e', 'E') : s;
}

const DROP = {};   // the cell is left out of the file

class PatchHandler {
  // The computed region of the sheet (sheet 1-2 CH:AGS, vt) is written from the computed
  // values: the first row of the region is rewritten cell by cell (and gives the style of
  // each column), the other rows get new cells with those styles without reading the old ones.
  // opts.detail === false: the region keeps only the columns vt.keep(c) (the Plan columns).
  constructor(sheet, opts, out) {
    this.sh = sheet;
    this.out = out;
    this.vt = sheet.virt;
    this.detail = !opts || opts.detail !== false;
    const byRow = new Map();
    for (const k of sheet.dirty) {
      const r = keyRow(k);
      let a = byRow.get(r);
      if (a === undefined) byRow.set(r, (a = []));
      a.push(keyCol(k));
    }
    for (const a of byRow.values()) a.sort((x, y) => x - y);
    this.byRow = byRow;
    this.pending = Array.from(byRow.keys()).sort((x, y) => x - y);
    this.pi = 0;
    this.vr1 = this.vt ? this.vt.r1 : 1;
    this.vr2 = this.vt ? this.vt.r2 : 0;
    this.nv = this.vt ? this.vt.r1 : Infinity;
    this.row = 0;
    this.cols = null;
    this.ci = 0;
    this.vc = Infinity;
    this.vc2 = -1;
    this.openRow = 0;       // a new row whose <row> tag is written with its first cell
    this.cells = 0;         // cells written
    this.light = this.vt !== null && !this.detail && typeof this.vt.keep === 'function';
    this.styles = [];       // style (s) of the columns of the region, from its first row
    this.stylesSeen = false;
  }

  // ---- rows and cells to write ---------------------------------------------------
  nextMissing() {
    const a = this.pi < this.pending.length ? this.pending[this.pi] : Infinity;
    const b = this.nv <= this.vr2 ? this.nv : Infinity;
    return a < b ? a : b;
  }
  consumeRow(r) {
    while (this.pi < this.pending.length && this.pending[this.pi] <= r) this.pi++;
    if (this.nv <= r) this.nv = r + 1;
  }
  rowHasPatch(r) {
    return this.byRow.has(r) || (r >= this.vr1 && r <= this.vr2);
  }
  beginRow(r) {
    this.row = r;
    this.cols = this.byRow.get(r) || null;
    this.ci = 0;
    if (this.vt !== null && r >= this.vr1 && r <= this.vr2) { this.vc = this.vt.c1; this.vc2 = this.vt.c2; }
    else { this.vc = Infinity; this.vc2 = -1; }
  }
  nextCol() {
    const a = this.cols !== null && this.ci < this.cols.length ? this.cols[this.ci] : Infinity;
    const b = this.vc <= this.vc2 ? this.vc : Infinity;
    return a < b ? a : b;
  }
  takeCol(c) {
    if (this.cols !== null && this.ci < this.cols.length && this.cols[this.ci] === c) this.ci++;
    if (this.vc === c) this.vc++;
  }
  value(r, c) {
    const vt = this.vt;
    if (vt !== null && r >= vt.r1 && r <= vt.r2 && c >= vt.c1 && c <= vt.c2) {
      if (!this.detail && vt.keep && !vt.keep(c)) return DROP;
      return vt.get(r, c);
    }
    return this.sh.getStored(r, c);
  }

  // ---- output ----------------------------------------------------------------------
  // Rest of a cell after its attributes (no t, no closing '>').
  body(v, fb, fs, fe) {
    const out = this.out;
    this.cells++;
    if (v === null || v === DROP) {
      if (fb !== null) { out.ascii('>'); out.bytes(fb, fs, fe); out.ascii('</c>'); } else out.ascii('/>');
      return;
    }
    if (typeof v === 'number') {
      out.ascii('>');
      if (fb !== null) out.bytes(fb, fs, fe);
      out.ascii('<v>'); out.ascii(numXml(v)); out.ascii('</v></c>');
    } else if (typeof v === 'string') {
      if (fb !== null) {
        out.ascii(' t="str">'); out.bytes(fb, fs, fe);
        out.ascii('<v>'); out.text(xmlText(v)); out.ascii('</v></c>');
      } else {
        out.ascii(' t="inlineStr"><is><t xml:space="preserve">'); out.text(xmlText(v)); out.ascii('</t></is></c>');
      }
    } else if (typeof v === 'boolean') {
      out.ascii(' t="b">');
      if (fb !== null) out.bytes(fb, fs, fe);
      out.ascii(v ? '<v>1</v></c>' : '<v>0</v></c>');
    } else {
      out.ascii(' t="e">');
      if (fb !== null) out.bytes(fb, fs, fe);
      out.ascii('<v>'); out.text(xmlText(v.text)); out.ascii('</v></c>');
    }
  }

  newCell(r, c) {
    const v = this.value(r, c);
    if (v === null || v === DROP) return;
    if (this.openRow !== 0) {
      this.out.ascii('<row r="' + this.openRow + '">');
      this.openRow = 0;
    }
    const st = this.styles[c] !== undefined ? ' s="' + this.styles[c] + '"' : '';
    this.out.ascii('<c r="' + colLetter(c) + r + '"' + st);
    this.body(v, null, 0, 0);
  }

  cellsBefore(limit) {
    for (;;) {
      const c = this.nextCol();
      if (c >= limit) return;
      if (c === this.vc) {
        // a run of computed cells
        const d = this.cols !== null && this.ci < this.cols.length ? this.cols[this.ci] : Infinity;
        const end = Math.min(limit - 1, this.vc2, d - 1);
        if (end >= c) {
          this.virtRun(this.row, c, end);
          this.vc = end + 1;
          continue;
        }
      }
      this.takeCol(c);
      this.newCell(this.row, c);
    }
  }

  // New cells of the computed region for columns a..b of row r.
  virtRun(r, a, b) {
    const vt = this.vt, out = this.out, light = this.light, styles = this.styles, rs = String(r);
    for (let c = a; c <= b; c++) {
      if (light && !vt.keep(c)) { c = vt.nextKeep(c) - 1; continue; }
      const v = vt.get(r, c);
      if (v === null) continue;
      if (this.openRow !== 0) { out.ascii('<row r="' + this.openRow + '">'); this.openRow = 0; }
      const st = styles[c];
      if (typeof v === 'number') {
        out.ascii(st !== undefined ? '<c r="' + colLetter(c) + rs + '" s="' + st + '"><v>' + numXml(v) + '</v></c>' :
          '<c r="' + colLetter(c) + rs + '"><v>' + numXml(v) + '</v></c>');
        this.cells++;
      } else {
        out.ascii('<c r="' + colLetter(c) + rs + '"' + (st !== undefined ? ' s="' + st + '"' : ''));
        this.body(v, null, 0, 0);
      }
    }
  }

  newRow(r) {
    this.beginRow(r);
    this.openRow = r;
    this.cellsBefore(Infinity);
    if (this.openRow === 0) this.out.ascii('</row>');
    this.openRow = 0;
  }

  rowsBefore(limit) {
    for (;;) {
      const r = this.nextMissing();
      if (r >= limit) return;
      this.consumeRow(r);
      this.newRow(r);
    }
  }

  rewriteCell(buf, s, tagEnd, e, r, c, sc) {
    const v = this.value(r, c);
    const aEnd = sc ? tagEnd - 2 : tagEnd - 1;
    const vt = this.vt;
    if (vt !== null && c >= vt.c1 && c <= vt.c2 && r >= vt.r1 && r <= vt.r2) {
      if (r === vt.r1) {
        const sv = findAttr(buf, s + 2, aEnd, A_S);
        if (sv >= 0) this.styles[c] = asciiStr(buf, sv, attrEnd);
      }
      // fast path for the computed cells: <c r=".." s=".."> + <v> or nothing, no type, no formula
      if (typeof v === 'number' && (sc || (buf[tagEnd] === B_LT && (buf[tagEnd + 1] === 118 || buf[tagEnd + 1] === B_SL)))) {
        let hasT = false;
        for (let i = s + 2; i < aEnd - 2; i++) {
          if (buf[i] === 116 && buf[i + 1] === B_EQ && isWs(buf[i - 1])) { hasT = true; break; }
        }
        if (!hasT) {
          const out = this.out;
          out.bytes(buf, s, aEnd);
          out.ascii('><v>'); out.ascii(numXml(v)); out.ascii('</v></c>');
          this.cells++;
          return;
        }
      }
    }
    // the formula, if any, stays as it is
    let fS = -1, fE = -1;
    if (!sc) {
      let i = tagEnd;
      const end = e - 4;
      while (i < end) {
        i = buf.indexOf(B_LT, i);
        if (i < 0 || i >= end) break;
        const b1 = buf[i + 1], b2 = buf[i + 2];
        if (b1 === 102 && (b2 === B_GT || isWs(b2) || b2 === B_SL)) {
          const gt = buf.indexOf(B_GT, i);
          fS = i;
          if (buf[gt - 1] === B_SL) fE = gt + 1;
          else fE = buf.indexOf(B_LT, gt + 1) + 4;
          break;
        }
        i++;
      }
    }
    if (v === DROP && fS < 0) return;
    const out = this.out;
    const tv = findAttr(buf, s + 2, aEnd, A_T);
    if (tv >= 0) {
      const ts = attrNameStart, te = attrEnd + 1;
      out.bytes(buf, s, ts);
      out.bytes(buf, te, aEnd);
    } else {
      out.bytes(buf, s, aEnd);
    }
    if (fS >= 0) this.body(v === DROP ? null : v, buf, fS, fE);
    else this.body(v, null, 0, 0);
  }

  // ---- scanner callbacks --------------------------------------------------------------
  raw(buf, s, e) { this.out.bytes(buf, s, e); }

  sheetData(buf, s, e, sc) {
    if (!sc) { this.out.bytes(buf, s, e); return; }
    this.out.bytes(buf, s, e - 2);
    this.out.ascii('>');
    this.rowsBefore(Infinity);
    this.out.ascii('</sheetData>');
  }

  rowOpen(buf, s, e, r, sc) {
    this.rowsBefore(r);
    this.consumeRow(r);
    if (!this.rowHasPatch(r)) { this.out.bytes(buf, s, e); return true; }
    this.beginRow(r);
    // after the first row of the region the old cells of the region are not read at all
    let skip = false;
    if (this.vt !== null && r >= this.vr1 && r <= this.vr2) {
      if (this.stylesSeen) skip = true; else this.stylesSeen = true;
    }
    const aEnd = sc ? e - 2 : e - 1;
    const vs = findAttr(buf, s + 4, aEnd, A_SPANS);
    if (vs >= 0) {
      const ns = attrNameStart, ne = attrEnd + 1;
      this.out.bytes(buf, s, ns);
      this.out.bytes(buf, ne, aEnd);
    } else {
      this.out.bytes(buf, s, aEnd);
    }
    this.out.ascii('>');
    if (sc) {
      this.cellsBefore(Infinity);
      this.out.ascii('</row>');
    }
    return skip ? this.vt : false;
  }

  cell(buf, s, tagEnd, e, r, c, sc) {
    this.cellsBefore(c);
    if (this.nextCol() === c) {
      this.takeCol(c);
      this.rewriteCell(buf, s, tagEnd, e, r, c, sc);
    } else {
      this.out.bytes(buf, s, e);
    }
  }

  rowClose(buf, s, e) {
    this.cellsBefore(Infinity);
    this.out.bytes(buf, s, e);
  }

  sheetDataEnd(buf, s, e) {
    this.rowsBefore(Infinity);
    this.out.bytes(buf, s, e);
  }
}

// The workbook with every change of the recalculation, as a Blob.
async function exportWorkbook(wb, opts, progress) {
  const repl = new Map();
  const targets = wb.sheets.filter((sh) => sh.dirty.size > 0 || sh.virt !== null);
  let total = 0, done = 0;
  for (const sh of targets) total += sh.entry.usize;
  for (const sh of targets) {
    const text = 'เขียนชีต ' + sh.name + '...';
    if (progress) await progress({ stage: 'write', text, done, total });
    const res = await deflateChunks(async (emit) => {
      const out = new ByteWriter();
      const h = new PatchHandler(sh, opts, out);
      const sc = new SheetScanner(h);
      let last = 0;
      for await (const chunk of bigChunks(await zipStream(wb.zip, sh.entry), 1 << 20)) {
        sc.push(chunk);
        for (const o of out.take()) await emit(o);
        done += chunk.length;
        last += chunk.length;
        if (progress && last > 4000000) { last = 0; await progress({ stage: 'write', text, done, total }); }
      }
      sc.finish();
      for (const o of out.take()) await emit(o);
    });
    repl.set(sh.entry.name, res);
  }
  if (progress) await progress({ stage: 'write', text: 'รวมไฟล์...', done: total, total });
  const mime = /\.xlsm$/i.test(wb.fileName) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' :
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  return zipWrite(wb.zip, repl, mime);
}
