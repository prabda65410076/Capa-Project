'use strict';
//==============================================================================
//  Workbook model: every sheet keeps its values column by column
//  (numbers in a Float64Array, other values in a Map), like Range.Value2.
//==============================================================================

const KEY_COLS = 16384;
function cellKey(r, c) { return r * KEY_COLS + (c - 1); }
function keyRow(k) { return Math.floor(k / KEY_COLS); }
function keyCol(k) { return (k % KEY_COLS) + 1; }

class Col {
  constructor(n) {
    this.num = new Float64Array(n).fill(NaN);   // NaN: not a number (blank or in oth)
    this.oth = null;                            // row -> text / boolean / error
    this.last = 0;                              // highest row ever written
  }
}

class Sheet {
  constructor(wb, index, name) {
    this.wb = wb;
    this.index = index;
    this.name = name;
    this.state = 'visible';     // visible / hidden / veryHidden
    this.path = '';             // part in the zip file
    this.cols = [];             // cols[c], c from 1
    this.maxRow = 0;
    this.maxCol = 0;
    this.formulas = new Map();  // cell key -> {kind: 'n' | 'shared' | 'array' | 'other', text, si}
    this.shared = new Map();    // shared formula si -> {text, r, c}
    this.dirty = new Set();     // cell keys changed by the recalculation (written back on export)
    this.orig = new Map();      // cell key -> value in the file, for cells changed by the recalculation
    this.virt = null;           // {r1, r2, c1, c2, get(r, c)}: values computed when read
    this.skip = null;           // {r1, c1, c2}: cells not stored while loading (replaced by virt)
    this.rows = 0;              // rows read from the file (for messages)
  }

  get(r, c) {
    const vt = this.virt;
    if (vt !== null && r >= vt.r1 && r <= vt.r2 && c >= vt.c1 && c <= vt.c2) return vt.get(r, c);
    const col = this.cols[c];
    if (col === undefined || r >= col.num.length) return null;
    const x = col.num[r];
    if (x === x) return x;
    if (col.oth === null) return null;
    const v = col.oth.get(r);
    return v === undefined ? null : v;
  }

  // Value without the computed region (what the file holds / the recalculation wrote).
  getStored(r, c) {
    const col = this.cols[c];
    if (col === undefined || r >= col.num.length) return null;
    const x = col.num[r];
    if (x === x) return x;
    if (col.oth === null) return null;
    const v = col.oth.get(r);
    return v === undefined ? null : v;
  }

  put(r, c, v) {
    let col = this.cols[c];
    if (col === undefined) {
      if (v === null) return;
      col = this.cols[c] = new Col(Math.max(64, r + 1));
    }
    if (r >= col.num.length) {
      if (v === null) return;
      let n = col.num.length * 2;
      while (n <= r) n *= 2;
      const a = new Float64Array(n).fill(NaN);
      a.set(col.num);
      col.num = a;
    }
    if (typeof v === 'number') {
      col.num[r] = v;
      if (col.oth !== null) col.oth.delete(r);
    } else {
      col.num[r] = NaN;
      if (v === null) {
        if (col.oth !== null) col.oth.delete(r);
        return;
      }
      if (col.oth === null) col.oth = new Map();
      col.oth.set(r, v);
    }
    if (r > col.last) col.last = r;
    if (r > this.maxRow) this.maxRow = r;
    if (c > this.maxCol) this.maxCol = c;
  }

  // Changes a cell (it is written back into the file on export).
  set(r, c, v) {
    const k = cellKey(r, c);
    if (!this.orig.has(k)) this.orig.set(k, this.getStored(r, c));
    this.put(r, c, v);
    this.dirty.add(k);
  }

  // Last row of column c that holds a value (End(xlUp) from the bottom; 1 when the column is empty).
  lastRowIn(c) {
    const col = this.cols[c];
    if (col === undefined) return 1;
    for (let r = Math.min(col.last, col.num.length - 1); r >= 1; r--) {
      if (col.num[r] === col.num[r]) return r;
      if (col.oth !== null && col.oth.has(r)) return r;
    }
    return 1;
  }

  // Last row of the used range (UsedRange.Row + UsedRange.Rows.Count - 1).
  usedLastRow() {
    return this.maxRow > 0 ? this.maxRow : 1;
  }

  hasFormula(r, c) {
    return this.formulas.has(cellKey(r, c));
  }
}

class Workbook {
  constructor(file, zip) {
    this.file = file;
    this.zip = zip;
    this.sheets = [];        // worksheets in workbook order
    this.byName = new Map(); // upper case name -> sheet
    this.names = [];         // defined names {name, text, local}
    this.date1904 = false;
    this.fileName = file && file.name ? file.name : 'workbook.xlsx';
  }

  addSheet(name) {
    const sh = new Sheet(this, this.sheets.length, name);
    this.sheets.push(sh);
    this.byName.set(name.toUpperCase(), sh);
    return sh;
  }

  // Worksheets("name"): the name ignores case.
  sheet(name) {
    return this.byName.get(String(name).toUpperCase()) || null;
  }

  // First sheet whose name starts with prefix (FindSheet of the macro).
  findPrefix(prefix) {
    for (const sh of this.sheets) if (sh.name.slice(0, prefix.length) === prefix) return sh;
    return null;
  }

  definedName(name) {
    const u = name.toUpperCase();
    for (const n of this.names) if (n.local < 0 && n.name.toUpperCase() === u) return n;
    return null;
  }
}
