'use strict';
//==============================================================================
//  Formula text: A1 <-> R1C1
//  The engine (engine.js) compiles R1C1 formulas, as the VBA engine did. A
//  formula in a cell of the file is A1 text; it is turned into R1C1 relative to
//  its cell, so every copy of a formula down a column is the same rule.
//==============================================================================

const RE_ERRLIT = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|UNKNOWN!|BUSY!)/i;
const RE_NUMLIT = /^(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/;

function isNameCh(ch) {
  if (ch === undefined || ch === '') return false;
  const c = ch.charCodeAt(0);
  return (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || (c >= 48 && c <= 57) || ch === '_' || ch === '.' ||
    ch === '\\' || c > 127;
}

function quoteSheet(nm) {
  return "'" + nm.replace(/'/g, "''") + "'";
}

// One part of an A1 reference at f[i]: cell ($A$1), column ($A) or row ($1).
function readA1Part(f, i) {
  let j = i, colAbs = false, rowAbs = false, letters = '', digits = '';
  if (f[j] === '$') { colAbs = true; j++; }
  while (j < f.length && /[A-Za-z]/.test(f[j]) && letters.length < 4) letters += f[j++];
  if (letters.length === 0) {
    // a row: $1 or 1
    rowAbs = colAbs; colAbs = false;
    while (j < f.length && f[j] >= '0' && f[j] <= '9') digits += f[j++];
    if (digits.length === 0) return null;
    const row = +digits;
    if (row < 1 || row > 1048576) return null;
    return { kind: 'row', row, rowAbs, end: j };
  }
  if (letters.length > 3) return null;
  const col = colNumber(letters.toUpperCase());
  if (col < 1 || col > 16384) return null;
  if (f[j] === '$') { rowAbs = true; j++; }
  while (j < f.length && f[j] >= '0' && f[j] <= '9') digits += f[j++];
  if (digits.length === 0) {
    if (rowAbs) return null;
    return { kind: 'col', col, colAbs, end: j };
  }
  const row = +digits;
  if (row < 1 || row > 1048576) return null;
  return { kind: 'cell', col, colAbs, row, rowAbs, end: j };
}

function rcPart(letter, v, abs, base) {
  if (abs) return letter + v;
  const d = v - base;
  return d === 0 ? letter : letter + '[' + d + ']';
}

// A1 reference at f[i] -> {text (R1C1), end} or null.
function readA1Ref(f, i, r0, c0) {
  const p1 = readA1Part(f, i);
  if (p1 === null) return null;
  let p2 = null, end = p1.end;
  if (f[end] === ':') {
    p2 = readA1Part(f, end + 1);
    if (p2 === null || p2.kind !== p1.kind) {
      if (p1.kind !== 'cell') return null;
      p2 = null;
    } else {
      end = p2.end;
    }
  }
  if (p2 === null && p1.kind !== 'cell') return null;
  const nx = f[end];
  if (isNameCh(nx) || nx === '(' || nx === '!' || nx === '$') return null;
  const one = (p) => {
    if (p.kind === 'cell') return rcPart('R', p.row, p.rowAbs, r0) + rcPart('C', p.col, p.colAbs, c0);
    if (p.kind === 'col') return rcPart('C', p.col, p.colAbs, c0);
    return rcPart('R', p.row, p.rowAbs, r0);
  };
  const t1 = one(p1);
  if (p2 === null) return { text: t1, end };
  const t2 = one(p2);
  return { text: p1.kind !== 'cell' && t1 === t2 ? t1 : t1 + ':' + t2, end };
}

// A1 formula text (without "=") of the cell at row r0, column c0 -> "=" R1C1 text.
function a1ToR1C1(f, r0, c0) {
  let out = '=', i = 0;
  const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') {
      let j = i + 1;
      for (;;) {
        if (j >= n) return out + f.slice(i);
        if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; }
        j++;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      let j = i + 1, nm = '';
      for (;;) {
        if (j >= n) return out + f.slice(i);
        if (f[j] === "'") { if (f[j + 1] === "'") { nm += "'"; j += 2; continue; } break; }
        nm += f[j++];
      }
      j++;
      if (f[j] !== '!') { out += f.slice(i, j); i = j; continue; }
      const ref = readA1Ref(f, j + 1, r0, c0);
      if (ref === null) { out += f.slice(i, j + 1); i = j + 1; continue; }
      out += quoteSheet(nm) + '!' + ref.text;
      i = ref.end;
      continue;
    }
    if (ch === '#') {
      const m = RE_ERRLIT.exec(f.slice(i, i + 16));
      if (m) { out += m[0]; i += m[0].length; continue; }
      out += ch; i++;
      continue;
    }
    if (ch >= '0' && ch <= '9' || ch === '.') {
      const ref = readA1Ref(f, i, r0, c0);
      if (ref !== null) { out += ref.text; i = ref.end; continue; }
      const m = RE_NUMLIT.exec(f.slice(i, i + 40));
      if (m) { out += m[0]; i += m[0].length; continue; }
      out += ch; i++;
      continue;
    }
    if (ch === '$' || isNameCh(ch)) {
      let j = i;
      while (j < n && (isNameCh(f[j]) || f[j] === '$' || f[j] === '?')) j++;
      const w = f.slice(i, j);
      if (f[j] === '!') {
        const ref = readA1Ref(f, j + 1, r0, c0);
        if (ref === null) { out += f.slice(i, j + 1); i = j + 1; continue; }
        out += quoteSheet(w) + '!' + ref.text;
        i = ref.end;
        continue;
      }
      if (f[j] === '(') { out += w; i = j; continue; }
      const ref = readA1Ref(f, i, r0, c0);
      if (ref !== null) { out += ref.text; i = ref.end; continue; }
      out += w;
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

// R1C1 reference at f[i] -> {text (A1), end} or null (for showing a rule to the user).
function readR1C1Ref(f, i, r0, c0) {
  const part = (j) => {
    let hasR = false, hasC = false, r = 0, c = 0, rRel = true, cRel = true;
    const off = (k) => {
      if (f[k] === '[') {
        const m = /^\[([+-]?[0-9]+)\]/.exec(f.slice(k, k + 12));
        if (!m) return null;
        return { v: +m[1], rel: true, end: k + m[0].length };
      }
      const m = /^[0-9]+/.exec(f.slice(k, k + 10));
      if (m) return { v: +m[0], rel: false, end: k + m[0].length };
      return { v: 0, rel: true, end: k };
    };
    if (f[j] === 'R') {
      const o = off(j + 1);
      if (o === null) return null;
      hasR = true; r = o.v; rRel = o.rel; j = o.end;
    }
    if (f[j] === 'C') {
      const o = off(j + 1);
      if (o === null) return null;
      hasC = true; c = o.v; cRel = o.rel; j = o.end;
    }
    if (!hasR && !hasC) return null;
    return { hasR, hasC, r, c, rRel, cRel, end: j };
  };
  const p1 = part(i);
  if (p1 === null) return null;
  let p2 = null, end = p1.end;
  if (f[end] === ':') {
    p2 = part(end + 1);
    if (p2 === null || p2.hasR !== p1.hasR || p2.hasC !== p1.hasC) return null;
    end = p2.end;
  }
  const nx = f[end];
  if (isNameCh(nx) || nx === '(' || nx === '!') return null;
  const one = (p) => {
    const rr = p.rRel ? r0 + p.r : p.r, cc = p.cRel ? c0 + p.c : p.c;
    if (rr < 1 || cc < 1 || rr > 1048576 || cc > 16384) return '#REF!';
    const cs = p.hasC ? (p.cRel ? '' : '$') + colLetter(cc) : '';
    const rs = p.hasR ? (p.rRel ? '' : '$') + rr : '';
    return cs + rs;
  };
  let t = one(p1);
  if (!p1.hasR || !p1.hasC) t = t + ':' + one(p2 || p1);
  else if (p2 !== null) t = t + ':' + one(p2);
  return { text: t, end };
}

// "=" R1C1 formula of the cell at r0, c0 -> "=" A1 formula.
function r1c1ToA1(f, r0, c0) {
  let out = '', i = 0;
  const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (ch === '"') {
      let j = i + 1;
      for (;;) {
        if (j >= n) return out + f.slice(i);
        if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; }
        j++;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      const k = f.indexOf("'!", i + 1);
      if (k < 0) return out + f.slice(i);
      out += f.slice(i, k + 2);
      i = k + 2;
      const ref = readR1C1Ref(f, i, r0, c0);
      if (ref !== null) { out += ref.text; i = ref.end; }
      continue;
    }
    if (isNameCh(ch)) {
      if (ch === 'R' || ch === 'C') {
        const ref = readR1C1Ref(f, i, r0, c0);
        if (ref !== null) { out += ref.text; i = ref.end; continue; }
      }
      let j = i;
      while (j < n && isNameCh(f[j])) j++;
      out += f.slice(i, j);
      i = j;
      if (f[i] === '!') {
        out += '!';
        i++;
        const ref = readR1C1Ref(f, i, r0, c0);
        if (ref !== null) { out += ref.text; i = ref.end; }
      }
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}
