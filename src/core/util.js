'use strict';
//==============================================================================
//  Cell values as the VBA macros see them (Range.Value2):
//    blank = null, number, string, boolean, error = XErr
//==============================================================================

class XErr {
  constructor(text) { this.text = text; }
  toString() { return this.text; }
}
const ERR_NULL = new XErr('#NULL!');
const ERR_DIV0 = new XErr('#DIV/0!');
const ERR_VALUE = new XErr('#VALUE!');
const ERR_REF = new XErr('#REF!');
const ERR_NAME = new XErr('#NAME?');
const ERR_NUM = new XErr('#NUM!');
const ERR_NA = new XErr('#N/A');
const errByText = new Map();
for (const e of [ERR_NULL, ERR_DIV0, ERR_VALUE, ERR_REF, ERR_NAME, ERR_NUM, ERR_NA]) errByText.set(e.text, e);

// The error value with this text (#N/A, #SPILL!, ...), always the same object.
function errOf(text) {
  const u = String(text).toUpperCase();
  let e = errByText.get(u);
  if (e === undefined) { e = new XErr(u); errByText.set(u, e); }
  return e;
}

// The seven error literals a formula may hold (IsErrLiteral of the VBA engine).
const ERR_LITERALS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);
function errLiteral(s) {
  const u = s.toUpperCase();
  return ERR_LITERALS.has(u) ? errByText.get(u) : null;
}

// A message for the user (shown as it is, no stack trace).
class UserError extends Error {}

//------------------------------------------------------------------------------
//  Numbers as text
//------------------------------------------------------------------------------
function trimZeros(s) {
  if (s.indexOf('.') < 0) return s;
  let i = s.length;
  while (s.charCodeAt(i - 1) === 48) i--;
  if (s.charCodeAt(i - 1) === 46) i--;
  return s.slice(0, i);
}

// VBA CStr of a Double: 15 significant digits, E notation for very large / small numbers.
function cstr(x) {
  if (x === 0) return '0';
  if (!isFinite(x)) return String(x);
  const ex = x.toExponential(14);
  const k = ex.indexOf('e');
  const exp = +ex.slice(k + 1);
  if (exp < -4 || exp >= 15) {
    const a = Math.abs(exp);
    return trimZeros(ex.slice(0, k)) + 'E' + (exp < 0 ? '-' : '+') + (a < 10 ? '0' : '') + a;
  }
  return trimZeros(x.toFixed(14 - exp));
}

// Key of a number that is the same for numbers equal to 15 significant digits.
function k15(x) {
  return x === 0 ? '0' : x.toPrecision(15);
}

// Excel "&" / MID on a number (NumText of the VBA engine).
function numText(x) {
  if (x === Math.floor(x) && Math.abs(x) < 1e15) return x === 0 ? '0' : x.toFixed(0);
  return cstr(x);
}

//------------------------------------------------------------------------------
//  Text that Excel reads as a number
//------------------------------------------------------------------------------
const RE_NOT_NUMCHAR = /[^0-9 .,+$%()eE-]/;
const RE_DIGIT = /[0-9]/;
// [spaces] [+-] digits [. digits] [e [+-] 1-3 digits] [spaces]   (PlainNum of the VBA engine)
const RE_PLAIN = /^ *[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]{1,3})? *$/;
// "$5", "-$5", "(5)", "1,000", "5%", ...: what Application.Evaluate("--""text""") gives a number for.
const RE_EXCELNUM = /^ *(\()? *([+-])? *(\$)? *([+-])? *((?:[0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)(?:\.[0-9]*)?|\.[0-9]+)(?:[eE]([+-]?[0-9]+))? *(%)? *(\))? *$/;
const numTextCache = new Map();

function excelNumText(s) {
  const m = RE_EXCELNUM.exec(s);
  if (m === null) return null;
  if (!!m[1] !== !!m[8]) return null;
  if (m[2] && m[4]) return null;
  if (m[1] && (m[2] || m[4])) return null;
  let x = Number(m[5].replace(/,/g, '') + (m[6] !== undefined ? 'e' + m[6] : ''));
  if (m[7]) x /= 100;
  if (m[1] || m[2] === '-' || m[4] === '-') x = -x;
  return isFinite(x) ? x : null;
}

// Text Excel reads as a number: "5", " 5 ", "1e3", "5%", "$5", "1,000", "(5)". null = not a number.
// (Dates such as "1-2" are not read as numbers here.)
function textToNum(s) {
  const n = s.length;
  if (n === 0 || n > 100) return null;
  if (RE_NOT_NUMCHAR.test(s) || !RE_DIGIT.test(s)) return null;
  if (RE_PLAIN.test(s)) {
    const x = Number(s);
    if (isFinite(x)) return x;
  }
  let r = numTextCache.get(s);
  if (r === undefined) { r = excelNumText(s); numTextCache.set(s, r); }
  return r;
}

//------------------------------------------------------------------------------
//  Conversions of the formula engine
//------------------------------------------------------------------------------
// Excel arithmetic on a value: blank = 0, TRUE = 1, number text = number, other text = #VALUE!,
// an error stays that error. Returns a number or an XErr.
function toNum(v) {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const x = textToNum(v);
    return x === null ? ERR_VALUE : x;
  }
  return v;
}

// Excel text of a value. Returns a string or an XErr.
function toText(v) {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof XErr) return v;
  return numText(v);
}

// Equal to 15 significant digits, like Excel's = on numbers.
function numEq(a, b) {
  if (a === b) return true;
  let m = Math.abs(a);
  if (Math.abs(b) > m) m = Math.abs(b);
  if (Math.abs(a - b) > m * 0.000000000000005) return false;
  return k15(a) === k15(b);
}

function numCmp(a, b) {
  if (numEq(a, b)) return 0;
  return a < b ? -1 : 1;
}

const textCollator = new Intl.Collator(undefined, { sensitivity: 'accent' });

// StrComp(a, b, vbTextCompare): case is ignored.
function strCompText(a, b) {
  if (a === b) return 0;
  const A = a.toUpperCase(), B = b.toUpperCase();
  if (A === B) return 0;
  const c = textCollator.compare(a, b);
  if (c === 0) return A < B ? -1 : 1;
  return c < 0 ? -1 : 1;
}

function typeRank(v) {
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 1;
}

function blankLike(v) {
  if (typeof v === 'string') return '';
  if (typeof v === 'boolean') return false;
  return 0;
}

// Excel order: numbers < text < TRUE/FALSE; blank = 0, "" or FALSE; text ignores case;
// numbers are equal when they agree to 15 significant digits. (No errors here.)
function xCompare(a, b) {
  if (a === null && b === null) return 0;
  if (a === null) a = blankLike(b);
  if (b === null) b = blankLike(a);
  const ta = typeRank(a), tb = typeRank(b);
  if (ta !== tb) return ta < tb ? -1 : 1;
  if (ta === 1) return numCmp(a, b);
  if (ta === 2) return strCompText(a, b);
  if (a === b) return 0;
  return a === false ? -1 : 1;
}

// Exact difference (type, value, error, case of text).
function differs(a, b) {
  const ea = a instanceof XErr, eb = b instanceof XErr;
  if (ea || eb) return !(ea && eb && a.text === b.text);
  if (a === null || b === null) return a !== b;
  if (typeof a !== typeof b) return true;
  return a !== b;
}

// Same result, numbers within 1E-9 (relative, or absolute below 1).
function sameValue(a, b) {
  if (typeof a === 'number' && typeof b === 'number') {
    let m = Math.abs(a);
    if (Math.abs(b) > m) m = Math.abs(b);
    if (m < 1) m = 1;
    return Math.abs(a - b) <= m * 0.000000001;
  }
  return !differs(a, b);
}

// VBA Round (to even on .5).
function vbRound(x) {
  const f = Math.floor(x), d = x - f;
  if (d > 0.5) return f + 1;
  if (d < 0.5) return f;
  return f % 2 === 0 ? f : f + 1;
}

// Text of a value for messages: "text", 12.5, TRUE, #N/A, (blank).
function valText(v) {
  if (v === null) return '(ว่าง)';
  if (typeof v === 'string') return '"' + v + '"';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof XErr) return v.text;
  return cstr(v);
}

//------------------------------------------------------------------------------
//  Cell addresses
//------------------------------------------------------------------------------
const colLetterCache = [];
function colLetter(c) {
  let s = colLetterCache[c];
  if (s === undefined) {
    s = '';
    let n = c;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    colLetterCache[c] = s;
  }
  return s;
}

function colNumber(letters) {
  let c = 0;
  for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) & 31);
  return c;
}

function addr(r, c) {
  return colLetter(c) + r;
}
