// Loads the core modules (src/core) into one script context, as the page and the worker do.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CORE_FILES = ['util', 'model', 'zip', 'xlsxread', 'formula', 'engine', 'capa', 'xlsxwrite', 'api'];

function coreSource() {
  return CORE_FILES.map((f) => {
    const p = path.join(__dirname, '..', 'src', 'core', f + '.js');
    return '// ---- ' + f + '.js ----\n' + fs.readFileSync(p, 'utf8');
  }).join('\n');
}

function loadCore() {
  const ctx = vm.createContext({
    console, TextDecoder, TextEncoder, DecompressionStream, CompressionStream, Blob, File, ReadableStream, Intl,
    setTimeout, Date, Math,
  });
  vm.runInContext(coreSource() + '\n;globalThis.__x = { XErr, UserError, Workbook, Sheet, session, FIRST_ROW, COL_PLAN, BLOCK, N_MONTH, COL_QTY, ERR_NA, ERR_VALUE };',
    ctx, { filename: 'core.js' });
  return ctx;
}

function fileFrom(p) {
  const buf = fs.readFileSync(p);
  return new File([buf], path.basename(p));
}

module.exports = { CORE_FILES, coreSource, loadCore, fileFrom };
