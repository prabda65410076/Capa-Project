#!/usr/bin/env node
// Timing on a big workbook: node tools/perf.js <file.xlsx> [--export]
'use strict';
const fs = require('fs');
const path = require('path');
const { loadCore, fileFrom } = require('./core.js');

(async () => {
  const file = process.argv[2];
  const c = loadCore();
  const mb = () => (process.memoryUsage().rss / 1048576).toFixed(0) + ' MB';
  let t = Date.now();
  const f = fileFrom(file);
  const info = await c.apiCall('load', { file: f });
  console.log('load', ((Date.now() - t) / 1000).toFixed(1) + 's', 'rss', mb(), info.sheets.map((s) => s.name + ':' + s.rows + 'x' + s.cols).join(' '));
  t = Date.now();
  const r = await c.apiCall('recalc', {});
  console.log('recalc', ((Date.now() - t) / 1000).toFixed(1) + 's', 'rss', mb(), 'items', r.items, 'groups', r.groups, 'cells', r.cells, 'changed', r.changedCells);
  for (const [n, s] of r.times) console.log('   ', n, s.toFixed(2) + 's');
  console.log('    ', r.evalNote);
  if (process.argv.includes('--export')) {
    for (const detail of [false, true]) {
      t = Date.now();
      const ex = await c.apiCall('export', { opts: { detail } });
      const outPath = file.replace(/\.xlsx$/, detail ? '_web.xlsx' : '_light.xlsx');
      const w = fs.createWriteStream(outPath);
      for await (const chunk of ex.blob.stream()) if (!w.write(chunk)) await new Promise((res) => w.once('drain', res));
      w.end();
      await new Promise((res) => w.once('finish', res));
      console.log('export detail=' + detail, ((Date.now() - t) / 1000).toFixed(1) + 's', (ex.blob.size / 1048576).toFixed(0) + ' MB', 'rss', mb());
    }
  }
})().catch((e) => { console.log(e.stack || e.message || e); process.exit(1); });
