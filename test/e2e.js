#!/usr/bin/env node
// Browser test of index.html (Chromium via Playwright): node test/e2e.js [screenshot dir]
'use strict';
const path = require('path');
const fs = require('fs');
let playwright;
try { playwright = require('playwright'); } catch (e) { playwright = require('/opt/node22/lib/node_modules/playwright'); }
const { loadCore, fileFrom } = require('../tools/core.js');

const ROOT = path.join(__dirname, '..');
const shots = process.argv[2] || null;
let failed = 0;
const check = (name, ok, detail) => { console.log((ok ? 'ok   ' : 'FAIL ') + name + (ok || detail === undefined ? '' : ' ' + detail)); if (!ok) failed++; };

(async () => {
  const browser = await playwright.chromium.launch();
  for (const scheme of ['light', 'dark']) {
    const ctx = await browser.newContext({ colorScheme: scheme, viewport: { width: 1360, height: 900 }, acceptDownloads: true });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(String(e)));
    page.on('console', (m) => { if (m.type() === 'error' && !/fonts\.g/.test(m.location().url || '')) errors.push(m.text() + ' ' + (m.location().url || '')); });
    await page.goto('file://' + path.join(ROOT, 'index.html'));
    // the sample opens by itself
    await page.waitForSelector('#calc-status .pill.ok', { timeout: 60000 });
    check(scheme + ': sample calculated', await page.isVisible('#sample-banner'));
    const changed = await page.textContent('#fig-changed');
    check(scheme + ': sample shows changed cells', Number(changed.replace(/,/g, '')) > 1000, changed);
    check(scheme + ': workload table', (await page.$$('#work tbody tr')).length > 10);
    if (shots) await page.screenshot({ path: path.join(shots, 'desktop-' + scheme + '.png'), fullPage: false });
    if (scheme === 'dark') { check(scheme + ': no page errors', errors.length === 0, errors.join(' | ')); await ctx.close(); continue; }

    // a real file
    await page.setInputFiles('#file-input', path.join(ROOT, 'test/fixtures/AB_values.xlsx'));
    await page.waitForFunction(() => !document.querySelector('#sample-banner').hidden === false &&
      document.querySelector('#calc-status .pill.ok') !== null, null, { timeout: 60000 });
    check('file: banner gone', !(await page.isVisible('#sample-banner')));
    const nChanged = Number((await page.textContent('#fig-changed')).replace(/,/g, ''));
    check('file: changed cells', nChanged > 1000, nChanged);
    // tabs
    await page.click('#tabbtn-changes');
    check('changes tab lists sheets', (await page.$$('#changes section')).length >= 3);
    await page.click('#changes tbody tr');
    await page.waitForSelector('#cellinfo .addr');
    check('click on a change opens the cell', /!/.test(await page.textContent('#cellinfo .addr')));
    await page.click('#tabbtn-sheet');
    await page.fill('#goto-input', "'1-3.CAPA'!BU7");
    await page.press('#goto-input', 'Enter');
    await page.waitForFunction(() => /1-3\.CAPA!BU7/.test(document.querySelector('#cellinfo .addr')?.textContent || ''));
    const code = await page.textContent('#cellinfo');
    check('cell info shows the formula', /SUMIFS/.test(code) && /CAPA_RULES/.test(code), code.slice(0, 200));
    await page.waitForTimeout(300);
    if (shots) await page.screenshot({ path: path.join(shots, 'viewer.png') });
    await page.click('#tabbtn-notes');
    check('notes', /PART METHOD/.test(await page.textContent('#notes')));
    // download
    const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btn-save')]);
    const out = path.join(ROOT, 'test/fixtures/out/e2e_download.xlsx');
    fs.mkdirSync(path.dirname(out), { recursive: true });
    await dl.saveAs(out);
    check('download name', dl.suggestedFilename() === 'AB_values_web.xlsx', dl.suggestedFilename());
    const c = loadCore();
    await c.apiCall('load', { file: fileFrom(out) });
    const r = await c.apiCall('recalc', {});
    const unexpected = r.changes.filter((ch) => !(ch.sheet === 'TEST')).length;
    check('downloaded file: recalculating again changes nothing', unexpected === 0, JSON.stringify(r.changes.map((ch) => [ch.sheet, ch.n])));
    check('save status shown', /ดาวน์โหลด/.test(await page.textContent('#save-status')));
    // a file that is not a workbook
    const bad = path.join(ROOT, 'test/fixtures/out/not-a-workbook.xlsx');
    fs.writeFileSync(bad, 'hello');
    await page.setInputFiles('#file-input', bad);
    await page.waitForSelector('#error:not([hidden])');
    check('bad file shows a message', /ไม่ใช่ไฟล์ Excel/.test(await page.textContent('#error')));
    check('no page errors', errors.length === 0, errors.join(' | '));
    await ctx.close();
  }
  // phone width
  const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2 });
  const page = await ctx.newPage();
  await page.goto('file://' + path.join(ROOT, 'index.html'));
  await page.waitForSelector('#calc-status .pill.ok', { timeout: 60000 });
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
  check('phone: no horizontal page scroll', overflow <= 0, overflow);
  if (shots) await page.screenshot({ path: path.join(shots, 'phone.png'), fullPage: true });
  await ctx.close();
  await browser.close();
  console.log(failed ? failed + ' failed' : 'all passed');
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.log(e.stack); process.exit(1); });
