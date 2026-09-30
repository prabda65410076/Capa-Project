'use strict';
// The page: file -> worker (core in src/core) -> results. The core also runs here when a
// Web Worker cannot be started.
(function () {
  const $ = (id) => document.getElementById(id);
  const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

  //----------------------------------------------------------------------------
  //  Calls into the core (worker or this page)
  //----------------------------------------------------------------------------
  function workerMain() {
    self.onmessage = async (ev) => {
      const { id, cmd, args } = ev.data;
      let last = 0;
      const progress = (p) => {
        const now = Date.now();
        if (now - last > 80 || p.done === p.total) { last = now; self.postMessage({ id, progress: p }); }
      };
      try {
        const res = await apiCall(cmd, args, progress);
        self.postMessage({ id, ok: true, res });
      } catch (e) {
        self.postMessage({ id, ok: false, err: errorText(e) });
      }
    };
  }

  let worker = null, seq = 0;
  const pending = new Map();
  try {
    const src = $('capa-core').textContent + '\n;(' + workerMain.toString() + ')();';
    worker = new Worker(URL.createObjectURL(new Blob([src], { type: 'text/javascript' })));
    worker.onmessage = (ev) => {
      const m = ev.data, p = pending.get(m.id);
      if (!p) return;
      if (m.progress) { if (p.onProgress) p.onProgress(m.progress); return; }
      pending.delete(m.id);
      if (m.ok) p.res(m.res); else p.rej(m.err);
    };
    worker.onerror = (ev) => {
      ev.preventDefault();
      const w = worker;
      worker = null;
      if (w) w.terminate();
      for (const [id, p] of pending) { pending.delete(id); runLocal(p.cmd, p.args, p.onProgress).then(p.res, p.rej); }
    };
  } catch (e) {
    worker = null;
  }

  function runLocal(cmd, args, onProgress) {
    const progress = (p) => { if (onProgress) onProgress(p); return new Promise((r) => setTimeout(r, 0)); };
    return apiCall(cmd, args, progress).catch((e) => { throw errorText(e); });
  }

  function call(cmd, args, onProgress) {
    if (worker === null) return runLocal(cmd, args, onProgress);
    return new Promise((res, rej) => {
      const id = ++seq;
      pending.set(id, { res, rej, onProgress, cmd, args });
      worker.postMessage({ id, cmd, args });
    });
  }

  //----------------------------------------------------------------------------
  //  Helpers
  //----------------------------------------------------------------------------
  const el = (tag, props, kids) => {
    const e = document.createElement(tag);
    if (props) for (const k of Object.keys(props)) {
      if (k === 'class') e.className = props[k];
      else if (k === 'text') e.textContent = props[k];
      else e.setAttribute(k, props[k]);
    }
    if (kids) for (const k of kids) if (k !== null && k !== undefined) e.append(k);
    return e;
  };
  const fmtBytes = (n) => n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
  const fmtInt = (n) => Number(n).toLocaleString('en-US');
  const fmtSec = (s) => s < 10 ? s.toFixed(2) + ' วินาที' : s.toFixed(1) + ' วินาที';
  const letters = (c) => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; };
  const colNum = (s) => { let c = 0; for (const ch of s.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64); return c; };

  const S = {
    busy: false, sample: false, file: null, info: null, rep: null, totalTime: 0,
    sheet: 0, sel: null, tiles: new Map(), wanted: new Set(),
  };

  function setProgress(text, done, total) {
    $('progress').hidden = false;
    $('progress-text').textContent = text || '';
    const fill = $('progress-fill');
    if (total > 0) { fill.classList.remove('indet'); fill.style.width = Math.min(100, (done / total) * 100).toFixed(1) + '%'; }
    else { fill.classList.add('indet'); fill.style.width = ''; }
  }
  function hideProgress() { $('progress').hidden = true; }

  function showError(err) {
    const box = $('error');
    box.replaceChildren();
    if (!err) { box.hidden = true; return; }
    const user = err.user !== false && typeof err.message === 'string';
    box.append(el('strong', { text: user ? 'คำนวณไม่ได้' : 'เกิดข้อผิดพลาดภายในโปรแกรม' }));
    if (user) box.append(document.createTextNode(err.message));
    else {
      box.append(document.createTextNode('กรุณาส่งรายละเอียดด้านล่างให้ผู้ดูแลโปรแกรม'));
      const d = el('details', null, [el('summary', { text: 'รายละเอียด' }), el('pre', { text: String(err.message || err) })]);
      box.append(d);
    }
    box.hidden = false;
  }

  function setBusy(b) {
    S.busy = b;
    $('file-input').disabled = b;
    $('btn-save').disabled = b || S.rep === null;
  }

  //----------------------------------------------------------------------------
  //  Open and calculate
  //----------------------------------------------------------------------------
  async function openFile(file, isSample) {
    if (S.busy) return;
    setBusy(true);
    showError(null);
    S.rep = null; S.info = null; S.sample = isSample; S.file = file; S.tiles.clear(); S.wanted.clear(); S.sel = null;
    $('sample-banner').hidden = !isSample;
    $('save-status').textContent = '';
    $('step-file').classList.remove('done'); $('step-calc').classList.remove('done'); $('step-save').classList.remove('done');
    $('calc-status').replaceChildren();
    renderFileInfo(file, null);
    const t0 = performance.now();
    try {
      setProgress('เปิดไฟล์...', 0, 0);
      S.info = await call('load', { file }, (p) => setProgress(p.text, p.done, p.total));
      renderFileInfo(file, S.info);
      $('step-file').classList.add('done');
      setProgress('คำนวณ...', 0, 0);
      S.rep = await call('recalc', {}, (p) => setProgress(p.text, 0, 0));
      S.totalTime = (performance.now() - t0) / 1000;
      $('step-calc').classList.add('done');
      renderAll();
    } catch (err) {
      showError(err);
      renderStatus(null);
    } finally {
      hideProgress();
      setBusy(false);
    }
  }

  function renderFileInfo(file, info) {
    const box = $('fileinfo');
    box.hidden = false;
    const mode = info ? info.mode : '';
    const modePill = !info ? el('span', { class: 'pill', text: 'กำลังอ่าน' }) :
      mode === 'VALUES' ? el('span', { class: 'pill ok', text: 'สูตรอยู่ใน CAPA_RULES' }) :
      mode === 'FORMULAS' ? el('span', { class: 'pill warn', text: 'สูตรอยู่ในเซลล์ (หลัง CAPA_EditFormulas)' }) :
      el('span', { class: 'pill bad', text: 'ยังไม่ได้รัน CAPA_Setup' });
    const kv = el('dl', { class: 'kv' }, [
      el('dt', { text: 'ขนาด' }), el('dd', { class: 'data', text: fmtBytes(file.size) }),
      el('dt', { text: 'รูปแบบ' }), el('dd', null, [modePill]),
    ]);
    if (info) {
      kv.append(el('dt', { text: 'ชีต' }), el('dd', { text: info.sheets.filter((s) => s.state === 'visible').length + ' ชีต (+' +
        info.sheets.filter((s) => s.state !== 'visible').length + ' ที่ซ่อน)' }));
      kv.append(el('dt', { text: 'อ่านไฟล์' }), el('dd', { class: 'data', text: fmtSec(info.readTime) }));
    }
    box.replaceChildren(el('div', { class: 'name', text: (S.sample ? 'ตัวอย่าง: ' : '') + file.name }), kv);
  }

  function renderStatus(rep) {
    const box = $('calc-status');
    box.replaceChildren();
    if (!rep) return;
    const t = el('table', { class: 'times' });
    for (const [name, sec] of rep.times) t.append(el('tr', null, [el('td', { text: name }), el('td', { text: fmtSec(sec) })]));
    t.append(el('tr', { class: 'total' }, [el('td', { text: 'รวมตั้งแต่เปิดไฟล์' }), el('td', { text: fmtSec(S.totalTime) })]));
    const nNotes = noteCount(rep);
    box.append(el('p', { class: 'status' }, [
      el('span', { class: 'pill ok', text: 'คำนวณเสร็จ' }),
      el('span', { text: fmtInt(rep.changedCells) + ' เซลล์ค่าเปลี่ยน' + (nNotes ? ' · คำเตือน ' + nNotes + ' เรื่อง' : '') }),
    ]), t);
  }

  function noteCount(rep) {
    let n = 0;
    if (rep.methodNote) n += rep.methodNote.split(/\n\n/).length;
    if (rep.unsupportedCount) n++;
    return n;
  }

  function renderAll() {
    const rep = S.rep;
    renderStatus(rep);
    $('fig-items').textContent = fmtInt(rep.items);
    $('fig-groups').textContent = fmtInt(rep.groups);
    $('fig-cells').textContent = fmtInt(rep.cells);
    $('fig-changed').textContent = fmtInt(rep.changedCells);
    $('fig-time').replaceChildren(document.createTextNode(S.totalTime.toFixed(1)), el('small', { text: ' วินาที' }));
    $('count-changes').textContent = rep.changedCells ? fmtInt(rep.changedCells) : '';
    $('count-notes').textContent = noteCount(rep) ? String(noteCount(rep)) : '';
    renderWork();
    renderChanges();
    renderNotes();
    renderSheetSelect();
    const name = S.file.name;
    $('btn-save').textContent = 'ดาวน์โหลด ' + name.replace(/(\.[^.]+)?$/, (ext) => '_web' + (ext || '.xlsx'));
  }

  //----------------------------------------------------------------------------
  //  Tab: hours of work per process
  //----------------------------------------------------------------------------
  function hours(x) {
    return x.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  }

  function renderWork() {
    const box = $('work');
    const w = S.rep.workload;
    const all = $('opt-allproc').checked;
    const table = el('table', { class: 'ledger-table' });
    const head = el('tr', null, [el('th', { class: 'l sticky-col', text: 'กระบวนการ' }), el('th', { class: 'l', text: 'กลุ่ม' })]);
    for (const m of w.months) head.append(el('th', { text: m }));
    head.append(el('th', { text: 'รวมทั้งปี' }));
    table.append(el('thead', null, [head]));
    const body = el('tbody');
    const colSum = new Array(12).fill(0), colErr = new Array(12).fill(false);
    let shown = 0;
    for (const p of w.procs) {
      let total = 0, err = false, max = -1, maxAt = -1;
      p.months.forEach(([v, e], i) => {
        total += v;
        colSum[i] += v;
        if (e) { err = true; colErr[i] = true; }
        if (v > max) { max = v; maxAt = i; }
      });
      if (!all && !err && total === 0) continue;
      shown++;
      const tr = el('tr', null, [el('td', { class: 'l sticky-col', text: p.k + ' ' + p.name }), el('td', { class: 'l group', text: p.group })]);
      p.months.forEach(([v, e], i) => {
        const td = el('td', { class: e ? 'err' : v === 0 ? 'zero' : i === maxAt && max > 0 ? 'max' : '', text: (v === 0 ? '0' : hours(v)) + (e ? ' *' : '') });
        if (e) td.title = 'บาง spec เป็น ' + e + ' จึงไม่ได้รวมไว้';
        tr.append(td);
      });
      tr.append(el('td', { class: err ? 'err' : '', text: hours(total) + (err ? ' *' : '') }));
      body.append(tr);
    }
    const sum = el('tr', { class: 'sum' }, [el('td', { class: 'l sticky-col', text: 'รวมทุกกระบวนการ' }), el('td', { class: 'l', text: '' })]);
    let grand = 0;
    colSum.forEach((v, i) => { grand += v; sum.append(el('td', { class: colErr[i] ? 'err' : '', text: hours(v) + (colErr[i] ? ' *' : '') })); });
    sum.append(el('td', { text: hours(grand) }));
    body.append(sum);
    table.append(body);
    const parts = [el('div', { class: 'table-wrap' }, [table])];
    if (colErr.some(Boolean)) {
      parts.push(el('p', { class: 'hint', text: '* มีบาง spec เป็น error (เช่น Plan หรือ ST ที่ไม่ใช่ตัวเลข) ตัวเลขที่เห็นไม่รวม spec นั้น ชี้ที่ตัวเลขเพื่อดูชนิดของ error และดูต้นเหตุได้ในชีต ST_SUM' }));
    }
    if (shown === 0) parts.unshift(el('p', { class: 'empty', text: 'ไม่มีชั่วโมงงานในไฟล์นี้ (Plan เป็น 0 ทั้งหมด?)' }));
    box.replaceChildren(...parts);
  }

  //----------------------------------------------------------------------------
  //  Tab: changed cells, notes
  //----------------------------------------------------------------------------
  function renderChanges() {
    const box = $('changes');
    const rep = S.rep;
    if (rep.changes.length === 0) {
      box.replaceChildren(el('p', { class: 'note ok' }, [
        el('h4', { text: 'ไม่มีเซลล์ที่ค่าเปลี่ยน' }),
        document.createTextNode('ทุกผลลัพธ์ที่เว็บคำนวณได้ตรงกับค่าที่อยู่ในไฟล์ (เช่นไฟล์ที่เพิ่งกด Recalc CAPA ใน Excel)'),
      ]));
      return;
    }
    const parts = [];
    for (const ch of rep.changes) {
      const t = el('table', { class: 'ledger-table' });
      t.append(el('thead', null, [el('tr', null, [el('th', { class: 'l', text: 'เซลล์' }), el('th', { text: 'ค่าเดิมในไฟล์' }), el('th', { text: 'ค่าใหม่' })])]));
      const tb = el('tbody');
      for (const x of ch.ex) {
        const tr = el('tr', { class: 'link', tabindex: '0' }, [
          el('td', { class: 'l data', text: x.a }), el('td', { text: x.old === '' ? '(ว่าง)' : x.old }), el('td', { text: x.val === '' ? '(ว่าง)' : x.val }),
        ]);
        const go = () => gotoCell(ch.index, x.r, x.c);
        tr.addEventListener('click', go);
        tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
        tb.append(tr);
      }
      t.append(tb);
      const more = ch.n > ch.ex.length ? el('p', { class: 'hint', text: 'แสดง ' + ch.ex.length + ' เซลล์แรก จาก ' + fmtInt(ch.n) }) : null;
      parts.push(el('section', null, [el('h4', null, [document.createTextNode(ch.sheet), el('span', { text: fmtInt(ch.n) + ' เซลล์' })]),
        el('div', { class: 'table-wrap' }, [t]), more]));
    }
    box.replaceChildren(...parts);
  }

  function renderNotes() {
    const box = $('notes');
    const rep = S.rep;
    const parts = [];
    if (rep.methodNote) {
      for (const block of rep.methodNote.split(/\n\n/)) parts.push(el('div', { class: 'note', text: block }));
    }
    if (rep.unsupportedCount) {
      const t = el('table', { class: 'ledger-table' });
      t.append(el('thead', null, [el('tr', null, [el('th', { class: 'l', text: 'เซลล์' }), el('th', { class: 'l', text: 'เหตุผล' }), el('th', { class: 'l', text: 'สูตร' })])]));
      const tb = el('tbody');
      for (const u of rep.unsupported) {
        tb.append(el('tr', null, [el('td', { class: 'l data', text: u.sheet + '!' + u.a }), el('td', { class: 'l', text: u.why }), el('td', { class: 'l data', text: u.formula })]));
      }
      t.append(tb);
      parts.push(el('div', { class: 'note' }, [
        el('h4', { text: 'สูตรที่เว็บคำนวณไม่ได้ ' + fmtInt(rep.unsupportedCount) + ' เซลล์' }),
        document.createTextNode('เซลล์เหล่านี้ใช้ค่าเดิมในไฟล์ ถ้าข้อมูลที่สูตรอ่านเปลี่ยนไป ให้เปิดไฟล์ใน Excel แล้วกด F9 หรือแก้สูตรให้ใช้ฟังก์ชันที่รองรับ'),
        el('div', { class: 'table-wrap' }, [t]),
      ]));
    }
    if (parts.length === 0) parts.push(el('p', { class: 'note ok', text: 'ไม่มีคำเตือน' }));
    box.replaceChildren(...parts);
  }

  //----------------------------------------------------------------------------
  //  Tab: sheet viewer (canvas; cells are fetched in tiles)
  //----------------------------------------------------------------------------
  const G = { head: 24, rowHead: 58, colW: 96, rowH: 22, tileR: 64, tileC: 16 };
  const canvas = $('grid-canvas'), scroller = $('grid-scroll'), spacer = $('grid-spacer');
  let drawQueued = false;

  function renderSheetSelect() {
    const sel = $('sheet-select');
    sel.replaceChildren();
    S.rep.sheets.forEach((sh, i) => {
      const label = sh.name + (sh.state !== 'visible' ? ' (ซ่อน)' : '');
      sel.append(el('option', { value: String(i), text: label }));
    });
    const first = S.rep.sheets.findIndex((sh) => sh.name.startsWith('1-3.'));
    setSheet(first >= 0 ? first : 0);
  }

  function sheetDims() {
    const sh = S.rep.sheets[S.sheet];
    return { rows: Math.min(600000, Math.max(sh.rows + 30, 60)), cols: Math.min(16384, Math.max(sh.cols + 4, 14)) };
  }

  function setSheet(i) {
    S.sheet = i;
    $('sheet-select').value = String(i);
    const d = sheetDims();
    spacer.style.width = (G.rowHead + d.cols * G.colW) + 'px';
    spacer.style.height = (G.head + d.rows * G.rowH) + 'px';
    scroller.scrollTop = 0;
    scroller.scrollLeft = 0;
    S.sel = null;
    queueDraw();
  }

  function queueDraw() {
    if (drawQueued) return;
    drawQueued = true;
    requestAnimationFrame(() => { drawQueued = false; draw(); });
  }

  function tileKey(si, tr, tc) { return si + ':' + tr + ':' + tc; }

  function tile(si, tr, tc) {
    const k = tileKey(si, tr, tc);
    const t = S.tiles.get(k);
    if (t) return t;
    if (!S.wanted.has(k)) {
      S.wanted.add(k);
      const r1 = tr * G.tileR + 1, c1 = tc * G.tileC + 1;
      call('block', { si, r1, c1, r2: r1 + G.tileR - 1, c2: c1 + G.tileC - 1 }).then((b) => {
        S.wanted.delete(k);
        if (S.tiles.size > 400) S.tiles.delete(S.tiles.keys().next().value);
        S.tiles.set(k, { r1, c1, texts: b.texts, kinds: b.kinds });
        if (si === S.sheet) queueDraw();
      }, () => S.wanted.delete(k));
    }
    return null;
  }

  function draw() {
    const box = canvas.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    const w = Math.max(1, Math.floor(box.width)), h = Math.max(1, Math.floor(box.height));
    if (canvas.width !== Math.floor(w * dpr) || canvas.height !== Math.floor(h * dpr)) {
      canvas.width = Math.floor(w * dpr);
      canvas.height = Math.floor(h * dpr);
    }
    const g = canvas.getContext('2d');
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    const css = getComputedStyle(document.documentElement);
    const col = (n) => css.getPropertyValue(n).trim();
    const C = {
      panel: col('--panel'), panel2: col('--panel-2'), ink: col('--ink'), muted: col('--muted'), line: col('--line'),
      calc: col('--calc'), chg: col('--copper-soft'), copper: col('--copper'), keep: col('--warn-soft'), warn: col('--warn'),
      bad: col('--bad'), focus: col('--focus'),
    };
    const dataFont = '12px ' + col('--font-data');
    const textFont = '12px ' + col('--font-body');
    g.fillStyle = C.panel;
    g.fillRect(0, 0, w, h);
    if (!S.rep) return;
    const d = sheetDims();
    const sx = scroller.scrollLeft, sy = scroller.scrollTop;
    const c0 = Math.floor(sx / G.colW) + 1, r0 = Math.floor(sy / G.rowH) + 1;
    const offX = G.rowHead - (sx % G.colW), offY = G.head - (sy % G.rowH);
    const nc = Math.min(d.cols - c0 + 1, Math.ceil((w - G.rowHead) / G.colW) + 1);
    const nr = Math.min(d.rows - r0 + 1, Math.ceil((h - G.head) / G.rowH) + 1);
    g.textBaseline = 'middle';
    // cells
    for (let i = 0; i < nr; i++) {
      const r = r0 + i, y = offY + i * G.rowH;
      for (let j = 0; j < nc; j++) {
        const c = c0 + j, x = offX + j * G.colW;
        const t = tile(S.sheet, Math.floor((r - 1) / G.tileR), Math.floor((c - 1) / G.tileC));
        if (!t) continue;
        const idx = (r - t.r1) * G.tileC + (c - t.c1);
        const kind = t.kinds[idx], text = t.texts[idx];
        if (kind & 16) { g.fillStyle = C.keep; g.fillRect(x, y, G.colW, G.rowH); }
        else if (kind & 8) { g.fillStyle = C.chg; g.fillRect(x, y, G.colW, G.rowH); g.fillStyle = C.copper; g.fillRect(x, y, 3, G.rowH); }
        else if (kind & 4) { g.fillStyle = C.calc; g.fillRect(x, y, G.colW, G.rowH); }
        if (!text) continue;
        g.fillStyle = kind & 2 ? C.bad : C.ink;
        const num = (kind & 1) !== 0;
        g.font = num || (kind & 2) ? dataFont : textFont;
        g.save();
        g.beginPath();
        g.rect(x + 1, y, G.colW - 2, G.rowH);
        g.clip();
        if (num) { g.textAlign = 'right'; g.fillText(text, x + G.colW - 6, y + G.rowH / 2 + 1); }
        else if (kind & 32) { g.textAlign = 'center'; g.fillText(text, x + G.colW / 2, y + G.rowH / 2 + 1); }
        else { g.textAlign = 'left'; g.fillText(text, x + 6, y + G.rowH / 2 + 1); }
        g.restore();
      }
    }
    // grid lines
    g.strokeStyle = C.line;
    g.lineWidth = 1;
    g.beginPath();
    for (let j = 0; j <= nc; j++) { const x = Math.floor(offX + j * G.colW) + 0.5; g.moveTo(x, G.head); g.lineTo(x, h); }
    for (let i = 0; i <= nr; i++) { const y = Math.floor(offY + i * G.rowH) + 0.5; g.moveTo(G.rowHead, y); g.lineTo(w, y); }
    g.stroke();
    // selection
    if (S.sel && S.sel.si === S.sheet) {
      const x = offX + (S.sel.c - c0) * G.colW, y = offY + (S.sel.r - r0) * G.rowH;
      g.strokeStyle = C.focus;
      g.lineWidth = 2;
      g.strokeRect(x + 1, y + 1, G.colW - 2, G.rowH - 2);
    }
    // headers
    g.fillStyle = C.panel2;
    g.fillRect(0, 0, w, G.head);
    g.fillRect(0, 0, G.rowHead, h);
    g.font = dataFont;
    g.fillStyle = C.muted;
    g.textAlign = 'center';
    for (let j = 0; j < nc; j++) {
      const x = offX + j * G.colW;
      if (x + G.colW < G.rowHead) continue;
      g.fillText(letters(c0 + j), x + G.colW / 2, G.head / 2 + 1);
    }
    g.textAlign = 'right';
    for (let i = 0; i < nr; i++) {
      const y = offY + i * G.rowH;
      if (y + G.rowH < G.head) continue;
      g.fillText(String(r0 + i), G.rowHead - 8, y + G.rowH / 2 + 1);
    }
    g.fillStyle = C.panel2;
    g.fillRect(0, 0, G.rowHead, G.head);
    g.strokeStyle = C.line;
    g.lineWidth = 1;
    g.beginPath();
    g.moveTo(0, G.head + 0.5); g.lineTo(w, G.head + 0.5);
    g.moveTo(G.rowHead + 0.5, 0); g.lineTo(G.rowHead + 0.5, h);
    g.stroke();
  }

  scroller.addEventListener('scroll', queueDraw, { passive: true });
  window.addEventListener('resize', queueDraw);
  try {
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', queueDraw);
    new MutationObserver(queueDraw).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
  } catch (e) { /* older browsers: the grid redraws on the next scroll */ }

  scroller.addEventListener('click', (e) => {
    if (!S.rep) return;
    const box = scroller.getBoundingClientRect();
    const x = e.clientX - box.left, y = e.clientY - box.top;
    if (x < G.rowHead || y < G.head) return;
    const c = Math.floor((x - G.rowHead + scroller.scrollLeft) / G.colW) + 1;
    const r = Math.floor((y - G.head + scroller.scrollTop) / G.rowH) + 1;
    selectCell(S.sheet, r, c);
  });

  scroller.addEventListener('keydown', (e) => {
    if (!S.sel || !S.rep) return;
    const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (!mv) return;
    e.preventDefault();
    const r = Math.max(1, S.sel.r + mv[0]), c = Math.max(1, S.sel.c + mv[1]);
    selectCell(S.sheet, r, c);
    scrollIntoView(r, c);
  });

  function scrollIntoView(r, c) {
    const x = G.rowHead + (c - 1) * G.colW, y = G.head + (r - 1) * G.rowH;
    const vw = scroller.clientWidth, vh = scroller.clientHeight;
    if (x - scroller.scrollLeft < G.rowHead) scroller.scrollLeft = x - G.rowHead;
    else if (x + G.colW - scroller.scrollLeft > vw) scroller.scrollLeft = x + G.colW - vw;
    if (y - scroller.scrollTop < G.head) scroller.scrollTop = y - G.head;
    else if (y + G.rowH - scroller.scrollTop > vh) scroller.scrollTop = y + G.rowH - vh;
  }

  async function selectCell(si, r, c) {
    S.sel = { si, r, c };
    queueDraw();
    const box = $('cellinfo');
    try {
      const info = await call('cell', { si, r, c });
      if (!S.sel || S.sel.r !== r || S.sel.c !== c) return;
      const kv = el('dl', { class: 'kv' }, [el('dt', { text: 'ค่า' }), el('dd', { class: 'data', text: info.value === '' ? '(ว่าง)' : info.value + '  ·  ' + info.type })]);
      if (info.old !== undefined) kv.append(el('dt', { text: 'ค่าเดิมในไฟล์' }), el('dd', { class: 'data old', text: info.old }));
      if (info.source) kv.append(el('dt', { text: 'ที่มา' }), el('dd', { text: info.source }));
      const parts = [el('div', { class: 'addr', text: info.sheet + '!' + info.a }), kv];
      if (info.formula) parts.push(el('code', { text: info.formula }));
      if (info.r1c1 && info.r1c1 !== info.formula) parts.push(el('code', { text: 'R1C1: ' + info.r1c1 }));
      if (info.note) parts.push(el('p', { class: 'hint', text: info.note }));
      box.replaceChildren(...parts);
    } catch (err) {
      box.replaceChildren(el('p', { class: 'hint', text: String(err && err.message || err) }));
    }
  }

  function gotoCell(si, r, c) {
    showTab('sheet');
    if (S.sheet !== si) setSheet(si);
    requestAnimationFrame(() => {
      scroller.scrollLeft = Math.max(0, (c - 3) * G.colW);
      scroller.scrollTop = Math.max(0, (r - 5) * G.rowH);
      selectCell(si, r, c);
      scroller.focus({ preventScroll: true });
    });
  }

  $('sheet-select').addEventListener('change', (e) => setSheet(Number(e.target.value)));
  $('goto-form').addEventListener('submit', (e) => {
    e.preventDefault();
    if (!S.rep) return;
    const text = $('goto-input').value.trim();
    const m = /^(?:(?:'([^']+)'|([^!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(text);
    if (!m) { $('cellinfo').replaceChildren(el('p', { class: 'hint', text: 'พิมพ์ที่อยู่เซลล์ เช่น BU7 หรือ ST_SUM!CI7' })); return; }
    let si = S.sheet;
    const name = m[1] || m[2];
    if (name) {
      const i = S.rep.sheets.findIndex((sh) => sh.name.toUpperCase() === name.trim().toUpperCase());
      if (i < 0) { $('cellinfo').replaceChildren(el('p', { class: 'hint', text: 'ไม่พบชีต ' + name })); return; }
      si = i;
    }
    gotoCell(si, Number(m[4]), colNum(m[3]));
  });

  //----------------------------------------------------------------------------
  //  Tabs
  //----------------------------------------------------------------------------
  const TABS = ['work', 'sheet', 'changes', 'notes'];
  function showTab(name) {
    for (const t of TABS) {
      $('tab-' + t).hidden = t !== name;
      $('tabbtn-' + t).setAttribute('aria-selected', String(t === name));
    }
    if (name === 'sheet') queueDraw();
    try { localStorage.setItem('capa-tab', name); } catch (e) { /* no storage */ }
  }
  for (const t of TABS) $('tabbtn-' + t).addEventListener('click', () => showTab(t));
  $('opt-allproc').addEventListener('change', () => { if (S.rep) renderWork(); });

  //----------------------------------------------------------------------------
  //  Saving
  //----------------------------------------------------------------------------
  // One file in a zip (stored): the viewer of claude.ai saves .xlsx and .zip, not .xlsm.
  async function zipOne(name, blob) {
    const data = new Uint8Array(await blob.arrayBuffer());
    const crc = crc32(0, data);
    const nb = new TextEncoder().encode(name);
    const le16 = (v) => [v & 255, (v >>> 8) & 255], le32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    const d = new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    const lh = new Uint8Array([0x50, 0x4B, 3, 4, ...le16(20), ...le16(0x800), 0, 0, ...le16(time), ...le16(date), ...le32(crc),
      ...le32(data.length), ...le32(data.length), ...le16(nb.length), 0, 0, ...nb]);
    const cd = new Uint8Array([0x50, 0x4B, 1, 2, ...le16(20), ...le16(20), ...le16(0x800), 0, 0, ...le16(time), ...le16(date),
      ...le32(crc), ...le32(data.length), ...le32(data.length), ...le16(nb.length), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...nb]);
    const eocd = new Uint8Array([0x50, 0x4B, 5, 6, 0, 0, 0, 0, 1, 0, 1, 0, ...le32(cd.length), ...le32(lh.length + data.length), 0, 0]);
    return new Blob([lh, data, cd, eocd], { type: 'application/zip' });
  }

  async function saveFile(blob, name) {
    const cl = window.claude;
    if (cl && typeof cl.use === 'function') {
      let dl = null;
      try { dl = await cl.use('downloads'); } catch (e) { dl = null; }
      if (dl) {
        let data = blob, fname = name, note = '';
        if (!/\.(xlsx|zip)$/i.test(name)) {
          data = await zipOne(name, blob);
          fname = name.replace(/\.[^.]+$/, '') + '.zip';
          note = ' (ไฟล์ ' + name + ' อยู่ในไฟล์ .zip นี้ ให้แตกไฟล์ก่อนเปิด)';
        }
        try {
          await dl.save({ filename: fname, data });
          return 'บันทึก ' + fname + ' แล้ว' + note;
        } catch (e) {
          const code = e && e.code;
          if (code === 'declined') return 'ยกเลิกการบันทึก';
          if (code === 'rate_limited') return 'มีหน้าต่างบันทึกเปิดอยู่แล้ว ลองใหม่อีกครั้ง';
          if (code === 'too_large') return 'ไฟล์ใหญ่เกินกว่าที่ที่นี่บันทึกได้ ลองไม่เลือก "เขียนค่ารายชิ้นงาน" แล้วบันทึกใหม่';
          throw new Error('บันทึกไฟล์ไม่ได้ (' + (code || e) + ')');
        }
      }
    }
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: name });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 120000);
    return 'ดาวน์โหลด ' + name + ' แล้ว (' + fmtBytes(blob.size) + ')';
  }

  $('btn-save').addEventListener('click', async () => {
    if (S.busy || !S.rep) return;
    setBusy(true);
    showError(null);
    $('save-status').textContent = '';
    const t0 = performance.now();
    try {
      setProgress('เขียนไฟล์...', 0, 0);
      const res = await call('export', { opts: { detail: $('opt-detail').checked } }, (p) => setProgress(p.text, p.done, p.total));
      setProgress('บันทึกไฟล์...', 0, 0);
      const msg = await saveFile(res.blob, res.name);
      $('save-status').textContent = msg + ' · เขียนไฟล์ ' + fmtSec((performance.now() - t0) / 1000);
      $('step-save').classList.add('done');
    } catch (err) {
      showError(err);
    } finally {
      hideProgress();
      setBusy(false);
    }
  });

  //----------------------------------------------------------------------------
  //  Opening files
  //----------------------------------------------------------------------------
  $('file-input').addEventListener('change', (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) openFile(f, false);
    e.target.value = '';
  });
  const drop = $('drop');
  window.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  window.addEventListener('dragleave', (e) => { if (e.target === document.documentElement || e.clientX <= 0) drop.classList.remove('over'); });
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('over');
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) openFile(f, false);
  });

  // the sample workbook (made-up data) so the page opens in a working state
  function sampleFile() {
    const b64 = ($('capa-sample').textContent || '').trim();
    if (!b64) return null;
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new File([bytes], 'CAPA-sample.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  try { const t = localStorage.getItem('capa-tab'); if (TABS.includes(t)) showTab(t); } catch (e) { /* no storage */ }
  const sample = sampleFile();
  if (sample) openFile(sample, true);
})();
