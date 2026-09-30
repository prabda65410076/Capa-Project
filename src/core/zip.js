'use strict';
//==============================================================================
//  ZIP container of .xlsx / .xlsm files. Entries are read as streams (the big
//  sheets never sit in memory as a whole); a new file is written from the old
//  entries (copied as they are) and the changed ones.
//  Uses the browser's DecompressionStream / CompressionStream ('deflate-raw').
//==============================================================================

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(crc, buf) {
  let c = ~crc;
  const t = CRC_TABLE;
  for (let i = 0, n = buf.length; i < n; i++) c = t[(c ^ buf[i]) & 255] ^ (c >>> 8);
  return ~c >>> 0;
}

async function readBytes(file, start, end) {
  return new Uint8Array(await file.slice(start, end).arrayBuffer());
}

function u16(b, i) { return b[i] | (b[i + 1] << 8); }
function u32(b, i) { return (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0; }
function u64(b, i) { return u32(b, i) + u32(b, i + 4) * 4294967296; }

async function zipOpen(file) {
  const size = file.size;
  const head = await readBytes(file, 0, Math.min(size, 8));
  if (head[0] === 0xD0 && head[1] === 0xCF && head[2] === 0x11 && head[3] === 0xE0) {
    throw new UserError('ไฟล์นี้เป็นไฟล์ Excel แบบเก่า (.xls) หรือไฟล์ที่ตั้งรหัสผ่านไว้ ' +
      'กรุณาเปิดใน Excel แล้วบันทึกเป็น .xlsm หรือ .xlsx (ไม่มีรหัสผ่าน) ก่อน');
  }
  if (size < 22 || head[0] !== 0x50 || head[1] !== 0x4B) {
    throw new UserError('ไฟล์นี้ไม่ใช่ไฟล์ Excel (.xlsm / .xlsx)');
  }
  const tailLen = Math.min(size, 65557);
  const tail = await readBytes(file, size - tailLen, size);
  let eo = -1;
  for (let i = tail.length - 22; i >= 0; i--) {
    if (tail[i] === 0x50 && tail[i + 1] === 0x4B && tail[i + 2] === 5 && tail[i + 3] === 6) { eo = i; break; }
  }
  if (eo < 0) throw new UserError('ไฟล์เสียหรือไม่ครบ (ไม่พบสารบัญของไฟล์ zip)');
  let count = u16(tail, eo + 10), cdSize = u32(tail, eo + 12), cdOff = u32(tail, eo + 16);
  if (count === 0xFFFF || cdSize === 0xFFFFFFFF || cdOff === 0xFFFFFFFF) {
    const lo = eo - 20;
    if (lo < 0 || u32(tail, lo) !== 0x07064B50) throw new UserError('ไฟล์ zip64 นี้อ่านไม่ได้');
    const off64 = u64(tail, lo + 8);
    const rec = await readBytes(file, off64, off64 + 56);
    count = u64(rec, 32); cdSize = u64(rec, 40); cdOff = u64(rec, 48);
  }
  const cd = await readBytes(file, cdOff, cdOff + cdSize);
  const entries = [], byName = new Map();
  let p = 0;
  for (let i = 0; i < count; i++) {
    if (u32(cd, p) !== 0x02014B50) throw new UserError('ไฟล์เสีย (สารบัญ zip ผิดรูปแบบ)');
    const flags = u16(cd, p + 8), method = u16(cd, p + 10);
    let crc = u32(cd, p + 16), csize = u32(cd, p + 20), usize = u32(cd, p + 24);
    const nl = u16(cd, p + 28), xl = u16(cd, p + 30), cl = u16(cd, p + 32);
    let lho = u32(cd, p + 42);
    const nameBytes = cd.subarray(p + 46, p + 46 + nl);
    const name = new TextDecoder(flags & 0x800 ? 'utf-8' : 'latin1').decode(nameBytes);
    let zip64 = false;
    if (usize === 0xFFFFFFFF || csize === 0xFFFFFFFF || lho === 0xFFFFFFFF) {
      zip64 = true;
      let q = p + 46 + nl;
      const qe = q + xl;
      while (q + 4 <= qe) {
        const id = u16(cd, q), len = u16(cd, q + 2);
        if (id === 1) {
          let f = q + 4;
          if (usize === 0xFFFFFFFF) { usize = u64(cd, f); f += 8; }
          if (csize === 0xFFFFFFFF) { csize = u64(cd, f); f += 8; }
          if (lho === 0xFFFFFFFF) { lho = u64(cd, f); f += 8; }
        }
        q += 4 + len;
      }
    }
    const rec = cd.slice(p, p + 46 + nl + xl + cl);
    const e = { name, flags, method, crc, csize, usize, lho, rec, zip64, dataStart: -1 };
    entries.push(e);
    byName.set(name.toLowerCase(), e);
    p += 46 + nl + xl + cl;
  }
  return { file, entries, byName, cdOff };
}

function zipEntry(zip, name) {
  return zip.byName.get(name.replace(/^\/+/, '').toLowerCase()) || null;
}

async function zipDataStart(zip, e) {
  if (e.dataStart < 0) {
    const h = await readBytes(zip.file, e.lho, e.lho + 30);
    if (u32(h, 0) !== 0x04034B50) throw new UserError('ไฟล์เสีย (ส่วน ' + e.name + ')');
    e.dataStart = e.lho + 30 + u16(h, 26) + u16(h, 28);
  }
  return e.dataStart;
}

// The uncompressed bytes of an entry as a stream of Uint8Array chunks.
async function zipStream(zip, e) {
  const start = await zipDataStart(zip, e);
  const raw = zip.file.slice(start, start + e.csize).stream();
  if (e.method === 0) return raw;
  if (e.method !== 8) throw new UserError('ส่วน ' + e.name + ' ของไฟล์ถูกบีบอัดด้วยวิธีที่อ่านไม่ได้');
  return raw.pipeThrough(new DecompressionStream('deflate-raw'));
}

async function* streamChunks(stream) {
  const rd = stream.getReader();
  try {
    for (;;) {
      const { done, value } = await rd.read();
      if (done) return;
      if (value.length > 0) yield value;
    }
  } finally {
    rd.releaseLock();
  }
}

// The chunks of a stream joined into blocks of at least min bytes (fewer, larger pieces to scan).
async function* bigChunks(stream, min) {
  let parts = [], n = 0;
  for await (const c of streamChunks(stream)) {
    parts.push(c);
    n += c.length;
    if (n >= min) {
      yield joinChunks(parts, n);
      parts = [];
      n = 0;
    }
  }
  if (n > 0) yield joinChunks(parts, n);
}

function joinChunks(parts, n) {
  if (parts.length === 1) return parts[0];
  const out = new Uint8Array(n);
  let p = 0;
  for (const c of parts) { out.set(c, p); p += c.length; }
  return out;
}

async function zipReadAll(zip, e) {
  const parts = [];
  let n = 0;
  for await (const c of streamChunks(await zipStream(zip, e))) { parts.push(c); n += c.length; }
  const out = new Uint8Array(n);
  let p = 0;
  for (const c of parts) { out.set(c, p); p += c.length; }
  return out;
}

// Compresses what produce(emit) emits. Returns {chunks, crc, csize, usize}.
async function deflateChunks(produce) {
  const cs = new CompressionStream('deflate-raw');
  const w = cs.writable.getWriter();
  const chunks = [];
  let csize = 0, crc = 0, usize = 0;
  const reading = (async () => {
    for await (const c of streamChunks(cs.readable)) { chunks.push(c); csize += c.length; }
  })();
  try {
    await produce(async (chunk) => {
      if (chunk.length === 0) return;
      crc = crc32(crc, chunk);
      usize += chunk.length;
      await w.write(chunk);
    });
    await w.close();
  } catch (err) {
    try { await w.abort(err); } catch (e2) { /* already failed */ }
    try { await reading; } catch (e3) { /* the reader fails with the same error */ }
    throw err;
  }
  await reading;
  return { chunks, crc, csize, usize };
}

function le16(v) { return [v & 255, (v >>> 8) & 255]; }
function le32(v) { return [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]; }

// A new zip file: the entries of zip in their order, entries named in repl replaced by
// {chunks, crc, csize, usize} (deflate). Returns a Blob.
function zipWrite(zip, repl, mime) {
  const order = zip.entries.slice().sort((a, b) => a.lho - b.lho);
  const ends = new Map();
  for (let i = 0; i < order.length; i++) ends.set(order[i], i + 1 < order.length ? order[i + 1].lho : zip.cdOff);
  const parts = [], cds = [];
  let off = 0;
  for (const e of order) {
    const r = repl.get(e.name);
    const nameBytes = e.rec.subarray(46, 46 + u16(e.rec, 28));
    if (off > 0xFFFFFFFF) throw new UserError('ไฟล์ใหม่ใหญ่เกิน 4 GB เขียนไม่ได้');
    if (r !== undefined) {
      if (r.csize > 0xFFFFFFFF || r.usize > 0xFFFFFFFF) throw new UserError('ชีต ' + e.name + ' ใหญ่เกิน 4 GB เขียนไม่ได้');
      const flags = e.flags & 0x800;
      const time = u16(e.rec, 12), date = u16(e.rec, 14);
      const lh = new Uint8Array(30 + nameBytes.length);
      lh.set([0x50, 0x4B, 3, 4, ...le16(20), ...le16(flags), ...le16(8), ...le16(time), ...le16(date),
        ...le32(r.crc), ...le32(r.csize), ...le32(r.usize), ...le16(nameBytes.length), ...le16(0)]);
      lh.set(nameBytes, 30);
      const cr = new Uint8Array(46 + nameBytes.length);
      cr.set([0x50, 0x4B, 1, 2, ...le16(u16(e.rec, 4)), ...le16(20), ...le16(flags), ...le16(8), ...le16(time),
        ...le16(date), ...le32(r.crc), ...le32(r.csize), ...le32(r.usize), ...le16(nameBytes.length), ...le16(0),
        ...le16(0), ...le16(0), ...le16(u16(e.rec, 36)), ...le32(u32(e.rec, 38)), ...le32(off)]);
      cr.set(nameBytes, 46);
      parts.push(lh, ...r.chunks);
      cds.push(cr);
      off += lh.length + r.csize;
    } else {
      if (e.zip64) throw new UserError('ไฟล์นี้ใช้รูปแบบ zip64 ซึ่งเขียนกลับไม่ได้');
      const end = ends.get(e);
      parts.push(zip.file.slice(e.lho, end));
      const cr = e.rec.slice();
      cr.set(le32(off), 42);
      cds.push(cr);
      off += end - e.lho;
    }
  }
  let cdSize = 0;
  for (const c of cds) cdSize += c.length;
  if (off > 0xFFFFFFFF || order.length > 0xFFFF) throw new UserError('ไฟล์ใหม่ใหญ่เกินไป เขียนไม่ได้');
  const eocd = new Uint8Array([0x50, 0x4B, 5, 6, 0, 0, 0, 0, ...le16(order.length), ...le16(order.length),
    ...le32(cdSize), ...le32(off), 0, 0]);
  return new Blob([...parts, ...cds, eocd], { type: mime || 'application/octet-stream' });
}
