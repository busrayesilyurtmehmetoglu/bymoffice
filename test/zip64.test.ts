import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { utf8 } from './fixtures';

const MAX32 = 0xffffffff;

const setUint64 = (view: DataView, offset: number, value: number) => {
  view.setUint32(offset, value >>> 0, true);
  view.setUint32(offset + 4, Math.floor(value / 0x100000000), true);
};

/**
 * Rewrites the directory of an archive produced by ZipArchive.write() (no extra
 * fields, no comments) the way a ZIP64 writer would: `fields` chooses which
 * values move from the central header into the ZIP64 extra field.
 */
function toZip64(bytes: Uint8Array, fields: { sizes: boolean; offset: boolean }, leadingExtra = false): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = bytes.length - 22;
  const count = view.getUint16(eocd + 10, true);
  const centralOffset = view.getUint32(eocd + 16, true);

  const headers: Uint8Array[] = [];
  let p = centralOffset;
  for (let n = 0; n < count; n++) {
    const nameLength = view.getUint16(p + 28, true);
    const values: number[] = [];
    if (fields.sizes) values.push(view.getUint32(p + 24, true), view.getUint32(p + 20, true));
    if (fields.offset) values.push(view.getUint32(p + 42, true));
    // An unrelated field in front proves the reader walks the list instead of assuming a position.
    const other = leadingExtra ? [0xfe, 0xca, 2, 0, 9, 9] : [];
    const extra = new Uint8Array(other.length + 4 + values.length * 8);
    extra.set(other);
    const xv = new DataView(extra.buffer);
    xv.setUint16(other.length, 1, true);
    xv.setUint16(other.length + 2, values.length * 8, true);
    values.forEach((value, i) => setUint64(xv, other.length + 4 + i * 8, value));

    const header = new Uint8Array(46 + nameLength + extra.length);
    header.set(bytes.subarray(p, p + 46 + nameLength));
    header.set(extra, 46 + nameLength);
    const hv = new DataView(header.buffer);
    hv.setUint16(6, 45, true);
    if (fields.sizes) {
      hv.setUint32(20, MAX32, true);
      hv.setUint32(24, MAX32, true);
    }
    if (fields.offset) hv.setUint32(42, MAX32, true);
    hv.setUint16(30, extra.length, true);
    headers.push(header);
    p += 46 + nameLength;
  }
  const centralSize = headers.reduce((sum, h) => sum + h.length, 0);

  const tail = new Uint8Array(56 + 20 + 22);
  const tv = new DataView(tail.buffer);
  tv.setUint32(0, 0x06064b50, true);
  setUint64(tv, 4, 44);
  tv.setUint16(12, 45, true);
  tv.setUint16(14, 45, true);
  setUint64(tv, 24, count);
  setUint64(tv, 32, count);
  setUint64(tv, 40, centralSize);
  setUint64(tv, 48, centralOffset);
  tv.setUint32(56, 0x07064b50, true);
  setUint64(tv, 64, centralOffset + centralSize);
  tv.setUint32(72, 1, true);
  tv.setUint32(76, 0x06054b50, true);
  tv.setUint16(84, 0xffff, true);
  tv.setUint16(86, 0xffff, true);
  tv.setUint32(88, MAX32, true);
  tv.setUint32(92, MAX32, true);

  const out = new Uint8Array(centralOffset + centralSize + tail.length);
  out.set(bytes.subarray(0, centralOffset));
  let q = centralOffset;
  for (const header of headers) {
    out.set(header, q);
    q += header.length;
  }
  out.set(tail, q);
  return out;
}

async function sample(): Promise<{ bytes: Uint8Array; content: Map<string, Uint8Array> }> {
  const content = new Map([
    ['[Content_Types].xml', utf8('<Types/>'.repeat(200))],
    ['xl/çalışma.xml', utf8('<row>şğü</row>'.repeat(300))],
    ['bin/raw.bin', new Uint8Array([9, 8, 7])],
  ]);
  const zip = new ZipArchive();
  for (const [name, data] of content) zip.set(name, data);
  return { bytes: await zip.write(), content };
}

/** Runs `body` as if the engine had no built-in compression streams. */
async function withoutNativeStreams<T>(body: () => Promise<T>): Promise<T> {
  const scope = globalThis as Record<string, unknown>;
  const saved = { c: scope.CompressionStream, d: scope.DecompressionStream };
  scope.CompressionStream = undefined;
  scope.DecompressionStream = undefined;
  try {
    return await body();
  } finally {
    scope.CompressionStream = saved.c;
    scope.DecompressionStream = saved.d;
  }
}

test('zip64: sizes and offsets are taken from the ZIP64 extra field', async () => {
  const { bytes, content } = await sample();
  const variants = [
    toZip64(bytes, { sizes: true, offset: true }),
    toZip64(bytes, { sizes: false, offset: true }),
    toZip64(bytes, { sizes: true, offset: false }, true),
    toZip64(bytes, { sizes: false, offset: false }),
  ];
  for (const variant of variants) {
    const zip = ZipArchive.read(variant);
    assert.deepEqual(zip.names(), [...content.keys()]);
    for (const [name, data] of content) assert.deepEqual(await zip.read(name), data, name);

    // Saving produces a plain archive again, with the stored bytes untouched and
    // every extra field but the ZIP64 one kept.
    const saved = await zip.write();
    if (variant === variants[2]) {
      assert.deepEqual(ZipArchive.read(saved).entry('bin/raw.bin')?.extra, new Uint8Array([0xfe, 0xca, 2, 0, 9, 9]));
      assert.equal(saved.length, bytes.length + 6 * content.size);
    } else assert.deepEqual(saved, bytes);
  }
});

/** Builds an archive by hand so that headers can carry what ZipArchive.set() never produces. */
function rawArchive(
  entries: { name: Uint8Array; data: Uint8Array; flags?: number; extra?: number[]; localExtra?: number[]; comment?: string }[],
  options: { comment?: string; announcedCount?: number } = {},
): Uint8Array {
  const crcOf = (data: Uint8Array) => {
    let c = 0xffffffff;
    for (const byte of data) {
      c ^= byte;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    return (c ^ 0xffffffff) >>> 0;
  };
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  for (const entry of entries) {
    const localExtra = entry.localExtra ?? [];
    const extra = entry.extra ?? [];
    const comment = utf8(entry.comment ?? '');
    const local = new Uint8Array(30 + entry.name.length + localExtra.length + entry.data.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, entry.flags ?? 0, true);
    lv.setUint32(14, crcOf(entry.data), true);
    lv.setUint32(18, entry.data.length, true);
    lv.setUint32(22, entry.data.length, true);
    lv.setUint16(26, entry.name.length, true);
    lv.setUint16(28, localExtra.length, true);
    local.set(entry.name, 30);
    local.set(localExtra, 30 + entry.name.length);
    local.set(entry.data, 30 + entry.name.length + localExtra.length);

    const central = new Uint8Array(46 + entry.name.length + extra.length + comment.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 0x031e, true); // made on Unix
    cv.setUint16(6, 20, true);
    cv.setUint16(8, entry.flags ?? 0, true);
    cv.setUint32(16, crcOf(entry.data), true);
    cv.setUint32(20, entry.data.length, true);
    cv.setUint32(24, entry.data.length, true);
    cv.setUint16(28, entry.name.length, true);
    cv.setUint16(30, extra.length, true);
    cv.setUint16(32, comment.length, true);
    cv.setUint16(36, 1, true); // internal attributes: text file
    cv.setUint32(38, 0x81a40000, true);
    cv.setUint32(42, offset, true);
    central.set(entry.name, 46);
    central.set(extra, 46 + entry.name.length);
    central.set(comment, 46 + entry.name.length + extra.length);

    locals.push(local);
    centrals.push(central);
    offset += local.length;
  }
  const comment = utf8(options.comment ?? '');
  const centralSize = centrals.reduce((sum, c) => sum + c.length, 0);
  const eocd = new Uint8Array(22 + comment.length);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, options.announcedCount ?? entries.length, true);
  ev.setUint16(10, options.announcedCount ?? entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, comment.length, true);
  eocd.set(comment, 22);

  const out = new Uint8Array(offset + centralSize + eocd.length);
  let p = 0;
  for (const chunk of [...locals, ...centrals, eocd]) {
    out.set(chunk, p);
    p += chunk.length;
  }
  return out;
}

test('zip: comments, extra fields, attributes and non-UTF-8 names survive a save', async () => {
  const cp437Name = new Uint8Array([0x81, 0x2e, 0x74, 0x78, 0x74]); // "ü.txt" in CP437, invalid as UTF-8
  const bytes = rawArchive(
    [
      { name: cp437Name, data: utf8('one'), extra: [0xfe, 0xca, 2, 0, 1, 2], localExtra: [0x20, 0xa2, 1, 0, 0], comment: 'entry comment' },
      { name: utf8('ğ.txt'), data: utf8('two'), flags: 0x0800 },
      { name: utf8('plain.txt'), data: utf8('three') },
    ],
    { comment: 'archive comment' },
  );
  const zip = ZipArchive.read(bytes);
  assert.deepEqual(await zip.read('plain.txt'), utf8('three'));

  // Nothing touched: the archive is reproduced byte for byte.
  assert.deepEqual(await zip.write(), bytes);

  // Replacing content keeps the name bytes, comment and attributes of that entry.
  const legacyName = zip.names()[0];
  zip.set(legacyName, utf8('ONE'));
  zip.set('ğ.txt', utf8('TWO'));
  const saved = await zip.write();
  const again = ZipArchive.read(saved);
  const first = again.entry(legacyName)!;
  assert.deepEqual(first.rawName, cp437Name);
  assert.equal(first.flags & 0x0800, 0);
  assert.equal(again.entry('ğ.txt')!.flags & 0x0800, 0x0800);
  assert.deepEqual(first.comment, utf8('entry comment'));
  assert.deepEqual([first.versionMadeBy, first.internalAttrs, first.externalAttrs], [0x031e, 1, 0x81a40000]);
  assert.deepEqual(await again.read(legacyName), utf8('ONE'));
  assert.deepEqual(await again.read('ğ.txt'), utf8('TWO'));
  assert.deepEqual(saved.subarray(saved.length - 15), utf8('archive comment'));
});

test('zip: entries beyond a wrapped 16-bit count are still read', async () => {
  const entries = ['a.txt', 'b.txt', 'c.txt'].map((name) => ({ name: utf8(name), data: utf8(name) }));
  const zip = ZipArchive.read(rawArchive(entries, { announcedCount: 2 }));
  assert.deepEqual(zip.names(), ['a.txt', 'b.txt', 'c.txt']);
  assert.deepEqual(ZipArchive.read(await zip.write()).names(), ['a.txt', 'b.txt', 'c.txt']);
});

test('zip: two entries with the same name are rejected instead of losing one', () => {
  const entries = ['a.txt', 'b.txt', 'a.txt'].map((name, i) => ({ name: utf8(name), data: utf8(`content ${i}`) }));
  assert.throws(() => ZipArchive.read(rawArchive(entries)), /duplicate entry name a\.txt/);
});

test('zip64: more than 65535 entries are written and read back', async () => {
  const zip = new ZipArchive();
  const total = 66000;
  for (let i = 0; i < total; i++) zip.set(`e/${i}`, i % 1000 === 0 ? utf8(`entry ${i}`) : new Uint8Array(0));
  const bytes = await zip.write();

  const view = new DataView(bytes.buffer);
  assert.equal(view.getUint16(bytes.length - 22 + 10, true), 0xffff);
  assert.equal(view.getUint32(bytes.length - 42, true), 0x07064b50);

  const again = ZipArchive.read(bytes);
  assert.equal(again.names().length, total);
  assert.equal(again.names()[total - 1], `e/${total - 1}`);
  assert.deepEqual(await again.read('e/65000'), utf8('entry 65000'));
  assert.deepEqual(await again.read('e/65999'), new Uint8Array(0));
});

test('zip64: exactly 65535 entries without a ZIP64 record are read literally', async () => {
  const zip = new ZipArchive();
  for (let i = 0; i < 0xffff; i++) zip.set(`${i}`, new Uint8Array(0));
  const bytes = await zip.write();
  // Drop the ZIP64 record and locator, as writers that predate ZIP64 would.
  const plain = new Uint8Array(bytes.length - 76);
  plain.set(bytes.subarray(0, bytes.length - 98));
  plain.set(bytes.subarray(bytes.length - 22), bytes.length - 98);
  assert.equal(ZipArchive.read(plain).names().length, 0xffff);
});

test('zip64: damaged directories are rejected with a clear error', async () => {
  const { bytes } = await sample();
  const good = toZip64(bytes, { sizes: true, offset: true });
  const locator = good.length - 42;
  const record = good.length - 98;
  const central = new DataView(bytes.buffer).getUint32(bytes.length - 22 + 16, true);
  const damage = (edit: (view: DataView, copy: Uint8Array) => void) => {
    const copy = good.slice();
    edit(new DataView(copy.buffer), copy);
    return copy;
  };

  // Locator points outside the file, and at something that is not the record.
  assert.throws(() => ZipArchive.read(damage((v) => setUint64(v, locator + 8, good.length))), /Corrupt ZIP archive/);
  assert.throws(() => ZipArchive.read(damage((v) => setUint64(v, locator + 8, 0))), /Corrupt ZIP archive/);
  assert.throws(() => ZipArchive.read(damage((v) => v.setUint32(locator + 12, 0x7fffffff, true))), /out of range/);
  // Record claims a directory that does not fit in the file, or more entries than it can hold.
  assert.throws(() => ZipArchive.read(damage((v) => setUint64(v, record + 40, good.length))), /out of bounds/);
  assert.throws(() => ZipArchive.read(damage((v) => setUint64(v, record + 32, 2 ** 40))), /out of bounds/);
  // ZIP64 extra field renamed, and cut short.
  const nameLength = new DataView(good.buffer).getUint16(central + 28, true);
  const extra = central + 46 + nameLength;
  assert.throws(() => ZipArchive.read(damage((v) => v.setUint16(extra, 0x7777, true))), /extra field is missing/);
  assert.throws(() => ZipArchive.read(damage((v) => v.setUint16(extra + 2, 16, true))), /Corrupt ZIP archive/);
  // Offsets and sizes beyond the end of the file.
  assert.throws(() => ZipArchive.read(damage((v) => setUint64(v, extra + 4 + 16, good.length))), /bad local header/);
  assert.throws(() => ZipArchive.read(damage((v) => setUint64(v, extra + 4 + 8, good.length))), /truncated data/);
});

test('zip: truncated and garbage input is rejected without reading out of bounds', async () => {
  const { bytes } = await sample();
  for (let length = 0; length < bytes.length; length += 7) {
    assert.throws(() => ZipArchive.read(bytes.subarray(0, length)), Error, `length ${length}`);
  }
  // Every single-byte change in the directory either fails cleanly or still parses.
  const central = new DataView(bytes.buffer).getUint32(bytes.length - 22 + 16, true);
  for (let i = central; i < bytes.length; i++) {
    const copy = bytes.slice();
    copy[i] ^= 0xff;
    try {
      const zip = ZipArchive.read(copy);
      for (const name of zip.names()) await zip.read(name).catch((error) => assert.ok(error instanceof Error));
    } catch (error) {
      assert.ok(error instanceof Error && !(error instanceof RangeError), `byte ${i}: ${error}`);
    }
  }
});

test('zip: an entry that inflates beyond its declared size is rejected', async () => {
  const zip = new ZipArchive();
  zip.set('bomb.bin', new Uint8Array(4 * 1024 * 1024));
  const bytes = await zip.write();
  const view = new DataView(bytes.buffer);
  const central = view.getUint32(bytes.length - 22 + 16, true);
  view.setUint32(central + 24, 100, true);

  await assert.rejects(ZipArchive.read(bytes).read('bomb.bin'), /Corrupt ZIP entry/);
  await withoutNativeStreams(() => assert.rejects(ZipArchive.read(bytes).read('bomb.bin'), /Corrupt ZIP entry/));
});

test('zip: works without built-in compression streams and stays compatible with them', async () => {
  const { bytes, content } = await sample();

  const fallbackBytes = await withoutNativeStreams(async () => {
    // Archive written with the built-in streams, read with our own inflate.
    const zip = ZipArchive.read(bytes);
    for (const [name, data] of content) assert.deepEqual(await zip.read(name), data, name);

    const fresh = new ZipArchive();
    for (const [name, data] of content) fresh.set(name, data);
    const written = await fresh.write();
    const again = ZipArchive.read(written);
    for (const [name, data] of content) assert.deepEqual(await again.read(name), data, name);
    assert.equal(again.entry('xl/çalışma.xml')?.method, 8);
    return written;
  });

  // Archive written with our own deflate, read with the built-in streams.
  const zip = ZipArchive.read(fallbackBytes);
  for (const [name, data] of content) assert.deepEqual(await zip.read(name), data, name);
});

test('zip: corrupt deflate data is reported as a corrupt entry in both modes', async () => {
  const zip = new ZipArchive();
  zip.set('a.txt', utf8('lorem ipsum dolor sit amet '.repeat(200)));
  const bytes = await zip.write();
  for (let i = 0; i < 40; i++) bytes[30 + 'a.txt'.length + i] ^= 0x5a;

  await assert.rejects(ZipArchive.read(bytes).read('a.txt'), /Corrupt ZIP entry/);
  await withoutNativeStreams(() => assert.rejects(ZipArchive.read(bytes).read('a.txt'), /Corrupt ZIP entry/));
});
