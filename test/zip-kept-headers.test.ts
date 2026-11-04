// The headers of an entry are what the archive's writer chose: the common zip tool puts
// "version needed" 1.0 on directory entries, writers that stream (Java, .NET, Python on a pipe)
// leave the sizes out of the local header and put a data descriptor after the data. An entry whose
// stored bytes are the archive's own keeps all of that, so an untouched package is the file that
// was opened, byte for byte.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, createBlank, open } from '../src/index';
import { crc32 } from '../src/zip';
import { utf8 } from './fixtures';

interface Item {
  name: string;
  data: Uint8Array;
  /** Of the local header; the central one when not given. */
  local?: number;
  central: number;
  /** Sizes and checksum after the data instead of in the local header. */
  descriptor?: 'signed' | 'bare';
}

/** A stored-only archive whose headers carry the given "version needed" values. */
function zipOf(items: Item[]): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const { name, data, local: localVersion, central: centralVersion, descriptor } of items) {
    const nameBytes = utf8(name);
    const sum = crc32(data);
    const flags = descriptor ? 0x0008 : 0;
    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, localVersion ?? centralVersion, true);
    lv.setUint16(6, flags, true);
    lv.setUint16(10, 0x6000, true);
    lv.setUint16(12, 0x5a21, true);
    if (!descriptor) {
      lv.setUint32(14, sum, true);
      lv.setUint32(18, data.length, true);
      lv.setUint32(22, data.length, true);
    }
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    const header = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(header.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 0x031e, true);
    cv.setUint16(6, centralVersion, true);
    cv.setUint16(8, flags, true);
    cv.setUint16(12, 0x6000, true);
    cv.setUint16(14, 0x5a21, true);
    cv.setUint32(16, sum, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(38, name.endsWith('/') ? 0x41ed0010 : 0x81a40000, true);
    cv.setUint32(42, offset, true);
    header.set(nameBytes, 46);
    const tail = new Uint8Array(descriptor === 'signed' ? 16 : descriptor ? 12 : 0);
    if (descriptor) {
      const dv = new DataView(tail.buffer);
      const at = descriptor === 'signed' ? 4 : 0;
      if (at) dv.setUint32(0, 0x08074b50, true);
      dv.setUint32(at, sum, true);
      dv.setUint32(at + 4, data.length, true);
      dv.setUint32(at + 8, data.length, true);
    }
    chunks.push(local, data, tail);
    central.push(header);
    offset += local.length + data.length + tail.length;
  }
  const size = central.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, items.length, true);
  ev.setUint16(10, items.length, true);
  ev.setUint32(12, size, true);
  ev.setUint32(16, offset, true);
  return new Uint8Array(Buffer.concat([...chunks, ...central, end]));
}

/** "Version needed" of the local and the central header of an entry of written bytes. */
function versionsOf(bytes: Uint8Array, name: string): { local: number; central: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = bytes.length - 22;
  assert.equal(view.getUint32(eocd, true), 0x06054b50);
  let p = view.getUint32(eocd + 16, true);
  const count = view.getUint16(eocd + 10, true);
  for (let n = 0; n < count; n++) {
    const nameLength = view.getUint16(p + 28, true);
    const next = p + 46 + nameLength + view.getUint16(p + 30, true) + view.getUint16(p + 32, true);
    if (Buffer.from(bytes.subarray(p + 46, p + 46 + nameLength)).toString('utf8') === name) {
      const local = view.getUint32(p + 42, true);
      return { local: view.getUint16(local + 4, true), central: view.getUint16(p + 6, true) };
    }
    p = next;
  }
  throw new Error(`no entry ${name}`);
}

const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0;
const TEXT = utf8('the quick brown fox '.repeat(40));

test('zip: the "version needed" of every header is written back as the archive had it', async () => {
  const bytes = zipOf([
    { name: 'word/', data: new Uint8Array(0), central: 10 },
    { name: 'word/a.txt', data: TEXT, central: 10 },
    // Headers that do not agree, and a value of a writer that announced ZIP64 support.
    { name: 'word/b.txt', data: utf8('second'), local: 10, central: 20 },
    { name: 'c.txt', data: utf8('third'), central: 45 },
  ]);
  assert.ok(same(await ZipArchive.read(bytes).write(), bytes));
});

test('zip: an entry compressed again is written with 2.0; the others keep their values', async () => {
  const bytes = zipOf([
    { name: 'word/', data: new Uint8Array(0), central: 10 },
    { name: 'word/a.txt', data: TEXT, central: 10 },
    { name: 'word/b.txt', data: utf8('second'), local: 10, central: 20 },
  ]);
  const zip = ZipArchive.read(bytes);
  zip.set('word/a.txt', utf8('changed '.repeat(40)));
  zip.set('word/b.txt', utf8('x'));
  zip.set('new.txt', utf8('added'));
  const written = await zip.write();
  assert.deepEqual(versionsOf(written, 'word/'), { local: 10, central: 10 });
  // Deflated now: 1.0 would tell a reader it need not inflate.
  assert.deepEqual(versionsOf(written, 'word/a.txt'), { local: 20, central: 20 });
  // Stored, but the bytes are the engine's, and so are the headers.
  assert.deepEqual(versionsOf(written, 'word/b.txt'), { local: 20, central: 20 });
  assert.deepEqual(versionsOf(written, 'new.txt'), { local: 20, central: 20 });
  assert.deepEqual(await ZipArchive.read(written).read('word/a.txt'), utf8('changed '.repeat(40)));

  // Given its content back, an entry is the archive's own again, headers included.
  zip.set('word/a.txt', TEXT.slice());
  zip.set('word/b.txt', utf8('second'));
  zip.delete('new.txt');
  assert.ok(same(await zip.write(), bytes));
});

test('package: a docx whose archive has directory entries is saved byte for byte', async () => {
  // As the zip tool writes a folder: directory entries, everything at 1.0 where nothing needs more.
  const blank = ZipArchive.read(await createBlank('docx'));
  const items: Item[] = [];
  const folders = new Set<string>();
  for (const name of blank.names()) {
    const parts = name.split('/');
    for (let n = 1; n < parts.length; n++) {
      const folder = `${parts.slice(0, n).join('/')}/`;
      if (!folders.has(folder)) {
        folders.add(folder);
        items.push({ name: folder, data: new Uint8Array(0), central: 10 });
      }
    }
    items.push({ name, data: await blank.read(name), central: 10 });
  }
  assert.ok(folders.size > 0);
  const bytes = zipOf(items);
  const file = await open(bytes);
  assert.equal(file.format, 'docx');
  assert.ok(same(await file.save(), bytes));
});

test('zip: data descriptors, with and without their signature, are written back as they were', async () => {
  const bytes = zipOf([
    { name: 'a.txt', data: TEXT, central: 20, descriptor: 'signed' },
    { name: 'b.txt', data: utf8('second'), central: 20, descriptor: 'bare' },
    { name: 'c.txt', data: utf8('third'), central: 20 },
    { name: 'd.txt', data: new Uint8Array(0), central: 20, descriptor: 'signed' },
  ]);
  const zip = ZipArchive.read(bytes);
  assert.ok(same(await zip.write(), bytes));
  assert.deepEqual(await zip.read('a.txt'), TEXT);

  // A changed entry is written with its sizes in the header and without a descriptor.
  zip.set('a.txt', utf8('changed'));
  const written = await zip.write();
  const again = ZipArchive.read(written);
  assert.equal((again.entry('a.txt')?.flags ?? 0) & 0x0008, 0);
  assert.equal((again.entry('b.txt')?.flags ?? 0) & 0x0008, 0x0008);
  assert.deepEqual(await again.read('a.txt'), utf8('changed'));
  assert.deepEqual(await again.read('b.txt'), utf8('second'));
  assert.deepEqual(await again.read('d.txt'), new Uint8Array(0));
  zip.set('a.txt', TEXT.slice());
  assert.ok(same(await zip.write(), bytes));
});

test('zip: a descriptor that does not say what the central header says is not copied', async () => {
  const bytes = zipOf([
    { name: 'a.txt', data: TEXT, central: 20, descriptor: 'bare' },
    { name: 'b.txt', data: utf8('second'), central: 20 },
  ]);
  // The checksum of the descriptor no longer agrees: what follows the data is not taken for one.
  const broken = bytes.slice();
  const at = 30 + 'a.txt'.length + TEXT.length;
  broken[at] ^= 0xff;
  const zip = ZipArchive.read(broken);
  assert.equal(zip.entry('a.txt')?.localHeader, undefined);
  const written = await zip.write();
  const again = ZipArchive.read(written);
  assert.equal((again.entry('a.txt')?.flags ?? 0) & 0x0008, 0);
  assert.deepEqual(await again.read('a.txt'), TEXT);
  assert.deepEqual(await again.read('b.txt'), utf8('second'));
});

test('zip: headers that disagree on the data descriptor are written anew, consistent', async () => {
  const bytes = zipOf([
    { name: 'a.txt', data: TEXT, central: 20, descriptor: 'signed' },
    { name: 'b.txt', data: utf8('second'), central: 20 },
  ]);
  // The central header of a.txt says there is no descriptor; its local header says there is.
  const broken = bytes.slice();
  const view = new DataView(broken.buffer);
  const central = view.getUint32(broken.length - 22 + 16, true);
  view.setUint16(central + 8, 0, true);
  const zip = ZipArchive.read(broken);
  assert.equal(zip.entry('a.txt')?.localHeader, undefined);
  const written = await zip.write();
  const wv = new DataView(written.buffer, written.byteOffset, written.byteLength);
  assert.equal(wv.getUint16(6, true) & 0x0008, 0, 'the local header no longer announces a descriptor');
  const again = ZipArchive.read(written);
  assert.deepEqual(await again.read('a.txt'), TEXT);
  assert.deepEqual(await again.read('b.txt'), utf8('second'));
});
