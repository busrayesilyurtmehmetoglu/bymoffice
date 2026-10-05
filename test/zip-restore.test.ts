// An entry whose content is what the archive had again is written as the archive had it:
// compressed bytes, extra fields, place in the archive. A package that was changed, saved
// and changed back is then the file that was opened, byte for byte.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeXlsx } from '../src/formats/xlsx/write';
import { OpcPackage } from '../src/opc';
import { ZipArchive, open } from '../src/index';
import { crc32 } from '../src/zip';
import type { Workbook } from '../src/index';
import { History } from '../src/ops/ops';
import { deleteSheet, insertSheet } from '../src/ops/sheets';
import { buildXlsx, utf8, worksheetXml } from './fixtures';

/**
 * An archive as another program writes it: every entry stored without compression although it
 * would compress, with an extra field in both headers, a comment and a date of its own.
 */
function foreignZip(entries: [name: string, data: Uint8Array][]): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  entries.forEach(([name, data], index) => {
    const nameBytes = utf8(name);
    const extra = new Uint8Array([0xfe, 0xca, 2, 0, index & 0xff, 7]);
    const comment = utf8(`entry ${index}`);
    const sum = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length + extra.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(10, 0x6000 + index, true);
    lv.setUint16(12, 0x5a21, true);
    lv.setUint32(14, sum, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, extra.length, true);
    local.set(nameBytes, 30);
    local.set(extra, 30 + nameBytes.length);
    const header = new Uint8Array(46 + nameBytes.length + extra.length + comment.length);
    const cv = new DataView(header.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 0x031e, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(12, 0x6000 + index, true);
    cv.setUint16(14, 0x5a21, true);
    cv.setUint32(16, sum, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, extra.length, true);
    cv.setUint16(32, comment.length, true);
    cv.setUint32(38, 0x81a40000, true);
    cv.setUint32(42, offset, true);
    header.set(nameBytes, 46);
    header.set(extra, 46 + nameBytes.length);
    header.set(comment, 46 + nameBytes.length + extra.length);
    chunks.push(local, data);
    central.push(header);
    offset += local.length + data.length;
  });
  const size = central.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, size, true);
  ev.setUint32(16, offset, true);
  return new Uint8Array(Buffer.concat([...chunks, ...central, end]));
}

/** The same package written by another program. */
async function repack(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  const entries: [string, Uint8Array][] = [];
  for (const name of zip.names()) entries.push([name, await zip.read(name)]);
  return foreignZip(entries);
}

const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0;
const TEXT = utf8('the quick brown fox '.repeat(40));

test('zip: an archive of another program is written back as it is', async () => {
  const bytes = foreignZip([
    ['a.txt', TEXT],
    ['b/c.txt', utf8('second')],
    ['d.txt', new Uint8Array(0)],
  ]);
  assert.ok(same(await ZipArchive.read(bytes).write(), bytes));
});

test('zip: an entry that is given the content it had is the entry of the archive again', async () => {
  const bytes = foreignZip([
    ['a.txt', TEXT],
    ['b/c.txt', utf8('second')],
    ['d.txt', utf8('third')],
  ]);
  const zip = ZipArchive.read(bytes);
  zip.set('b/c.txt', utf8('changed'));
  zip.set('a.txt', utf8('changed too'));
  const changed = await zip.write();
  assert.ok(!same(changed, bytes));
  assert.deepEqual(await ZipArchive.read(changed).read('b/c.txt'), utf8('changed'));
  // The same content from another array: it is the content that counts.
  zip.set('b/c.txt', utf8('second'));
  zip.set('a.txt', TEXT.slice());
  assert.ok(same(await zip.write(), bytes));
  // And once more: a write does not use the original up.
  zip.set('a.txt', utf8('x'));
  await zip.write();
  zip.set('a.txt', TEXT.slice());
  assert.ok(same(await zip.write(), bytes));
  // Set to what it is without a write in between.
  zip.set('d.txt', utf8('third'));
  assert.ok(same(await zip.write(), bytes));
  assert.deepEqual(await zip.read('a.txt'), TEXT);
});

test('zip: an entry that was deleted and comes back takes its place in the archive again', async () => {
  const bytes = foreignZip([
    ['a.txt', TEXT],
    ['b/c.txt', utf8('second')],
    ['d.txt', utf8('third')],
  ]);
  const zip = ZipArchive.read(bytes);
  zip.delete('a.txt');
  zip.set('new.txt', utf8('new'));
  const without = await zip.write();
  assert.deepEqual(ZipArchive.read(without).names(), ['b/c.txt', 'd.txt', 'new.txt']);
  zip.set('a.txt', TEXT);
  assert.deepEqual(ZipArchive.read(await zip.write()).names(), ['a.txt', 'b/c.txt', 'd.txt', 'new.txt']);
  zip.delete('new.txt');
  assert.ok(same(await zip.write(), bytes));
  // Back with other content: its place, and its own headers.
  zip.delete('b/c.txt');
  await zip.write();
  zip.set('b/c.txt', utf8('other'));
  const other = ZipArchive.read(await zip.write());
  assert.deepEqual(other.names(), ['a.txt', 'b/c.txt', 'd.txt']);
  assert.deepEqual(await other.read('b/c.txt'), utf8('other'));
  assert.equal(other.entry('b/c.txt')!.extra!.length, 0);
});

test('zip: content of the same length and another checksum, and of the same checksum only, is not taken for the original', async () => {
  const bytes = foreignZip([['a.txt', utf8('abcdef')]]);
  const zip = ZipArchive.read(bytes);
  zip.set('a.txt', utf8('abcdeg'));
  const out = ZipArchive.read(await zip.write());
  assert.deepEqual(await out.read('a.txt'), utf8('abcdeg'));
  assert.ok(!same(await zip.write(), bytes));
});

const sheets = () =>
  buildXlsx({
    sheets: ['One', 'Two', 'Three'].map((name, index) => ({
      name,
      xml: worksheetXml(`<sheetData><row r="1"><c r="A1"><v>${index + 1}</v></c></row></sheetData>`),
    })),
  });

async function load(bytes: Uint8Array): Promise<{ wb: Workbook; history: History; save(): Promise<Uint8Array> }> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return {
    wb,
    history: new History(wb),
    async save() {
      await writeXlsx(file.package!, wb);
      return file.package!.save();
    },
  };
}

test('package: a sheet added, saved, taken back and saved again leaves the file byte for byte', async () => {
  for (const make of [sheets, async () => repack(await sheets())]) {
    const bytes = await make();
    const doc = await load(bytes);
    assert.ok(same(await doc.save(), bytes), 'saved as opened');
    doc.history.apply(insertSheet(doc.wb, 1, { name: 'New' }), undefined, undefined);
    const saved = await doc.save();
    assert.equal((await (await open(saved)).workbook()).sheets.length, 4);
    doc.history.undo();
    assert.ok(same(await doc.save(), bytes), 'the whole archive');
  }
});

test('package: a sheet of the file deleted, saved, brought back and saved again leaves the file byte for byte', async () => {
  for (const make of [sheets, async () => repack(await sheets())]) {
    const bytes = await make();
    const doc = await load(bytes);
    doc.history.apply(deleteSheet(doc.wb, 1), undefined, undefined);
    const saved = await doc.save();
    assert.deepEqual((await (await open(saved)).workbook()).sheets.map((sheet) => sheet.name), ['One', 'Three']);
    doc.history.undo();
    assert.ok(same(await doc.save(), bytes), 'the whole archive');
    // And again, the other way round.
    doc.history.redo();
    await doc.save();
    doc.history.undo();
    assert.ok(same(await doc.save(), bytes), 'after redo and undo');
  }
});

test('package: content types that are the ones of the file again are written with the bytes of the file', async () => {
  const bytes = await repack(await sheets());
  const pkg = await OpcPackage.open(bytes);
  pkg.writePart('/xl/extra.bin', utf8('x'), 'application/x-test');
  const withPart = await pkg.save();
  assert.ok(new TextDecoder().decode(await ZipArchive.read(withPart).read('[Content_Types].xml')).includes('application/x-test'));
  pkg.deletePart('/xl/extra.bin');
  assert.ok(same(await pkg.save(), bytes));
  // An override that goes and comes back at the end of the list: the same set, the bytes of the file.
  const sheet = await pkg.readPart('/xl/worksheets/sheet2.xml');
  const type = pkg.contentTypeOf('/xl/worksheets/sheet2.xml');
  pkg.deletePart('/xl/worksheets/sheet2.xml');
  await pkg.save();
  pkg.writePart('/xl/worksheets/sheet2.xml', sheet, type);
  assert.ok(same(await pkg.save(), bytes));
  // Another type for the same part is another set.
  pkg.deletePart('/xl/worksheets/sheet2.xml');
  pkg.writePart('/xl/worksheets/sheet2.xml', sheet, 'application/x-other');
  const other = await pkg.save();
  assert.ok(new TextDecoder().decode(await ZipArchive.read(other).read('[Content_Types].xml')).includes('application/x-other'));
});
