// The compound file writer: what it writes is read back by the engine's own reader, by the
// reader of test/tools/offcrypto_ref.py (which also checks the order and the colours of the
// directory trees and the counts of the header), and named by file(1) where there is one.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { allocateCfb, writeCfb } from '../src/cfb-write';
import { refOk, refTest } from './reftools';

const fill = (size: number, seed = 1): Uint8Array => new Uint8Array(size).map((_, i) => (i * 7 + seed * 13 + (i >> 9)) & 0xff);
const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

interface Listing {
  version: number;
  sector: number;
  cutoff: number;
  fat_sectors: number;
  difat_sectors: number;
  order: string[];
  storages: string[];
  streams: Record<string, { size: number; sha256: string }>;
}

const listing = (file: Uint8Array): Listing => JSON.parse(new TextDecoder().decode(refOk('cfb-list', {}, file))) as Listing;

/** Both readers find every stream, with its content. */
function sameInBothReaders(streams: { path: string; data: Uint8Array }[], file: Uint8Array, withReference: boolean): Listing | undefined {
  const cfb = CfbFile.read(file);
  for (const { path, data } of streams) {
    assert.ok(cfb.has(path), `the engine's reader finds ${JSON.stringify(path)}`);
    const read = cfb.stream(path);
    assert.equal(read.length, data.length, `the size of ${JSON.stringify(path)}`);
    assert.equal(sha256(read), sha256(data), `the content of ${JSON.stringify(path)}`);
  }
  assert.equal(cfb.entries.filter((entry) => entry.type === 'stream').length, streams.length, 'no stream more than was written');
  if (!withReference) return undefined;
  const listed = listing(file);
  assert.deepEqual(Object.keys(listed.streams).sort(), streams.map((s) => s.path).sort(), 'the streams the reference reader finds');
  for (const { path, data } of streams) assert.deepEqual(listed.streams[path], { size: data.length, sha256: sha256(data) }, `${JSON.stringify(path)} in the reference reader`);
  assert.deepEqual([listed.version, listed.sector, listed.cutoff], [3, 512, 4096], 'version 3, sectors of 512 bytes, a mini stream below 4096 bytes');
  return listed;
}

test('cfb writer: an empty file, and one stream of every size around the edges of the mini stream and of a sector, are read back', () => {
  for (const size of [0, 1, 63, 64, 65, 511, 512, 513, 4095, 4096, 4097, 4608, 100000]) {
    const streams = [{ path: 'Stream', data: fill(size) }];
    const file = writeCfb(streams);
    assert.equal(file.length % 512, 0, `${size} bytes: whole sectors`);
    sameInBothReaders(streams, file, false);
  }
  const empty = writeCfb([]);
  assert.equal(CfbFile.read(empty).entries.length, 0);
});

refTest('cfb writer: streams of 4095, 4096 and 4097 bytes are on the right side of the mini stream, by the reference reader', () => {
  for (const size of [0, 1, 64, 4095, 4096, 4097]) {
    const streams = [
      { path: 'A', data: fill(size, 1) },
      { path: 'Bb', data: fill(size, 2) },
      { path: 'Small', data: fill(10, 3) },
    ];
    sameInBothReaders(streams, writeCfb(streams), true);
  }
});

refTest('cfb writer: storages inside storages, names with control characters, and the layout of an encrypted package', () => {
  const streams = [
    { path: 'EncryptedPackage', data: fill(20000, 1) },
    { path: 'EncryptionInfo', data: fill(1100, 2) },
    { path: '\u0006DataSpaces/Version', data: fill(76, 3) },
    { path: '\u0006DataSpaces/DataSpaceMap', data: fill(112, 4) },
    { path: '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', data: fill(64, 5) },
    { path: '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', data: fill(200, 6) },
  ];
  const listed = sameInBothReaders(streams, writeCfb(streams), true) as Listing;
  assert.deepEqual(listed.storages.sort(), ['\u0006DataSpaces', '\u0006DataSpaces/DataSpaceInfo', '\u0006DataSpaces/TransformInfo', '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform'].sort());
});

refTest('cfb writer: siblings are ordered by the length of their names first and without regard to case, and every tree is a red-black tree', () => {
  const names = ['b', 'AA', 'a0', 'Z', 'c', 'abc', 'ABD', 'É', 'zz', 'n'.repeat(31)];
  const streams = names.map((name, i) => ({ path: name, data: fill(i * 3, i) }));
  const listed = sameInBothReaders(streams, writeCfb(streams), true) as Listing;
  assert.deepEqual(listed.order, ['b', 'c', 'Z', 'É', 'a0', 'AA', 'zz', 'abc', 'ABD', 'n'.repeat(31)]);
  // From no entry to more than three sectors of directory: the reference reader checks order, colours and black depth of each.
  for (let count = 0; count <= 40; count++) {
    const many = [];
    for (let i = 0; i < count; i++) many.push({ path: `S${(i * 37) % 101}`, data: fill(i, i) });
    const again = sameInBothReaders(many, writeCfb(many), true) as Listing;
    assert.equal(again.order.length, count, `${count} entries`);
  }
});

refTest('cfb writer: files whose allocation table just fills a sector, and just does not', () => {
  // One FAT sector describes 128 sectors, itself included.
  for (let sectors = 120; sectors <= 132; sectors++) {
    const streams = [{ path: 'Big', data: fill(sectors * 512 - 5, sectors) }];
    const listed = sameInBothReaders(streams, writeCfb(streams), true) as Listing;
    assert.ok(listed.fat_sectors === 1 || listed.fat_sectors === 2, `${sectors} sectors of stream: ${listed.fat_sectors} FAT sectors`);
  }
});

refTest('cfb writer: a stream of more than 7 megabytes needs sectors for the table of allocation tables (DIFAT), and both readers follow them', () => {
  // 109 FAT sectors are named in the header: they describe 109 * 128 sectors, a little under 7 MB.
  for (const size of [109 * 128 * 512 - 3 * 512, 7300000, 16 * 1024 * 1024 + 17]) {
    const streams = [
      { path: 'EncryptedPackage', data: fill(size, 9) },
      { path: 'EncryptionInfo', data: fill(1000, 2) },
    ];
    const listed = sameInBothReaders(streams, writeCfb(streams), true) as Listing;
    assert.ok(listed.fat_sectors > 108, `${size} bytes: ${listed.fat_sectors} FAT sectors`);
    assert.equal(listed.difat_sectors, listed.fat_sectors <= 109 ? 0 : Math.ceil((listed.fat_sectors - 109) / 127), `${size} bytes: DIFAT sectors for ${listed.fat_sectors} FAT sectors`);
  }
});

test('cfb writer: a stream can be filled in after the file was laid out', () => {
  const layout = allocateCfb([
    { path: 'Big', size: 5000 },
    { path: 'Small', size: 100 },
    { path: 'Dir/Empty', size: 0 },
  ]);
  assert.equal(layout.stream('Big').length, 5000);
  assert.equal(layout.stream('Small').length, 100);
  assert.equal(layout.stream('Dir/Empty').length, 0);
  layout.stream('Big').set(fill(5000, 1));
  layout.stream('Small').set(fill(100, 2));
  const cfb = CfbFile.read(layout.bytes);
  assert.equal(sha256(cfb.stream('Big')), sha256(fill(5000, 1)));
  assert.equal(sha256(cfb.stream('Small')), sha256(fill(100, 2)));
  assert.equal(cfb.stream('Dir/Empty').length, 0);
  assert.throws(() => layout.stream('Missing'), /No such stream/);
});

test('cfb writer: names a compound file cannot hold are refused', () => {
  const one = (path: string) => () => writeCfb([{ path, data: fill(1) }]);
  assert.throws(one('x'.repeat(32)), /longer than 31 characters/);
  assert.throws(one(''), /empty name/);
  assert.throws(one('a//b'), /empty name/);
  assert.throws(one('a:b'), /cannot be in a name/);
  assert.throws(one('a\\b'), /cannot be in a name/);
  assert.throws(one('a!b'), /cannot be in a name/);
  assert.throws(() => writeCfb([{ path: 'Same', data: fill(1) }, { path: 'SAME', data: fill(1) }]), /twice/);
  assert.throws(() => writeCfb([{ path: 'Dir', data: fill(1) }, { path: 'Dir/Inner', data: fill(1) }]), /both a stream and a storage/);
});

test('cfb writer: file(1) calls what is written a compound document', (t) => {
  const done = spawnSync('file', ['-'], { input: writeCfb([{ path: 'EncryptionInfo', data: fill(900) }, { path: 'EncryptedPackage', data: fill(9000) }]) });
  if (done.error && (done.error as { code?: string }).code === 'ENOENT') return t.skip('file(1) not found');
  // file(1) may stop reading before all of the file was written to it: what it printed is what counts.
  assert.match(String(done.stdout), /Composite Document File V2|CDFV2|Compound Document/, `file(1) said: ${String(done.stdout)} ${String(done.stderr)} ${done.error ? done.error.message : ''}`);
});
