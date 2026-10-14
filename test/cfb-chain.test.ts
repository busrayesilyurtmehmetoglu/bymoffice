// A stream of a compound file is read as far as its size reaches, whatever its chain goes on to:
// many entries that name one long chain each cost only what their own size states.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-write';

/** A compound file whose stream `small` starts where the stream `pad` starts: its chain runs on through the whole of `pad`. */
function lying(padSize: number, smallSize: number, statedSmall = smallSize): Uint8Array {
  const pad = new Uint8Array(padSize);
  for (let i = 0; i < pad.length; i++) pad[i] = i & 0xff;
  const out = writeCfb([
    { path: 'Pad', data: pad },
    { path: 'Small', data: new Uint8Array(smallSize) },
  ]);
  const view = new DataView(out.buffer, out.byteOffset, out.byteLength);
  const named = (p: number, name: string) => view.getUint16(p + 64, true) === name.length * 2 + 2 && Array.from(name).every((c, k) => view.getUint16(p + k * 2, true) === c.charCodeAt(0));
  let padStart = -1;
  let small = -1;
  for (let p = 512; p + 128 <= out.length; p += 128) {
    if (named(p, 'Pad') && view.getUint32(p + 120, true) === padSize) padStart = view.getUint32(p + 116, true);
    if (named(p, 'Small') && view.getUint32(p + 120, true) === smallSize) small = p;
  }
  assert.ok(padStart >= 0 && small >= 0, 'the entries are found');
  view.setUint32(small + 116, padStart, true);
  view.setUint32(small + 120, statedSmall, true);
  return out;
}

test('cfb: a stream that names a longer chain is read as far as its size, not the whole chain', () => {
  const cfb = CfbFile.read(lying(1024 * 1024, 4096));
  const bytes = cfb.stream('Small');
  assert.equal(bytes.length, 4096);
  assert.ok(bytes.buffer.byteLength <= 4096, `no more than its sectors are copied: ${bytes.buffer.byteLength}`);
  assert.deepEqual(Array.from(bytes.subarray(0, 4)), [0, 1, 2, 3]);
  assert.equal(bytes[4095], 4095 & 0xff);
  // Thousands of reads of it stay quick.
  const started = Date.now();
  for (let i = 0; i < 5000; i++) cfb.stream('Small');
  assert.ok(Date.now() - started < 5000, `quickly: ${Date.now() - started} ms`);
});

test('cfb: the same in the mini stream', () => {
  const cfb = CfbFile.read(lying(3000, 100));
  const bytes = cfb.stream('Small');
  assert.equal(bytes.length, 100);
  assert.ok(bytes.buffer.byteLength <= 128, `no more than its mini sectors are copied: ${bytes.buffer.byteLength}`);
  assert.deepEqual(Array.from(bytes.subarray(0, 3)), [0, 1, 2]);
});

test('cfb: a chain shorter than the size it states gives what the chain holds, as before', () => {
  const cfb = CfbFile.read(lying(8192, 4096, 1024 * 1024));
  assert.equal(cfb.stream('Small').length, 8192);
  assert.equal(cfb.stream('Pad').length, 8192);
});
