// A copied picture as the clipboard carries it (ui/picture-clip.ts): one copy held for the page,
// known by the mark of its HTML only; the PNG put on the clipboard of the system where the
// browser lets a page do so, and given up when another copy took the place of this one.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { PictureCopy } from '../src/ops/pictures';
import { heldPictureCopy, holdPicture, holdsPicture, picturePasteKind, releasePicture, writeSystemClipboard } from '../src/ui/picture-clip';

const copyOf = (type = 'image/png'): PictureCopy => ({ bytes: new Uint8Array([1, 2, 3]), type, rect: { x: 0, y: 0, width: 10, height: 10 }, anchor: { kind: 'twoCell' } });

test('a picture is known by the mark of its HTML alone: the copy that is held, a copy that is gone, anything else', async () => {
  releasePicture();
  assert.equal(holdsPicture(), false);
  assert.deepEqual(picturePasteKind('<span data-bym-picture="x">Picture</span>'), { kind: 'stale' });
  const copy = Promise.resolve(copyOf());
  const put = holdPicture(copy, 'Pic<ture> & "co"');
  assert.equal(put.text, 'Pic<ture> & "co"');
  // The word is in the HTML without anything that could be markup; the mark is letters and digits.
  assert.match(put.html, /^<span data-bym-picture="[a-z0-9]+">Picture  co<\/span>$/);
  assert.equal(holdsPicture(), true);
  assert.equal(heldPictureCopy(), copy);
  const found = picturePasteKind(`<meta charset="utf-8"><b>${put.html}</b>`);
  assert.equal(found.kind, 'picture');
  assert.equal((found as { copy: Promise<PictureCopy> }).copy, copy);
  assert.equal(picturePasteKind(put.html.replace(/"[a-z0-9]+"/, "'other'")).kind, 'stale');
  assert.equal(picturePasteKind('<b>picture</b>').kind, 'other');
  assert.equal(picturePasteKind(undefined).kind, 'other');
  assert.equal(picturePasteKind('').kind, 'other');
  // A second copy takes the place of the first.
  const second = holdPicture(Promise.resolve(copyOf()), 'Picture');
  assert.equal(picturePasteKind(put.html).kind, 'stale');
  assert.equal(picturePasteKind(second.html).kind, 'picture');
  releasePicture();
  assert.equal(picturePasteKind(second.html).kind, 'stale');
  assert.equal(heldPictureCopy(), undefined);
});

test('a copy that cannot be read is held all the same, and leaves nothing unhandled', async () => {
  const failed = Promise.reject(new RangeError('Cannot copy this picture: no'));
  const put = holdPicture(failed, 'Picture');
  const found = picturePasteKind(put.html) as { kind: 'picture'; copy: Promise<PictureCopy> };
  await assert.rejects(found.copy, /Cannot copy/);
  releasePicture();
});

test('the many marks of a long HTML are looked at no further than 64, and a tag without its end costs no more than its length', () => {
  const put = holdPicture(Promise.resolve(copyOf()), 'Picture');
  const many = '<i data-bym-picture="a"></i>'.repeat(100) + put.html;
  assert.equal(picturePasteKind(many).kind, 'stale');
  const open = '<b data-bym-picture ' + ' x'.repeat(500000);
  const started = Date.now();
  assert.equal(picturePasteKind(open).kind, 'other');
  assert.ok(Date.now() - started < 2000, 'an unclosed tag is scanned once');
  releasePicture();
});

test('a write of the clipboard of the system is given up once anything else is copied before its bytes came: cells, text, another page of the window', async () => {
  const g = globalThis as unknown as { navigator?: unknown; window?: unknown };
  const before = { navigator: g.navigator, window: g.window };
  const written: { [type: string]: unknown }[] = [];
  const listeners = new Map<string, Set<() => void>>();
  class Item {
    constructor(readonly items: { [type: string]: unknown }) {}
  }
  const fakeWindow = {
    ClipboardItem: Item,
    addEventListener: (type: string, listener: () => void) => {
      if (!listeners.has(type)) listeners.set(type, new Set());
      listeners.get(type)!.add(listener);
    },
    removeEventListener: (type: string, listener: () => void) => listeners.get(type)?.delete(listener),
  };
  const fire = (type: string) => [...(listeners.get(type) || [])].forEach((listener) => listener());
  try {
    Object.defineProperty(globalThis, 'navigator', { value: { clipboard: { write: (items: Item[]) => (written.push(items[0].items), Promise.resolve()) } }, configurable: true, writable: true });
    Object.defineProperty(globalThis, 'window', { value: fakeWindow, configurable: true, writable: true });
    const slow = (): { copy: Promise<PictureCopy>; release: () => void } => {
      let release: () => void = () => undefined;
      const copy = new Promise<PictureCopy>((resolve) => (release = () => resolve(copyOf())));
      return { copy, release };
    };
    const clip = await import('../src/ui/picture-clip');
    assert.equal(typeof clip.clipboardMovedOn, 'function');
    // Cells copied by the view meanwhile.
    const first = slow();
    writeSystemClipboard(holdPicture(first.copy, 'Picture'), first.copy);
    clip.clipboardMovedOn();
    first.release();
    await assert.rejects(written[0]['image/png'] as Promise<Blob>);
    // Anything copied in the window meanwhile (a text field, another editor).
    const second = slow();
    writeSystemClipboard(holdPicture(second.copy, 'Picture'), second.copy);
    fire('copy');
    second.release();
    await assert.rejects(written[1]['image/png'] as Promise<Blob>);
    // Nothing in between: written; and nothing is listened to any more.
    const third = slow();
    writeSystemClipboard(holdPicture(third.copy, 'Picture'), third.copy);
    third.release();
    assert.equal(((await written[2]['image/png']) as Blob).type, 'image/png');
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal([...listeners.values()].reduce((sum, set) => sum + set.size, 0), 0);
  } finally {
    Object.defineProperty(globalThis, 'navigator', { value: before.navigator, configurable: true, writable: true });
    Object.defineProperty(globalThis, 'window', { value: before.window, configurable: true, writable: true });
    if (before.navigator === undefined) delete g.navigator;
    if (before.window === undefined) delete g.window;
    releasePicture();
  }
});

test('the clipboard of the system gets a PNG with the word and the mark; other kinds, a copy that is no longer held, and a browser without it get nothing', async () => {
  const g = globalThis as unknown as { navigator?: unknown; window?: unknown };
  const before = { navigator: g.navigator, window: g.window };
  const written: { [type: string]: unknown }[] = [];
  class Item {
    constructor(readonly items: { [type: string]: unknown }) {}
  }
  try {
    // Without the asynchronous clipboard nothing is tried.
    writeSystemClipboard(holdPicture(Promise.resolve(copyOf()), 'Picture'), Promise.resolve(copyOf()));
    Object.defineProperty(globalThis, 'navigator', { value: { clipboard: { write: (items: Item[]) => (written.push(items[0].items), Promise.resolve()) } }, configurable: true, writable: true });
    Object.defineProperty(globalThis, 'window', { value: { ClipboardItem: Item }, configurable: true, writable: true });
    const png = Promise.resolve(copyOf());
    const put = holdPicture(png, 'Picture');
    writeSystemClipboard(put, png);
    assert.equal(written.length, 1);
    const blob = (await written[0]['image/png']) as Blob;
    assert.equal(blob.type, 'image/png');
    assert.deepEqual(Array.from(new Uint8Array(await blob.arrayBuffer())), [1, 2, 3]);
    assert.equal(await (written[0]['text/html'] as Blob).text(), put.html);
    assert.equal(await (written[0]['text/plain'] as Blob).text(), 'Picture');
    // A JPEG is not converted: the write is given up.
    const jpeg = Promise.resolve(copyOf('image/jpeg'));
    writeSystemClipboard(holdPicture(jpeg, 'Picture'), jpeg);
    await assert.rejects(written[1]['image/png'] as Promise<Blob>);
    // Another copy was made before the bytes came: the first write is given up.
    let release: (copy: PictureCopy) => void = () => undefined;
    const slow = new Promise<PictureCopy>((resolve) => (release = resolve));
    writeSystemClipboard(holdPicture(slow, 'Picture'), slow);
    holdPicture(Promise.resolve(copyOf()), 'Picture');
    release(copyOf());
    await assert.rejects(written[2]['image/png'] as Promise<Blob>);
    // A browser that refuses: nothing is thrown.
    (g.navigator as { clipboard: { write: () => Promise<void> } }).clipboard.write = () => Promise.reject(new Error('NotAllowedError'));
    const refused = Promise.resolve(copyOf());
    writeSystemClipboard(holdPicture(refused, 'Picture'), refused);
  } finally {
    Object.defineProperty(globalThis, 'navigator', { value: before.navigator, configurable: true, writable: true });
    Object.defineProperty(globalThis, 'window', { value: before.window, configurable: true, writable: true });
    if (before.navigator === undefined) delete g.navigator;
    if (before.window === undefined) delete g.window;
    releasePicture();
  }
});
