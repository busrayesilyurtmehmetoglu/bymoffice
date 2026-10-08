// ImageStore.destroy() while a picture is still being decoded in an image
// element: the address of the picture is given back when the element is done
// with it, not under it (which browsers report as a failed load).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ImageStore } from '../src/render/images';
import { png } from '../e2e/fixtures/docx';

interface FakeImage {
  src: string;
  onload: (() => void) | null;
  onerror: (() => void) | null;
  naturalWidth: number;
  naturalHeight: number;
  removeAttribute(name: string): void;
}

async function withFakes(run: (state: { images: FakeImage[]; made: string[]; revoked: string[] }) => Promise<void>): Promise<void> {
  const scope = globalThis as unknown as Record<string, unknown>;
  const saved = { Image: scope.Image, create: URL.createObjectURL, revoke: URL.revokeObjectURL, bitmap: scope.createImageBitmap };
  const state = { images: [] as FakeImage[], made: [] as string[], revoked: [] as string[] };
  scope.Image = function (this: FakeImage) {
    this.src = '';
    this.onload = this.onerror = null;
    this.naturalWidth = this.naturalHeight = 4;
    this.removeAttribute = () => undefined;
    state.images.push(this);
  };
  // No bitmap decoder: the image element is what decodes.
  scope.createImageBitmap = undefined;
  URL.createObjectURL = () => {
    const url = 'blob:fake/' + state.made.length;
    state.made.push(url);
    return url;
  };
  URL.revokeObjectURL = (url: string) => {
    state.revoked.push(url);
  };
  try {
    await run(state);
  } finally {
    scope.Image = saved.Image;
    scope.createImageBitmap = saved.bitmap;
    URL.createObjectURL = saved.create;
    URL.revokeObjectURL = saved.revoke;
  }
}

const until = async (done: () => boolean): Promise<void> => {
  for (let i = 0; i < 200 && !done(); i++) await new Promise((resolve) => setTimeout(resolve, 5));
  assert.ok(done(), 'timed out');
};

for (const outcome of ['onload', 'onerror'] as const) {
  test(`destroy during a decode: the address is given back when the decode ends (${outcome}), once`, async () => {
    await withFakes(async (state) => {
      const loaded: string[] = [];
      const store = new ImageStore({ readPart: async () => png(4, 4, [200, 0, 0]) }, 64 * 1024 * 1024, (part) => loaded.push(part));
      assert.equal(store.get('/word/media/a.png', 4, 4), undefined);
      await until(() => state.images.length === 1 && state.images[0].src !== '');
      assert.deepEqual(state.made, [state.images[0].src]);
      store.destroy();
      // Still decoding: the address must stay valid.
      assert.deepEqual(state.revoked, []);
      (state.images[0][outcome] as () => void)();
      await until(() => state.revoked.length > 0);
      await new Promise((resolve) => setTimeout(resolve, 20));
      assert.deepEqual(state.revoked, state.made);
      // Nothing of a store that is gone reaches its owner.
      assert.deepEqual(loaded, []);
      assert.equal(store.bytes, 0);
    });
  });
}

test('a decode that ends before destroy gives its address back then, and destroy gives back nothing twice', async () => {
  await withFakes(async (state) => {
    const loaded: string[] = [];
    const store = new ImageStore({ readPart: async () => png(4, 4, [200, 0, 0]) }, 64 * 1024 * 1024, (part) => loaded.push(part));
    store.get('/word/media/a.png', 4, 4);
    await until(() => state.images.length === 1 && state.images[0].src !== '');
    (state.images[0].onload as () => void)();
    await until(() => loaded.length === 1);
    assert.deepEqual(state.revoked, state.made);
    store.destroy();
    assert.deepEqual(state.revoked, state.made);
  });
});

// ------------------------------------------------------- the bitmap decoder

interface FakeBitmap {
  width: number;
  height: number;
  closed: number;
  close(): void;
}

/** A createImageBitmap that answers when the test says so, with a bitmap of the given size. */
async function withBitmaps(width: number, height: number, run: (state: { bitmaps: FakeBitmap[]; release: (() => void)[] }) => Promise<void>): Promise<void> {
  const scope = globalThis as unknown as Record<string, unknown>;
  const saved = scope.createImageBitmap;
  const state = { bitmaps: [] as FakeBitmap[], release: [] as (() => void)[] };
  scope.createImageBitmap = () =>
    new Promise<FakeBitmap>((resolve) => {
      state.release.push(() => {
        const bitmap: FakeBitmap = {
          width,
          height,
          closed: 0,
          close() {
            this.closed++;
          },
        };
        state.bitmaps.push(bitmap);
        resolve(bitmap);
      });
    });
  try {
    await run(state);
  } finally {
    scope.createImageBitmap = saved;
  }
}

// The second size is a decoder that ignores the size it was asked for: the store would draw such a bitmap down.
for (const [width, height] of [
  [4, 4],
  [64, 64],
]) {
  test(`destroy during a bitmap decode: the bitmap of ${width}x${height} that arrives afterwards is closed, once`, async () => {
    await withBitmaps(width, height, async (state) => {
      const loaded: string[] = [];
      const store = new ImageStore({ readPart: async () => png(4, 4, [200, 0, 0]) }, 64 * 1024 * 1024, (part) => loaded.push(part));
      assert.equal(store.get('/word/media/a.png', 4, 4), undefined);
      await until(() => state.release.length === 1);
      store.destroy();
      state.release[0]();
      await until(() => state.bitmaps.length === 1 && state.bitmaps[0].closed > 0);
      await new Promise((resolve) => setTimeout(resolve, 20));
      assert.equal(state.bitmaps[0].closed, 1);
      assert.deepEqual(loaded, []);
      assert.equal(store.bytes, 0);
    });
  });
}

test('a bitmap decoded before destroy is closed by destroy, once', async () => {
  await withBitmaps(4, 4, async (state) => {
    const loaded: string[] = [];
    const store = new ImageStore({ readPart: async () => png(4, 4, [200, 0, 0]) }, 64 * 1024 * 1024, (part) => loaded.push(part));
    store.get('/word/media/a.png', 4, 4);
    await until(() => state.release.length === 1);
    state.release[0]();
    await until(() => loaded.length === 1);
    assert.equal(state.bitmaps[0].closed, 0);
    store.destroy();
    assert.equal(state.bitmaps[0].closed, 1);
    store.destroy();
    assert.equal(state.bitmaps[0].closed, 1);
  });
});
