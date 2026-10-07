import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type BridgeHost } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import type { FontRegistration } from '../src/render/fonts';

interface Fake {
  host: BridgeHost;
  sent: any[];
  fonts: FontRegistration[];
  views: number;
  fetched: { url: string; maxBytes: number }[];
  /** What the address of a font gives; default five bytes. */
  files: { [url: string]: Uint8Array | Error };
  timers: { callback: () => void; ms: number }[];
  /** What registering a family does: nothing listed loads at once. */
  outcome: { [family: string]: 'fail' | 'never' };
}

function fake(withRegister = true): Fake {
  const state: Fake = { host: undefined as unknown as BridgeHost, sent: [], fonts: [], views: 0, fetched: [], files: {}, timers: [], outcome: {} };
  state.host = {
    container: {} as HTMLElement,
    version: '9.9.9',
    baseUrl: 'https://app.example.com/bymoffice.webview.html',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: async () =>
      ({
        format: 'xlsx',
        workbook: async () => ({ sheets: [{ name: 'Data', state: 'visible', kind: 'worksheet' }] }),
        save: async () => new Uint8Array([1]),
      }) as unknown as OfficeFile,
    createView() {
      state.views++;
      return { history: { canUndo: false, canRedo: false }, activeSheet: 0, isEditing: false, destroy() {} } as any;
    },
    fetchBytes: async (url, limits) => {
      state.fetched.push({ url, maxBytes: limits.maxBytes });
      const file = state.files[url];
      if (file instanceof Error) throw file;
      return file || new Uint8Array([7, 7, 7, 7, 7]);
    },
    frame: () => undefined,
    timer: (callback, ms) => state.timers.push({ callback, ms }),
  };
  if (withRegister) {
    state.host.registerFont = (font) => {
      state.fonts.push(font);
      const outcome = state.outcome[font.family];
      if (outcome === 'fail') return Promise.reject(new Error('A network error occurred.'));
      if (outcome === 'never') return new Promise(() => undefined);
      return Promise.resolve();
    };
  }
  return state;
}

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));
const FONT = base64Encode(new Uint8Array([0, 1, 0, 0, 9]));
const answer = (f: Fake) => f.sent.filter((message) => message.type === 'loaded' || message.type === 'error');
const settle = async () => {
  for (let i = 0; i < 20; i++) await new Promise((resolve) => setImmediate(resolve));
};

test('bridge fonts: a load registers its fonts before the file is shown', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({
    type: 'load',
    id: '1',
    data: DATA,
    fonts: [
      { family: 'Mine', data: FONT, as: ['Bookman Old Style', 'Book Antiqua'], weight: 'bold', style: 'italic', prefer: true },
      { family: 'Theirs', url: 'fonts/theirs.woff2', as: 'Garamond', weight: 600 },
      { family: 'Plain', url: 'data:font/ttf;base64,AAEAAAA=' },
    ],
  });
  assert.equal(f.fonts.length, 3);
  assert.deepEqual(f.fonts[0], { family: 'Mine', source: new Uint8Array([0, 1, 0, 0, 9]), weight: 'bold', style: 'italic', as: ['Bookman Old Style', 'Book Antiqua'], prefer: true });
  // A font at an address is fetched by the page, as the file would be, and registered as bytes.
  assert.deepEqual(f.fonts[1], { family: 'Theirs', source: new Uint8Array([7, 7, 7, 7, 7]), weight: 600, as: ['Garamond'] });
  assert.deepEqual(f.fonts[2], { family: 'Plain', source: new Uint8Array([7, 7, 7, 7, 7]) });
  assert.deepEqual(f.fetched, [
    { url: 'https://app.example.com/fonts/theirs.woff2', maxBytes: 24 * 1024 * 1024 },
    { url: 'data:font/ttf;base64,AAEAAAA=', maxBytes: 24 * 1024 * 1024 },
  ]);
  const loaded = answer(f);
  assert.equal(loaded.length, 1);
  assert.equal(loaded[0].type, 'loaded');
  assert.equal(loaded[0].id, '1');
  assert.equal(loaded[0].warnings, undefined);
  assert.equal(f.views, 1);
});

test('bridge fonts: a font that fails or does not arrive in time is a warning, and the file is shown', async () => {
  const f = fake();
  f.outcome = { Broken: 'fail', Slow: 'never' };
  const bridge = createBridge(f.host);
  const done = bridge.receive({
    type: 'load',
    id: '2',
    data: DATA,
    fonts: [
      { family: 'Good', data: FONT },
      { family: 'Broken', data: FONT },
      { family: 'Slow', data: FONT },
    ],
  });
  await settle();
  assert.deepEqual(answer(f), [], 'the fonts are waited for');
  assert.equal(f.views, 0);
  // The wait has an end.
  const waits = f.timers.filter((entry) => entry.ms === 3000);
  assert.equal(waits.length, 1);
  waits[0].callback();
  await done;
  const loaded = answer(f);
  assert.equal(loaded.length, 1);
  assert.equal(loaded[0].type, 'loaded');
  assert.deepEqual(loaded[0].warnings, ['The font "Broken" could not be loaded: A network error occurred.', 'The font "Slow" is still loading; the file is shown without it until it arrives']);
  assert.equal(f.views, 1);
});

test('bridge fonts: a font whose address gives nothing, or too much, is a warning', async () => {
  const f = fake();
  f.files['https://app.example.com/missing.ttf'] = new Error('HTTP 404');
  f.files['https://app.example.com/huge.ttf'] = new Uint8Array(24 * 1024 * 1024 + 1);
  f.files['https://app.example.com/empty.ttf'] = new Uint8Array(0);
  await createBridge(f.host).receive({
    type: 'load',
    id: '4',
    data: DATA,
    fonts: [
      { family: 'Missing', url: 'missing.ttf' },
      { family: 'Huge', url: 'huge.ttf' },
      { family: 'Empty', url: 'empty.ttf' },
    ],
  });
  const loaded = answer(f);
  assert.equal(loaded[0].type, 'loaded');
  assert.deepEqual(loaded[0].warnings, [
    'The font "Missing" could not be loaded: HTTP 404',
    'The font "Huge" could not be loaded: the font file is larger than 25165824 bytes',
    'The font "Empty" could not be loaded: the font file is empty',
  ]);
  assert.equal(f.fonts.length, 0);

  // A page that cannot fetch says so.
  const g = fake();
  g.host.fetchBytes = undefined;
  await createBridge(g.host).receive({ type: 'load', id: '5', data: DATA, fonts: [{ family: 'Remote', url: 'a.ttf' }] });
  assert.deepEqual(answer(g)[0].warnings, ['The font "Remote" could not be loaded: loading from a URL is not available here']);
});

test('bridge fonts: without web fonts every font is a warning', async () => {
  const f = fake(false);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: '3', data: DATA, fonts: [{ family: 'Mine', data: FONT }] });
  const loaded = answer(f);
  assert.equal(loaded[0].type, 'loaded');
  assert.equal(loaded[0].warnings.length, 1);
  assert.match(loaded[0].warnings[0], /^The font "Mine" could not be loaded: Fonts cannot be registered here/);
});

test('bridge fonts: a load whose fonts do not follow the contract is refused whole', async () => {
  const many = new Array(17).fill({ family: 'A', data: FONT });
  const bad: [unknown, RegExp][] = [
    ['Mine', /"fonts" must be a list/],
    [{ family: 'A', data: FONT }, /"fonts" must be a list/],
    [many, /more than 16/],
    [['Mine'], /fonts\[0\] must be an object/],
    [[null], /fonts\[0\] must be an object/],
    [[{ data: FONT }], /fonts\[0\]: "family" is missing/],
    [[{ family: 7, data: FONT }], /fonts\[0\]: "family" must be a string/],
    [[{ family: '', data: FONT }], /fonts\[0\]: "family" must be a font name/],
    [[{ family: '";()', data: FONT }], /fonts\[0\]: "family" must be a font name/],
    [[{ family: 'a'.repeat(256), data: FONT }], /fonts\[0\]: "family" is too long/],
    [[{ family: 'A', data: FONT }, { family: 'B' }], /fonts\[1\] needs exactly one of "data" and "url"/],
    [[{ family: 'A', data: FONT, url: 'a.ttf' }], /fonts\[0\] needs exactly one of "data" and "url"/],
    [[{ family: 'A', data: 'not base64!' }], /fonts\[0\]: Invalid base64/],
    [[{ family: 'A', data: '' }], /fonts\[0\]: "data" is empty/],
    [[{ family: 'A', data: 'A'.repeat(32 * 1024 * 1024 + 4) }], /fonts\[0\]: "data" is too long/],
    [[{ family: 'A', url: 'https://elsewhere.example.com/a.ttf' }], /fonts\[0\]: "url" must be a blob: or data: URL or an address of the page's own origin/],
    [[{ family: 'A', url: 'javascript:alert(1)' }], /fonts\[0\]: "url" must be/],
    [[{ family: 'A', url: 7 }], /fonts\[0\]: "url" must be a string/],
    [[{ family: 'A', data: FONT, weight: 'heavy' }], /fonts\[0\]: "weight" must be "normal", "bold" or a number from 1 to 1000/],
    [[{ family: 'A', data: FONT, weight: 0 }], /fonts\[0\]: "weight"/],
    [[{ family: 'A', data: FONT, style: 'oblique' }], /fonts\[0\]: "style" must be "normal" or "italic"/],
    [[{ family: 'A', data: FONT, as: 7 }], /fonts\[0\]: "as" must be a font name or a list of font names/],
    [[{ family: 'A', data: FONT, as: ['B', 7] }], /fonts\[0\]: "as" must be a font name or a list of font names/],
    [[{ family: 'A', data: FONT, as: new Array(33).fill('B') }], /fonts\[0\]: "as" names more than 32 fonts/],
    [[{ family: 'A', data: FONT, as: ['b'.repeat(256)] }], /fonts\[0\]: "as" must be a font name or a list of font names/],
    [[{ family: 'A', data: FONT, prefer: 'yes' }], /fonts\[0\]: "prefer" must be a boolean/],
  ];
  for (const [fonts, message] of bad) {
    const f = fake();
    await createBridge(f.host).receive({ type: 'load', id: 'x', data: DATA, fonts });
    const out = answer(f);
    assert.equal(out.length, 1, String(message));
    assert.equal(out[0].type, 'error', String(message));
    assert.equal(out[0].code, 'bad-message', String(message));
    assert.match(out[0].message, message);
    assert.equal(f.fonts.length, 0, 'nothing is registered');
    assert.equal(f.views, 0);
  }
});

test('bridge fonts: a framed page takes fonts as data only', async () => {
  const f = fake();
  f.host.allowUrl = false;
  await createBridge(f.host).receive({ type: 'load', id: 'x', data: DATA, fonts: [{ family: 'A', url: 'a.ttf' }] });
  assert.equal(answer(f)[0].code, 'bad-message');
  assert.match(answer(f)[0].message, /fonts\[0\]: "url" is not accepted by a framed page/);
  assert.equal(f.fonts.length, 0);
});

test('bridge fonts: an empty list and no list are the same', async () => {
  const f = fake();
  await createBridge(f.host).receive({ type: 'load', id: 'x', data: DATA, fonts: [] });
  assert.equal(answer(f)[0].type, 'loaded');
  assert.equal(f.timers.filter((entry) => entry.ms === 3000).length, 0);
});

test('bridge fonts: one face twice in a load is refused', async () => {
  const twice: unknown[][] = [
    [{ family: 'A', data: FONT }, { family: 'a', data: FONT }],
    [{ family: 'A', data: FONT, weight: 'bold' }, { family: 'B', data: FONT }, { family: ' A ', url: 'a.ttf', weight: 700 }],
    [{ family: 'A', data: FONT, weight: 'normal', style: 'normal' }, { family: 'A', data: FONT, weight: 400 }],
    [{ family: 'A "x"', data: FONT, style: 'italic' }, { family: 'A x', data: FONT, style: 'italic' }],
  ];
  for (const fonts of twice) {
    const f = fake();
    await createBridge(f.host).receive({ type: 'load', id: 'x', data: DATA, fonts });
    const out = answer(f);
    assert.equal(out[0].type, 'error');
    assert.equal(out[0].code, 'bad-message');
    assert.match(out[0].message, /^fonts\[[12]\] is the same face as fonts\[0\] \(family, weight and style\)$/);
    assert.equal(f.fonts.length, 0);
    assert.equal(f.views, 0);
  }
  // Other faces of one family are what a load is expected to bring.
  const f = fake();
  await createBridge(f.host).receive({
    type: 'load',
    id: 'y',
    data: DATA,
    fonts: [
      { family: 'A', data: FONT },
      { family: 'A', data: FONT, weight: 'bold' },
      { family: 'A', data: FONT, style: 'italic' },
      { family: 'A', data: FONT, weight: 700, style: 'italic' },
    ],
  });
  assert.equal(answer(f)[0].type, 'loaded');
  assert.equal(f.fonts.length, 4);
});

test('bridge fonts: a load that waits for its fonts gives way to a later load and to close', async () => {
  const f = fake();
  f.outcome = { Slow: 'never' };
  const bridge = createBridge(f.host);
  const first = bridge.receive({ type: 'load', id: 'first', data: DATA, fonts: [{ family: 'Slow', data: FONT }] });
  await settle();
  assert.deepEqual(answer(f), []);
  const second = bridge.receive({ type: 'load', id: 'second', data: DATA });
  await settle();
  // The wait of the first load ends; it was overtaken and shows nothing.
  for (const entry of f.timers.filter((timer) => timer.ms === 3000)) entry.callback();
  await first;
  await second;
  const out = answer(f);
  assert.deepEqual(
    out.map((message) => [message.type, message.id, message.code]),
    [
      ['error', 'first', 'load-failed'],
      ['loaded', 'second', undefined],
    ],
  );
  assert.equal(out[0].message, 'The load was cancelled');
  assert.equal(f.views, 1);

  const g = fake();
  g.outcome = { Slow: 'never' };
  const other = createBridge(g.host);
  const waiting = other.receive({ type: 'load', id: 'only', data: DATA, fonts: [{ family: 'Slow', data: FONT }] });
  await settle();
  const closing = other.receive({ type: 'close', id: 'c' });
  await settle();
  for (const entry of g.timers.filter((timer) => timer.ms === 3000)) entry.callback();
  await waiting;
  await closing;
  assert.deepEqual(answer(g).map((message) => [message.type, message.id, message.code]), [['error', 'only', 'load-failed']]);
  assert.equal(g.views, 0);
});
