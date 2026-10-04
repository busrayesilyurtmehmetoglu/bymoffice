// Printing a document, the parts that need no browser: the plan (resolution
// within the limits on memory, paper sizes), the style sheet, the element
// with the pages and its clean-up, and the wait for the print dialog.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  DEFAULT_PRINT_DPI,
  DocumentPrintError,
  MAX_PRINT_DPI,
  MAX_PRINT_PAGES,
  MAX_PRINT_SIZES,
  MIN_PRINT_DPI,
  PrintFrame,
  canvasPicture,
  openPrintDialog,
  planPrint,
  printCss,
  printLimits,
  revokePicture,
  sheetBox,
  supportsNamedPages,
  type PrintWindow,
} from '../src/ui/doc-print';

// A4 and US Letter in px of the layout (96 to the inch; twips / 15).
const A4 = { width: 11906 / 15, height: 16838 / 15 };
const A4_LANDSCAPE = { width: A4.height, height: A4.width };
const LETTER = { width: 816, height: 1056 };

const many = <T>(count: number, value: T): T[] => Array.from({ length: count }, () => value);

function code(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof DocumentPrintError, String(error));
    return error.code;
  }
  return 'no error';
}

test('plan: the default resolution, the pixels of a page and the paper size in points', () => {
  const plan = planPrint([A4, A4, A4]);
  assert.equal(plan.dpi, DEFAULT_PRINT_DPI);
  assert.equal(plan.pages.length, 3);
  // 8.27 x 11.69 inches at 200 dpi.
  assert.deepEqual([plan.pages[0].pixelWidth, plan.pages[0].pixelHeight], [1654, 2339]);
  assert.deepEqual(plan.sizes, [{ width: 595.3, height: 841.9 }]);
  assert.deepEqual(
    plan.pages.map((page) => page.size),
    [0, 0, 0],
  );
  const letter = planPrint([LETTER], { dpi: 96 });
  assert.deepEqual([letter.pages[0].pixelWidth, letter.pages[0].pixelHeight], [816, 1056]);
  assert.deepEqual(letter.sizes, [{ width: 612, height: 792 }]);
});

test('plan: the resolution option is checked and kept between the smallest and the largest', () => {
  assert.equal(planPrint([A4], { dpi: 300 }).dpi, 300);
  assert.equal(planPrint([A4], { dpi: 10 }).dpi, MIN_PRINT_DPI);
  assert.equal(planPrint([A4], { dpi: 5000 }).dpi, MAX_PRINT_DPI);
  for (const dpi of [0, -1, NaN, Infinity, '200' as unknown as number]) assert.throws(() => planPrint([A4], { dpi }), RangeError);
});

test('plan: a long document is painted at a lower resolution, and one beyond the limits is refused', () => {
  const limits = printLimits(false);
  // 155 A4 pages fit at 200 dpi; 500 do not.
  assert.equal(planPrint(many(150, A4), {}, limits).dpi, DEFAULT_PRINT_DPI);
  const long = planPrint(many(500, A4), {}, limits);
  assert.ok(long.dpi < DEFAULT_PRINT_DPI && long.dpi > 100, String(long.dpi));
  let total = 0;
  for (const page of long.pages) total += page.pixelWidth * page.pixelHeight;
  // Rounding to whole pixels may go a little over.
  assert.ok(total <= limits.maxTotalPixels * 1.01, String(total));
  // A phone has less memory: the same document gets fewer pixels, and a very long one is refused.
  const coarse = printLimits(true);
  assert.ok(planPrint(many(200, A4), {}, coarse).dpi < planPrint(many(200, A4), {}, limits).dpi);
  assert.equal(
    code(() => planPrint(many(500, A4), {}, coarse)),
    'too-long',
  );
  assert.equal(
    code(() => planPrint(many(MAX_PRINT_PAGES + 1, { width: 10, height: 10 }))),
    'too-long',
  );
  assert.equal(
    code(() => planPrint([])),
    'empty',
  );
  assert.equal(
    code(() => planPrint([{ width: NaN, height: 10 }])),
    'failed',
  );
});

test('plan: one very large page is painted at less, the others keep the resolution', () => {
  const limits = printLimits(false);
  const plan = planPrint([A4, { width: 20000, height: 20000 }], { dpi: 100 }, limits);
  assert.deepEqual([plan.pages[0].pixelWidth, plan.pages[0].pixelHeight], [827, 1169]);
  const large = plan.pages[1];
  assert.ok(large.pixelWidth * large.pixelHeight <= limits.maxPagePixels * 1.001);
  assert.equal(large.pixelWidth, large.pixelHeight);
});

test('plan: every distinct paper size is listed once, in the order of the pages', () => {
  const plan = planPrint([A4, A4_LANDSCAPE, A4, LETTER, A4_LANDSCAPE]);
  assert.deepEqual(plan.sizes, [
    { width: 595.3, height: 841.9 },
    { width: 841.9, height: 595.3 },
    { width: 612, height: 792 },
  ]);
  assert.deepEqual(
    plan.pages.map((page) => page.size),
    [0, 1, 0, 2, 1],
  );
  // More sizes than there are page rules for: the rest is fitted into the first size.
  const sizes = Array.from({ length: MAX_PRINT_SIZES + 3 }, (_, i) => ({ width: 500 + i * 10, height: 700 }));
  const wide = planPrint(sizes);
  assert.equal(wide.sizes.length, MAX_PRINT_SIZES);
  assert.equal(wide.pages[MAX_PRINT_SIZES - 1].size, MAX_PRINT_SIZES - 1);
  assert.equal(wide.pages[MAX_PRINT_SIZES].size, -1);
});

test('style sheet: the paper of the document, nothing but the pages in print, nothing of them on screen', () => {
  const single = printCss(planPrint([A4]), true);
  assert.ok(single.includes('@page{size:595.3pt 841.9pt;margin:0}'));
  assert.ok(!single.includes('@page bym-print-'));
  assert.ok(single.includes('@media screen{.bym-print{display:none!important}}'));
  assert.ok(single.includes('body>:not(.bym-print){display:none!important}'));

  const plan = planPrint([A4, A4_LANDSCAPE]);
  const mixed = printCss(plan, true);
  assert.ok(mixed.includes('@page{size:595.3pt 841.9pt;margin:0}'));
  assert.ok(mixed.includes('@page bym-print-0{size:595.3pt 841.9pt;margin:0}'));
  assert.ok(mixed.includes('@page bym-print-1{size:841.9pt 595.3pt;margin:0}'));
  assert.ok(mixed.includes('.bym-print-s1{page:bym-print-1}'));
  // A browser that does not know named pages gets the first size only.
  const plain = printCss(plan, false);
  assert.ok(plain.includes('@page{size:595.3pt 841.9pt;margin:0}'));
  assert.ok(!plain.includes('bym-print-1'));
});

test('sheets: paper of the page’s own size with named pages, fitted into the first size without', () => {
  const plan = planPrint([A4, A4_LANDSCAPE, { width: 300, height: 300 }]);
  assert.deepEqual(sheetBox(plan, 0, false), { width: 595.3, height: 841.9, left: 0, top: 0, imageWidth: 595.3, imageHeight: 841.9 });
  assert.deepEqual(sheetBox(plan, 1, true), { width: 841.9, height: 595.3, left: 0, top: 0, imageWidth: 841.9, imageHeight: 595.3 });
  // Landscape on portrait paper: scaled to the width, in the middle.
  const fitted = sheetBox(plan, 1, false);
  assert.deepEqual([fitted.width, fitted.height, fitted.left], [595.3, 841.9, 0]);
  assert.ok(Math.abs(fitted.imageWidth - 595.3) < 0.01 && Math.abs(fitted.imageHeight - 595.3 * (595.3 / 841.9)) < 0.01);
  assert.ok(Math.abs(fitted.top - (841.9 - fitted.imageHeight) / 2) < 0.01);
  // A smaller page is not enlarged.
  const small = sheetBox(plan, 2, false);
  assert.deepEqual([small.imageWidth, small.imageHeight, small.left, small.top], [225, 225, 185.15, 308.45]);
});

test('named pages are used only where the browser knows the page property', () => {
  assert.equal(supportsNamedPages({}), false);
  assert.equal(supportsNamedPages({ CSS: { supports: (property) => property === 'page' } }), true);
  assert.equal(supportsNamedPages({ CSS: { supports: () => false } }), false);
  assert.equal(
    supportsNamedPages({
      CSS: {
        supports: () => {
          throw new Error('no');
        },
      },
    }),
    false,
  );
});

// --------------------------------------------------------------- a small DOM

class FakeElement {
  className = '';
  textContent = '';
  style = { cssText: '' };
  attributes = new Map<string, string>();
  children: FakeElement[] = [];
  parentNode: FakeElement | null = null;
  listeners = new Map<string, (() => void)[]>();
  constructor(readonly tag: string) {}
  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
  appendChild(child: FakeElement): void {
    child.parentNode = this;
    this.children.push(child);
  }
  removeChild(child: FakeElement): void {
    this.children.splice(this.children.indexOf(child), 1);
    child.parentNode = null;
  }
  addEventListener(type: string, listener: () => void): void {
    this.listeners.set(type, (this.listeners.get(type) || []).concat(listener));
  }
  fire(type: string): void {
    for (const listener of this.listeners.get(type) || []) listener();
  }
}

function fakeDocument() {
  const head = new FakeElement('head');
  const body = new FakeElement('body');
  const doc = { head, body, documentElement: new FakeElement('html'), createElement: (tag: string) => new FakeElement(tag) };
  return { head, body, doc: doc as unknown as Document };
}

const imageOf = (sheet: FakeElement): FakeElement => sheet.children[0];

test('frame: one image per page at the level of the document; clean-up takes everything out and revokes the pictures', async () => {
  const { head, body, doc } = fakeDocument();
  body.appendChild(new FakeElement('main'));
  const revoked: string[] = [];
  const plan = planPrint([A4, A4_LANDSCAPE]);
  const frame = new PrintFrame(doc, plan, true, (url) => revoked.push(url));
  assert.equal(head.children.length, 1);
  assert.equal(head.children[0].tag, 'style');
  assert.equal(head.children[0].textContent, printCss(plan, true));
  assert.equal(body.children.length, 2);
  const element = body.children[1];
  assert.equal(element.className, 'bym-print');
  assert.equal(element.style.cssText, 'display:none');

  const first = frame.add('blob:one', 'Page 1');
  let done = false;
  void first.then(() => (done = true));
  await Promise.resolve();
  // The page counts only once the browser has its picture.
  assert.equal(done, false);
  imageOf(element.children[0]).fire('load');
  await first;
  const second = frame.add('blob:two', 'Page 2');
  imageOf(element.children[1]).fire('load');
  await second;
  assert.equal(frame.count, 2);

  const sheets = element.children;
  assert.deepEqual(
    sheets.map((sheet) => sheet.className),
    ['bym-print-sheet bym-print-s0', 'bym-print-sheet bym-print-s1'],
  );
  assert.equal(sheets[0].style.cssText, 'width:595.3pt;height:841.9pt');
  assert.equal(sheets[1].style.cssText, 'width:841.9pt;height:595.3pt');
  const image = imageOf(sheets[1]);
  assert.equal(image.tag, 'img');
  assert.equal(image.attributes.get('src'), 'blob:two');
  // Only the view's own words are in the element.
  assert.equal(image.attributes.get('alt'), 'Page 2');
  assert.equal(image.attributes.get('width'), String(plan.pages[1].pixelWidth));
  assert.equal(image.style.cssText, 'left:0pt;top:0pt;width:841.9pt;height:595.3pt');

  // More pages than the plan has are refused, and their picture is given back.
  await assert.rejects(frame.add('blob:three', 'Page 3'), (error: unknown) => error instanceof DocumentPrintError && error.code === 'cancelled');
  assert.deepEqual(revoked, ['blob:three']);

  frame.dispose();
  assert.equal(head.children.length, 0);
  assert.equal(body.children.length, 1);
  assert.deepEqual(revoked, ['blob:three', 'blob:one', 'blob:two']);
  // Twice is once.
  frame.dispose();
  assert.equal(revoked.length, 3);
  await assert.rejects(frame.add('blob:late', 'Page 1'), DocumentPrintError);
  assert.equal(revoked[3], 'blob:late');
});

test('frame: a picture that cannot be loaded fails the page; clean-up ends the wait for one that is loading', async () => {
  const { body, doc } = fakeDocument();
  const revoked: string[] = [];
  const plan = planPrint([A4, A4]);
  const frame = new PrintFrame(doc, plan, false, (url) => revoked.push(url));
  const broken = frame.add('blob:broken', 'Page 1');
  imageOf(body.children[0].children[0]).fire('error');
  await assert.rejects(broken, (error: unknown) => error instanceof DocumentPrintError && error.code === 'failed');
  const waiting = frame.add('blob:waiting', 'Page 2');
  frame.dispose();
  await assert.rejects(waiting, (error: unknown) => error instanceof DocumentPrintError && error.code === 'cancelled');
  assert.deepEqual(revoked, ['blob:broken', 'blob:waiting']);
  assert.equal(body.children.length, 0);
});

test('frame: without named pages a page of another size is put on a sheet of the first size', async () => {
  const { body, doc } = fakeDocument();
  const frame = new PrintFrame(doc, planPrint([A4, A4_LANDSCAPE]), false, () => undefined);
  // Never loaded: the clean-up below ends the wait.
  frame.add('blob:1', 'Page 1').catch(() => undefined);
  frame.add('blob:2', 'Page 2').catch(() => undefined);
  const sheet = body.children[0].children[1];
  assert.equal(sheet.style.cssText, 'width:595.3pt;height:841.9pt');
  assert.match(imageOf(sheet).style.cssText, /^left:0pt;top:210\.\d+pt;width:595\.3pt;height:420\.\d+pt$/);
  frame.dispose();
});

// ------------------------------------------------------------- the dialog

class FakeWindow implements PrintWindow {
  print?: unknown;
  listeners: { type: string; listener: () => void }[] = [];
  timers: { callback: () => void; ms: number }[] = [];
  cleared: unknown[] = [];
  addEventListener(type: string, listener: () => void): void {
    this.listeners.push({ type, listener });
  }
  removeEventListener(type: string, listener: () => void): void {
    this.listeners = this.listeners.filter((entry) => entry.type !== type || entry.listener !== listener);
  }
  setTimeout(callback: () => void, ms: number): unknown {
    this.timers.push({ callback, ms });
    return this.timers.length;
  }
  clearTimeout(handle: never): void {
    this.cleared.push(handle);
  }
  fire(type: string): void {
    for (const entry of this.listeners.slice()) if (entry.type === type) entry.listener();
  }
}

const pending = async (promise: Promise<unknown>): Promise<boolean> => {
  let settled = false;
  const mark = () => (settled = true);
  void promise.then(mark, mark);
  await new Promise((resolve) => setImmediate(resolve));
  return !settled;
};

test('dialog: a window without print() is refused; a print() that throws fails', async () => {
  const none = new FakeWindow();
  await assert.rejects(openPrintDialog(none).closed, (error: unknown) => error instanceof DocumentPrintError && error.code === 'unavailable');
  assert.equal(none.listeners.length, 0);
  const throwing = new FakeWindow();
  throwing.print = () => {
    throw new Error('blocked');
  };
  await assert.rejects(openPrintDialog(throwing).closed, (error: unknown) => error instanceof DocumentPrintError && error.code === 'failed' && /blocked/.test(error.message));
  assert.equal(throwing.listeners.length, 0);
});

test('dialog: closed with afterprint, whether it comes inside print() or after it; no listener stays behind', async () => {
  // A browser whose print() blocks: the event comes before print() returns.
  const blocking = new FakeWindow();
  let calls = 0;
  blocking.print = () => {
    calls++;
    blocking.fire('beforeprint');
    blocking.fire('afterprint');
  };
  await openPrintDialog(blocking).closed;
  assert.equal(calls, 1);
  assert.equal(blocking.listeners.length, 0);
  assert.equal(blocking.timers.length, 0);

  const later = new FakeWindow();
  later.print = () => undefined;
  const dialog = openPrintDialog(later, 1234);
  assert.equal(await pending(dialog.closed), true);
  assert.equal(later.timers[0].ms, 1234);
  later.fire('afterprint');
  await dialog.closed;
  assert.equal(later.listeners.length, 0);
  assert.equal(later.cleared.length, 1);
});

test('dialog: where no afterprint comes, a press in the window, the time limit or finish() end the wait', async () => {
  for (const end of ['pointerdown', 'mousedown', 'touchstart', 'keydown', 'timer', 'finish']) {
    const win = new FakeWindow();
    win.print = () => undefined;
    const dialog = openPrintDialog(win);
    assert.equal(await pending(dialog.closed), true, end);
    if (end === 'timer') win.timers[0].callback();
    else if (end === 'finish') dialog.finish();
    else win.fire(end);
    await dialog.closed;
    assert.equal(win.listeners.length, 0, end);
    // Once is enough.
    dialog.finish();
  }
});

// ------------------------------------------------------------ the pictures

test('pictures: a canvas becomes a PNG at a blob: address, or at a data: address without toBlob()', async () => {
  const made: Blob[] = [];
  const urls = {
    createObjectURL: (blob: Blob) => {
      made.push(blob);
      return `blob:made-${made.length}`;
    },
    revokeObjectURL: (url: string) => revoked.push(url),
  };
  const revoked: string[] = [];
  const blob = new Blob([new Uint8Array(42)], { type: 'image/png' });
  const withBlob = { toBlob: (callback: (blob: Blob | null) => void, type: string) => callback(type === 'image/png' ? blob : null) } as unknown as HTMLCanvasElement;
  assert.deepEqual(await canvasPicture(withBlob, urls), { url: 'blob:made-1', bytes: 42 });
  // While the browser encodes, frames are asked for, so that it has idle time to encode in; none after the picture is there.
  let deliver: (blob: Blob | null) => void = () => undefined;
  const slow = { toBlob: (callback: (blob: Blob | null) => void) => (deliver = callback) } as unknown as HTMLCanvasElement;
  const frames: (() => void)[] = [];
  const picture = canvasPicture(slow, urls, (callback) => frames.push(callback));
  assert.equal(frames.length, 1);
  frames[0]();
  assert.equal(frames.length, 2);
  deliver(blob);
  frames[1]();
  assert.equal(frames.length, 2);
  assert.deepEqual(await picture, { url: 'blob:made-2', bytes: 42 });
  const empty = { toBlob: (callback: (blob: Blob | null) => void) => callback(null) } as unknown as HTMLCanvasElement;
  await assert.rejects(canvasPicture(empty, urls), DocumentPrintError);
  const old = { toDataURL: () => 'data:image/png;base64,AAAA' } as unknown as HTMLCanvasElement;
  assert.deepEqual(await canvasPicture(old, urls), { url: 'data:image/png;base64,AAAA', bytes: 26 });
  const broken = {
    toDataURL: () => {
      throw new Error('tainted');
    },
  } as unknown as HTMLCanvasElement;
  await assert.rejects(canvasPicture(broken, urls), DocumentPrintError);

  revokePicture('blob:made-1', urls);
  revokePicture('data:image/png;base64,AAAA', urls);
  assert.deepEqual(revoked, ['blob:made-1']);
});
